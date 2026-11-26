"""
ENTREGA SABESP - fluxo input / output

COMO FUNCIONA
-------------
1. Voce joga uma ou mais planilhas (.xlsx) na pasta  input/
2. Roda:  python entrega_sabesp.py
3. Pra CADA planilha do input, o codigo cria em  output/<nome_da_planilha>/ :
      - fotos/              -> todas as imagens baixadas (nomeadas por codigo + tipo)
      - linkagem_fotos.xlsx -> uma linha por foto, ligando arquivo <-> codigo <-> tipo,
                               com link clicavel (relativo) que abre a foto da pasta
      - fotos_com_erro.csv  -> (so se algo falhar) lista o que nao baixou
   => Depois e so zipar a pasta de output daquela planilha e mandar pra Sabesp.

O codigo identifica as colunas PELO NOME DO CABECALHO (nao pela letra). Por isso,
toda planilha que voce subir no input deve seguir o  template_planilha.xlsx
(mesmos cabecalhos na linha 1). O template vem vazio, so pra guiar.

ESTRUTURA DE PASTAS (crie assim, ou o codigo cria sozinho):
    projeto/
      entrega_sabesp.py
      template_planilha.xlsx
      input/      <- planilhas entram aqui
      output/     <- resultado sai aqui

IMPORTANTE: rode NA SUA MAQUINA. So a sua rede acessa o s3.deepessoas.com.
Instale antes:  pip install openpyxl requests
Pode parar e rodar de novo: ele PULA as fotos ja baixadas (retoma de onde parou).
"""

import os
import re
import csv
import glob
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# ============================================================
# CONFIGURACOES
# ============================================================
PASTA_INPUT = "input"
PASTA_OUTPUT = "output"
SUBPASTA_FOTOS = "fotos"
LINHA_INICIAL = 2          # primeira linha de dados (pula cabecalho)

THREADS = 8                # downloads simultaneos
TIMEOUT = 60               # segundos por foto
TENTATIVAS = 3             # re-tentativas por foto

# --- Colunas identificadas PELO NOME DO CABECALHO ---
# Identificacao (pra nomear o arquivo e preencher a linkagem):
H_CODIGO      = "Código de cadastro interno"
H_CODE_DEEP   = "Code Deep"
H_CIDADE      = "Cidade"
H_LOGRA       = "Logradouro"
H_NUMERO      = "Número"
H_BAIRRO      = "Bairro:"
H_TIPO_ADESAO = "Qual tipo de adesão?"
H_STATUS      = "Status do cadastro"

# Cabecalhos das colunas que contem FOTOS/DOCUMENTOS pra baixar.
# (A coluna "Link Backoffice" NAO entra: e link de sistema, nao e foto.)
CABECALHOS_FOTO = [
    "Foto da fachada do imóvel",
    "Foto do documento RG e CPF (frente )",
    "Foto do documento RG e CPF (verso)",
    "Foto do cartão CNPJ",
    "Foto do documento de constituição",
    "Foto do documento do representante (RG/CPF)",
    "Documento que comprova a representação",
    "Foto do documento do imóvel",
    "Foto do hidrômetro",
    "Foto da conta de água",
    "Foto da solução de água",
    "Fotos da solução de esgoto",
    "Foto do termo de adesão",
    "Foto de responsabilidade do imóvel",
]
# ============================================================


def norm(s):
    """normaliza cabecalho: sem acento, minusculo, sem espacos extras."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def slug(texto, padrao="x"):
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    texto = re.sub(r"[^A-Za-z0-9]+", "_", texto).strip("_")
    return texto or padrao


def extensao(url):
    ext = os.path.splitext(urlparse(url).path)[1].lower()
    return ext if ext in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".pdf") else ".jpeg"


def mapa_colunas(ws):
    """le a linha 1 e devolve {cabecalho_normalizado: indice_coluna}."""
    mapa = {}
    for c, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), start=1):
        h = norm(cell.value)
        if h and h not in mapa:
            mapa[h] = c
    return mapa


def processar_planilha(caminho):
    nome = os.path.splitext(os.path.basename(caminho))[0]
    print(f"\n================  {nome}  ================")
    pasta_saida = os.path.join(PASTA_OUTPUT, nome)
    pasta_fotos = os.path.join(pasta_saida, SUBPASTA_FOTOS)
    os.makedirs(pasta_fotos, exist_ok=True)

    wb = load_workbook(caminho, read_only=True)
    ws = wb.active
    mapa = mapa_colunas(ws)

    def idx(header):
        return mapa.get(norm(header))

    def val(row, header):
        i = idx(header)
        if not i or i > len(row):
            return ""
        v = row[i - 1]
        return "" if v is None else str(v).strip()

    # confere cabecalhos de foto presentes
    cols_foto = [(h, idx(h)) for h in CABECALHOS_FOTO if idx(h)]
    faltando = [h for h in CABECALHOS_FOTO if not idx(h)]
    if faltando:
        print("  AVISO - cabecalhos de foto nao encontrados (ignorados):")
        for h in faltando:
            print("     -", h)
    if not cols_foto:
        print("  Nenhuma coluna de foto encontrada. Confira se a planilha segue o template.")
        wb.close()
        return

    # monta tarefas (uma por foto), com nome de arquivo unico
    tarefas = []
    usados = {}
    for row in ws.iter_rows(min_row=LINHA_INICIAL, values_only=True):
        codigo = val(row, H_CODIGO) or val(row, H_CODE_DEEP)
        if not codigo:
            continue
        info = {
            "codigo": codigo,
            "code_deep": val(row, H_CODE_DEEP),
            "cidade": val(row, H_CIDADE),
            "endereco": f"{val(row, H_LOGRA)}, {val(row, H_NUMERO)}".strip(", "),
            "bairro": val(row, H_BAIRRO),
            "tipo_adesao": val(row, H_TIPO_ADESAO),
            "status": val(row, H_STATUS),
        }
        for h, i in cols_foto:
            raw = row[i - 1]
            if raw is None:
                continue
            # uma celula pode ter VARIAS fotos (separadas por quebra de linha
            # ou espaco). Separa cada URL e baixa como arquivos distintos.
            urls = [u.strip() for u in re.split(r"\s+", str(raw))
                    if u.strip().lower().startswith("http")]
            for url in urls:
                base = f"{slug(codigo)}__{slug(h)}"
                usados[base] = usados.get(base, 0) + 1
                if usados[base] > 1:
                    base = f"{base}_{usados[base]}"  # 2a, 3a foto do mesmo tipo
                tarefas.append({**info, "tipo_foto": h, "url": url,
                                "arquivo": base + extensao(url)})
    wb.close()
    print(f"  Fotos a processar: {len(tarefas)}")
    if not tarefas:
        return

    # download (pulando o que ja existe)
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})

    def baixar(t):
        destino = os.path.join(pasta_fotos, t["arquivo"])
        if os.path.exists(destino) and os.path.getsize(destino) > 0:
            return (t, "pulado", "")
        erro = ""
        for tent in range(1, TENTATIVAS + 1):
            try:
                r = sess.get(t["url"], timeout=TIMEOUT, stream=True)
                r.raise_for_status()
                with open(destino, "wb") as f:
                    for chunk in r.iter_content(8192):
                        f.write(chunk)
                if os.path.getsize(destino) > 0:
                    return (t, "ok", "")
                erro = "arquivo vazio"
            except Exception as e:
                erro = f"{type(e).__name__}: {e}"
                time.sleep(1.0 * tent)
        if os.path.exists(destino):
            os.remove(destino)
        return (t, "erro", erro)

    ok = pulado = erro = 0
    resultados = []
    with ThreadPoolExecutor(max_workers=THREADS) as ex:
        futuros = [ex.submit(baixar, t) for t in tarefas]
        for i, fut in enumerate(as_completed(futuros), 1):
            t, status, msg = fut.result()
            resultados.append((t, status, msg))
            if status == "ok": ok += 1
            elif status == "pulado": pulado += 1
            else:
                erro += 1
                print(f"    ERRO {t['arquivo']}: {msg}")
            if i % 200 == 0:
                print(f"    {i}/{len(tarefas)}  (ok={ok} pulados={pulado} erros={erro})")
    print(f"  Download: ok={ok} | pulados={pulado} | erros={erro}")

    # gera linkagem_fotos.xlsx
    wb2 = Workbook()
    wsl = wb2.active
    wsl.title = "Linkagem"
    cols = ["Codigo cadastro", "Code Deep", "Cidade", "Endereco", "Bairro",
            "Tipo de adesao", "Status", "Tipo da foto", "Arquivo",
            "Abrir foto", "Link original"]
    wsl.append(cols)
    azul = PatternFill("solid", fgColor="1F4E78")
    for c in range(1, len(cols) + 1):
        cell = wsl.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = azul

    r_out = 2
    for t, status, msg in resultados:
        if status == "erro":
            continue
        wsl.append([t["codigo"], t["code_deep"], t["cidade"], t["endereco"],
                    t["bairro"], t["tipo_adesao"], t["status"], t["tipo_foto"],
                    t["arquivo"], "", t["url"]])
        cell = wsl.cell(row=r_out, column=10)
        cell.value = "abrir"
        cell.hyperlink = f"{SUBPASTA_FOTOS}/{t['arquivo']}"
        cell.font = Font(color="0563C1", underline="single")
        r_out += 1

    for col, w in {"A": 16, "B": 18, "C": 16, "D": 28, "E": 16, "F": 16,
                   "G": 16, "H": 28, "I": 42, "J": 10, "K": 60}.items():
        wsl.column_dimensions[col].width = w
    wsl.freeze_panes = "A2"
    if r_out > 2:
        wsl.auto_filter.ref = f"A1:K{r_out-1}"
    wb2.save(os.path.join(pasta_saida, "linkagem_fotos.xlsx"))

    falhas = [(t, msg) for t, s, msg in resultados if s == "erro"]
    if falhas:
        with open(os.path.join(pasta_saida, "fotos_com_erro.csv"), "w",
                  newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["codigo", "tipo_foto", "arquivo", "url", "erro"])
            for t, msg in falhas:
                w.writerow([t["codigo"], t["tipo_foto"], t["arquivo"], t["url"], msg])

        # ---- lista das faltas AGRUPADA por codigo de cadastro interno ----
        # (404 = foto nao existe no servidor). Serve pra cobrar recaptura.
        por_codigo = {}
        for t, msg in falhas:
            d = por_codigo.setdefault(t["codigo"], {
                "code_deep": t["code_deep"], "cidade": t["cidade"],
                "endereco": t["endereco"], "tipos": [], "n": 0})
            d["n"] += 1
            d["tipos"].append(t["tipo_foto"])

        wb3 = Workbook()
        wsf = wb3.active
        wsf.title = "Faltas por codigo"
        cabf = ["Codigo cadastro", "Code Deep", "Cidade", "Endereco",
                "Qtd fotos faltando", "Tipos de foto faltando"]
        wsf.append(cabf)
        vermelho = PatternFill("solid", fgColor="C00000")
        for c in range(1, len(cabf) + 1):
            cell = wsf.cell(row=1, column=c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = vermelho
        for cod in sorted(por_codigo):
            d = por_codigo[cod]
            # conta por tipo quando ha repeticao (ex.: "Foto X (2)")
            cont = {}
            for tp in d["tipos"]:
                cont[tp] = cont.get(tp, 0) + 1
            tipos_txt = "; ".join(
                tp if n == 1 else f"{tp} ({n})" for tp, n in sorted(cont.items()))
            wsf.append([cod, d["code_deep"], d["cidade"], d["endereco"],
                        d["n"], tipos_txt])
        for col, w in {"A": 18, "B": 18, "C": 16, "D": 28,
                       "E": 18, "F": 60}.items():
            wsf.column_dimensions[col].width = w
        wsf.freeze_panes = "A2"
        wsf.auto_filter.ref = f"A1:F{len(por_codigo)+1}"
        wb3.save(os.path.join(pasta_saida, "faltas_por_codigo.xlsx"))
        print(f"  -> {len(falhas)} fotos faltando em {len(por_codigo)} cadastros "
              f"(ver faltas_por_codigo.xlsx)")

    print(f"  -> {pasta_saida}  ({ok+pulado} fotos, {r_out-1} linhas na linkagem)")


def main():
    os.makedirs(PASTA_INPUT, exist_ok=True)
    os.makedirs(PASTA_OUTPUT, exist_ok=True)

    planilhas = [p for p in glob.glob(os.path.join(PASTA_INPUT, "*.xlsx"))
                 if not os.path.basename(p).startswith("~$")]
    if not planilhas:
        print(f"Nenhuma planilha .xlsx encontrada em '{PASTA_INPUT}/'.")
        print("Coloque a(s) planilha(s) la dentro e rode de novo.")
        return

    print(f"Planilhas encontradas em '{PASTA_INPUT}/': {len(planilhas)}")
    for caminho in planilhas:
        try:
            processar_planilha(caminho)
        except Exception as e:
            print(f"  ERRO ao processar {caminho}: {type(e).__name__}: {e}")

    print("\nConcluido. Veja os resultados na pasta 'output/'.")


if __name__ == "__main__":
    main()
