# Entrega Sabesp — Download de fotos e linkagem

Ferramenta que lê planilhas de cadastro, baixa as fotos a partir dos links e
gera uma pasta de entrega com as imagens + uma planilha de linkagem
(`linkagem_fotos.xlsx`) que referencia cada foto ao código do cadastro e ao
tipo da foto, usando links locais (relativos) — sem depender de links externos.

## ⚠️ Aviso de LGPD (leia antes de usar ou publicar)

Este repositório contém **apenas código**. Ele processa **dados pessoais e
dados pessoais sensíveis** (nome, CPF, RG, CNPJ, endereço, fotos de documentos).
Por isso:

- **Nunca** faça commit de planilhas com dados reais, fotos, arquivos da pasta
  `output/`, `input/` ou qualquer `.csv`/`.zip` gerado. O `.gitignore` já
  bloqueia isso — não remova essas regras.
- Use preferencialmente um **repositório privado**. Dados pessoais não devem
  ficar públicos.
- Trate as planilhas e fotos sob os princípios da LGPD: **finalidade**
  (uso exclusivo para a entrega contratada à Sabesp), **necessidade/minimização**
  (baixe só o necessário) e **segurança** (armazene em local restrito, com
  acesso controlado).
- Ao final do uso/retenção, **elimine** os dados locais (pasta `output/`,
  planilhas e fotos) conforme a política de retenção acordada.
- Se algum dado real já tiver sido commitado por engano, ele permanece no
  histórico do Git — será preciso **limpar o histórico** (ex.: `git filter-repo`)
  e, se já foi enviado, considerar o dado como exposto.

> O `template_planilha.xlsx` versionado contém **apenas os cabeçalhos** (linha 1),
> sem nenhum dado pessoal.

## Como funciona

1. Coloque a(s) planilha(s) `.xlsx` na pasta `input/` (seguindo os cabeçalhos
   do `template_planilha.xlsx`).
2. Rode o script.
3. Para cada planilha, é criada em `output/<nome-da-planilha>/`:
   - `fotos/` — todas as imagens, nomeadas por `codigo__tipo-da-foto`
   - `linkagem_fotos.xlsx` — uma linha por foto, com link clicável relativo
   - `fotos_com_erro.csv` — (se houver) fotos que não existem no servidor
   - `faltas_por_codigo.xlsx` — (se houver) faltas agrupadas por código interno

## Requisitos

- Python 3.9+
- Instale as dependências:

```bash
pip install -r requirements.txt
```

## Uso

```bash
python entrega_sabesp.py
```

O script identifica as colunas **pelo nome do cabeçalho** (não pela letra) e
pode ser interrompido e retomado — ele **pula o que já foi baixado**.

## Observações técnicas

- Uma mesma célula pode conter **várias fotos** (separadas por quebra de linha);
  o script separa e baixa cada uma.
- Execute em uma rede com acesso ao servidor de imagens.
