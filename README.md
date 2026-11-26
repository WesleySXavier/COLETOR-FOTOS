
  ENTREGA SABESP - baixar fotos e gerar linkagem


O QUE ESTE PROJETO FAZ
----------------------
Le as planilhas que voce colocar na pasta "input", baixa TODAS as fotos
dos links e gera, na pasta "output", uma pasta por planilha contendo:
  - fotos/              -> todas as imagens (nomeadas por codigo + tipo da foto)
  - linkagem_fotos.xlsx -> uma linha por foto, ligando o arquivo ao codigo e
                           ao tipo, com um link que abre a foto da propria pasta
  - fotos_com_erro.csv  -> (so se algo falhar) lista o que nao baixou

No fim e so zipar a pasta de output daquela planilha e enviar pra Sabesp.
Como os links sao relativos (apontam pra pasta "fotos"), nao depende de
internet nem de link externo.

PASSO A PASSO
-------------
1. Instale o Python 3 (se ainda nao tiver): https://www.python.org/downloads/
   Na instalacao, marque a opcao "Add Python to PATH".

2. Abra o Prompt de Comando (CMD) nesta pasta e instale as dependencias:
       pip install openpyxl requests

3. Coloque a(s) planilha(s) .xlsx dentro da pasta "input".
   (Elas devem seguir os cabecalhos do arquivo template_planilha.xlsx.)

4. Rode:
       python entrega_sabesp.py

5. Aguarde. Sao muitas fotos (~20 mil na planilha de exemplo), entao demora.
   Pode fechar e rodar de novo depois: ele PULA o que ja baixou e continua
   de onde parou.

6. Veja o resultado na pasta "output".

OBSERVACOES
-----------
- Rode na sua maquina / rede, que tem acesso aos links das fotos.
- O codigo acha as colunas pelo NOME do cabecalho (nao pela letra), entao a
  ordem das colunas pode variar, desde que os cabecalhos sejam os mesmos do
  template.
- Se algum cabecalho de foto nao for encontrado, ele avisa no terminal.

ESTRUTURA DA PASTA
------------------
  entrega_sabesp.py        -> o programa
  template_planilha.xlsx   -> modelo vazio (so cabecalhos) pra guiar
  input/                   -> planilhas entram aqui
  output/                  -> resultados saem aqui
  LEIA-ME.txt              -> este arquivo
