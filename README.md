# Conversor e Unificador PPU — HC4

Aplicação web para converter, correlacionar e unificar documentos PPU diretamente no navegador.

https://pedroandradeoriginal.github.io/converte_unifica/

O sistema identifica automaticamente arquivos correspondentes pelo nome, converte documentos Excel e Word para PDF e gera um único arquivo para cada correlação encontrada.

> Desenvolvido por **Pedro Andrade** para o Consórcio HC4.

## Funcionalidades

- Importação simultânea de vários arquivos.
- Compatibilidade com `.xlsx`, `.xlsm`, `.docx` e `.pdf`.
- Correlação automática pelo nome dos documentos.
- Normalização de pequenas diferenças nos nomes, incluindo:
  - `_` no início do arquivo;
  - espaços adicionais;
  - diferenças entre `GI` e `G1`;
  - diferenças entre os prefixos `5135` e `5950`.
- Conversão local de Excel e Word para PDF.
- Formatação de datas do Excel em `dia/mês/ano`.
- União dos documentos na ordem obrigatória:
  1. Excel;
  2. Word.
- Download individual dos PDFs gerados.
- Download de todos os resultados em um arquivo ZIP.
- Interface responsiva para computadores, tablets e celulares.

## Privacidade

Todo o processamento acontece localmente no navegador.

- Nenhum documento é enviado para servidores.
- Nenhuma informação é armazenada externamente.
- Os arquivos são descartados ao fechar ou atualizar a página.
- Os originais não são alterados.

## Como usar

1. Abra o site pelo navegador.
2. Arraste os arquivos para a área indicada ou clique para selecioná-los.
3. Confira as correlações encontradas.
4. Clique em **Converter e unificar PDFs**.
5. Baixe cada PDF separadamente ou selecione **Baixar tudo em ZIP**.

## Publicação no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie todo o conteúdo da pasta `dist` para a raiz do repositório.
3. No repositório, acesse **Settings → Pages**.
4. Em **Build and deployment**, selecione **Deploy from a branch**.
5. Escolha a branch `main`, a pasta `/ (root)` e clique em **Save**.
6. Aguarde o GitHub disponibilizar o endereço do site.

## Execução local

Para abrir sem publicar, utilize um servidor HTTP local dentro da pasta `dist`.

Exemplo com Python:

```bash
python -m http.server 8000
```

Depois, acesse `http://localhost:8000` no navegador.

## Estrutura do pacote

```text
dist/
├── index.html
├── README.md
├── logo-hc4.png
├── src/
│   ├── app.js
│   ├── styles.css
│   └── logo-hc4.png
└── vendor/
    └── bibliotecas utilizadas pelo processamento local
```

## Observação sobre a conversão

A conversão de arquivos Office é feita pelo navegador. Para documentos que exigem reprodução absolutamente idêntica à impressão do Microsoft Office, recomenda-se utilizar os PDFs previamente convertidos.

## Tecnologias

- HTML5
- CSS3
- JavaScript
- SheetJS
- Mammoth.js
- jsPDF
- html2canvas
- PDF-Lib
- JSZip

## Autor

**Pedro Andrade**  
Consórcio HC4 — Planejamento
