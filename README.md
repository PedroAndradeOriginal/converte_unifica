# Conversor e Unificador PPU — HC4

Ferramenta web para converter documentos do Excel e do Word em PDF, correlacionar os arquivos pelo nome e gerar um PDF por PPU.

## Funcionamento

1. Selecione juntos os arquivos do Excel e do Word.
2. O site identifica automaticamente cada tipo pela extensão.
3. O site normaliza pequenas diferenças de nome, como `_` no início, espaços, `GI/G1` e `5135/5950`.
4. Cada documento é aberto pelo LibreOffice WebAssembly e exportado para PDF.
5. O resultado é unido sempre nesta ordem: **Excel primeiro, Word depois**.

O processo não reescreve células, datas, imagens ou estilos. Os documentos originais não são modificados e permanecem no computador do usuário.

## Publicação no GitHub Pages

1. Envie todos os arquivos deste projeto para a branch `main` do repositório.
2. No GitHub, abra **Settings → Pages**.
3. Em **Build and deployment → Source**, escolha **GitHub Actions**.
4. Abra a aba **Actions** e aguarde o fluxo **Publicar no GitHub Pages** terminar.

O fluxo baixa automaticamente o motor LibreOffice/ZetaOffice, compila o site e publica a pasta `dist`. A primeira publicação pode demorar por causa dos arquivos do motor Office.

## Uso local

Requer Node.js 20 ou superior.

```bash
npm install
npm run setup:office
npm run dev
```

## Privacidade

- Conversão inteiramente no navegador.
- Nenhum documento é enviado a servidor externo.
- Arquivos descartados ao atualizar ou fechar a página.
- O site só baixa os PDFs quando o usuário solicita.

## Tecnologias

- LibreOffice/ZetaOffice WebAssembly para abrir e exportar documentos.
- ZetaJS para controlar a exportação no navegador.
- PDF-Lib para unir os PDFs.
- JSZip para baixar vários resultados de uma vez.

Desenvolvido por **Pedro Andrade** para o Consórcio HC4.
