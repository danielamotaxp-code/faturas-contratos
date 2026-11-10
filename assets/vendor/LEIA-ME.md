# Bibliotecas de terceiros

- `xlsx.bundle.js` e `cpexcel.js`: [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style), licença Apache 2.0 (`LICENSE-xlsx-js-style`).

**Ajuste local em `xlsx.bundle.js`:** quando a aba tem `ws["!freeze"] = N`, o arquivo gerado congela as N primeiras linhas, o que no Excel corresponde a "Congelar painéis". Na biblioteca original, o trecho era `Yt("sheetViews",Yt("sheetView",null,o),{})`. Se a biblioteca for atualizada, reaplique esse ajuste.
