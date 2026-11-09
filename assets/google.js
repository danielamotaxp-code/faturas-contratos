/**
 * Leitura de planilhas do Google Planilhas pelo link.
 * A planilha precisa estar compartilhada como "Qualquer pessoa com o link – Leitor".
 * Usa a API oficial do Google Sheets com uma chave de API (restrita a este site),
 * o que garante o nome da planilha, o nome de todas as abas e os valores exatos.
 */
(function (raiz) {
  'use strict';

  var API = 'https://sheets.googleapis.com/v4/spreadsheets/';

  function idDoLink(link) {
    var m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/.exec(String(link || ''));
    if (!m) throw new Error('Esse não parece um link do Google Planilhas. Ele começa com https://docs.google.com/spreadsheets/d/...');
    return m[1];
  }

  function erroGoogle(r) {
    return r.json().catch(function () { return {}; }).then(function (j) {
      var msg = (j.error && j.error.message) || '';
      if (r.status === 403 && /API key|referer|referrer|blocked/i.test(msg)) {
        throw new Error('A chave do Google recusou o pedido. Confira se a chave está liberada para a Google Sheets API e para o site danielamotaxp-code.github.io.');
      }
      if (r.status === 403 || r.status === 401) {
        throw new Error('A planilha não está aberta para leitura. No Google Planilhas, clique em Compartilhar → Acesso geral → "Qualquer pessoa com o link" (Leitor) e tente de novo.');
      }
      if (r.status === 404) throw new Error('Planilha não encontrada. Confira o link.');
      if (r.status === 400 && /API key not valid/i.test(msg)) throw new Error('A chave de API do Google é inválida. Atualize em Configurações.');
      throw new Error('O Google respondeu com erro (' + r.status + '). ' + msg);
    });
  }

  function aspas(nome) { return "'" + String(nome).replace(/'/g, "''") + "'"; }

  /**
   * Lê a planilha inteira: { titulo, link, abas: { nomeDaAba: [[linha], ...] }, nomesAbas: [...] }
   */
  function lerPlanilha(link, chave) {
    if (!chave) return Promise.reject(new Error('A leitura por link ainda não foi configurada. Peça para a administradora salvar a chave do Google em Configurações.'));
    var id = idDoLink(link);
    var base = API + id;
    return fetch(base + '?key=' + encodeURIComponent(chave) + '&fields=' + encodeURIComponent('properties.title,sheets.properties(title,sheetId,hidden)'))
      .then(function (r) { return r.ok ? r.json() : erroGoogle(r); })
      .then(function (meta) {
        var nomes = (meta.sheets || []).map(function (s) { return s.properties.title; });
        if (!nomes.length) throw new Error('A planilha não tem abas.');
        var q = nomes.map(function (n) { return 'ranges=' + encodeURIComponent(aspas(n)); }).join('&');
        var url = base + '/values:batchGet?' + q + '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING&majorDimension=ROWS&key=' + encodeURIComponent(chave);
        return fetch(url).then(function (r) { return r.ok ? r.json() : erroGoogle(r); }).then(function (v) {
          var abas = {};
          (v.valueRanges || []).forEach(function (vr, i) { abas[nomes[i]] = vr.values || []; });
          return { titulo: meta.properties.title, link: 'https://docs.google.com/spreadsheets/d/' + id + '/edit', abas: abas, nomesAbas: nomes };
        });
      });
  }

  var api = { lerPlanilha: lerPlanilha, idDoLink: idDoLink };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Google = api;
})(this);
