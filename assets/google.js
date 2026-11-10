/**
 * Integração com o Google Planilhas.
 *
 * - Leitura: com a pessoa conectada pelo "Entrar com Google" (token), ou com a chave de API
 *   quando a planilha está aberta para "qualquer pessoa com o link".
 * - Escrita (só com "Entrar com Google"):
 *     Faturas Pendentes   → pinta a aba "Semana N" da própria planilha (azul / branco).
 *     Contratos sem Faturas → cria a aba "Contratos sem Fatura – Semana N" na própria planilha.
 */
(function (raiz) {
  'use strict';

  var API = 'https://sheets.googleapis.com/v4/spreadsheets/';
  var ESCOPO = 'https://www.googleapis.com/auth/spreadsheets';
  var sessao = { token: null, expira: 0, cliente: null, clientId: '' };

  /* ---------------- login com Google ---------------- */

  function conectado() { return !!sessao.token && Date.now() < sessao.expira - 60000; }

  /** Abre a janela do Google para autorizar. Precisa ser chamado a partir de um clique. */
  function entrar(clientId) {
    return new Promise(function (ok, falha) {
      if (!clientId) return falha(new Error('O "Entrar com Google" ainda não foi configurado. Peça para a administradora salvar o ID do cliente em Configurações.'));
      if (!raiz.google || !raiz.google.accounts || !raiz.google.accounts.oauth2) return falha(new Error('Não consegui carregar o login do Google. Recarregue a página.'));
      if (!sessao.cliente || sessao.clientId !== clientId) {
        sessao.clientId = clientId;
        sessao.cliente = raiz.google.accounts.oauth2.initTokenClient({ client_id: clientId, scope: ESCOPO, callback: function () {} });
      }
      sessao.cliente.callback = function (resp) {
        if (resp.error) return falha(new Error('O login com Google não foi concluído (' + resp.error + ').'));
        sessao.token = resp.access_token;
        sessao.expira = Date.now() + Number(resp.expires_in || 3600) * 1000;
        ok(true);
      };
      sessao.cliente.error_callback = function (e) { falha(new Error('O login com Google foi fechado ou bloqueado (' + ((e && e.type) || 'erro') + ').')); };
      sessao.cliente.requestAccessToken({ prompt: '' });
    });
  }

  function sair() {
    if (sessao.token && raiz.google && raiz.google.accounts && raiz.google.accounts.oauth2) {
      try { raiz.google.accounts.oauth2.revoke(sessao.token, function () {}); } catch (e) { /* ignora */ }
    }
    sessao.token = null; sessao.expira = 0;
  }

  /* ---------------- chamadas à API ---------------- */

  function idDoLink(link) {
    var m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/.exec(String(link || ''));
    if (!m) throw new Error('Esse não parece um link do Google Planilhas. Ele começa com https://docs.google.com/spreadsheets/d/...');
    return m[1];
  }

  function erroGoogle(r, escrevendo) {
    return r.json().catch(function () { return {}; }).then(function (j) {
      var msg = (j.error && j.error.message) || '';
      if (r.status === 401) { sessao.token = null; throw new Error('O login com Google expirou. Clique em "Entrar com Google" de novo.'); }
      if (r.status === 403 && /API key|referer|referrer|blocked/i.test(msg)) {
        throw new Error('A chave do Google recusou o pedido. Confira se ela está liberada para a Google Sheets API e para o site danielamotaxp-code.github.io.');
      }
      if (r.status === 403) {
        throw new Error(escrevendo
          ? 'A conta Google conectada não tem permissão para editar essa planilha. Entre com a conta que edita as planilhas.'
          : (conectado() ? 'A conta Google conectada não tem acesso a essa planilha.'
            : 'A planilha não está aberta para leitura. Clique em "Entrar com Google" ou deixe a planilha em Compartilhar → "Qualquer pessoa com o link" (Leitor).'));
      }
      if (r.status === 404) throw new Error('Planilha não encontrada. Confira o link.');
      if (r.status === 400 && /API key not valid/i.test(msg)) throw new Error('A chave de API do Google é inválida. Atualize em Configurações.');
      throw new Error('O Google respondeu com erro (' + r.status + '). ' + msg);
    });
  }

  function chamar(url, opcoes, chave, escrevendo) {
    var o = opcoes || {};
    o.headers = Object.assign({}, o.headers || {});
    if (conectado()) o.headers.Authorization = 'Bearer ' + sessao.token;
    else if (chave) url += (url.indexOf('?') >= 0 ? '&' : '?') + 'key=' + encodeURIComponent(chave);
    else return Promise.reject(new Error('Clique em "Entrar com Google" para ler a planilha pelo link.'));
    if (o.body) o.headers['Content-Type'] = 'application/json';
    return fetch(url, o).then(function (r) { return r.ok ? r.json() : erroGoogle(r, escrevendo); });
  }

  function aspas(nome) { return "'" + String(nome).replace(/'/g, "''") + "'"; }

  function metadados(id, chave) {
    return chamar(API + id + '?fields=' + encodeURIComponent('properties.title,sheets(properties(title,sheetId,gridProperties),bandedRanges(bandedRangeId))'), {}, chave);
  }

  /** Lê a planilha inteira: { id, titulo, link, abas: { nome: [[...]] }, nomesAbas, folhas: { nome: {sheetId, bandas} } } */
  function lerPlanilha(link, chave) {
    var id;
    try { id = idDoLink(link); } catch (e) { return Promise.reject(e); }
    return metadados(id, chave).then(function (meta) {
      var nomes = [], folhas = {};
      (meta.sheets || []).forEach(function (s) {
        nomes.push(s.properties.title);
        folhas[s.properties.title] = { sheetId: s.properties.sheetId,
          bandas: (s.bandedRanges || []).map(function (b) { return b.bandedRangeId; }) };
      });
      if (!nomes.length) throw new Error('A planilha não tem abas.');
      var q = nomes.map(function (n) { return 'ranges=' + encodeURIComponent(aspas(n)); }).join('&');
      return chamar(API + id + '/values:batchGet?' + q + '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING&majorDimension=ROWS', {}, chave)
        .then(function (v) {
          var abas = {};
          (v.valueRanges || []).forEach(function (vr, i) { abas[nomes[i]] = vr.values || []; });
          return { id: id, titulo: meta.properties.title, link: 'https://docs.google.com/spreadsheets/d/' + id + '/edit',
            abas: abas, nomesAbas: nomes, folhas: folhas };
        });
    });
  }

  function corRGB(hex) {
    return { red: parseInt(hex.slice(0, 2), 16) / 255, green: parseInt(hex.slice(2, 4), 16) / 255, blue: parseInt(hex.slice(4, 6), 16) / 255 };
  }

  function lote(id, pedidos) {
    // envia em partes para não passar do limite de tamanho do Google
    var partes = [];
    for (var i = 0; i < pedidos.length; i += 4000) partes.push(pedidos.slice(i, i + 4000));
    return partes.reduce(function (p, parte) {
      return p.then(function () {
        return chamar(API + id + ':batchUpdate', { method: 'POST', body: JSON.stringify({ requests: parte }) }, null, true);
      });
    }, Promise.resolve());
  }

  /**
   * Pinta as linhas de uma aba. "indices" são as linhas reais da planilha (0 = cabeçalho)
   * e "cores" a cor de cada uma (hex sem #). Remove as cores alternadas (zebra) da aba, se houver.
   */
  function pintarAba(planilha, nomeAba, colunas, indices, cores) {
    var folha = planilha.folhas[nomeAba];
    if (!folha) return Promise.reject(new Error('Não encontrei a aba "' + nomeAba + '" na planilha do Google.'));
    var pedidos = folha.bandas.map(function (b) { return { deleteBanding: { bandedRangeId: b } }; });
    // cabeçalho fixo ao rolar
    pedidos.push({ updateSheetProperties: { properties: { sheetId: folha.sheetId, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });
    var i = 0;
    while (i < indices.length) {
      var j = i;
      while (j + 1 < indices.length && indices[j + 1] === indices[j] + 1 && cores[j + 1] === cores[i]) j++;
      pedidos.push({ repeatCell: {
        range: { sheetId: folha.sheetId, startRowIndex: indices[i], endRowIndex: indices[j] + 1, startColumnIndex: 0, endColumnIndex: colunas },
        cell: { userEnteredFormat: { backgroundColor: corRGB(cores[i]) } },
        fields: 'userEnteredFormat.backgroundColor'
      } });
      i = j + 1;
    }
    return lote(planilha.id, pedidos);
  }

  /** Cria (ou recria) uma aba com os valores informados, como texto exato. */
  function escreverAba(planilha, titulo, valores) {
    var id = planilha.id, existente = planilha.folhas[titulo];
    var pedidos = [];
    if (existente) pedidos.push({ deleteSheet: { sheetId: existente.sheetId } });
    var linhas = Math.max(valores.length, 2), colunas = Math.max((valores[0] || []).length, 1);
    pedidos.push({ addSheet: { properties: { title: titulo, gridProperties: { rowCount: linhas, columnCount: colunas, frozenRowCount: 1 } } } });
    return chamar(API + id + ':batchUpdate', { method: 'POST', body: JSON.stringify({ requests: pedidos }) }, null, true)
      .then(function (resp) {
        var novo = resp.replies[resp.replies.length - 1].addSheet.properties.sheetId;
        var texto = valores.map(function (l) { return l.map(function (c) { return c == null ? '' : (c instanceof Date ? c.toLocaleDateString('pt-BR') : c); }); });
        return chamar(API + id + '/values/' + encodeURIComponent(aspas(titulo) + '!A1') + '?valueInputOption=RAW',
          { method: 'PUT', body: JSON.stringify({ values: texto }) }, null, true).then(function () { return novo; });
      })
      .then(function (sheetId) {
        return lote(id, [{ repeatCell: {
          range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: colunas },
          cell: { userEnteredFormat: { backgroundColor: corRGB('1F4E79'), textFormat: { bold: true, foregroundColor: corRGB('FFFFFF') } } },
          fields: 'userEnteredFormat(backgroundColor,textFormat)'
        } }]);
      });
  }

  var api = { lerPlanilha: lerPlanilha, idDoLink: idDoLink, entrar: entrar, sair: sair, conectado: conectado,
    pintarAba: pintarAba, escreverAba: escreverAba };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Google = api;
})(this);
