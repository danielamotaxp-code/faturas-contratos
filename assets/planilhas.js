/**
 * Leitura das planilhas enviadas e geração das planilhas novas (.xlsx com cores),
 * tudo dentro do navegador.
 */
(function (raiz) {
  'use strict';
  var R = raiz.Regras || require('./regras.js');
  var X = raiz.XLSX || require('./vendor/xlsx.bundle.js');

  var COR_CABECALHO = '1F4E79';
  var BORDA = { style: 'thin', color: { rgb: 'D9DEE5' } };

  /** Lê o arquivo enviado (.xlsx, .xls, .ods ou .csv). */
  function lerArquivo(nome, buffer) {
    if (/\.csv$/i.test(nome)) {
      var txt = new TextDecoder('utf-8').decode(buffer);
      var sep = (txt.split('\n')[0].match(/;/g) || []).length > (txt.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
      return X.read(txt, { type: 'string', raw: true, FS: sep });
    }
    return X.read(buffer, { type: 'array', cellDates: true });
  }

  function dadosDaAba(wb, nome) {
    var ws = wb.Sheets[nome];
    if (!ws || !ws['!ref']) return { cab: [], linhas: [] };
    var aoa = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    var r0 = X.utils.decode_range(ws['!ref']).s.r;
    var cab = (aoa[0] || []).map(function (c) { return String(c == null ? '' : c).trim(); });
    while (cab.length && !cab[cab.length - 1]) cab.pop();
    var w = cab.length, linhas = [], indices = [];
    for (var i = 1; i < aoa.length; i++) {
      var l = aoa[i] || [];
      if (!l.some(function (c) { return c !== '' && c != null && String(c).trim() !== ''; })) continue;
      l = l.slice(0, w);
      while (l.length < w) l.push('');
      linhas.push(l);
      indices.push(r0 + i); // linha real na planilha (0 = cabeçalho)
    }
    return { cab: cab, linhas: linhas, indices: indices };
  }

  function estiloCabecalho() {
    return { fill: { patternType: 'solid', fgColor: { rgb: COR_CABECALHO } },
      font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { vertical: 'center' },
      border: { top: BORDA, bottom: BORDA, left: BORDA, right: BORDA } };
  }

  /** Deixa CPF/CNPJ completos e instalação / unidade consumidora como texto, para nada se perder. */
  function normalizarCodigos(cab, linhas) {
    var docs = [], codigos = [];
    cab.forEach(function (h, i) {
      var n = R.normalizar(h);
      if (/cpf|cnpj/.test(n)) docs.push(i);
      else if (/instala|unid.*consum|unidade\s*consumidora|^uc$/.test(n)) codigos.push(i);
    });
    if (!docs.length && !codigos.length) return linhas;
    return linhas.map(function (l) {
      var c = l.slice();
      docs.forEach(function (i) { if (c[i] !== '' && c[i] != null) c[i] = R.corrigirDocumento(c[i]); });
      codigos.forEach(function (i) { if (c[i] !== '' && c[i] != null) c[i] = R.codigoTexto(c[i]); });
      return c;
    });
  }

  function abaDeDados(cab, linhas, cores) {
    linhas = normalizarCodigos(cab, linhas);
    var ws = X.utils.aoa_to_sheet([cab].concat(linhas), { cellDates: true });
    var nLin = linhas.length + 1, nCol = cab.length;
    for (var r = 0; r < nLin; r++) {
      for (var c = 0; c < nCol; c++) {
        var end = X.utils.encode_cell({ r: r, c: c });
        var cel = ws[end];
        if (!cel) { if (r === 0 || cores) { cel = ws[end] = { t: 's', v: '' }; } else continue; }
        if (r === 0) { cel.s = estiloCabecalho(); continue; }
        if (cel.t === 'd') cel.z = 'dd/mm/yyyy';
        if (cel.t === 'n' && Number.isInteger(cel.v) && Math.abs(cel.v) >= 1e10) cel.z = '0';
        if (cores) {
          cel.s = { fill: { patternType: 'solid', fgColor: { rgb: cores[r - 1] } },
            border: { top: BORDA, bottom: BORDA, left: BORDA, right: BORDA } };
        }
      }
    }
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(nLin - 1, 0), c: Math.max(nCol - 1, 0) } });
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!cols'] = cab.map(function (h) { return { wch: Math.min(38, Math.max(12, String(h).length + 3)) }; });
    return ws;
  }

  function abaSimples(cab, linhas) {
    var ws = X.utils.aoa_to_sheet([cab].concat(linhas));
    cab.forEach(function (h, c) { ws[X.utils.encode_cell({ r: 0, c: c })].s = estiloCabecalho(); });
    ws['!cols'] = cab.map(function (h, c) {
      var max = String(h).length;
      linhas.forEach(function (l) { max = Math.max(max, String(l[c] == null ? '' : l[c]).length); });
      return { wch: Math.min(70, max + 2) };
    });
    return ws;
  }

  function gravar(wb) {
    return new Uint8Array(X.write(wb, { bookType: 'xlsx', type: 'array', compression: true, cellDates: false }));
  }

  /** ABA 1: gera "Faturas Pendentes – <Mês> – Semana N". */
  function gerarFaturas(wb, mes, semanaManual) {
    var semanas = R.semanasFaturas(wb.SheetNames, semanaManual);
    var abas = [], resultado = [], google = null;
    semanas.forEach(function (s) {
      var d = dadosDaAba(wb, s.aba);
      var f = R.filtrarUltimosMeses(d.cab, d.linhas, d.indices, mes.chave);
      if (!f.linhas.length) throw new Error('A aba "' + s.aba + '" não tem faturas de ' + f.janela.map(R.rotuloMes).join(', ') + '.');
      var a = R.analisarFaturas(d.cab, f.linhas, mes.chave);
      google = { aba: s.aba, colunas: d.cab.length, indices: f.indices, cores: a.cores };
      abas.push(['Semana ' + s.n, abaDeDados(d.cab, f.linhas, a.cores)]);
      resultado.push({ n: s.n, total: a.total, comFatura: a.comFatura, pendentes: a.pendentes,
        linhas: f.linhas.length, removidas: f.removidas, janela: f.janela });
    });
    var ult = resultado[resultado.length - 1];
    var out = X.utils.book_new();
    var meses = ult.janela.map(R.rotuloMes).join(', ');
    X.utils.book_append_sheet(out, abaSimples(['Semana', 'Meses considerados', 'Linhas (faturas)', 'Instalações', 'Com fatura de ' + mes.rotulo, 'Sem fatura (gerar)', 'Linhas de outros meses (fora do filtro)'],
      resultado.map(function (s) { return ['Semana ' + s.n, meses, s.linhas, s.total, s.comFatura, s.pendentes.length, s.removidas]; })), 'Resumo');
    abas.forEach(function (a) { X.utils.book_append_sheet(out, a[1], a[0]); });
    X.utils.book_append_sheet(out, abaSimples(['Instalação', 'Cliente', 'CPF/CNPJ', 'Detalhes'],
      ult.pendentes.map(function (p) { return [p.instalacao, p.cliente, p.cpfCnpj, p.info]; })), 'Pendentes – Semana ' + ult.n);
    return { nomeGerado: 'Faturas Pendentes – ' + mes.rotulo + ' – Semana ' + ult.n, bytes: gravar(out), semanas: resultado, google: google };
  }

  /** ABA 2: gera "Contratos sem Faturas – <Mês> – Semana N". */
  function gerarContratos(wb, mes, semanaManual) {
    var semanas = R.semanasContratos(wb.SheetNames, semanaManual);
    var abas = [], resultado = [], google = null;
    semanas.forEach(function (s) {
      var c = dadosDaAba(wb, s.c), f = dadosDaAba(wb, s.f);
      var a = R.analisarContratos(c.cab, c.linhas, f.cab, f.linhas, mes.chave, s.c);
      google = { titulo: 'Contratos sem Fatura – Semana ' + s.n,
        valores: [c.cab].concat(normalizarCodigos(c.cab, a.indices.map(function (i) { return c.linhas[i]; }))) };
      abas.push(['Semana ' + s.n, abaDeDados(c.cab, a.indices.map(function (i) { return c.linhas[i]; }), null)]);
      resultado.push({ n: s.n, total: a.total, comFatura: a.comFatura, pendentes: a.pendentes, faturasDoMes: a.faturasDoMes });
    });
    var ult = resultado[resultado.length - 1];
    var out = X.utils.book_new();
    X.utils.book_append_sheet(out, abaSimples(['Semana', 'Contratos vinculados', 'Com fatura de ' + mes.rotulo, 'Sem fatura', 'Faturas de ' + mes.rotulo + ' no gateway'],
      resultado.map(function (s) { return ['Semana ' + s.n, s.total, s.comFatura, s.pendentes.length, s.faturasDoMes]; })), 'Resumo');
    abas.forEach(function (a) { X.utils.book_append_sheet(out, a[1], a[0]); });
    return { nomeGerado: 'Contratos sem Faturas – ' + mes.rotulo + ' – Semana ' + ult.n, bytes: gravar(out), semanas: resultado, google: google };
  }

  /** Processa uma planilha já aberta. "nome" é o nome do arquivo ou o título da planilha do Google (de onde sai o mês). */
  function processarLivro(aba, nome, wb, semanaManual) {
    var mes = R.mesDoNome(nome);
    if (!mes) throw new Error('Não identifiquei o mês no nome "' + nome + '". O nome da planilha precisa ter o mês e o ano, por exemplo "OUT 2026".');
    var r = aba === 'faturas' ? gerarFaturas(wb, mes, semanaManual) : gerarContratos(wb, mes, semanaManual);
    r.mes = mes;
    return r;
  }

  /** Processa o arquivo enviado. */
  function processar(aba, nomeArquivo, buffer, semanaManual) {
    return processarLivro(aba, nomeArquivo, lerArquivo(nomeArquivo, buffer), semanaManual);
  }

  /** Monta uma planilha a partir das abas lidas do Google (cópia fiel, guardada como "planilha enviada"). */
  function livroDeAbas(nomes, abas) {
    var wb = X.utils.book_new();
    nomes.forEach(function (n) {
      var ws = X.utils.aoa_to_sheet(abas[n] && abas[n].length ? abas[n] : [[]]);
      X.utils.book_append_sheet(wb, ws, n.slice(0, 31));
    });
    return wb;
  }

  var api = { processar: processar, processarLivro: processarLivro, livroDeAbas: livroDeAbas, gravar: gravar,
    lerArquivo: lerArquivo, dadosDaAba: dadosDaAba };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Planilhas = api;
})(this);
