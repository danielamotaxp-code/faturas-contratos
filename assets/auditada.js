/*
 * Planilha auditada → lista de ouvidorias, sem depender de um modelo fixo.
 * Entende:
 *  - colunas com nomes diferentes (ex.: "Nome do cliente" = Cliente, "Unidade consumidora" = Unid. Consumidora);
 *  - a classificação em qualquer coluna (ex.: "Coluna1"/"Coluna2" com "Abrir ouvidoria", "Gerei fatura manual",
 *    "Cliente sem saldo de geração"), juntando as colunas quando há mais de uma;
 *  - cores das linhas, com a legenda da própria planilha (ex.: célula verde "Fatura gerada", ou "Verde = …").
 *    Sem legenda, a cor vira "Linha verde", "Linha vermelha"…
 * Só ficam as linhas classificadas; desconectados/inativos saem.
 */
(function (raiz) {
  'use strict';
  function nrm(x) { return String(x == null ? '' : x).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
  function texto(c) { return c == null ? '' : (c instanceof Date ? c.toLocaleDateString('pt-BR') : String(c).trim()); }

  /* ---------- cores ---------- */
  function hexDaCelula(cel) {
    var s = cel && cel.s; if (!s || s.patternType === 'none' || !s.fgColor) return '';
    var rgb = s.fgColor.rgb; if (!rgb || rgb.length < 6) return '';
    return rgb.slice(-6).toUpperCase();
  }
  function nomeCor(hex) {
    if (!hex) return '';
    var r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn, sat = mx ? d / mx : 0;
    if (sat < 0.15) return mx > 0.92 ? '' : 'cinza';
    var h = d === 0 ? 0 : mx === r ? 60 * (((g - b) / d) % 6) : mx === g ? 60 * ((b - r) / d + 2) : 60 * ((r - g) / d + 4);
    if (h < 0) h += 360;
    if (h < 15 || h >= 340) return 'vermelho';
    if (h < 42) return 'laranja';
    if (h < 68) return 'amarelo';
    if (h < 170) return 'verde';
    if (h < 255) return 'azul';
    return 'roxo';
  }
  var PALAVRA_COR = { verde: 'verde', vermelh: 'vermelho', amarel: 'amarelo', laranja: 'laranja', azul: 'azul', roxo: 'roxo', lilas: 'roxo', rosa: 'roxo', cinza: 'cinza' };
  var FEM = { vermelho: 'vermelha', amarelo: 'amarela', roxo: 'roxa' };

  /** Linhas (texto) e cores (nome) de uma aba do SheetJS lida com cellStyles. */
  function daAbaXlsx(ws, X) {
    var ref = ws['!ref']; if (!ref) return { m: [], cores: [] };
    var rg = X.utils.decode_range(ref), m = [], cores = [];
    for (var r = rg.s.r; r <= rg.e.r; r++) {
      var linha = [], cl = [];
      for (var c = rg.s.c; c <= rg.e.c; c++) {
        var cel = ws[X.utils.encode_cell({ r: r, c: c })];
        linha.push(cel ? texto(cel.w != null ? cel.w : cel.v) : '');
        cl.push(nomeCor(hexDaCelula(cel)));
      }
      m.push(linha); cores.push(cl);
    }
    return { m: m, cores: cores };
  }

  /* ---------- colunas ---------- */
  var SINONIMOS = {
    'Cliente': [/^cliente$/, /^nome( do)? (cliente|titular|consumidor)/, /^titular/, /razao social/, /^nome$/, /^cliente/],
    'Usina': [/^usina/, /nome( da)? usina/],
    'Instalação': [/^(n[o°º.]? ?)?(numero )?(da )?instalacao$/, /instalacao/],
    'Unid. Consumidora': [/unid.*consum/, /^uc$/],
    'Mês Referencia': [/mes( de)? refer/, /referencia/],
    'CPF / CNPJ': [/cpf|cnpj/],
    'Data Vencimento Boleto': [/vencimento/],
    'Parceiro': [/parceiro/],
    'Associado': [/associado/],
    'Status Contrato Faturamento': [/status( do)? contrato faturamento/],
    'Status Contrato Atual': [/status( do)? contrato( atual)?$/, /status( do)? contrato atual/],
    'Modalidade da Fatura': [/modalidade/],
    'Status da Fatura': [/status( da)? fatura/],
    'Geração de Faturas': [/geracao de faturas/],
    'Colaborador': [/colaborador/, /responsavel/]
  };
  function colunaDe(cab, nome) {
    var lista = SINONIMOS[nome] || [new RegExp('^' + nrm(nome) + '$')], n = cab.map(nrm);
    for (var p = 0; p < lista.length; p++) for (var i = 0; i < n.length; i++) if (lista[p].test(n[i])) return i;
    return -1;
  }
  var KW = /ouvidoria|saldo|manual|gerad|gerei|lancad|lancament|boleto|desconect|inativ|reclamac|sem fatura|fatura gerada/;
  var CAB_LIVRE = /^(coluna ?\d*|classifica.*|observ.*|obs\.?|motivo.*|auditoria.*|acao|a fazer|situacao( da auditoria)?|resultado|legenda)?$/;

  /**
   * m: matriz de textos; cores: matriz de nomes de cor (opcional).
   * Devolve { cab, linhas, classificadas, descartadas, legenda } com a coluna "Classificação" no fim.
   */
  function interpretar(m, cores) {
    m = (m || []).map(function (l) { return (l || []).map(texto); });
    cores = cores || [];
    var h = m.findIndex(function (l) { return l.some(function (c) { return /^cliente$|instalacao|^nome( do)? cliente/.test(nrm(c)); }); });
    if (h < 0) return null;
    var cab = m[h], n = Math.max.apply(null, m.map(function (l) { return l.length; }));
    var ci = colunaDe(cab, 'Instalação');
    var ehDado = function (l) { return ci >= 0 ? /\d{6,}/.test(String(l[ci] || '').replace(/\D/g, '')) : l.filter(Boolean).length >= 3; };
    var dados = [];
    for (var i = h + 1; i < m.length; i++) if (ehDado(m[i])) dados.push(i);

    // legenda: cores com texto fora das linhas de dados, ou textos "Verde = …"
    var legenda = {};
    for (var r = 0; r < m.length; r++) {
      if (r === h || dados.indexOf(r) >= 0) continue;
      for (var c = 0; c < n; c++) {
        var t = (m[r] || [])[c] || '', cor = (cores[r] || [])[c] || '';
        var mt = /^\s*(verde|vermelh\w*|amarel\w*|laranja|azul|roxo|lilas|rosa|cinza)\s*[=:–\-]\s*(.+)$/i.exec(nrm(t) ? t.normalize('NFD').replace(/[̀-ͯ]/g, '') : '');
        if (mt) { var k = Object.keys(PALAVRA_COR).find(function (p) { return nrm(mt[1]).indexOf(p) === 0; }); if (k) legenda[PALAVRA_COR[k]] = t.replace(/^[^=:–\-]+[=:–\-]\s*/, '').trim(); continue; }
        if (cor && cor !== 'cinza') {
          var txt = t || ((m[r] || [])[c + 1] || '');
          if (txt && !/^legenda$/i.test(txt.trim())) legenda[cor] = legenda[cor] || txt;
        }
      }
    }

    // colunas de classificação
    var padrao = ['Cliente', 'Usina', 'Instalação', 'Unid. Consumidora', 'Mês Referencia', 'CPF / CNPJ', 'Data Vencimento Boleto', 'Parceiro', 'Associado',
      'Status Contrato Faturamento', 'Status Contrato Atual', 'Modalidade da Fatura', 'Status da Fatura', 'Geração de Faturas', 'Colaborador'].map(function (nm) { return colunaDe(cab, nm); });
    var classCols = [];
    for (var c2 = 0; c2 < n; c2++) {
      if (padrao.indexOf(c2) >= 0) continue;
      var vals = dados.map(function (r2) { return (m[r2] || [])[c2] || ''; }), cheios = vals.filter(Boolean);
      if (!cheios.length) continue;
      var livre = CAB_LIVRE.test(nrm(cab[c2] || '')), temKw = cheios.some(function (v) { return KW.test(nrm(v)); });
      var poucos = cheios.length <= Math.max(1, 0.8 * dados.length);
      if ((livre && (temKw || poucos)) || (temKw && poucos)) classCols.push(c2);
    }

    // cor de cada linha: a que pinta a maior parte das células preenchidas
    function corDaLinha(r3) {
      var cl = cores[r3] || [], cont = {}, cheias = 0;
      (m[r3] || []).forEach(function (v, c3) { if (v) cheias++; var k = cl[c3]; if (k && k !== 'cinza') cont[k] = (cont[k] || 0) + 1; });
      var melhor = Object.keys(cont).sort(function (a, b) { return cont[b] - cont[a]; })[0];
      return melhor && cont[melhor] >= Math.max(2, 0.4 * cheias) ? melhor : '';
    }

    var usar = [];
    for (var c4 = 0; c4 < n; c4++) {
      if (classCols.indexOf(c4) >= 0) continue;
      if ((cab[c4] || '') !== '' || dados.some(function (r4) { return (m[r4] || [])[c4]; })) usar.push(c4);
    }
    var linhas = [], descartadas = 0, semClassif = 0;
    dados.forEach(function (r5) {
      var partes = classCols.map(function (c5) { return (m[r5] || [])[c5] || ''; }).filter(Boolean);
      var cor = corDaLinha(r5);
      if (cor && (legenda[cor] || !partes.length)) partes.push(legenda[cor] || ('Linha ' + (FEM[cor] || cor)));
      partes = partes.filter(function (p, i) { return partes.indexOf(p) === i; });
      var classif = partes.join(' · ');
      if (!classif) { semClassif++; return; }
      if (/desconect|inativ/.test(nrm(classif)) && !/ouvidoria|saldo|manual|lanc/.test(nrm(classif))) { descartadas++; return; }
      linhas.push(usar.map(function (c6) { return (m[r5] || [])[c6] || ''; }).concat([classif]));
    });
    return { cab: usar.map(function (c7) { return cab[c7] || 'Coluna ' + (c7 + 1); }).concat(['Classificação']), linhas: linhas,
      semClassificacao: semClassif, descartadas: descartadas, legenda: legenda, colunasClassificacao: classCols.map(function (c8) { return cab[c8] || 'Coluna ' + (c8 + 1); }) };
  }

  raiz.Auditada = { interpretar: interpretar, daAbaXlsx: daAbaXlsx, colunaDe: colunaDe, nomeCor: nomeCor };
})(typeof window !== 'undefined' ? window : this);
