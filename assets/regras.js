/**
 * Regras de negócio da plataforma Faturas e Contratos – WattWise.
 * Lógica pura (sem tela e sem rede), usada pelo site e pelos testes.
 */
(function (raiz) {
  'use strict';

  var AZUL = 'CFE0F3';
  var BRANCO = 'FFFFFF';
  var MESES_NOME = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto',
    'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  var ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  var RE_MES = '(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro|jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)';

  function normalizar(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  }
  function chaveMes(ano, mes) { return ano + '-' + ('0' + mes).slice(-2); }
  function rotuloMes(chave) { var p = String(chave).split('-'); return MESES_NOME[Number(p[1]) - 1] + ' ' + p[0]; }
  function rotuloCurtoMes(chave) { var p = String(chave).split('-'); return ABREV[Number(p[1]) - 1] + '/' + p[0]; }
  function chavePeriodo(mes, semana) { return mes + '-S' + ('0' + semana).slice(-2); }

  /** Mês da planilha a partir do NOME do arquivo. Ex.: "Contratos e Faturas - OUT 2026.xlsx" -> 2026-10 */
  function mesDoNome(nome) {
    var n = normalizar(nome).replace(/\.(xlsx|xlsm|xls|csv|ods)$/, '');
    var m = new RegExp('(?:^|[^a-z])' + RE_MES + '[a-z.]*(?:\\s+de)?[\\s/\\-_.]*(\\d{4})').exec(n);
    if (m) {
      var mes = ABREV.indexOf(m[1].slice(0, 3)) + 1;
      return { chave: chaveMes(m[2], mes), rotulo: MESES_NOME[mes - 1] + ' ' + m[2] };
    }
    m = /(?:^|[^\d])(\d{1,2})[\/\-_.](\d{4})(?:[^\d]|$)/.exec(n);
    if (m && Number(m[1]) >= 1 && Number(m[1]) <= 12) {
      return { chave: chaveMes(m[2], Number(m[1])), rotulo: MESES_NOME[Number(m[1]) - 1] + ' ' + m[2] };
    }
    return null;
  }

  /** Mês de referência de uma célula: "out./2026", "outubro/2026", "10/2026" ou data. */
  function mesDaCelula(v) {
    if (v instanceof Date && !isNaN(v)) return chaveMes(v.getFullYear(), v.getMonth() + 1);
    var n = normalizar(v);
    if (!n) return null;
    var m = new RegExp('^' + RE_MES + '[a-z.]*\\s*(?:de\\s*)?[/\\-]?\\s*(\\d{2,4})').exec(n);
    if (m) return chaveMes(m[2].length === 2 ? '20' + m[2] : m[2], ABREV.indexOf(m[1].slice(0, 3)) + 1);
    m = /^(\d{1,2})[\/\-](\d{4})$/.exec(n);
    if (m) return chaveMes(m[2], Number(m[1]));
    m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(n);
    if (m) return chaveMes(m[3], Number(m[2]));
    return null;
  }

  /** Instalação só com dígitos e sem zeros à esquerda, para comparar planilhas diferentes. */
  function chaveInstalacao(v) {
    if (typeof v === 'number') v = v.toFixed(0);
    return String(v == null ? '' : v).replace(/\D/g, '').replace(/^0+/, '');
  }

  function texto(v) {
    if (v instanceof Date && !isNaN(v)) return ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + v.getFullYear();
    if (typeof v === 'number') return Number.isInteger(v) ? v.toFixed(0) : String(v);
    return String(v == null ? '' : v).trim();
  }

  function coluna(cab, padroes) {
    var norm = cab.map(normalizar);
    for (var p = 0; p < padroes.length; p++) {
      for (var i = 0; i < norm.length; i++) if (padroes[p].test(norm[i])) return i;
    }
    return -1;
  }

  function numeroSemana(nome) {
    var m = /semana\s*(\d+)/.exec(normalizar(nome));
    return m ? Number(m[1]) : null;
  }

  /** Abas "Semana N" de uma planilha de faturas. Sem abas assim, usa a primeira com a semana informada. */
  function semanasFaturas(nomesAbas, semanaManual) {
    var lista = nomesAbas.map(function (nome) {
      var n = numeroSemana(nome);
      return n ? { n: n, aba: nome } : null;
    }).filter(Boolean);
    if (!lista.length) {
      if (!semanaManual) throw new Error('A planilha não tem abas "Semana N". Escolha a semana no campo "Semana" e envie de novo.');
      lista = [{ n: Number(semanaManual), aba: nomesAbas[0] }];
    }
    return lista.sort(function (a, b) { return a.n - b.n; });
  }

  /** Pares "Contratos - Semana N" + "Faturas - Semana N". */
  function semanasContratos(nomesAbas, semanaManual) {
    var por = {};
    nomesAbas.forEach(function (nome) {
      var norm = normalizar(nome);
      var n = numeroSemana(norm) || Number(semanaManual) || 0;
      var tipo = /contrat/.test(norm) ? 'c' : (/fatur|gateway|pagamento/.test(norm) ? 'f' : null);
      if (!tipo) return;
      por[n] = por[n] || { n: n };
      por[n][tipo] = nome;
    });
    var lista = Object.keys(por).map(function (k) { return por[k]; }).filter(function (s) { return s.c && s.f; });
    if (!lista.length) throw new Error('Não encontrei as abas "Contratos - Semana N" e "Faturas - Semana N" no arquivo.');
    if (lista.some(function (s) { return !s.n; })) throw new Error('As abas não têm o número da semana. Escolha a semana no campo "Semana" e envie de novo.');
    return lista.sort(function (a, b) { return a.n - b.n; });
  }

  /**
   * ABA 1 – Faturas Pendentes.
   * Instalação com fatura no mês da planilha (qualquer status, inclusive Cancelado): TODAS as linhas em azul.
   * Instalação sem fatura no mês: TODAS as linhas em branco, e ela entra como pendente.
   */
  function analisarFaturas(cab, linhas, mesAtual) {
    var ci = coluna(cab, [/^instalacao$/, /instala/]);
    var cm = coluna(cab, [/mes\s*refer/, /referencia/]);
    if (ci < 0 || cm < 0) throw new Error('Não encontrei as colunas "Instalação" e "Mês Referencia" na planilha.');
    var cc = coluna(cab, [/^cliente/, /titular/]);
    var cd = coluna(cab, [/cpf|cnpj/]);
    var cu = coluna(cab, [/^usina/, /usina/]);

    var comAtual = {}, grupos = {}, ordem = [];
    var chaves = linhas.map(function (l) {
      var k = chaveInstalacao(l[ci]);
      var m = mesDaCelula(l[cm]);
      if (!k) return '';
      if (!grupos[k]) { grupos[k] = { cliente: '', cpfCnpj: '', usina: '', meses: {} }; ordem.push(k); }
      var g = grupos[k];
      if (!g.cliente && cc >= 0) g.cliente = texto(l[cc]);
      if (!g.cpfCnpj && cd >= 0) g.cpfCnpj = texto(l[cd]);
      if (!g.usina && cu >= 0) g.usina = texto(l[cu]);
      if (m) g.meses[m] = true;
      if (m === mesAtual) comAtual[k] = true;
      return k;
    });

    var cores = linhas.map(function (l, i) {
      var k = chaves[i];
      if (k) return comAtual[k] ? AZUL : BRANCO;
      return mesDaCelula(l[cm]) === mesAtual ? AZUL : BRANCO;
    });

    var pendentes = ordem.filter(function (k) { return !comAtual[k]; }).map(function (k) {
      var g = grupos[k];
      var meses = Object.keys(g.meses).sort().map(rotuloCurtoMes).join(', ');
      return { instalacao: k, cliente: g.cliente, cpfCnpj: g.cpfCnpj,
        info: 'Faturas em: ' + (meses || '—') + (g.usina ? ' · Usina: ' + g.usina : '') };
    });
    return { cores: cores, pendentes: pendentes, total: ordem.length, comFatura: ordem.length - pendentes.length };
  }

  /**
   * ABA 2 – Contratos sem Faturas.
   * Compara contratos com faturas (gateway) pela instalação e devolve só os contratos sem fatura.
   */
  function analisarContratos(cabC, linhasC, cabF, linhasF) {
    var iC = coluna(cabC, [/numero\s*(da\s*)?instala/, /^instalacao$/, /instala/]);
    var iF = coluna(cabF, [/^instalacao$/, /numero\s*(da\s*)?instala/, /instala/]);
    if (iC < 0) throw new Error('Não encontrei a coluna "Número instalação" na aba de contratos.');
    if (iF < 0) throw new Error('Não encontrei a coluna "Instalação" na aba de faturas.');
    var cT = coluna(cabC, [/titular/, /^cliente/]);
    var cD = coluna(cabC, [/cpf|cnpj/]);
    var cS = coluna(cabC, [/^status\s*contrato/, /status/]);
    var cU = coluna(cabC, [/nome\s*usina/, /usina/]);

    var faturadas = {};
    linhasF.forEach(function (l) { var k = chaveInstalacao(l[iF]); if (k) faturadas[k] = true; });

    var indices = [], pendentes = [];
    linhasC.forEach(function (l, i) {
      var k = chaveInstalacao(l[iC]);
      if (k && faturadas[k]) return;
      indices.push(i);
      var info = [];
      if (cS >= 0 && texto(l[cS])) info.push('Status: ' + texto(l[cS]));
      if (cU >= 0 && texto(l[cU])) info.push('Usina: ' + texto(l[cU]));
      pendentes.push({ instalacao: k || '(sem instalação)', cliente: cT >= 0 ? texto(l[cT]) : '',
        cpfCnpj: cD >= 0 ? texto(l[cD]) : '', info: info.join(' · ') });
    });
    return { indices: indices, pendentes: pendentes, total: linhasC.length, comFatura: linhasC.length - indices.length };
  }

  /**
   * Painel de progresso.
   * historico: { faturas: { "2026-10-S01": { mes, semana, total, comFatura, semFatura, pendentes: [instalações] } }, contratos: {...} }
   * clientes:  { faturas: { instalacao: { cliente, cpfCnpj, info } }, contratos: {...} }
   */
  function calcularProgresso(historico, clientes) {
    var saida = {};
    ['faturas', 'contratos'].forEach(function (aba) {
      var h = (historico && historico[aba]) || {};
      var info = (clientes && clientes[aba]) || {};
      var ordem = Object.keys(h).sort();
      var conj = {};
      ordem.forEach(function (ch) { conj[ch] = {}; (h[ch].pendentes || []).forEach(function (k) { conj[ch][k] = true; }); });

      var comparar = function (chaves, mensal) {
        return chaves.map(function (ch, i) {
          var p = h[ch], atual = conj[ch], ant = i > 0 ? conj[chaves[i - 1]] : null;
          var novos = 0, persist = 0, resolvidos = 0;
          Object.keys(atual).forEach(function (k) { if (ant && ant[k]) persist++; else novos++; });
          if (ant) Object.keys(ant).forEach(function (k) { if (!atual[k]) resolvidos++; });
          return { chave: ch, mes: p.mes, semana: p.semana,
            rotulo: rotuloMes(p.mes) + (mensal ? '' : ' · Semana ' + p.semana),
            total: p.total, comFatura: p.comFatura, semFatura: p.semFatura,
            novos: ant ? novos : null, persistentes: ant ? persist : null, resolvidos: ant ? resolvidos : null };
        });
      };

      var ultimaDoMes = {};
      ordem.forEach(function (ch) { ultimaDoMes[h[ch].mes] = ch; });
      var ultimo = ordem[ordem.length - 1] || null;

      var porCliente = {};
      ordem.forEach(function (ch) {
        Object.keys(conj[ch]).forEach(function (k) {
          var c = porCliente[k] || (porCliente[k] = { periodos: [], meses: {} });
          c.periodos.push(ch);
          c.meses[h[ch].mes] = true;
        });
      });
      var lista = Object.keys(porCliente).map(function (k) {
        var c = porCliente[k], d = info[k] || {};
        var seq = 0;
        for (var i = ordem.length - 1; i >= 0 && conj[ordem[i]][k]; i--) seq++;
        return { instalacao: k, cliente: d.cliente || '', cpfCnpj: d.cpfCnpj || '', info: d.info || '',
          vezes: c.periodos.length, meses: Object.keys(c.meses).length, primeira: c.periodos[0],
          ultima: c.periodos[c.periodos.length - 1], seguidas: seq, pendenteAgora: !!(ultimo && conj[ultimo][k]) };
      }).sort(function (a, b) {
        return (b.pendenteAgora - a.pendenteAgora) || (b.seguidas - a.seguidas) || (b.vezes - a.vezes) ||
          String(a.cliente).localeCompare(String(b.cliente));
      });

      saida[aba] = {
        semanal: comparar(ordem, false),
        mensal: comparar(Object.keys(ultimaDoMes).sort().map(function (m) { return ultimaDoMes[m]; }), true),
        clientes: lista, ultimoPeriodo: ultimo
      };
    });
    return saida;
  }

  var api = {
    AZUL: AZUL, BRANCO: BRANCO, MESES_NOME: MESES_NOME,
    normalizar: normalizar, rotuloMes: rotuloMes, chavePeriodo: chavePeriodo, mesDoNome: mesDoNome,
    mesDaCelula: mesDaCelula, chaveInstalacao: chaveInstalacao, texto: texto, coluna: coluna,
    semanasFaturas: semanasFaturas, semanasContratos: semanasContratos,
    analisarFaturas: analisarFaturas, analisarContratos: analisarContratos, calcularProgresso: calcularProgresso
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Regras = api;
})(this);
