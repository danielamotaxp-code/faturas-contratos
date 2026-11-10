/**
 * Regras de negócio da plataforma Faturas e Contratos – WattWise.
 * Lógica pura (sem tela e sem rede), usada pelo site e pelos testes.
 */
(function (raiz) {
  'use strict';

  var AZUL = '9BC2E6';
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

  /* ---------- CPF / CNPJ ---------- */
  function cpfValido(c) {
    if (/^(\d)\1{10}$/.test(c)) return false;
    var d = c.split('').map(Number);
    for (var n = 9; n <= 10; n++) {
      var s = 0;
      for (var i = 0; i < n; i++) s += d[i] * (n + 1 - i);
      if ((s * 10) % 11 % 10 !== d[n]) return false;
    }
    return true;
  }
  function cnpjValido(c) {
    if (/^(\d)\1{13}$/.test(c)) return false;
    var d = c.split('').map(Number);
    var pesos = [[5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2], [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]];
    for (var k = 0; k < 2; k++) {
      var s = 0;
      for (var i = 0; i < pesos[k].length; i++) s += d[i] * pesos[k][i];
      var r = s % 11;
      if ((r < 2 ? 0 : 11 - r) !== d[12 + k]) return false;
    }
    return true;
  }
  function preencher(d, n) { while (d.length < n) d = '0' + d; return d; }

  /**
   * Devolve o CPF (11 dígitos) ou CNPJ (14 dígitos) completo, recolocando os zeros
   * à esquerda que a planilha perde quando guarda o documento como número.
   * Usa os dígitos verificadores para decidir entre CPF e CNPJ.
   */
  function corrigirDocumento(v) {
    if (typeof v === 'number') v = v.toFixed(0);
    var d = String(v == null ? '' : v).replace(/\D/g, '');
    if (!d) return String(v == null ? '' : v).trim();
    var cpf = preencher(d, 11), cnpj = preencher(d, 14);
    if (d.length <= 11 && cpfValido(cpf)) return cpf;
    if (d.length <= 14 && cnpjValido(cnpj)) return cnpj;
    return d.length <= 11 ? cpf : cnpj;
  }

  /** Número longo (instalação, unidade consumidora) como texto, sem notação científica nem casas decimais. */
  function codigoTexto(v) {
    if (typeof v === 'number') return v.toFixed(0);
    return String(v == null ? '' : v).trim();
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

  /** Os últimos N meses até o mês da planilha. Ex.: 2026-10 -> [2026-08, 2026-09, 2026-10] */
  function mesesJanela(mesAtual, n) {
    var p = String(mesAtual).split('-'), ano = Number(p[0]), mes = Number(p[1]), out = [];
    for (var i = (n || 3) - 1; i >= 0; i--) {
      var m = mes - i, a = ano;
      while (m < 1) { m += 12; a--; }
      out.push(chaveMes(a, m));
    }
    return out;
  }

  /** Mantém só as linhas dos últimos 3 meses (pela coluna "Mês Referencia"). */
  function filtrarUltimosMeses(cab, linhas, indices, mesAtual) {
    var cm = coluna(cab, [/mes\s*refer/, /referencia/]);
    if (cm < 0) throw new Error('Não encontrei a coluna "Mês Referencia" na planilha.');
    var janela = mesesJanela(mesAtual, 3), ok = {};
    janela.forEach(function (m) { ok[m] = true; });
    var L = [], I = [];
    linhas.forEach(function (l, i) {
      if (ok[mesDaCelula(l[cm])]) { L.push(l); I.push(indices ? indices[i] : i); }
    });
    return { linhas: L, indices: I, janela: janela, removidas: linhas.length - L.length };
  }

  function numeroSemana(nome) {
    var m = /semana\s*(\d+)/.exec(normalizar(nome));
    return m ? Number(m[1]) : null;
  }

  /** Semana do mês pela data: dias 1–7 = 1, 8–14 = 2, 15–21 = 3, 22 em diante = 4. */
  function semanaDaData(d) {
    d = d || new Date();
    return Math.min(4, Math.ceil(d.getDate() / 7));
  }

  /**
   * Aba de faturas a conferir: a aba "Semana <preferida>" se existir; senão a aba "Semana N" mais recente;
   * senão a primeira aba da planilha.
   */
  function abaFaturas(nomesAbas, preferida) {
    var semanas = nomesAbas.map(function (nome) { return { n: numeroSemana(nome), aba: nome }; })
      .filter(function (x) { return x.n; }).sort(function (a, b) { return a.n - b.n; });
    var igual = semanas.filter(function (x) { return x.n === Number(preferida); })[0];
    if (igual) return igual.aba;
    if (semanas.length) return semanas[semanas.length - 1].aba;
    return nomesAbas[0];
  }

  /** Par de abas Contratos + Faturas a conferir (mesma lógica de escolha da semana). */
  function abaContratos(nomesAbas, preferida) {
    var por = {};
    nomesAbas.forEach(function (nome) {
      var norm = normalizar(nome);
      if (/sem\s*fatura/.test(norm)) return; // abas geradas pela própria plataforma
      var n = numeroSemana(norm) || 0;
      var tipo = /contrat/.test(norm) ? 'c' : (/fatur|gateway|pagamento/.test(norm) ? 'f' : null);
      if (!tipo) return;
      por[n] = por[n] || { n: n };
      if (!por[n][tipo]) por[n][tipo] = nome;
    });
    var pares = Object.keys(por).map(function (k) { return por[k]; }).filter(function (s) { return s.c && s.f; })
      .sort(function (a, b) { return a.n - b.n; });
    if (!pares.length) throw new Error('Não encontrei uma aba de contratos e uma aba de faturas (ex.: "Contratos - Semana 1" e "Faturas - Semana 1").');
    var igual = pares.filter(function (x) { return x.n === Number(preferida); })[0];
    return igual || pares[pares.length - 1];
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
      if (!g.cpfCnpj && cd >= 0) g.cpfCnpj = corrigirDocumento(l[cd]);
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
  function analisarContratos(cabC, linhasC, cabF, linhasF, mesAtual, nomeAbaC) {
    var iC = coluna(cabC, [/numero\s*(da\s*)?instala/, /^instalacao$/, /instala/]);
    var iF = coluna(cabF, [/^instalacao$/, /numero\s*(da\s*)?instala/, /instala/]);
    var abaC = nomeAbaC ? '"' + nomeAbaC + '"' : 'de contratos';
    if (iC < 0) throw new Error('Não encontrei a coluna "Número instalação" na aba ' + abaC + '.');
    if (iF < 0) throw new Error('Não encontrei a coluna "Instalação" na aba de faturas.');

    // Confere se a aba de contratos tem mesmo contratos (e não cabeçalhos repetidos ou faturas)
    var comNumero = linhasC.filter(function (l) { return chaveInstalacao(l[iC]).length >= 5; }).length;
    if (!linhasC.length || comNumero < linhasC.length * 0.5) {
      var repetido = linhasC.length && normalizar(linhasC[0][iC]) === normalizar(cabC[iC]);
      throw new Error('A aba ' + abaC + ' não parece ter contratos: só ' + comNumero + ' de ' + linhasC.length +
        ' linhas têm número de instalação' + (repetido ? ' (as linhas repetem o cabeçalho "' + cabC[iC] + '")' : '') +
        '. Confira se os contratos foram colados nessa aba e envie de novo.');
    }
    if (coluna(cabC, [/mes\s*refer/]) >= 0 && coluna(cabC, [/numero\s*(da\s*)?instala/, /titular/]) < 0) {
      throw new Error('A aba ' + abaC + ' tem colunas de faturas (ex.: "Mês Referencia"), e não de contratos. Confira se as abas não foram trocadas.');
    }

    // Só valem as faturas do mês da planilha (o gateway pode trazer vários meses)
    var cmF = coluna(cabF, [/mes\s*refer/, /referencia/]);
    var faturasDoMes = linhasF;
    if (mesAtual && cmF >= 0) {
      faturasDoMes = linhasF.filter(function (l) { return mesDaCelula(l[cmF]) === mesAtual; });
      if (!faturasDoMes.length) {
        throw new Error('A aba de faturas não tem nenhuma fatura de ' + rotuloMes(mesAtual) +
          ' na coluna "' + cabF[cmF] + '". Confira o mês no nome da planilha.');
      }
    }
    linhasF = faturasDoMes;
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
        cpfCnpj: cD >= 0 ? corrigirDocumento(l[cD]) : '', info: info.join(' · ') });
    });
    return { indices: indices, pendentes: pendentes, total: linhasC.length, comFatura: linhasC.length - indices.length,
      faturasDoMes: linhasF.length };
  }

  /**
   * O histórico guarda, para cada semana, uma versão por envio:
   *   historico[aba][periodo].versoes[registroId] = { mes, semana, total, comFatura, semFatura, pendentes, criadoEm }
   * Vale a versão mais recente. Ao excluir um envio, a semana volta para a versão anterior (se houver).
   */
  function historicoVigente(historico) {
    var out = {};
    ['faturas', 'contratos'].forEach(function (aba) {
      out[aba] = {};
      var h = (historico && historico[aba]) || {};
      Object.keys(h).forEach(function (ch) {
        var p = h[ch];
        if (!p.versoes) { out[aba][ch] = p; return; }
        var ids = Object.keys(p.versoes);
        if (!ids.length) return;
        ids.sort(function (a, b) { return String(p.versoes[a].criadoEm).localeCompare(String(p.versoes[b].criadoEm)); });
        out[aba][ch] = p.versoes[ids[ids.length - 1]];
      });
    });
    return out;
  }

  /**
   * Painel de progresso.
   * historico: { faturas: { "2026-10-S01": { mes, semana, total, comFatura, semFatura, pendentes: [instalações] } }, contratos: {...} }
   * clientes:  { faturas: { instalacao: { cliente, cpfCnpj, info } }, contratos: {...} }
   */
  function calcularProgresso(historico, clientes) {
    historico = historicoVigente(historico);
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
    corrigirDocumento: corrigirDocumento, codigoTexto: codigoTexto, historicoVigente: historicoVigente,
    semanaDaData: semanaDaData, abaFaturas: abaFaturas, abaContratos: abaContratos,
    mesesJanela: mesesJanela, filtrarUltimosMeses: filtrarUltimosMeses,
    analisarFaturas: analisarFaturas, analisarContratos: analisarContratos, calcularProgresso: calcularProgresso
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Regras = api;
})(this);
