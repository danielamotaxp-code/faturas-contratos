/**
 * Plataforma Faturas e Contratos – WattWise
 * Motor que roda no Google Apps Script (servidores do Google), sem depender de ninguém.
 *
 * Este arquivo NÃO contém senhas. Os logins ficam no arquivo Config.gs,
 * que está no repositório privado "faturas-contratos-acessos" e só existe
 * dentro do projeto do Apps Script.
 *
 * Abas da plataforma:
 *   1. Faturas Pendentes   – quem tem fatura nos meses anteriores mas não no mês da planilha
 *   2. Contratos sem Faturas – contratos vinculados que não aparecem no gateway de pagamento
 *   3. Progresso            – evolução semana a semana / mês a mês e clientes recorrentes
 */

var AZUL = '#cfe0f3';
var BRANCO = '#ffffff';
var COR_CABECALHO = '#1f4e79';
var NOME_PASTA = 'WattWise – Faturas e Contratos';
var NOME_BANCO = 'WattWise – Banco da Plataforma (não editar)';
var CAB_REGISTROS = ['id', 'criadoEm', 'usuario', 'aba', 'mes', 'mesRotulo', 'semana', 'semanas',
  'nomeOrigem', 'linkOrigem', 'nomeGerado', 'linkGerado', 'total', 'comFatura', 'semFatura'];
var CAB_PERIODOS = ['aba', 'mes', 'semana', 'total', 'comFatura', 'semFatura', 'atualizadoEm'];
var CAB_HISTORICO = ['aba', 'mes', 'semana', 'instalacao', 'cliente', 'cpfCnpj', 'info'];

var MESES_NOME = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto',
  'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
var MESES_ABREV = { jan: 1, fev: 2, mar: 3, abr: 4, mai: 5, jun: 6, jul: 7, ago: 8, set: 9, out: 10, nov: 11, dez: 12 };
var RE_MES = '(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro|jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)';

/* ------------------------------------------------------------------ */
/* Configuração inicial (rodar UMA vez pelo editor do Apps Script)     */
/* ------------------------------------------------------------------ */

function configurarPlataforma() {
  var pasta = pasta_();
  var banco = banco_();
  Logger.log('Pasta da plataforma: ' + pasta.getUrl());
  Logger.log('Banco da plataforma: ' + banco.getUrl());
  Logger.log('Usuários cadastrados: ' + Object.keys(CONFIG.USUARIOS).join(', '));
}

/* ------------------------------------------------------------------ */
/* API                                                                 */
/* ------------------------------------------------------------------ */

function doGet() {
  return json_({ ok: true, servico: 'WattWise – Faturas e Contratos' });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, erro: 'Requisição inválida.' });
  }
  try {
    if (req.acao === 'login') return json_(login_(req.usuario, req.senha));
    var sessao = validarToken_(req.token);
    if (!sessao) return json_({ ok: false, sessao: false, erro: 'Sua sessão expirou. Entre novamente.' });
    switch (req.acao) {
      case 'listar':
        return json_({ ok: true, registros: listarRegistros_() });
      case 'processar':
        return json_(processar_(req.aba, req.link, sessao));
      case 'progresso':
        return json_({ ok: true, dados: progresso_() });
      case 'sair':
        CacheService.getScriptCache().remove('tok_' + req.token);
        return json_({ ok: true });
    }
    return json_({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return json_({ ok: false, erro: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Login                                                               */
/* ------------------------------------------------------------------ */

function hash_(texto) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function login_(usuario, senha) {
  var chave = normalizar_(usuario);
  var cache = CacheService.getScriptCache();
  var chaveTentativas = 'tent_' + chave;
  var tentativas = Number(cache.get(chaveTentativas) || 0);
  if (tentativas >= 5) return { ok: false, erro: 'Muitas tentativas. Aguarde 5 minutos e tente de novo.' };
  var u = CONFIG.USUARIOS[chave];
  if (!u || hash_(CONFIG.SALT + String(senha || '')) !== u.hash) {
    cache.put(chaveTentativas, String(tentativas + 1), 300);
    return { ok: false, erro: 'Usuário ou senha incorretos.' };
  }
  cache.remove(chaveTentativas);
  var token = Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify({ usuario: chave, nome: u.nome }), 21600);
  return { ok: true, token: token, nome: u.nome };
}

function validarToken_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var v = cache.get('tok_' + token);
  if (!v) return null;
  cache.put('tok_' + token, v, 21600);
  return JSON.parse(v);
}

/* ------------------------------------------------------------------ */
/* Pasta e banco                                                       */
/* ------------------------------------------------------------------ */

function pasta_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PASTA_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recria abaixo */ }
  }
  var pasta = DriveApp.createFolder(NOME_PASTA);
  (CONFIG.EMAILS_COMPARTILHAR || []).forEach(function (email) {
    try { pasta.addEditor(email); } catch (e) { Logger.log('Não compartilhou com ' + email + ': ' + e); }
  });
  props.setProperty('PASTA_ID', pasta.getId());
  return pasta;
}

function subpasta_(nome) {
  var raiz = pasta_();
  var it = raiz.getFoldersByName(nome);
  return it.hasNext() ? it.next() : raiz.createFolder(nome);
}

function banco_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('BANCO_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recria abaixo */ }
  }
  var ss = SpreadsheetApp.create(NOME_BANCO);
  DriveApp.getFileById(ss.getId()).moveTo(pasta_());
  var abas = [['Registros', CAB_REGISTROS], ['Periodos', CAB_PERIODOS], ['Historico', CAB_HISTORICO]];
  abas.forEach(function (a, i) {
    var sh = i === 0 ? ss.getSheets()[0].setName(a[0]) : ss.insertSheet(a[0]);
    sh.getRange('A:Z').setNumberFormat('@');
    sh.getRange(1, 1, 1, a[1].length).setValues([a[1]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  props.setProperty('BANCO_ID', ss.getId());
  return ss;
}

function lerTabela_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];
  var v = sh.getRange(1, 1, lr, lc).getValues();
  var cab = v[0];
  return v.slice(1).map(function (l) {
    var o = {};
    cab.forEach(function (c, i) { o[c] = l[i]; });
    return o;
  });
}

function gravarTabela_(sh, cab, objs) {
  var lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, Math.max(cab.length, sh.getLastColumn())).clearContent();
  if (!objs.length) return;
  var linhas = objs.map(function (o) { return cab.map(function (c) { return o[c] == null ? '' : String(o[c]); }); });
  if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
  var rg = sh.getRange(2, 1, linhas.length, cab.length);
  rg.setNumberFormat('@');
  rg.setValues(linhas);
}

/* ------------------------------------------------------------------ */
/* Utilitários de texto, mês e instalação (lógica pura, testável)      */
/* ------------------------------------------------------------------ */

function normalizar_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

function chaveMes_(ano, mes) {
  return ano + '-' + ('0' + mes).slice(-2);
}

function rotuloMes_(chave) {
  var p = String(chave).split('-');
  return MESES_NOME[Number(p[1]) - 1] + ' ' + p[0];
}

/** Mês da planilha a partir do NOME do arquivo. Ex.: "Contratos e Faturas - OUT 2026" -> 2026-10 */
function mesDoNome_(nome) {
  var n = normalizar_(nome);
  var m = new RegExp('(?:^|[^a-z])' + RE_MES + '[a-z.]*(?:\\s+de)?[\\s/\\-_.]*(\\d{4})').exec(n);
  if (m) {
    var mes = MESES_ABREV[m[1].slice(0, 3)];
    return { chave: chaveMes_(m[2], mes), rotulo: MESES_NOME[mes - 1] + ' ' + m[2] };
  }
  m = /(?:^|[^\d])(\d{1,2})[\/\-_.](\d{4})(?:[^\d]|$)/.exec(n);
  if (m && Number(m[1]) >= 1 && Number(m[1]) <= 12) {
    return { chave: chaveMes_(m[2], Number(m[1])), rotulo: MESES_NOME[Number(m[1]) - 1] + ' ' + m[2] };
  }
  return null;
}

/** Mês de referência de uma célula. Aceita "out./2026", "outubro/2026", "10/2026" ou data. */
function mesDaCelula_(v) {
  if (v instanceof Date && !isNaN(v)) return chaveMes_(v.getFullYear(), v.getMonth() + 1);
  var n = normalizar_(v);
  if (!n) return null;
  var m = new RegExp('^' + RE_MES + '[a-z.]*\\s*[/\\-]?\\s*(\\d{2,4})').exec(n);
  if (m) {
    var ano = m[2].length === 2 ? '20' + m[2] : m[2];
    return chaveMes_(ano, MESES_ABREV[m[1].slice(0, 3)]);
  }
  m = /^(\d{1,2})[\/\-](\d{4})$/.exec(n);
  if (m) return chaveMes_(m[2], Number(m[1]));
  m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(n);
  if (m) return chaveMes_(m[3], Number(m[2]));
  return null;
}

function rotuloCurtoMes_(chave) {
  var p = String(chave).split('-');
  return Object.keys(MESES_ABREV)[Number(p[1]) - 1] + '/' + p[0];
}

/** Instalação só com dígitos e sem zeros à esquerda, para comparar planilhas diferentes. */
function chaveInstalacao_(v) {
  if (typeof v === 'number') v = v.toFixed(0);
  return String(v == null ? '' : v).replace(/\D/g, '').replace(/^0+/, '');
}

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'dd/MM/yyyy');
  if (typeof v === 'number') return Number.isInteger(v) ? v.toFixed(0) : String(v);
  return String(v == null ? '' : v).trim();
}

function coluna_(cab, padroes) {
  var norm = cab.map(normalizar_);
  for (var p = 0; p < padroes.length; p++) {
    for (var i = 0; i < norm.length; i++) if (padroes[p].test(norm[i])) return i;
  }
  return -1;
}

/* ------------------------------------------------------------------ */
/* Regras de negócio                                                   */
/* ------------------------------------------------------------------ */

/**
 * ABA 1 – Faturas Pendentes.
 * Se a instalação tem fatura no mês da planilha (qualquer status, inclusive Cancelado),
 * TODAS as suas linhas ficam azuis. Se não tem, TODAS ficam brancas e ela entra como pendente.
 */
function analisarFaturas_(cab, linhas, mesAtual) {
  var ci = coluna_(cab, [/^instalacao$/, /instala/]);
  var cm = coluna_(cab, [/mes\s*refer/, /referencia/]);
  if (ci < 0 || cm < 0) throw new Error('Não encontrei as colunas "Instalação" e "Mês Referencia" na planilha.');
  var cc = coluna_(cab, [/^cliente/, /titular/]);
  var cd = coluna_(cab, [/cpf|cnpj/]);
  var cu = coluna_(cab, [/^usina/, /usina/]);

  var comAtual = {}, grupos = {}, ordem = [];
  var chaves = linhas.map(function (l) {
    var k = chaveInstalacao_(l[ci]);
    var m = mesDaCelula_(l[cm]);
    if (!k) return '';
    if (!grupos[k]) { grupos[k] = { instalacao: texto_(l[ci]), cliente: '', cpfCnpj: '', usina: '', meses: {} }; ordem.push(k); }
    var g = grupos[k];
    if (!g.cliente && cc >= 0) g.cliente = texto_(l[cc]);
    if (!g.cpfCnpj && cd >= 0) g.cpfCnpj = texto_(l[cd]);
    if (!g.usina && cu >= 0) g.usina = texto_(l[cu]);
    if (m) g.meses[m] = true;
    if (m === mesAtual) comAtual[k] = true;
    return k;
  });

  var cores = linhas.map(function (l, i) {
    var k = chaves[i];
    if (k) return comAtual[k] ? AZUL : BRANCO;
    return mesDaCelula_(l[cm]) === mesAtual ? AZUL : BRANCO;
  });

  var pendentes = ordem.filter(function (k) { return !comAtual[k]; }).map(function (k) {
    var g = grupos[k];
    var meses = Object.keys(g.meses).sort().map(rotuloCurtoMes_).join(', ');
    return {
      instalacao: k, cliente: g.cliente, cpfCnpj: g.cpfCnpj,
      info: 'Faturas em: ' + (meses || '—') + (g.usina ? ' · Usina: ' + g.usina : '')
    };
  });

  return { cores: cores, pendentes: pendentes, total: ordem.length, comFatura: ordem.length - pendentes.length };
}

/**
 * ABA 2 – Contratos sem Faturas.
 * Compara a aba de contratos com a aba de faturas (gateway) pela instalação
 * e devolve só os contratos que não têm fatura.
 */
function analisarContratos_(cabC, linhasC, cabF, linhasF) {
  var iC = coluna_(cabC, [/numero\s*(da\s*)?instala/, /^instalacao$/, /instala/]);
  var iF = coluna_(cabF, [/^instalacao$/, /numero\s*(da\s*)?instala/, /instala/]);
  if (iC < 0) throw new Error('Não encontrei a coluna "Número instalação" na aba de contratos.');
  if (iF < 0) throw new Error('Não encontrei a coluna "Instalação" na aba de faturas.');
  var cT = coluna_(cabC, [/titular/, /^cliente/]);
  var cD = coluna_(cabC, [/cpf|cnpj/]);
  var cS = coluna_(cabC, [/^status\s*contrato/, /status/]);
  var cU = coluna_(cabC, [/nome\s*usina/, /usina/]);

  var faturadas = {};
  linhasF.forEach(function (l) { var k = chaveInstalacao_(l[iF]); if (k) faturadas[k] = true; });

  var indices = [], pendentes = [];
  linhasC.forEach(function (l, i) {
    var k = chaveInstalacao_(l[iC]);
    if (k && faturadas[k]) return;
    indices.push(i);
    var info = [];
    if (cS >= 0 && texto_(l[cS])) info.push('Status: ' + texto_(l[cS]));
    if (cU >= 0 && texto_(l[cU])) info.push('Usina: ' + texto_(l[cU]));
    pendentes.push({
      instalacao: k || texto_(l[iC]) || '(sem instalação)',
      cliente: cT >= 0 ? texto_(l[cT]) : '',
      cpfCnpj: cD >= 0 ? texto_(l[cD]) : '',
      info: info.join(' · ')
    });
  });
  return { indices: indices, pendentes: pendentes, total: linhasC.length, comFatura: linhasC.length - indices.length };
}

/* ------------------------------------------------------------------ */
/* Leitura das planilhas enviadas                                      */
/* ------------------------------------------------------------------ */

function lerAba_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { cab: [], formatoCab: [], valores: [], formatos: [] };
  var rg = sh.getRange(1, 1, lr, lc);
  var v = rg.getValues(), f = rg.getNumberFormats();
  var valores = [], formatos = [];
  for (var i = 1; i < v.length; i++) {
    var vazia = v[i].every(function (c) { return c === '' || c === null; });
    if (!vazia) { valores.push(v[i]); formatos.push(f[i]); }
  }
  return { cab: v[0], formatoCab: f[0], valores: valores, formatos: formatos };
}

function numeroSemana_(nome) {
  var m = /semana\s*(\d+)/.exec(normalizar_(nome));
  return m ? Number(m[1]) : null;
}

function semanasFaturas_(ss) {
  var lista = ss.getSheets().map(function (sh) {
    var n = numeroSemana_(sh.getName());
    return n ? { n: n, sheet: sh } : null;
  }).filter(Boolean);
  if (!lista.length) lista = [{ n: 1, sheet: ss.getSheets()[0] }];
  return lista.sort(function (a, b) { return a.n - b.n; });
}

function semanasContratos_(ss) {
  var por = {};
  ss.getSheets().forEach(function (sh) {
    var nome = normalizar_(sh.getName());
    var n = numeroSemana_(nome) || 1;
    var tipo = /contrat/.test(nome) ? 'c' : (/fatur|gateway|pagamento/.test(nome) ? 'f' : null);
    if (!tipo) return;
    por[n] = por[n] || { n: n };
    por[n][tipo] = sh;
  });
  var lista = Object.keys(por).map(function (k) { return por[k]; }).filter(function (s) { return s.c && s.f; });
  if (!lista.length) throw new Error('Não encontrei as abas "Contratos - Semana N" e "Faturas - Semana N" na planilha.');
  return lista.sort(function (a, b) { return a.n - b.n; });
}

/* ------------------------------------------------------------------ */
/* Geração da planilha nova                                            */
/* ------------------------------------------------------------------ */

function prepararAba_(sh, linhas, colunas) {
  if (sh.getMaxRows() < linhas) sh.insertRowsAfter(sh.getMaxRows(), linhas - sh.getMaxRows());
  if (sh.getMaxColumns() < colunas) sh.insertColumnsAfter(sh.getMaxColumns(), colunas - sh.getMaxColumns());
  if (sh.getMaxRows() > linhas) sh.deleteRows(linhas + 1, sh.getMaxRows() - linhas);
  if (sh.getMaxColumns() > colunas) sh.deleteColumns(colunas + 1, sh.getMaxColumns() - colunas);
}

function estilizarCabecalho_(sh, colunas) {
  sh.getRange(1, 1, 1, colunas).setBackground(COR_CABECALHO).setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
}

function escreverDados_(sh, cab, formatoCab, valores, formatos, cores) {
  var c = cab.length, n = valores.length + 1;
  prepararAba_(sh, n, c);
  var rg = sh.getRange(1, 1, n, c);
  rg.setNumberFormats([formatoCab].concat(formatos));
  rg.setValues([cab].concat(valores));
  estilizarCabecalho_(sh, c);
  if (valores.length && cores) {
    sh.getRange(2, 1, valores.length, c).setBackgrounds(cores.map(function (cor) {
      var linha = []; for (var i = 0; i < c; i++) linha.push(cor); return linha;
    }));
  }
}

function escreverSimples_(sh, cab, linhas) {
  var n = linhas.length + 1, c = cab.length;
  prepararAba_(sh, Math.max(n, 2), c);
  sh.getRange(1, 1, Math.max(n, 2), c).setNumberFormat('@');
  sh.getRange(1, 1, n, c).setValues([cab].concat(linhas));
  estilizarCabecalho_(sh, c);
  sh.autoResizeColumns(1, c);
}

function novaPlanilha_(nome, mes) {
  var ss = SpreadsheetApp.create(nome);
  try { ss.setSpreadsheetLocale('pt_BR'); } catch (e) { /* mantém o padrão */ }
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  DriveApp.getFileById(ss.getId()).moveTo(subpasta_(mes.chave + ' – ' + mes.rotulo));
  return ss;
}

/* ------------------------------------------------------------------ */
/* Processamento                                                       */
/* ------------------------------------------------------------------ */

function abrirPlanilha_(link) {
  var url = String(link || '').trim();
  if (!/docs\.google\.com\/spreadsheets\/d\//.test(url)) {
    throw new Error('Cole o link completo da planilha do Google Sheets (começa com https://docs.google.com/spreadsheets/d/...).');
  }
  try {
    return SpreadsheetApp.openByUrl(url);
  } catch (e) {
    throw new Error('Não consegui abrir a planilha. Verifique se ela está compartilhada com a conta dona da plataforma.');
  }
}

function processar_(aba, link, sessao) {
  if (aba !== 'faturas' && aba !== 'contratos') throw new Error('Aba inválida.');
  var origem = abrirPlanilha_(link);
  var nomeOrigem = origem.getName();
  var mes = mesDoNome_(nomeOrigem);
  if (!mes) {
    throw new Error('Não identifiquei o mês no nome da planilha "' + nomeOrigem +
      '". O nome precisa ter o mês e o ano, por exemplo "OUT 2026".');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Outra planilha está sendo processada agora. Tente de novo em 1 minuto.');
  try {
    var r = aba === 'faturas' ? gerarFaturas_(origem, mes) : gerarContratos_(origem, mes);
    var ultima = r.semanas[r.semanas.length - 1];
    var agora = new Date().toISOString();

    var banco = banco_();
    // períodos e histórico de cada semana (substitui a versão anterior da mesma semana)
    var shP = banco.getSheetByName('Periodos'), shH = banco.getSheetByName('Historico');
    var periodos = lerTabela_(shP), hist = lerTabela_(shH);
    var processadas = {};
    r.semanas.forEach(function (s) { processadas[s.n] = s; });
    var mesmo = function (o) { return o.aba === aba && String(o.mes) === mes.chave && processadas[Number(o.semana)]; };
    periodos = periodos.filter(function (o) { return !mesmo(o); });
    hist = hist.filter(function (o) { return !mesmo(o); });
    r.semanas.forEach(function (s) {
      periodos.push({ aba: aba, mes: mes.chave, semana: s.n, total: s.total, comFatura: s.comFatura,
        semFatura: s.pendentes.length, atualizadoEm: agora });
      s.pendentes.forEach(function (p) {
        hist.push({ aba: aba, mes: mes.chave, semana: s.n, instalacao: p.instalacao, cliente: p.cliente,
          cpfCnpj: p.cpfCnpj, info: p.info });
      });
    });
    gravarTabela_(shP, CAB_PERIODOS, periodos);
    gravarTabela_(shH, CAB_HISTORICO, hist);

    var registro = {
      id: Utilities.getUuid(), criadoEm: agora, usuario: sessao.nome, aba: aba, mes: mes.chave,
      mesRotulo: mes.rotulo, semana: ultima.n, semanas: r.semanas.map(function (s) { return s.n; }).join(', '),
      nomeOrigem: nomeOrigem, linkOrigem: origem.getUrl(), nomeGerado: r.planilha.getName(),
      linkGerado: r.planilha.getUrl(), total: ultima.total, comFatura: ultima.comFatura,
      semFatura: ultima.pendentes.length
    };
    banco.getSheetByName('Registros').appendRow(CAB_REGISTROS.map(function (c) { return String(registro[c]); }));
    return { ok: true, registro: registro };
  } finally {
    lock.releaseLock();
  }
}

function gerarFaturas_(origem, mes) {
  var semanas = semanasFaturas_(origem);
  var ultima = semanas[semanas.length - 1].n;
  var ss = novaPlanilha_('Faturas Pendentes – ' + mes.rotulo + ' – Semana ' + ultima, mes);
  var padrao = ss.getSheets()[0];
  var resultado = [];

  semanas.forEach(function (s) {
    var d = lerAba_(s.sheet);
    var a = analisarFaturas_(d.cab, d.valores, mes.chave);
    escreverDados_(ss.insertSheet('Semana ' + s.n), d.cab, d.formatoCab, d.valores, d.formatos, a.cores);
    resultado.push({ n: s.n, total: a.total, comFatura: a.comFatura, pendentes: a.pendentes });
  });

  var ult = resultado[resultado.length - 1];
  escreverSimples_(ss.insertSheet('Pendentes – Semana ' + ult.n),
    ['Instalação', 'Cliente', 'CPF/CNPJ', 'Detalhes'],
    ult.pendentes.map(function (p) { return [p.instalacao, p.cliente, p.cpfCnpj, p.info]; }));
  escreverSimples_(ss.insertSheet('Resumo', 0),
    ['Semana', 'Instalações', 'Com fatura de ' + mes.rotulo, 'Sem fatura (gerar)'],
    resultado.map(function (s) { return ['Semana ' + s.n, String(s.total), String(s.comFatura), String(s.pendentes.length)]; }));
  ss.deleteSheet(padrao);
  ss.setActiveSheet(ss.getSheetByName('Semana ' + ult.n));
  return { planilha: ss, semanas: resultado };
}

function gerarContratos_(origem, mes) {
  var semanas = semanasContratos_(origem);
  var ultima = semanas[semanas.length - 1].n;
  var ss = novaPlanilha_('Contratos sem Faturas – ' + mes.rotulo + ' – Semana ' + ultima, mes);
  var padrao = ss.getSheets()[0];
  var resultado = [];

  semanas.forEach(function (s) {
    var c = lerAba_(s.c), f = lerAba_(s.f);
    var a = analisarContratos_(c.cab, c.valores, f.cab, f.valores);
    var valores = a.indices.map(function (i) { return c.valores[i]; });
    var formatos = a.indices.map(function (i) { return c.formatos[i]; });
    escreverDados_(ss.insertSheet('Semana ' + s.n), c.cab, c.formatoCab, valores, formatos, null);
    resultado.push({ n: s.n, total: a.total, comFatura: a.comFatura, pendentes: a.pendentes });
  });

  escreverSimples_(ss.insertSheet('Resumo', 0),
    ['Semana', 'Contratos vinculados', 'Com fatura gerada', 'Sem fatura'],
    resultado.map(function (s) { return ['Semana ' + s.n, String(s.total), String(s.comFatura), String(s.pendentes.length)]; }));
  ss.deleteSheet(padrao);
  ss.setActiveSheet(ss.getSheetByName('Semana ' + ultima));
  return { planilha: ss, semanas: resultado };
}

/* ------------------------------------------------------------------ */
/* Lista de planilhas e progresso                                      */
/* ------------------------------------------------------------------ */

function listarRegistros_() {
  var regs = lerTabela_(banco_().getSheetByName('Registros'));
  regs.sort(function (a, b) { return String(b.criadoEm).localeCompare(String(a.criadoEm)); });
  return regs;
}

function progresso_() {
  var banco = banco_();
  return calcularProgresso_(lerTabela_(banco.getSheetByName('Periodos')), lerTabela_(banco.getSheetByName('Historico')));
}

/** Lógica pura do painel de progresso. */
function calcularProgresso_(periodos, hist) {
  var saida = {};
  ['faturas', 'contratos'].forEach(function (aba) {
    var ps = periodos.filter(function (p) { return p.aba === aba; }).map(function (p) {
      return { chave: p.mes + '-S' + ('0' + p.semana).slice(-2), mes: String(p.mes), semana: Number(p.semana),
        total: Number(p.total), comFatura: Number(p.comFatura), semFatura: Number(p.semFatura) };
    }).sort(function (a, b) { return a.chave.localeCompare(b.chave); });

    var conjuntos = {};
    ps.forEach(function (p) { conjuntos[p.chave] = {}; });
    var clientes = {};
    hist.forEach(function (h) {
      if (h.aba !== aba) return;
      var ch = h.mes + '-S' + ('0' + h.semana).slice(-2);
      if (!conjuntos[ch]) return;
      var k = String(h.instalacao);
      conjuntos[ch][k] = true;
      var c = clientes[k] || (clientes[k] = { instalacao: k, periodos: [], meses: {} });
      c.periodos.push(ch);
      c.meses[h.mes] = true;
      if (!c.ultimo || ch >= c.ultimo) { c.ultimo = ch; c.cliente = h.cliente || c.cliente; c.cpfCnpj = h.cpfCnpj || c.cpfCnpj; c.info = h.info; }
    });

    var comparar = function (lista) {
      return lista.map(function (p, i) {
        var atual = conjuntos[p.chave] || {};
        var ant = i > 0 ? (conjuntos[lista[i - 1].chave] || {}) : null;
        var novos = 0, persist = 0, resolvidos = 0;
        Object.keys(atual).forEach(function (k) { if (ant && ant[k]) persist++; else novos++; });
        if (ant) Object.keys(ant).forEach(function (k) { if (!atual[k]) resolvidos++; });
        return {
          chave: p.chave, rotulo: rotuloMes_(p.mes) + (p.rotuloSemana === false ? '' : ' · Semana ' + p.semana),
          mes: p.mes, semana: p.semana, total: p.total, comFatura: p.comFatura, semFatura: p.semFatura,
          novos: ant ? novos : null, persistentes: ant ? persist : null, resolvidos: ant ? resolvidos : null
        };
      });
    };

    var semanal = comparar(ps);
    var ultimaDoMes = {};
    ps.forEach(function (p) { ultimaDoMes[p.mes] = p; });
    var mensal = comparar(Object.keys(ultimaDoMes).sort().map(function (m) {
      var p = ultimaDoMes[m];
      return { chave: p.chave, mes: p.mes, semana: p.semana, total: p.total, comFatura: p.comFatura, semFatura: p.semFatura, rotuloSemana: false };
    }));

    var ordem = ps.map(function (p) { return p.chave; });
    var ultimoPeriodo = ordem[ordem.length - 1];
    var lista = Object.keys(clientes).map(function (k) {
      var c = clientes[k];
      var seq = 0;
      for (var i = ordem.length - 1; i >= 0 && conjuntos[ordem[i]][k]; i--) seq++;
      c.periodos.sort();
      return {
        instalacao: k, cliente: c.cliente || '', cpfCnpj: c.cpfCnpj || '', info: c.info || '',
        vezes: c.periodos.length, meses: Object.keys(c.meses).length,
        primeira: c.periodos[0], ultima: c.ultimo, seguidas: seq,
        pendenteAgora: c.ultimo === ultimoPeriodo
      };
    }).sort(function (a, b) {
      return (b.pendenteAgora - a.pendenteAgora) || (b.seguidas - a.seguidas) || (b.vezes - a.vezes) ||
        String(a.cliente).localeCompare(String(b.cliente));
    });

    saida[aba] = { semanal: semanal, mensal: mensal, clientes: lista, ultimoPeriodo: ultimoPeriodo || null };
  });
  return saida;
}
