(function () {
  'use strict';

  var API = (window.WATTWISE_CONFIG || {}).API_URL || '';
  var estado = { token: null, nome: '', registros: [], progresso: null, progAba: 'faturas', progPeriodo: 'semanal' };

  var TEXTOS = {
    faturas: {
      selo: '1',
      titulo: 'Faturas Pendentes',
      ajuda: 'Cole o link da planilha de faturas do mês (ex.: <b>Clientes Não Faturados – OUT 2026</b>). O mês é lido do <b>nome da planilha</b> e cada aba <b>Semana N</b> é conferida. A plataforma gera uma planilha nova com todas as semanas: clientes com fatura do mês ficam em <b>azul</b> (todas as linhas) e os que ainda não têm ficam em <b>branco</b>.',
      gerado: 'Faturas Pendentes',
      rotTotal: 'instalações', rotCom: 'com fatura do mês', rotSem: 'sem fatura (gerar)'
    },
    contratos: {
      selo: '2',
      titulo: 'Contratos sem Faturas',
      ajuda: 'Cole o link da planilha com as abas <b>Contratos – Semana N</b> e <b>Faturas – Semana N</b> (gateway de pagamento). A plataforma compara as duas pela instalação e gera uma planilha nova só com os clientes que têm contrato com a WattWise e não têm fatura gerada.',
      gerado: 'Contratos sem Faturas',
      rotTotal: 'contratos', rotCom: 'com fatura', rotSem: 'sem fatura'
    }
  };

  /* ---------- utilitários ---------- */
  var $ = function (s, el) { return (el || document).querySelector(s); };
  var $$ = function (s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); };
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function num(n) { return n == null || n === '' || isNaN(n) ? '—' : Number(n).toLocaleString('pt-BR'); }
  function dataHora(iso) {
    var d = new Date(iso);
    return isNaN(d) ? '' : d.toLocaleDateString('pt-BR') + ' às ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }
  var MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  function rotuloMes(chave) { var p = String(chave).split('-'); return MESES[Number(p[1]) - 1] + ' de ' + p[0]; }
  function rotuloPeriodo(ch) {
    var m = /^(\d{4})-(\d{2})-S(\d+)$/.exec(ch || '');
    return m ? MESES[Number(m[2]) - 1].slice(0, 3) + '/' + m[1] + ' · S' + Number(m[3]) : '—';
  }
  function guardar(k, v) { try { v == null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch (e) { /* sem armazenamento */ } }
  function ler(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } }

  function api(acao, dados) {
    if (!API) return Promise.reject(new Error('A plataforma ainda não foi conectada ao motor (falta o endereço em assets/config.js).'));
    var corpo = Object.assign({ acao: acao, token: estado.token }, dados || {});
    return fetch(API, { method: 'POST', body: JSON.stringify(corpo) })
      .then(function (r) { if (!r.ok) throw new Error('Falha de comunicação (' + r.status + ').'); return r.json(); })
      .then(function (j) {
        if (j.sessao === false) { sair(true); throw new Error(j.erro); }
        if (!j.ok) throw new Error(j.erro || 'Erro desconhecido.');
        return j;
      });
  }

  /* ---------- login ---------- */
  function mostrarLogin(msg) {
    $('#tela-app').hidden = true;
    $('#tela-login').hidden = false;
    $('#login-erro').textContent = msg || (API ? '' : 'Plataforma ainda não conectada ao motor.');
  }
  function entrar(token, nome) {
    estado.token = token; estado.nome = nome;
    guardar('ww_token', token); guardar('ww_nome', nome);
    $('#nome-usuario').textContent = nome;
    $('#tela-login').hidden = true;
    $('#tela-app').hidden = false;
    carregarLista();
  }
  function sair(expirou) {
    if (!expirou && estado.token) api('sair').catch(function () {});
    estado.token = null; estado.progresso = null;
    guardar('ww_token', null); guardar('ww_nome', null);
    mostrarLogin(expirou ? 'Sua sessão expirou. Entre novamente.' : '');
  }
  $('#form-login').addEventListener('submit', function (e) {
    e.preventDefault();
    var btn = $('button', e.target);
    btn.disabled = true; $('#login-erro').textContent = '';
    api('login', { usuario: $('#login-usuario').value, senha: $('#login-senha').value })
      .then(function (j) { $('#login-senha').value = ''; entrar(j.token, j.nome); })
      .catch(function (err) { $('#login-erro').textContent = err.message; })
      .then(function () { btn.disabled = false; });
  });
  $('#btn-sair').addEventListener('click', function () { sair(false); });

  /* ---------- abas ---------- */
  $$('.abas button').forEach(function (b) {
    b.addEventListener('click', function () {
      $$('.abas button').forEach(function (x) { x.classList.toggle('ativa', x === b); });
      $$('.painel').forEach(function (p) { p.hidden = p.dataset.painel !== b.dataset.aba; });
      if (b.dataset.aba === 'progresso' && !estado.progresso) carregarProgresso();
    });
  });

  /* ---------- abas 1 e 2 ---------- */
  function montarPainel(aba) {
    var t = TEXTOS[aba];
    var el = $('[data-painel="' + aba + '"]');
    el.innerHTML =
      '<div class="card">' +
        '<div class="card-titulo"><span class="selo">' + t.selo + '</span><h2>' + t.titulo + '</h2></div>' +
        '<p class="ajuda">' + t.ajuda + '</p>' +
        '<form class="envio">' +
          '<input class="campo" type="url" required placeholder="https://docs.google.com/spreadsheets/d/..." aria-label="Link da planilha">' +
          '<button class="btn btn-primario" type="submit">Gerar planilha</button>' +
        '</form>' +
        '<div class="status" hidden></div>' +
      '</div>' +
      '<div class="card">' +
        '<div class="lista-topo">' +
          '<div class="card-titulo"><span class="selo">i</span><h3>Planilhas enviadas e geradas por mês</h3></div>' +
          '<div><button class="btn btn-peq" data-acao="atualizar">Atualizar lista</button> <span class="atualizado"></span></div>' +
        '</div>' +
        '<div class="lista"><p class="vazio">Carregando…</p></div>' +
      '</div>';

    $('form', el).addEventListener('submit', function (e) {
      e.preventDefault();
      var campo = $('input', e.target), btn = $('button', e.target), st = $('.status', el);
      btn.disabled = true;
      st.hidden = false; st.className = 'status processando';
      st.innerHTML = '<span class="spinner"></span>Conferindo a planilha e gerando a nova versão… isso pode levar até 2 minutos.';
      api('processar', { aba: aba, link: campo.value.trim() })
        .then(function (j) {
          var r = j.registro;
          st.className = 'status ok';
          st.innerHTML = '<b>' + esc(r.nomeGerado) + '</b> gerada. Semana ' + esc(r.semana) + ': ' +
            num(r.semFatura) + ' ' + t.rotSem + ' de ' + num(r.total) + ' ' + t.rotTotal + '. ' +
            '<a href="' + esc(r.linkGerado) + '" target="_blank" rel="noopener">Abrir planilha gerada</a>';
          campo.value = '';
          estado.progresso = null;
          carregarLista();
        })
        .catch(function (err) { st.className = 'status falha'; st.textContent = err.message; })
        .then(function () { btn.disabled = false; });
    });
    $('[data-acao="atualizar"]', el).addEventListener('click', carregarLista);
  }

  function carregarLista() {
    $$('.painel .lista').forEach(function (l) { if (!estado.registros.length) l.innerHTML = '<p class="vazio">Carregando…</p>'; });
    return api('listar').then(function (j) {
      estado.registros = j.registros || [];
      ['faturas', 'contratos'].forEach(renderLista);
      var hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      $$('.atualizado').forEach(function (s) { s.textContent = 'Atualizado às ' + hora; });
    }).catch(function (err) {
      $$('.painel .lista').forEach(function (l) { l.innerHTML = '<p class="vazio">' + esc(err.message) + '</p>'; });
    });
  }

  function renderLista(aba) {
    var t = TEXTOS[aba];
    var regs = estado.registros.filter(function (r) { return r.aba === aba; });
    var el = $('[data-painel="' + aba + '"] .lista');
    if (!regs.length) { el.innerHTML = '<p class="vazio">Nenhuma planilha enviada ainda.</p>'; return; }
    var porMes = {};
    regs.forEach(function (r) { (porMes[r.mes] = porMes[r.mes] || []).push(r); });
    el.innerHTML = Object.keys(porMes).sort().reverse().map(function (mes, i) {
      var itens = porMes[mes];
      return '<details class="mes"' + (i === 0 ? ' open' : '') + '><summary>' + esc(rotuloMes(mes)) +
        '<small>' + itens.length + (itens.length === 1 ? ' planilha' : ' planilhas') + '</small></summary>' +
        itens.map(function (r) {
          return '<div class="item"><div>' +
            '<div class="item-titulo">' + esc(r.nomeGerado) + '</div>' +
            '<div class="item-meta">Gerada em ' + esc(dataHora(r.criadoEm)) + ' por ' + esc(r.usuario) +
              ' · semanas ' + esc(r.semanas) + ' · origem: ' + esc(r.nomeOrigem) + '</div>' +
            '<div class="numeros"><span><b>' + num(r.total) + '</b> ' + t.rotTotal + '</span>' +
              '<span><b>' + num(r.comFatura) + '</b> ' + t.rotCom + '</span>' +
              '<span><b>' + num(r.semFatura) + '</b> ' + t.rotSem + '</span></div>' +
            '</div><div class="item-botoes">' +
            '<a class="btn btn-peq" href="' + esc(r.linkOrigem) + '" target="_blank" rel="noopener">Planilha enviada</a>' +
            '<a class="btn btn-peq btn-primario" href="' + esc(r.linkGerado) + '" target="_blank" rel="noopener">Planilha gerada</a>' +
            '</div></div>';
        }).join('') + '</details>';
    }).join('');
  }

  /* ---------- aba 3: progresso ---------- */
  function segmentado(id, chave) {
    $$('#' + id + ' button').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('#' + id + ' button').forEach(function (x) { x.classList.toggle('ativo', x === b); });
        estado[chave] = b.dataset.v;
        renderProgresso();
      });
    });
  }
  segmentado('prog-aba', 'progAba');
  segmentado('prog-periodo', 'progPeriodo');
  $('#prog-atualizar').addEventListener('click', carregarProgresso);

  function carregarProgresso() {
    $('#prog-conteudo').innerHTML = '<p class="vazio">Carregando…</p>';
    api('progresso').then(function (j) { estado.progresso = j.dados; renderProgresso(); })
      .catch(function (err) { $('#prog-conteudo').innerHTML = '<p class="vazio">' + esc(err.message) + '</p>'; });
  }

  function variacao(atual, anterior, menorMelhor) {
    if (anterior == null) return '<small>primeira conferência</small>';
    var d = atual - anterior;
    if (!d) return '<small>igual à conferência anterior</small>';
    var cls = (d > 0) === menorMelhor ? 'sobe' : 'desce';
    return '<small class="' + cls + '">' + (d > 0 ? '▲ ' : '▼ ') + num(Math.abs(d)) + ' em relação à anterior</small>';
  }

  function renderProgresso() {
    var dados = estado.progresso;
    if (!dados) return;
    var aba = estado.progAba, t = TEXTOS[aba], d = dados[aba];
    var serie = d[estado.progPeriodo] || [];
    var el = $('#prog-conteudo');
    if (!serie.length) { el.innerHTML = '<div class="card"><p class="vazio">Ainda não há conferências de ' + t.titulo + '.</p></div>'; return; }

    var u = serie[serie.length - 1], a = serie.length > 1 ? serie[serie.length - 2] : null;
    var cobertura = u.total ? Math.round(u.comFatura / u.total * 1000) / 10 : null;
    var cobAnt = a && a.total ? Math.round(a.comFatura / a.total * 1000) / 10 : null;
    var unidade = estado.progPeriodo === 'semanal' ? 'semana' : 'mês';

    var html = '<div class="kpis">' +
      '<div class="kpi"><span>Sem fatura na última conferência</span><strong>' + num(u.semFatura) + '</strong>' + variacao(u.semFatura, a && a.semFatura, true) + '</div>' +
      '<div class="kpi"><span>Resolvidos desde a anterior</span><strong>' + num(u.resolvidos) + '</strong><small>tinham pendência e passaram a ter fatura</small></div>' +
      '<div class="kpi"><span>Novos sem fatura</span><strong>' + num(u.novos) + '</strong><small>não estavam pendentes na anterior</small></div>' +
      '<div class="kpi"><span>Continuam sem fatura</span><strong>' + num(u.persistentes) + '</strong><small>pendentes nas duas últimas</small></div>' +
      '<div class="kpi"><span>Cobertura de faturas</span><strong>' + (cobertura == null ? '—' : cobertura.toLocaleString('pt-BR') + '%') + '</strong>' +
        (cobAnt == null ? '<small>' + num(u.comFatura) + ' de ' + num(u.total) + '</small>' : variacao(cobertura, cobAnt, false).replace(/(\d+(?:[.,]\d+)?) em/, '$1 p.p. em')) + '</div>' +
      '</div>';

    html += '<div class="card"><div class="card-titulo"><h3>Clientes sem fatura por ' + unidade + '</h3></div>' +
      '<div class="grafico">' + grafico(serie) + '</div></div>';

    html += '<div class="card"><div class="lista-topo"><div class="card-titulo"><h3>Números por ' + unidade + '</h3></div>' +
      '<button class="btn btn-peq" id="baixar-numeros">Baixar números (Excel)</button></div>' +
      '<div class="tabela-wrap"><table><thead><tr><th>Período</th><th class="num">' + cap(t.rotTotal) + '</th><th class="num">Com fatura</th>' +
      '<th class="num">Sem fatura</th><th class="num">Novos</th><th class="num">Resolvidos</th><th class="num">Continuam</th><th class="num">Cobertura</th></tr></thead><tbody>' +
      serie.slice().reverse().map(function (s) {
        return '<tr><td>' + esc(s.rotulo) + '</td><td class="num">' + num(s.total) + '</td><td class="num">' + num(s.comFatura) +
          '</td><td class="num">' + num(s.semFatura) + '</td><td class="num">' + num(s.novos) + '</td><td class="num">' + num(s.resolvidos) +
          '</td><td class="num">' + num(s.persistentes) + '</td><td class="num">' + (s.total ? (Math.round(s.comFatura / s.total * 1000) / 10).toLocaleString('pt-BR') + '%' : '—') + '</td></tr>';
      }).join('') + '</tbody></table></div></div>';

    html += '<div class="card"><div class="lista-topo"><div class="card-titulo"><h3>Clientes recorrentes sem fatura</h3></div>' +
      '<button class="btn btn-peq btn-primario" id="baixar-lista">Baixar lista (Excel)</button></div>' +
      '<p class="ajuda">Cada cliente que já apareceu sem fatura em alguma conferência, com quantas vezes isso aconteceu e se continua pendente na última (' + esc(rotuloPeriodo(d.ultimoPeriodo)) + ').</p>' +
      '<div class="lista-filtros"><input class="campo" id="busca" type="search" placeholder="Buscar por nome, instalação ou CPF/CNPJ">' +
      '<label><input type="checkbox" id="so-pendentes" checked> Só quem continua sem fatura</label>' +
      '<label>Mínimo de vezes <select class="campo" id="min-vezes"><option>1</option><option selected>2</option><option>3</option><option>4</option><option>6</option></select></label>' +
      '<span class="atualizado" id="contagem"></span></div>' +
      '<div class="rolagem"><table><thead><tr><th>Situação</th><th>Cliente</th><th>Instalação</th><th>CPF/CNPJ</th>' +
      '<th class="num">Vezes sem fatura</th><th class="num">Seguidas</th><th class="num">Meses</th><th>Primeira</th><th>Última</th><th>Detalhes</th></tr></thead>' +
      '<tbody id="corpo-clientes"></tbody></table></div></div>';

    el.innerHTML = html;
    ligarGrafico(serie);
    ['busca', 'so-pendentes', 'min-vezes'].forEach(function (id) { $('#' + id).addEventListener('input', renderClientes); });
    $('#baixar-lista').addEventListener('click', function () { baixarLista(filtrarClientes(), aba); });
    $('#baixar-numeros').addEventListener('click', function () { baixarNumeros(serie, aba); });
    renderClientes();
  }

  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  function filtrarClientes() {
    var d = estado.progresso[estado.progAba];
    var busca = ($('#busca').value || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    var so = $('#so-pendentes').checked, min = Number($('#min-vezes').value);
    return d.clientes.filter(function (c) {
      if (so && !c.pendenteAgora) return false;
      if (c.vezes < min) return false;
      if (!busca) return true;
      return (c.cliente + ' ' + c.instalacao + ' ' + c.cpfCnpj).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').indexOf(busca) >= 0;
    });
  }

  function renderClientes() {
    var lista = filtrarClientes();
    $('#contagem').textContent = num(lista.length) + (lista.length === 1 ? ' cliente' : ' clientes');
    var max = 500;
    $('#corpo-clientes').innerHTML = lista.slice(0, max).map(function (c) {
      return '<tr><td>' + (c.pendenteAgora ? '<span class="tag tag-pendente">Sem fatura</span>' : '<span class="tag tag-ok">Resolvido</span>') +
        '</td><td>' + esc(c.cliente || '—') + '</td><td>' + esc(c.instalacao) + '</td><td>' + esc(c.cpfCnpj) +
        '</td><td class="num">' + c.vezes + '</td><td class="num">' + c.seguidas + '</td><td class="num">' + c.meses +
        '</td><td>' + esc(rotuloPeriodo(c.primeira)) + '</td><td>' + esc(rotuloPeriodo(c.ultima)) + '</td><td class="quebra">' + esc(c.info) + '</td></tr>';
    }).join('') + (lista.length > max ? '<tr><td colspan="10" class="vazio">Mostrando ' + max + ' de ' + num(lista.length) + '. Baixe a lista para ver todos.</td></tr>' : '') +
      (lista.length ? '' : '<tr><td colspan="10" class="vazio">Nenhum cliente com esses filtros.</td></tr>');
  }

  /* gráfico de barras simples (uma série, um tom de azul) */
  function grafico(serie) {
    var W = 900, H = 260, m = { t: 16, r: 12, b: 34, l: 48 };
    var max = Math.max.apply(null, serie.map(function (s) { return s.semFatura; }).concat([1]));
    var passo = Math.pow(10, Math.floor(Math.log10(max)));
    var topo = Math.ceil(max / passo) * passo;
    var iw = W - m.l - m.r, ih = H - m.t - m.b;
    var bw = Math.min(56, iw / serie.length * 0.6), gap = iw / serie.length;
    var y = function (v) { return m.t + ih - v / topo * ih; };
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Clientes sem fatura por período">';
    for (var i = 0; i <= 4; i++) {
      var v = topo / 4 * i, yy = y(v);
      svg += '<line class="grade" x1="' + m.l + '" x2="' + (W - m.r) + '" y1="' + yy + '" y2="' + yy + '"/>' +
        '<text x="' + (m.l - 8) + '" y="' + (yy + 4) + '" text-anchor="end">' + num(Math.round(v)) + '</text>';
    }
    var cada = Math.ceil(serie.length / 12);
    serie.forEach(function (s, i) {
      var x = m.l + gap * i + (gap - bw) / 2, h = Math.max(0, m.t + ih - y(s.semFatura));
      var r = Math.min(4, bw / 2, h);
      var y0 = m.t + ih, y1 = y0 - h;
      var d = 'M' + x + ',' + y0 + 'V' + (y1 + r) + 'Q' + x + ',' + y1 + ' ' + (x + r) + ',' + y1 + 'H' + (x + bw - r) +
        'Q' + (x + bw) + ',' + y1 + ' ' + (x + bw) + ',' + (y1 + r) + 'V' + y0 + 'Z';
      svg += '<path class="barra" data-i="' + i + '" d="' + d + '"/>' +
        '<rect data-i="' + i + '" x="' + (m.l + gap * i) + '" y="' + m.t + '" width="' + gap + '" height="' + ih + '" fill="transparent"/>';
      if (i % cada === 0 || i === serie.length - 1) {
        svg += '<text x="' + (x + bw / 2) + '" y="' + (H - 12) + '" text-anchor="middle">' + esc(rotuloCurto(s)) + '</text>';
      }
    });
    return svg + '</svg>';
  }
  function rotuloCurto(s) {
    var p = String(s.mes).split('-');
    return MESES[Number(p[1]) - 1].slice(0, 3) + '/' + p[0].slice(2) + (estado.progPeriodo === 'semanal' ? ' S' + s.semana : '');
  }
  function ligarGrafico(serie) {
    var tip = $('#tooltip');
    $$('.grafico [data-i]').forEach(function (n) {
      if (n.tagName.toLowerCase() !== 'rect') return;
      n.addEventListener('mousemove', function (e) {
        var i = Number(n.dataset.i), s = serie[i];
        $$('.grafico .barra').forEach(function (b) { b.classList.toggle('foco', Number(b.dataset.i) === i); });
        tip.hidden = false;
        tip.innerHTML = '<b>' + esc(s.rotulo) + '</b><br>' + num(s.semFatura) + ' sem fatura de ' + num(s.total) +
          (s.resolvidos != null ? '<br>' + num(s.resolvidos) + ' resolvidos · ' + num(s.novos) + ' novos' : '');
        tip.style.left = Math.min(e.clientX + 14, window.innerWidth - 270) + 'px';
        tip.style.top = (e.clientY + 14) + 'px';
      });
      n.addEventListener('mouseleave', function () {
        tip.hidden = true;
        $$('.grafico .barra').forEach(function (b) { b.classList.remove('foco'); });
      });
    });
  }

  /* downloads em CSV que o Excel abre direto (ponto e vírgula + BOM) */
  function baixarCSV(nome, cab, linhas) {
    var cel = function (v) { v = String(v == null ? '' : v); return /[";\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var txt = '﻿' + [cab].concat(linhas).map(function (l) { return l.map(cel).join(';'); }).join('\r\n');
    var url = URL.createObjectURL(new Blob([txt], { type: 'text/csv;charset=utf-8' }));
    var a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }
  function hoje() { return new Date().toISOString().slice(0, 10); }
  function baixarLista(lista, aba) {
    baixarCSV('clientes-recorrentes-' + aba + '-' + hoje() + '.csv',
      ['Situação', 'Cliente', 'Instalação', 'CPF/CNPJ', 'Vezes sem fatura', 'Seguidas', 'Meses', 'Primeira', 'Última', 'Detalhes'],
      lista.map(function (c) {
        return [c.pendenteAgora ? 'Sem fatura' : 'Resolvido', c.cliente, '="' + c.instalacao + '"', c.cpfCnpj ? '="' + c.cpfCnpj + '"' : '',
          c.vezes, c.seguidas, c.meses, rotuloPeriodo(c.primeira), rotuloPeriodo(c.ultima), c.info];
      }));
  }
  function baixarNumeros(serie, aba) {
    baixarCSV('progresso-' + aba + '-' + estado.progPeriodo + '-' + hoje() + '.csv',
      ['Período', 'Total', 'Com fatura', 'Sem fatura', 'Novos', 'Resolvidos', 'Continuam', 'Cobertura (%)'],
      serie.map(function (s) {
        return [s.rotulo, s.total, s.comFatura, s.semFatura, s.novos == null ? '' : s.novos, s.resolvidos == null ? '' : s.resolvidos,
          s.persistentes == null ? '' : s.persistentes, s.total ? String(Math.round(s.comFatura / s.total * 1000) / 10).replace('.', ',') : ''];
      }));
  }

  /* ---------- início ---------- */
  montarPainel('faturas');
  montarPainel('contratos');
  var tk = ler('ww_token');
  if (tk) entrar(tk, ler('ww_nome') || ''); else mostrarLogin();
})();
