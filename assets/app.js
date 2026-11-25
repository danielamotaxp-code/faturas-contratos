(function () {
  'use strict';

  var A = window.Armazenamento, R = window.Regras, P = window.Planilhas;
  var USUARIOS_PADRAO = [['daniela', 'Daniela'], ['paula', 'Paula'], ['ytala', 'Ytala']];
  var ARQ_INDICE = 'dados/indice.json', ARQ_HIST = 'dados/historico.json', ARQ_CLIENTES = 'dados/clientes.json';
  var estado = { chaveGoogle: '', clientIdGoogle: '', modo: { faturas: 'link', contratos: 'link' }, alvo: {}, repo: null, nome: '', cofre: null, registros: [], progresso: null, progAba: 'faturas', progPeriodo: 'semanal' };

  var TEXTOS = {
    faturas: {
      selo: '1',
      titulo: 'Faturas Pendentes',
      ajuda: 'Cole o link da planilha de faturas do mês no Google (ex.: <b>Clientes Não Faturados – OUT 2026</b>) ou envie o arquivo. O mês é lido do <b>nome da planilha</b> e a <b>semana</b> pela data do envio (ou a que vocês escolherem). Entram só os <b>últimos 3 meses</b>. A plataforma gera uma planilha nova com as semanas do mês: clientes com fatura do mês ficam em <b>azul</b> (todas as linhas) e os que ainda não têm ficam em <b>branco</b>.',
      rotTotal: 'instalações', rotCom: 'com fatura do mês', rotSem: 'sem fatura (gerar)'
    },
    contratos: {
      selo: '2',
      titulo: 'Contratos sem Faturas',
      ajuda: 'Cole o link da planilha do Google (ou envie o arquivo) com as abas <b>Contratos – Semana N</b> e <b>Faturas – Semana N</b> (gateway de pagamento), ex.: <b>Contratos e Faturas – OUT 2026</b>. A plataforma compara as duas pela instalação e gera uma planilha nova só com os clientes que têm contrato com a WattWise e não têm fatura gerada.',
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
  var MESES = R.MESES_NOME;
  function rotuloMes(chave) { var p = String(chave).split('-'); return MESES[Number(p[1]) - 1] + ' de ' + p[0]; }
  function rotuloPeriodo(ch) {
    var m = /^(\d{4})-(\d{2})-S(\d+)$/.exec(ch || '');
    return m ? MESES[Number(m[2]) - 1].slice(0, 3) + '/' + m[1] + ' · S' + Number(m[3]) : '—';
  }
  function guardar(k, v) { try { v == null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch (e) { /* sem armazenamento */ } }
  function ler(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } }
  function arquivoSeguro(nome) {
    return String(nome).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[–—]/g, '-')
      .replace(/[^A-Za-z0-9 ._-]+/g, '').replace(/\s+/g, ' ').trim();
  }
  function carimbo(d) {
    var p = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
  }
  function baixarBytes(bytes, nome, tipo) {
    var url = URL.createObjectURL(new Blob([bytes], { type: tipo || 'application/octet-stream' }));
    var a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  function esperarTela() { return new Promise(function (ok) { requestAnimationFrame(function () { setTimeout(ok, 30); }); }); }

  /* ---------- telas ---------- */
  function mostrar(id) {
    ['tela-login', 'tela-config', 'tela-app'].forEach(function (t) { $('#' + t).hidden = t !== id; });
  }

  function iniciarLogin(msg) {
    mostrar('tela-login');
    $('#login-erro').textContent = msg || '';
    A.lerCofre().then(function (cofre) {
      estado.cofre = cofre;
      var sel = $('#login-usuario');
      if (!cofre) {
        $('#form-login').hidden = true;
        $('#login-sem-config').hidden = false;
        return;
      }
      $('#form-login').hidden = false;
      $('#login-sem-config').hidden = true;
      sel.innerHTML = '<option value="">Selecione</option>' + Object.keys(cofre.usuarios).map(function (k) {
        return '<option value="' + esc(k) + '">' + esc(cofre.usuarios[k].nome) + '</option>';
      }).join('');
    }).catch(function (e) { $('#login-erro').textContent = e.message; });
  }

  $('#form-login').addEventListener('submit', function (e) {
    e.preventDefault();
    var btn = $('button[type="submit"]', e.target), u = $('#login-usuario').value;
    var reg = estado.cofre && estado.cofre.usuarios[u];
    if (!reg) { $('#login-erro').textContent = 'Selecione o usuário.'; return; }
    btn.disabled = true; btn.textContent = 'Entrando…'; $('#login-erro').textContent = '';
    A.abrir(reg, $('#login-senha').value)
      .then(function (token) { $('#login-senha').value = ''; entrar(token, reg.nome); })
      .catch(function (err) { $('#login-erro').textContent = err.message; })
      .then(function () { btn.disabled = false; btn.textContent = 'Entrar'; });
  });

  // Planilhas das ouvidorias: esta aba aparece dentro de Ouvidorias Cemig (app.html?modo=planilhas-ouv), não em Faturas e Contratos
  var MODO_PLANILHAS = /[?&]modo=planilhas-ouv\b/.test(location.search);
  if (MODO_PLANILHAS) document.documentElement.classList.add('modo-planilhas-ouv');

  function entrar(token, nome) {
    estado.repo = new A.Repositorio(token);
    estado.nome = nome;
    guardar('ww_token', token); guardar('ww_nome', nome);
    $('#nome-usuario').textContent = nome;
    mostrar('tela-app');
    $('.abas button[data-aba="' + (MODO_PLANILHAS ? 'ouvidorias' : 'faturas') + '"]').click();
    carregarLista();
  }

  function sair() {
    if (window.PLAT) return window.PLAT.sair();
    estado.repo = null; estado.progresso = null; estado.registros = [];
    guardar('ww_token', null); guardar('ww_nome', null);
    iniciarLogin('');
  }
  $('#btn-sair').addEventListener('click', sair);

  /* ---------- configuração inicial / troca de senha ---------- */
  $$('[data-ir="config"]').forEach(function (b) { b.addEventListener('click', function () { mostrar('tela-config'); $('#config-erro').textContent = ''; }); });
  $('#config-voltar').addEventListener('click', function () {
    if (location.hash) history.replaceState(null, '', location.pathname + location.search);
    iniciarLogin('');
  });

  $('#form-config').addEventListener('submit', function (e) {
    e.preventDefault();
    var token = $('#config-token').value.trim(), s1 = $('#config-senha').value, s2 = $('#config-senha2').value;
    var erro = $('#config-erro'), btn = $('button[type="submit"]', e.target);
    erro.textContent = '';
    if (s1 !== s2) { erro.textContent = 'As senhas não conferem.'; return; }
    if (s1.length < 4) { erro.textContent = 'A senha precisa ter pelo menos 4 caracteres.'; return; }
    btn.disabled = true; btn.textContent = 'Configurando…';
    var dados = new A.Repositorio(token, A.REPO_DADOS), site = new A.Repositorio(token, A.REPO_SITE);
    dados.verificar().then(function () { return site.verificar(); }).then(function () {
      return Promise.all(USUARIOS_PADRAO.map(function (u) {
        return A.trancar(token, s1).then(function (r) { r.nome = u[1]; return [u[0], r]; });
      }));
    }).then(function (pares) {
      var cofre = { versao: 1, atualizadoEm: new Date().toISOString(), usuarios: {} };
      pares.forEach(function (p) { cofre.usuarios[p[0]] = p[1]; });
      return site.info('assets/cofre.json').then(function (meta) {
        return site.gravarBytes('assets/cofre.json', new TextEncoder().encode(JSON.stringify(cofre, null, 2)),
          'Cofre de acesso atualizado', meta && meta.sha);
      }).then(function () { estado.cofre = cofre; });
    }).then(function () {
      $('#config-token').value = '';
      entrar(token, 'Daniela');
      avisar('Plataforma configurada. O site leva cerca de 1 minuto para publicar o novo acesso para as demais usuárias.');
    }).catch(function (err) { erro.textContent = err.message; })
      .then(function () { btn.disabled = false; btn.textContent = 'Salvar configuração'; });
  });

  var ultimoAvisoSinc = '';
  function avisar(msg) {
    ultimoAvisoSinc = msg;
    var el = $('#aviso');
    el.textContent = msg; el.hidden = false;
    setTimeout(function () { el.hidden = true; }, 9000);
  }

  /* ---------- abas ---------- */
  $$('.abas button').forEach(function (b) {
    b.addEventListener('click', function () {
      $$('.abas button').forEach(function (x) { x.classList.toggle('ativa', x === b); });
      $$('.painel').forEach(function (p) { p.hidden = p.dataset.painel !== b.dataset.aba; });
      if (b.dataset.aba === 'progresso' && !estado.progresso) carregarProgresso();
    });
  });

  /* ---------- abas 1 e 2 ---------- */
  var SEM_ARQUIVO = 'Escolha ou arraste a planilha (.xlsx)';

  function montarPainel(aba) {
    var t = TEXTOS[aba];
    var el = $('[data-painel="' + aba + '"]');
    var opcoes = '<option value="">Automática (pela data: Semana ' + R.semanaDaData(new Date()) + ')</option>' +
      [1, 2, 3, 4].map(function (n) { return '<option value="' + n + '">Semana ' + n + '</option>'; }).join('');
    el.innerHTML =
      '<div class="card">' +
        '<div class="card-titulo"><span class="selo">' + t.selo + '</span><h2>' + t.titulo + '</h2></div>' +
        '<p class="ajuda">' + t.ajuda + '</p>' +
        '<div class="segmentado modo"><button type="button" data-modo="link" class="ativo">Link do Google Planilhas</button>' +
          '<button type="button" data-modo="arquivo">Arquivo (.xlsx)</button></div>' +
        '<form class="envio">' +
          '<input class="campo campo-link" type="url" placeholder="Cole o link da planilha do Google (https://docs.google.com/spreadsheets/d/...)" aria-label="Link da planilha do Google">' +
          '<label class="arquivo" hidden><input type="file" accept=".xlsx,.xls,.xlsm,.ods,.csv">' +
            '<span class="arquivo-texto">' + SEM_ARQUIVO + '</span></label>' +
          '<label class="semana">Semana <select class="campo sel-semana">' + opcoes + '</select></label>' +
          '<label class="semana">Mês <select class="campo sel-mes"><option value="">Automático (pelo nome; se não tiver, o mês atual)</option>' +
            MESES.map(function (m, i) { return '<option value="' + (i + 1) + '">' + m + '</option>'; }).join('') + '</select></label>' +
          '<label class="semana">Ano <select class="campo sel-ano">' + [-1, 0, 1].map(function (d) { var a = new Date().getFullYear() + d; return '<option value="' + a + '"' + (d === 0 ? ' selected' : '') + '>' + a + '</option>'; }).join('') + '</select></label>' +
          '<button class="btn btn-primario" type="submit">Gerar planilha</button>' +
        '</form>' +
        '<p class="dica dica-link">A planilha precisa estar em <b>Compartilhar → Acesso geral → Qualquer pessoa com o link (Leitor)</b>. Escolha o mês e o ano, ou deixe automático (lido do nome da planilha).</p>' +
        '<p class="dica dica-arquivo" hidden>O arquivo pode ter qualquer nome: escolha o mês e o ano. Com o Google conectado, a semana entra também na planilha do Google do mês.</p>' +
        '<p class="alvo" hidden></p>' +
        '<div class="status" hidden></div>' +
      '</div>' +
      '<div class="card">' +
        '<div class="lista-topo">' +
          '<div class="card-titulo"><span class="selo">i</span><h3>Planilhas enviadas e geradas por mês e semana</h3></div>' +
          '<div><button class="btn btn-peq" data-acao="atualizar">Atualizar lista</button> <span class="atualizado"></span></div>' +
        '</div>' +
        '<div class="lista"><p class="vazio">Carregando…</p></div>' +
      '</div>';

    var input = $('form input[type="file"]', el), rotulo = $('.arquivo-texto', el), caixa = $('.arquivo', el);
    input.addEventListener('change', function () {
      rotulo.textContent = input.files[0] ? input.files[0].name : SEM_ARQUIVO;
      caixa.classList.toggle('escolhido', !!input.files[0]);
    });
    ['dragover', 'dragenter'].forEach(function (ev) { caixa.addEventListener(ev, function () { caixa.classList.add('arrastando'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { caixa.addEventListener(ev, function () { caixa.classList.remove('arrastando'); }); });

    $$('.modo button', el).forEach(function (b) {
      b.addEventListener('click', function () {
        estado.modo[aba] = b.dataset.modo;
        $$('.modo button', el).forEach(function (x) { x.classList.toggle('ativo', x === b); });
        var link = b.dataset.modo === 'link';
        $('.campo-link', el).hidden = !link; caixa.hidden = link;
        $('.dica-link', el).hidden = !link; $('.dica-arquivo', el).hidden = link;
      });
    });

    $('form', el).addEventListener('submit', function (e) {
      e.preventDefault();
      var porLink = estado.modo[aba] === 'link', link = $('.campo-link', el).value.trim();
      if (porLink ? !link : !input.files[0]) {
        mostrarStatus(aba, 'falha', porLink ? 'Cole o link da planilha do Google.' : 'Escolha o arquivo da planilha.');
        return;
      }
      var btn = $('button[type="submit"]', e.target);
      btn.disabled = true;
      var semana = $('.sel-semana', e.target).value, mesSel = $('.sel-mes', e.target).value, anoSel = $('.sel-ano', e.target).value;
      var escolha = mesSel ? { mes: Number(mesSel), ano: Number(anoSel) } : null;
      (porLink ? enviarLink(aba, link, semana, escolha) : enviarArquivo(aba, input.files[0], semana, escolha)).then(function (ok) {
        if (ok) {
          e.target.reset(); $('.sel-ano', e.target).value = String(new Date().getFullYear()); rotulo.textContent = SEM_ARQUIVO; caixa.classList.remove('escolhido');
          estado.alvo[aba] = null; $('.alvo', el).hidden = true;
        }
      }).then(function () { btn.disabled = false; });
    });

    $('[data-acao="atualizar"]', el).addEventListener('click', carregarLista);
  }

  function mostrarStatus(aba, tipo, html) {
    var st = $('[data-painel="' + aba + '"] .status');
    st.hidden = false; st.className = 'status ' + tipo;
    st.innerHTML = html;
    st.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return st;
  }

  function passoDe(aba) {
    return function (txt) { mostrarStatus(aba, 'processando', '<span class="spinner"></span>' + esc(txt)); };
  }

  function conferirAlvo(aba, resultado) {
    var alvo = estado.alvo[aba];
    if (alvo && alvo.mes && resultado.mes.chave !== alvo.mes) {
      throw new Error('Essa planilha é de ' + resultado.mes.rotulo + ', mas você escolheu inserir em ' + R.rotuloMes(alvo.mes) + '. Confira o nome da planilha.');
    }
  }

  /** Nome usado para saber o mês: o mês/ano escolhidos vêm na frente; sem escolha e sem mês no nome, vale o mês atual. */
  function nomeComMes(nome, escolha) {
    if (escolha && escolha.mes) return R.MESES_NOME[escolha.mes - 1] + ' ' + escolha.ano + ' – ' + nome;
    if (R.mesDoNome(nome)) return nome;
    var h = new Date(); return R.MESES_NOME[h.getMonth()] + ' ' + h.getFullYear() + ' – ' + nome;
  }
  /** Semana do envio: a escolhida, ou a calculada pela data de hoje (dias 1–7 = 1, 8–14 = 2, 15–21 = 3, 22+ = 4). */
  function semanaDoEnvio(escolhida) { return Number(escolhida) || R.semanaDaData(new Date()); }

  /** Baixa as planilhas geradas das semanas anteriores do mesmo mês, para entrarem na planilha nova. */
  function carregarAnteriores(aba, mes, semana) {
    if (!mes) return Promise.resolve([]);
    var porSemana = {};
    estado.registros.filter(function (r) { return r.aba === aba && r.mes === mes.chave && Number(r.semana) < semana; })
      .forEach(function (r) {
        var atual = porSemana[r.semana];
        if (!atual || String(r.criadoEm) > String(atual.criadoEm)) porSemana[r.semana] = r;
      });
    return Promise.all(Object.keys(porSemana).map(function (n) {
      var reg = porSemana[n];
      return estado.repo.lerBytes(reg.caminhoGerado).then(function (bytes) {
        return bytes ? { n: Number(n), registro: reg, wb: window.XLSX.read(bytes, { type: 'array', cellDates: true }) } : null;
      });
    })).then(function (l) { return l.filter(Boolean); });
  }

  /** Lê o arquivo, confere, gera e salva. Devolve true se deu certo. */
  function enviarArquivo(aba, arq, semana, escolha) {
    var passo = passoDe(aba), resultado, nomeMes = nomeComMes(arq.name, escolha);
    passo('Lendo e conferindo "' + arq.name + '"…');
    var sem = semanaDoEnvio(semana), buffer;
    return concluir(aba, arq.arrayBuffer().then(function (buf) {
      buffer = buf;
      passo('Juntando as semanas anteriores do mês…');
      return carregarAnteriores(aba, R.mesDoNome(nomeMes), sem);
    }).then(function (anteriores) {
      var buf = buffer;
      passo('Conferindo "' + arq.name + '" (Semana ' + sem + ')…');
      return esperarTela().then(function () {
        resultado = P.processar(aba, nomeMes, buf, sem, anteriores);
        resultado.origemWb = function () { return P.lerArquivo(arq.name, buf); };
        conferirAlvo(aba, resultado);
        resultado.original = new Uint8Array(buf);
        passo('Salvando a planilha enviada e a gerada…');
        return salvarProcessamento(aba, arq.name, resultado, passo, '').then(function (reg) { return [reg, resultado]; }); // a planilha do Google do mês é preenchida em concluir()
      });
    }));
  }

  /** Lê a planilha pelo link do Google, confere, gera e salva. */
  function enviarLink(aba, link, semana, escolha) {
    var passo = passoDe(aba), resultado, nomeMes;
    passo('Lendo a planilha no Google Planilhas…');
    var sem = semanaDoEnvio(semana), planilha;
    return concluir(aba, window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      planilha = g; nomeMes = nomeComMes(g.titulo, escolha);
      passo('Juntando as semanas anteriores do mês…');
      return carregarAnteriores(aba, R.mesDoNome(nomeMes), sem);
    }).then(function (anteriores) {
      var g = planilha;
      passo('Conferindo "' + g.titulo + '" (Semana ' + sem + ')…');
      return esperarTela().then(function () {
        var wb = P.livroDeAbas(g.nomesAbas, g.abas);
        resultado = P.processarLivro(aba, nomeMes, wb, sem, anteriores);
        resultado.origemWb = function () { return wb; };
        conferirAlvo(aba, resultado);
        resultado.original = P.gravar(wb);
        passo('Salvando a cópia (Excel) da planilha como estava e a planilha gerada…');
        return salvarProcessamento(aba, g.titulo + '.xlsx', resultado, passo, g.link).then(function (reg) {
          return atualizarGoogle(aba, g, resultado, reg, passo).then(function () { return [reg, resultado]; });
        });
      });
    }));
  }

  /** Atualiza a própria planilha do Google (só com "Entrar com Google"). Nunca impede o resto de ser salvo. */
  function atualizarGoogle(aba, g, resultado, reg, passo) {
    if (!window.Google.conectado()) {
      reg.avisoGoogle = 'A planilha original do Google não foi alterada, porque ninguém entrou com Google. Clique em "Entrar com Google" para que a plataforma atualize a planilha direto no link.';
      return Promise.resolve();
    }
    var x = resultado.google;
    passo(aba === 'faturas' ? 'Pintando a aba "' + x.aba + '" na planilha do Google…' : 'Criando a aba "' + x.titulo + '" na planilha do Google…');
    var escrita = aba === 'faturas'
      ? window.Google.pintarAba(g, x.aba, x.colunas, x.indices, x.cores)
      : window.Google.escreverAba(g, x.titulo, x.valores);
    return escrita.then(function () {
      passo('Colocando o filtro de desconectados e inativos na planilha do Google…');
      return window.Google.filtrarAbas(g, x.filtros).catch(function () { /* sem o filtro, o resto continua valendo */ });
    }).then(function () {
      reg.googleAtualizada = aba === 'faturas' ? 'Aba "' + x.aba + '" pintada' : 'Aba "' + x.titulo + '" criada';
      return estado.repo.atualizarJSON(ARQ_INDICE, { registros: [] }, function (ind) {
        (ind.registros || []).forEach(function (r) { if (r.id === reg.id) r.googleAtualizada = reg.googleAtualizada; });
        return ind;
      }, 'Índice: planilha do Google atualizada (' + estado.nome + ')').then(function (ind) { estado.registros = ind.registros; });
    }).catch(function (err) {
      reg.avisoGoogle = 'A planilha gerada e a cópia foram salvas, mas não consegui atualizar a planilha do Google: ' + err.message;
    });
  }

  function concluir(aba, promessa) {
    var t = TEXTOS[aba], resultado;
    return promessa.then(function (par) {
      return paraPlanilhaDoMes(aba, par[0], par[1]).then(function () { return par; });
    }).then(function (par) {
      var reg = par[0]; resultado = par[1];
      var st = mostrarStatus(aba, 'ok', '<b>' + esc(reg.nomeGerado) + '</b> gerada. Semana ' + esc(reg.semana) + ': ' +
        num(reg.semFatura) + ' ' + t.rotSem + ' de ' + num(reg.total) + ' ' + t.rotTotal + '. ' +
        (reg.desconsiderados ? '<br>' + num(reg.desconsiderados) + (aba === 'faturas' ? ' instalações' : ' contratos') + ' desconectados ou inativos ficaram de fora. ' : '') +
        '<button class="btn btn-peq btn-primario" type="button">Baixar planilha gerada</button>' +
        (reg.googleAtualizada ? '<br>✔ Planilha do Google atualizada no mesmo link: ' + esc(reg.googleAtualizada) + '. <a href="' + esc(reg.linkOrigem) + '" target="_blank" rel="noopener">Abrir no Google</a>' : '') +
        (reg.avisoGoogle ? '<br><span class="aviso-google">' + esc(reg.avisoGoogle) + '</span>' : '') +
        (reg.planilhaMes ? '<br>✔ Também na planilha do Google de ' + esc(rotuloMes(reg.mes)) + ' (aba "' + esc(reg.planilhaMes.aba) + '"). <a href="' + esc(reg.planilhaMes.link) + '" target="_blank" rel="noopener">Abrir</a>' : '') +
        (reg.avisoMes ? '<br><span class="aviso-google">' + esc(reg.avisoMes) + '</span>' : ''));
      $('button', st).addEventListener('click', function () { baixarBytes(resultado.bytes, reg.nomeGerado + '.xlsx', XLSX_MIME); });
      estado.progresso = null;
      renderListas();
      return true;
    }).catch(function (err) {
      mostrarStatus(aba, 'falha', esc(err.message));
      return false;
    });
  }

  /* ---------- planilha do Google do mês: uma por mês, com todas as semanas (faturas e contratos) ---------- */
  var ARQ_MESES = 'dados/planilhas-mes.json';
  function lerMeses() { return estado.repo.lerJSON(ARQ_MESES, { meses: {} }).then(function (d) { estado.meses = d.meses || {}; return estado.meses; }); }
  function planilhaMes(chave) {
    return lerMeses().then(function (meses) {
      if (meses[chave]) return meses[chave];
      var p = String(chave).split('-'), titulo = 'Faturas e Contratos – ' + MESES[Number(p[1]) - 1] + ' ' + p[0];
      return window.Google.criarPlanilha(titulo).then(function (nova) {
        nova.criadaEm = new Date().toISOString(); nova.criadaPor = estado.nome;
        return estado.repo.atualizarJSON(ARQ_MESES, { meses: {} }, function (d) { d.meses = d.meses || {}; if (!d.meses[chave]) d.meses[chave] = nova; return d; }, 'Planilha do Google de ' + titulo + ' (' + estado.nome + ')')
          .then(function (d) { estado.meses = d.meses; desenharMeses(); return d.meses[chave]; });
      });
    });
  }
  /** Escreve uma semana (faturas ou contratos) na planilha do Google do mês: uma aba por semana. */
  function semanaNoMes(aba, semana, link, resultado) {
    var x = resultado.google, titulo;
    return window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      if (aba === 'contratos') { titulo = x.titulo; return window.Google.escreverAba(g, titulo, x.valores, 10).then(function () { return g; }); }
      titulo = 'Faturas - Semana ' + semana;
      var wb = resultado.origemWb(), ws = wb.Sheets[x.aba];
      var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
      for (var r0 = window.XLSX.utils.decode_range(ws['!ref']).s.r; r0 > 0; r0--) aoa.unshift([]);
      return window.Google.escreverAba(g, titulo, aoa, 10)
        .then(function () { return window.Google.lerPlanilha(link, estado.chaveGoogle); })
        .then(function (g2) { return window.Google.pintarAba(g2, titulo, x.colunas, x.indices, x.cores).then(function () { return g2; }); });
    }).then(function (g) {
      var vazia = ['Página1', 'Sheet1', 'Planilha1'].filter(function (n) { return g.folhas[n]; })[0];
      if (vazia && Object.keys(g.folhas).length > 1) return window.Google.lote(g.id, [{ deleteSheet: { sheetId: g.folhas[vazia].sheetId } }]).catch(function () {});
    }).then(function () { return titulo; });
  }
  /** Preenche a planilha do mês com as semanas já enviadas (a última de cada semana), refazendo a partir do arquivo enviado. */
  function preencherMes(chave, exceto) {
    var pm = (estado.meses || {})[chave]; if (!pm) return Promise.resolve(0);
    var ult = {};
    estado.registros.filter(function (r) { return r.mes === chave && r.id !== exceto; }).forEach(function (r) {
      var k = r.aba + '|' + r.semana; if (!ult[k] || String(r.criadoEm) > String(ult[k].criadoEm)) ult[k] = r;
    });
    var lista = Object.keys(ult).map(function (k) { return ult[k]; }).sort(function (a, b) { return a.semana - b.semana; }), n = 0;
    return lista.reduce(function (p, r) {
      return p.then(function () {
        return estado.repo.lerBytes(r.caminhoOrigem).then(function (bytes) {
          if (!bytes) return;
          var buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
          return carregarAnteriores(r.aba, { chave: chave }, Number(r.semana)).then(function (ant) {
            var nome = nomeComMes(r.nomeOrigem, { mes: Number(chave.split('-')[1]), ano: Number(chave.split('-')[0]) });
            var res = P.processar(r.aba, nome, buf, Number(r.semana), ant);
            res.origemWb = function () { return P.lerArquivo(r.nomeOrigem, buf); };
            return semanaNoMes(r.aba, r.semana, pm.link, res).then(function () { n++; });
          });
        }).catch(function () { /* uma semana com problema não impede as outras */ });
      });
    }, Promise.resolve()).then(function () { return n; });
  }
  function paraPlanilhaDoMes(aba, reg, resultado) {
    if (!window.Google.conectado()) { reg.avisoMes = 'Conecte o Google para a semana entrar também na planilha do Google do mês.'; return Promise.resolve(); }
    var x = resultado.google, titulo, link, nova = !(estado.meses || {})[reg.mes];
    return planilhaMes(reg.mes).then(function (pm) {
      link = pm.link;
      // mês novo: entram também as semanas que já tinham sido enviadas
      return (nova ? preencherMes(reg.mes, reg.id) : Promise.resolve()).then(function () { return semanaNoMes(aba, reg.semana, link, resultado); });
    }).then(function (t) { titulo = t;
      reg.planilhaMes = { aba: titulo, link: link };
      return estado.repo.atualizarJSON(ARQ_INDICE, { registros: [] }, function (ind) {
        (ind.registros || []).forEach(function (r) { if (r.id === reg.id) r.planilhaMes = reg.planilhaMes; });
        return ind;
      }, 'Índice: semana na planilha do mês (' + estado.nome + ')').then(function (ind) { estado.registros = ind.registros; });
    }).catch(function (err) { reg.avisoMes = 'A semana não entrou na planilha do Google do mês: ' + err.message; });
  }
  function desenharMeses() { renderListas(); }
  function desenharMesesAntigo() {
    var meses = estado.meses || {}, chaves = Object.keys(meses).sort().reverse();
    $$('.meses-google').forEach(function (el) {
      el.innerHTML = '<div class="card-titulo"><span class="selo">G</span><h3>Planilha do Google de cada mês</h3></div>' +
        '<p class="ajuda">Toda semana gerada (faturas e contratos) entra também numa planilha do Google do mês, uma aba por semana. Ela é criada sozinha no primeiro envio do mês (com o Google conectado) e fica aberta para a equipe editar.</p>' +
        (chaves.length ? '<ul class="lista-meses">' + chaves.map(function (k) { return '<li><b>' + esc(rotuloMes(k)) + ':</b> <a href="' + esc(meses[k].link) + '" target="_blank" rel="noopener">' + esc(meses[k].titulo) + '</a></li>'; }).join('') + '</ul>'
          : '<p class="vazio">Ainda nenhuma. A do mês é criada no próximo envio.</p>');
    });
  }

  /** Grava no repositório privado: arquivo enviado, arquivo gerado, histórico, clientes e índice. */
  function salvarProcessamento(aba, nomeArquivo, r, passo, linkOrigem) {
    var repo = estado.repo, agora = new Date(), id = carimbo(agora);
    var pasta = 'dados/arquivos/' + r.mes.chave + '/';
    var ext = (/\.[a-z0-9]+$/i.exec(nomeArquivo) || ['.xlsx'])[0].toLowerCase();
    var caminhoOrigem = pasta + 'enviadas/' + id + ' ' + arquivoSeguro(nomeArquivo.replace(/\.[a-z0-9]+$/i, '')) + ext;
    var caminhoGerado = pasta + 'geradas/' + id + ' ' + arquivoSeguro(r.nomeGerado) + '.xlsx';
    var ult = r.semanas[r.semanas.length - 1];
    var registro = {
      id: id, criadoEm: agora.toISOString(), usuario: estado.nome, aba: aba, mes: r.mes.chave, mesRotulo: r.mes.rotulo,
      semana: ult.n, semanas: (r.semanasNoArquivo || r.semanas.map(function (s) { return s.n; })).join(', '),
      abaOrigem: r.abaOrigem || '', faturasDoMes: ult.faturasDoMes == null ? null : ult.faturasDoMes,
      nomeOrigem: nomeArquivo, caminhoOrigem: caminhoOrigem, linkOrigem: linkOrigem || '', nomeGerado: r.nomeGerado, caminhoGerado: caminhoGerado,
      total: ult.total, comFatura: ult.comFatura, semFatura: ult.pendentes.length,
      desconsiderados: ult.desconsiderados || 0
    };
    var msg = r.nomeGerado + ' (' + estado.nome + ')';
    return repo.gravarBytes(caminhoOrigem, r.original, 'Enviada: ' + nomeArquivo + ' (' + estado.nome + ')')
      .then(function () { return repo.gravarBytes(caminhoGerado, r.bytes, 'Gerada: ' + msg); })
      .then(function () {
        passo('Atualizando o histórico de progresso…');
        return repo.atualizarJSON(ARQ_HIST, { faturas: {}, contratos: {} }, function (h) {
          h[aba] = h[aba] || {};
          r.semanas.forEach(function (s) {
            var ch = R.chavePeriodo(r.mes.chave, s.n);
            var p = h[aba][ch] = h[aba][ch] && h[aba][ch].versoes ? h[aba][ch] : { versoes: {} };
            p.versoes[id] = {
              mes: r.mes.chave, semana: s.n, total: s.total, comFatura: s.comFatura, semFatura: s.pendentes.length,
              pendentes: s.pendentes.map(function (x) { return x.instalacao; }), criadoEm: registro.criadoEm
            };
          });
          return h;
        }, 'Histórico: ' + msg);
      })
      .then(function () {
        return repo.atualizarJSON(ARQ_CLIENTES, { faturas: {}, contratos: {} }, function (c) {
          c[aba] = c[aba] || {};
          r.semanas.forEach(function (s) {
            s.pendentes.forEach(function (p) { c[aba][p.instalacao] = { cliente: p.cliente, cpfCnpj: p.cpfCnpj, info: p.info }; });
          });
          return c;
        }, 'Clientes: ' + msg);
      })
      .then(function () {
        passo('Registrando na lista de planilhas…');
        return repo.atualizarJSON(ARQ_INDICE, { registros: [] }, function (ind) {
          ind.registros = [registro].concat(ind.registros || []);
          return ind;
        }, 'Índice: ' + msg);
      })
      .then(function (ind) { estado.registros = ind.registros; return registro; });
  }

  /** Exclui um envio: tira da lista, apaga os dois arquivos e devolve o histórico da semana à versão anterior. */
  function excluirRegistro(reg) {
    var repo = estado.repo, msg = 'Excluída: ' + reg.nomeGerado + ' (' + estado.nome + ')';
    return repo.atualizarJSON(ARQ_INDICE, { registros: [] }, function (ind) {
      ind.registros = (ind.registros || []).filter(function (r) { return r.id !== reg.id; });
      return ind;
    }, msg).then(function (ind) {
      estado.registros = ind.registros;
      return repo.atualizarJSON(ARQ_HIST, { faturas: {}, contratos: {} }, function (h) {
        var porAba = h[reg.aba] || {};
        Object.keys(porAba).forEach(function (ch) {
          var p = porAba[ch];
          if (p.versoes && p.versoes[reg.id]) {
            delete p.versoes[reg.id];
            if (!Object.keys(p.versoes).length) delete porAba[ch];
          } else if (!p.versoes && p.registro === reg.id) {
            delete porAba[ch];
          }
        });
        return h;
      }, 'Histórico: ' + msg);
    }).then(function () { return repo.apagar(reg.caminhoGerado, msg); })
      .then(function () { return repo.apagar(reg.caminhoOrigem, msg); });
  }

  function carregarLista() {
    if (!estado.repo) return;
    $$('.painel .lista').forEach(function (l) { if (!estado.registros.length) l.innerHTML = '<p class="vazio">Carregando…</p>'; });
    lerMeses().then(desenharMeses).catch(function () { /* sem a lista de meses */ });
    return estado.repo.lerJSON(ARQ_INDICE, { registros: [] }).then(function (ind) {
      estado.registros = ind.registros || [];
      renderListas();
    }).catch(function (err) {
      $$('.painel .lista').forEach(function (l) { l.innerHTML = '<p class="vazio">' + esc(err.message) + '</p>'; });
    });
  }

  function renderListas() {
    ['faturas', 'contratos'].forEach(renderLista);
    var hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    $$('.atualizado').forEach(function (s) { if (!s.id) s.textContent = 'Atualizado às ' + hora; });
  }

  function itemHtml(r, t) {
    return '<div class="item" data-id="' + esc(r.id) + '"><div>' +
      '<div class="item-titulo">' + esc(r.nomeGerado) + (r.googleAtualizada ? ' <span class="tag tag-ok" title="' + esc(r.googleAtualizada) + '">Google atualizado</span>' : '') + '</div>' +
      '<div class="item-meta">Gerada em ' + esc(dataHora(r.criadoEm)) + ' por ' + esc(r.usuario) +
        ' · semanas ' + esc(r.semanas) + ' · enviada: ' + esc(r.nomeOrigem) + '</div>' +
      '<div class="numeros"><span><b>' + num(r.total) + '</b> ' + t.rotTotal + '</span>' +
        '<span><b>' + num(r.comFatura) + '</b> ' + t.rotCom + '</span>' +
        '<span><b>' + num(r.semFatura) + '</b> ' + t.rotSem + '</span></div>' +
      '</div><div class="item-botoes">' +

      '<button class="btn btn-peq" data-baixar="' + esc(r.caminhoOrigem) + '" data-nome="' + esc(r.nomeOrigem) + '">Baixar enviada</button>' +
      '<button class="btn btn-peq btn-primario" data-baixar="' + esc(r.caminhoGerado) + '" data-nome="' + esc(r.nomeGerado + '.xlsx') + '">Baixar gerada</button>' +
      '<button class="btn btn-peq btn-perigo" data-excluir="' + esc(r.id) + '">Excluir</button>' +
      '</div></div>';
  }

  function renderLista(aba) {
    var t = TEXTOS[aba];
    var regs = estado.registros.filter(function (r) { return r.aba === aba; });
    var el = $('[data-painel="' + aba + '"] .lista');
    if (!regs.length) {
      el.innerHTML = '<p class="vazio">Nenhuma planilha enviada ainda. Use o campo acima para enviar a primeira.</p>';
      return;
    }
    var porMes = {};
    regs.forEach(function (r) { (porMes[r.mes] = porMes[r.mes] || []).push(r); });
    el.innerHTML = Object.keys(porMes).sort().reverse().map(function (mes, i) {
      var itens = porMes[mes], porSemana = {};
      itens.forEach(function (r) { (porSemana[r.semana] = porSemana[r.semana] || []).push(r); });
      var maxSemana = Math.max.apply(null, [4].concat(itens.map(function (r) { return Number(r.semana); })));
      var semanas = '';
      for (var s = 1; s <= maxSemana; s++) {
        var lista = porSemana[s] || [];
        semanas += '<div class="semana-bloco">' +
          '<div class="semana-topo"><b>Semana ' + s + '</b>' +
          '<span class="semana-qtd">' + (lista.length ? lista.length + (lista.length === 1 ? ' planilha' : ' planilhas') : 'sem planilha') + '</span>' +
          '<button class="btn btn-peq" data-inserir="' + s + '" data-mes="' + mes + '">+ Inserir planilha</button></div>' +
          (lista.length ? lista.map(function (r) { return itemHtml(r, t); }).join('') : '<p class="semana-vazia">Nenhuma planilha nesta semana.</p>') +
          '</div>';
      }
      var pm = (estado.meses || {})[mes];
      var google = '<div class="mes-google">' + (pm ? '📗 <b>Planilha do Google de ' + esc(rotuloMes(mes)) + '</b> (todas as semanas, uma aba por semana): <a class="btn btn-peq btn-primario" href="' + esc(pm.link) + '" target="_blank" rel="noopener">Abrir no Google</a>'
        : 'Este mês ainda não tem planilha do Google. <button class="btn btn-peq" type="button" data-criar-mes="' + esc(mes) + '">Criar a planilha do Google de ' + esc(rotuloMes(mes)) + '</button> <small>(precisa do Google conectado; entram as semanas já enviadas)</small>') + '</div>';
      return '<details class="mes"' + (i === 0 ? ' open' : '') + '><summary>' + esc(rotuloMes(mes)) +
        '<small>' + itens.length + (itens.length === 1 ? ' planilha' : ' planilhas') + '</small></summary>' + google + semanas + '</details>';
    }).join('');

    $$('[data-baixar]', el).forEach(function (b) {
      b.addEventListener('click', function () {
        var txt = b.textContent;
        b.disabled = true; b.textContent = 'Baixando…';
        estado.repo.lerBytes(b.dataset.baixar).then(function (bytes) {
          if (!bytes) throw new Error('Arquivo não encontrado.');
          baixarBytes(bytes, b.dataset.nome, /\.csv$/i.test(b.dataset.nome) ? 'text/csv' : XLSX_MIME);
        }).catch(function (err) { avisar(err.message); })
          .then(function () { b.disabled = false; b.textContent = txt; });
      });
    });

    $$('[data-criar-mes]', el).forEach(function (b) {
      b.addEventListener('click', function () {
        if (!window.Google.conectado()) { avisar('Conecte o Google primeiro.'); return; }
        b.disabled = true; b.textContent = 'Criando e preenchendo…';
        planilhaMes(b.dataset.criarMes).then(function () { return preencherMes(b.dataset.criarMes); })
          .then(function (n) { avisar('Planilha do Google do mês criada' + (n ? ' com ' + n + ' semana(s)' : '') + '.'); renderListas(); })
          .catch(function (e) { avisar('Não consegui criar: ' + e.message); b.disabled = false; });
      });
    });
    $$('[data-inserir]', el).forEach(function (b) {
      b.addEventListener('click', function () {
        var painel = $('[data-painel="' + aba + '"]');
        estado.alvo[aba] = { mes: b.dataset.mes, semana: b.dataset.inserir };
        $('form .sel-semana', painel).value = b.dataset.inserir;
        var pm = String(b.dataset.mes).split('-'); $('form .sel-mes', painel).value = String(Number(pm[1])); $('form .sel-ano', painel).value = pm[0];
        var alvo = $('.alvo', painel);
        alvo.hidden = false;
        alvo.innerHTML = 'Inserindo na <b>Semana ' + esc(b.dataset.inserir) + '</b> de <b>' + esc(rotuloMes(b.dataset.mes)) +
          '</b>. Cole o link ou escolha o arquivo e clique em Gerar planilha. <button type="button" class="btn btn-link">Cancelar</button>';
        $('button', alvo).addEventListener('click', function () { estado.alvo[aba] = null; alvo.hidden = true; $('form .sel-semana', painel).value = ''; $('form .sel-mes', painel).value = ''; });
        $('.card', painel).scrollIntoView({ behavior: 'smooth', block: 'start' });
        var campo = estado.modo[aba] === 'link' ? $('.campo-link', painel) : null;
        if (campo) setTimeout(function () { campo.focus(); }, 300);
      });
    });

    // Excluir pede confirmação no próprio botão (2 cliques)
    $$('[data-excluir]', el).forEach(function (b) {
      b.addEventListener('click', function () {
        var reg = estado.registros.filter(function (r) { return r.id === b.dataset.excluir; })[0];
        if (!reg) return;
        if (!b.classList.contains('confirmar')) {
          b.classList.add('confirmar');
          b.textContent = 'Confirmar exclusão';
          setTimeout(function () { if (b.isConnected && !b.disabled) { b.classList.remove('confirmar'); b.textContent = 'Excluir'; } }, 5000);
          return;
        }
        b.disabled = true; b.textContent = 'Excluindo…';
        excluirRegistro(reg).then(function () {
          estado.progresso = null;
          renderListas();
          avisar('"' + reg.nomeGerado + '" foi excluída.');
        }).catch(function (err) {
          avisar(err.message);
          carregarLista();
        });
      });
    });
  }

  function alertar(msg) { avisar(msg); }

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
    Promise.all([estado.repo.lerJSON(ARQ_HIST, {}), estado.repo.lerJSON(ARQ_CLIENTES, {})])
      .then(function (r) { estado.progresso = R.calcularProgresso(r[0], r[1]); renderProgresso(); })
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

  /* ---------- configurações (chave do Google) ---------- */
  function carregarChaveGoogle() {
    return fetch('assets/google.json?v=' + Date.now(), { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : {}; })
      .then(function (j) { estado.chaveGoogle = j.chave || ''; estado.clientIdGoogle = j.clientId || ''; atualizarConfigGoogle(); })
      .catch(function () { /* sem chave ainda */ });
  }
  function atualizarConfigGoogle() {
    var ok = '<span class="tag tag-ok">Configurado</span>', nao = '<span class="tag tag-pendente">Ainda não configurado</span>';
    $('#google-situacao').innerHTML = estado.chaveGoogle ? ok : nao;
    $('#cliente-situacao').innerHTML = estado.clientIdGoogle ? ok : nao;
    var b = $('#btn-google');
    b.hidden = !estado.clientIdGoogle;
    b.textContent = window.Google.conectado() ? 'Google conectado ✔' : 'Entrar com Google';
    b.classList.toggle('conectado', window.Google.conectado());
  }
  $('#btn-google').addEventListener('click', function () {
    if (window.Google.conectado()) { window.Google.sair(); atualizarConfigGoogle(); avisar('Google desconectado.'); return; }
    window.Google.entrar(estado.clientIdGoogle).then(function () {
      atualizarConfigGoogle();
      avisar('Google conectado. Agora a plataforma atualiza a planilha direto no link.');
    }).catch(function (err) { avisar(err.message); });
  });
  function salvarConfigGoogle(mudar) {
    var site = new A.Repositorio(estado.repo.token, A.REPO_SITE);
    return site.atualizarJSON('assets/google.json', {}, function (j) { mudar(j); return j; }, 'Configuração do Google atualizada');
  }
  $('#btn-config').addEventListener('click', function () {
    $$('.abas button').forEach(function (x) { x.classList.remove('ativa'); });
    $$('.painel').forEach(function (p) { p.hidden = p.dataset.painel !== 'config'; });
    atualizarConfigGoogle();
  });
  $('#form-google').addEventListener('submit', function (e) {
    e.preventDefault();
    var chave = $('#google-chave').value.trim(), btn = $('button[type="submit"]', e.target), msg = $('#google-msg');
    if (!/^AIza[0-9A-Za-z_-]{30,}$/.test(chave)) { msg.className = 'erro'; msg.textContent = 'A chave do Google começa com "AIza" e tem cerca de 39 caracteres.'; return; }
    btn.disabled = true; msg.className = ''; msg.textContent = 'Salvando…';
    salvarConfigGoogle(function (j) { j.chave = chave; }).then(function () {
      estado.chaveGoogle = chave; atualizarConfigGoogle();
      $('#google-chave').value = '';
      msg.textContent = 'Chave salva. Já dá para enviar planilhas pelo link (para as demais usuárias, em cerca de 1 minuto).';
    }).catch(function (err) { msg.className = 'erro'; msg.textContent = err.message; })
      .then(function () { btn.disabled = false; });
  });

  $('#form-cliente').addEventListener('submit', function (e) {
    e.preventDefault();
    var id = $('#cliente-id').value.trim(), btn = $('button[type="submit"]', e.target), msg = $('#cliente-msg');
    if (!/^[0-9]+-[0-9a-z]+\.apps\.googleusercontent\.com$/.test(id)) { msg.className = 'erro'; msg.textContent = 'O ID do cliente termina com ".apps.googleusercontent.com".'; return; }
    btn.disabled = true; msg.className = ''; msg.textContent = 'Salvando…';
    salvarConfigGoogle(function (j) { j.clientId = id; }).then(function () {
      estado.clientIdGoogle = id; atualizarConfigGoogle();
      $('#cliente-id').value = '';
      msg.textContent = 'ID salvo. O botão "Entrar com Google" já aparece no topo.';
    }).catch(function (err) { msg.className = 'erro'; msg.textContent = err.message; })
      .then(function () { btn.disabled = false; });
  });


  /* ---------- aba Ouvidorias: listas de quem precisa de ouvidoria (Reclamação 6) ---------- */
  var ARQ_OUV = 'dados/ouvidorias.json';
  var PLANILHA_OUV_PADRAO = 'https://docs.google.com/spreadsheets/d/1T55s2hcu4n0_O50vk5ZKbidlbMpnwzs00btX-IMLQQ4/edit';
  // Formato das abas da planilha de ouvidorias (uma aba por data). As 4 últimas colunas a plataforma vai atualizando.
  var COLS_OUV = ['Colaborador', 'Cliente', 'Usina', 'Instalação', 'Unid. Consumidora', 'Mês Referencia', 'CPF / CNPJ', 'Data Vencimento Boleto',
    'Parceiro', 'Associado', 'Status Contrato Faturamento', 'Status Contrato Atual', 'Modalidade da Fatura', 'Status da Fatura', 'Geração de Faturas', 'Classificação'];
  var COLS_AUTO = ['Ouvidoria', 'Protocolo da ouvidoria', 'Data da ouvidoria', 'Ouvidoria na plataforma'];
  var PORTAL_OUV = 'https://danielamotaxp-code.github.io/central-atendimento/#ouvidorias/cad/';
  var ouv = { dados: null, previa: null, regs: null, filtro: {} };
  function hojeISO() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function isoParaBR(iso) { var p = String(iso || '').split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : iso; }
  function nrm(x) { return R.normalizar(x); }

  function lerTabelaColada(txt) {
    var linhas = String(txt || '').replace(/\r/g, '').split('\n');
    var tab = linhas.some(function (l) { return l.indexOf('\t') >= 0; });
    var m = linhas.map(function (l) {
      if (tab) return l.split('\t');
      var out = [], cur = '', q = false; // CSV simples com aspas
      for (var i = 0; i < l.length; i++) {
        var c = l[i];
        if (c === '"') { if (q && l[i + 1] === '"') { cur += '"'; i++; } else q = !q; }
        else if (c === ',' && !q) { out.push(cur); cur = ''; } else cur += c;
      }
      out.push(cur); return out;
    });
    return limparTabela(m);
  }
  /** Acha o cabeçalho, tira colunas vazias e dá nome à coluna da classificação (que costuma vir sem título). */
  function limparTabela(m) {
    m = (m || []).map(function (l) { return (l || []).map(function (c) { return c == null ? '' : (c instanceof Date ? c.toLocaleDateString('pt-BR') : String(c).trim()); }); })
      .filter(function (l) { return l.some(function (c) { return c !== ''; }); });
    var h = m.findIndex(function (l) { return l.some(function (c) { return /^cliente$|instala/.test(nrm(c)); }); });
    if (h < 0) return null;
    var cab = m[h], corpo = m.slice(h + 1), n = Math.max.apply(null, m.map(function (l) { return l.length; }));
    var usar = [];
    for (var c = 0; c < n; c++) {
      var temCab = (cab[c] || '') !== '', temDado = corpo.some(function (l) { return (l[c] || '') !== ''; });
      if (temCab || temDado) usar.push(c);
    }
    var ultimaSemNome = usar.filter(function (c) { return !(cab[c] || ''); }).pop();
    var novoCab = usar.map(function (c, i) { return cab[c] || (c === ultimaSemNome ? 'Classificação' : 'Coluna ' + (i + 1)); });
    var linhas = corpo.map(function (l) { return usar.map(function (c) { return l[c] || ''; }); })
      .filter(function (l) { return l.some(function (x) { return x !== ''; }); });
    return { cab: novoCab, linhas: linhas };
  }
  /** Planilha auditada → formato da aba de ouvidorias: só quem foi classificado, colunas na ordem certa. */
  function paraFormatoOuv(t) {
    if (!t) return null;
    var c = colunasOuv(t.cab);
    var linhas = t.linhas;
    if (c.classif >= 0 && linhas.some(function (l) { return l[c.classif]; })) linhas = linhas.filter(function (l) { return l[c.classif]; });
    var idx = COLS_OUV.map(function (nome) {
      if (nome === 'Classificação') return c.classif;
      if (nome === 'Colaborador') return c.colab;
      if (window.Auditada) return window.Auditada.colunaDe(t.cab, nome);
      var n = nrm(nome); return t.cab.findIndex(function (h) { return nrm(h) === n; });
    });
    var usar = COLS_OUV.map(function (nome, i) { return { nome: nome, i: idx[i] }; })
      .filter(function (o) { return o.i >= 0 || o.nome === 'Colaborador' || o.nome === 'Classificação'; });
    return { cab: usar.map(function (o) { return o.nome; }).concat(COLS_AUTO),
      linhas: linhas.map(function (l) { return usar.map(function (o) { return o.i >= 0 ? (l[o.i] || '') : ''; }).concat(COLS_AUTO.map(function () { return ''; })); }),
      descartadas: t.linhas.length - linhas.length };
  }
  function colunasOuv(cab) {
    var col = function (re) { return R.coluna(cab, re); };
    return { colab: col([/colaborador/, /responsavel/]), cliente: col([/^cliente$/, /^cliente/, /titular/, /^nome$/, /^nome\s*(do\s*)?(cliente|consumidor|titular)/, /razao\s*social/, /consumidor/, /^nome(?!.*(usina|parceiro|backoffice|associado|colaborador))/]),
      usina: col([/^usina/]), inst: col([/^instalacao$/, /numero\s*(da\s*)?instala/, /instala/]),
      uc: col([/unid.*consum/, /^uc$/]), doc: col([/cpf|cnpj/]), mes: col([/mes\s*refer/]),
      classif: col([/classifica/, /^observ/, /^motivo/]) };
  }
  function contar(linhas, i, vazio) {
    var c = {}; if (i < 0) return c;
    linhas.forEach(function (l) { var v = l[i] || vazio || '(sem classificação)'; c[v] = (c[v] || 0) + 1; });
    return c;
  }
  function chips(cont, cinza) {
    return Object.keys(cont).sort(function (a, b) { return cont[b] - cont[a]; })
      .map(function (k) { return '<span class="ouv-chip' + (cinza ? ' cinza' : '') + '">' + esc(k) + ': <b>' + num(cont[k]) + '</b></span>'; }).join('');
  }

  // explica o que a plataforma entendeu da planilha (aba, colunas de classificação, legenda das cores)
  function mostrarInterpretacao(x) {
    var t = x.t, leg = Object.keys(t.legenda || {});
    $('#ouv-previa').insertAdjacentHTML('afterbegin', '<p class="dica">Aba lida: <b>' + esc(x.n) + '</b>' +
      (t.colunasClassificacao.length ? ' · classificação nas colunas <b>' + esc(t.colunasClassificacao.join(', ')) + '</b>' : '') +
      (leg.length ? ' · legenda das cores: ' + leg.map(function (k) { return '<b>' + esc(k) + '</b> = ' + esc(t.legenda[k]); }).join('; ') : '') +
      (t.semClassificacao ? ' · ' + num(t.semClassificacao) + ' linha(s) sem classificação ficaram de fora' : '') +
      (t.descartadas ? ' · ' + num(t.descartadas) + ' desconectado(s)/inativo(s) fora' : '') + '</p>');
  }
  function mostrarPrevia(t) {
    ouv.previa = t;
    var box = $('#ouv-previa');
    if (!t) { box.innerHTML = arguments.length ? '<p class="erro">Não achei o cabeçalho (coluna "Cliente" ou "Instalação") na planilha.</p>' : ''; $('#ouv-salvar').disabled = true; return; }
    var c = colunasOuv(t.cab);
    var faltam = [c.inst < 0 && 'Instalação'].filter(Boolean);
    box.innerHTML = '<div class="ouv-resumo"><span class="ouv-chip"><b>' + num(t.linhas.length) + '</b> clientes</span>' + chips(contar(t.linhas, c.classif)) + '</div>' +
      (c.colab >= 0 ? '<div class="ouv-resumo">' + chips(contar(t.linhas, c.colab, '(sem colaborador)'), true) + '</div>' : '') +
      (faltam.length ? '<p class="erro">Faltou a coluna ' + faltam.join(' e ') + '. Colunas encontradas: ' + esc(t.cab.filter(Boolean).join(', ')) + '.</p>' : '') +
      (c.cliente < 0 ? '<p class="dica">Não achei a coluna com o nome do cliente; a lista será salva sem ela (o cliente é achado pela instalação).</p>' : '') +
      (c.classif < 0 ? '<p class="dica">Não achei a coluna da classificação. A lista será salva mesmo assim.</p>' : '');
    if (t.descartadas) box.insertAdjacentHTML('beforeend', '<p class="dica">' + num(t.descartadas) + ' linhas sem classificação (fatura gerada normalmente) ficaram de fora.</p>');
    $('#ouv-salvar').disabled = !!faltam.length || !t.linhas.length;
  }
  $('#ouv-ler-link').addEventListener('click', function () {
    var link = $('#ouv-link').value.trim(), btn = this;
    if (!/\/spreadsheets\/d\//.test(link)) { $('#ouv-previa').innerHTML = '<p class="erro">Cole o link da planilha auditada (Google Planilhas).</p>'; return; }
    btn.disabled = true; $('#ouv-previa').innerHTML = '<p class="dica">Lendo a planilha…</p>';
    window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      // as cores só vêm com o Google conectado; sem elas, lê só as colunas
      return Promise.all(g.nomesAbas.map(function (n) {
        return (window.Google.conectado() ? window.Google.coresAba(g.id, n, estado.chaveGoogle).catch(function () { return []; }) : Promise.resolve([]))
          .then(function (hex) { return { n: n, cores: hex.map(function (l) { return (l || []).map(window.Auditada.nomeCor); }) }; });
      })).then(function (cs) {
        // aba com mais linhas classificadas (a última auditada)
        var melhor = null;
        g.nomesAbas.forEach(function (n, i) {
          var t = window.Auditada.interpretar(g.abas[n], cs[i].cores); if (!t) return;
          if (!melhor || t.linhas.length > melhor.t.linhas.length || (t.linhas.length === melhor.t.linhas.length && i > melhor.i)) melhor = { n: n, t: t, i: i };
        });
        $('#ouv-arquivo').value = '';
        mostrarPrevia(melhor ? paraFormatoOuv(melhor.t) : null);
        if (melhor) mostrarInterpretacao(melhor);
      });
    }).catch(function (e) { $('#ouv-previa').innerHTML = '<p class="erro">' + esc(e.message) + '</p>'; })
      .then(function () { btn.disabled = false; });
  });
  $('#ouv-arquivo').addEventListener('change', function () {
    var f = this.files[0]; if (!f) return;
    f.arrayBuffer().then(function (buf) {
      $('#ouv-link').value = '';
      var wb = /\.csv$/i.test(f.name) ? window.Planilhas.lerArquivo(f.name, buf) : window.XLSX.read(buf, { type: 'array', cellDates: true, cellStyles: true });
      // a aba com mais linhas classificadas
      var melhor = null;
      wb.SheetNames.forEach(function (nm) {
        var a = window.Auditada.daAbaXlsx(wb.Sheets[nm], window.XLSX), t = window.Auditada.interpretar(a.m, a.cores);
        if (t && (!melhor || t.linhas.length > melhor.t.linhas.length)) melhor = { n: nm, t: t };
      });
      mostrarPrevia(melhor ? paraFormatoOuv(melhor.t) : null);
      if (melhor) mostrarInterpretacao(melhor);
    }).catch(function (e) { $('#ouv-previa').innerHTML = '<p class="erro">Não consegui ler o arquivo: ' + esc(e.message) + '</p>'; });
  });

  /* ---------- Reclamação 6: uma planilha por mês (cada uma com as suas listas por data) ---------- */
  // a planilha que existia (uma só) vira a de Outubro, com o período da rodada de setembro/outubro
  function migrarMeses(d) {
    d.listas = d.listas || [];
    if (!d.mesesR6 || !d.mesesR6.length) {
      d.mesesR6 = [{ id: 'r6-2026-10', nome: 'Ouvidorias Faturas Outubro – Reclamação 6', link: d.planilha || PLANILHA_OUV_PADRAO, desde: '2026-11-14', ate: '2026-11-24', atual: true }];
    }
    if (!d.mesesR6.some(function (m) { return m.atual; })) d.mesesR6[d.mesesR6.length - 1].atual = true;
    return d;
  }
  function mesesR6() { return (ouv.dados && ouv.dados.mesesR6) || []; }
  function idPlan(link) { return (String(link || '').match(/\/d\/([\w-]+)/) || [])[1] || ''; }
  function mesAtual() { var ms = mesesR6(); return ms.filter(function (m) { return m.atual; })[0] || ms[ms.length - 1] || null; }
  // planilha do mês de uma lista: a escolhida ao criar; nas listas antigas, pela planilha onde a aba foi criada
  function mesDaLista(l) {
    var ms = mesesR6();
    return ms.filter(function (m) { return m.id === l.mes; })[0] ||
      ms.filter(function (m) { return l.planilhaId && idPlan(m.link) === l.planilhaId; })[0] || ms[0] || null;
  }
  function linkPlanilhaOuv(l) { var m = l ? mesDaLista(l) : mesAtual(); return (m && m.link) || (ouv.dados && ouv.dados.planilha) || PLANILHA_OUV_PADRAO; }
  // a faixa do Google acompanha a conexão feita em qualquer página do portal
  var gAntes = null;
  setInterval(function () {
    var on = !!(window.Google && window.Google.conectado());
    if (gAntes !== null && on !== gAntes && $('#ouv-gbar') && $('#ouv-gbar').innerHTML) desenharPlanilhaOuv();
    // acabou de conectar o Google: atualiza as abas das listas e as outras planilhas
    if (gAntes === false && on && ouv.dados) ouvCarregarRegs().then(function () { return sincronizarPlanilhaOuv(false); }).then(function () { return sincronizarAcomp(false); });
    gAntes = on;
  }, 3000);
  var sincBotao = { rodando: false, msg: '' };
  function desenharPlanilhaOuv() {
    var ma = mesAtual();
    $('#ouv-planilha').innerHTML = (ma ? '📗 As listas novas vão para: <b>' + esc(ma.nome) + '</b> (as planilhas de cada mês ficam em "Planilhas de ouvidorias", abaixo).' : '') +
      (window.Google.conectado() ? ' <button class="btn btn-peq" type="button" id="ouv-sinc"' + (sincBotao.rodando ? ' disabled>Atualizando…' : '>Atualizar a planilha agora') + '</button>' +
        ' <span id="ouv-sinc-st" style="margin-left:8px;font-size:.9rem">' + esc(sincBotao.msg) + '</span>' : '');
    // conexão com o Google, em destaque no topo da aba
    var on = window.Google.conectado(), gb = $('#ouv-gbar');
    gb.className = 'card ouv-gbar ' + (on ? 'ok' : 'nao');
    gb.innerHTML = on
      ? '<span class="gtit">✓ Google conectado</span><span class="dica" style="margin:0">As planilhas das ouvidorias são atualizadas com a sua conta do Google.</span>'
      : '<span class="gtit">Google não conectado</span><button class="btn btn-primario" type="button" id="ouv-google">Conectar Google</button>' +
        '<span class="dica" style="margin:0">Conecte para criar as abas e para as planilhas das ouvidorias serem atualizadas (vale para toda a plataforma enquanto ela estiver aberta).</span>';
    if ($('#ouv-sinc')) $('#ouv-sinc').onclick = function () {
      // o andamento aparece ao lado do botão (o aviso flutuante fica fora da tela dentro do portal)
      sincBotao.rodando = true; sincBotao.msg = ''; desenharPlanilhaOuv();
      var espera = function () { return sincronizando ? new Promise(function (ok) { setTimeout(ok, 800); }).then(espera) : Promise.resolve(); };
      espera().then(ouvCarregarRegs).then(function () { ultimoAvisoSinc = ''; return sincronizarPlanilhaOuv(true); })
        .then(function () { sincBotao.msg = ultimoAvisoSinc || 'Pronto.'; })
        .catch(function (e) { sincBotao.msg = 'Não consegui atualizar: ' + (e && e.message || e); })
        .then(function () { sincBotao.rodando = false; desenharPlanilhaOuv(); });
    };
    if ($('#ouv-google')) $('#ouv-google').onclick = function () {
      // dentro do portal: conecta pela página do Google e volta (a janela do Google falha em alguns Chrome)
      if (window.PLAT && window.PLAT.google && window.PLAT.google.entrarPelaPagina) {
        try { sessionStorage.setItem('ouv_aba', MODO_PLANILHAS ? 'planilhas' : ''); } catch (e) { /* nada */ }
        window.PLAT.google.entrarPelaPagina(); return;
      }
      var gb0 = $('#ouv-gbar'); if (gb0) gb0.insertAdjacentHTML('beforeend', '<p class="dica" style="flex-basis:100%;margin:4px 0 0">Termine na janela do Google (escolher a conta, Acessar e Continuar). Esta faixa fica verde sozinha quando terminar.</p>');
      window.Google.entrar(estado.clientIdGoogle).then(function () { atualizarConfigGoogle(); avisar('Google conectado.'); renderOuv(); sincronizarPlanilhaOuv(true).then(function () { return sincronizarAcomp(true); }); })
        .catch(function (e) {
          var msg = e.message || 'O login com Google não foi concluído.'; avisar(msg);
          var gb = $('#ouv-gbar'); if (gb) gb.insertAdjacentHTML('beforeend', '<p class="erro" style="flex-basis:100%;margin:4px 0 0">' + esc(msg) + '</p>');
        });
    };
    // escolha da planilha do mês no formulário da nova lista
    var sel = $('#ouv-mes');
    if (sel) {
      var val = sel.value;
      sel.innerHTML = mesesR6().map(function (m) { return '<option value="' + esc(m.id) + '">' + esc(m.nome) + '</option>'; }).join('');
      sel.value = val && mesesR6().some(function (m) { return m.id === val; }) ? val : (ma ? ma.id : '');
    }
  }
  function salvarOuv(mudar, msg) {
    return estado.repo.atualizarJSON(ARQ_OUV, { planilha: PLANILHA_OUV_PADRAO, listas: [] }, function (d) {
      migrarMeses(d); return mudar(d);
    }, msg + ' (' + estado.nome + ')').then(function (d) { ouv.dados = d; return d; });
  }
  // recarrega as ouvidorias registradas (podem ter entrado novas desde que a página abriu)
  function ouvCarregarRegs() {
    return new A.Repositorio(estado.repo.token, 'danielamotaxp-code/ouvidorias-cemig-dados').lerJSON('registros.json', { registros: [] })
      .then(function (r) { if (r && r.registros && r.registros.length) ouv.regs = r.registros; }).catch(function () {});
  }
  function carregarOuv() {
    $('#ouv-data').value = $('#ouv-data').value || hojeISO();
    return Promise.all([
      estado.repo.lerJSON(ARQ_OUV, { planilha: PLANILHA_OUV_PADRAO, listas: [] }),
      new A.Repositorio(estado.repo.token, 'danielamotaxp-code/ouvidorias-cemig-dados').lerJSON('registros.json', { registros: [] }).catch(function () { return { registros: null }; })
    ]).then(function (r) { ouv.dados = migrarMeses(r[0] || {}); ouv.regs = r[1].registros || null; if (!ouv.regs) avisar("Não consegui carregar a lista de ouvidorias da plataforma. Recarregue a página."); renderOuv(); desenharAcomp(); sincronizarPlanilhaOuv(false).then(function () { return sincronizarAcomp(false); }); })
      .catch(function (e) { $('#ouv-listas').innerHTML = '<p class="erro">' + esc(e.message) + '</p>'; });
  }

  /** Cria a aba com a data na planilha de ouvidorias (se já existir uma aba com esse nome, usa "(2)", "(3)"…). */
  function criarAbaGoogle(lista) {
    if (!window.Google.conectado()) return Promise.reject(new Error('Clique em "Entrar com Google" no topo para criar a aba na planilha.'));
    return window.Google.lerPlanilha(linkPlanilhaOuv(lista), estado.chaveGoogle).then(function (g) {
      var nome = lista.data, n = 2;
      while (g.folhas[nome]) nome = lista.data + ' (' + (n++) + ')';
      return window.Google.escreverAba(g, nome, [lista.cab].concat(lista.linhas), 10).then(function (sheetId) {
        // contorno em toda a tabela (inclusive Colaborador, Ouvidoria, Protocolo, Data e Abrir ouvidoria); as 10 linhas e colunas a mais ficam em branco
        var B = { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } };
        return window.Google.lote(g.id, [{ updateBorders: { range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: lista.linhas.length + 1, startColumnIndex: 0, endColumnIndex: lista.cab.length },
          top: B, bottom: B, left: B, right: B, innerHorizontal: B, innerVertical: B } }]).then(function () { return sheetId; });
      }).then(function (sheetId) {
        return salvarOuv(function (d) {
          d.listas.forEach(function (x) { if (x.id === lista.id) { x.aba = nome; x.gid = sheetId; x.planilhaId = g.id; } });
          return d;
        }, 'Ouvidorias: aba ' + nome + ' criada na planilha');
      });
    });
  }

  $('#ouv-salvar').addEventListener('click', function () {
    var t = ouv.previa, btn = this, msg = $('#ouv-msg');
    if (!t) return;
    var iso = $('#ouv-data').value || hojeISO();
    var mesEsc = ($('#ouv-mes') && $('#ouv-mes').value) || (mesAtual() && mesAtual().id) || '';
    var lista = { id: carimbo(new Date()), dataISO: iso, data: isoParaBR(iso), mes: mesEsc,
      cab: t.cab, linhas: t.linhas, usuario: estado.nome, criadoEm: new Date().toISOString() };
    btn.disabled = true; msg.textContent = 'Salvando a lista…';
    salvarOuv(function (d) { d.listas.push(lista); return d; }, 'Ouvidorias: lista de ' + lista.data)
      .then(function () {
        if (!window.Google.conectado()) { msg.textContent = 'Lista salva. Para criar a aba na planilha, clique em "Entrar com Google" e depois em "Criar aba na planilha" na lista.'; return; }
        msg.textContent = 'Criando a aba "' + lista.data + '" na planilha de ouvidorias…';
        return criarAbaGoogle(lista).then(function () {
          // já preenche e colore as linhas de quem já tem ouvidoria registrada na plataforma
          msg.textContent = 'Aba criada. Marcando as ouvidorias já registradas…';
          return ouvCarregarRegs().then(function () { return sincronizarPlanilhaOuv(false); })
            .then(function () { msg.textContent = 'Lista salva, aba criada e ouvidorias já registradas marcadas ✓'; });
        })
          .catch(function (e) { msg.textContent = 'Lista salva, mas não consegui criar a aba: ' + e.message; });
      })
      .then(function () { $('#ouv-link').value = ''; $('#ouv-arquivo').value = ''; mostrarPrevia(); renderOuv(); })
      .catch(function (e) { msg.textContent = 'Não consegui salvar: ' + e.message; btn.disabled = false; });
  });

  // ouvidorias feitas até 3 dias antes da data da lista também contam (as ouvidorias podem começar antes da aba ser criada)
  function seteDiasAntes(iso) { var d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() - 3); return d.toISOString().slice(0, 10); }
  // data da lista mais antiga em que a instalação aparece: a ouvidoria aberta por causa dela vale para todas as listas seguintes
  function primeiraListaDe(k) {
    var menor = '';
    ((ouv.dados && ouv.dados.listas) || []).forEach(function (l) {
      var ci = (l.cab || []).findIndex(function (h) { return R.normalizar ? R.normalizar(h) === 'instalacao' : /instala/i.test(h); });
      if (ci < 0 || !l.dataISO) return;
      if ((l.linhas || []).some(function (x) { return R.chaveInstalacao(x[ci]) === k; }) && (!menor || l.dataISO < menor)) menor = l.dataISO;
    });
    return menor;
  }
  /* ---------- período das ouvidorias (de / até), escolhido em cada lista e em cada planilha ---------- */
  function brParaISO(t) { var d = String(t || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return d ? d[3] + '-' + ('0' + d[2]).slice(-2) + '-' + ('0' + d[1]).slice(-2) : ''; }
  function textoPeriodo(desde, ate) {
    if (!desde && !ate) return 'Todas as ouvidorias';
    if (!ate) return 'De ' + isoParaBR(desde) + ' em diante';
    if (!desde) return 'Até ' + isoParaBR(ate);
    return 'De ' + isoParaBR(desde) + ' até ' + isoParaBR(ate);
  }
  // pergunta "de" e "até"; devolve null se a pessoa cancelar
  function pedirPeriodo(onde, desde, ate) {
    var a = window.prompt('Período das ouvidorias ' + onde + '\n\nConsiderar as ouvidorias registradas a partir de (dd/mm/aaaa):', isoParaBR(desde || ''));
    if (a === null) return null;
    var b = window.prompt('Até (dd/mm/aaaa)?\n\nDeixe em branco enquanto a rodada estiver aberta (vale até hoje e os próximos dias).', isoParaBR(ate || ''));
    if (b === null) return null;
    var de = brParaISO(a), at = brParaISO(b);
    if (de && at && at < de) { avisar('A data "até" ficou antes da data "de". Escolha o período de novo.'); return null; }
    return { desde: de, ate: at };
  }
  // abre, ao lado do botão, os campos "de" e "até" (calendário) e o botão para aplicar e puxar as ouvidorias
  function caixaPeriodo(botao, desde, ate, aplicar) {
    // dentro do título de uma lista (summary), a caixa vai logo abaixo do título, com a lista aberta
    var sum = botao.closest('summary'), lugar = sum ? sum.parentNode : botao.parentNode;
    var aberta = lugar.querySelector('.ouv-periodo');
    if (aberta) { aberta.remove(); return; }
    var cx = document.createElement('span'); cx.className = 'ouv-periodo';
    cx.style.cssText = 'display:inline-flex;flex-wrap:wrap;gap:6px;align-items:center;margin:4px 0 4px 6px;padding:6px 8px;border:1px solid #c9d3df;border-radius:8px;background:#f6f9fc';
    cx.innerHTML = '<label>De <input type="date" class="campo" style="width:auto" data-de value="' + esc(desde || '') + '"></label>' +
      '<label>até <input type="date" class="campo" style="width:auto" data-ate value="' + esc(ate || '') + '"></label>' +
      '<button class="btn btn-peq" type="button" data-aplicar>Aplicar e puxar as ouvidorias</button>' +
      '<button class="btn btn-link btn-peq" type="button" data-fechar>Cancelar</button>' +
      '<small style="flex-basis:100%">"Até" em branco = rodada aberta (as ouvidorias novas continuam entrando).</small>';
    ['click', 'keydown'].forEach(function (ev) { cx.addEventListener(ev, function (e) { e.stopPropagation(); if (ev === 'click' && e.target.tagName !== 'INPUT') e.preventDefault(); }); });
    if (sum) { sum.insertAdjacentElement('afterend', cx); sum.parentNode.open = true; } else botao.insertAdjacentElement('afterend', cx);
    cx.querySelector('[data-fechar]').onclick = function () { cx.remove(); };
    cx.querySelector('[data-aplicar]').onclick = function () {
      var de = cx.querySelector('[data-de]').value, at = cx.querySelector('[data-ate]').value;
      if (de && at && at < de) { avisar('A data "até" ficou antes da data "de".'); return; }
      this.disabled = true; this.textContent = 'Puxando…';
      aplicar({ desde: de, ate: at });
    };
  }
  function noPeriodo(r, desde, ate) { var d = String(r.criadoEm || '').slice(0, 10); return (!desde || d >= desde) && (!ate || d <= ate); }
  // l = a lista (com o período escolhido). Listas antigas, sem período: regra de antes (3 dias antes da primeira lista do cliente).
  function regOuv(inst, l) {
    var k = R.chaveInstalacao(inst); if (!k || !ouv.regs) return null;
    var desde, ate = '';
    var m = mesDaLista(l);
    if (l.desde || l.ate) { desde = l.desde || ''; ate = l.ate || ''; }
    else if (m && (m.desde || m.ate)) { desde = m.desde || ''; ate = m.ate || ''; }
    else { desde = l.dataISO; var p = primeiraListaDe(k); if (p && p < desde) desde = p; desde = seteDiasAntes(desde); }
    var rs = ouv.regs.filter(function (r) { return +r.reclamacao === 6 && R.chaveInstalacao(r.instalacao) === k && noPeriodo(r, desde, ate); });
    if (!rs.length) return null;
    rs.sort(function (a, b) { return String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')); }); // a mais recente
    return rs.filter(function (r) { return r.status !== 'pendente'; })[0] || rs[0];
  }
  function letraCol(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  /**
   * Atualiza a aba de cada lista na planilha de ouvidorias, conforme as ouvidorias registradas em Ouvidorias Cemig:
   * Colaborador (quem registrou, se estiver vazio), situação, protocolo, data, link "Abrir ouvidoria" e a linha em verde.
   */
  var sincronizando = false;
  var VERDE = { red: 0.204, green: 0.659, blue: 0.325 }, LARANJA = { red: 0.992, green: 0.851, blue: 0.663 };
  var BORDA = { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } };
  function sincronizarPlanilhaOuv(manual) {
    var listas = ((ouv.dados && ouv.dados.listas) || []).filter(function (l) { return l.aba; });
    if (manual && !window.Google.conectado()) { avisar('Conecte o Google no topo de Ouvidorias Cemig e tente de novo.'); return Promise.resolve(false); }
    if (manual && !listas.length) { avisar('Nenhuma lista com aba na planilha ainda.'); return Promise.resolve(false); }
    if (sincronizando || !listas.length || !window.Google.conectado()) return Promise.resolve(false);
    sincronizando = true;
    // agrupa as listas pela planilha do mês e atualiza uma planilha de cada vez
    var grupos = {};
    listas.forEach(function (l) { var k = linkPlanilhaOuv(l); (grupos[k] = grupos[k] || []).push(l); });
    var mudouAlgo = false;
    return Object.keys(grupos).reduce(function (cad, link) {
      return cad.then(function () { return sincUmaPlanilhaR6(link, grupos[link]).then(function (x) { if (x) mudouAlgo = true; }); });
    }, Promise.resolve()).then(function () {
      if (manual) avisar(mudouAlgo ? 'Planilhas de ouvidorias atualizadas.' : 'As planilhas de ouvidorias já estavam em dia.');
      return true;
    }).catch(function (e) { if (manual) avisar('Não consegui atualizar a planilha de ouvidorias: ' + e.message); return false; })
      .then(function (x) { sincronizando = false; return x; });
  }
  function sincUmaPlanilhaR6(link, listas) {
    return window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      var dados = [], formatos = [];
      listas.forEach(function (l) {
        var m = g.abas[l.aba], folha = g.folhas[l.aba]; if (!m || !m.length || !folha) return;
        var hl = m.findIndex(function (row) { return (row || []).some(function (h) { return nrm(h) === nrm('Instalação'); }); });
        if (hl < 0) return;
        var cab = m[hl].map(function (h) { return nrm(h); }), ci = cab.indexOf(nrm('Instalação')), cc = cab.indexOf('colaborador');
        var cs = COLS_AUTO.map(function (h) { return cab.indexOf(nrm(h)); });
        var ultimaCol = cab.length, primeiraCol = Math.max(0, cab.findIndex(function (h) { return h; })); // pinta só a tabela, não colunas vazias à esquerda
        m.slice(hl + 1).forEach(function (linha, j) {
          var r = regOuv(linha[ci], l); if (!r) return;
          var lin = hl + j + 2, quem = r.responsavel || '', pend = r.status === 'pendente';
          var val = [pend ? 'Não enviada (problema no site)' : 'Enviada', pend ? '' : (r.protocolo || ''),
            isoParaBR(String((pend ? r.ultimaTentativa : r.dataAbertura) || r.criadoEm || '').slice(0, 10))];
          var mudou = false;
          if (cc >= 0 && !String(linha[cc] || '').trim() && quem) { dados.push({ range: window.Google.aspas(l.aba) + '!' + letraCol(cc) + lin, values: [[quem]] }); mudou = true; }
          val.forEach(function (v, k) { var c = cs[k]; if (c >= 0 && String(linha[c] == null ? '' : linha[c]) !== v) { dados.push({ range: window.Google.aspas(l.aba) + '!' + letraCol(c) + lin, values: [[v]] }); mudou = true; } });
          var cl = cs[3];
          if (cl >= 0 && String(linha[cl] || '') !== 'Abrir ouvidoria') {
            formatos.push({ updateCells: { range: { sheetId: folha.sheetId, startRowIndex: lin - 1, endRowIndex: lin, startColumnIndex: cl, endColumnIndex: cl + 1 },
              rows: [{ values: [{ userEnteredValue: { stringValue: 'Abrir ouvidoria' }, textFormatRuns: [{ startIndex: 0, format: { link: { uri: PORTAL_OUV + r.id } } }] }] }],
              fields: 'userEnteredValue,textFormatRuns' } });
            mudou = true;
          }
          if (mudou) formatos.push({ repeatCell: { range: { sheetId: folha.sheetId, startRowIndex: lin - 1, endRowIndex: lin, startColumnIndex: primeiraCol, endColumnIndex: ultimaCol },
            cell: { userEnteredFormat: { backgroundColor: pend ? LARANJA : VERDE } }, fields: 'userEnteredFormat.backgroundColor' } });
          // contorno nas células da ouvidoria (Colaborador, Ouvidoria, Protocolo, Data, Abrir ouvidoria)
          if (mudou) [cc].concat(cs).filter(function (c) { return c >= 0; }).forEach(function (c) {
            formatos.push({ updateBorders: { range: { sheetId: folha.sheetId, startRowIndex: lin - 1, endRowIndex: lin, startColumnIndex: c, endColumnIndex: c + 1 },
              top: BORDA, bottom: BORDA, left: BORDA, right: BORDA } });
          });
          // colunas da ouvidoria centralizadas e sem negrito
          if (mudou) [cc].concat(cs).filter(function (c) { return c >= 0; }).forEach(function (c) {
            formatos.push({ repeatCell: { range: { sheetId: folha.sheetId, startRowIndex: lin - 1, endRowIndex: lin, startColumnIndex: c, endColumnIndex: c + 1 },
              cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: false } } }, fields: 'userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat.bold' } });
          });
        });
      });
      return window.Google.gravarIntervalos(g.id, dados)
        .then(function () { return formatos.length ? window.Google.lote(g.id, formatos) : null; })
        .then(function () { return !!(dados.length || formatos.length); });
    });
  }

  function situacaoOuv(inst, l) {
    var r0 = regOuv(inst, l); if (!r0) return null;
    var rs = [r0];
    var env = rs.filter(function (r) { return r.status !== 'pendente'; })[0];
    return env ? { cls: 'ok', txt: 'Enviada' + (env.protocolo ? ' · ' + env.protocolo : '') } : { cls: 'pend', txt: 'Não enviada (problema no site)' };
  }

  function renderOuv() {
    desenharPlanilhaOuv();
    var box = $('#ouv-listas'), listas = ((ouv.dados && ouv.dados.listas) || []).slice().sort(function (a, b) { return String(b.dataISO + b.id).localeCompare(String(a.dataISO + a.id)); });
    if (!listas.length) { box.innerHTML = '<div class="card"><p class="vazio">Nenhuma lista ainda. Cole a primeira lista acima.</p></div>'; return; }
    box.innerHTML = listas.map(function (l, li) {
      var c = colunasOuv(l.cab), q = nrm(ouv.filtro[l.id] || '');
      var feitos = l.linhas.filter(function (x) { var s = situacaoOuv(x[c.inst], l); return s && s.cls === 'ok'; }).length;
      var linhas = l.linhas.map(function (x, i) { return { x: x, i: i }; }).filter(function (o) { return !q || nrm(o.x.join(' ')).indexOf(q) >= 0; });
      var abaLink = l.gid != null ? 'https://docs.google.com/spreadsheets/d/' + esc(l.planilhaId || idPlan(linkPlanilhaOuv(l))) + '/edit#gid=' + l.gid : '';
      return '<details class="card ouv-lista"' + (li === 0 ? ' open' : '') + '><summary>' +
        '<h3>Lista de ' + esc(l.data) + '</h3><span class="ouv-chip"><b>' + num(l.linhas.length) + '</b> clientes</span>' +
        '<span class="ouv-chip" style="background:#e7f4ec;color:#1d5a37">' + num(feitos) + ' com ouvidoria enviada</span>' +
        (function () { var m = mesDaLista(l), de = l.desde || (m && m.desde), at = l.ate || (m && m.ate);
          return '<span class="ouv-chip">📗 ' + esc(m ? m.nome : 'Planilha de ouvidorias') + '</span>' +
            '<span class="ouv-chip" title="Só as ouvidorias registradas nesse período marcam os clientes desta lista">📅 ' + esc((de || at) ? textoPeriodo(de, at) : 'Período: regra antiga (3 dias antes da lista)') + '</span>'; })() +
        '<span class="dica">por ' + esc(l.usuario || '') + '</span>' +
        '<span class="acoes">' + (l.aba ? '<a class="btn btn-peq" href="' + abaLink + '" target="_blank" rel="noopener">Abrir aba "' + esc(l.aba) + '"</a>'
          : '<button class="btn btn-peq" type="button" data-criar="' + esc(l.id) + '">Criar aba na planilha</button>') +
        '<button class="btn btn-link btn-peq" type="button" data-excluir="' + esc(l.id) + '">Excluir lista</button></span></summary>' +
        '<div class="ouv-resumo">' + chips(contar(l.linhas, c.classif)) + '</div>' +
        '<div class="ouv-filtro"><input class="campo" type="search" placeholder="Buscar cliente, instalação, colaborador…" data-busca="' + esc(l.id) + '" value="' + esc(ouv.filtro[l.id] || '') + '"></div>' +
        '<div class="tabela-wrap"><table class="ouv-tab"><thead><tr>' + (c.colab >= 0 ? '<th>Colaborador</th>' : '') + '<th>Cliente</th><th>Instalação</th><th>UC</th>' + (c.classif >= 0 ? '<th>Classificação</th>' : '') + '<th>Ouvidoria</th><th></th></tr></thead><tbody>' +
        linhas.map(function (o) {
          var x = o.x, s = situacaoOuv(x[c.inst], l);
          var rg = regOuv(x[c.inst], l);
          return '<tr>' + (c.colab >= 0 ? '<td>' + esc(x[c.colab] || (rg && rg.responsavel) || '') + '</td>' : '') + '<td class="cl">' + esc(x[c.cliente]) + '</td><td>' + esc(x[c.inst]) + '</td><td>' + esc(c.uc >= 0 ? x[c.uc] : '') + '</td>' +
            (c.classif >= 0 ? '<td class="cl">' + esc(x[c.classif]) + '</td>' : '') +
            '<td>' + (s ? '<span class="ouv-sit ' + s.cls + '">' + esc(s.txt) + '</span>' : '<span class="ouv-sit nao">Não gerada</span>') + '</td>' +
            '<td>' + (rg ? '<button class="btn btn-peq" type="button" data-abrir="' + esc(rg.id) + '">Abrir ouvidoria</button>' : '') + '</td></tr>';
        }).join('') + '</tbody></table></div></details>';
    }).join('');
    $$('[data-busca]', box).forEach(function (inp) {
      inp.addEventListener('input', function () { ouv.filtro[inp.dataset.busca] = inp.value; var pos = inp.selectionStart; renderOuv(); var n = $('[data-busca="' + inp.dataset.busca + '"]'); if (n) { n.focus(); n.setSelectionRange(pos, pos); } });
      inp.addEventListener('click', function (e) { e.stopPropagation(); });
    });
    $$('[data-criar]', box).forEach(function (b) {
      b.onclick = function (e) {
        e.preventDefault(); var l = ouv.dados.listas.filter(function (x) { return x.id === b.dataset.criar; })[0];
        b.disabled = true; b.textContent = 'Criando…';
        criarAbaGoogle(l).then(function () { return ouvCarregarRegs().then(function () { return sincronizarPlanilhaOuv(false); }); }).then(function () { avisar('Aba criada na planilha de ouvidorias, com as ouvidorias já registradas marcadas.'); renderOuv(); })
          .catch(function (err) { avisar(err.message); b.disabled = false; b.textContent = 'Criar aba na planilha'; });
      };
    });
    $$('[data-excluir]', box).forEach(function (b) {
      b.onclick = function (e) {
        e.preventDefault();
        if (b.dataset.conf !== '1') { b.dataset.conf = '1'; b.textContent = 'Excluir mesmo? (a aba na planilha continua)'; return; }
        salvarOuv(function (d) { d.listas = d.listas.filter(function (x) { return x.id !== b.dataset.excluir; }); return d; }, 'Ouvidorias: lista excluída').then(renderOuv);
      };
    });
    $$('[data-abrir]', box).forEach(function (b) {
      b.onclick = function () {
        if (window.PLAT) { try { window.top.location.hash = '#ouvidorias/cad/' + b.dataset.abrir; return; } catch (err) { /* abre em outra aba */ } }
        window.open(PORTAL_OUV + b.dataset.abrir, '_blank');
      };
    });
  }
  /* ---------- outras planilhas de ouvidorias (só atualiza os clientes que já estão nelas) ---------- */
  function gPortal() {
    if (window.PLAT && window.PLAT.google && window.PLAT.google.conectado && window.PLAT.google.conectado()) return window.PLAT.google;
    if (window.Google && window.Google.conectado && window.Google.conectado()) return window.Google; // Google conectado nesta página
    return null;
  }
  function desenharAcomp() {
    var box = $('#ouv-acomp'), ps = (ouv.dados && ouv.dados.acompanhadas) || [], ms = mesesR6();
    var linhasMes = ms.map(function (m, i) {
      var nl = ((ouv.dados && ouv.dados.listas) || []).filter(function (l) { var x = mesDaLista(l); return x && x.id === m.id; }).length;
      return '<tr><td class="cl"><a href="' + esc(m.link) + '" target="_blank" rel="noopener">' + esc(m.nome) + '</a>' + (m.atual ? ' <span class="ouv-chip" style="background:#e7f4ec;color:#1d5a37">recebe as listas novas</span>' : '') + '</td>' +
        '<td>6 (listas por data)</td><td>Sim (verde / laranja)</td><td>' + esc(textoPeriodo(m.desde, m.ate)) + '</td><td>' + num(nl) + ' lista(s)</td>' +
        '<td><button class="btn btn-peq" type="button" data-mes-per="' + i + '">📅 Período</button> ' +
        (m.atual ? '' : '<button class="btn btn-link btn-peq" type="button" data-mes-atual="' + i + '">Receber as listas novas</button> ') +
        '<button class="btn btn-link btn-peq" type="button" data-mes-ed="' + i + '">Editar</button></td></tr>';
    }).join('');
    box.innerHTML = (ps.length || ms.length) ? '<div class="tabela-wrap"><table class="ouv-tab"><thead><tr><th>Planilha</th><th>Reclamações</th><th>Pinta a linha</th><th>Período das ouvidorias</th><th>Última atualização</th><th></th></tr></thead><tbody>' + linhasMes +
      ps.map(function (p, i) {
        var u = ouv.ultimaAcomp && ouv.ultimaAcomp[p.nome];
        return '<tr><td class="cl"><a href="' + esc(p.link) + '" target="_blank" rel="noopener">' + esc(p.nome) + '</a></td>' +
          '<td>' + ((p.reclamacoes || []).length ? p.reclamacoes.join(', ') : 'Pela coluna da planilha') + '</td>' +
          '<td>' + (p.colorir ? 'Sim (verde / laranja)' : 'Não (mantém as cores)') + '</td><td>' + esc(textoPeriodo(p.desde, p.ate)) + '</td>' +
          '<td>' + (u ? (u.erro ? '<span class="erro">' + esc(u.erro) + '</span>' : num(u.linhas) + ' linha(s) atualizada(s)') : '—') + '</td>' +
          '<td><button class="btn btn-peq" type="button" data-acomp-per="' + i + '">📅 Período</button> <button class="btn btn-link btn-peq" type="button" data-acomp-ed="' + i + '">Editar</button> <button class="btn btn-link btn-peq" type="button" data-acomp-del="' + i + '">Tirar</button></td></tr>';
      }).join('') + '</tbody></table></div>' : '<p class="vazio">Nenhuma planilha ainda.</p>';
    $$('[data-acomp-ed]', box).forEach(function (b) { b.onclick = function () { editarAcomp(+b.dataset.acompEd); }; });
    // planilhas do mês (Reclamação 6)
    $$('[data-mes-per]', box).forEach(function (b) {
      b.onclick = function () {
        var m = ms[+b.dataset.mesPer]; if (!m) return;
        caixaPeriodo(b, m.desde || '', m.ate || '', function (per) {
          salvarOuv(function (d) { d.mesesR6.forEach(function (x) { if (x.id === m.id) { x.desde = per.desde; x.ate = per.ate; } }); return d; }, 'Ouvidorias: período da planilha ' + m.nome)
            .then(function () { return ouvCarregarRegs().then(function () { return sincronizarPlanilhaOuv(false); }); })
            .then(function () { renderOuv(); desenharAcomp(); avisar(m.nome + ': ' + textoPeriodo(per.desde, per.ate) + '. Ouvidorias puxadas para as abas.'); });
        });
      };
    });
    $$('[data-mes-atual]', box).forEach(function (b) {
      b.onclick = function () {
        var m = ms[+b.dataset.mesAtual]; if (!m) return;
        salvarOuv(function (d) { d.mesesR6.forEach(function (x) { x.atual = x.id === m.id; }); return d; }, 'Ouvidorias: listas novas vão para ' + m.nome)
          .then(function () { renderOuv(); desenharAcomp(); avisar('As listas novas vão para ' + m.nome + '.'); });
      };
    });
    $$('[data-mes-ed]', box).forEach(function (b) {
      b.onclick = function () {
        var m = ms[+b.dataset.mesEd]; if (!m) return;
        var nome = window.prompt('Nome da planilha:', m.nome); if (!nome) return;
        var link = window.prompt('Link da planilha (Google Planilhas):', m.link); if (!link) return;
        if (!/\/spreadsheets\/d\//.test(link)) { avisar('Esse não parece um link do Google Planilhas.'); return; }
        salvarOuv(function (d) { d.mesesR6.forEach(function (x) { if (x.id === m.id) { x.nome = nome.trim(); x.link = link.trim(); } }); return d; }, 'Ouvidorias: planilha ' + nome.trim())
          .then(function () { renderOuv(); desenharAcomp(); });
      };
    });
    $$('[data-acomp-per]', box).forEach(function (b) {
      b.onclick = function () {
        var p = ps[+b.dataset.acompPer]; if (!p) return;
        caixaPeriodo(b, p.desde || '', p.ate || '', function (per) {
          salvarOuv(function (d) { (d.acompanhadas || []).forEach(function (x) { if (x.nome === p.nome) { x.desde = per.desde; x.ate = per.ate; } }); return d; }, 'Ouvidorias: período da planilha ' + p.nome)
            .then(function () { desenharAcomp(); avisar(p.nome + ': ' + textoPeriodo(per.desde, per.ate) + '.'); return ouvCarregarRegs().then(function () { return sincronizarAcomp(true); }); });
        });
      };
    });
    $$('[data-acomp-del]', box).forEach(function (b) {
      b.onclick = function () {
        if (b.dataset.conf !== '1') { b.dataset.conf = '1'; b.textContent = 'Tirar mesmo?'; return; }
        var nome = ps[+b.dataset.acompDel].nome;
        salvarOuv(function (d) { d.acompanhadas = (d.acompanhadas || []).filter(function (x) { return x.nome !== nome; }); return d; }, 'Ouvidorias: tira a planilha ' + nome).then(desenharAcomp);
      };
    });
  }
  function editarAcomp(i) {
    var ps = (ouv.dados && ouv.dados.acompanhadas) || [], p = i != null ? ps[i] : {};
    var nome = window.prompt('Nome da planilha:', p.nome || ''); if (!nome) return;
    var link = window.prompt('Link da planilha (Google Planilhas):', p.link || ''); if (!link) return;
    if (!/\/spreadsheets\/d\//.test(link)) { avisar('Esse não parece um link do Google Planilhas. Se for um arquivo do Excel (.xlsx), abra no Drive e use Arquivo → Salvar como Planilhas Google.'); return; }
    var colorir = window.confirm('Pintar a linha de verde quando a ouvidoria for enviada (e laranja quando der erro)?\n\nOK = pintar · Cancelar = manter as cores da planilha');
    var rq = window.prompt('Quais reclamações valem para essa planilha? (ex.: 7, 8)\nDeixe vazio para usar a coluna "Reclamação N" de cada linha, ou qualquer reclamação se não houver essa coluna.', (p.reclamacoes || []).join(', '));
    if (rq === null) return;
    var reclamacoes = (rq.match(/\d+/g) || []).map(Number);
    var per = pedirPeriodo('da planilha ' + nome.trim(), p.desde || hojeISO(), p.ate || ''); if (!per) return;
    var novo = { nome: nome.trim(), link: link.trim(), colorir: colorir, reclamacoes: reclamacoes, desde: per.desde, ate: per.ate };
    salvarOuv(function (dd) {
      dd.acompanhadas = dd.acompanhadas || [];
      if (i != null) dd.acompanhadas = dd.acompanhadas.map(function (x) { return x.nome === p.nome ? novo : x; }); else dd.acompanhadas.push(novo);
      return dd;
    }, 'Ouvidorias: planilha ' + novo.nome).then(function () { desenharAcomp(); sincronizarAcomp(true); });
  }
  var sincAcomp = false;
  function sincronizarAcomp(manual) {
    var ps = (ouv.dados && ouv.dados.acompanhadas) || [], g = gPortal(), msg = $('#ouv-acomp-msg');
    if (!ps.length || sincAcomp || !window.Acompanhamento) return Promise.resolve();
    if (!ouv.regs || !ouv.regs.length) { if (manual) msg.textContent = 'A lista de ouvidorias da plataforma ainda não carregou. Espere um instante e tente de novo.'; return Promise.resolve(); }
    if (!g) { if (manual) msg.textContent = 'Clique em "Conectar Google" (lá em cima, nesta aba) para atualizar as planilhas.'; return Promise.resolve(); }
    sincAcomp = true; msg.textContent = 'Atualizando as planilhas…';
    return window.Acompanhamento.sincronizar({ google: g, planilhas: ps, registros: ouv.regs || [] }).then(function (res) {
      ouv.ultimaAcomp = {}; res.forEach(function (x) { ouv.ultimaAcomp[x.nome] = x; });
      var tot = res.reduce(function (n, x) { return n + (x.linhas || 0); }, 0), erros = res.filter(function (x) { return x.erro; });
      msg.textContent = (tot ? num(tot) + ' linha(s) atualizada(s).' : 'Tudo em dia.') + (erros.length ? ' Não consegui abrir: ' + erros.map(function (x) { return x.nome; }).join(', ') + '.' : '');
      desenharAcomp();
    }).catch(function (e) { msg.textContent = 'Não consegui atualizar: ' + e.message; })
      .then(function () { sincAcomp = false; });
  }
  $('#ouv-acomp-novo').addEventListener('click', function () { editarAcomp(null); });
  // cria a planilha do Google do mês novo (ou usa um link já existente) e passa a receber as listas novas
  $('#ouv-mes-novo').addEventListener('click', function () {
    var msg = $('#ouv-acomp-msg');
    var mesNome = window.prompt('Planilha nova da Reclamação 6. De qual mês? (ex.: Novembro)', ''); if (!mesNome) return;
    var nome = 'Ouvidorias Faturas ' + mesNome.trim() + ' – Reclamação 6';
    var link = window.prompt('Link da planilha do Google, se ela já existir.\n\nDeixe em branco para a plataforma criar "' + nome + '" no seu Google Drive.', '');
    if (link === null) return; link = link.trim();
    if (link && !/\/spreadsheets\/d\//.test(link)) { avisar('Esse não parece um link do Google Planilhas.'); return; }
    var criar = link ? Promise.resolve(link) : (window.Google.conectado()
      ? window.Google.criarPlanilha(nome).then(function (g) { return g.link || g.spreadsheetUrl || ('https://docs.google.com/spreadsheets/d/' + (g.id || g.spreadsheetId) + '/edit'); })
      : Promise.reject(new Error('Conecte o Google (no topo) para a plataforma criar a planilha, ou cole o link de uma planilha que já existe.')));
    msg.textContent = link ? 'Salvando…' : 'Criando a planilha no Google…';
    criar.then(function (lk) {
      var novo = { id: 'r6-' + carimbo(new Date()), nome: nome, link: lk, desde: hojeISO(), ate: '', atual: true };
      return salvarOuv(function (d) { d.mesesR6.forEach(function (x) { x.atual = false; }); d.mesesR6.push(novo); return d; }, 'Ouvidorias: planilha nova ' + nome)
        .then(function () { renderOuv(); desenharAcomp(); msg.textContent = nome + ' criada. As listas novas vão para ela, com ouvidorias de ' + isoParaBR(novo.desde) + ' em diante (mude em 📅 Período).'; });
    }).catch(function (e) { msg.textContent = 'Não consegui criar: ' + e.message; });
  });
  $('#ouv-acomp-sinc').addEventListener('click', function () { ouvCarregarRegs().then(function () { return sincronizarPlanilhaOuv(true); }).then(function () { return sincronizarAcomp(true); }); });

  $$('.abas button[data-aba="ouvidorias"]').forEach(function (b) { b.addEventListener('click', function () { carregarOuv(); }); });
  window.addEventListener('focus', function () { if (!$('[data-painel="ouvidorias"]').hidden && estado.repo) carregarOuv(); });

  /* ---------- início ---------- */
  carregarChaveGoogle();
  montarPainel('faturas');
  montarPainel('contratos');
  var tk = ler('ww_token');
  if (window.PLAT) entrar(window.PLAT.token, window.PLAT.usuario.nome); // dentro do portal: usa o login de lá
  else if (location.hash === '#configurar') { mostrar('tela-config'); } // acesso da administradora (não aparece no site)
  else if (tk) entrar(tk, ler('ww_nome') || ''); else iniciarLogin('');
  window.addEventListener('hashchange', function () { if (location.hash === '#configurar') mostrar('tela-config'); });
})();
