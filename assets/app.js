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

  function entrar(token, nome) {
    estado.repo = new A.Repositorio(token);
    estado.nome = nome;
    guardar('ww_token', token); guardar('ww_nome', nome);
    $('#nome-usuario').textContent = nome;
    mostrar('tela-app');
    $('.abas button[data-aba="faturas"]').click();
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

  function avisar(msg) {
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
          '<label class="semana">Semana <select class="campo">' + opcoes + '</select></label>' +
          '<button class="btn btn-primario" type="submit">Gerar planilha</button>' +
        '</form>' +
        '<p class="dica dica-link">A planilha precisa estar em <b>Compartilhar → Acesso geral → Qualquer pessoa com o link (Leitor)</b>. O mês é lido do nome da planilha.</p>' +
        '<p class="dica dica-arquivo" hidden>No Google Planilhas: <b>Arquivo → Fazer download → Microsoft Excel (.xlsx)</b>. Assim todas as abas vão juntas.</p>' +
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
      var semana = $('select', e.target).value;
      (porLink ? enviarLink(aba, link, semana) : enviarArquivo(aba, input.files[0], semana)).then(function (ok) {
        if (ok) {
          e.target.reset(); rotulo.textContent = SEM_ARQUIVO; caixa.classList.remove('escolhido');
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
  function enviarArquivo(aba, arq, semana) {
    var passo = passoDe(aba), resultado;
    passo('Lendo e conferindo "' + arq.name + '"…');
    var sem = semanaDoEnvio(semana), buffer;
    return concluir(aba, arq.arrayBuffer().then(function (buf) {
      buffer = buf;
      passo('Juntando as semanas anteriores do mês…');
      return carregarAnteriores(aba, R.mesDoNome(arq.name), sem);
    }).then(function (anteriores) {
      var buf = buffer;
      passo('Conferindo "' + arq.name + '" (Semana ' + sem + ')…');
      return esperarTela().then(function () {
        resultado = P.processar(aba, arq.name, buf, sem, anteriores);
        conferirAlvo(aba, resultado);
        resultado.original = new Uint8Array(buf);
        passo('Salvando a planilha enviada e a gerada…');
        return salvarProcessamento(aba, arq.name, resultado, passo, '').then(function (reg) { return [reg, resultado]; });
      });
    }));
  }

  /** Lê a planilha pelo link do Google, confere, gera e salva. */
  function enviarLink(aba, link, semana) {
    var passo = passoDe(aba), resultado;
    passo('Lendo a planilha no Google Planilhas…');
    var sem = semanaDoEnvio(semana), planilha;
    return concluir(aba, window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      planilha = g;
      passo('Juntando as semanas anteriores do mês…');
      return carregarAnteriores(aba, R.mesDoNome(g.titulo), sem);
    }).then(function (anteriores) {
      var g = planilha;
      passo('Conferindo "' + g.titulo + '" (Semana ' + sem + ')…');
      return esperarTela().then(function () {
        var wb = P.livroDeAbas(g.nomesAbas, g.abas);
        resultado = P.processarLivro(aba, g.titulo, wb, sem, anteriores);
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
      var reg = par[0]; resultado = par[1];
      var st = mostrarStatus(aba, 'ok', '<b>' + esc(reg.nomeGerado) + '</b> gerada. Semana ' + esc(reg.semana) + ': ' +
        num(reg.semFatura) + ' ' + t.rotSem + ' de ' + num(reg.total) + ' ' + t.rotTotal + '. ' +
        (reg.desconsiderados ? '<br>' + num(reg.desconsiderados) + (aba === 'faturas' ? ' instalações' : ' contratos') + ' desconectados ou inativos ficaram de fora. ' : '') +
        '<button class="btn btn-peq btn-primario" type="button">Baixar planilha gerada</button>' +
        (reg.googleAtualizada ? '<br>✔ Planilha do Google atualizada no mesmo link: ' + esc(reg.googleAtualizada) + '. <a href="' + esc(reg.linkOrigem) + '" target="_blank" rel="noopener">Abrir no Google</a>' : '') +
        (reg.avisoGoogle ? '<br><span class="aviso-google">' + esc(reg.avisoGoogle) + '</span>' : ''));
      $('button', st).addEventListener('click', function () { baixarBytes(resultado.bytes, reg.nomeGerado + '.xlsx', XLSX_MIME); });
      estado.progresso = null;
      renderListas();
      return true;
    }).catch(function (err) {
      mostrarStatus(aba, 'falha', esc(err.message));
      return false;
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
      (r.linkOrigem ? '<a class="btn btn-peq" href="' + esc(r.linkOrigem) + '" target="_blank" rel="noopener">Abrir no Google</a>' : '') +
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
      return '<details class="mes"' + (i === 0 ? ' open' : '') + '><summary>' + esc(rotuloMes(mes)) +
        '<small>' + itens.length + (itens.length === 1 ? ' planilha' : ' planilhas') + '</small></summary>' + semanas + '</details>';
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

    $$('[data-inserir]', el).forEach(function (b) {
      b.addEventListener('click', function () {
        var painel = $('[data-painel="' + aba + '"]');
        estado.alvo[aba] = { mes: b.dataset.mes, semana: b.dataset.inserir };
        $('form select', painel).value = b.dataset.inserir;
        var alvo = $('.alvo', painel);
        alvo.hidden = false;
        alvo.innerHTML = 'Inserindo na <b>Semana ' + esc(b.dataset.inserir) + '</b> de <b>' + esc(rotuloMes(b.dataset.mes)) +
          '</b>. Cole o link ou escolha o arquivo e clique em Gerar planilha. <button type="button" class="btn btn-link">Cancelar</button>';
        $('button', alvo).addEventListener('click', function () { estado.alvo[aba] = null; alvo.hidden = true; $('form select', painel).value = ''; });
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
    return { colab: col([/colaborador/, /responsavel/]), cliente: col([/^cliente$/, /^cliente/, /titular/, /^nome/]),
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

  function mostrarPrevia(t) {
    ouv.previa = t;
    var box = $('#ouv-previa');
    if (!t) { box.innerHTML = arguments.length ? '<p class="erro">Não achei o cabeçalho (coluna "Cliente" ou "Instalação") na planilha.</p>' : ''; $('#ouv-salvar').disabled = true; return; }
    var c = colunasOuv(t.cab);
    var faltam = [c.cliente < 0 && 'Cliente', c.inst < 0 && 'Instalação'].filter(Boolean);
    box.innerHTML = '<div class="ouv-resumo"><span class="ouv-chip"><b>' + num(t.linhas.length) + '</b> clientes</span>' + chips(contar(t.linhas, c.classif)) + '</div>' +
      (c.colab >= 0 ? '<div class="ouv-resumo">' + chips(contar(t.linhas, c.colab, '(sem colaborador)'), true) + '</div>' : '') +
      (faltam.length ? '<p class="erro">Faltou a coluna ' + faltam.join(' e ') + '.</p>' : '') +
      (c.classif < 0 ? '<p class="dica">Não achei a coluna da classificação. A lista será salva mesmo assim.</p>' : '');
    if (t.descartadas) box.insertAdjacentHTML('beforeend', '<p class="dica">' + num(t.descartadas) + ' linhas sem classificação (fatura gerada normalmente) ficaram de fora.</p>');
    $('#ouv-salvar').disabled = !!faltam.length || !t.linhas.length;
  }
  $('#ouv-ler-link').addEventListener('click', function () {
    var link = $('#ouv-link').value.trim(), btn = this;
    if (!/\/spreadsheets\/d\//.test(link)) { $('#ouv-previa').innerHTML = '<p class="erro">Cole o link da planilha auditada (Google Planilhas).</p>'; return; }
    btn.disabled = true; $('#ouv-previa').innerHTML = '<p class="dica">Lendo a planilha…</p>';
    window.Google.lerPlanilha(link, estado.chaveGoogle).then(function (g) {
      // aba com mais linhas classificadas (a última auditada)
      var melhor = null;
      g.nomesAbas.forEach(function (n) {
        var t = limparTabela(g.abas[n]); if (!t) return;
        var c = colunasOuv(t.cab), q = c.classif >= 0 ? t.linhas.filter(function (l) { return l[c.classif]; }).length : 0;
        if (!melhor || q > melhor.q || (q === melhor.q && g.nomesAbas.indexOf(n) > g.nomesAbas.indexOf(melhor.n))) melhor = { n: n, t: t, q: q };
      });
      $('#ouv-arquivo').value = '';
      mostrarPrevia(melhor ? paraFormatoOuv(melhor.t) : null);
      if (melhor) $('#ouv-previa').insertAdjacentHTML('afterbegin', '<p class="dica">Aba lida: <b>' + esc(melhor.n) + '</b></p>');
    }).catch(function (e) { $('#ouv-previa').innerHTML = '<p class="erro">' + esc(e.message) + '</p>'; })
      .then(function () { btn.disabled = false; });
  });
  $('#ouv-arquivo').addEventListener('change', function () {
    var f = this.files[0]; if (!f) return;
    f.arrayBuffer().then(function (buf) {
      var wb = window.Planilhas.lerArquivo(f.name, buf), ws = wb.Sheets[wb.SheetNames[0]];
      $('#ouv-link').value = '';
      mostrarPrevia(paraFormatoOuv(limparTabela(window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }))));
    }).catch(function (e) { $('#ouv-previa').innerHTML = '<p class="erro">Não consegui ler o arquivo: ' + esc(e.message) + '</p>'; });
  });

  function linkPlanilhaOuv() { return (ouv.dados && ouv.dados.planilha) || PLANILHA_OUV_PADRAO; }
  function desenharPlanilhaOuv() {
    var l = linkPlanilhaOuv();
    $('#ouv-planilha').innerHTML = '📗 Planilha de ouvidorias: <a href="' + esc(l) + '" target="_blank" rel="noopener">Ouvidorias Faturas – Reclamação 6</a>' +
      ' <button class="btn btn-link btn-peq" type="button" id="ouv-trocar">Trocar link</button>' +
      (window.Google.conectado() ? ' <button class="btn btn-peq" type="button" id="ouv-sinc">Atualizar a planilha agora</button>' : '') +
      (window.Google.conectado() ? '' : ' <span class="dica" style="margin:0">Para criar as abas na planilha, clique em <b>Entrar com Google</b> no topo.</span>');
    if ($('#ouv-sinc')) $('#ouv-sinc').onclick = function () { sincronizarPlanilhaOuv(true); };
    $('#ouv-trocar').onclick = function () {
      var novo = window.prompt('Link da planilha de ouvidorias (Google Planilhas):', l);
      if (novo == null) return; novo = novo.trim();
      if (!/\/spreadsheets\/d\//.test(novo)) { avisar('Esse não parece um link do Google Planilhas.'); return; }
      salvarOuv(function (d) { d.planilha = novo; return d; }, 'Ouvidorias: link da planilha').then(renderOuv);
    };
  }
  function salvarOuv(mudar, msg) {
    return estado.repo.atualizarJSON(ARQ_OUV, { planilha: PLANILHA_OUV_PADRAO, listas: [] }, function (d) {
      d.listas = d.listas || []; return mudar(d);
    }, msg + ' (' + estado.nome + ')').then(function (d) { ouv.dados = d; return d; });
  }
  function carregarOuv() {
    $('#ouv-data').value = $('#ouv-data').value || hojeISO();
    return Promise.all([
      estado.repo.lerJSON(ARQ_OUV, { planilha: PLANILHA_OUV_PADRAO, listas: [] }),
      new A.Repositorio(estado.repo.token, 'ouvidorias-cemig-dados').lerJSON('registros.json', { registros: [] }).catch(function () { return { registros: [] }; })
    ]).then(function (r) { ouv.dados = r[0]; ouv.regs = r[1].registros || []; renderOuv(); sincronizarPlanilhaOuv(false); })
      .catch(function (e) { $('#ouv-listas').innerHTML = '<p class="erro">' + esc(e.message) + '</p>'; });
  }

  /** Cria a aba com a data na planilha de ouvidorias (se já existir uma aba com esse nome, usa "(2)", "(3)"…). */
  function criarAbaGoogle(lista) {
    if (!window.Google.conectado()) return Promise.reject(new Error('Clique em "Entrar com Google" no topo para criar a aba na planilha.'));
    return window.Google.lerPlanilha(linkPlanilhaOuv(), estado.chaveGoogle).then(function (g) {
      var nome = lista.data, n = 2;
      while (g.folhas[nome]) nome = lista.data + ' (' + (n++) + ')';
      return window.Google.escreverAba(g, nome, [lista.cab].concat(lista.linhas)).then(function (sheetId) {
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
    var lista = { id: carimbo(new Date()), dataISO: iso, data: isoParaBR(iso),
      cab: t.cab, linhas: t.linhas, usuario: estado.nome, criadoEm: new Date().toISOString() };
    btn.disabled = true; msg.textContent = 'Salvando a lista…';
    salvarOuv(function (d) { d.listas.push(lista); return d; }, 'Ouvidorias: lista de ' + lista.data)
      .then(function () {
        if (!window.Google.conectado()) { msg.textContent = 'Lista salva. Para criar a aba na planilha, clique em "Entrar com Google" e depois em "Criar aba na planilha" na lista.'; return; }
        msg.textContent = 'Criando a aba "' + lista.data + '" na planilha de ouvidorias…';
        return criarAbaGoogle(lista).then(function () { msg.textContent = 'Lista salva e aba criada na planilha ✓'; })
          .catch(function (e) { msg.textContent = 'Lista salva, mas não consegui criar a aba: ' + e.message; });
      })
      .then(function () { $('#ouv-link').value = ''; $('#ouv-arquivo').value = ''; mostrarPrevia(); renderOuv(); })
      .catch(function (e) { msg.textContent = 'Não consegui salvar: ' + e.message; btn.disabled = false; });
  });

  // ouvidorias feitas até 7 dias antes da data da lista também contam (a auditoria e as ouvidorias podem começar antes da aba ser criada)
  function seteDiasAntes(iso) { var d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() - 7); return d.toISOString().slice(0, 10); }
  function regOuv(inst, desde) {
    var k = R.chaveInstalacao(inst); if (!k || !ouv.regs) return null;
    desde = seteDiasAntes(desde);
    var rs = ouv.regs.filter(function (r) { return +r.reclamacao === 6 && R.chaveInstalacao(r.instalacao) === k && String(r.criadoEm || '').slice(0, 10) >= desde; });
    if (!rs.length) return null;
    return rs.filter(function (r) { return r.status !== 'pendente'; })[0] || rs[0];
  }
  function letraCol(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  /**
   * Atualiza a aba de cada lista na planilha de ouvidorias, conforme as ouvidorias registradas em Ouvidorias Cemig:
   * Colaborador (quem registrou, se estiver vazio), situação, protocolo, data, link "Abrir ouvidoria" e a linha em verde.
   */
  var sincronizando = false;
  var VERDE = { red: 0.204, green: 0.659, blue: 0.325 }, LARANJA = { red: 0.992, green: 0.851, blue: 0.663 };
  function sincronizarPlanilhaOuv(manual) {
    var listas = ((ouv.dados && ouv.dados.listas) || []).filter(function (l) { return l.aba; });
    if (sincronizando || !listas.length || !window.Google.conectado()) return Promise.resolve(false);
    sincronizando = true;
    return window.Google.lerPlanilha(linkPlanilhaOuv(), estado.chaveGoogle).then(function (g) {
      var dados = [], formatos = [];
      listas.forEach(function (l) {
        var m = g.abas[l.aba], folha = g.folhas[l.aba]; if (!m || !m.length || !folha) return;
        var hl = m.findIndex(function (row) { return (row || []).some(function (h) { return nrm(h) === nrm('Instalação'); }); });
        if (hl < 0) return;
        var cab = m[hl].map(function (h) { return nrm(h); }), ci = cab.indexOf(nrm('Instalação')), cc = cab.indexOf('colaborador');
        var cs = COLS_AUTO.map(function (h) { return cab.indexOf(nrm(h)); });
        var ultimaCol = cab.length;
        m.slice(hl + 1).forEach(function (linha, j) {
          var r = regOuv(linha[ci], l.dataISO); if (!r) return;
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
          if (mudou) formatos.push({ repeatCell: { range: { sheetId: folha.sheetId, startRowIndex: lin - 1, endRowIndex: lin, startColumnIndex: 0, endColumnIndex: ultimaCol },
            cell: { userEnteredFormat: { backgroundColor: pend ? LARANJA : VERDE } }, fields: 'userEnteredFormat.backgroundColor' } });
        });
      });
      return window.Google.gravarIntervalos(g.id, dados)
        .then(function () { return formatos.length ? window.Google.lote(g.id, formatos) : null; })
        .then(function () {
          if (manual) avisar(dados.length || formatos.length ? 'Planilha de ouvidorias atualizada.' : 'A planilha de ouvidorias já estava em dia.');
          return true;
        });
    }).catch(function (e) { if (manual) avisar('Não consegui atualizar a planilha de ouvidorias: ' + e.message); return false; })
      .then(function (x) { sincronizando = false; return x; });
  }

  function situacaoOuv(inst, desde) {
    var r0 = regOuv(inst, desde); if (!r0) return null;
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
      var feitos = l.linhas.filter(function (x) { var s = situacaoOuv(x[c.inst], l.dataISO); return s && s.cls === 'ok'; }).length;
      var linhas = l.linhas.map(function (x, i) { return { x: x, i: i }; }).filter(function (o) { return !q || nrm(o.x.join(' ')).indexOf(q) >= 0; });
      var abaLink = l.gid != null ? 'https://docs.google.com/spreadsheets/d/' + esc(l.planilhaId || (linkPlanilhaOuv().match(/\/d\/([\w-]+)/) || [])[1] || '') + '/edit#gid=' + l.gid : '';
      return '<details class="card ouv-lista"' + (li === 0 ? ' open' : '') + '><summary>' +
        '<h3>Lista de ' + esc(l.data) + '</h3><span class="ouv-chip"><b>' + num(l.linhas.length) + '</b> clientes</span>' +
        '<span class="ouv-chip" style="background:#e7f4ec;color:#1d5a37">' + num(feitos) + ' com ouvidoria enviada</span>' +
        '<span class="dica">por ' + esc(l.usuario || '') + '</span>' +
        '<span class="acoes">' + (l.aba ? '<a class="btn btn-peq" href="' + abaLink + '" target="_blank" rel="noopener">Abrir aba "' + esc(l.aba) + '"</a>'
          : '<button class="btn btn-peq" type="button" data-criar="' + esc(l.id) + '">Criar aba na planilha</button>') +
        '<button class="btn btn-link btn-peq" type="button" data-excluir="' + esc(l.id) + '">Excluir lista</button></span></summary>' +
        '<div class="ouv-resumo">' + chips(contar(l.linhas, c.classif)) + '</div>' +
        '<div class="ouv-filtro"><input class="campo" type="search" placeholder="Buscar cliente, instalação, colaborador…" data-busca="' + esc(l.id) + '" value="' + esc(ouv.filtro[l.id] || '') + '"></div>' +
        '<div class="tabela-wrap"><table class="ouv-tab"><thead><tr>' + (c.colab >= 0 ? '<th>Colaborador</th>' : '') + '<th>Cliente</th><th>Instalação</th><th>UC</th>' + (c.classif >= 0 ? '<th>Classificação</th>' : '') + '<th>Ouvidoria</th><th></th></tr></thead><tbody>' +
        linhas.map(function (o) {
          var x = o.x, s = situacaoOuv(x[c.inst], l.dataISO);
          var rg = regOuv(x[c.inst], l.dataISO);
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
        criarAbaGoogle(l).then(function () { avisar('Aba criada na planilha de ouvidorias.'); renderOuv(); })
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
