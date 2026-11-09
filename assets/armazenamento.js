/**
 * Armazenamento da plataforma no repositório PRIVADO do GitHub e cofre de acesso.
 *
 * - As planilhas enviadas e geradas, o índice e o histórico ficam em
 *   danielamotaxp-code/faturas-contratos-acessos, na pasta dados/.
 * - O acesso a esse repositório usa um token do GitHub que fica criptografado
 *   (AES-GCM, chave derivada da senha com PBKDF2) em assets/cofre.json.
 *   Sem a senha, o cofre não abre.
 */
(function (raiz) {
  'use strict';

  var REPO_DADOS = 'danielamotaxp-code/faturas-contratos-acessos';
  var REPO_SITE = 'danielamotaxp-code/faturas-contratos';
  var ITERACOES = 600000;
  var API = 'https://api.github.com/repos/';

  /* ---------------- base64 / texto ---------------- */
  function bytesParaBase64(bytes) {
    var s = '', passo = 0x8000;
    for (var i = 0; i < bytes.length; i += passo) s += String.fromCharCode.apply(null, bytes.subarray(i, i + passo));
    return btoa(s);
  }
  function base64ParaBytes(b64) {
    var s = atob(String(b64).replace(/\s/g, '')), out = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
  }
  var enc = new TextEncoder(), dec = new TextDecoder();

  /* ---------------- cofre ---------------- */
  function chaveDaSenha(senha, salt) {
    return crypto.subtle.importKey('raw', enc.encode(senha), 'PBKDF2', false, ['deriveKey']).then(function (base) {
      return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: ITERACOES, hash: 'SHA-256' },
        base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
    });
  }

  function trancar(token, senha) {
    var salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
    return chaveDaSenha(senha, salt).then(function (k) {
      return crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv }, k, enc.encode(token));
    }).then(function (c) {
      return { salt: bytesParaBase64(salt), iv: bytesParaBase64(iv), dados: bytesParaBase64(new Uint8Array(c)) };
    });
  }

  function abrir(registro, senha) {
    return chaveDaSenha(senha, base64ParaBytes(registro.salt)).then(function (k) {
      return crypto.subtle.decrypt({ name: 'AES-GCM', iv: base64ParaBytes(registro.iv) }, k, base64ParaBytes(registro.dados));
    }).then(function (b) { return dec.decode(b); }, function () { throw new Error('Usuário ou senha incorretos.'); });
  }

  function lerCofre() {
    return fetch('assets/cofre.json?v=' + Date.now(), { cache: 'no-store' }).then(function (r) {
      if (r.status === 404) return null;
      if (!r.ok) throw new Error('Não consegui carregar o cofre de acesso.');
      return r.json();
    });
  }

  /* ---------------- GitHub ---------------- */
  function Repositorio(token, repo) {
    this.token = token;
    this.repo = repo || REPO_DADOS;
  }

  Repositorio.prototype.req = function (caminho, opcoes, aceitar) {
    var o = opcoes || {};
    o.cache = 'no-store';
    o.headers = Object.assign({
      Authorization: 'Bearer ' + this.token,
      Accept: aceitar || 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    }, o.headers || {});
    return fetch(API + this.repo + caminho, o).then(function (r) {
      if (r.status === 401) throw new Error('O token do GitHub expirou ou foi revogado. Peça para refazer a configuração da plataforma.');
      return r;
    });
  };

  function caminhoUrl(p) { return '/contents/' + p.split('/').map(encodeURIComponent).join('/'); }

  /** Confere se o token consegue gravar no repositório. */
  Repositorio.prototype.verificar = function () {
    var repo = this.repo;
    return this.req('').then(function (r) {
      if (r.status === 404) throw new Error('O token não tem acesso ao repositório ' + repo + '.');
      if (!r.ok) throw new Error('Falha ao acessar ' + repo + ' (' + r.status + ').');
      return r.json();
    }).then(function (j) {
      if (!j.permissions || !j.permissions.push) throw new Error('O token não tem permissão de gravação em ' + repo + '.');
      return true;
    });
  };

  /** Metadados (sha) de um arquivo, ou null se não existir. */
  Repositorio.prototype.info = function (caminho) {
    return this.req(caminhoUrl(caminho)).then(function (r) {
      if (r.status === 404) return null;
      if (!r.ok) throw new Error('Falha ao ler ' + caminho + ' (' + r.status + ').');
      return r.json();
    });
  };

  /** Conteúdo bruto de um arquivo (até 100 MB). */
  Repositorio.prototype.lerBytes = function (caminho) {
    return this.req(caminhoUrl(caminho), {}, 'application/vnd.github.raw+json').then(function (r) {
      if (r.status === 404) return null;
      if (!r.ok) throw new Error('Falha ao baixar ' + caminho + ' (' + r.status + ').');
      return r.arrayBuffer().then(function (b) { return new Uint8Array(b); });
    });
  };

  Repositorio.prototype.lerJSON = function (caminho, padrao) {
    return this.lerBytes(caminho).then(function (b) { return b ? JSON.parse(dec.decode(b)) : padrao; });
  };

  Repositorio.prototype.gravarBytes = function (caminho, bytes, mensagem, sha) {
    var corpo = { message: mensagem, content: bytesParaBase64(bytes) };
    if (sha) corpo.sha = sha;
    return this.req(caminhoUrl(caminho), { method: 'PUT', body: JSON.stringify(corpo) }).then(function (r) {
      if (r.status === 409 || r.status === 422) { var e = new Error('conflito'); e.conflito = true; throw e; }
      if (!r.ok) throw new Error('Falha ao salvar ' + caminho + ' (' + r.status + ').');
      return r.json();
    });
  };

  /**
   * Atualiza um JSON com segurança: lê a versão atual, aplica a mudança e grava.
   * Se outra pessoa salvou ao mesmo tempo, lê de novo e repete (até 4 vezes).
   */
  Repositorio.prototype.atualizarJSON = function (caminho, padrao, mudar, mensagem) {
    var self = this, tentativa = 0;
    function rodar() {
      tentativa++;
      return self.info(caminho).then(function (meta) {
        var ler = meta
          ? self.req('/git/blobs/' + meta.sha).then(function (r) { return r.json(); }).then(function (b) { return JSON.parse(dec.decode(base64ParaBytes(b.content))); })
          : Promise.resolve(JSON.parse(JSON.stringify(padrao)));
        return ler.then(function (atual) {
          var novo = mudar(atual) || atual;
          return self.gravarBytes(caminho, enc.encode(JSON.stringify(novo)), mensagem, meta && meta.sha).then(function () { return novo; });
        });
      }).catch(function (e) {
        if (e.conflito && tentativa < 4) return new Promise(function (ok) { setTimeout(ok, 600 * tentativa); }).then(rodar);
        throw e;
      });
    }
    return rodar();
  };

  var api = {
    REPO_DADOS: REPO_DADOS, REPO_SITE: REPO_SITE,
    Repositorio: Repositorio, trancar: trancar, abrir: abrir, lerCofre: lerCofre,
    bytesParaBase64: bytesParaBase64
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Armazenamento = api;
})(this);
