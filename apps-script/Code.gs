/**
 * Divino Grão — banco de dados na planilha
 *
 * Este script fica dentro da planilha (Extensões › Apps Script) e é publicado
 * como "App da Web". O sistema (GitHub Pages) conversa só com ele.
 *
 * Primeira vez:
 *   1. Rode a função `configurar` (cria todas as abas).
 *   2. Na aba Usuarios, escreva a senha de cada pessoa na coluna senha_nova.
 *      No primeiro login ela vira um código (senha_hash) e some da planilha.
 *   3. Implantar › Nova implantação › App da Web · Executar como: Eu ·
 *      Quem pode acessar: Qualquer pessoa. Copie a URL /exec para o config.js.
 *
 * Cada linha das abas tem colunas legíveis (para conferir e usar em relatórios)
 * e a coluna _json, que é o registro completo usado pelo sistema. Não edite
 * a _json à mão.
 */

var VERSAO = "1.0";
var DURACAO_SESSAO = 21600; // 6 h, renovadas a cada uso
var LIMITE_CELULA = 45000;

var TABELAS = {
  Produtos: ["nome", "tipo", "categoria", "grupo", "unidade", "preco", "custo", "codigo", "plu", "ean"],
  Comandas: ["nome", "tipo", "status", "aberta_em", "itens", "total"],
  Vendas: ["situacao", "data", "hora", "numero", "comanda", "subtotal", "desconto", "total", "granel", "cafeteria", "pagamento", "fiado", "troco", "itens"],
  ItensVendidos: ["venda", "data", "hora", "numero", "produto_id", "produto", "tipo", "unidade", "quantidade", "preco", "total", "custo_total", "lote", "obs"],
  Fechamentos: ["data", "abertura", "fechamento", "abriu", "fechou", "fundo", "vendas", "granel", "cafeteria", "n_vendas", "sangrias", "suprimentos", "esperado", "contado", "diferenca", "motivo"],
  Fiado: ["nome", "telefone", "dia_pagamento", "compras", "pagamentos", "saldo", "lancamentos"],
  Fornecedores: ["nome", "contato", "telefone", "endereco", "categorias", "obs"],
  Estoque_Lotes: ["produto", "lotes", "saldo"],
  Estoque_Movimentos: ["tipo", "produto", "detalhe", "quantidade"],
  Lancamentos_Despesas: ["data", "descricao", "categoria", "valor", "forma"],
  Lancamentos_Entradas: ["data", "origem", "tipo", "valor", "parcelas", "valor_parcela"],
  Despesas_Fixas: ["nome", "categoria", "valor", "dia", "pago"],
  Dividas: ["nome", "credor", "parcela", "pagas", "total", "vencimento"],
  Contas_a_Pagar: ["nome", "categoria", "valor", "vencimento", "pago"],
  Imobilizado: ["nome", "valor"],
  Registro: ["hora", "texto"],
  Estado: ["chave"]
};
var FIXAS_INICIO = ["id"];
var FIXAS_FIM = ["atualizado_em", "atualizado_por", "_ordem", "_json"];

// ───────────────────────── configuração ─────────────────────────
function configurar() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(TABELAS).forEach(function (nome) { garantirAba_(ss, nome); });

  var u = ss.getSheetByName("Usuarios");
  if (!u) {
    u = ss.insertSheet("Usuarios", 0);
    u.getRange(1, 1, 1, 7).setValues([["usuario", "nome", "perfil", "ativo", "senha_nova", "senha_hash", "sal"]]);
    u.getRange(2, 1, 1, 4).setValues([["lucca", "Lucca", "admin", "sim"]]);
    u.setFrozenRows(1);
    u.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#F3EEE5");
    u.getRange("E1").setNote("Escreva aqui a senha da pessoa. No primeiro login ela vira código (senha_hash) e esta célula fica vazia. Para trocar a senha, escreva uma nova aqui.");
    u.getRange("D1").setNote("sim = pode entrar · não = bloqueado");
    u.getRange(2, 1, 50, 7).setNumberFormat("@");
    u.hideColumns(6, 2);
    u.autoResizeColumns(1, 5);
  }
  var x = ss.getSheetByName("_Excluidos");
  if (!x) {
    x = ss.insertSheet("_Excluidos");
    x.getRange(1, 1, 1, 4).setValues([["aba", "id", "quando", "por"]]).setFontWeight("bold");
    x.setFrozenRows(1);
  }
  // a aba vazia que vem com a planilha nova sai, se não tiver nada
  ["Página1", "Sheet1", "Planilha1"].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  return "Abas prontas: " + Object.keys(TABELAS).join(", ") + ", Usuarios, _Excluidos";
}

function garantirAba_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (sh) return sh;
  sh = ss.insertSheet(nome);
  var cab = FIXAS_INICIO.concat(TABELAS[nome] || [], FIXAS_FIM);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold").setBackground("#F3EEE5");
  sh.setFrozenRows(1);
  sh.getRange(1, cab.indexOf("_json") + 1).setNote("Registro completo usado pelo sistema. Não edite.");
  sh.hideColumns(cab.indexOf("_ordem") + 1, 2);
  return sh;
}

// ───────────────────────── entrada HTTP ─────────────────────────
function doGet() {
  return json_({ ok: true, app: "Divino Grão", versao: VERSAO });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    out = rotear_(req);
  } catch (err) {
    out = { erro: String((err && err.message) || err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rotear_(req) {
  if (req.acao === "login") return login_(req.usuario, req.senha);
  var u = usuarioDaSessao_(req.token);
  if (!u) return { erro: "sessao" };
  if (req.acao === "carregar") return carregar_(req.desde || 0, req.tabs || Object.keys(TABELAS));
  if (req.acao === "salvar") return salvar_(req.ops || [], u);
  if (req.acao === "sair") { CacheService.getScriptCache().remove("s_" + req.token); return { ok: true }; }
  return { erro: "Ação desconhecida" };
}

// ───────────────────────── login ─────────────────────────
function sha_(txt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ((x + 256) % 256).toString(16).padStart(2, "0"); }).join("");
}

function login_(usuario, senha) {
  usuario = String(usuario || "").trim().toLowerCase();
  senha = String(senha || "");
  if (!usuario || !senha) return { erro: "Informe usuário e senha" };
  var cache = CacheService.getScriptCache();
  var chaveTent = "t_" + usuario;
  var tent = Number(cache.get(chaveTent) || 0);
  if (tent >= 8) return { erro: "Muitas tentativas. Espere 10 minutos." };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName("Usuarios");
    if (!sh) return { erro: "A planilha ainda não foi configurada (rode configurar)" };
    var v = sh.getDataRange().getValues();
    var cab = v[0];
    var c = {};
    cab.forEach(function (n, i) { c[n] = i; });
    for (var r = 1; r < v.length; r++) {
      if (String(v[r][c.usuario]).trim().toLowerCase() !== usuario) continue;
      if (String(v[r][c.ativo]).trim().toLowerCase() === "não" || String(v[r][c.ativo]).trim().toLowerCase() === "nao") return { erro: "Usuário bloqueado" };
      var nova = String(v[r][c.senha_nova] || "");
      var hash = String(v[r][c.senha_hash] || "");
      var sal = String(v[r][c.sal] || "");
      if (nova) {
        sal = Utilities.getUuid();
        hash = sha_(sal + nova);
        sh.getRange(r + 1, c.senha_nova + 1).setValue("");
        sh.getRange(r + 1, c.senha_hash + 1).setValue(hash);
        sh.getRange(r + 1, c.sal + 1).setValue(sal);
      }
      if (!hash) return { erro: "Este usuário ainda não tem senha. Escreva uma na coluna senha_nova da aba Usuarios." };
      if (sha_(sal + senha) !== hash) break;
      var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
      var dados = { usuario: usuario, nome: String(v[r][c.nome] || usuario), perfil: String(v[r][c.perfil] || "") };
      cache.put("s_" + token, JSON.stringify(dados), DURACAO_SESSAO);
      cache.remove(chaveTent);
      return { token: token, usuario: dados.usuario, nome: dados.nome, perfil: dados.perfil };
    }
  } finally {
    lock.releaseLock();
  }
  cache.put(chaveTent, String(tent + 1), 600);
  return { erro: "Usuário ou senha incorretos" };
}

function usuarioDaSessao_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var raw = cache.get("s_" + token);
  if (!raw) return null;
  cache.put("s_" + token, raw, DURACAO_SESSAO);
  return JSON.parse(raw);
}

// ───────────────────────── leitura ─────────────────────────
function carregar_(desde, tabs) {
  var agora = Date.now();
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties().getProperties();
  var out = { agora: agora, tabs: {}, excluidos: {}, vazio: true };
  tabs.forEach(function (nome) {
    var mod = Number(props["mod_" + nome] || 0);
    var sh = ss.getSheetByName(nome);
    if (!sh) return;
    if (desde && mod && mod <= desde) { out.vazio = false; return; }
    var last = sh.getLastRow();
    if (last < 2) return;
    out.vazio = false;
    var v = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var cab = v[0];
    var iId = cab.indexOf("id"), iAt = cab.indexOf("atualizado_em"), iOrd = cab.indexOf("_ordem");
    var iJson = [];
    cab.forEach(function (n, i) { if (n === "_json" || /^_json_\d+$/.test(n)) iJson.push(i); });
    iJson.sort(function (a, b) { return parteJson_(cab[a]) - parteJson_(cab[b]); });
    var rows = [];
    for (var r = 1; r < v.length; r++) {
      var id = String(v[r][iId]);
      if (!id) continue;
      if (desde && !(Number(v[r][iAt]) > desde)) continue;
      var j = iJson.map(function (i) { return String(v[r][i] || ""); }).join("");
      if (!j) continue;
      rows.push({ id: id, json: j, ordem: v[r][iOrd] });
    }
    if (rows.length) out.tabs[nome] = rows;
  });
  if (desde) {
    var x = ss.getSheetByName("_Excluidos");
    if (x && x.getLastRow() > 1) {
      var xv = x.getRange(2, 1, x.getLastRow() - 1, 3).getValues();
      xv.forEach(function (l) {
        if (Number(l[2]) > desde && tabs.indexOf(l[0]) >= 0) (out.excluidos[l[0]] = out.excluidos[l[0]] || []).push(String(l[1]));
      });
    }
  }
  return out;
}

function parteJson_(n) { return n === "_json" ? 1 : Number(n.split("_")[2]); }

// ───────────────────────── gravação ─────────────────────────
function texto_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" || typeof v === "boolean") return v;
  if (typeof v === "object") v = JSON.stringify(v);
  v = String(v);
  return v === "" ? "" : "'" + v; // apóstrofo: a planilha guarda como texto, sem virar data ou número
}

function salvar_(ops, u) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive();
    var agora = Date.now();
    var props = PropertiesService.getScriptProperties();
    var excl = [];
    ops.forEach(function (op) {
      if (!TABELAS[op.tab]) return;
      var sh = garantirAba_(ss, op.tab);
      var cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      var pos = {};
      cab.forEach(function (n, i) { pos[n] = i; });

      var ups = op.ups || [];
      // colunas novas que o sistema mandou e a aba ainda não tem
      var novas = [];
      var maxPartes = 1;
      ups.forEach(function (r) {
        Object.keys(r.campos || {}).forEach(function (k) { if (!(k in pos) && novas.indexOf(k) < 0) novas.push(k); });
        maxPartes = Math.max(maxPartes, Math.ceil(String(r.json || "").length / LIMITE_CELULA) || 1);
      });
      for (var p = 2; p <= maxPartes; p++) if (!(("_json_" + p) in pos) && novas.indexOf("_json_" + p) < 0) novas.push("_json_" + p);
      if (novas.length) {
        sh.getRange(1, cab.length + 1, 1, novas.length).setValues([novas]).setFontWeight("bold").setBackground("#F3EEE5");
        novas.forEach(function (n, k) {
          pos[n] = cab.length + k;
          if (n.indexOf("_json_") === 0) sh.hideColumns(cab.length + k + 1);
        });
        cab = cab.concat(novas);
      }

      var last = sh.getLastRow();
      var ids = last > 1 ? sh.getRange(2, pos.id + 1, last - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
      var linhaDe = {};
      ids.forEach(function (id, i) { if (id) linhaDe[id] = i + 2; });

      var anexar = [];
      ups.forEach(function (r) {
        var linha = new Array(cab.length);
        for (var i = 0; i < cab.length; i++) linha[i] = "";
        linha[pos.id] = texto_(r.id);
        Object.keys(r.campos || {}).forEach(function (k) { linha[pos[k]] = texto_(r.campos[k]); });
        linha[pos.atualizado_em] = agora;
        linha[pos.atualizado_por] = texto_(u.nome);
        linha[pos._ordem] = r.ordem == null ? "" : r.ordem;
        var j = String(r.json || "");
        for (var k = 0; k * LIMITE_CELULA < j.length || k === 0; k++) {
          var col = k === 0 ? "_json" : "_json_" + (k + 1);
          linha[pos[col]] = texto_(j.slice(k * LIMITE_CELULA, (k + 1) * LIMITE_CELULA));
        }
        var ex = linhaDe[String(r.id)];
        if (ex > 0) sh.getRange(ex, 1, 1, cab.length).setValues([linha]);
        else if (!ex) { anexar.push(linha); linhaDe[String(r.id)] = -1; }
      });
      if (anexar.length) sh.getRange(sh.getLastRow() + 1, 1, anexar.length, cab.length).setValues(anexar);

      var dels = (op.dels || []).map(String);
      if (dels.length) {
        last = sh.getLastRow();
        var ids2 = last > 1 ? sh.getRange(2, pos.id + 1, last - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
        var apagar = [];
        ids2.forEach(function (id, i) { if (dels.indexOf(id) >= 0) apagar.push(i + 2); });
        apagar.sort(function (a, b) { return b - a; }).forEach(function (l) { sh.deleteRow(l); });
        dels.forEach(function (id) { excl.push([op.tab, texto_(id), agora, texto_(u.nome)]); });
      }
      props.setProperty("mod_" + op.tab, String(agora));
    });
    if (excl.length) {
      var x = ss.getSheetByName("_Excluidos") || ss.insertSheet("_Excluidos");
      x.getRange(x.getLastRow() + 1, 1, excl.length, 4).setValues(excl);
    }
    return { ok: true, agora: agora };
  } finally {
    lock.releaseLock();
  }
}
