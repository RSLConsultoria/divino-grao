/* Divino Grão — ligação com a planilha (Google Sheets via Apps Script)
 *
 * 1. Mostra a tela de login e confere usuário e senha na aba Usuarios.
 * 2. Carrega todas as abas da planilha e monta o estado inicial do sistema.
 * 3. A cada mudança, descobre o que mudou (linha a linha) e grava na planilha.
 * 4. A cada 15 s busca o que outros computadores gravaram e aplica na tela.
 *
 * O que ainda não subiu fica guardado neste navegador e é reenviado sozinho.
 */
(function () {
  "use strict";

  var CFG = window.DG_CONFIG || {};
  var LS_SESSAO = "dg:sessao:v1";
  var LS_PEND = "dg:pendentes:v1";
  var INTERVALO_PULL = 15000;

  // ───────────────────────── utilidades ─────────────────────────
  function hash(str) {
    var h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (var i = 0; i < str.length; i++) {
      var ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  }
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch (e) { return null; } }
  function lsSet(k, v) { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function num(v) { var n = Number(v); return isFinite(n) ? Math.round(n * 100) / 100 : 0; }
  function dataHora(ms) {
    if (!ms) return "";
    var d = new Date(ms);
    if (isNaN(d.getTime())) return "";
    var p = function (n) { return String(n).padStart(2, "0"); };
    return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear() + " " + p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function totalItem(i) {
    if (!i) return 0;
    if (i.receitaItem != null) return num(i.receitaItem);
    var q = i.unidade === "kg" ? Number(i.peso || 0) : Number(i.qtd || 1);
    return num((Number(i.preco) || 0) * q - (Number(i.desconto) || 0));
  }

  // ───────────────────────── o que vai para cada aba ─────────────────────────
  // get(state) → lista de itens · set(state, lista) → pedaço do estado · cols → colunas legíveis
  function semSit(v) { var o = Object.assign({}, v); delete o._sit; return o; }
  function nomeProduto(s, id) { var p = (s.produtos || []).find(function (x) { return x.id === id; }); return p ? p.nome : id; }

  var ESTADO_CHAVES = {
    cfg: { get: function (s) { return s.cfg; }, set: function (s, v) { return { cfg: Object.assign({}, s.cfg, v) }; } },
    sessao: {
      get: function (s) { var o = Object.assign({}, s.sessao); delete o.vendas; return o; },
      set: function (s, v) { return { sessao: Object.assign({}, v, { vendas: (s.sessao && s.sessao.vendas) || [] }) }; }
    },
    estoque_saldos: {
      get: function (s) { return (s.estoque && s.estoque.saldoLote) || {}; },
      set: function (s, v) { return { estoque: Object.assign({}, s.estoque, { saldoLote: v || {} }) }; }
    },
    financeiro_ajustes: {
      get: function (s) {
        var f = s.fin || {};
        return { mes: f.mes, rateio: f.rateio, area: f.area, acum: f.acum, baseRateio: f.baseRateio, variaveis: f.variaveis, catsVar: f.catsVar, catsFixa: f.catsFixa };
      },
      set: function (s, v) { return { fin: Object.assign({}, s.fin, v) }; }
    },
    fichas_tecnicas: { get: function (s) { return s.fichasTec || {}; }, set: function (s, v) { return { fichasTec: v || {} }; } },
    compras: { get: function (s) { return s.compras; }, set: function (s, v) { return v ? { compras: v } : {}; } }
  };

  var SPECS = [
    {
      tab: "Produtos",
      get: function (s) { return s.produtos || []; },
      set: function (s, a) { return { produtos: a }; },
      cols: function (p) {
        return { nome: p.nome, tipo: p.tipo, categoria: p.cat || "", grupo: p.grupo, unidade: p.unidade, preco: num(p.preco), custo: num(p.custo), codigo: p.cod || "", plu: p.plu || "", ean: p.ean || "" };
      }
    },
    {
      tab: "Comandas",
      get: function (s) { return s.comandas || []; },
      set: function (s, a) { return { comandas: a }; },
      cols: function (c) {
        var itens = c.itens || [];
        return { nome: c.nome, tipo: c.tipo, status: c.status, aberta_em: dataHora(c.abertaEm), itens: itens.length, total: num(itens.reduce(function (t, i) { return t + totalItem(i); }, 0)) };
      }
    },
    {
      tab: "Vendas",
      get: function (s) {
        var arq = (s.vendasArquivo || []).map(function (v) { return Object.assign({}, v, { _sit: "fechada" }); });
        var tur = ((s.sessao && s.sessao.vendas) || []).map(function (v) { return Object.assign({}, v, { _sit: "turno" }); });
        return arq.concat(tur);
      },
      set: function (s, a) {
        return {
          vendasArquivo: a.filter(function (v) { return v._sit !== "turno"; }).map(semSit),
          sessao: Object.assign({}, s.sessao, { vendas: a.filter(function (v) { return v._sit === "turno"; }).map(semSit) })
        };
      },
      cols: function (v) {
        return {
          situacao: v._sit === "turno" ? "caixa aberto" : "caixa fechado", data: v.data || "", hora: v.hora || "", numero: v.num || "",
          comanda: v.rotulo || "", subtotal: num(v.subtotal), desconto: num(v.desconto), total: num(v.total),
          granel: num(v.granel), cafeteria: num(v.cafe), pagamento: v.pgto || "", fiado: num(v.fiado), troco: num(v.troco), itens: v.itens || ""
        };
      }
    },
    {
      tab: "ItensVendidos",
      somenteEscrita: true,
      get: function (s) {
        var out = [];
        SPECS[2].get(s).forEach(function (v) {
          var vid = v.id != null ? String(v.id) : "h" + hash(JSON.stringify(semSit(v)));
          (v.itensLista || []).forEach(function (i, k) {
            out.push({
              id: vid + "-" + (k + 1), venda: vid, data: v.data || "", hora: v.hora || "", numero: v.num || "",
              produto_id: i.id, produto: i.nome, tipo: i.tipo, unidade: i.unidade,
              quantidade: i.unidade === "kg" ? Number(i.peso || 0) : Number(i.qtd || 1),
              preco: num(i.preco), total: totalItem(i), custo_total: num(i.custoTotal), lote: i.loteNum || "", obs: i.obs || ""
            });
          });
        });
        return out;
      },
      cols: function (r) { var o = Object.assign({}, r); delete o.id; return o; }
    },
    {
      tab: "Fechamentos",
      get: function (s) { return s.fichas || []; },
      set: function (s, a) { return { fichas: a }; },
      cols: function (f) {
        return {
          data: f.data || "", abertura: f.abertura || "", fechamento: f.fechamento || "", abriu: f.resp || "", fechou: f.respFech || "",
          fundo: num(f.fundo), vendas: num(f.vendas), granel: num(f.totGranel), cafeteria: num(f.totCafe), n_vendas: f.nVendas || 0,
          sangrias: num(f.sangrias), suprimentos: num(f.suprimentos), esperado: num(f.esperado), contado: num(f.contado), diferenca: num(f.dif), motivo: f.motivo || ""
        };
      }
    },
    {
      tab: "Fiado",
      get: function (s) { return (s.fiado && s.fiado.contas) || []; },
      set: function (s, a) { return { fiado: Object.assign({}, s.fiado, { contas: a }) }; },
      cols: function (c) {
        var l = c.lancamentos || [];
        var compras = l.filter(function (x) { return x.tipo === "compra"; }).reduce(function (t, x) { return t + Number(x.valor || 0); }, 0);
        var pagos = l.filter(function (x) { return x.tipo !== "compra"; }).reduce(function (t, x) { return t + Number(x.valor || 0); }, 0);
        return { nome: c.nome, telefone: c.tel || "", dia_pagamento: c.dia || "", compras: num(compras), pagamentos: num(pagos), saldo: num(compras - pagos), lancamentos: l.length };
      }
    },
    {
      tab: "Fornecedores",
      get: function (s) { return s.fornecedores || []; },
      set: function (s, a) { return { fornecedores: a }; },
      cols: function (f) { return { nome: f.nome, contato: f.contato || "", telefone: f.tel || "", endereco: f.end || "", categorias: (f.cats || []).join(", "), obs: f.obs || "" }; }
    },
    {
      tab: "Estoque_Lotes",
      get: function (s) {
        var l = (s.estoque && s.estoque.lotes) || {};
        return Object.keys(l).map(function (k) { return { id: k, lotes: l[k] }; });
      },
      set: function (s, a) {
        var o = {};
        a.forEach(function (r) { o[r.id] = r.lotes || []; });
        return { estoque: Object.assign({}, s.estoque, { lotes: o }) };
      },
      cols: function (r, s) {
        var lotes = r.lotes || [];
        var saldos = (s.estoque && s.estoque.saldoLote) || {};
        var saldo = lotes.reduce(function (t, x) { var v = saldos[x.id]; return t + Number(v != null ? v : (x.qtd || 0)); }, 0);
        return { produto: nomeProduto(s, r.id), lotes: lotes.length, saldo: Math.round(saldo * 1000) / 1000 };
      }
    },
    {
      tab: "Estoque_Movimentos",
      get: function (s) { return (s.estoque && s.estoque.movimentos) || []; },
      set: function (s, a) { return { estoque: Object.assign({}, s.estoque, { movimentos: a }) }; },
      cols: function (m) { return { tipo: m.tipo || "", produto: m.nome || "", detalhe: m.detalhe || "", quantidade: m.valor || "" }; }
    },
    {
      tab: "Lancamentos_Despesas",
      get: function (s) { return (s.fin && s.fin.micro) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { micro: a }) }; },
      cols: function (d) { return { data: d.data || "", descricao: d.nome || d.desc || "", categoria: d.cat || "", valor: num(d.valor), forma: d.forma || "" }; }
    },
    {
      tab: "Lancamentos_Entradas",
      get: function (s) { return (s.fin && s.fin.emprestimos) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { emprestimos: a }) }; },
      cols: function (e) { return { data: e.data || "", origem: e.nome || e.origem || "", tipo: e.tipo || "", valor: num(e.valor), parcelas: e.parcelas || "", valor_parcela: num(e.parcela) }; }
    },
    {
      tab: "Despesas_Fixas",
      get: function (s) { return (s.fin && s.fin.fixas) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { fixas: a }) }; },
      cols: function (f) { return { nome: f.nome, categoria: f.cat || "", valor: num(f.valor), dia: f.dia || "", pago: f.pago ? "sim" : "não" }; }
    },
    {
      tab: "Dividas",
      get: function (s) { return (s.fin && s.fin.dividas) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { dividas: a }) }; },
      cols: function (d) { return { nome: d.nome, credor: d.credor || "", parcela: num(d.parcela), pagas: d.pagas || 0, total: d.total || 0, vencimento: d.venc || "" }; }
    },
    {
      tab: "Contas_a_Pagar",
      get: function (s) { return (s.fin && s.fin.contas) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { contas: a }) }; },
      cols: function (c) { return { nome: c.nome || "", categoria: c.cat || "", valor: num(c.valor), vencimento: c.venc || c.data || "", pago: c.pago ? "sim" : "não" }; }
    },
    {
      tab: "Imobilizado",
      get: function (s) { return (s.fin && s.fin.imobilizado) || []; },
      set: function (s, a) { return { fin: Object.assign({}, s.fin, { imobilizado: a }) }; },
      cols: function (b) { return { nome: b.nome || "", valor: num(b.valor) }; }
    },
    {
      tab: "Registro",
      get: function (s) { return s.logs || []; },
      set: function (s, a) { return { logs: a }; },
      cols: function (l) { return { hora: l.hora || "", texto: l.txt || "" }; }
    },
    {
      tab: "Estado",
      get: function (s) {
        return Object.keys(ESTADO_CHAVES).map(function (k) { return { id: k, valor: ESTADO_CHAVES[k].get(s) }; });
      },
      setParcial: true,
      set: function (s, a) {
        var patch = {};
        var cur = s;
        a.forEach(function (r) {
          var def = ESTADO_CHAVES[r.id];
          if (!def || r.valor == null) return;
          var p = def.set(cur, r.valor);
          Object.assign(patch, p);
          cur = Object.assign({}, cur, p);
        });
        return patch;
      },
      cols: function (r) { return { chave: r.id }; }
    }
  ];

  function idDe(item, vistos) {
    if (item && item.id != null && item.id !== "") return String(item.id);
    var base = "h" + hash(JSON.stringify(item));
    var n = (vistos[base] || 0) + 1;
    vistos[base] = n;
    return n > 1 ? base + "_" + n : base;
  }
  function mapa(spec, s) {
    var m = new Map();
    var vistos = {};
    spec.get(s).forEach(function (item, ordem) {
      if (item == null) return;
      var id = idDe(item, vistos);
      m.set(id, { item: item, json: JSON.stringify(item), ordem: ordem });
    });
    return m;
  }

  // ───────────────────────── API ─────────────────────────
  var sessao = lsGet(LS_SESSAO); // { token, nome, usuario, perfil }
  function api(acao, dados) {
    var corpo = Object.assign({ acao: acao, token: sessao && sessao.token }, dados || {});
    return fetch(CFG.api, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(corpo), redirect: "follow" })
      .then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); })
      .then(function (j) {
        if (j && j.erro === "sessao") { lsSet(LS_SESSAO, null); sessao = null; var e = new Error("Sessão expirada · entre de novo"); e.sessao = true; throw e; }
        if (j && j.erro) throw new Error(j.erro);
        return j;
      });
  }

  // ───────────────────────── sincronização ─────────────────────────
  var ultimo = {};      // tab → Map(id → json) — o que a planilha tem
  var pend = lsGet(LS_PEND) || {}; // tab → { up: {id: row}, del: {id: 1} }
  var voando = {};
  var comp = null, ultimoPull = 0, enviando = false, tFlush = null, tDiff = null, estadoCarga = null;
  var status = { txt: "", erro: false };

  function temPendente() { return Object.keys(pend).some(function (t) { return Object.keys(pend[t].up).length || Object.keys(pend[t].del).length; }); }
  function contarPend() { var n = 0; Object.keys(pend).forEach(function (t) { n += Object.keys(pend[t].up).length + Object.keys(pend[t].del).length; }); return n; }

  function diff(s) {
    var mudou = false;
    SPECS.forEach(function (spec) {
      var atual = mapa(spec, s);
      var antes = ultimo[spec.tab] || new Map();
      var p = pend[spec.tab] || (pend[spec.tab] = { up: {}, del: {} });
      atual.forEach(function (v, id) {
        if (antes.get(id) !== v.json) {
          p.up[id] = { id: id, json: v.json, ordem: v.ordem, campos: spec.cols ? spec.cols(v.item, s) : {} };
          delete p.del[id];
          antes.set(id, v.json);
          mudou = true;
        }
      });
      Array.from(antes.keys()).forEach(function (id) {
        if (!atual.has(id)) { p.del[id] = 1; delete p.up[id]; antes.delete(id); mudou = true; }
      });
      ultimo[spec.tab] = antes;
    });
    if (mudou) { lsSet(LS_PEND, pend); agendarEnvio(400); }
  }

  function agendarEnvio(ms) { clearTimeout(tFlush); tFlush = setTimeout(enviar, ms); }

  function enviar() {
    if (enviando || !temPendente() || !sessao) { pintar(); return; }
    enviando = true;
    var lote = pend;
    voando = lote;
    pend = {};
    var ops = Object.keys(lote).map(function (t) {
      return { tab: t, ups: Object.keys(lote[t].up).map(function (k) { return lote[t].up[k]; }), dels: Object.keys(lote[t].del) };
    }).filter(function (o) { return o.ups.length || o.dels.length; });
    pintar("salvando");
    api("salvar", { ops: ops })
      .then(function () {
        enviando = false;
        voando = {};
        lsSet(LS_PEND, pend);
        status = { txt: "", erro: false };
        if (temPendente()) agendarEnvio(300);
        pintar();
      })
      .catch(function (err) {
        enviando = false;
        voando = {};
        // devolve o lote para a fila, sem passar por cima do que mudou depois
        Object.keys(lote).forEach(function (t) {
          var p = pend[t] || (pend[t] = { up: {}, del: {} });
          Object.keys(lote[t].up).forEach(function (id) { if (!(id in p.up) && !(id in p.del)) p.up[id] = lote[t].up[id]; });
          Object.keys(lote[t].del).forEach(function (id) { if (!(id in p.up) && !(id in p.del)) p.del[id] = 1; });
        });
        lsSet(LS_PEND, pend);
        status = { txt: err.message, erro: true };
        pintar();
        if (err.sessao) { mostrarLogin(err.message); return; }
        agendarEnvio(10000);
      });
  }

  function aplicarRemoto(s, resp, ignorarPendentes) {
    var patch = {};
    var cur = s;
    SPECS.forEach(function (spec) {
      if (spec.somenteEscrita) return;
      var rows = (resp.tabs && resp.tabs[spec.tab]) || [];
      var dels = (resp.excluidos && resp.excluidos[spec.tab]) || [];
      if (!rows.length && !dels.length && !resp.completo) return;
      var p0 = pend[spec.tab] || { up: {}, del: {} };
      var v0 = voando[spec.tab] || { up: {}, del: {} };
      var p = { up: Object.assign({}, v0.up, p0.up), del: Object.assign({}, v0.del, p0.del) };
      var ant = ultimo[spec.tab] || new Map();
      var atual = resp.completo ? new Map() : mapa(spec, cur);
      var alterou = !!resp.completo;
      rows.forEach(function (r) {
        if (!ignorarPendentes && (r.id in p.up || r.id in p.del)) return;
        var item;
        try { item = JSON.parse(r.json); } catch (e) { return; }
        ant.set(r.id, r.json);
        var x = atual.get(r.id);
        if (x && x.json === r.json) return;
        atual.set(r.id, { item: item, json: r.json, ordem: x ? x.ordem : (r.ordem != null && r.ordem !== "" ? Number(r.ordem) : 1e9) });
        alterou = true;
      });
      dels.forEach(function (id) {
        if (!ignorarPendentes && id in p.up) return;
        ant.delete(id);
        if (atual.delete(id)) alterou = true;
      });
      ultimo[spec.tab] = ant;
      if (!alterou) return;
      var lista = Array.from(atual.values()).sort(function (a, b) { return a.ordem - b.ordem; }).map(function (v) { return v.item; });
      var p2 = spec.set(cur, lista);
      Object.assign(patch, p2);
      cur = Object.assign({}, cur, p2);
    });
    return { patch: patch, estado: cur };
  }

  function puxar() {
    if (!comp || !sessao || enviando) return;
    // primeiro registra o que mudou aqui, para não ser atropelado pelo que vem de fora
    diff(comp.state);
    var tabs = SPECS.filter(function (s) { return !s.somenteEscrita; }).map(function (s) { return s.tab; });
    api("carregar", { desde: ultimoPull, tabs: tabs })
      .then(function (resp) {
        ultimoPull = resp.agora || ultimoPull;
        diff(comp.state); // o que mudou aqui enquanto a resposta vinha também tem prioridade
        var r = aplicarRemoto(comp.state, resp, false);
        if (Object.keys(r.patch).length) comp.setState(r.patch);
        if (!status.erro || !temPendente()) status = { txt: "", erro: false };
        pintar();
      })
      .catch(function (err) {
        status = { txt: err.message, erro: true };
        pintar();
        if (err.sessao) mostrarLogin(err.message);
      });
  }

  // ───────────────────────── estado inicial ─────────────────────────
  function mesAtual() {
    var m = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
    var d = new Date();
    return m[d.getMonth()] + "/" + d.getFullYear();
  }
  // tira tudo que é demonstração: o sistema começa vazio, só com o catálogo e as 4 mesas
  function estadoLimpo(s) {
    var fin = s.fin || {};
    return Object.assign({}, s, {
      comandas: [1, 2, 3, 4].map(function (n) { return { id: "m" + n, tipo: "mesa", nome: "Mesa " + n, itens: [], status: "livre", abertaEm: null, taxa: false }; }),
      cart: [], espera: [], pagos: [], desconto: 0, comandaOrigem: null, cartRotulo: null,
      sessao: { abertaEm: null, fundo: 0, resp: null, vendas: [], movimentos: [], recebFiado: [] },
      fichas: [], vendasArquivo: [], logs: [], exemplosOn: false,
      fiado: Object.assign({}, s.fiado, { sel: null, contas: [] }),
      fornecedores: [],
      estoque: Object.assign({}, s.estoque, { lotes: {}, saldoLote: {}, movimentos: [] }),
      fin: Object.assign({}, fin, {
        mes: mesAtual(),
        acum: { granel: { receita: 0, cmv: 0, diretas: 0 }, cafe: { receita: 0, cmv: 0, diretas: 0 } },
        fixas: [], dividas: [], micro: [], emprestimos: [], contas: [], imobilizado: []
      }),
      cfg: Object.assign({}, s.cfg, { user: (sessao && sessao.nome) || (s.cfg && s.cfg.user) })
    });
  }

  // ───────────────────────── tela de login e selo de status ─────────────────────────
  var CSS = [
    "#dg-login{position:fixed;inset:0;z-index:9999;background:#F3EEE5;display:flex;align-items:center;justify-content:center;padding:16px;font-family:Geist,'Helvetica Neue',Helvetica,sans-serif;color:#302D29}",
    "#dg-login form{width:100%;max-width:360px;background:#FCFAF6;border:1px solid #DDD1BE;box-shadow:0 30px 70px -30px rgba(48,45,41,.45);padding:28px 26px}",
    "#dg-login .marca{width:44px;height:44px;background:#713D42;color:#FCFAF6;display:flex;align-items:center;justify-content:center;font-weight:700;letter-spacing:.02em;margin-bottom:16px}",
    "#dg-login h1{margin:0 0 4px;font-size:20px;font-weight:600;letter-spacing:-.01em}",
    "#dg-login p{margin:0 0 20px;font-size:13px;color:#6F6A61;line-height:1.5}",
    "#dg-login label{display:block;font-size:11.5px;font-weight:600;color:#6F6A61;letter-spacing:.04em;margin:0 0 6px}",
    "#dg-login input{width:100%;box-sizing:border-box;height:46px;border:1.5px solid #DDD1BE;background:#fff;padding:0 12px;font-size:15px;font-family:inherit;margin-bottom:14px;outline:none}",
    "#dg-login input:focus{border-color:#713D42}",
    "#dg-login button{width:100%;height:50px;border:none;background:#713D42;color:#FCFAF6;font-size:14.5px;font-weight:600;cursor:pointer;font-family:inherit}",
    "#dg-login button:disabled{opacity:.6;cursor:default}",
    "#dg-login .erro{min-height:18px;font-size:12.5px;color:#B5462F;margin:2px 0 10px}",
    "#dg-selo{position:fixed;left:0;bottom:0;width:92px;box-sizing:border-box;z-index:9000;padding:8px 6px 10px;text-align:center;font:500 10px/1.35 Geist,'Helvetica Neue',sans-serif;color:rgba(252,250,246,.72)}",
    "#dg-selo .st{display:flex;align-items:center;justify-content:center;gap:5px;color:rgba(252,250,246,.9)}",
    "#dg-selo i{width:6px;height:6px;border-radius:50%;background:#8BA173;display:inline-block;flex:none}",
    "#dg-selo.salvando i{background:#D9A441}",
    "#dg-selo.erro i{background:#E07A62}",
    "#dg-selo.erro .st{color:#F2B3A3}",
    "#dg-selo .nm{margin-top:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
    "#dg-selo a{color:rgba(252,250,246,.9);cursor:pointer;text-decoration:underline}"
  ].join("\n");
  function injetarCss() {
    if (document.getElementById("dg-css")) return;
    var st = document.createElement("style");
    st.id = "dg-css";
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  var aoEntrar = null;
  function mostrarLogin(msg) {
    injetarCss();
    var el = document.getElementById("dg-login");
    if (!el) {
      el = document.createElement("div");
      el.id = "dg-login";
      el.innerHTML =
        '<form autocomplete="on">' +
        '<div class="marca">DG</div>' +
        "<h1>Armazém Divino Grão</h1>" +
        "<p>Entre com seu usuário para abrir o sistema.</p>" +
        '<label for="dg-u">USUÁRIO</label><input id="dg-u" name="username" autocomplete="username" autocapitalize="off" spellcheck="false" required>' +
        '<label for="dg-s">SENHA</label><input id="dg-s" name="password" type="password" autocomplete="current-password" required>' +
        '<div class="erro" id="dg-erro"></div>' +
        '<button type="submit" id="dg-btn">Entrar</button>' +
        "</form>";
      document.body.appendChild(el);
      el.querySelector("form").addEventListener("submit", function (e) {
        e.preventDefault();
        var u = el.querySelector("#dg-u").value.trim().toLowerCase();
        var s = el.querySelector("#dg-s").value;
        var btn = el.querySelector("#dg-btn");
        btn.disabled = true;
        btn.textContent = "Entrando…";
        el.querySelector("#dg-erro").textContent = "";
        api("login", { usuario: u, senha: s })
          .then(function (r) {
            sessao = { token: r.token, nome: r.nome, usuario: r.usuario, perfil: r.perfil };
            lsSet(LS_SESSAO, sessao);
            el.remove();
            if (comp) { status = { txt: "", erro: false }; agendarEnvio(100); puxar(); }
            else if (aoEntrar) aoEntrar();
          })
          .catch(function (err) {
            btn.disabled = false;
            btn.textContent = "Entrar";
            el.querySelector("#dg-erro").textContent = err.message === "Failed to fetch" ? "Sem conexão com a planilha. Confira a internet." : err.message;
          });
      });
    }
    el.querySelector("#dg-erro").textContent = msg || "";
    setTimeout(function () { var i = el.querySelector("#dg-u"); if (i) i.focus(); }, 50);
  }

  function telaCarregando(txt, erro) {
    injetarCss();
    var el = document.getElementById("dg-login");
    if (!el) { el = document.createElement("div"); el.id = "dg-login"; document.body.appendChild(el); }
    el.innerHTML = '<form><div class="marca">DG</div><h1>' + (erro ? "Não foi possível abrir" : "Abrindo o sistema…") + "</h1><p>" + txt + "</p>" +
      (erro ? '<button type="button" onclick="location.reload()">Tentar de novo</button>' : "") + "</form>";
  }

  function pintar(modo) {
    var el = document.getElementById("dg-selo");
    if (!el) {
      if (!comp) return;
      injetarCss();
      el = document.createElement("div");
      el.id = "dg-selo";
      document.body.appendChild(el);
    }
    var n = contarPend();
    var cls = status.erro ? "erro" : (modo === "salvando" || enviando || n ? "salvando" : "");
    var txt = status.erro ? "Sem conexão" : (cls ? "Salvando…" : "Salvo");
    el.className = cls;
    el.title = status.erro
      ? n + " alteraç" + (n === 1 ? "ão" : "ões") + " aguardando para subir · " + status.txt
      : (cls ? "Gravando na planilha" : "Tudo gravado na planilha");
    var nome = String((sessao && sessao.nome) || "").replace(/[<>&]/g, "");
    el.innerHTML = '<div class="st"><i></i><span>' + txt + (status.erro && n ? " · " + n : "") + '</span></div><div class="nm">' + nome + ' · <a id="dg-sair">sair</a></div>';
    el.querySelector("#dg-sair").onclick = function () { DGBanco.sair(); };
  }

  // ───────────────────────── partida ─────────────────────────
  function carregarRuntime() {
    var s = document.createElement("script");
    s.src = "./support.js";
    document.head.appendChild(s);
  }

  function iniciar() {
    telaCarregando("Buscando os dados na planilha.");
    // reenvia o que ficou pendente de outra vez antes de ler
    var pre = temPendente() ? new Promise(function (ok) {
      var ops = Object.keys(pend).map(function (t) { return { tab: t, ups: Object.keys(pend[t].up).map(function (k) { return pend[t].up[k]; }), dels: Object.keys(pend[t].del) }; });
      api("salvar", { ops: ops }).then(function () { pend = {}; lsSet(LS_PEND, pend); ok(); }).catch(function () { ok(); });
    }) : Promise.resolve();
    pre.then(function () {
      var tabs = SPECS.filter(function (s) { return !s.somenteEscrita; }).map(function (s) { return s.tab; });
      return api("carregar", { tabs: tabs });
    }).then(function (resp) {
      ultimoPull = resp.agora;
      estadoCarga = resp;
      var el = document.getElementById("dg-login");
      if (el) el.remove();
      carregarRuntime();
    }).catch(function (err) {
      if (err.sessao) { aoEntrar = iniciar; mostrarLogin(err.message); return; }
      telaCarregando(err.message === "Failed to fetch" ? "Sem conexão com a planilha. Confira a internet e tente de novo." : err.message, true);
    });
  }

  var DGBanco = {
    ativo: !!CFG.api,
    usuario: function () { return sessao; },
    // chamado no construtor do sistema: devolve o estado já com os dados da planilha
    estadoInicial: function (padrao) {
      var base = estadoLimpo(padrao);
      var resp = estadoCarga || { tabs: {}, excluidos: {} };
      if (resp.vazio) return base; // primeira vez: começa limpo e tudo sobe para a planilha
      var r = aplicarRemoto(base, Object.assign({}, resp, { completo: true }), true);
      var s = r.estado;
      return Object.assign({}, s, { cfg: Object.assign({}, s.cfg, { user: (sessao && sessao.nome) || s.cfg.user }) });
    },
    mudou: function (s) {
      clearTimeout(tDiff);
      tDiff = setTimeout(function () { diff(s); pintar(); }, 350);
    },
    conectar: function (c) {
      comp = c;
      diff(c.state);
      pintar();
      setInterval(puxar, INTERVALO_PULL);
      window.addEventListener("online", function () { agendarEnvio(200); });
      window.addEventListener("beforeunload", function (e) {
        if (temPendente() || enviando) { e.preventDefault(); e.returnValue = ""; }
      });
    },
    sair: function () {
      if (temPendente() && !window.confirm("Ainda há alterações sendo salvas na planilha. Sair mesmo assim? Elas ficam guardadas neste computador e sobem no próximo login.")) return;
      var t = sessao && sessao.token;
      lsSet(LS_SESSAO, null);
      sessao = null;
      if (t) fetch(CFG.api, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ acao: "sair", token: t }) }).catch(function () {});
      location.reload();
    }
  };
  window.DGBanco = DGBanco;

  function partida() {
    if (!DGBanco.ativo) { carregarRuntime(); return; } // sem planilha configurada: roda só neste navegador
    if (!sessao) { aoEntrar = iniciar; mostrarLogin(); return; }
    iniciar();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", partida);
  else partida();
})();
