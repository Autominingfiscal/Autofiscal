# -*- coding: utf-8 -*-
"""Modelo do HTML do Dashboard Fluig. O dashboard_fluig.py troca o marcador
/*__DADOS__*/null pelos dados. Fica num .py (e nao num .html solto) para ir
junto quando a ferramenta e exportada (Manutencao > Exportar ferramenta)."""

MODELO = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Fluig</title>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink-2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --axis: #c3c2b7; --ring: rgba(11,11,11,0.10);
  --s1: #2a78d6; --s1-track: #cde2fb; --s2: #eb6834; --hover: rgba(42,120,214,0.08);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19; --ink: #ffffff; --ink-2: #c3c2b7; --muted: #898781;
    --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
    --s1: #3987e5; --s1-track: #184f95; --s2: #d95926; --hover: rgba(57,135,229,0.14);
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19; --ink: #ffffff; --ink-2: #c3c2b7; --muted: #898781;
  --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
  --s1: #3987e5; --s1-track: #184f95; --s2: #d95926; --hover: rgba(57,135,229,0.14);
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--page); color: var(--ink);
       font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
.pagina { max-width: 1280px; margin: 0 auto; padding: 24px 16px 48px; }
header { display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 16px; margin-bottom: 16px; }
h1 { font-size: 22px; font-weight: 650; margin: 0; }
.sub { color: var(--ink-2); font-size: 13px; }
.tema { margin-left: auto; }
button, select { font: inherit; color: var(--ink); background: var(--surface);
                 border: 1px solid var(--ring); border-radius: 8px; padding: 6px 10px; }
button { cursor: pointer; }
button:hover, select:hover { background: var(--hover); }
.filtros { display: flex; flex-wrap: wrap; gap: 12px; align-items: end; margin-bottom: 20px; }
.filtros label { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
.filtros select { min-width: 150px; max-width: 280px; }
.limpar { align-self: end; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 20px; }
.tile { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; padding: 14px 16px; min-width: 0; }
.tile .rot { font-size: 12px; color: var(--ink-2); }
.tile .val { font-size: 26px; font-weight: 650; margin-top: 2px; white-space: nowrap; }
.tile .det { font-size: 12px; color: var(--muted); margin-top: 2px; }
.tile.hero { grid-column: span 2; }
@media (min-width: 1100px) {
  .kpis { grid-template-columns: repeat(5, 1fr); }
  .tile.hero { grid-row: span 2; display: flex; flex-direction: column; justify-content: center; }
}
.tile.hero .val { font-size: 48px; line-height: 1.1; }
.grade { display: grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap: 16px; }
.card { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; padding: 16px; min-width: 0; }
.card.largo { grid-column: 1 / -1; }
.card-topo { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
.card h2 { font-size: 15px; font-weight: 650; margin: 0; }
.card .nota { font-size: 12px; color: var(--muted); }
.card-topo .acoes { margin-left: auto; display: flex; gap: 6px; }
.card-topo .acoes button { font-size: 12px; padding: 3px 9px; }
.card-topo .acoes button[aria-pressed="true"] { border-color: var(--s1); }
.vazio { color: var(--muted); padding: 24px 0; text-align: center; }
/* barras horizontais */
.barras { display: flex; flex-direction: column; gap: 2px; }
.barra { display: grid; grid-template-columns: minmax(90px, 38%) 1fr auto; align-items: center;
         gap: 10px; padding: 4px 6px; border-radius: 6px; cursor: pointer; }
.barra:hover, .barra:focus-visible { background: var(--hover); outline: none; }
.barra.sel { background: var(--hover); }
.barra .nome { font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.barra .trilho { height: 14px; }
.barra .preenchido { height: 14px; background: var(--s1); border-radius: 0 4px 4px 0; min-width: 2px; }
.barra .num { font-size: 13px; font-variant-numeric: tabular-nums; color: var(--ink); min-width: 48px; text-align: right; }
/* graficos svg */
.svgbox { position: relative; }
svg text { fill: var(--muted); font-size: 11px; font-family: inherit; }
.legenda { display: flex; gap: 16px; font-size: 12px; color: var(--ink-2); margin-bottom: 6px; }
.legenda span::before { content: ""; display: inline-block; width: 14px; height: 2px; vertical-align: middle;
                        margin-right: 6px; background: var(--c); }
.legenda .ret::before { height: 10px; width: 10px; border-radius: 2px; }
/* tabela */
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid var(--grid); }
td.n, th.n { text-align: right; font-variant-numeric: tabular-nums; }
th { color: var(--ink-2); font-weight: 600; }
.tabwrap { max-height: 420px; overflow: auto; }
/* tooltip */
#dica { position: fixed; pointer-events: none; z-index: 10; background: var(--surface); color: var(--ink);
        border: 1px solid var(--ring); border-radius: 8px; padding: 8px 10px; font-size: 12px;
        box-shadow: 0 4px 16px rgba(0,0,0,.12); display: none; max-width: 320px; }
#dica .v { font-weight: 650; font-size: 14px; }
#dica .l { color: var(--ink-2); }
#dica .linha { display: flex; align-items: center; gap: 6px; }
#dica .chave { width: 12px; height: 2px; background: var(--c); }
footer { margin-top: 24px; font-size: 12px; color: var(--muted); }
@media (max-width: 640px) {
  .tile.hero { grid-column: auto; }
  .tile.hero .val { font-size: 36px; }
  .grade { grid-template-columns: 1fr; }
  .barra { grid-template-columns: minmax(80px, 42%) 1fr auto; }
}
</style>
</head>
<body>
<div class="pagina">
  <header>
    <h1>Dashboard Fluig · Entrada de Notas Fiscais</h1>
    <span class="sub" id="sub"></span>
    <button class="tema" id="tema" type="button" title="Alternar tema claro/escuro">◐ Tema</button>
  </header>

  <div class="filtros">
    <label>Ano<select id="f-ano"></select></label>
    <label>Mês<select id="f-mes"></select></label>
    <label>Técnico fiscal<select id="f-tec"></select></label>
    <label>Tipo<select id="f-tipo"></select></label>
    <button class="limpar" id="limpar" type="button">Limpar filtros</button>
  </div>

  <section class="kpis" id="kpis" aria-label="Indicadores"></section>

  <div class="grade">
    <section class="card" id="c-tec"></section>
    <section class="card" id="c-tipo"></section>
    <section class="card largo" id="c-mes"></section>
    <section class="card" id="c-forn"></section>
    <section class="card" id="c-sla"></section>
  </div>

  <footer id="rodape"></footer>
</div>
<div id="dica" role="tooltip"></div>

<script>
const DADOS = /*__DADOS__*/null;

// ------------------------------------------------------------------ dados ----
const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
const LINHAS = DADOS.linhas.map(l => ({
  n: l[0], i: l[1], f: l[2], s: l[3], t: l[4], r: l[5], fo: l[6], v: l[7], k: l[8], h: l[9],
  ai: +l[1].slice(0, 4), mi: +l[1].slice(5, 7),
  af: l[2] ? +l[2].slice(0, 4) : 0, mf: l[2] ? +l[2].slice(5, 7) : 0,
}));
const HOJE = { a: +DADOS.hoje.slice(0, 4), m: +DADOS.hoje.slice(5, 7) };
const fmt = new Intl.NumberFormat("pt-BR");
const fmt1 = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1, minimumFractionDigits: 1 });
const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });

function brlCurto(v) {
  const a = Math.abs(v);
  if (a >= 1e9) return "R$ " + fmt1.format(v / 1e9) + " bi";
  if (a >= 1e6) return "R$ " + fmt1.format(v / 1e6) + " mi";
  if (a >= 1e4) return "R$ " + fmt.format(Math.round(v / 1e3)) + " mil";
  return brl.format(v);
}
const dias = h => h / 24;
function diasTxt(h) {
  if (h == null) return "–";
  const d = dias(h);
  return d < 1 ? fmt.format(Math.round(h)) + " h" : fmt1.format(d) + " d";
}
function mediana(xs) {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b), m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}
function percentil(xs, p) {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.ceil(p * s.length) - 1)];
}

// ---------------------------------------------------------------- filtros ----
const F = { ano: 0, mes: 0, tec: -1, tipo: -1 };
const el = id => document.getElementById(id);
function opcao(sel, valor, texto) {
  const o = document.createElement("option");
  o.value = valor; o.textContent = texto; sel.appendChild(o);
}
function montarFiltros() {
  const anos = [...new Set(LINHAS.map(l => l.ai))].sort((a, b) => b - a);
  opcao(el("f-ano"), 0, "Todos");
  anos.forEach(a => opcao(el("f-ano"), a, String(a)));
  F.ano = anos.includes(HOJE.a) ? HOJE.a : (anos[0] || 0);
  el("f-ano").value = F.ano;
  opcao(el("f-mes"), 0, "Ano inteiro");
  MESES.forEach((m, j) => opcao(el("f-mes"), j + 1, m));
  opcao(el("f-tec"), -1, "Todos");
  DADOS.tecnicos.forEach((t, j) => opcao(el("f-tec"), j, t));
  opcao(el("f-tipo"), -1, "Todos");
  ordemTipos().forEach(j => opcao(el("f-tipo"), j, DADOS.categorias[j]));
  for (const [id, chave] of [["f-ano", "ano"], ["f-mes", "mes"], ["f-tec", "tec"], ["f-tipo", "tipo"]]) {
    el(id).addEventListener("change", e => { F[chave] = +e.target.value; desenhar(); });
  }
  el("limpar").addEventListener("click", () => {
    F.mes = 0; F.tec = -1; F.tipo = -1; F.ano = anos.includes(HOJE.a) ? HOJE.a : (anos[0] || 0);
    sincronizar(); desenhar();
  });
}
function sincronizar() {
  el("f-ano").value = F.ano; el("f-mes").value = F.mes;
  el("f-tec").value = F.tec; el("f-tipo").value = F.tipo;
}
function ordemTipos() {
  const ordem = DADOS.ordem_categorias.map(n => DADOS.categorias.indexOf(n)).filter(j => j >= 0);
  DADOS.categorias.forEach((_, j) => { if (!ordem.includes(j)) ordem.push(j); });
  return ordem;
}
const noPeriodo = (a, m) => (!F.ano || a === F.ano) && (!F.mes || m === F.mes);
const base = () => LINHAS.filter(l => (F.tec < 0 || l.t === F.tec) && (F.tipo < 0 || l.k === F.tipo));
const nomePeriodo = () => F.ano ? (F.mes ? MESES[F.mes - 1] + "/" + F.ano : String(F.ano)) : "todo o período";

// ---------------------------------------------------------------- tooltip ----
const dica = el("dica");
function mostrarDica(ev, linhas) {
  dica.replaceChildren();
  for (const ln of linhas) {
    const d = document.createElement("div");
    d.className = "linha";
    if (ln.cor) { const k = document.createElement("span"); k.className = "chave"; k.style.setProperty("--c", ln.cor); d.appendChild(k); }
    const v = document.createElement("span"); v.className = "v"; v.textContent = ln.v; d.appendChild(v);
    if (ln.l) { const l = document.createElement("span"); l.className = "l"; l.textContent = ln.l; d.appendChild(l); }
    dica.appendChild(d);
  }
  dica.style.display = "block";
  posicionar(ev);
}
function posicionar(ev) {
  const r = dica.getBoundingClientRect();
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + r.width > innerWidth - 8) x = ev.clientX - r.width - 14;
  if (y + r.height > innerHeight - 8) y = ev.clientY - r.height - 14;
  dica.style.left = x + "px"; dica.style.top = y + "px";
}
const esconderDica = () => { dica.style.display = "none"; };

// ------------------------------------------------------------ componentes ----
function tile(rotulo, valor, detalhe, hero, titulo) {
  const t = document.createElement("div");
  t.className = "tile" + (hero ? " hero" : "");
  const r = document.createElement("div"); r.className = "rot"; r.textContent = rotulo;
  const v = document.createElement("div"); v.className = "val"; v.textContent = valor;
  if (titulo) v.title = titulo;
  t.append(r, v);
  if (detalhe) { const d = document.createElement("div"); d.className = "det"; d.textContent = detalhe; t.appendChild(d); }
  return t;
}

// card com titulo, nota, botoes e alternancia grafico/tabela
function card(id, titulo, nota, botoes) {
  const c = el(id);
  const estado = c._tabela || false;
  c.replaceChildren();
  const topo = document.createElement("div"); topo.className = "card-topo";
  const h = document.createElement("h2"); h.textContent = titulo;
  const n = document.createElement("span"); n.className = "nota"; n.textContent = nota;
  const acoes = document.createElement("div"); acoes.className = "acoes";
  for (const b of (botoes || [])) acoes.appendChild(b);
  const bt = document.createElement("button");
  bt.type = "button"; bt.textContent = estado ? "Ver gráfico" : "Ver tabela";
  bt.addEventListener("click", () => { c._tabela = !estado; desenhar(); });
  acoes.appendChild(bt);
  topo.append(h, n, acoes);
  const corpo = document.createElement("div");
  c.append(topo, corpo);
  return { corpo, tabela: estado };
}
function vazio(corpo) {
  const d = document.createElement("div"); d.className = "vazio";
  d.textContent = "Nenhuma solicitação com estes filtros."; corpo.appendChild(d);
}
function tabela(corpo, cabecalho, linhas, numericas) {
  const w = document.createElement("div"); w.className = "tabwrap";
  const t = document.createElement("table");
  const tr = document.createElement("tr");
  cabecalho.forEach((c, j) => { const th = document.createElement("th"); th.textContent = c; if (numericas.includes(j)) th.className = "n"; tr.appendChild(th); });
  const thead = document.createElement("thead"); thead.appendChild(tr); t.appendChild(thead);
  const tb = document.createElement("tbody");
  for (const l of linhas) {
    const r = document.createElement("tr");
    l.forEach((v, j) => { const td = document.createElement("td"); td.textContent = v; if (numericas.includes(j)) td.className = "n"; r.appendChild(td); });
    tb.appendChild(r);
  }
  t.appendChild(tb); w.appendChild(t); corpo.appendChild(w);
}
// itens: [{nome, valor, texto, dica:[...], sel, clique}]
function barras(corpo, itens) {
  if (!itens.length) return vazio(corpo);
  const max = Math.max(...itens.map(i => i.valor), 1);
  const box = document.createElement("div"); box.className = "barras";
  for (const it of itens) {
    const r = document.createElement("div");
    r.className = "barra" + (it.sel ? " sel" : ""); r.tabIndex = 0;
    const nome = document.createElement("span"); nome.className = "nome"; nome.textContent = it.nome; nome.title = it.nome;
    const trilho = document.createElement("span"); trilho.className = "trilho";
    const p = document.createElement("span"); p.className = "preenchido"; p.style.display = "block";
    p.style.width = (100 * it.valor / max) + "%"; trilho.appendChild(p);
    const num = document.createElement("span"); num.className = "num"; num.textContent = it.texto;
    r.append(nome, trilho, num);
    r.addEventListener("pointermove", ev => mostrarDica(ev, it.dica));
    r.addEventListener("pointerleave", esconderDica);
    r.addEventListener("focus", () => { const b = r.getBoundingClientRect(); mostrarDica({ clientX: b.right - 40, clientY: b.top }, it.dica); });
    r.addEventListener("blur", esconderDica);
    if (it.clique) {
      r.addEventListener("click", it.clique);
      r.addEventListener("keydown", ev => { if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); it.clique(); } });
    }
    box.appendChild(r);
  }
  corpo.appendChild(box);
}
const NS = "http://www.w3.org/2000/svg";
function svgEl(tag, attrs) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); return e; }
function escalaY(max) {
  const brutos = max / 4, pot = Math.pow(10, Math.floor(Math.log10(Math.max(brutos, 1))));
  const passo = [1, 2, 2.5, 5, 10].map(f => f * pot).find(p => p >= brutos) || pot;
  return { passo, topo: Math.max(passo, Math.ceil(max / passo) * passo) };
}

// ---------------------------------------------------------------- desenho ----
function desenhar() {
  esconderDica();
  const B = base();
  const criadas = B.filter(l => l.s !== "C" && noPeriodo(l.ai, l.mi));
  const finalizadas = B.filter(l => l.s === "F" && l.f && noPeriodo(l.af, l.mf));
  desenharKpis(B, criadas, finalizadas);
  desenharTecnicos(finalizadas);
  desenharTipos(criadas);
  desenharMeses(B);
  desenharFornecedores(criadas);
  desenharSla(finalizadas);
  const p = el("rodape");
  p.textContent = "Abertas, valor, tipos e fornecedores contam pela data de início da solicitação (sem as canceladas); "
    + "finalizadas, ranking dos técnicos e SLA, pela data de fim. SLA = tempo corrido entre início e fim, só das finalizadas. "
    + "Clique numa barra de técnico ou de tipo para filtrar.";
}

function desenharKpis(B, criadas, finalizadas) {
  const k = el("kpis"); k.replaceChildren();
  const doAno = B.filter(l => l.s !== "C" && (!F.ano || l.ai === F.ano));
  const canceladas = B.filter(l => l.s === "C" && (!F.ano || l.ai === F.ano)).length;
  k.appendChild(tile(F.ano ? "Solicitações em " + F.ano : "Solicitações (todo o período)", fmt.format(doAno.length),
    canceladas ? fmt.format(canceladas) + " cancelada(s) fora da conta" : "", true));
  // abertas no mes: o mes escolhido; com "ano inteiro", o mes atual (ou o ultimo do ano com dado)
  let a = F.ano || HOJE.a, m = F.mes;
  if (!m) {
    if (a === HOJE.a) m = HOJE.m;
    else m = Math.max(0, ...B.filter(l => l.ai === a).map(l => l.mi));
  }
  const noMes = m ? B.filter(l => l.s !== "C" && l.ai === a && l.mi === m).length : 0;
  k.appendChild(tile("Abertas em " + (m ? MESES[m - 1] + "/" + a : "–"), fmt.format(noMes), "solicitações criadas no mês"));
  k.appendChild(tile("Em aberto agora", fmt.format(B.filter(l => l.s === "A").length), "aguardando tratativa"));
  k.appendChild(tile("Finalizadas", fmt.format(finalizadas.length), nomePeriodo()));
  const rep = criadas.filter(l => l.r).length;
  k.appendChild(tile("Reprovadas", fmt.format(rep),
    criadas.length ? fmt1.format(100 * rep / criadas.length) + "% das solicitações do período" : ""));
  const total = criadas.reduce((s, l) => s + (l.v || 0), 0);
  k.appendChild(tile("Valor total das notas", brlCurto(total), nomePeriodo(), false, brl.format(total)));
  const hs = finalizadas.map(l => l.h).filter(h => h != null);
  k.appendChild(tile("SLA mediano", diasTxt(mediana(hs)),
    hs.length ? "média " + diasTxt(hs.reduce((s, h) => s + h, 0) / hs.length) + " · 90% em até " + diasTxt(percentil(hs, 0.9)) : "sem finalizadas"));
}

function desenharTecnicos(finalizadas) {
  const { corpo, tabela: tab } = card("c-tec", "Finalizadas por técnico fiscal", nomePeriodo());
  const por = new Map();
  for (const l of finalizadas) {
    const e = por.get(l.t) || { n: 0, hs: [] };
    e.n++; if (l.h != null) e.hs.push(l.h); por.set(l.t, e);
  }
  const lista = [...por.entries()].sort((a, b) => b[1].n - a[1].n);
  if (tab) return lista.length ? tabela(corpo, ["#", "Técnico", "Finalizadas", "SLA mediano"],
    lista.map(([t, e], j) => [j + 1, DADOS.tecnicos[t], fmt.format(e.n), diasTxt(mediana(e.hs))]), [0, 2, 3]) : vazio(corpo);
  barras(corpo, lista.slice(0, 15).map(([t, e]) => ({
    nome: DADOS.tecnicos[t], valor: e.n, texto: fmt.format(e.n), sel: F.tec === t,
    dica: [{ v: fmt.format(e.n) + " finalizada(s)", l: DADOS.tecnicos[t] }, { v: diasTxt(mediana(e.hs)), l: "SLA mediano" }],
    clique: () => { F.tec = F.tec === t ? -1 : t; sincronizar(); desenhar(); },
  })));
}

function desenharTipos(criadas) {
  const { corpo, tabela: tab } = card("c-tipo", "Solicitações por tipo", nomePeriodo());
  const cont = new Map();
  for (const l of criadas) cont.set(l.k, (cont.get(l.k) || 0) + 1);
  const lista = ordemTipos().filter(j => cont.get(j)).map(j => [j, cont.get(j)]);
  const tot = criadas.length || 1;
  if (tab) return lista.length ? tabela(corpo, ["Tipo", "Solicitações", "%"],
    lista.map(([j, n]) => [DADOS.categorias[j], fmt.format(n), fmt1.format(100 * n / tot) + "%"]), [1, 2]) : vazio(corpo);
  barras(corpo, lista.map(([j, n]) => ({
    nome: DADOS.categorias[j], valor: n, texto: fmt.format(n), sel: F.tipo === j,
    dica: [{ v: fmt.format(n), l: DADOS.categorias[j] }, { v: fmt1.format(100 * n / tot) + "%", l: "do período" }],
    clique: () => { F.tipo = F.tipo === j ? -1 : j; sincronizar(); desenhar(); },
  })));
}

function desenharFornecedores(criadas) {
  const c = el("c-forn");
  const porValor = c._valor || false;
  const b1 = document.createElement("button"); b1.type = "button"; b1.textContent = "Quantidade";
  const b2 = document.createElement("button"); b2.type = "button"; b2.textContent = "Valor";
  b1.setAttribute("aria-pressed", String(!porValor)); b2.setAttribute("aria-pressed", String(porValor));
  b1.addEventListener("click", () => { c._valor = false; desenhar(); });
  b2.addEventListener("click", () => { c._valor = true; desenhar(); });
  const { corpo, tabela: tab } = card("c-forn", "Ranking de fornecedores", nomePeriodo(), [b1, b2]);
  const por = new Map();
  for (const l of criadas) {
    const e = por.get(l.fo) || { n: 0, v: 0 };
    e.n++; e.v += l.v || 0; por.set(l.fo, e);
  }
  const lista = [...por.entries()].sort((a, b) => porValor ? b[1].v - a[1].v : b[1].n - a[1].n);
  if (tab) return lista.length ? tabela(corpo, ["#", "Fornecedor", "Solicitações", "Valor"],
    lista.map(([f, e], j) => [j + 1, DADOS.fornecedores[f], fmt.format(e.n), brl.format(e.v)]), [0, 2, 3]) : vazio(corpo);
  barras(corpo, lista.slice(0, 15).map(([f, e]) => ({
    nome: DADOS.fornecedores[f], valor: porValor ? e.v : e.n, texto: porValor ? brlCurto(e.v) : fmt.format(e.n),
    dica: [{ v: fmt.format(e.n) + " solicitação(ões)", l: DADOS.fornecedores[f] }, { v: brl.format(e.v), l: "valor das notas" }],
  })));
}

function desenharMeses(B) {
  const ano = F.ano;
  const { corpo, tabela: tab } = card("c-mes", "Abertas × finalizadas por mês", ano ? String(ano) : "todo o período");
  // meses do eixo: o ano escolhido inteiro, ou todos os meses com dado
  let meses;
  // ano atual: so ate o mes corrente (mes que ainda nao chegou nao e zero, e vazio)
  if (ano) meses = MESES.map((_, j) => [ano, j + 1]).filter(([a, mm]) => a < HOJE.a || mm <= HOJE.m);
  else {
    const chaves = new Set();
    B.forEach(l => { chaves.add(l.ai * 100 + l.mi); if (l.af) chaves.add(l.af * 100 + l.mf); });
    meses = [...chaves].sort((a, b) => a - b).map(c => [Math.floor(c / 100), c % 100]);
  }
  const ab = meses.map(([a, m]) => B.filter(l => l.s !== "C" && l.ai === a && l.mi === m).length);
  const fi = meses.map(([a, m]) => B.filter(l => l.s === "F" && l.af === a && l.mf === m).length);
  const rotulo = ([a, m]) => MESES[m - 1] + (ano ? "" : "/" + String(a).slice(2));
  if (tab) return tabela(corpo, ["Mês", "Abertas", "Finalizadas"],
    meses.map((mm, j) => [rotulo(mm), fmt.format(ab[j]), fmt.format(fi[j])]), [1, 2]);
  if (!meses.length) return vazio(corpo);
  const leg = document.createElement("div"); leg.className = "legenda";
  for (const [txt, cor] of [["Abertas", "var(--s1)"], ["Finalizadas", "var(--s2)"]]) {
    const s = document.createElement("span"); s.textContent = txt; s.style.setProperty("--c", cor); leg.appendChild(s);
  }
  corpo.appendChild(leg);
  const box = document.createElement("div"); box.className = "svgbox"; corpo.appendChild(box);
  const W = Math.max(box.clientWidth || corpo.clientWidth || 600, 300), H = 240;
  const m = { e: 44, d: 64, t: 10, b: 26 };
  const { passo, topo } = escalaY(Math.max(...ab, ...fi, 1));
  const x = j => m.e + (meses.length === 1 ? (W - m.e - m.d) / 2 : j * (W - m.e - m.d) / (meses.length - 1));
  const y = v => m.t + (H - m.t - m.b) * (1 - v / topo);
  const svg = svgEl("svg", { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: "img",
                             "aria-label": "Linhas de solicitações abertas e finalizadas por mês" });
  for (let v = 0; v <= topo + 1e-9; v += passo) {
    svg.appendChild(svgEl("line", { x1: m.e, x2: W - m.d, y1: y(v), y2: y(v), stroke: v ? "var(--grid)" : "var(--axis)", "stroke-width": 1 }));
    const t = svgEl("text", { x: m.e - 8, y: y(v) + 4, "text-anchor": "end" }); t.textContent = fmt.format(v); svg.appendChild(t);
  }
  const passoRot = Math.ceil(meses.length / Math.max(1, Math.floor((W - m.e - m.d) / 44)));
  meses.forEach((mm, j) => {
    if (j % passoRot) return;
    const t = svgEl("text", { x: x(j), y: H - 8, "text-anchor": "middle" }); t.textContent = rotulo(mm); svg.appendChild(t);
  });
  if (F.mes && ano) {
    svg.appendChild(svgEl("rect", { x: x(F.mes - 1) - 14, y: m.t, width: 28, height: H - m.t - m.b, fill: "var(--hover)" }));
  }
  const series = [{ v: ab, cor: "var(--s1)", nome: "Abertas" }, { v: fi, cor: "var(--s2)", nome: "Finalizadas" }];
  // mes corrente ainda pela metade: o ultimo trecho vai tracejado
  const ultimoParcial = meses.length > 1 && meses[meses.length - 1][0] === HOJE.a && meses[meses.length - 1][1] === HOJE.m;
  for (const s of series) {
    const fim = ultimoParcial ? s.v.length - 1 : s.v.length;
    const d = s.v.slice(0, fim).map((v, j) => (j ? "L" : "M") + x(j) + " " + y(v)).join(" ");
    svg.appendChild(svgEl("path", { d, fill: "none", stroke: s.cor, "stroke-width": 2, "stroke-linejoin": "round", "stroke-linecap": "round" }));
    if (ultimoParcial) {
      const a = s.v.length - 2, b = s.v.length - 1;
      svg.appendChild(svgEl("path", { d: `M${x(a)} ${y(s.v[a])} L${x(b)} ${y(s.v[b])}`, fill: "none", stroke: s.cor,
        "stroke-width": 2, "stroke-dasharray": "4 4", "stroke-linecap": "round" }));
    }
    const ult = s.v.length - 1;
    svg.appendChild(svgEl("circle", { cx: x(ult), cy: y(s.v[ult]), r: 4, fill: s.cor, stroke: "var(--surface)", "stroke-width": 2 }));
  }
  // rotulos so no fim de cada linha (afastados se coincidirem)
  let ya = y(ab[ab.length - 1]), yf = y(fi[fi.length - 1]);
  if (Math.abs(ya - yf) < 14) { if (ya <= yf) { ya -= 7; yf += 7; } else { ya += 7; yf -= 7; } }
  for (const [txt, yy] of [[fmt.format(ab[ab.length - 1]) + " ab.", ya], [fmt.format(fi[fi.length - 1]) + " fin.", yf]]) {
    const t = svgEl("text", { x: x(meses.length - 1) + 8, y: yy + 4 }); t.textContent = txt; t.style.fill = "var(--ink-2)"; svg.appendChild(t);
  }
  // camada de hover: linha vertical que encosta no mes mais proximo
  const cruz = svgEl("line", { y1: m.t, y2: H - m.b, stroke: "var(--axis)", "stroke-width": 1, visibility: "hidden" });
  svg.appendChild(cruz);
  const pontos = series.map(s => { const c = svgEl("circle", { r: 4, fill: s.cor, stroke: "var(--surface)", "stroke-width": 2, visibility: "hidden" }); svg.appendChild(c); return c; });
  const area = svgEl("rect", { x: m.e - 10, y: 0, width: W - m.e - m.d + 20, height: H, fill: "transparent" });
  svg.appendChild(area);
  area.addEventListener("pointermove", ev => {
    const r = svg.getBoundingClientRect(), px = (ev.clientX - r.left) * W / r.width;
    let j = 0, melhor = Infinity;
    meses.forEach((_, k) => { const d = Math.abs(x(k) - px); if (d < melhor) { melhor = d; j = k; } });
    cruz.setAttribute("x1", x(j)); cruz.setAttribute("x2", x(j)); cruz.setAttribute("visibility", "visible");
    series.forEach((s, k) => { pontos[k].setAttribute("cx", x(j)); pontos[k].setAttribute("cy", y(s.v[j])); pontos[k].setAttribute("visibility", "visible"); });
    const parcial = ultimoParcial && j === meses.length - 1 ? " (mês em andamento)" : "";
    mostrarDica(ev, [{ v: MESES[meses[j][1] - 1] + "/" + meses[j][0] + parcial },
      ...series.map(s => ({ v: fmt.format(s.v[j]), l: s.nome, cor: s.cor }))]);
  });
  area.addEventListener("pointerleave", () => { esconderDica(); cruz.setAttribute("visibility", "hidden"); pontos.forEach(p => p.setAttribute("visibility", "hidden")); });
  box.appendChild(svg);
}

const FAIXAS = [["até 1 dia", 24], ["1 a 2 dias", 48], ["2 a 5 dias", 120], ["5 a 10 dias", 240],
                ["10 a 20 dias", 480], ["mais de 20 dias", Infinity]];
function desenharSla(finalizadas) {
  const { corpo, tabela: tab } = card("c-sla", "SLA das tratativas", "finalizadas · " + nomePeriodo());
  const hs = finalizadas.map(l => l.h).filter(h => h != null);
  const cont = FAIXAS.map(() => 0);
  for (const h of hs) cont[FAIXAS.findIndex(([, lim]) => h <= lim)]++;
  if (tab) return hs.length ? tabela(corpo, ["Tempo de tratativa", "Finalizadas", "%"],
    FAIXAS.map(([n], j) => [n, fmt.format(cont[j]), fmt1.format(100 * cont[j] / hs.length) + "%"]), [1, 2]) : vazio(corpo);
  if (!hs.length) return vazio(corpo);
  const box = document.createElement("div"); box.className = "svgbox"; corpo.appendChild(box);
  const W = Math.max(box.clientWidth || corpo.clientWidth || 400, 280), H = 230;
  const m = { e: 40, d: 8, t: 18, b: 34 };
  const { passo, topo } = escalaY(Math.max(...cont, 1));
  const banda = (W - m.e - m.d) / FAIXAS.length, larg = Math.min(24, banda * 0.6);
  const y = v => m.t + (H - m.t - m.b) * (1 - v / topo);
  const svg = svgEl("svg", { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: "img",
                             "aria-label": "Colunas com quantas finalizadas em cada faixa de tempo" });
  for (let v = 0; v <= topo + 1e-9; v += passo) {
    svg.appendChild(svgEl("line", { x1: m.e, x2: W - m.d, y1: y(v), y2: y(v), stroke: v ? "var(--grid)" : "var(--axis)", "stroke-width": 1 }));
    const t = svgEl("text", { x: m.e - 8, y: y(v) + 4, "text-anchor": "end" }); t.textContent = fmt.format(v); svg.appendChild(t);
  }
  FAIXAS.forEach(([nome], j) => {
    const cx = m.e + banda * j + banda / 2, v = cont[j], hgt = (H - m.t - m.b) * v / topo;
    if (v) {
      const r = Math.min(4, hgt);
      const x0 = cx - larg / 2, x1 = cx + larg / 2, yb = y(0), yt = y(v);
      svg.appendChild(svgEl("path", { fill: "var(--s1)",
        d: `M${x0} ${yb} L${x0} ${yt + r} Q${x0} ${yt} ${x0 + r} ${yt} L${x1 - r} ${yt} Q${x1} ${yt} ${x1} ${yt + r} L${x1} ${yb} Z` }));
      const t = svgEl("text", { x: cx, y: yt - 5, "text-anchor": "middle" }); t.textContent = fmt.format(v); t.style.fill = "var(--ink-2)"; svg.appendChild(t);
    }
    const rot = svgEl("text", { x: cx, y: H - 12, "text-anchor": "middle" }); rot.textContent = nome; svg.appendChild(rot);
    const alvo = svgEl("rect", { x: cx - banda / 2, y: m.t, width: banda, height: H - m.t - m.b, fill: "transparent" });
    alvo.addEventListener("pointermove", ev => mostrarDica(ev, [{ v: fmt.format(v) + " finalizada(s)", l: nome },
      { v: fmt1.format(100 * v / hs.length) + "%", l: "das finalizadas" }]));
    alvo.addEventListener("pointerleave", esconderDica);
    svg.appendChild(alvo);
  });
  box.appendChild(svg);
}

// ------------------------------------------------------------------ inicio ----
el("sub").textContent = "Atualizado em " + DADOS.gerado + " · " + DADOS.origem + " · " + fmt.format(LINHAS.length) + " solicitações";
el("tema").addEventListener("click", () => {
  const r = document.documentElement;
  const escuro = r.dataset.theme ? r.dataset.theme === "dark" : matchMedia("(prefers-color-scheme: dark)").matches;
  r.dataset.theme = escuro ? "light" : "dark";
  try { localStorage.setItem("dashboard-fluig-tema", r.dataset.theme); } catch (e) {}
});
try { const t = localStorage.getItem("dashboard-fluig-tema"); if (t) document.documentElement.dataset.theme = t; } catch (e) {}
montarFiltros();
desenhar();
let espera;
addEventListener("resize", () => { clearTimeout(espera); espera = setTimeout(desenhar, 150); });
</script>
</body>
</html>
"""
