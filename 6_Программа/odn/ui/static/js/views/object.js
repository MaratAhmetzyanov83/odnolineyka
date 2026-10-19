// Экран «Весь объект». Вкладки — TABS: новая вкладка = новый элемент {id, title, bar, render, mount}.
import { esc, num } from "../fmt.js";
import { runKJ } from "./panel.js";
import { S, counts, roots, drawState, kjState, specState, kc } from "../store.js";
import { toObject, toPanel } from "../nav.js";
import { mount } from "../grid.js";
import { lineColumns, lineRef, lineAttrs } from "../columns.js";
import * as issues from "./issues.js";
import * as scheme from "./scheme.js";

const kv = (k, v, cls = "") => `<div><span class="k">${k}</span><b class="${cls}">${v}</b></div>`;
const ec = p => counts(S.issuesBy[p.id]);

// сводка: тип замечания × щит; щелчок по числу — замечания щита этого типа
function summary(){
  const ps = Object.values(S.panels).sort((a, b) => a.name.localeCompare(b.name, "ru", {numeric: true}));
  const m = {};
  for (const i of S.P.issues) for (const k of issues.kindsOf(i)){
    const r = m[k] ||= {lvl: 0, n: {}, tot: 0};
    r.lvl = Math.max(r.lvl, i.level); r.n[i.panel] = (r.n[i.panel] || 0) + 1; r.tot++;
  }
  const kinds = Object.entries(m).sort((a, b) => b[1].lvl - a[1].lvl || b[1].tot - a[1].tot);
  if (!kinds.length) return `<div class="empty">Замечаний нет.</div>`;
  let h = `<table class="mx"><thead><tr><th></th><th>Тип замечания</th><th class=n>Всего</th>${ps.map(p => `<th class=n title="${esc(p.doc)}">${esc(p.name)}</th>`).join("")}</tr></thead><tbody>`;
  for (const [k, r] of kinds){
    h += `<tr><td><span class="dot d${r.lvl}"></span></td><td class="w" title="${esc(k)}">${esc(k)}</td><td class=n><b>${r.tot}</b></td>` +
      ps.map(p => { const n = r.n[p.id]; return n ? `<td class="n hit${r.lvl}" data-kindgo="${esc(k)}" data-go="${esc(toPanel(p.id, "", "issues"))}">${n}</td>` : `<td class="n c0">·</td>`; }).join("") + "</tr>";
  }
  return `<div class="tbl">${h}</tbody></table></div>`;
}

function panelsGrid(host){
  const cols = [
    {id: "name", t: "Щит", w: 90, v: p => p.name, h: p => `<a data-go="${esc(toPanel(p.id))}"><b>${esc(p.name)}</b></a>`, agg: "count"},
    {id: "from", t: "Питается от", w: 120, v: p => S.panels[p.parent] ? `${S.panels[p.parent].name} / ${p.parent_line}` : p.supply.manual},
    {id: "brk", t: "Фидер", w: 100, v: p => p.supply.breaker},
    {id: "cab", t: "Кабель фидера", w: 150, v: p => p.supply.cable},
    {id: "len", t: "L, м", n: 1, w: 50, d: 0, v: p => p.supply.length, agg: "sum"},
    {id: "pi", t: "Руст", title: "Руст, кВт", n: 1, w: 60, d: 1, v: p => p.result.p_inst},
    {id: "pc", t: "Рр", title: "Рр, кВт", n: 1, w: 60, d: 1, v: p => p.result.p_calc},
    {id: "kc", t: "Кс", title: "Коэффициент спроса щита: Рр / Руст", n: 1, w: 44, d: 2, v: p => kc(p)},
    {id: "ic", t: "Iр, А", n: 1, w: 56, d: 1, v: p => p.result.i_calc, agg: "max"},
    ...["L1", "L2", "L3"].map(k => ({id: k, t: k + ", А", title: `ΣIр фазы ${k} с нижестоящими щитами, без Кс`, n: 1, w: 50, d: 1, hidden: 1,
      v: p => (p.phase_i || {})[k]})),
    {id: "skew", t: "Перекос", title: "Перекос фаз с нижестоящими щитами, % = (max − min) / max", n: 1, w: 60, d: 0,
      v: p => p.phase_skew, agg: "max", cls: p => p.phase_skew > 15 ? "cw" : "",
      tip: p => ["L1", "L2", "L3"].map(k => `${k} ${num((p.phase_i || {})[k] || 0, 1)} А`).join(" · ")},
    {id: "cos", t: "cos φ", n: 1, w: 50, d: 2, v: p => p.result.cos},
    {id: "du", t: "ΔU шин, %", n: 1, w: 66, v: p => p.result.du_bus, agg: "max"},
    {id: "kz1", t: "Iкз min, А", n: 1, w: 70, d: 0, v: p => p.result.ikz_min_bus, agg: "min"},
    {id: "kz3", t: "Iкз max, кА", n: 1, w: 74, v: p => p.result.ikz_max_bus, agg: "max"},
    {id: "inp", t: "Вводы / АВР", w: 80, v: p => `${p.scheme.inputs ?? ""} ${p.scheme.avr || ""}`.trim()},
    {id: "cat", t: "Кат.", title: "Категория, обеспечиваемая вводами", w: 40, v: p => p.result.category},
    {id: "vru", t: "ВРУ", title: "ВРУ здания: потери считаются от него", w: 44, v: p => p.vru ? "ВРУ" : "", cls: () => "vru"},
    {id: "man", t: "Питание вручную", title: "«Сеть_щита» C12 (если нет связи)", w: 110, hidden: 1, v: p => p.supply.manual, edit: {path: "supply.manual"}},
    {id: "nl", t: "Линий", n: 1, w: 50, d: 0, v: p => p.lines.length, agg: "sum"},
    {id: "e", t: "Ош.", n: 1, w: 40, d: 0, v: p => ec(p)[0], agg: "sum", cls: p => ec(p)[0] ? "ce" : "c0"},
    {id: "w", t: "Вн.", n: 1, w: 40, d: 0, v: p => ec(p)[1], agg: "sum", cls: p => ec(p)[1] ? "cw" : "c0"},
    {id: "dwg", t: "Однолинейка", title: "DWG: актуальна / устарела / проверить (отметки SXDRAW v3.26+). Подсказка — по чертежам", w: 90,
     v: p => drawState(p).txt, cls: p => drawState(p).cls, tip: p => drawState(p).tip},
    {id: "kj", t: "КЖ", title: "Кабельный журнал в DWG: актуален / устарел (отметка «КЖ в DWG»). Подсказка — что изменилось", w: 80,
     v: p => kjState(p).txt, cls: p => kjState(p).cls, tip: p => kjState(p).tip},
    {id: "spec", t: "Спец.", title: "Спецификация книги: актуальна / устарела / проверить (аппараты менялись после её проверки). Подсказка — что изменилось", w: 80,
     v: p => specState(p).txt, cls: p => specState(p).cls, tip: p => specState(p).tip},
    {id: "rev", t: "Отпечаток", w: 100, hidden: 1, v: p => p.revision},
    {id: "doc", t: "Файл", w: 200, v: p => p.doc, cls: () => "mut"},
  ];
  mount(host, {key: "panels", title: "Щиты", columns: cols, rows: Object.values(S.panels), freeze: 1,
    rowRef: p => ({panel: p.id, line: "", name: p.name}),
    rowAttrs: p => { const [e, w] = ec(p); return `class="${e ? "l2" : w ? "l1" : ""}"`; }});
}

function allLinesGrid(host){
  const rows = Object.values(S.panels).flatMap(p => p.lines);
  mount(host, {key: "linesAll", title: "Все линии", columns: lineColumns({withPanel: true}), rows, freeze: 3,
    rowRef: lineRef, rowAttrs: lineAttrs, rowId: l => l.panel + "#" + l.number});
}

// книги с устаревшим КЖ — обновить все разом (DWG из отметок .kj.txt)
const kjStale = () => S.P.documents.filter(d => d.kj?.state === "stale");
const kjBar = () => { const n = kjStale().length;
  return n ? `<button id="kjall" title="Перестроить устаревшие КЖ в их DWG: надстройка «КЖ в DWG» в своём Excel, AutoCAD в фоне">Обновить устаревшие КЖ (${n})</button>` : ""; };

export const TABS = [
  { id: "summary", title: () => "Сводка", bar: () => `<span class="sub">Тип замечания × щит. Щелчок по числу — эти замечания в щите.</span>`,
    render: summary },
  { id: "issues", title: () => `Замечания (${S.P.issues.length})`, bar: issues.filters, mount: host => issues.mountTable(host, S.P.issues, true) },
  { id: "scheme", title: () => "Схема", bar: () => "", mount: scheme.mount },
  { id: "panels", title: () => `Щиты (${Object.keys(S.panels).length})`, bar: kjBar, mount: panelsGrid },
  { id: "lines", title: () => `Все линии (${Object.values(S.panels).reduce((s, p) => s + p.lines.length, 0)})`, bar: () => "", mount: allLinesGrid },
];

export function render({ params }){
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  const [e, w] = counts(S.P.issues), ps = Object.values(S.panels);
  const main = roots().sort((a, b) => (S.kids[b.id]?.length || 0) - (S.kids[a.id]?.length || 0))[0];
  return `<div class="top"><h2>Весь объект</h2><span class="sub">${esc(S.P.folder)}</span></div>
    <div class="kv">${kv("Ошибок", e, e ? "e" : "")}${kv("Внимание", w, w ? "w" : "")}${kv("Щитов", ps.length)}
      ${kv("Линий", ps.reduce((s, p) => s + p.lines.length, 0))}
      ${main ? kv("Главный щит", `<a data-go="${esc(toPanel(main.id))}">${esc(main.name)}</a>`) + kv("Руст", num(main.result.p_inst, 1) + " кВт")
        + kv("Рр", num(main.result.p_calc, 1) + " кВт") + kv("Iр", num(main.result.i_calc, 1) + " А") : ""}
      ${kv("Документов", S.P.documents.length)}</div>
    <div class="bar">${TABS.map(t => `<button data-go="${toObject(t.id)}" class="${t === tab ? "on" : ""}">${esc(t.title())}</button>`).join("")}
      <span class="sep"></span>${tab.bar()}</div>
    ${tab.mount ? `<div class="ghost" id="ghost"></div>` : tab.render()}`;
}

export function bind(root, rerender, ctx){
  issues.bind(root, rerender);
  root.querySelectorAll("[data-kindgo]").forEach(td => td.addEventListener("click", () => {
    issues.filt.kind = td.dataset.kindgo; issues.filt.lvl = 0; issues.filt.group = ""; issues.filt.q = ""; }, true));
  const ka = root.querySelector("#kjall");
  if (ka) ka.onclick = () => { const ds = kjStale();
    if (confirm(`Обновить КЖ в ${ds.length} книгах?\n${ds.map(d => d.kj.dwg).join("\n")}\nНужен AutoCAD.`)) runKJ(ds.map(d => ({doc: d.file, dwg: ""})), ka); };
  const tab = TABS.find(t => t.id === ctx.params.tab) || TABS[0];
  if (tab.mount) tab.mount(root.querySelector("#ghost"));
}
