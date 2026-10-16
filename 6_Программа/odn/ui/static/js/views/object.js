// Экран «Весь объект». Вкладки — TABS: новая вкладка = новый элемент.
import { esc, num } from "../fmt.js";
import { S, counts, roots } from "../store.js";
import { toObject, toPanel } from "../nav.js";
import { grid } from "../grid.js";
import * as issues from "./issues.js";

const kv = (k, v, cls = "") => `<div><span class="k">${k}</span><b class="${cls}">${v}</b></div>`;

function panelsTable(){
  const cols = [
    {id: "name", t: "Щит", v: p => p.name, h: p => `<b>${esc(p.name)}</b>`},
    {id: "from", t: "Питается от", v: p => S.panels[p.parent] ? `${S.panels[p.parent].name} / ${p.parent_line}` : p.supply.manual},
    {id: "brk", t: "Фидер", v: p => p.supply.breaker},
    {id: "cab", t: "Кабель фидера", v: p => p.supply.cable},
    {id: "len", t: "L, м", n: 1, v: p => p.supply.length, h: p => num(p.supply.length, 0)},
    {id: "pi", t: "Руст", title: "Руст, кВт", n: 1, v: p => p.result.p_inst, h: p => num(p.result.p_inst, 1)},
    {id: "pc", t: "Рр", title: "Рр, кВт", n: 1, v: p => p.result.p_calc, h: p => num(p.result.p_calc, 1)},
    {id: "ic", t: "Iр, А", n: 1, v: p => p.result.i_calc, h: p => num(p.result.i_calc, 1)},
    {id: "cos", t: "cos φ", n: 1, v: p => p.result.cos, h: p => num(p.result.cos, 2)},
    {id: "du", t: "ΔU шин, %", n: 1, v: p => p.result.du_bus, h: p => num(p.result.du_bus, 2)},
    {id: "kz1", t: "Iкз min, А", n: 1, v: p => p.result.ikz_min_bus, h: p => num(p.result.ikz_min_bus, 0)},
    {id: "kz3", t: "Iкз max, кА", n: 1, v: p => p.result.ikz_max_bus, h: p => num(p.result.ikz_max_bus, 2)},
    {id: "inp", t: "Вводы / АВР", v: p => `${p.scheme.inputs ?? ""} ${p.scheme.avr || ""}`.trim()},
    {id: "cat", t: "Кат.", title: "Категория, обеспечиваемая вводами", v: p => p.result.category},
    {id: "nl", t: "Линий", n: 1, v: p => p.lines.length},
    {id: "e", t: "Ош.", n: 1, v: p => counts(S.issuesBy[p.id])[0], h: p => { const e = counts(S.issuesBy[p.id])[0]; return `<span class="${e ? "ce" : "c0"}">${e}</span>`; }},
    {id: "w", t: "Вн.", n: 1, v: p => counts(S.issuesBy[p.id])[1], h: p => { const w = counts(S.issuesBy[p.id])[1]; return `<span class="${w ? "cw" : "c0"}">${w}</span>`; }},
    {id: "doc", t: "Файл", v: p => p.doc, cls: () => "mut"},
  ];
  return grid("panels", cols, Object.values(S.panels), {rowAttrs: p => {
    const [e, w] = counts(S.issuesBy[p.id]); return `class="${e ? "l2" : w ? "l1" : ""}" data-go="${esc(toPanel(p.id))}"`; }});
}

// сводка: тип замечания × щит; щелчок по числу — замечания щита этого типа
function summary(){
  const ps = Object.values(S.panels).sort((a, b) => a.name.localeCompare(b.name, "ru", {numeric: true}));
  const m = {};                                       // kind -> {lvl, n: {pid: count}, tot}
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
  return h + "</tbody></table>";
}

export const TABS = [
  { id: "summary", title: () => "Сводка", bar: () => `<span class="sub">Тип замечания × щит. Щелчок по числу — эти замечания в щите.</span>`, render: summary },
  { id: "issues", title: () => `Замечания (${S.P.issues.length})`, bar: issues.filters, render: () => issues.table(S.P.issues, true) },
  { id: "panels", title: () => `Щиты (${Object.keys(S.panels).length})`, bar: () => "", render: panelsTable },
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
    <div class="tbl">${tab.render()}</div>`;
}
export function bind(root, rerender){
  issues.bind(root, rerender);
  root.querySelectorAll("[data-kindgo]").forEach(td => td.addEventListener("click", () => {
    issues.filt.kind = td.dataset.kindgo; issues.filt.lvl = 0; issues.filt.group = ""; issues.filt.q = ""; }, true));
}
