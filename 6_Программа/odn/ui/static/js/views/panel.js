// Экран щита: строка показателей (KV — список функций), вкладки (TABS).
import { esc, num, refLink } from "../fmt.js";
import { S, counts, lineIn } from "../store.js";
import { toPanel } from "../nav.js";
import { grid } from "../grid.js";
import * as issues from "./issues.js";

const kv = (k, v, cls = "", tip = "") => v === "" || v == null ? "" :
  `<div ${tip ? `title="${esc(tip)}"` : ""}><span class="k">${k}</span><b class="${cls}">${v}</b></div>`;
const lf = { q: "", only: false };              // фильтр линий

// Σ Iр по фазам (без Кс): однофазные — на свою фазу, трёхфазные — на все
function phases(p){
  const s = {L1: 0, L2: 0, L3: 0};
  for (const l of p.lines){
    const i = l.result.i_calc || 0, ph = (l.scheme.phase || "").split(/[,\s]+/).filter(x => s[x] != null);
    if (ph.length >= 3) for (const k in s) s[k] += i; else for (const k of ph) s[k] += i;
  }
  const v = Object.values(s), mx = Math.max(...v), mn = Math.min(...v);
  return {s, skew: mx > 0 ? (mx - mn) / mx * 100 : 0};
}

export const KV = [
  p => { const par = S.panels[p.parent];
    return kv("Питание", par ? `<a data-go="${esc(toPanel(par.id, p.parent_line, "lines"))}">${esc(par.name)}</a> / ${esc(p.parent_line || "?")}`
                             : esc(p.supply.manual || p.supply.from_doc || "не задано")); },
  p => kv("Фидер", esc([p.supply.breaker, p.supply.cable].filter(Boolean).join(", "))),
  p => kv("L", p.supply.length != null ? num(p.supply.length, 0) + " м" : ""),
  p => kv("Руст", num(p.result.p_inst, 2) + " кВт"),
  p => kv("Рр", num(p.result.p_calc, 2) + " кВт"),
  p => kv("Iр", num(p.result.i_calc, 1) + " А"),
  p => kv("cos φ", num(p.result.cos, 2)),
  p => kv("ΔU до шин", p.result.du_bus != null ? num(p.result.du_bus, 2) + " %" : ""),
  p => kv("Iкз min", p.result.ikz_min_bus != null ? num(p.result.ikz_min_bus, 0) + " А" : ""),
  p => kv("Iкз max", p.result.ikz_max_bus != null ? num(p.result.ikz_max_bus, 2) + " кА" : ""),
  p => kv("Вводы", esc([p.scheme.inputs, p.scheme.avr && p.scheme.avr !== "нет" ? "АВР" : ""].filter(Boolean).join(" + "))),
  p => kv("Кат.", esc(p.result.category)),
  p => { const {s, skew} = phases(p);
    return kv("Фазы ΣIр", Object.entries(s).map(([k, v]) => `${k} ${num(v, 1)}`).join(" · ") + (skew > 15 ? ` (перекос ${num(skew, 0)} %)` : ""),
      skew > 15 ? "w" : "", "Сумма Iр линий по фазам, без Кс; перекос = (max − min) / max"); },
  p => { const [e, w] = counts(S.issuesBy[p.id]); return kv("Ошибок / вн.", `<span class="${e ? "e" : ""}">${e}</span> / <span class="${w ? "w" : ""}">${w}</span>`); },
  p => kv("Отпечаток", esc(p.revision), "", "Отпечаток данных схемы. Позже — сверка с отметкой в DWG / КЖ"),
];

const cableShort = c => (c || "").replace(/\s*L=.*$/, "");

function linesTable(p, hl){
  const q = lf.q;
  const rows = p.lines.filter(l => (!lf.only || l.result.level) &&
    (!q || `${l.number} ${l.name} ${l.scheme.qf} ${l.scheme.breaker} ${l.scheme.cable} ${l.result.status}`.toLowerCase().includes(q)));
  const cols = [
    {id: "pl", t: "М.", title: "Место", n: 1, v: l => l.scheme.place},
    {id: "qf", t: "QF", v: l => l.scheme.qf},
    {id: "num", t: "Линия", v: l => l.number},
    {id: "name", t: "Наименование", v: l => l.name, cls: () => "w",
      h: l => { const c = l.feeds && S.panels[l.feeds];
        return esc(l.name) + (c ? ` → <a data-go="${esc(toPanel(c.id))}">${esc(c.name)}</a>` : ""); }},
    {id: "p", t: "P, кВт", n: 1, v: l => l.input.p_kw, h: l => num(l.input.p_kw, 2)},
    {id: "ph", t: "Фаза", v: l => l.scheme.phase},
    {id: "ic", t: "Iр, А", n: 1, v: l => l.result.i_calc, h: l => num(l.result.i_calc, 2)},
    {id: "in", t: "In, А", n: 1, v: l => lineIn(l), h: l => num(lineIn(l), 0),
      cls: l => lineIn(l) != null && l.result.i_calc != null && lineIn(l) < l.result.i_calc ? "bad" : ""},
    {id: "brk", t: "Автомат", v: l => l.scheme.breaker},
    {id: "cab", t: "Кабель", v: l => cableShort(l.scheme.cable), tip: l => l.scheme.cable},
    {id: "s", t: "S, мм²", n: 1, v: l => l.result.s_mm2, h: l => num(l.result.s_mm2, 1)},
    {id: "len", t: "L, м", n: 1, v: l => l.input.length, h: l => num(l.input.length, 0)},
    {id: "du", t: "ΔU, %", n: 1, v: l => l.result.du, h: l => num(l.result.du, 2)},
    {id: "dus", t: "ΔUΣ, %", title: "ΔU от ВРУ / допустимая", n: 1, v: l => l.result.du_sum, h: l => num(l.result.du_sum, 2),
      cls: l => l.result.du_sum != null && l.result.du_max != null && l.result.du_sum > l.result.du_max ? "bad" : "",
      tip: l => l.result.du_max != null ? `допустимо ${num(l.result.du_max, 1)} %` : ""},
    {id: "kz", t: "Iкз, А", title: "Iкз min в конце линии", n: 1, v: l => l.result.ikz_min, h: l => num(l.result.ikz_min, 0)},
    {id: "kzr", t: "Iкз/In", n: 1, v: l => l.result.ikz_ratio, h: l => num(l.result.ikz_ratio, 1),
      cls: l => (l.result.ikz_check || "").startsWith("мало") ? "bad" : "", tip: l => l.result.ikz_check},
    {id: "cat", t: "Кат.", v: l => l.result.category},
    {id: "st", t: "Статус", v: l => l.result.status, cls: l => "w " + (l.result.level >= 2 ? "lv2" : l.result.level ? "lv1" : ""),
      h: l => `${refLink(l.refs.scheme || l.refs.input, "✎")} ${esc(l.result.status)}`, tip: l => l.result.status},
  ];
  return grid("lines", cols, rows, {empty: "Нет линий по фильтру.",
    rowAttrs: l => `class="l${l.result.level} ${hl === l.number ? "hl" : ""}" data-line="${esc(l.number)}"`});
}

export const TABS = [
  { id: "lines", title: p => `Линии (${p.lines.length})`,
    bar: () => `<input type="search" id="lq" placeholder="Поиск: линия, QF, автомат, кабель" value="${esc(lf.q)}">
      <label class="chk"><input type="checkbox" id="lonly" ${lf.only ? "checked" : ""}>только с замечаниями</label>`,
    render: (p, prm) => linesTable(p, prm.line) },
  { id: "issues", title: p => `Замечания (${(S.issuesBy[p.id] || []).length})`, bar: issues.filters,
    render: p => issues.table(S.issuesBy[p.id] || [], false) },
];

export function render({ id, params }){
  const p = S.panels[id];
  if (!p) return `<div class="empty">Щит не найден (объект перечитан?).</div>`;
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  return `<div class="top"><h2>${esc(p.name)}</h2><span class="sub">${esc(p.doc)}</span>
      ${refLink({file: p.doc, sheet: "Однолинейка", cell: "A1"}, "открыть в Excel")}</div>
    <div class="kv">${KV.map(f => f(p)).join("")}</div>
    <div class="bar">${TABS.map(t => `<button data-go="${esc(toPanel(p.id, "", t.id))}" class="${t === tab ? "on" : ""}">${esc(t.title(p))}</button>`).join("")}
      <span class="sep"></span>${tab.bar(p)}</div>
    <div class="tbl">${tab.render(p, params)}</div>`;
}

export function bind(root, rerender, { params }){
  issues.bind(root, rerender);
  const q = root.querySelector("#lq");
  if (q) q.oninput = () => { lf.q = q.value.toLowerCase(); const pos = q.selectionStart; rerender();
    const q2 = document.querySelector("#lq"); q2.focus(); q2.setSelectionRange(pos, pos); };
  const o = root.querySelector("#lonly");
  if (o) o.onchange = () => { lf.only = o.checked; rerender(); };
  if (params.line) root.querySelector(`tr[data-line="${CSS.escape(params.line)}"]`)?.scrollIntoView({block: "center"});
}
