// Экран щита: строка показателей (KV — список функций), вкладки (TABS).
import { esc, num, refLink } from "../fmt.js";
import { S, counts } from "../store.js";
import { toPanel } from "../nav.js";
import { mount } from "../grid.js";
import { lineColumns, lineRef, lineAttrs } from "../columns.js";
import * as issues from "./issues.js";
import * as structure from "./structure.js";
import * as linkdlg from "./linkdlg.js";
import * as feederdlg from "./feederdlg.js";

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
    const src = p.link_src ? ` <span class="tag" title="${p.link_src === "программа" ? "связь хранит программа" : "связь записана в книге"}">${esc(p.link_src)}</span>` : "";
    return kv("Питание", (par ? `<a data-go="${esc(toPanel(par.id, p.parent_line, "lines"))}">${esc(par.name)}</a> / ${esc(p.parent_line || "?")}${src}`
                              : esc(p.supply.manual || p.supply.from_doc || "не задано"))
      + ` <button class="mini" data-linkdlg title="Задать, от какого щита и по какой линии питается щит">Связь…</button>`); },
  p => { const par = S.panels[p.parent], ln = par && par.lines.find(l => l.number === p.parent_line);
    const txt = ln ? [ln.scheme.breaker, (ln.scheme.cable || "").replace(/\s*L=.*$/, "")].filter(Boolean).join(", ")
                   : [p.supply.breaker, p.supply.cable].filter(Boolean).join(", ");
    return kv("Фидер", esc(txt) + (ln ? ` <button class="mini" data-feederdlg title="Параметры фидера (строка ${esc(ln.number)} щита ${esc(par.name)})">✎</button>` : "")); },
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

export const TABS = [
  { id: "lines", title: p => `Линии (${p.lines.length})`,
    bar: () => `<input type="search" id="lq" placeholder="Поиск: линия, QF, автомат, кабель (Ctrl+F)" value="${esc(lf.q)}">
      <label class="chk"><input type="checkbox" id="lonly" ${lf.only ? "checked" : ""}>только с замечаниями</label>`,
    mount: (host, p, prm) => {
      const q = lf.q;
      const rows = p.lines.filter(l => (!lf.only || l.result.level) &&
        (!q || `${l.number} ${l.name} ${l.scheme.qf} ${l.scheme.breaker} ${l.scheme.cable} ${l.result.status}`.toLowerCase().includes(q)));
      mount(host, {key: "lines", title: p.name, columns: lineColumns(), rows, freeze: 3, rowRef: lineRef, rowAttrs: lineAttrs,
        rowId: l => l.panel + "#" + l.number, focusRowId: prm.line ? p.id + "#" + prm.line : ""});
    } },
  { id: "structure", title: p => `Структура${p.nodes && p.nodes.length ? ` (${p.nodes.length} узл.)` : ""}`, bar: structure.bar0,
    html: p => structure.render(p), mount: (host, p) => structure.bind(host.closest("main"), p) },
  { id: "issues", title: p => `Замечания (${(S.issuesBy[p.id] || []).length})`, bar: issues.filters,
    mount: (host, p) => issues.mountTable(host, S.issuesBy[p.id] || [], false) },
];

export function render({ id, params }){
  const p = S.panels[id];
  if (!p) return `<div class="empty">Щит не найден (объект перечитан?).</div>`;
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  return `<div class="top"><h2>${esc(p.name)}</h2><span class="sub">${esc(p.doc)}</span>
      ${refLink({file: p.doc, sheet: "Однолинейка", cell: "A1"}, "открыть в Excel")}
      <a data-exclude="${esc(p.doc)}" title="Резервная или лишняя книга: программа перестанет её читать (список — _окно\\исключить.txt)">не читать эту книгу</a></div>
    <div class="kv">${KV.map(f => f(p)).join("")}</div>
    <div class="bar">${TABS.map(t => `<button data-go="${esc(toPanel(p.id, "", t.id))}" class="${t === tab ? "on" : ""}">${esc(t.title(p))}</button>`).join("")}
      <span class="sep"></span>${tab.bar(p)}</div>
    <div class="ghost" id="ghost">${tab.html ? tab.html(p) : ""}</div>`;
}

export function bind(root, rerender, { id, params }){
  issues.bind(root, rerender);
  const q = root.querySelector("#lq");
  if (q) q.oninput = () => { lf.q = q.value.toLowerCase(); const pos = q.selectionStart; rerender();
    const q2 = document.querySelector("#lq"); q2.focus(); q2.setSelectionRange(pos, pos); };
  const o = root.querySelector("#lonly");
  if (o) o.onchange = () => { lf.only = o.checked; rerender(); };
  const p = S.panels[id];
  if (!p) return;
  root.querySelectorAll("[data-linkdlg]").forEach(b => b.onclick = () => linkdlg.open(p));
  root.querySelectorAll("[data-feederdlg]").forEach(b => b.onclick = () => feederdlg.open(p));
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  tab.mount(root.querySelector("#ghost"), p, params);
}
