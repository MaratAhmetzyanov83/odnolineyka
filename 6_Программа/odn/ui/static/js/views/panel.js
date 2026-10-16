// Экран щита. Карточки — список CARDS, вкладки — TABS: расширять добавлением элементов.
import { esc, num, refLink } from "../fmt.js";
import { S, counts, lineIn } from "../store.js";
import { toPanel } from "../nav.js";
import * as issues from "./issues.js";

const card = (k, v, cls = "", small = false) =>
  `<div class="card ${cls}"><div class="k">${k}</div><div class="v ${small ? "small" : ""}">${v}</div></div>`;

export const CARDS = [
  p => {
    const par = S.panels[p.parent], s = p.supply;
    const from = par ? `<a data-go="${esc(toPanel(par.id, p.parent_line, "lines"))}">${esc(par.name)}</a>, линия ${esc(p.parent_line || "?")}`
                     : esc(s.manual || s.from_doc || "не задано");
    const det = [s.breaker, s.cable, s.length != null ? `L=${num(s.length, 0)} м` : ""].filter(Boolean).join(" · ");
    return `<div class="card"><div class="k">Питается от</div><div class="v" style="font-size:14px">${from}</div><div class="k">${esc(det)}</div></div>`;
  },
  p => card("Руст / Рр, кВт", `${num(p.result.p_inst, 2)} / ${num(p.result.p_calc, 2)}`),
  p => card("Iр, А · cos φ", `${num(p.result.i_calc, 1)} · ${num(p.result.cos, 2)}`),
  p => card("ΔU до шин, %", num(p.result.du_bus, 2) || "—"),
  p => card("Iкз min / max на шинах",
    `${p.result.ikz_min_bus != null ? num(p.result.ikz_min_bus, 0) + " А" : "—"} / ${p.result.ikz_max_bus != null ? num(p.result.ikz_max_bus, 2) + " кА" : "—"}`),
  p => { const [e, w] = counts(S.issuesBy[p.id]); return card("Ошибок / внимание", `${e} / ${w}`, e ? "e" : w ? "w" : ""); },
  p => card("Схема DWG / КЖ", `отпечаток данных ${esc(p.revision)}; сверка с чертежом — следующий шаг`, "", true),
];

function linesTable(p, hl){
  let h = `<table><tr><th class=n>Место</th><th>QF</th><th>Линия</th><th>Наименование</th><th class=n>P, кВт</th><th>Фаза</th>
    <th class=n>Iр, А</th><th>Автомат</th><th>Кабель</th><th class=n>L, м</th><th class=n>ΔU Σ, %</th><th class=n>Iкз/In</th><th>Кат.</th><th>Статус</th></tr>`;
  for (const l of p.lines){
    const child = l.feeds && S.panels[l.feeds], r = l.result, sc = l.scheme;
    const duBad = r.du_sum != null && r.du_max != null && r.du_sum > r.du_max;
    h += `<tr class="l${r.level} ${hl === l.number ? "hl" : ""}" data-line="${esc(l.number)}">
      <td class=n>${sc.place ?? ""}</td><td>${esc(sc.qf)}</td><td>${esc(l.number)}</td>
      <td>${esc(l.name)}${child ? ` → <a data-go="${esc(toPanel(child.id))}">${esc(child.name)}</a>` : ""}</td>
      <td class=n>${num(l.input.p_kw, 2)}</td><td>${esc(sc.phase)}</td><td class=n>${num(r.i_calc, 2)}</td>
      <td class=nw title="In ${num(lineIn(l), 0)} А">${esc(sc.breaker)}</td><td class=sm>${esc(sc.cable)}</td>
      <td class=n>${num(l.input.length, 0)}</td><td class="n ${duBad ? "bad" : ""}">${num(r.du_sum, 2)}</td>
      <td class=n>${num(r.ikz_ratio, 1)}</td><td>${esc(r.category)}</td>
      <td class=sm>${esc(r.status)} ${refLink(l.refs.scheme || l.refs.input, "✎")}</td></tr>`;
  }
  return h + "</table>";
}

export const TABS = [
  { id: "lines", title: p => `Линии (${p.lines.length})`, render: (p, prm) => linesTable(p, prm.line) },
  { id: "issues", title: p => `Замечания (${(S.issuesBy[p.id] || []).length})`,
    render: p => issues.filters() + issues.table(S.issuesBy[p.id] || [], false) },
];

export function render({ id, params }){
  const p = S.panels[id];
  if (!p) return `<div class="empty">Щит не найден (объект перечитан?).</div>`;
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  return `<h2>${esc(p.name)}</h2>
    <div class="sub">${esc(p.doc)} · ${refLink({file: p.doc, sheet: "Однолинейка", cell: "A1"}, "открыть в Excel")}</div>
    <div class="cards">${CARDS.map(c => c(p)).join("")}</div>
    <div class="tabs">${TABS.map(t => `<button data-go="${esc(toPanel(p.id, "", t.id))}" class="${t === tab ? "on" : ""}">${esc(t.title(p))}</button>`).join("")}</div>
    ${tab.render(p, params)}`;
}
export function bind(root, rerender, { params }){
  issues.bind(root, rerender);
  if (params.line) root.querySelector(`tr[data-line="${CSS.escape(params.line)}"]`)?.scrollIntoView({block: "center"});
}
