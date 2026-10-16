// Экран «Весь объект». Вкладки — список TABS: новая вкладка = новый элемент.
import { esc, num } from "../fmt.js";
import { S, counts } from "../store.js";
import { toObject, toPanel } from "../nav.js";
import * as issues from "./issues.js";

function panelsTable(){
  let h = `<table><tr><th>Щит</th><th>Питается от</th><th class=n>Руст, кВт</th><th class=n>Рр, кВт</th><th class=n>Iр, А</th>
    <th>Фидер</th><th class=n>ΔU шин, %</th><th class=n>Iкз min, А</th><th class=n>Ошибок</th><th class=n>Внимание</th><th>Файл</th></tr>`;
  for (const p of Object.values(S.panels)){
    const [e, w] = counts(S.issuesBy[p.id]), par = S.panels[p.parent], r = p.result;
    h += `<tr class="${e ? "l2" : w ? "l1" : ""}" data-go="${esc(toPanel(p.id))}"><td><b>${esc(p.name)}</b></td>
      <td>${par ? esc(par.name) + " / " + esc(p.parent_line) : esc(p.supply.manual)}</td>
      <td class=n>${num(r.p_inst, 2)}</td><td class=n>${num(r.p_calc, 2)}</td><td class=n>${num(r.i_calc, 1)}</td>
      <td>${esc(p.supply.breaker)} ${esc(p.supply.cable)}</td><td class=n>${num(r.du_bus, 2)}</td><td class=n>${num(r.ikz_min_bus, 0)}</td>
      <td class=n>${e}</td><td class=n>${w}</td><td class=sub>${esc(p.doc)}</td></tr>`;
  }
  return h + "</table>";
}

export const TABS = [
  { id: "issues", title: () => `Замечания (${S.P.issues.length})`, render: () => issues.filters() + issues.table(S.P.issues, true) },
  { id: "panels", title: () => `Щиты (${Object.keys(S.panels).length})`, render: panelsTable },
];

export function render({ params }){
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  const [e, w] = counts(S.P.issues), ps = Object.values(S.panels);
  return `<h2>Весь объект</h2><div class="sub">${esc(S.P.folder)}</div>
    <div class="cards">
      <div class="card ${e ? "e" : ""}"><div class="k">Ошибок</div><div class="v">${e}</div></div>
      <div class="card ${w ? "w" : ""}"><div class="k">Внимание</div><div class="v">${w}</div></div>
      <div class="card"><div class="k">Щитов</div><div class="v">${ps.length}</div></div>
      <div class="card"><div class="k">Линий</div><div class="v">${ps.reduce((s, p) => s + p.lines.length, 0)}</div></div>
    </div>
    <div class="tabs">${TABS.map(t => `<button data-go="${toObject(t.id)}" class="${t === tab ? "on" : ""}">${esc(t.title())}</button>`).join("")}</div>
    ${tab.render()}`;
}
export const bind = issues.bind;
