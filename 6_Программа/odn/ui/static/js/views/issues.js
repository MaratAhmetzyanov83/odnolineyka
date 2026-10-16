// Таблица замечаний с фильтрами (общая для объекта и щита).
import { esc, refLink } from "../fmt.js";
import { S } from "../store.js";
import { toPanel } from "../nav.js";

export const filt = { lvl: 0, group: "", q: "" };

function groups(){ return [...new Set(Object.values(S.rules).map(r => r.group))]; }

export function filters(){
  const b = (attr, val, cur, text) => `<button data-${attr}="${esc(val)}" class="${cur === val ? "on" : ""}">${esc(text)}</button>`;
  return `<div class="filters">
    ${b("lvl", 0, filt.lvl, "Все")}${b("lvl", 2, filt.lvl, "Ошибки")}${b("lvl", 1, filt.lvl, "Внимание")}
    <span style="width:8px"></span>
    ${b("group", "", filt.group, "Все проверки")}${groups().map(g => b("group", g, filt.group, g)).join("")}
    <input type="search" id="q" placeholder="Поиск: линия, щит, текст" value="${esc(filt.q)}"></div>`;
}

export function table(list, withPanel){
  const L = list.filter(i => (!filt.lvl || i.level === filt.lvl) && (!filt.group || i.group === filt.group) &&
    (!filt.q || `${i.text} ${i.line} ${S.panels[i.panel]?.name || ""}`.toLowerCase().includes(filt.q)));
  if (!L.length) return `<div class="empty">Замечаний нет.</div>`;
  let h = `<table><tr><th></th>${withPanel ? "<th>Щит</th>" : ""}<th>Линия</th><th>Замечание</th><th>Проверка</th><th>Где править</th></tr>`;
  for (const i of L){
    const p = S.panels[i.panel];
    h += `<tr class="l${i.level}" ${p ? `data-go="${esc(toPanel(p.id, i.line, i.line ? "lines" : "issues"))}"` : ""}>
      <td>${i.level >= 2 ? "<span class=l2t>ошибка</span>" : "<span class=l1t>внимание</span>"}</td>
      ${withPanel ? `<td>${esc(p?.name || "")}</td>` : ""}<td>${esc(i.line)}</td><td>${esc(i.text)}</td>
      <td><span class="tag" title="${esc(i.rule)}">${esc(S.rules[i.rule]?.title || i.rule)}</span></td>
      <td>${refLink(i.ref)}</td></tr>`;
  }
  return h + "</table>";
}

export function bind(root, rerender){
  root.querySelectorAll("[data-lvl]").forEach(b => b.onclick = () => { filt.lvl = +b.dataset.lvl; rerender(); });
  root.querySelectorAll("[data-group]").forEach(b => b.onclick = () => { filt.group = b.dataset.group; rerender(); });
  const q = root.querySelector("#q");
  if (q) q.oninput = () => { filt.q = q.value.toLowerCase(); const pos = q.selectionStart; rerender();
    const q2 = document.querySelector("#q"); q2.focus(); q2.setSelectionRange(pos, pos); };
}
