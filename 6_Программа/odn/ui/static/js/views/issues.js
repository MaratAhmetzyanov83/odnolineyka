// Замечания: фильтры и таблица (общие для объекта и щита).
import { esc, refLink, toast } from "../fmt.js";
import * as api from "../api.js";
import { S } from "../store.js";
import { toPanel } from "../nav.js";
import { mount } from "../grid.js";

export const filt = { lvl: 0, group: "", q: "", kind: "", dis: false };

// тип замечания: статус строки — по частям без номеров («дубль клеммы XT…»), остальное — название правила
export function kindsOf(i){
  if (i.rule !== "source.row_status") return [S.rules[i.rule]?.title || i.rule];
  return i.text.split(/;\s*/).map(t => t.replace(/\bXT\d+/g, "XT…").replace(/\d+([.,]\d+)?(?=\s*(%|А|кА|кВт|м)(\s|$|[^\wа-яё]))/gi, "…").trim()).filter(Boolean);
}

export function filters(p){
  const b = (attr, val, cur, text) => `<button data-${attr}="${esc(val)}" class="${cur === val ? "on" : ""}">${esc(text)}</button>`;
  const groups = [...new Set(Object.values(S.rules).map(r => r.group))];
  const kind = filt.kind ? `<button data-kind="" class="on" title="Снять фильтр">тип: ${esc(filt.kind)} ✕</button><span class="sep"></span>` : "";
  return `${kind}${b("lvl", 0, filt.lvl, "Все")}${b("lvl", 2, filt.lvl, "Ошибки")}${b("lvl", 1, filt.lvl, "Внимание")}<span class="sep"></span>
    ${b("group", "", filt.group, "Все проверки")}${groups.map(g => b("group", g, filt.group, g)).join("")}<span class="sep"></span>
    <button data-dis="1" class="${filt.dis ? "on" : ""}" title="Замечания, снятые вручную (_окно\\снятые.json): посмотреть, вернуть">Снятые (${(p ? S.dismissedBy[p.id] : S.P.dismissed)?.length || 0})</button><span class="sep"></span>
    <input type="search" id="q" placeholder="Поиск: линия, щит, текст (Ctrl+F)" value="${esc(filt.q)}">`;
}

const pname = i => S.panels[i.panel]?.name || "";

let shown = [];                     // строки таблицы: data-disone / data-restore — индекс в них

export function mountTable(host, list, withPanel, panelId){
  if (filt.dis) list = (withPanel ? S.P.dismissed : S.dismissedBy[panelId]) || [];
  const rows = shown = list.filter(i => (!filt.lvl || i.level === filt.lvl) && (!filt.group || i.group === filt.group) &&
    (!filt.kind || kindsOf(i).includes(filt.kind)) &&
    (!filt.q || `${i.text} ${i.line} ${pname(i)}`.toLowerCase().includes(filt.q)));
  const cols = [
    {id: "lvl", t: "", title: "Ошибка / внимание", w: 26, v: i => i.level >= 2 ? "ошибка" : "внимание",
      h: i => `<span class="dot d${i.level}"></span>`},
    ...(withPanel ? [{id: "p", t: "Щит", w: 90, v: pname}] : []),
    {id: "line", t: "Линия", w: 90, v: i => i.line,
      h: i => S.panels[i.panel] ? `<a data-go="${esc(toPanel(i.panel, i.line, i.group === "Узлы" && /^У\d+$/.test(i.line) ? "structure" : i.line ? "lines" : "issues"))}" title="Перейти">${esc(i.line || "→ щит")}</a>` : esc(i.line)},
    {id: "text", t: "Замечание", title: "Наведите — пояснение: что значит и как исправить", w: 520, v: i => i.text, agg: "count", nogroup: 1,
      h: i => esc(i.text) + (i.hint ? ' <span class="hint">?</span>' : ""), tip: i => i.hint ? `${i.text}\n\n${i.hint}` : i.text},
    {id: "kind", t: "Тип", w: 200, hidden: 1, v: i => kindsOf(i).join("; ")},
    {id: "rule", t: "Проверка", w: 220, v: i => S.rules[i.rule]?.title || i.rule, cls: () => "mut"},
    {id: "grp", t: "Группа", w: 100, hidden: 1, v: i => i.group},
    {id: "ref", t: "Где править", w: 170, v: i => i.ref?.cell ? `${i.ref.sheet}!${i.ref.cell}` : "", h: i => refLink(i.ref)},
    ...(filt.dis ? [{id: "dwhen", t: "Снято", w: 220, v: i => i.dismissed}] : []),
    {id: "act", t: "", w: 60, nogroup: 1, v: () => "",
      h: i => filt.dis ? `<a data-restore="${shown.indexOf(i)}" title="Вернуть замечание в список">вернуть</a>`
                       : `<a class="dis" data-disone="${shown.indexOf(i)}" title="Снять замечание вручную (проверено, согласен). Книга не меняется">✕ снять</a>`},
  ];
  mount(host, {key: "issues" + (withPanel ? "A" : "P"), title: "Замечания", columns: cols, rows, freeze: withPanel ? 2 : 1,
    empty: filt.dis ? "Снятых замечаний нет." : "Замечаний нет.",
    rowAttrs: i => `class="l${i.level}"`});
}

export function bind(root, rerender){
  root.querySelectorAll("[data-lvl]").forEach(b => b.onclick = () => { filt.lvl = +b.dataset.lvl; rerender(); });
  root.querySelectorAll("[data-group]").forEach(b => b.onclick = () => { filt.group = b.dataset.group; rerender(); });
  root.querySelectorAll("[data-kind]").forEach(b => b.onclick = () => { filt.kind = b.dataset.kind; rerender(); });
  root.querySelectorAll("[data-dis]").forEach(b => b.onclick = () => { filt.dis = !filt.dis; rerender(); });
  const q = root.querySelector("#q");
  if (q) q.oninput = () => { filt.q = q.value.toLowerCase(); const pos = q.selectionStart; rerender();
    const q2 = document.querySelector("#q"); q2.focus(); q2.setSelectionRange(pos, pos); };
}

// снять / вернуть вручную (services/dismiss.py) — щелчок в таблице линий (✕ у статуса) или замечаний
const keyOf = i => ({rule: i.rule, panel: i.panel, line: i.line, text: i.text});
async function sendDismiss(items, restore){
  if (!items.length) return;
  let note = "";
  if (!restore){
    note = prompt(`Снять вручную (${items.length}):\n${items.map(i => "• " + i.text).join("\n")}\n\nПричина (можно пусто):`, "");
    if (note === null) return;
  }
  try { const r = await api.dismiss(items.map(keyOf), note, restore);
        toast(r.ok === false ? "Не сделано: " + (r.errors || []).join("; ") : (restore ? "Возвращено: " : "Снято: ") + r.n, 6000);
        window.dispatchEvent(new CustomEvent("odn:reload")); }
  catch (e){ toast("Не сделано: " + e.message, 10000); }
}
document.addEventListener("click", ev => {
  const a = ev.target.closest("[data-disline],[data-disone],[data-restore]");
  if (!a) return;
  ev.preventDefault(); ev.stopPropagation();
  if (a.dataset.disline !== undefined)
    sendDismiss((S.issuesBy[a.dataset.dispanel] || []).filter(i => i.line === a.dataset.disline), false);
  else if (a.dataset.disone !== undefined) sendDismiss([shown[+a.dataset.disone]].filter(Boolean), false);
  else sendDismiss([shown[+a.dataset.restore]].filter(Boolean), true);
}, true);
