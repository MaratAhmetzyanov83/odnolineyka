// Таблица с сортировкой по щелчку на заголовке.
// cols: [{id, t: заголовок, n: число?, v: row => значение для сортировки, h: row => html, cls: row => класс ячейки, tip: row => подсказка}]
import { esc } from "./fmt.js";

const sort = {};   // key таблицы -> {col, dir}

export function toggle(key, col){
  const s = sort[key];
  sort[key] = s && s.col === col ? {col, dir: -s.dir} : {col, dir: 1};
}

export function grid(key, cols, rows, {rowAttrs = () => "", empty = "Нет строк."} = {}){
  if (!rows.length) return `<div class="empty">${esc(empty)}</div>`;
  const s = sort[key], c = s && cols.find(x => x.id === s.col);
  if (c){
    const val = r => { const v = (c.v || c.h)(r); return v == null || v === "" ? null : v; };
    rows = [...rows].sort((a, b) => { const x = val(a), y = val(b);
      if (x == null) return 1; if (y == null) return -1;
      return (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "ru", {numeric: true})) * s.dir; });
  }
  let h = `<table><thead><tr>${cols.map(x => `<th class="${x.n ? "n" : ""} ${s?.col === x.id ? (s.dir > 0 ? "s-asc" : "s-desc") : ""}"
    data-sort="${key}|${x.id}" title="${esc(x.title || x.t)}">${esc(x.t)}</th>`).join("")}</tr></thead><tbody>`;
  for (const r of rows){
    h += `<tr ${rowAttrs(r)}>` + cols.map(x => {
      const html = x.h ? x.h(r) : esc(x.v(r) ?? "");
      const tip = x.tip ? x.tip(r) : (x.n ? "" : String(html).replace(/<[^>]+>/g, ""));
      return `<td class="${x.n ? "n" : ""} ${x.cls ? x.cls(r) : ""}" ${tip ? `title="${esc(tip)}"` : ""}>${html}</td>`;
    }).join("") + "</tr>";
  }
  return h + "</tbody></table>";
}
