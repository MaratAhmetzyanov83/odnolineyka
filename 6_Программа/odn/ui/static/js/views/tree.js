// Дерево щитов (слева).
import { esc } from "../fmt.js";
import { S, counts, worst, roots } from "../store.js";
import { toObject, toPanel } from "../nav.js";

export function render(selId){
  const [e, w] = counts(S.P.issues);
  let h = `<div class="node obj ${selId ? "" : "sel"}" data-go="${toObject()}"><span class="dot d${e ? 2 : w ? 1 : 0}"></span>Весь объект<span class="cnt">${e} / ${w}</span></div>`;
  const byName = (a, b) => a.name.localeCompare(b.name, "ru");
  const walk = (p, lvl) => {
    const [pe, pw] = counts(S.issuesBy[p.id]);
    h += `<div class="node ${selId === p.id ? "sel" : ""}" data-go="${esc(toPanel(p.id))}" style="padding-left:${12 + lvl * 18}px" title="${esc(p.doc)}">
      <span class="dot d${worst(p.id)}"></span>${esc(p.name)}<span class="cnt">${pe} / ${pw}</span></div>`;
    for (const k of (S.kids[p.id] || []).sort(byName)) walk(k, lvl + 1);
  };
  for (const r of roots().sort((a, b) => (S.kids[b.id]?.length || 0) - (S.kids[a.id]?.length || 0) || byName(a, b))) walk(r, 0);
  for (const d of S.P.documents) if (d.error) h += `<div class="node" title="${esc(d.error)}"><span class="dot d2"></span>${esc(d.file)}</div>`;
  return h;
}
