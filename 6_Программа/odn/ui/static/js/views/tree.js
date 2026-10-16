// Дерево щитов (слева): имя, Рр, ошибки, внимание.
import { esc, num } from "../fmt.js";
import { S, counts, worst, roots } from "../store.js";
import { toObject, toPanel } from "../nav.js";

const cnt = (n, cls) => `<span class="${n ? cls : "c0"}">${n}</span>`;

export function render(selId){
  const [e, w] = counts(S.P.issues);
  let h = `<div class="tree-h"><span style="text-align:left">Щит</span><span>Рр, кВт</span><span>ош.</span><span>вн.</span></div>
    <div class="node obj ${selId ? "" : "sel"}" data-go="${toObject()}"><span class="nm"><span class="dot d${e ? 2 : w ? 1 : 0}"></span>Весь объект</span>
    <span></span>${cnt(e, "ce")}${cnt(w, "cw")}</div>`;
  const byName = (a, b) => a.name.localeCompare(b.name, "ru", {numeric: true});
  const walk = (p, lvl) => {
    const [pe, pw] = counts(S.issuesBy[p.id]);
    h += `<div class="node ${selId === p.id ? "sel" : ""}" data-go="${esc(toPanel(p.id))}" title="${esc(p.doc)}">
      <span class="nm" style="padding-left:${lvl * 14}px"><span class="dot d${worst(p.id)}"></span>${esc(p.name)}</span>
      <span>${num(p.result.p_calc, 1)}</span>${cnt(pe, "ce")}${cnt(pw, "cw")}</div>`;
    for (const k of (S.kids[p.id] || []).sort(byName)) walk(k, lvl + 1);
  };
  for (const r of roots().sort((a, b) => (S.kids[b.id]?.length || 0) - (S.kids[a.id]?.length || 0) || byName(a, b))) walk(r, 0);
  for (const d of S.P.documents) if (d.error) h += `<div class="node" title="${esc(d.error)}"><span class="nm"><span class="dot d2"></span>${esc(d.file)}</span></div>`;
  return h;
}
