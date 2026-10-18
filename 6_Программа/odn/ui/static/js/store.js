// Данные объекта и индексы. Экранам — только чтение.
export const S = { P: null, panels: {}, issuesBy: {}, rules: {}, kids: {} };
export const OPTIONS = {};          // правимые поля: {путь: {title, type, options, when}} — с сервера
export function setOptions(o){ for (const k in OPTIONS) delete OPTIONS[k]; Object.assign(OPTIONS, o); }

export function setProject(P){
  S.P = P; S.panels = {}; S.issuesBy = {}; S.kids = {}; S.rules = {};
  for (const d of P.documents) for (const p of d.panels) S.panels[p.id] = p;
  for (const p of Object.values(S.panels)) if (p.parent && S.panels[p.parent]) (S.kids[p.parent] ||= []).push(p);
  for (const i of P.issues) (S.issuesBy[i.panel] ||= []).push(i);
  for (const r of P.rules || []) S.rules[r.id] = r;
}
export const lineIn = l => l.scheme.in_a ?? l.result.in_by_current;   // как Line.in_a в модели
export const counts = list => { let e = 0, w = 0; for (const i of list || []) i.level >= 2 ? e++ : w++; return [e, w]; };
export const worst = id => Math.max(0, ...(S.issuesBy[id] || []).map(i => i.level));
export const roots = () => Object.values(S.panels).filter(p => !p.parent || !S.panels[p.parent]);
// однолинейка в DWG (отметки SXDRAW): {txt, cls, tip} — худшее по чертежам щита
const DW = {stale: ["устарела", "ce", 2], check: ["проверить", "cw", 1], ok: ["актуальна", "ok", 0]};
export function drawState(p){
  const ds = p.drawings || [];
  if (!ds.length) return {txt: "нет отметки", cls: "c0", tip: "Схему не строили SXDRAW v3.26+ (отметка — «<книга>.drawn.txt» рядом с книгой)"};
  const w = ds.reduce((a, d) => DW[d.state][2] > DW[a.state][2] ? d : a, ds[0]);
  const name = d => d.dwg.split(/[\\/]/).pop();
  return {txt: DW[w.state][0] + (ds.length > 1 ? ` (${ds.length} DWG)` : ""), cls: DW[w.state][1],
          tip: ds.map(d => `${name(d)}: ${DW[d.state][0]}, построена ${d.at}${d.why ? " — " + d.why : ""}`).join("\n")};
}
// КЖ в DWG (отметка надстройки «КЖ в DWG») — по книге щита
export function kjState(p){
  const k = S.P.documents.find(d => d.file === p.doc)?.kj || {};
  if (!k.state) return {txt: "нет отметки", cls: "c0", tip: "КЖ не записывали кнопкой «КЖ в DWG» новой надстройки (отметка — «<книга>.kj.txt»)"};
  const [txt, cls] = k.state === "stale" ? ["устарел", "ce"] : ["актуален", "ok"];
  return {txt, cls, tip: `${k.dwg.split(/[\\/]/).pop()}: записан ${k.at}, кабелей в книге ${k.cables}${k.why ? "\n" + k.why : ""}`};
}
