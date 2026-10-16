// Данные объекта и индексы. Экранам — только чтение.
export const S = { P: null, panels: {}, issuesBy: {}, rules: {}, kids: {} };

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
