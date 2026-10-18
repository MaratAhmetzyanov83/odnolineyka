// Окно «Фидер» нижестоящего щита: параметры фидера живут в строке линии вышестоящего щита.
// Правки идут туда (в список несохранённых), Ctrl+S пишет книгу вышестоящего; книга этого щита пересчитывается вместе.
import { esc, num, toast } from "../fmt.js";
import { S, OPTIONS } from "../store.js";
import * as pending from "../pending.js";

const SERIES = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600];
const isInput = l => /^ввод/i.test(l.number);
const inOf = l => l.scheme.in_a ?? l.result.in_by_current;

// селективность: In фидера ≥ 1,6 × In ввода щита (если ввод с аппаратом), иначе × наибольшего отходящего; и ≥ Iр щита
export function selectivity(p){
  const inp = p.lines.find(l => isInput(l) && inOf(l));
  const outs = p.lines.filter(l => !isInput(l) && inOf(l)).map(inOf);
  const base = inp ? inOf(inp) : (outs.length ? Math.max(...outs) : null);
  const need = Math.max(base ? 1.6 * base : 0, p.result.i_calc || 0);
  const pick = SERIES.find(x => x >= need - 1e-9) || SERIES[SERIES.length - 1];
  return {inp, base, need, pick, by: inp ? `ввода ${inp.number}` : "наибольшего отходящего"};
}

const FIELDS = [
  ["input.length_plan", "Lпл, м", l => l.input.length_plan],
  ["input.cable_code", "Марка (код)", l => l.input.cable_code],
  ["input.cores", "Жил", l => l.input.cores],
  ["input.s_min", "Sмин, мм²", l => l.input.s_min],
  ["input.parallel", "Кабелей ∥", l => l.input.parallel],
  ["input.in_manual", "In ручной, А", l => l.input.in_manual],
  ["scheme.breaker_set", "Автомат (авто / текст)", l => l.scheme.breaker_set],
];

export function open(p){
  const par = S.panels[p.parent];
  const ln = par && par.lines.find(l => l.number === p.parent_line);
  if (!ln){ toast("Нет линии-фидера: сначала «Связь…»"); return; }
  document.querySelectorAll(".gpop").forEach(x => x.remove());
  const brkOk = ln.scheme.group_head || !ln.scheme.qf;
  const cur = (path, l) => { const pp = pending.get(par.id, ln.number, path); return pp ? pp.new : (l ?? ""); };
  const d = document.createElement("div");
  d.className = "gpop dlg";
  d.innerHTML = `<div class="ph"><b>Фидер ${esc(p.name)}</b>
      <div class="sub">линия ${esc(ln.number)} щита ${esc(par.name)} (${esc(par.doc)}):
      ${esc(ln.scheme.breaker)} · ${esc(ln.scheme.cable)} · Iр ${num(ln.result.i_calc, 1)} А · ΔU ${num(ln.result.du, 2)} %</div></div>
    ${FIELDS.filter(f => f[0] !== "scheme.breaker_set" || brkOk).map(([path, t, get]) => {
      const o = (OPTIONS[path] || {}).options || [];
      return `<div class="row"><label>${t}</label><input data-p="${path}" value="${esc(cur(path, get(ln)))}" list="fd-${path.replace(/\W/g, "_")}">
        <datalist id="fd-${path.replace(/\W/g, "_")}">${o.map(x => `<option value="${esc(x)}">`).join("")}</datalist></div>`; }).join("")}
    ${(() => { const q = selectivity(p), have = ln.scheme.in_a ?? ln.result.in_by_current, bad = q.base && have < 1.6 * q.base;
      return `<div class="row ${bad ? "bad" : ""}" style="display:block">Селективность: In фидера ≥ 1,6 × ${num(q.base, 0) || "—"} А (${q.by})
        и ≥ Iр щита ${num(p.result.i_calc, 1)} А → нужно ≥ ${num(q.need, 1)} А; сейчас ${num(have, 0)} А ${bad ? "— НЕСЕЛЕКТИВНО" : "— ок"}
        <button class="mini" data-a="pick" title="In ручной = ${q.pick} А">Подобрать In: ${q.pick} А</button></div>`; })()}
    <div class="sub">Пусто — по расчёту. Правки — в книгу ${esc(par.name)}; книга ${esc(p.name)} пересчитается вместе (Ctrl+S).</div>
    <div class="row end"><button data-a="x">Отмена</button><button data-a="ok" class="on">В правки</button></div>`;
  document.body.appendChild(d);
  d.querySelector('[data-a="x"]').onclick = () => d.remove();
  d.querySelector('[data-a="pick"]').onclick = () => { d.querySelector('[data-p="input.in_manual"]').value = selectivity(p).pick; };
  d.querySelector('[data-a="ok"]').onclick = () => {
    const items = [];
    d.querySelectorAll("[data-p]").forEach(i => {
      const f = FIELDS.find(x => x[0] === i.dataset.p);
      const was = String(f[2](ln) ?? ""), now = i.value.trim();
      items.push({panel: par.id, line: ln.number, path: f[0], new: now, label: `${par.name} ${ln.number}: ${f[1]}`,
                  same: was.replace(",", ".") === now.replace(",", ".")});
    });
    pending.set(items);
    d.remove();
    const n = items.filter(x => !x.same).length;
    toast(n ? `Правок фидера: ${n}. Ctrl+S — записать (${par.name} и ${p.name} пересчитаются вместе)` : "Ничего не изменилось", 7000);
  };
}
