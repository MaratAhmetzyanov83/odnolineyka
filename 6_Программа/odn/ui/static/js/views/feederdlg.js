// Окно «Фидер» нижестоящего щита: параметры фидера живут в строке линии вышестоящего щита.
// Правки идут туда (в список несохранённых), Ctrl+S пишет книгу вышестоящего; книга этого щита пересчитывается вместе.
import { esc, num, toast } from "../fmt.js";
import { S, OPTIONS } from "../store.js";
import * as pending from "../pending.js";
import * as api from "../api.js";

const SEL_HOW = [["", "по правилу 1,6 × In"], ["по времени", "по времени (выдержка tsd вышестоящего)"],
                 ["по таблице", "по таблице производителя"], ["на ступень", "на ступень выше (вышестоящий — следующий номинал)"]];

const SERIES = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600];
const isInput = l => /^ввод/i.test(l.number);
const inOf = l => l.scheme.in_a ?? l.result.in_by_current;
// аппарат ввода отключает КЗ: автомат; QS (выключатель нагрузки, рубильник) и пустой ввод — нет (как checks/links._trips)
const trips = l => { const t = (l.scheme.breaker || "").trim().toLowerCase();
                     return !!t && !/^(qs|вн|рубильник|выключатель нагрузки)/.test(t); };

// селективность: In фидера ≥ 1,6 × In ввода щита (если ввод с аппаратом), иначе × наибольшего отходящего; и ≥ Iр щита
export function selectivity(p){
  const inp = p.lines.find(l => isInput(l) && inOf(l) && trips(l));
  const outs = p.lines.filter(l => !isInput(l) && inOf(l)).map(inOf);
  const how = (p.selectivity || {}).how, step = how === "на ступень";
  const other = step ? "" : how;                              // по времени / по таблице: правило 1,6 × In не нужно
  const base = other ? null : inp ? inOf(inp) : (outs.length ? Math.max(...outs) : null);
  const sel = base ? (step ? (SERIES.find(x => x > base + 1e-9) || SERIES[SERIES.length - 1]) : 1.6 * base) : 0;
  const need = Math.max(sel, p.result.i_calc || 0);
  const pick = SERIES.find(x => x >= need - 1e-9) || SERIES[SERIES.length - 1];
  return {inp, base, sel, need, pick, other, step, by: inp ? `ввода ${inp.number}` : "наибольшего отходящего"};
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
      ${esc(ln.scheme.breaker)} · ${esc(ln.scheme.cable)} · Iр ${num(ln.result.i_calc, 1)} А · ΔU ${num(ln.result.du, 2)} %<br>
      прокладка: лоток ${num(ln.input.tray || 0, 0)} м · гофра ПВХ ${num(ln.input.pvc || 0, 0)} м · гофра ПНД ${num(ln.input.pnd || 0, 0)} м
      ${Object.keys(ln.mismatch || {}).length ? `<br><b class="w">Строка ввода ${esc(p.name)} не совпадает: ${esc(Object.entries(ln.mismatch).map(([k, v]) => k.split(".").pop() + " там " + (v ?? "пусто")).join("; "))}</b>` : ""}</div></div>
    ${FIELDS.filter(f => f[0] !== "scheme.breaker_set" || brkOk).map(([path, t, get]) => {
      const o = (OPTIONS[path] || {}).options || [];
      return `<div class="row"><label>${t}</label><input data-p="${path}" value="${esc(cur(path, get(ln)))}" list="fd-${path.replace(/\W/g, "_")}">
        <datalist id="fd-${path.replace(/\W/g, "_")}">${o.map(x => `<option value="${esc(x)}">`).join("")}</datalist></div>`; }).join("")}
    ${(() => { const q = selectivity(p), have = ln.scheme.in_a ?? ln.result.in_by_current, bad = q.base && have < q.sel - 1e-9;
      const sel = p.selectivity || {};
      return `<div class="row ${bad ? "bad" : ""}" style="display:block">${q.other
          ? `Селективность — ${esc(q.other)}: правило 1,6 × In не проверяется. In фидера ≥ Iр щита ${num(p.result.i_calc, 1)} А; сейчас ${num(have, 0)} А`
          : `Селективность: In фидера ≥ ${q.step ? `${num(q.sel, 0)} А (на ступень выше ${num(q.base, 0) || "—"} А` : `1,6 × ${num(q.base, 0) || "—"} А (`}${q.by})
        и ≥ Iр щита ${num(p.result.i_calc, 1)} А → нужно ≥ ${num(q.need, 1)} А; сейчас ${num(have, 0)} А ${bad ? "— НЕСЕЛЕКТИВНО" : "— ок"}`}
        <button class="mini" data-a="pick" title="In ручной = ${q.pick} А">Подобрать In: ${q.pick} А</button></div>
      <div class="row"><label title="Селективность обеспечена иначе: вышестоящий с выдержкой времени (расцепитель LSI, tsd) или пара из таблицы селективности производителя. Хранит программа (_окно\\структура.json)">Селективность</label>
        <select data-sel>${SEL_HOW.map(([v, t]) => `<option value="${esc(v)}" ${v === (sel.how || "") ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></div>
      <div class="row"><label>Основание</label><input data-seln value="${esc(sel.note || "")}" placeholder="напр. tsd 0,2 с; таблица Chint NM8N / NB1-63"></div>`; })()}
    <div class="sub">Пусто — по расчёту. Правки — в книгу ${esc(par.name)}; книга ${esc(p.name)} пересчитается вместе (Ctrl+S).</div>
    <div class="row end"><button data-a="x">Отмена</button><button data-a="ok" class="on">В правки</button></div>`;
  document.body.appendChild(d);
  d.querySelector('[data-a="x"]').onclick = () => d.remove();
  d.querySelector('[data-a="pick"]').onclick = () => { d.querySelector('[data-p="input.in_manual"]').value = selectivity(p).pick; };
  d.querySelector('[data-a="ok"]').onclick = async () => {
    const sel = p.selectivity || {}, how = d.querySelector("[data-sel]").value, note = d.querySelector("[data-seln]").value.trim();
    if (how !== (sel.how || "") || (how && note !== (sel.note || ""))){       // отметка — сразу, в данные программы
      try { const r = await api.structure({op: "set_selectivity", panel: p.id, how, note});
            if (r.ok === false) throw new Error((r.errors || []).join("; "));
            window.dispatchEvent(new CustomEvent("odn:reload"));
            toast(how ? `Селективность ${p.name}: ${how}` : `Селективность ${p.name}: по правилу 1,6 × In`, 5000); }
      catch (e){ toast("Селективность не отмечена: " + e.message, 8000); }
    }
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
