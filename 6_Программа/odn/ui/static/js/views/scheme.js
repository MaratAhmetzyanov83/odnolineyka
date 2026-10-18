// Вкладка объекта «Схема»: структурная схема питания. Щиты — прямоугольники, линии — фидеры, у каждого ввода свой цвет.
// Расстановка — по основному вводу (дерево слева направо); вводы 2 и 3 — дополнительные линии к своему порту.
import { esc, num } from "../fmt.js";
import { S, roots, worst } from "../store.js";
import { toPanel } from "../nav.js";

const W = 184, H = 80, DX = W + 240, DY = 22, PAD = 16, SRC = 160;
let zoom = 1;

const cut = (s, n) => (s = String(s || "")).length > n ? s.slice(0, n - 1) + "…" : s;
const byName = (a, b) => a.name.localeCompare(b.name, "ru", {numeric: true});

// вводы щита: [{input, parent, line, breaker, cable, length, manual, from}]
function inputsOf(p){
  const main = {input: 1, parent: S.panels[p.parent] ? p.parent : "", line: p.parent_line, breaker: p.supply.breaker,
                cable: p.supply.cable, length: p.supply.length, manual: p.supply.manual, from: p.supply.from_doc || p.supply.from_panel};
  const more = (p.inputs_more || []).map(sp => ({input: sp.input, parent: S.panels[sp.parent] ? sp.parent : "", line: sp.parent_line,
                breaker: sp.breaker, cable: sp.cable, length: sp.length, manual: sp.manual, from: sp.from_doc || sp.from_panel}));
  // пусто в книге щита — взять с линии-фидера вышестоящего
  for (const i of [main, ...more]){
    const l = i.parent && S.panels[i.parent].lines.find(x => x.number === i.line);
    if (!l) continue;
    i.breaker ||= l.scheme.breaker; i.cable ||= l.scheme.cable; i.length ??= l.input.length;
    i.lay = [["лоток", l.input.tray], ["ПВХ", l.input.pvc], ["ПНД", l.input.pnd]].filter(x => x[1]).map(([k, v]) => `${k} ${num(v, 0)} м`).join(" · ");
    i.mm = Object.keys(l.mismatch || {}).length > 0;
  }
  return [main, ...more];
}
const nInputs = p => Math.max(p.scheme.inputs || 1, ...inputsOf(p).map(i => i.input));

function sections(p){
  if (nInputs(p) < 2) return "";
  const s = {};
  for (let k = 1; k <= nInputs(p); k++) s[k] = 0;
  for (const l of p.lines) s[l.input.section || 1] = (s[l.input.section || 1] || 0) + (l.input.p_kw || 0);
  return Object.keys(s).sort().map(k => `С${k} ${num(s[k], 1)}`).join(" · ") + " кВт";
}

function layout(){
  const pos = {};
  let row = 0, maxDepth = 0;
  const place = (p, depth) => {
    maxDepth = Math.max(maxDepth, depth);
    const kids = (S.kids[p.id] || []).slice().sort((a, b) =>
      String(a.parent_line).localeCompare(String(b.parent_line), "ru", {numeric: true}));
    const ys = kids.map(k => place(k, depth + 1));
    const y = ys.length ? (ys[0] + ys[ys.length - 1]) / 2 : row++ * (H + DY);
    pos[p.id] = {x: PAD + SRC + depth * DX, y: PAD + y};
    return y;
  };
  for (const r of roots().sort(byName)){ place(r, 0); row += 0.4; }
  return {pos, w: PAD * 2 + SRC + maxDepth * DX + W + 10, h: PAD * 2 + row * (H + DY)};
}

function box(p, at, ext){
  const lv = worst(p.id), n = nInputs(p), sec = sections(p);
  const ports = Array.from({length: n}, (_, k) =>
    `<circle class="port i${Math.min(k + 1, 3)}" cx="${at.x}" cy="${at.y + H * (k + 1) / (n + 1)}" r="3.5"><title>Ввод ${k + 1}</title></circle>`).join("");
  return `<g class="pbox l${lv}" data-go="${esc(toPanel(p.id))}"><title>${esc(p.doc)}</title>
    <rect x="${at.x}" y="${at.y}" width="${W}" height="${H}" rx="5"/>
    <circle class="dot d${lv}" cx="${at.x + 10}" cy="${at.y + 13}" r="4"/>
    <text x="${at.x + 19}" y="${at.y + 17}" class="nm">${esc(cut(p.name, 18))}${ext ? " ↗" : ""}</text>
    ${p.vru ? `<rect class="vru" x="${at.x + W - 38}" y="${at.y + 5}" width="32" height="15" rx="3"/><text class="vrut" x="${at.x + W - 22}" y="${at.y + 16}">ВРУ</text>` : ""}
    <text x="${at.x + 8}" y="${at.y + 33}">Рр ${num(p.result.p_calc, 1)} кВт · Iр ${num(p.result.i_calc, 1)} А</text>
    <text x="${at.x + 8}" y="${at.y + 73}" class="mu${p.phase_skew > 15 ? " sk" : ""}"><title>ΣIр по фазам с нижестоящими, перекос ${num(p.phase_skew, 0)} %</title>${["L1", "L2", "L3"].map(k => `${k} ${num((p.phase_i || {})[k] || 0, 0)}`).join(" · ")} А${p.phase_skew > 15 ? ` · ${num(p.phase_skew, 0)} %` : ""}</text>
    <text x="${at.x + 8}" y="${at.y + 47}" class="mu">${n} ввод${n > 1 ? "а" : ""}${p.scheme.avr ? " · АВР " + esc(p.scheme.avr) : ""} · ${p.lines.length} лин.</text>
    ${sec ? `<text x="${at.x + 8}" y="${at.y + 60}" class="mu">${esc(sec)}</text>` : ""}
    ${ports}</g>`;
}

// подпись у порта: 1 ввод — две строки (кабель второй); 2–3 ввода — одна строка (порты рядом), кабель — в подсказке
function edgeLabel(inp, x, y, one){
  const len = inp.length != null ? num(inp.length, 0) + " м" : "";
  const head = [`В${inp.input}`, inp.line, inp.breaker, one ? "" : len].filter(Boolean).join(" · ");
  const tail = one ? [inp.cable, len].filter(Boolean).join(" · ") : "";
  const full = [`Ввод ${inp.input}`, inp.line && "линия " + inp.line, inp.breaker, inp.cable, len, inp.lay,
                inp.mm && "кабель в книгах разный (см. замечания)"].filter(Boolean).join("\n");
  return `<g class="elab${inp.mm ? " warn" : ""}"><title>${esc(full)}</title>
    <text x="${x}" y="${one ? y - 4 : y - 3}">${esc(cut(head, 40))}</text>${tail ? `<text x="${x}" y="${y + 10}" class="mu">${esc(cut(tail, 40))}</text>` : ""}</g>`;
}

export function mount(host){
  const ps = Object.values(S.panels);
  if (!ps.length){ host.innerHTML = `<div class="empty">Щитов нет.</div>`; return; }
  const ext = new Set(S.P.documents.filter(d => d.external).map(d => d.file));
  const {pos, w, h} = layout();
  const out = {};                             // id щита -> число исходящих (для портов справа)
  const edges = [];
  for (const p of ps) for (const inp of inputsOf(p)) if (inp.parent) edges.push({p, inp});
  for (const e of edges) (out[e.inp.parent] ||= []).push(e);
  for (const list of Object.values(out)) list.sort((a, b) => pos[a.p.id].y - pos[b.p.id].y || a.inp.input - b.inp.input);
  let lines = "", labels = "", srcs = "";
  for (const [pid, list] of Object.entries(out)){
    const a = pos[pid];
    list.forEach((e, i) => {
      const b = pos[e.p.id], n = nInputs(e.p), k = Math.min(e.inp.input, 3);
      const x1 = a.x + W, y1 = a.y + H * (i + 1) / (list.length + 1);
      const x2 = b.x, y2 = b.y + H * e.inp.input / (n + 1);
      const mx = x2 > x1 ? (x1 + x2) / 2 : x1 + 60;
      const c2 = x2 > x1 ? mx : x2 - 60;
      lines += `<path class="edge i${k}" d="M${x1} ${y1} C${mx} ${y1} ${c2} ${y2} ${x2} ${y2}"/>`;
      labels += edgeLabel(e.inp, x2 - 8, y2, n === 1);
    });
  }
  // вводы без связи: источник вручную («ТП») или не найденный вышестоящий
  for (const p of ps) for (const inp of inputsOf(p)){
    if (inp.parent || (inp.input === 1 && S.panels[p.parent])) continue;
    const b = pos[p.id], n = nInputs(p), y2 = b.y + H * inp.input / (n + 1), k = Math.min(inp.input, 3);
    const name = inp.manual || (inp.from ? "нет щита: " + inp.from : "нет связи");
    const x1 = b.x - 150, tail = [inp.breaker, inp.length != null ? num(inp.length, 0) + " м" : ""].filter(Boolean).join(" · ");
    lines += `<path class="edge i${k}${inp.manual ? "" : " dash"}" d="M${x1} ${y2} L${b.x} ${y2}"/>`;
    srcs += `<g class="src${inp.manual ? "" : " warn"}"><title>${esc([name, inp.breaker, inp.cable].filter(Boolean).join("\n"))}</title>
      <text x="${b.x - 8}" y="${y2 - 4}">В${inp.input} ← ${esc(cut(name, 22))}</text>${tail ? `<text x="${b.x - 8}" y="${y2 + 10}" class="mu">${esc(cut(tail, 26))}</text>` : ""}</g>`;
  }
  const boxes = ps.map(p => pos[p.id] ? box(p, pos[p.id], ext.has(p.doc)) : "").join("");
  host.innerHTML = `<div class="schbar"><button data-z="-">−</button><button data-z="0">100 %</button><button data-z="+">+</button>
      <span class="lg"><i class="i1"></i>ввод 1 <i class="i2"></i>ввод 2 <i class="i3"></i>ввод 3 <i class="i1 dash"></i>нет связи</span>
      <span class="sub">Щелчок по щиту — открыть. Подпись у линии: ввод · линия-фидер · автомат · кабель · L.</span></div>
    <div class="tbl sch"><svg class="scheme" width="${w * zoom}" height="${h * zoom}" viewBox="0 0 ${w} ${h}">
      ${lines}${boxes}${labels}${srcs}</svg></div>`;
  host.querySelectorAll("[data-z]").forEach(b => b.onclick = () => {
    zoom = b.dataset.z === "0" ? 1 : Math.min(2, Math.max(0.3, zoom * (b.dataset.z === "+" ? 1.2 : 1 / 1.2)));
    const svg = host.querySelector("svg"); svg.setAttribute("width", w * zoom); svg.setAttribute("height", h * zoom);
  });
}
