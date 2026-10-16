// Таблица «как в Excel»: сортировка, автофильтр, выделение с итогами, Ctrl+C / Ctrl+V / Ctrl+D, правка в ячейке,
// закрепление столбцов, свой набор и ширина столбцов, группировка с итогами, полосы и шкалы, выгрузка в Excel.
//
// mount(host, opts) — нарисовать в host. Состояние (сортировка, фильтры, вид, выделение) живёт по opts.key
// и переживает перерисовку экрана; вид столбцов запоминается в браузере.
//
// Столбец: {id, t, title, n, w, v(row), h(row) html, cls(row), tip(row), agg: sum|max|min|avg|count,
//           hidden, bar(row) -> {f: 0..1, lvl: 0|1|2}, edit: {path, ok(row), cur(row)}}
// opts: {key, title, columns, rows, rowId(row), rowRef(row) -> {panel, line} (для правки), rowAttrs(row),
//        freeze, empty, focusRowId, tools (html слева в панели таблицы)}
import { esc, num, toast } from "./fmt.js";
import * as pending from "./pending.js";
import { OPTIONS } from "./store.js";
import * as api from "./api.js";

const ST = {};
const AGG = { sum: "Σ", max: "max", min: "min", avg: "ср.", count: "n" };

function state(key, cols){
  if (!ST[key]){
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem("odn-grid-" + key) || "{}"); } catch (e) {}
    ST[key] = { sort: null, filters: {}, group: saved.group || "", collapsed: new Set(),
      order: saved.order || [], hidden: new Set(saved.hidden || cols.filter(c => c.hidden).map(c => c.id)),
      width: saved.width || {}, sel: null, anchor: null };
  }
  return ST[key];
}
function save(key){
  const s = ST[key];
  try { localStorage.setItem("odn-grid-" + key, JSON.stringify({order: s.order, hidden: [...s.hidden], width: s.width, group: s.group})); } catch (e) {}
}

// ---------------------------------------------------------------- значения ячеек
const isNum = x => typeof x === "number" && isFinite(x);
function cellVal(g, row, c){
  if (c.edit && g.opts.rowRef){
    const ref = g.opts.rowRef(row), p = pending.get(ref.panel, ref.line, c.edit.path);
    if (p) return p.new === "" || p.new == null ? null : (c.n && !isNaN(+String(p.new).replace(",", ".")) ? +String(p.new).replace(",", ".") : p.new);
  }
  const v = c.v ? c.v(row) : null;
  return v === "" ? null : v;
}
const shown = (c, v) => v == null ? "" : isNum(v) ? num(v, c.d ?? 2) : String(v);
function textVal(g, row, c){ return shown(c, cellVal(g, row, c)); }

// ---------------------------------------------------------------- фильтры
function passes(f, v, txt){
  if (!f) return true;
  if (f.vals && !f.vals.has(txt)) return false;
  if (f.op){
    if (!isNum(v)) return false;
    const a = f.a, b = f.b;
    switch (f.op){
      case "=": return v === a; case "≠": return v !== a; case ">": return v > a; case "≥": return v >= a;
      case "<": return v < a; case "≤": return v <= a; case "между": return v >= a && v <= b;
    }
  }
  if (f.text && !txt.toLowerCase().includes(f.text)) return false;
  return true;
}

// ---------------------------------------------------------------- построение вида
function view(g){
  const s = g.s, cols = g.cols;
  let rows = g.opts.rows.filter(r => cols.every(c => passes(s.filters[c.id], cellVal(g, r, c), textVal(g, r, c))));
  if (s.sort){
    const c = g.all.find(x => x.id === s.sort.col);
    if (c){
      const dir = s.sort.dir;
      rows = rows.map(r => [cellVal(g, r, c), r]).sort((a, b) => {
        const x = a[0], y = b[0];
        if (x == null) return y == null ? 0 : 1; if (y == null) return -1;
        return (isNum(x) && isNum(y) ? x - y : String(x).localeCompare(String(y), "ru", {numeric: true})) * dir;
      }).map(x => x[1]);
    }
  }
  g.rows = rows;
  g.items = [];
  const gc = s.group && g.all.find(c => c.id === s.group);
  if (gc){
    const groups = new Map();
    for (const r of rows){ const k = textVal(g, r, gc) || "(пусто)"; (groups.get(k) || groups.set(k, []).get(k)).push(r); }
    const keys = [...groups.keys()].sort((a, b) => a.localeCompare(b, "ru", {numeric: true}));
    g.flat = [];
    for (const k of keys){
      g.items.push({group: k, rows: groups.get(k)});
      if (!s.collapsed.has(k)) for (const r of groups.get(k)){ g.items.push({row: r, i: g.flat.length}); g.flat.push(r); }
    }
  } else {
    g.flat = rows;
    g.items = rows.map((r, i) => ({row: r, i}));
  }
}

function aggregate(g, c, rows){
  if (!c.agg) return "";
  const vals = rows.map(r => cellVal(g, r, c)).filter(isNum);
  if (c.agg === "count") return String(rows.length);
  if (!vals.length) return "";
  const x = c.agg === "sum" ? vals.reduce((a, b) => a + b, 0) : c.agg === "max" ? Math.max(...vals)
          : c.agg === "min" ? Math.min(...vals) : vals.reduce((a, b) => a + b, 0) / vals.length;
  return `<span class="agg">${AGG[c.agg]}</span> ${num(x, c.d ?? 2)}`;
}

// ---------------------------------------------------------------- рисование
function colsOf(g){
  const all = g.opts.columns, s = g.s;
  const order = [...s.order.filter(id => all.some(c => c.id === id)), ...all.map(c => c.id).filter(id => !s.order.includes(id))];
  g.all = order.map(id => all.find(c => c.id === id));
  g.cols = g.all.filter(c => !s.hidden.has(c.id));
}
const W = c => c.w || (c.n ? 64 : 110);

function cellHtml(g, row, c, ri, ci){
  const v = cellVal(g, row, c);
  let pend = false, html;
  if (c.edit && g.opts.rowRef){ const ref = g.opts.rowRef(row); pend = !!pending.get(ref.panel, ref.line, c.edit.path); }
  html = pend || !c.h ? esc(shown(c, v)) : c.h(row);
  const editable = c.edit && (!c.edit.ok || c.edit.ok(row));
  let style = "";
  if (c.bar){ const b = c.bar(row); if (b && isFinite(b.f)) style = `--bar:${Math.min(100, Math.max(0, b.f * 100)).toFixed(0)}%;`;
    if (b) html = `<span class="barv l${b.lvl || 0}"></span>${html}`; }
  const tip = c.tip ? c.tip(row) : (c.n ? "" : shown(c, v));
  const cls = [c.n ? "n" : "", c.cls ? c.cls(row) : "", editable ? "ed" : "", pend ? "pend" : "", c.bar ? "hasbar" : "", ci < g.freeze ? "fz" : ""].join(" ");
  return `<td class="${cls}" data-r="${ri}" data-c="${ci}" ${tip ? `title="${esc(tip)}"` : ""} style="${style}${ci < g.freeze ? `left:${g.left[ci]}px;` : ""}">${html}</td>`;
}

function render(g){
  colsOf(g); view(g);
  const s = g.s, cols = g.cols;
  g.freeze = Math.min(g.opts.freeze || 0, cols.length);
  g.left = []; let acc = 0;
  for (const c of cols){ g.left.push(acc); acc += s.width[c.id] || W(c); }
  const nf = Object.keys(s.filters).length;
  const groupable = g.all.filter(c => !c.nogroup);
  let h = `<div class="gtools">${g.opts.tools || ""}
    <span class="sp"></span>
    <span class="sub">Строк: <b>${g.rows.length}</b>${g.rows.length !== g.opts.rows.length ? ` из ${g.opts.rows.length}` : ""}</span>
    ${nf ? `<button data-g="unfilter" title="Снять все фильтры">Фильтры: ${nf} ✕</button>` : ""}
    <label class="gsel">Группировать:<select data-g="group"><option value="">нет</option>${groupable.map(c =>
      `<option value="${c.id}" ${s.group === c.id ? "selected" : ""}>${esc(c.t || c.title)}</option>`).join("")}</select></label>
    <button data-g="cols" title="Показать / скрыть / переставить столбцы">Столбцы</button>
    <button data-g="xlsx" title="Выгрузить то, что видно (фильтр, сортировка, столбцы), в Excel">В Excel</button></div>`;
  h += `<div class="gwrap" tabindex="0"><table class="grid" style="width:${acc}px"><colgroup>${cols.map(c => `<col style="width:${s.width[c.id] || W(c)}px">`).join("")}</colgroup><thead><tr>`;
  cols.forEach((c, ci) => {
    const so = s.sort && s.sort.col === c.id ? (s.sort.dir > 0 ? " s-asc" : " s-desc") : "";
    h += `<th class="${c.n ? "n" : ""}${so}${ci < g.freeze ? " fz" : ""}${s.filters[c.id] ? " filt" : ""}" data-c="${ci}" title="${esc(c.title || c.t)}${c.edit ? " — правится (F2)" : ""}"
      style="${ci < g.freeze ? `left:${g.left[ci]}px` : ""}"><span class="tt">${esc(c.t)}${c.edit ? '<i class="pen">✎</i>' : ""}</span><span class="fb" data-fc="${ci}" title="Фильтр">▾</span><span class="rz" data-rz="${ci}"></span></th>`;
  });
  h += `</tr></thead><tbody>`;
  if (!g.flat.length && !g.items.length) h += `<tr><td colspan="${cols.length}" class="empty">${esc(g.opts.empty || "Нет строк.")}</td></tr>`;
  for (const it of g.items){
    if (it.group != null){
      const col = s.collapsed.has(it.group);
      let span = 1; while (span < cols.length && !cols[span].agg && span < 4) span++;
      h += `<tr class="grp" data-grp="${esc(it.group)}"><td colspan="${span}" class="${g.freeze ? "fz" : ""}" style="left:0"><b>${col ? "▸" : "▾"} ${esc(it.group)}</b> <span class="sub">(${it.rows.length})</span></td>` +
        cols.slice(span).map((c, j) => { const ci = j + span;
          return `<td class="${c.n ? "n" : ""}${ci < g.freeze ? " fz" : ""}" style="${ci < g.freeze ? `left:${g.left[ci]}px` : ""}">${aggregate(g, c, it.rows)}</td>`; }).join("") + "</tr>";
      continue;
    }
    const r = it.row, rid = g.opts.rowId ? g.opts.rowId(r) : "";
    h += `<tr ${g.opts.rowAttrs ? g.opts.rowAttrs(r) : ""} data-ri="${it.i}" ${rid === g.opts.focusRowId && rid ? 'data-focus="1"' : ""}>${cols.map((c, ci) => cellHtml(g, r, c, it.i, ci)).join("")}</tr>`;
  }
  h += `</tbody><tfoot><tr>${cols.map((c, ci) => `<td class="${c.n ? "n" : ""}${ci < g.freeze ? " fz" : ""}" style="${ci < g.freeze ? `left:${g.left[ci]}px` : ""}">${ci === 0 && !c.agg ? `<b>Итог</b>` : aggregate(g, c, g.rows)}</td>`).join("")}</tr></tfoot></table></div>
    <div class="gstat"><span class="st"></span><span class="sp"></span><span class="sub">Щелчок, Shift — диапазон, Ctrl — добавить. Ctrl+C / Ctrl+V, F2 или ввод — правка, Ctrl+Enter — во все выделенные, Ctrl+D — вниз, Del — очистить</span></div>`;
  g.host.innerHTML = h;
  g.wrap = g.host.querySelector(".gwrap");
  g.trs = [...g.host.querySelectorAll("tbody tr[data-ri]")];
  bindGrid(g);
  paintSel(g);
  const f = g.host.querySelector("tr[data-focus]");
  if (f && !g.focused){ g.focused = true; f.classList.add("hl"); f.scrollIntoView({block: "center"});
    const ri = +f.dataset.ri; s.sel = [{r0: ri, c0: 0, r1: ri, c1: 0}]; s.anchor = {r: ri, c: 0}; paintSel(g); }
}

// ---------------------------------------------------------------- выделение
function cellEl(g, r, c){ const tr = g.trs[r]; return tr ? tr.querySelector(`td[data-c="${c}"]`) : null; }
function norm(a){ return {r0: Math.min(a.r0, a.r1), r1: Math.max(a.r0, a.r1), c0: Math.min(a.c0, a.c1), c1: Math.max(a.c0, a.c1)}; }
function eachSel(g, fn){
  const seen = new Set();
  for (const a of (g.s.sel || []).map(norm))
    for (let r = a.r0; r <= a.r1 && r < g.flat.length; r++) for (let c = a.c0; c <= a.c1 && c < g.cols.length; c++){
      const k = r + ":" + c; if (!seen.has(k)){ seen.add(k); fn(r, c); } }
}
function paintSel(g){
  (g.painted || []).forEach(el => el.classList.remove("sel", "act"));
  g.painted = [];
  const vals = []; let n = 0;
  eachSel(g, (r, c) => { const el = cellEl(g, r, c); if (el){ el.classList.add("sel"); g.painted.push(el); }
    n++; const v = cellVal(g, g.flat[r], g.cols[c]); if (isNum(v)) vals.push(v); });
  const a = g.s.anchor; if (a){ const el = cellEl(g, a.r, a.c); if (el){ el.classList.add("act"); g.painted.push(el); } }
  const st = g.host.querySelector(".gstat .st");
  if (!st) return;
  if (n > 1 && vals.length){
    const sum = vals.reduce((x, y) => x + y, 0);
    st.innerHTML = `Ячеек: <b>${n}</b> · чисел: <b>${vals.length}</b> · Σ <b>${num(sum, 3)}</b> · ср. <b>${num(sum / vals.length, 3)}</b> · min <b>${num(Math.min(...vals), 3)}</b> · max <b>${num(Math.max(...vals), 3)}</b>`;
  } else st.innerHTML = n > 1 ? `Ячеек: <b>${n}</b>` : "";
}
function select(g, r, c, mode){
  r = Math.max(0, Math.min(g.flat.length - 1, r)); c = Math.max(0, Math.min(g.cols.length - 1, c));
  const s = g.s;
  if (mode === "extend" && s.anchor){ const last = s.sel[s.sel.length - 1]; last.r1 = r; last.c1 = c; }
  else if (mode === "add"){ s.sel = [...(s.sel || []), {r0: r, c0: c, r1: r, c1: c}]; s.anchor = {r, c}; }
  else { s.sel = [{r0: r, c0: c, r1: r, c1: c}]; s.anchor = {r, c}; }
  if (mode === "extend") g.cursor = {r, c}; else g.cursor = {r, c};
  paintSel(g);
  cellEl(g, r, c)?.scrollIntoView({block: "nearest", inline: "nearest"});
}

// ---------------------------------------------------------------- правка
function editTarget(g, r, c){
  const col = g.cols[c], row = g.flat[r];
  if (!col || !row || !col.edit || !g.opts.rowRef) return null;
  if (col.edit.ok && !col.edit.ok(row)) return null;
  const ref = g.opts.rowRef(row);
  return {col, row, ref};
}
function stage(g, cells, value, label){
  // cells: [[r, c]] -> правки (только правимые); value: строка или функция (r, c) -> строка
  const items = []; let skipped = 0;
  for (const [r, c] of cells){
    const t = editTarget(g, r, c);
    if (!t){ skipped++; continue; }
    const v = typeof value === "function" ? value(r, c) : value;
    if (v === undefined) continue;
    const cur = t.col.edit.cur ? t.col.edit.cur(t.row) : cellVal(g, t.row, t.col);
    const normv = x => t.col.n || (OPTIONS[t.col.edit.path] || {}).type !== "text" ? String(x ?? "").replace(",", ".").trim() : String(x ?? "").trim();
    const same = normv(cur) === normv(v);
    items.push({panel: t.ref.panel, line: t.ref.line, path: t.col.edit.path, new: v, label: `${t.ref.name || ""} ${t.ref.line}: ${t.col.t}`, same});
  }
  pending.set(items);
  if (skipped && label) toast(`${label}: ${skipped} яч. не правятся (расчётные или не первая строка группы) — пропущены`);
  return items.length;
}
function startEdit(g, r, c, initial){
  const t = editTarget(g, r, c);
  if (!t){ toast("Эта ячейка не правится: расчёт или данные Excel"); return; }
  const td = cellEl(g, r, c);
  const p = pending.get(t.ref.panel, t.ref.line, t.col.edit.path);
  const cur = p ? p.new : (t.col.edit.cur ? t.col.edit.cur(t.row) : cellVal(g, t.row, t.col));
  const opts = (OPTIONS[t.col.edit.path] || {}).options || [];
  const id = "dl-" + t.col.edit.path.replace(/\W/g, "_");
  td.classList.add("editing");
  td.innerHTML = `<input class="ced" list="${id}" value="${esc(initial ?? (cur ?? ""))}"><datalist id="${id}">${opts.map(o => `<option value="${esc(o)}">`).join("")}</datalist>`;
  const inp = td.querySelector("input");
  inp.focus();
  if (initial == null) inp.select(); else inp.setSelectionRange(inp.value.length, inp.value.length);
  let done = false;
  const finish = (commit, move, all) => {
    if (done) return; done = true;
    if (commit){
      const cells = all ? (() => { const a = []; eachSel(g, (rr, cc) => a.push([rr, cc])); return a; })() : [[r, c]];
      stage(g, cells, inp.value, "Правка");
    }
    refresh(g);
    if (move) select(g, r + move[0], c + move[1]);
    g.wrap.focus();
  };
  inp.onkeydown = ev => {
    ev.stopPropagation();
    if (ev.key === "Enter"){ ev.preventDefault(); finish(true, ev.ctrlKey ? null : [ev.shiftKey ? -1 : 1, 0], ev.ctrlKey); }
    else if (ev.key === "Tab"){ ev.preventDefault(); finish(true, [0, ev.shiftKey ? -1 : 1]); }
    else if (ev.key === "Escape"){ ev.preventDefault(); finish(false); }
  };
  inp.onblur = () => finish(true);
}

// ---------------------------------------------------------------- буфер обмена
function copyText(g){
  const sel = (g.s.sel || []).map(norm);
  if (!sel.length) return "";
  const a = sel[sel.length - 1], lines = [];
  for (let r = a.r0; r <= a.r1 && r < g.flat.length; r++){
    const out = [];
    for (let c = a.c0; c <= a.c1; c++){ const v = cellVal(g, g.flat[r], g.cols[c]); out.push(v == null ? "" : isNum(v) ? String(v).replace(".", ",") : String(v).replace(/[\t\n]/g, " ")); }
    lines.push(out.join("\t"));
  }
  return lines.join("\r\n");
}
function paste(g, text){
  const a = g.s.anchor; if (!a) return;
  const m = text.replace(/\r/g, "").replace(/\n$/, "").split("\n").map(l => l.split("\t"));
  // одно значение и выделено много — во все выделенные (как в Excel)
  if (m.length === 1 && m[0].length === 1){
    const cells = []; eachSel(g, (r, c) => cells.push([r, c]));
    const n = stage(g, cells, m[0][0], "Вставка"); toast(`Вставлено: ${n}`); refresh(g); return;
  }
  const cells = [];
  m.forEach((row, i) => row.forEach((_, j) => { if (a.r + i < g.flat.length && a.c + j < g.cols.length) cells.push([a.r + i, a.c + j]); }));
  const n = stage(g, cells, (r, c) => m[r - a.r]?.[c - a.c], "Вставка");
  toast(`Вставлено: ${n}`); refresh(g);
}
function fillDown(g){
  const cells = [], top = {};
  for (const s of (g.s.sel || []).map(norm)) for (let c = s.c0; c <= s.c1; c++) for (let r = s.r0 + 1; r <= s.r1; r++){ cells.push([r, c]); top[r + ":" + c] = s.r0; }
  if (!cells.length){ toast("Ctrl+D: выделите диапазон вниз от образца"); return; }
  const n = stage(g, cells, (r, c) => { const v = cellVal(g, g.flat[top[r + ":" + c]], g.cols[c]); return v == null ? "" : String(v); }, "Заполнить вниз");
  toast(`Заполнено: ${n}`); refresh(g);
}

// ---------------------------------------------------------------- всплывающие окна (фильтр, столбцы)
function popup(g, anchor, html){
  closePopup();
  const d = document.createElement("div");
  d.className = "gpop"; d.innerHTML = html;
  document.body.appendChild(d);
  const r = anchor.getBoundingClientRect();
  d.style.left = Math.min(r.left, innerWidth - d.offsetWidth - 8) + "px";
  d.style.top = Math.min(r.bottom + 2, innerHeight - d.offsetHeight - 8) + "px";
  setTimeout(() => document.addEventListener("mousedown", outside), 0);
  return d;
}
function outside(ev){ if (!ev.target.closest(".gpop")) closePopup(); }
function closePopup(){ document.querySelectorAll(".gpop").forEach(x => x.remove()); document.removeEventListener("mousedown", outside); }

function filterPopup(g, ci, anchor){
  const c = g.cols[ci], f = g.s.filters[c.id] || {};
  // значения — по строкам, прошедшим остальные фильтры
  const others = g.opts.rows.filter(r => g.cols.every(x => x === c || passes(g.s.filters[x.id], cellVal(g, r, x), textVal(g, r, x))));
  const cnt = new Map();
  for (const r of others){ const t = textVal(g, r, c); cnt.set(t, (cnt.get(t) || 0) + 1); }
  const vals = [...cnt.keys()].sort((a, b) => c.n ? (parseFloat(a.replace(",", ".")) || 0) - (parseFloat(b.replace(",", ".")) || 0) : a.localeCompare(b, "ru", {numeric: true}));
  const d = popup(g, anchor, `<div class="ph"><b>${esc(c.title || c.t)}</b></div>
    <div class="row"><button data-p="asc">▲ по возр.</button><button data-p="desc">▼ по убыв.</button></div>
    ${c.n ? `<div class="row">Условие: <select data-p="op">${["", "=", "≠", ">", "≥", "<", "≤", "между"].map(o => `<option ${f.op === o ? "selected" : ""}>${o}</option>`).join("")}</select>
      <input data-p="a" size="6" value="${f.a ?? ""}"> <input data-p="b" size="6" value="${f.b ?? ""}" placeholder="до"></div>` : ""}
    <div class="row"><input data-p="q" placeholder="Поиск значений" style="width:100%"></div>
    <div class="row"><label><input type="checkbox" data-p="all" ${!f.vals ? "checked" : ""}> (выделить все)</label></div>
    <div class="vals">${vals.slice(0, 1000).map(v => `<label data-v="${esc(v.toLowerCase())}"><input type="checkbox" value="${esc(v)}" ${!f.vals || f.vals.has(v) ? "checked" : ""}> ${esc(v || "(пусто)")} <span class="sub">${cnt.get(v)}</span></label>`).join("")}</div>
    <div class="row end"><button data-p="clear">Очистить</button><button data-p="ok" class="on">ОК</button></div>`);
  const q = d.querySelector('[data-p="q"]');
  q.oninput = () => d.querySelectorAll(".vals label").forEach(l => l.style.display = l.dataset.v.includes(q.value.toLowerCase()) ? "" : "none");
  d.querySelector('[data-p="all"]').onchange = ev => d.querySelectorAll(".vals label").forEach(l => { if (l.style.display !== "none") l.querySelector("input").checked = ev.target.checked; });
  const close = () => { closePopup(); refresh(g); };
  d.querySelector('[data-p="asc"]').onclick = () => { g.s.sort = {col: c.id, dir: 1}; close(); };
  d.querySelector('[data-p="desc"]').onclick = () => { g.s.sort = {col: c.id, dir: -1}; close(); };
  d.querySelector('[data-p="clear"]').onclick = () => { delete g.s.filters[c.id]; close(); };
  d.querySelector('[data-p="ok"]').onclick = () => {
    const nf = {};
    const boxes = [...d.querySelectorAll(".vals input")];
    if (boxes.some(b => !b.checked)) nf.vals = new Set(boxes.filter(b => b.checked).map(b => b.value));
    if (c.n){
      const op = d.querySelector('[data-p="op"]').value, a = parseFloat(d.querySelector('[data-p="a"]').value.replace(",", ".")),
            b = parseFloat(d.querySelector('[data-p="b"]').value.replace(",", "."));
      if (op && isFinite(a)) Object.assign(nf, {op, a, b: isFinite(b) ? b : a});
    }
    if (Object.keys(nf).length) g.s.filters[c.id] = nf; else delete g.s.filters[c.id];
    close();
  };
  q.focus();
}

function colsPopup(g, anchor){
  const s = g.s;
  const draw = () => g.all.map((c, i) => `<div class="crow"><label><input type="checkbox" data-id="${c.id}" ${s.hidden.has(c.id) ? "" : "checked"}> ${esc(c.title || c.t)}</label>
    <span><button data-up="${i}" title="Выше">↑</button><button data-dn="${i}" title="Ниже">↓</button></span></div>`).join("");
  const d = popup(g, anchor, `<div class="ph"><b>Столбцы</b> <span class="sub">галочка — показать, ↑↓ — порядок; закреплено первых: ${g.opts.freeze || 0}</span></div>
    <div class="clist">${draw()}</div><div class="row end"><button data-p="reset">Как было</button><button data-p="ok" class="on">Готово</button></div>`);
  const redraw = () => { d.querySelector(".clist").innerHTML = draw(); wire(); };
  const wire = () => {
    d.querySelectorAll("[data-id]").forEach(b => b.onchange = () => { b.checked ? s.hidden.delete(b.dataset.id) : s.hidden.add(b.dataset.id); save(g.key); refresh(g); });
    const move = (i, j) => { const ids = g.all.map(c => c.id); if (j < 0 || j >= ids.length) return; [ids[i], ids[j]] = [ids[j], ids[i]]; s.order = ids; save(g.key); refresh(g); redraw(); };
    d.querySelectorAll("[data-up]").forEach(b => b.onclick = () => move(+b.dataset.up, +b.dataset.up - 1));
    d.querySelectorAll("[data-dn]").forEach(b => b.onclick = () => move(+b.dataset.dn, +b.dataset.dn + 1));
  };
  wire();
  d.querySelector('[data-p="reset"]').onclick = () => { s.order = []; s.width = {}; s.hidden = new Set(g.opts.columns.filter(c => c.hidden).map(c => c.id)); save(g.key); closePopup(); refresh(g); };
  d.querySelector('[data-p="ok"]').onclick = closePopup;
}

async function exportXlsx(g){
  const cols = g.cols.map(c => ({t: c.title || c.t, n: !!c.n}));
  const rows = g.rows.map(r => g.cols.map(c => { const v = cellVal(g, r, c); return v == null ? null : v; }));
  try { await api.download("/api/export", {title: g.opts.title || "Таблица", columns: cols, rows}); }
  catch (e){ toast("Не выгружено: " + e.message); }
}

// ---------------------------------------------------------------- события
function bindGrid(g){
  const h = g.host, w = g.wrap;
  h.querySelector('[data-g="group"]').onchange = ev => { g.s.group = ev.target.value; g.s.collapsed.clear(); save(g.key); refresh(g); };
  h.querySelector('[data-g="cols"]').onclick = ev => colsPopup(g, ev.target);
  h.querySelector('[data-g="xlsx"]').onclick = () => exportXlsx(g);
  const uf = h.querySelector('[data-g="unfilter"]'); if (uf) uf.onclick = () => { g.s.filters = {}; refresh(g); };
  h.querySelectorAll("th").forEach(th => th.onclick = ev => {
    if (ev.target.closest(".rz")) return;
    const ci = +th.dataset.c;
    if (ev.target.closest(".fb")){ filterPopup(g, ci, th); return; }
    const c = g.cols[ci], so = g.s.sort;
    g.s.sort = so && so.col === c.id ? (so.dir > 0 ? {col: c.id, dir: -1} : null) : {col: c.id, dir: 1};
    refresh(g);
  });
  h.querySelectorAll(".rz").forEach(z => z.onmousedown = ev => {
    ev.preventDefault(); ev.stopPropagation();
    const ci = +z.dataset.rz, c = g.cols[ci], col = h.querySelectorAll("col")[ci], x0 = ev.clientX, w0 = col.offsetWidth || W(c);
    const mv = e => { col.style.width = Math.max(30, w0 + e.clientX - x0) + "px"; };
    const up = e => { document.removeEventListener("mousemove", mv); document.removeEventListener("mouseup", up);
      g.s.width[c.id] = Math.max(30, w0 + e.clientX - x0); save(g.key); refresh(g); };
    document.addEventListener("mousemove", mv); document.addEventListener("mouseup", up);
  });
  h.querySelectorAll("tr.grp").forEach(tr => tr.onclick = () => { const k = tr.dataset.grp; g.s.collapsed.has(k) ? g.s.collapsed.delete(k) : g.s.collapsed.add(k); refresh(g); });

  let drag = false;
  w.onmousedown = ev => {
    const td = ev.target.closest("td[data-r]");
    if (!td || ev.target.closest("a") || ev.target.closest("input")) return;
    const r = +td.dataset.r, c = +td.dataset.c;
    select(g, r, c, ev.shiftKey ? "extend" : ev.ctrlKey ? "add" : "");
    drag = true; w.focus({preventScroll: true});
    const mv = e => { const t = e.target.closest && e.target.closest("td[data-r]"); if (drag && t) select(g, +t.dataset.r, +t.dataset.c, "extend"); };
    const up = () => { drag = false; document.removeEventListener("mouseover", mv); document.removeEventListener("mouseup", up); };
    document.addEventListener("mouseover", mv); document.addEventListener("mouseup", up);
  };
  w.ondblclick = ev => { const td = ev.target.closest("td[data-r]"); if (td && !ev.target.closest("a")) startEdit(g, +td.dataset.r, +td.dataset.c); };
  w.onkeydown = ev => keys(g, ev);
  w.addEventListener("copy", ev => { if (ev.target.closest && ev.target.closest("input")) return; const t = copyText(g); if (t){ ev.clipboardData.setData("text/plain", t); ev.preventDefault(); toast("Скопировано"); } });
  w.addEventListener("paste", ev => { if (ev.target.closest && ev.target.closest("input")) return; ev.preventDefault(); paste(g, ev.clipboardData.getData("text/plain")); });
}

function keys(g, ev){
  if (ev.target.closest("input,select")) return;
  const a = g.cursor || g.s.anchor || {r: 0, c: 0}, ext = ev.shiftKey ? "extend" : "";
  const page = Math.max(5, Math.floor(g.wrap.clientHeight / 22) - 2);
  const k = ev.key;
  const mv = (r, c) => { ev.preventDefault(); select(g, r, c, ext); };
  if (k === "ArrowDown") return mv(ev.ctrlKey ? g.flat.length - 1 : a.r + 1, a.c);
  if (k === "ArrowUp") return mv(ev.ctrlKey ? 0 : a.r - 1, a.c);
  if (k === "ArrowRight") return mv(a.r, ev.ctrlKey ? g.cols.length - 1 : a.c + 1);
  if (k === "ArrowLeft") return mv(a.r, ev.ctrlKey ? 0 : a.c - 1);
  if (k === "PageDown") return mv(a.r + page, a.c);
  if (k === "PageUp") return mv(a.r - page, a.c);
  if (k === "Home") return mv(ev.ctrlKey ? 0 : a.r, 0);
  if (k === "End") return mv(ev.ctrlKey ? g.flat.length - 1 : a.r, g.cols.length - 1);
  if (k === "Tab") return mv(a.r, a.c + (ev.shiftKey ? -1 : 1));
  if (ev.ctrlKey && (k === "a" || k === "ф")){ ev.preventDefault(); g.s.sel = [{r0: 0, c0: 0, r1: g.flat.length - 1, c1: g.cols.length - 1}]; paintSel(g); return; }
  if (ev.ctrlKey && (k === "d" || k === "в")){ ev.preventDefault(); fillDown(g); return; }
  if (!g.s.anchor) return;
  if (k === "F2" || k === "Enter"){ ev.preventDefault(); startEdit(g, a.r, a.c); return; }
  if (k === "Delete" || k === "Backspace"){
    ev.preventDefault(); const cells = []; eachSel(g, (r, c) => cells.push([r, c]));
    const n = stage(g, cells, "", "Очистить"); if (n){ toast(`Очищено: ${n}`); refresh(g); } return;
  }
  if (k.length === 1 && !ev.ctrlKey && !ev.altKey && !ev.metaKey){ ev.preventDefault(); startEdit(g, a.r, a.c, k); }
}

// ---------------------------------------------------------------- снаружи
const GRIDS = {};
function refresh(g){ const top = g.wrap?.scrollTop, left = g.wrap?.scrollLeft; render(g); if (g.wrap){ g.wrap.scrollTop = top || 0; g.wrap.scrollLeft = left || 0; } }

export function mount(host, opts){
  const s = state(opts.key, opts.columns);
  const g = GRIDS[opts.key] && GRIDS[opts.key].host === host ? GRIDS[opts.key] : {key: opts.key, s};
  g.host = host; g.opts = opts; g.s = s;
  GRIDS[opts.key] = g;
  g.focused = false;
  render(g);
  return g;
}
// перерисовать все таблицы (после изменения правок)
pending.onChange(() => Object.values(GRIDS).forEach(g => { if (g.host && g.host.isConnected) refresh(g); }));
