// Вкладка щита «Структура»: шина → узлы (УЗО, АВДТ, автомат, рубильник) → линии. Данные — программы, не Excel.
import { esc, num, toast } from "../fmt.js";
import { S, lineIn } from "../store.js";
import * as api from "../api.js";

const KINDS = ["УЗО", "АВДТ", "Автомат", "Рубильник"];
const picked = new Set();                  // выбранные линии (номера) текущего щита
let pickedPanel = "";

const issuesOf = (p, id) => (S.issuesBy[p.id] || []).filter(i => i.line === id);
const lvlOf = list => Math.max(0, ...list.map(i => i.level));
const loadPct = (i, n) => i != null && n ? i / n * 100 : null;

function bar(x){
  if (x == null) return "";
  const l = x >= 100 ? 2 : x >= 80 ? 1 : 0;
  return `<span class="sbar"><i class="l${l}" style="width:${Math.min(100, x).toFixed(0)}%"></i></span> ${num(x, 0)} %`;
}

function lineRow(p, l, depth){
  const iss = issuesOf(p, l.number), lv = lvlOf(iss);
  return `<tr class="ln l${lv}"><td><input type="checkbox" data-pick="${esc(l.number)}" ${picked.has(l.number) ? "checked" : ""}></td>
    <td style="padding-left:${8 + depth * 18}px">${esc(l.number)} <span class="sub">${esc(l.name)}</span></td>
    <td>${esc(l.scheme.qf)}</td><td>${esc(l.scheme.breaker)}</td><td class=n>${num(l.result.i_calc, 2)}</td>
    <td class=n>${num(lineIn(l), 0)}</td><td>${bar(loadPct(l.result.i_calc, lineIn(l)))}</td><td>${esc(l.scheme.phase)}</td>
    <td class="w">${l.rcd_node ? `<span class="tag">30 мА от ${esc(l.rcd_node)}</span> ` : ""}${esc(iss.map(i => i.text).join("; "))}</td><td></td></tr>`;
}

function nodeRow(p, n, depth){
  const iss = issuesOf(p, n.id), lv = lvlOf(iss);
  const ph = n.phase_i || {};
  return `<tr class="nd l${lv}"><td></td>
    <td style="padding-left:${8 + depth * 18}px"><b>▾ ${esc(n.id)}</b> <span class="tag">${esc(n.kind)}</span>
      <span class="sub">${n.lines_all.length} лин.${n.note ? " · " + esc(n.note) : ""}</span></td>
    <td><b>${esc(n.designation)}</b></td><td>${esc(n.text)}${n.device && n.device !== "авто" ? ' <span class="tag">вручную</span>' : ""}</td>
    <td class=n><b>${num(n.i_calc, 2)}</b></td><td class=n><b>${num(n.in_a, 0)}</b></td><td>${bar(loadPct(n.i_calc, n.in_a))}</td>
    <td class="sub">${["L1", "L2", "L3"].map(k => `${k} ${num(ph[k] || 0, 1)}`).join(" · ")}</td>
    <td class="w">${esc(iss.map(i => i.text).join("; "))}</td>
    <td class="nw"><button data-edit="${esc(n.id)}" title="Изменить узел">✎</button><button data-del="${esc(n.id)}" title="Убрать узел (линии — к узлу выше или на шину)">✕</button></td></tr>`;
}

function tree(p){
  const nodes = p.nodes || [], ids = new Set(nodes.map(n => n.id));
  const kids = id => nodes.filter(n => (ids.has(n.parent) ? n.parent : "") === id);
  const direct = id => p.lines.filter(l => (ids.has(l.node) ? l.node : "") === id);
  const busOf = l => l.scheme.bus || "";
  const buses = [...new Set(["", ...nodes.filter(n => !ids.has(n.parent)).map(n => n.bus || ""), ...p.lines.map(busOf)])];
  let h = "";
  const walk = (n, depth) => {
    h += nodeRow(p, n, depth);
    for (const k of kids(n.id)) walk(k, depth + 1);
    for (const l of direct(n.id)) h += lineRow(p, l, depth + 1);
  };
  for (const b of buses){
    const top = kids("").filter(n => (n.bus || "") === b), free = direct("").filter(l => busOf(l) === b);
    if (!top.length && !free.length) continue;
    h += `<tr class="bus"><td></td><td colspan="9"><b>Шина ${esc(b || "основная")}</b> <span class="sub">${top.length} узл. · ${free.length} лин. без узла</span></td></tr>`;
    for (const n of top) walk(n, 1);
    for (const l of free) h += lineRow(p, l, 1);
  }
  return h;
}

export function bar0(p){
  const nodes = p.nodes || [];
  return `<button data-s="add" class="on" title="Выделите линии галочками — узел встанет над ними">Узел над выделенными</button>
    <select data-s="target"><option value="">на шину (без узла)</option>${nodes.map(n => `<option value="${esc(n.id)}">${esc(n.id)} — ${esc(n.designation)} ${esc(n.kind)}</option>`).join("")}</select>
    <button data-s="move" title="Перенести выделенные линии">Перенести выделенные</button>
    <button data-s="none" title="Снять выделение">Снять</button>
    <button data-s="undo" title="Отменить последний шаг структуры">Отменить шаг</button>
    <span class="sub" id="spick">${picked.size ? `выделено: ${picked.size}` : ""}</span>`;
}

export function render(p){
  if (pickedPanel !== p.id){ picked.clear(); pickedPanel = p.id; }
  return `<div class="tbl"><table class="grid auto stree"><thead><tr><th></th><th>Структура</th><th>Обозн.</th><th>Аппарат</th>
    <th class=n>Iр, А</th><th class=n>In, А</th><th>Загрузка</th><th>Фазы / ток</th><th>Замечания</th><th></th></tr></thead>
    <tbody>${tree(p)}</tbody></table></div>
    <div class="gstat"><span class="sub">Узлы и их связь с линиями хранит программа (папка объекта\\_окно\\структура.json), книги Excel не меняются.
    Iр узла — сумма Iр линий под ним по наиболее нагруженной фазе, без Кс. Под УЗО / АВДТ линиям 30 мА уже дано.</span></div>`;
}

// ---------------------------------------------------------------- окно узла
function dialog(p, node, onOk){
  const others = (p.nodes || []).filter(n => !node || n.id !== node.id);
  const v = node || {kind: "УЗО", device: "авто", rcd_ma: 30, parent: "", bus: "", note: ""};
  const d = document.createElement("div");
  d.className = "gpop dlg";
  d.innerHTML = `<div class="ph"><b>${node ? "Узел " + esc(node.id) : `Новый узел над ${picked.size} лин.`}</b></div>
    <div class="row"><label>Тип</label><select data-f="kind">${KINDS.map(k => `<option ${v.kind === k ? "selected" : ""}>${k}</option>`).join("")}</select></div>
    <div class="row"><label>Аппарат</label><input data-f="device" value="${esc(v.device || "авто")}" list="dl-dev" placeholder="авто"><datalist id="dl-dev"><option value="авто"></datalist></div>
    <div class="row" data-rcd><label>Утечка, мА</label><select data-f="rcd_ma">${[10, 30, 100, 300].map(x => `<option ${(v.rcd_ma || 30) === x ? "selected" : ""}>${x}</option>`).join("")}</select></div>
    <div class="row"><label>Узел выше</label><select data-f="parent"><option value="">— на шине —</option>${others.map(n => `<option value="${esc(n.id)}" ${v.parent === n.id ? "selected" : ""}>${esc(n.id)} ${esc(n.kind)}</option>`).join("")}</select></div>
    <div class="row" data-bus><label>Шина</label><input data-f="bus" value="${esc(v.bus || "")}" placeholder="основная"></div>
    <div class="row"><label>Примечание</label><input data-f="note" value="${esc(v.note || "")}"></div>
    <div class="row end"><button data-x>Отмена</button><button data-ok class="on">${node ? "Сохранить" : "Создать"}</button></div>`;
  document.body.appendChild(d);
  const f = k => d.querySelector(`[data-f="${k}"]`);
  const sync = () => { d.querySelector("[data-rcd]").style.display = ["УЗО", "АВДТ"].includes(f("kind").value) ? "" : "none";
                       d.querySelector("[data-bus]").style.display = f("parent").value ? "none" : ""; };
  f("kind").onchange = sync; f("parent").onchange = sync; sync();
  d.querySelector("[data-x]").onclick = () => d.remove();
  d.querySelector("[data-ok]").onclick = async () => {
    const fields = {kind: f("kind").value, device: f("device").value, parent: f("parent").value,
                    bus: f("parent").value ? "" : f("bus").value, note: f("note").value,
                    rcd_ma: ["УЗО", "АВДТ"].includes(f("kind").value) ? +f("rcd_ma").value : ""};
    if (await onOk(fields)) d.remove();
  };
  d.addEventListener("keydown", ev => { if (ev.key === "Escape") d.remove(); if (ev.key === "Enter") d.querySelector("[data-ok]").click(); });
  f("device").focus();
}

async function op(body){
  try {
    const r = await api.structure(body);
    if (r.ok === false){ toast("Не сделано: " + (r.errors || []).join("; "), 8000); return false; }
    window.dispatchEvent(new CustomEvent("odn:reload"));
    return true;
  } catch (e){ toast("Не сделано: " + e.message, 8000); return false; }
}

export function bind(root, p){
  root.querySelectorAll("[data-pick]").forEach(c => c.onchange = () => {
    c.checked ? picked.add(c.dataset.pick) : picked.delete(c.dataset.pick);
    const s = document.getElementById("spick"); if (s) s.textContent = picked.size ? `выделено: ${picked.size}` : "";
  });
  const q = s => root.querySelector(`[data-s="${s}"]`);
  q("add").onclick = () => {
    if (!picked.size){ toast("Выделите линии галочками — узел встанет над ними"); return; }
    dialog(p, null, async fields => {
      const ok = await op({op: "add_node", panel: p.id, lines: [...picked], ...fields});
      if (ok){ toast("Узел создан"); picked.clear(); }
      return ok;
    });
  };
  q("move").onclick = async () => {
    if (!picked.size){ toast("Выделите линии галочками"); return; }
    if (await op({op: "move_lines", panel: p.id, lines: [...picked], node: q("target").value})){ toast("Перенесено"); picked.clear(); }
  };
  q("none").onclick = () => { picked.clear(); window.dispatchEvent(new CustomEvent("odn:rerender")); };
  q("undo").onclick = async () => {
    const r = await api.structureUndo();
    toast(r.ok ? "Отменено: " + r.undone : (r.errors || []).join("; "));
    window.dispatchEvent(new CustomEvent("odn:reload"));
  };
  root.querySelectorAll("[data-edit]").forEach(b => b.onclick = () => {
    const n = p.nodes.find(x => x.id === b.dataset.edit);
    dialog(p, n, fields => op({op: "update_node", panel: p.id, id: n.id, fields}));
  });
  root.querySelectorAll("[data-del]").forEach(b => b.onclick = async () => {
    if (confirm(`Убрать узел ${b.dataset.del}? Его линии перейдут к узлу выше или на шину.`)) await op({op: "delete_node", panel: p.id, id: b.dataset.del});
  });
}
