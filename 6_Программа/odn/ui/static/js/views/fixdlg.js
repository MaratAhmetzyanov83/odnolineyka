// Окно «Исправить ошибки»: предложения сервера (services/fixes.py) списком «было → стало», галочки.
// Правки книг идут в несохранённые (Ctrl+S — записать), узлы — сразу (структура программы), связи — «Синхронизировать книги».
import { esc, toast } from "../fmt.js";
import * as api from "../api.js";
import * as pending from "../pending.js";
import { toPanel } from "../nav.js";

const KIND = {edit: "правка книги", structure: "узел (программа)", sync: "все книги", manual: "вручную"};

export async function open(){
  document.querySelectorAll(".gpop").forEach(x => x.remove());
  let items;
  try { items = (await api.getFixes()).items; } catch (e){ toast("Не получено: " + e.message); return; }
  const auto = items.filter(x => x.kind !== "manual"), manual = items.filter(x => x.kind === "manual");
  const d = document.createElement("div");
  d.className = "gpop dlg fixdlg";
  const row = x => `<tr class="${x.kind === "manual" ? "mut" : ""}">
      <td>${x.kind === "manual" ? "" : `<input type="checkbox" data-fx="${esc(x.id)}" ${x.checked ? "checked" : ""}>`}</td>
      <td>${x.panel ? `<a data-go="${esc(toPanel(x.panel, x.kind === "edit" ? x.line : ""))}">${esc(x.panel_name)}</a>` : "объект"}</td>
      <td>${esc(x.kind === "structure" ? "узел " + x.line : x.line)}</td>
      <td class="w">${esc(x.text)}${x.checked || x.kind === "manual" ? "" : ' <span class="tag">проверьте</span>'}</td>
      <td>${esc(x.was)}</td><td>→</td><td><b>${esc(x.now)}</b></td><td class="sub">${KIND[x.kind]}</td></tr>`;
  d.innerHTML = `<div class="ph"><b>Исправить ошибки</b>
      <div class="sub">Отметьте нужное. Книги сразу не меняются: правки встанут в список несохранённых — Ctrl+S запишет.</div></div>
    ${items.length ? `<div class="tbl"><table class="grid auto"><thead><tr><th><input type="checkbox" data-all></th><th>Щит</th><th>Линия</th>
      <th>Что</th><th>Было</th><th></th><th>Станет</th><th></th></tr></thead>
      <tbody>${auto.map(row).join("")}${manual.map(row).join("")}</tbody></table></div>`
      : `<div class="empty">Исправлять автоматически нечего.</div>`}
    <div class="row end"><span class="sub" data-n></span><button data-x>Закрыть</button>${auto.length ? `<button data-ok class="on">Применить отмеченные</button>` : ""}</div>`;
  document.body.appendChild(d);
  const boxes = () => [...d.querySelectorAll("[data-fx]")];
  const count = () => { const n = d.querySelector("[data-n]"); if (n) n.textContent = `отмечено: ${boxes().filter(b => b.checked).length}`; };
  boxes().forEach(b => b.onchange = count); count();
  const all = d.querySelector("[data-all]");
  if (all) all.onchange = () => { boxes().forEach(b => b.checked = all.checked); count(); };
  d.querySelector("[data-x]").onclick = () => d.remove();
  d.addEventListener("keydown", ev => { if (ev.key === "Escape") d.remove(); });
  d.addEventListener("click", ev => { if (ev.target.closest("[data-go]")) d.remove(); });
  const ok = d.querySelector("[data-ok]");
  if (ok) ok.onclick = async () => {
    const on = new Set(boxes().filter(b => b.checked).map(b => b.dataset.fx));
    const pick = auto.filter(x => on.has(x.id));
    if (!pick.length){ toast("Ничего не отмечено"); return; }
    ok.disabled = true;
    const msg = [], errs = [];
    // 1. узлы — операции структуры (сразу, отмена — «Отменить шаг» во вкладке «Структура»)
    const ops = pick.filter(x => x.kind === "structure");
    for (const x of ops){
      try { const r = await api.structure(x.op); if (r.ok === false) errs.push(`${x.panel_name} ${x.line}: ${(r.errors || []).join("; ")}`); }
      catch (e){ errs.push(`${x.panel_name} ${x.line}: ${e.message}`); }
    }
    if (ops.length) msg.push(`узлов: ${ops.length - errs.length}`);
    // 2. правки книг — в несохранённые
    const edits = pick.filter(x => x.kind === "edit");
    if (edits.length){
      pending.set(edits.map(x => ({panel: x.panel, line: x.line, path: x.path, new: x.now, label: `${x.panel_name} ${x.line}: ${x.text}`})));
      msg.push(`правок в книги: ${edits.length} — Ctrl+S, чтобы записать`);
    }
    d.remove();
    if (ops.length) window.dispatchEvent(new CustomEvent("odn:reload"));
    // 3. синхронизация книг — после записи правок
    if (pick.some(x => x.kind === "sync")){
      if (pending.size()) msg.push("после записи нажмите «Синхронизировать книги»");
      else document.getElementById("sync")?.click();
    }
    toast((msg.length ? "Исправлено: " + msg.join("; ") : "") + (errs.length ? "\nНе сделано: " + errs.join("; ") : ""), 10000);
  };
}
