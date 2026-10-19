// Окно «Кабели КЖ»: какие кабели листа «Кабельный_журнал» идут в КЖ (DWG). Убранные — <книга>.kj.skip.txt,
// книга не меняется. Выпуск из окна и кнопка надстройки «КЖ в DWG» их пропускают.
import { esc, toast } from "../fmt.js";
import * as api from "../api.js";
import { S } from "../store.js";

const HEAD = ["Маркировка", "Начало", "Конец", "Кабель", "Лоток", "Труба ПВХ", "Труба ПНД", "Длина, м"];
const norm = v => String(v ?? "").replace(/\s+/g, " ").trim();

export async function open(docname, runKJ){
  document.querySelectorAll(".gpop").forEach(x => x.remove());
  let r;
  try { r = await api.kjCables(docname); } catch (e){ toast("Не сделано: " + e.message); return; }
  if (r.ok === false){ toast("Не сделано: " + (r.errors || []).join("; ")); return; }
  if (!r.rows.length){ toast(`В книге «${docname}» нет кабелей на листе «Кабельный_журнал»`); return; }
  const skip = new Set(r.skip);
  const doc = S.P.documents.find(d => d.file === docname);
  const d = document.createElement("div");
  d.className = "gpop dlg";
  d.style.maxWidth = "min(900px, 95vw)";
  d.innerHTML = `<div class="ph"><b>Кабели КЖ: ${esc(docname)}</b>
      <div class="sub">Снимите галочку — кабель не попадёт в КЖ (DWG). Книга не меняется: список — ${esc(r.file)}</div></div>
    <div class="row" style="gap:8px;margin:6px 0"><span id="kjn" class="sub"></span><span class="sep"></span>
      <button data-a="all">Все в КЖ</button><input type="search" id="kjq" placeholder="Поиск"></div>
    <div style="max-height:60vh;overflow:auto"><table class="kjt"><thead><tr><th>В КЖ</th>${HEAD.map(h => `<th>${esc(h)}</th>`).join("")}</tr></thead>
      <tbody>${r.rows.map((row, k) => `<tr data-k="${k}"><td><input type="checkbox" data-mk="${esc(norm(row[0]))}" ${skip.has(norm(row[0])) ? "" : "checked"}></td>
        ${row.map(v => `<td>${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
    <div class="row end"><button data-a="save" title="Записать список убранных кабелей">Сохранить</button>
      ${doc?.kj?.dwg ? `<button data-a="saverun" title="Сохранить и перестроить КЖ в ${esc(doc.kj.dwg)}">Сохранить и обновить КЖ</button>` : ""}
      <button data-a="x">Закрыть</button></div>`;
  document.body.appendChild(d);
  const boxes = [...d.querySelectorAll("input[data-mk]")];
  const count = () => { const n = boxes.filter(b => !b.checked).length;
    d.querySelector("#kjn").textContent = `В КЖ ${boxes.length - n} из ${boxes.length}` + (n ? `, убрано ${n}` : ""); };
  boxes.forEach(b => b.onchange = count); count();
  d.querySelector("#kjq").oninput = ev => { const q = ev.target.value.toLowerCase();
    d.querySelectorAll("tbody tr").forEach(tr => tr.style.display = !q || tr.textContent.toLowerCase().includes(q) ? "" : "none"); };
  d.querySelector('[data-a="all"]').onclick = () => { boxes.forEach(b => b.checked = true); count(); };
  d.querySelector('[data-a="x"]').onclick = () => d.remove();
  const save = async run => {
    const res = await api.kjSkip(docname, boxes.filter(b => !b.checked).map(b => b.dataset.mk)).catch(e => ({ok: false, errors: [e.message]}));
    if (res.ok === false){ toast("Не сделано: " + (res.errors || []).join("; ")); return; }
    d.remove();
    toast(res.n ? `Убрано из КЖ: ${res.n}` : "Все кабели — в КЖ", 6000);
    if (run) runKJ([{doc: docname, dwg: ""}]);
    else window.dispatchEvent(new CustomEvent("odn:reload"));
  };
  d.querySelector('[data-a="save"]').onclick = () => save(false);
  d.querySelector('[data-a="saverun"]')?.addEventListener("click", () => save(true));
}
