// Окно «Связь…» щита: от какого щита и по какой линии он питается.
// «Сохранить в программе» — связь хранит программа (_окно/структура.json); «Записать в книги» — как кнопка надстройки.
import { esc, toast } from "../fmt.js";
import { S } from "../store.js";
import * as api from "../api.js";

const below = (pid, id) => { for (let x = S.panels[id]; x; x = S.panels[x.parent]) if (x.parent === pid) return true; return false; };

export function open(p){
  document.querySelectorAll(".gpop").forEach(x => x.remove());
  const cands = Object.values(S.panels).filter(q => q.id !== p.id && !below(p.id, q.id))
    .sort((a, b) => a.name.localeCompare(b.name, "ru", {numeric: true}));
  const n = Math.max(1, p.scheme.inputs || 1);
  const d = document.createElement("div");
  d.className = "gpop dlg";
  d.innerHTML = `<div class="ph"><b>Питание щита ${esc(p.name)}</b>
      <div class="sub">сейчас: ${p.parent ? `${esc(S.panels[p.parent]?.name)} / ${esc(p.parent_line || "?")} (${esc(p.link_src)})` : "не связан"}
      ${p.book_parent && p.link_src === "программа" ? `<br>в книге: ${esc(S.panels[p.book_parent]?.name)} / ${esc(p.book_line || "?")}` : ""}</div></div>
    <div class="row"><label>От щита</label><select data-f="parent"><option value="">— выберите —</option>
      ${cands.map(q => `<option value="${esc(q.id)}" ${q.id === p.parent ? "selected" : ""}>${esc(q.name)} — ${esc(q.doc)}</option>`).join("")}</select></div>
    <div class="row"><label>Линия-фидер</label><select data-f="line"></select></div>
    ${n > 1 ? `<div class="row"><label>Ввод</label><select data-f="input">${Array.from({length: n}, (_, i) => `<option>${i + 1}</option>`).join("")}</select></div>` : ""}
    <div class="row end" style="flex-wrap:wrap">
      ${p.link_src === "программа" ? `<button data-a="clear" title="Снова брать связь из книги">Убрать связь программы</button>` : ""}
      <button data-a="x">Отмена</button>
      <button data-a="prog" class="on" title="Связь хранит программа; книги не меняются">Сохранить в программе</button>
      <button data-a="book" title="Как кнопка надстройки «Связать с вышестоящим»: связь в обе книги (Excel получит ΔU и Iкз)">Записать в книги</button></div>
    <div class="sub" style="margin-top:6px">«Записать в книги»: обе книги закрыты в Excel; перед записью — их копии в _окно\\резерв.
      <a data-a="undo">Отменить последнюю запись связи</a></div>`;
  document.body.appendChild(d);
  const f = k => d.querySelector(`[data-f="${k}"]`);
  const fillLines = async () => {
    const par = S.panels[f("parent").value];
    if (!par){ f("line").innerHTML = ""; return; }
    let pick = par.id === p.parent ? p.parent_line : "";
    if (!pick){ try { pick = (await api.linkSuggest(par.id, p.id)).line; } catch (e) {} }
    const feeders = par.lines.filter(l => l.input.group_ks === "Щит" || l.number === pick);
    const rest = par.lines.filter(l => !feeders.includes(l));
    const opt = l => `<option value="${esc(l.number)}" ${l.number === pick ? "selected" : ""}>${esc(l.number)} — ${esc(l.name)}${l.feeds && l.feeds !== p.id ? ` (уже: ${esc(S.panels[l.feeds]?.name)})` : ""}</option>`;
    f("line").innerHTML = `<option value="">— не знаю —</option>${feeders.length ? `<optgroup label="Фидеры">${feeders.map(opt).join("")}</optgroup>` : ""}<optgroup label="Все линии">${rest.map(opt).join("")}</optgroup>`;
  };
  f("parent").onchange = fillLines;
  fillLines();
  const body = () => ({panel: p.id, parent: f("parent").value, line: f("line").value, input: f("input") ? +f("input").value : 1});
  const done = async (r, okText) => {
    if (r.ok === false){ toast("Не сделано: " + (r.errors || []).join("; "), 10000); return; }
    d.remove(); toast(okText + (r.msg ? "\n" + r.msg : ""), 8000);
    window.dispatchEvent(new CustomEvent("odn:reload"));
  };
  d.querySelector('[data-a="x"]').onclick = () => d.remove();
  d.querySelector('[data-a="prog"]').onclick = async () => done(await api.structure({op: "set_link", ...body()}), "Связь сохранена в программе");
  d.querySelector('[data-a="book"]').onclick = async ev => {
    const b = body();
    if (!b.parent || !b.line){ toast("Выберите щит и линию-фидер"); return; }
    ev.target.disabled = true; ev.target.textContent = "Пишу в книги…";
    try { await done(await api.linkWrite(b), "Связь записана в книги"); }
    catch (e){ toast("Не записано: " + e.message, 10000); }
    ev.target.disabled = false; ev.target.textContent = "Записать в книги";
  };
  const cl = d.querySelector('[data-a="clear"]');
  if (cl) cl.onclick = async () => done(await api.structure({op: "clear_link", panel: p.id}), "Связь программы убрана");
  d.querySelector('[data-a="undo"]').onclick = async () => {
    if (!confirm("Вернуть книги к копиям до последней записи связи?")) return;
    const r = await api.linkUndo();
    done(r, "Отменено: " + (r.undone || ""));
  };
}
