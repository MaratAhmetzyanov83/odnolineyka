// Экран «Журнал записей в книги».
import { esc } from "../fmt.js";
import { S } from "../store.js";
import * as api from "../api.js";

export function render(){
  return `<div class="top"><h2>Журнал записей в книги</h2>
    <span class="sub">${esc(S.P.folder)}\\_окно\\журнал.jsonl · резервные копии книг — _окно\\резерв</span></div>
    <div class="tbl" id="jr"><div class="empty">Загрузка…</div></div>`;
}

export async function bind(root){
  const rows = (await api.getJournal()).rows.filter(r => r.batch);
  const el = root.querySelector("#jr");
  if (!el) return;
  el.innerHTML = rows.length ? `<table class="grid auto"><thead><tr><th>Когда</th><th>Кто</th><th>Что</th><th>Книга</th><th>Правки (было → стало)</th><th>Резерв</th></tr></thead><tbody>${rows.map(r =>
    `<tr class="${r.undone ? "mut" : ""}"><td>${esc(r.at)}</td><td>${esc(r.user)}</td><td>${esc(r.kind)}${r.undone ? " (отменена)" : ""}</td><td>${esc(r.doc)}</td>
     <td class="wrap">${r.changes.map(c => `${esc(c.line || "щит")} · ${esc(c.path)}: ${esc(c.old ?? "")} → <b>${esc(c.new ?? "")}</b> <span class="sub">${esc(c.cell)}</span>`).join("<br>")}</td>
     <td class="mut">${esc(r.backup)}</td></tr>`).join("")}</tbody></table>` : `<div class="empty">Записей ещё не было.</div>`;
}
