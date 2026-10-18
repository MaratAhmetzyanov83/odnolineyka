// Точка входа окна: загрузка, маршруты экранов, общие щелчки.
// Новый экран: модуль в views/ с render(ctx) и bind(root, rerender, ctx) + строка в VIEWS.
import * as api from "./api.js";
import { S, setProject } from "./store.js";
import { parse } from "./nav.js";
import { toast, esc } from "./fmt.js";
import * as pending from "./pending.js";
import { setOptions } from "./store.js";
import * as tree from "./views/tree.js";
import * as objectView from "./views/object.js";
import * as panelView from "./views/panel.js";
import * as journalView from "./views/journal.js";

const VIEWS = { object: objectView, panel: panelView, journal: journalView };
const $ = s => document.querySelector(s);

function render(){
  if (!S.P) return;
  const ctx = parse(), view = VIEWS[ctx.view] || objectView;
  $("#tree").innerHTML = tree.render(ctx.view === "panel" ? ctx.id : "");
  const main = $("#main");
  main.innerHTML = view.render(ctx);
  view.bind?.(main, render, ctx);
}

async function load(){
  try { setProject(await api.getProject()); setOptions(await api.getOptions()); }
  catch (e){ $("#main").innerHTML = `<div class="empty">Ошибка загрузки: ${e.message}</div>`; return; }
  $("#title").textContent = "Объект: " + S.P.name;
  $("#stamp").textContent = `прочитано ${S.P.loaded} · документов ${S.P.documents.length} · щитов ${Object.keys(S.panels).length}`;
  render();
}

// общие щелчки: переход (data-go), открыть в источнике (data-ref)
document.addEventListener("click", async ev => {
  const ref = ev.target.closest("[data-ref]");
  if (ref){
    ev.preventDefault(); ev.stopPropagation();
    try { const r = await api.openRef(JSON.parse(ref.dataset.ref)); toast(r.msg === "ok" ? "Открыто в Excel" : r.msg); }
    catch (e){ toast("Не открыто: " + e.message); }
    return;
  }
  const ex = ev.target.closest("[data-exclude]");
  if (ex){
    ev.preventDefault();
    if (!confirm(`Не читать книгу «${ex.dataset.exclude}»?\nОна попадёт в список исключений объекта (_окно\\исключить.txt). Вернуть — убрать строку из этого файла.`)) return;
    try { const r = await api.exclude(ex.dataset.exclude); if (r.ok === false) throw new Error((r.errors || []).join("; "));
          location.hash = "#/"; await load(); toast("Книга больше не читается: " + ex.dataset.exclude); }
    catch (e){ toast("Не сделано: " + e.message); }
    return;
  }
  const go = ev.target.closest("[data-go]");
  if (go){ ev.preventDefault(); location.hash = go.dataset.go; }
});
window.addEventListener("hashchange", render);
window.addEventListener("odn:reload", () => load());          // экран поменял данные на сервере
window.addEventListener("odn:rerender", () => render());
$("#reload").onclick = async () => {
  const b = $("#reload"); b.disabled = true; b.textContent = "Читаю книги…";
  try { await api.reload(); await load(); toast("Обновлено"); } catch (e){ toast("Ошибка: " + e.message); }
  b.disabled = false; b.textContent = "Обновить";
};
// ---- несохранённые правки: панель в шапке, Ctrl+S — записать, Ctrl+Z — отменить
function pendBar(){
  const n = pending.size(), el = $("#pend");
  el.innerHTML = n ? `Правок: <b>${n}</b> <button id="psave" class="on" title="Записать в книги Excel (Ctrl+S)">Записать в книги</button>
    <button id="plist" title="Что изменено">Что</button><button id="pdrop" title="Отменить все несохранённые правки">Отменить все</button>` : "";
  if (!n) return;
  $("#psave").onclick = saveAll;
  $("#pdrop").onclick = () => { if (confirm(`Отменить ${n} несохранённых правок?`)) pending.clear(); };
  $("#plist").onclick = () => alert(pending.all().map(p => `${p.label}: ${p.new === "" ? "(пусто)" : p.new}`).join("\n"));
}
pending.onChange(pendBar);
window.addEventListener("beforeunload", ev => { if (pending.size()){ ev.preventDefault(); ev.returnValue = ""; } });

async function saveAll(){
  const list = pending.all();
  if (!list.length) return;
  const b = $("#psave"); if (b){ b.disabled = true; b.textContent = "Пишу в книги…"; }
  try {
    const r = await api.edit(list.map(({panel, line, path, new: v}) => ({panel, line, path, new: v})));
    if (r.written){ pending.clear(); await load(); }
    toast((r.written ? `Записано: ${r.written} (${r.docs.join(", ")}). Книги пересчитаны, проверки обновлены.` : "Ничего не записано.")
      + (r.related?.length ? `\nПересчитаны связанные: ${r.related.join(", ")}` : "")
      + (r.warnings?.length ? "\nВнимание: " + r.warnings.join("; ") : "")
      + (r.errors?.length ? "\nОшибки: " + r.errors.join("; ") : ""), r.errors?.length || r.warnings?.length ? 12000 : 6000);
  } catch (e){ toast("Не записано: " + e.message, 10000); }
  pendBar();
}
async function undoWrite(){
  if (!confirm("Отменить последнюю запись в книгу? Значения вернутся как были.")) return;
  try {
    const r = await api.undo();
    if (r.written) await load();
    toast(r.written ? `Отменено: ${r.written} знач. в ${r.docs.join(", ")}` : "Не отменено: " + (r.errors || []).join("; "), 8000);
  } catch (e){ toast("Не отменено: " + e.message); }
}
$("#undo").onclick = undoWrite;
$("#sync").onclick = async () => {
  if (!confirm("Пересчитать все связанные книги объекта вместе и сохранить?\nКниги должны быть закрыты в Excel. Перед этим — копии в _окно\\резерв.")) return;
  const b = $("#sync"); b.disabled = true; b.textContent = "Пересчитываю…";
  try { const r = await api.sync(); if (r.ok){ await load(); }
    toast(r.ok ? `Синхронизировано книг: ${r.docs.length}` + (r.warnings?.length ? "\nНе тронуты: " + r.warnings.join("; ") : "")
               : "Не сделано: " + (r.errors || []).join("; "), 10000); }
  catch (e){ toast("Не сделано: " + e.message, 10000); }
  b.disabled = false; b.textContent = "Синхронизировать книги";
};

$("#journal").onclick = () => { location.hash = "#/journal"; };

// горячие клавиши окна
document.addEventListener("keydown", ev => {
  if (ev.ctrlKey && (ev.key === "s" || ev.key === "ы")){ ev.preventDefault(); saveAll(); return; }
  if (ev.ctrlKey && (ev.key === "f" || ev.key === "а")){ const q = $("#lq") || $("#q"); if (q){ ev.preventDefault(); q.focus(); q.select(); } return; }
  if (ev.ctrlKey && (ev.key === "z" || ev.key === "я") && !ev.target.closest("input")){
    ev.preventDefault();
    if (pending.size()) pending.undo(); else undoWrite();
  }
});

// тема: авто -> светлая -> тёмная; запоминается в браузере
const THEMES = [["", "авто"], ["light", "светлая"], ["dark", "тёмная"]];
function showTheme(){ const t = document.documentElement.dataset.theme || "";
  $("#theme").textContent = "Тема: " + (THEMES.find(x => x[0] === t) || THEMES[0])[1]; }
$("#theme").onclick = () => {
  const cur = document.documentElement.dataset.theme || "";
  const next = THEMES[(THEMES.findIndex(x => x[0] === cur) + 1) % THEMES.length][0];
  if (next) document.documentElement.dataset.theme = next; else delete document.documentElement.dataset.theme;
  try { next ? localStorage.setItem("odn-theme", next) : localStorage.removeItem("odn-theme"); } catch (e) {}
  showTheme();
};
showTheme();
load();
