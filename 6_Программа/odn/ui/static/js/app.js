// Точка входа окна: загрузка, маршруты экранов, общие щелчки.
// Новый экран: модуль в views/ с render(ctx) и bind(root, rerender, ctx) + строка в VIEWS.
import * as api from "./api.js";
import { S, setProject } from "./store.js";
import { parse } from "./nav.js";
import { toast } from "./fmt.js";
import * as grid from "./grid.js";
import * as tree from "./views/tree.js";
import * as objectView from "./views/object.js";
import * as panelView from "./views/panel.js";

const VIEWS = { object: objectView, panel: panelView };
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
  try { setProject(await api.getProject()); }
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
  const th = ev.target.closest("[data-sort]");
  if (th){ const [key, col] = th.dataset.sort.split("|"); grid.toggle(key, col); render(); return; }
  const go = ev.target.closest("[data-go]");
  if (go){ ev.preventDefault(); location.hash = go.dataset.go; }
});
window.addEventListener("hashchange", render);
$("#reload").onclick = async () => {
  const b = $("#reload"); b.disabled = true; b.textContent = "Читаю книги…";
  try { await api.reload(); await load(); toast("Обновлено"); } catch (e){ toast("Ошибка: " + e.message); }
  b.disabled = false; b.textContent = "Обновить";
};
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
