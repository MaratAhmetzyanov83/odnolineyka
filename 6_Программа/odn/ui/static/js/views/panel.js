// Экран щита: строка показателей (KV — список функций), вкладки (TABS).
import { esc, num, refLink, toast } from "../fmt.js";
import { S, counts, drawState, kjState, specState, kc } from "../store.js";
import { toPanel } from "../nav.js";
import * as api from "../api.js";
import { mount } from "../grid.js";
import { lineColumns, lineRef, lineAttrs } from "../columns.js";
import * as issues from "./issues.js";
import * as structure from "./structure.js";
import * as linkdlg from "./linkdlg.js";
import * as feederdlg from "./feederdlg.js";
import * as vrudlg from "./vrudlg.js";
import * as kjdlg from "./kjdlg.js";

const kv = (k, v, cls = "", tip = "") => v === "" || v == null ? "" :
  `<div ${tip ? `title="${esc(tip)}"` : ""}><span class="k">${k}</span><b class="${cls}">${v}</b></div>`;
const lf = { q: "", only: false };              // фильтр линий

export const KV = [
  p => { const par = S.panels[p.parent];
    const src = p.link_src ? ` <span class="tag" title="${p.link_src === "программа" ? "связь хранит программа" : "связь записана в книге"}">${esc(p.link_src)}</span>` : "";
    return kv("Питание", (par ? `<a data-go="${esc(toPanel(par.id, p.parent_line, "lines"))}">${esc(par.name)}</a> / ${esc(p.parent_line || "?")}${src}`
                              : esc(p.supply.manual || p.supply.from_doc || "не задано"))
      + ` <button class="mini" data-linkdlg title="Задать, от какого щита и по какой линии питается щит">Связь…</button>`); },
  p => { const par = S.panels[p.parent], ln = par && par.lines.find(l => l.number === p.parent_line);
    const txt = ln ? [ln.scheme.breaker, (ln.scheme.cable || "").replace(/\s*L=.*$/, "")].filter(Boolean).join(", ")
                   : [p.supply.breaker, p.supply.cable].filter(Boolean).join(", ");
    return kv("Фидер", esc(txt) + (ln ? ` <button class="mini" data-feederdlg title="Параметры фидера (строка ${esc(ln.number)} щита ${esc(par.name)})">✎</button>` : "")); },
  p => kv("L", p.supply.length != null ? num(p.supply.length, 0) + " м" : ""),
  p => kv("ВРУ здания", (p.vru ? `<b class="vru">ВРУ</b> <span class="tag">${esc(p.vru_src)}</span>` : "нет")
      + ` <button class="mini" data-vrudlg title="Ввод в здание: потери считаются от шин этого щита">ВРУ…</button>`),
  p => kv("Руст", num(p.result.p_inst, 2) + " кВт"),
  p => kv("Рр", num(p.result.p_calc, 2) + " кВт"),
  p => kv("Кс", kc(p) != null ? num(kc(p), 2) : "", "", "Коэффициент спроса щита: Рр / Руст"),
  p => kv("Iр", num(p.result.i_calc, 1) + " А"),
  p => kv("cos φ", num(p.result.cos, 2)),
  p => kv("ΔU до шин", p.result.du_bus != null ? num(p.result.du_bus, 2) + " %" : ""),
  p => kv("Iкз min", p.result.ikz_min_bus != null ? num(p.result.ikz_min_bus, 0) + " А" : ""),
  p => kv("Iкз max", p.result.ikz_max_bus != null ? num(p.result.ikz_max_bus, 2) + " кА" : ""),
  p => kv("Вводы", esc([p.scheme.inputs, p.scheme.avr && p.scheme.avr !== "нет" ? "АВР" : ""].filter(Boolean).join(" + "))),
  p => kv("Кат.", esc(p.result.category)),
  p => { const s = p.phase_i || {}, own = p.phase_own || {}, skew = p.phase_skew || 0, fed = p.lines.some(l => Object.keys(l.phase_i || {}).length);
    const txt = o => ["L1", "L2", "L3"].map(k => `${k} ${num(o[k] || 0, 1)}`).join(" · ");
    return kv(fed ? "Фазы ΣIр (с нижними)" : "Фазы ΣIр", txt(s) + (skew > 15 ? ` (перекос ${num(skew, 0)} %)` : ""), skew > 15 ? "w" : "",
      "Сумма Iр по фазам, без Кс; перекос = (max − min) / max." + (fed ? `\nФидеры несут фазные токи своих щитов.\nТолько свои линии: ${txt(own)}` : "")
      + (p.phase_unknown ? `\nЛиний без фазы: ${p.phase_unknown}` : "")); },
  p => { const [e, w] = counts(S.issuesBy[p.id]); return kv("Ошибок / вн.", `<span class="${e ? "e" : ""}">${e}</span> / <span class="${w ? "w" : ""}">${w}</span>`); },
  p => { const d = drawState(p); return kv("Однолинейка", `<span class="${d.cls}">${esc(d.txt)}</span>`, "", d.tip); },
  p => { const d = kjState(p); return kv("КЖ", `<span class="${d.cls}">${esc(d.txt)}</span>`, "", d.tip); },
  p => { const d = specState(p); return d.txt ? kv("Спецификация", `<span class="${d.cls}">${esc(d.txt)}</span>`
      + (d.state !== "ok" ? ` <button class="mini" data-specok title="Спецификацию в книге поправили — запомнить аппараты как есть">Спецификация верна</button>` : ""), "", d.tip) : ""; },
  p => kv("Отпечаток", esc(p.revision), "", "Отпечаток данных схемы: по нему сверяется отметка построения DWG"),
];

export const TABS = [
  { id: "lines", title: p => `Линии (${p.lines.length})`,
    bar: () => `<input type="search" id="lq" placeholder="Поиск: линия, QF, автомат, кабель (Ctrl+F)" value="${esc(lf.q)}">
      <label class="chk"><input type="checkbox" id="lonly" ${lf.only ? "checked" : ""}>только с замечаниями</label>`,
    mount: (host, p, prm) => {
      const q = lf.q;
      const rows = p.lines.filter(l => (!lf.only || l.result.level) &&
        (!q || `${l.number} ${l.name} ${l.scheme.qf} ${l.scheme.breaker} ${l.scheme.cable} ${l.result.status}`.toLowerCase().includes(q)));
      mount(host, {key: "lines", title: p.name, columns: lineColumns(), rows, freeze: 3, rowRef: lineRef, rowAttrs: lineAttrs,
        rowId: l => l.panel + "#" + l.number, focusRowId: prm.line ? p.id + "#" + prm.line : ""});
    } },
  { id: "structure", title: p => `Структура${p.nodes && p.nodes.length ? ` (${p.nodes.length} узл.)` : ""}`, bar: structure.bar0,
    html: p => structure.render(p), mount: (host, p) => structure.bind(host.closest("main"), p) },
  { id: "issues", title: p => `Замечания (${(S.issuesBy[p.id] || []).length})`, bar: issues.filters,
    mount: (host, p) => issues.mountTable(host, S.issuesBy[p.id] || [], false, p.id) },
];

// КЖ в DWG (кнопка надстройки «КЖ в DWG» из окна): обновить DWG из отметки .kj.txt или выпустить в новый
const kjBase = f => { let s = f.split(/[\\/]/).pop().replace(/\.[^.]*$/, "");
  if (s.toLowerCase().startsWith("однолинейка")) s = s.slice(11);
  return "КЖ_" + (s.replace(/^[ _-]+/, "") || "щит") + ".dwg"; };
function kjLinks(doc){
  const dwg = doc.kj?.dwg;
  return `<a data-kjdlg="${esc(doc.file)}" title="Какие кабели идут в КЖ: убрать / вернуть (книга не меняется)">кабели КЖ…</a> ` + (dwg
    ? `<a data-kj="${esc(doc.file)}" title="Перестроить КЖ в ${esc(dwg)}: надстройка «КЖ в DWG» в своём Excel, AutoCAD в фоне. Ручные графы 9–13 сохранятся, прежний DWG — *.bak">обновить КЖ</a>
       <a data-kj="${esc(doc.file)}" data-kjask="1" title="Выпустить КЖ в другой DWG (новый — по КЖ_шаблон.dwg)">КЖ в другой DWG…</a>`
    : `<a data-kj="${esc(doc.file)}" data-kjask="1" title="Выпустить кабельный журнал в DWG по КЖ_шаблон.dwg: надстройка «КЖ в DWG» в своём Excel, AutoCAD в фоне">выпустить КЖ…</a>`);
}

export async function runKJ(items, el){
  const old = el?.textContent;
  if (el) el.textContent = "КЖ: выпускаю…";
  try {
    const r = await api.kj(items);
    const ok = (r.done || []).map(d => `${d.doc} → ${d.dwg}`);
    toast((ok.length ? "КЖ готов: " + ok.join("; ") : "") + ((r.errors || []).length ? (ok.length ? ". " : "") + "Не сделано: " + r.errors.join("; ") : ""), 15000);
    if (ok.length) window.dispatchEvent(new CustomEvent("odn:reload"));
  } catch (e){ toast("Не сделано: " + e.message, 15000); }
  finally { if (el && el.isConnected) el.textContent = old; }
}

export function render({ id, params }){
  const p = S.panels[id];
  if (!p) return `<div class="empty">Щит не найден (объект перечитан?).</div>`;
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  const doc = S.P.documents.find(d => d.file === p.doc);
  return `<div class="top"><h2>${esc(p.name)}</h2><span class="sub">${esc(p.doc)}</span>
      ${doc?.external ? `<span class="tag" title="Книга не из папки объекта: программа нашла её по связи">↗ из другой папки, ссылка из ${esc(doc.found_from || "?")}</span>` : ""}
      ${refLink({file: p.doc, sheet: "Однолинейка", cell: "A1"}, "открыть в Excel")}
      ${doc && !(doc.features || []).includes("low_current") ? `<a data-lowpatch="${esc(p.doc)}" title="Книга без столбца слаботочных: запустить 5_Проверка/patch_lowcurrent.py (копия книги — в _окно\\резерв). Книга должна быть закрыта">включить слаботочные в книге</a>` : ""}
      ${doc ? kjLinks(doc) : ""}
      <a data-exclude="${esc(p.doc)}" title="Резервная или лишняя книга: программа перестанет её читать (список — _окно\\исключить.txt)">не читать эту книгу</a></div>
    <div class="kv">${KV.map(f => f(p)).join("")}</div>
    <div class="bar">${TABS.map(t => `<button data-go="${esc(toPanel(p.id, "", t.id))}" class="${t === tab ? "on" : ""}">${esc(t.title(p))}</button>`).join("")}
      <span class="sep"></span>${tab.bar(p)}</div>
    <div class="ghost" id="ghost">${tab.html ? tab.html(p) : ""}</div>`;
}

export function bind(root, rerender, { id, params }){
  issues.bind(root, rerender);
  root.querySelectorAll("[data-lowpatch]").forEach(a => a.onclick = async ev => {
    ev.preventDefault();
    const d = a.dataset.lowpatch;
    if (!confirm(`Включить слаботочные линии в книге «${d}»?\nЗапустится 5_Проверка/patch_lowcurrent.py (своя копия Excel, 10–60 с).\nКнига должна быть закрыта. Копия — в _окно\\резерв.`)) return;
    a.textContent = "включаю…";
    try { const r = await api.bookPatch(d, "lowcurrent");
          toast(r.ok === false ? "Не сделано: " + (r.errors || []).join("; ") : "Слаботочные включены: " + r.msg, 10000);
          window.dispatchEvent(new CustomEvent("odn:reload")); }
    catch (e){ toast("Не сделано: " + e.message, 10000); }
  });
  root.querySelectorAll("[data-kjdlg]").forEach(a => a.onclick = ev => { ev.preventDefault(); kjdlg.open(a.dataset.kjdlg, runKJ); });
  root.querySelectorAll("[data-kj]").forEach(a => a.onclick = ev => {
    ev.preventDefault();
    const d = a.dataset.kj, doc = S.P.documents.find(x => x.file === d);
    let dwg = "";
    if (a.dataset.kjask){
      dwg = prompt(`Кабельный журнал книги «${d}» — путь DWG.\nОтносительный — от папки книги. Есть такой DWG — он обновится (прежний — *.bak).`,
                   doc?.kj?.dwg || kjBase(d));
      if (!dwg) return;
    } else if (!confirm(`Обновить КЖ книги «${d}» в\n${doc?.kj?.dwg}?\nНужен AutoCAD. Если DWG открыт в AutoCAD — правится и сохраняется открытый.`)) return;
    runKJ([{doc: d, dwg}], a);
  });
  const q = root.querySelector("#lq");
  if (q) q.oninput = () => { lf.q = q.value.toLowerCase(); const pos = q.selectionStart; rerender();
    const q2 = document.querySelector("#lq"); q2.focus(); q2.setSelectionRange(pos, pos); };
  const o = root.querySelector("#lonly");
  if (o) o.onchange = () => { lf.only = o.checked; rerender(); };
  const p = S.panels[id];
  if (!p) return;
  root.querySelectorAll("[data-linkdlg]").forEach(b => b.onclick = () => linkdlg.open(p));
  root.querySelectorAll("[data-feederdlg]").forEach(b => b.onclick = () => feederdlg.open(p));
  root.querySelectorAll("[data-vrudlg]").forEach(b => b.onclick = () => vrudlg.open(p));
  root.querySelectorAll("[data-specok]").forEach(b => b.onclick = async () => {
    if (!confirm(`Спецификация книги «${p.doc}» верна?\nПрограмма запомнит аппараты как есть; следующее их изменение снова покажет «устарела».\nСначала сохраните книгу в Excel и нажмите «Обновить».`)) return;
    try { const r = await api.specOk(p.doc); if (r.ok === false) throw new Error((r.errors || []).join("; "));
          window.dispatchEvent(new CustomEvent("odn:reload")); toast("Спецификация отмечена как верная"); }
    catch (e){ toast("Не сделано: " + e.message, 8000); }
  });
  const tab = TABS.find(t => t.id === params.tab) || TABS[0];
  tab.mount(root.querySelector("#ghost"), p, params);
}
