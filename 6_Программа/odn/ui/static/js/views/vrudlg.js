// Окно «ВРУ здания»: щит — ввод в здание; потери (ΔU) и длина для надбавки считаются от его шин.
import { esc, toast } from "../fmt.js";
import * as api from "../api.js";
import * as pending from "../pending.js";

export function open(p){
  document.querySelectorAll(".gpop").forEach(x => x.remove());
  const book = p.scheme.is_vru === "да";
  const d = document.createElement("div");
  d.className = "gpop dlg";
  d.innerHTML = `<div class="ph"><b>Ввод в здание: ${esc(p.name)}</b>
      <div class="sub">сейчас: ${p.vru ? "ВРУ здания" : "не ВРУ"}${p.vru_src ? ` (${esc(p.vru_src)})` : ""} · в книге: ${book ? "да" : "нет"}</div></div>
    <div class="sub" style="margin:6px 0">ВРУ здания — начало отсчёта: ΔU линий этого щита и всех щитов ниже считается от его шин
      (3 % свет / 4 %), длина для надбавки 0,005 %/м — тоже от него. Так считают отдельные здания комплекса.</div>
    <div class="row end" style="flex-wrap:wrap">
      <button data-v="1" class="${p.vru ? "on" : ""}">ВРУ здания</button>
      <button data-v="0" class="${!p.vru && p.vru_src === "программа" ? "on" : ""}">Не ВРУ</button>
      <button data-v="" title="Брать из книги («Сеть_щита» строка 26)">Как в книге</button></div>
    <div class="row end"><button data-a="book" title="«Сеть_щита» строка 26 = «да» / пусто — в правки (Ctrl+S)">Записать в книгу</button>
      <button data-a="x">Закрыть</button></div>`;
  document.body.appendChild(d);
  d.querySelector('[data-a="x"]').onclick = () => d.remove();
  d.querySelectorAll("[data-v]").forEach(b => b.onclick = async () => {
    const v = b.dataset.v === "" ? null : b.dataset.v === "1";
    const r = await api.structure({op: "set_vru", panel: p.id, vru: v});
    if (r.ok === false){ toast("Не сделано: " + (r.errors || []).join("; ")); return; }
    d.remove(); toast(v === null ? "ВРУ — как в книге" : `ВРУ здания: ${v ? "да" : "нет"} (программа)`);
    window.dispatchEvent(new CustomEvent("odn:reload"));
  });
  d.querySelector('[data-a="book"]').onclick = () => {
    const want = p.vru ? "да" : "";
    pending.set([{panel: p.id, line: "", path: "scheme.is_vru", new: want, label: `${p.name}: ВРУ здания`, same: want === (book ? "да" : "")}]);
    d.remove();
    toast(want === (book ? "да" : "") ? "В книге уже так" : "В правках: ВРУ здания. Ctrl+S — записать в книгу", 6000);
  };
}
