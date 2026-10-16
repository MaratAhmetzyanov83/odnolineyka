// Формат чисел и текста для экрана.
export const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
export const num = (x, d = 1) => x == null ? "" : (Math.round(x * 10 ** d) / 10 ** d).toLocaleString("ru-RU");
export const refText = r => !r ? "" : (r.sheet ? `${r.sheet}!${r.cell}` : r.file);
// ссылка «открыть в источнике»: ref кладётся в data-ref (JSON), обработчик — в app.js
export const refLink = (r, text) => r && r.file ? `<a data-ref="${esc(JSON.stringify(r))}" title="Открыть в Excel">${esc(text ?? refText(r))}</a>` : esc(text ?? "");
export function toast(t){ const m = document.getElementById("msg"); m.textContent = t; m.style.display = "block";
  clearTimeout(m._t); m._t = setTimeout(() => m.style.display = "none", 4000); }
