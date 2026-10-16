// Адреса экранов (#/…): кнопка «Назад» браузера работает, ссылку можно сохранить.
export const toObject = (tab = "") => "#/" + (tab ? "?tab=" + tab : "");
export const toPanel = (id, line = "", tab = "") =>
  "#/panel/" + encodeURIComponent(id) + "?" + new URLSearchParams({...(line && {line}), ...(tab && {tab})});
export function parse(){
  const [path, q] = location.hash.slice(1).split("?");
  const parts = (path || "/").split("/").filter(Boolean).map(decodeURIComponent);
  return { view: parts[0] || "object", id: parts[1] || "", params: Object.fromEntries(new URLSearchParams(q || "")) };
}
