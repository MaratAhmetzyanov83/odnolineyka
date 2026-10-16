// Запросы к серверу (odn/ui/routes.py).
async function call(method, path, body){
  const r = await fetch(path, method === "GET" ? {} : {method, body: JSON.stringify(body || {})});
  const j = await r.json();
  if (!r.ok) throw new Error(j.error || r.statusText);
  return j;
}
export const getProject = () => call("GET", "/api/project");
export const getOptions = () => call("GET", "/api/options");
export const getJournal = () => call("GET", "/api/journal");
export const reload = () => call("POST", "/api/reload");
export const openRef = ref => call("POST", "/api/open", {ref});
export const edit = changes => call("POST", "/api/edit", {changes});
export const undo = () => call("POST", "/api/undo");

export async function download(path, body){
  const j = await call("POST", path, body);           // сервер готовит файл и даёт ссылку с именем
  const a = Object.assign(document.createElement("a"), {href: j.url});
  document.body.appendChild(a); a.click(); setTimeout(() => a.remove(), 1000);
}
