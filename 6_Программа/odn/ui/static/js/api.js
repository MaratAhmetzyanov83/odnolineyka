// Запросы к серверу (odn/ui/routes.py).
async function call(method, path, body){
  const r = await fetch(path, method === "GET" ? {} : {method, body: JSON.stringify(body || {})});
  const j = await r.json();
  if (!r.ok || j.ok === false) throw new Error(j.error || r.statusText);
  return j;
}
export const getProject = () => call("GET", "/api/project");
export const reload = () => call("POST", "/api/reload");
export const openRef = ref => call("POST", "/api/open", {ref});
