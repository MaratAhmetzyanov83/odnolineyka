// Несохранённые правки (как изменённые ячейки до «Сохранить» в Excel). Ключ — щит # линия # поле.
const list = new Map();          // key -> {panel, line, path, new, label}
const order = [];                // порядок для Ctrl+Z
const subs = new Set();

export const key = (panel, line, path) => `${panel}#${line}#${path}`;
export const get = (panel, line, path) => list.get(key(panel, line, path));
export const size = () => list.size;
export const all = () => [...list.values()];
export const onChange = fn => subs.add(fn);
const notify = () => subs.forEach(fn => fn());

// set: пачка правок одним действием (для Ctrl+Z — одной отменой)
export function set(items){
  const batch = [];
  for (const it of items){
    const k = key(it.panel, it.line, it.path);
    batch.push([k, list.get(k)]);
    if (it.same) list.delete(k); else list.set(k, it);
  }
  if (batch.length){ order.push(batch); notify(); }
}
export function undo(){
  const batch = order.pop();
  if (!batch) return false;
  for (const [k, prev] of batch.reverse()) prev ? list.set(k, prev) : list.delete(k);
  notify();
  return true;
}
export function clear(){ list.clear(); order.length = 0; notify(); }
