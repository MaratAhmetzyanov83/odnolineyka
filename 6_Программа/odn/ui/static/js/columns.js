// Столбцы таблиц окна. Новый столбец — элемент массива (см. описание в grid.js).
import { esc, num, refLink } from "./fmt.js";
import { S, lineIn, lineIssues, lineDismissed, lineLevel } from "./store.js";
import { toPanel } from "./nav.js";

const cableShort = c => (c || "").replace(/\s*L=.*$/, "");
const load = l => { const i = l.result.i_calc, n = lineIn(l); return i != null && n ? i / n * 100 : null; };
const duMax = l => l.result.du_max_eff ?? l.result.du_max;
const duSum = l => l.result.du_sum_eff ?? l.result.du_sum;
const duReserve = l => duMax(l) != null && duSum(l) != null ? duMax(l) - duSum(l) : null;
// фидер ≠ строка ввода нижестоящего (calc/feeders.py): подсветка ячейки и значение в другой книге
const mmOf = (l, path) => (l.mismatch || {})[path];
const mmTip = (l, path) => mmOf(l, path) !== undefined ? `в другой книге: ${mmOf(l, path) === "" || mmOf(l, path) == null ? "пусто" : mmOf(l, path)}`
  + (l.feeds ? ` (строка ввода ${S.panels[l.feeds]?.name || ""})` : " (линия-фидер вышестоящего)") : "";
const mm = (path, cls) => ({cls: l => mmOf(l, path) !== undefined ? "mm" : (cls ? cls(l) : ""), tip: l => mmTip(l, path)});
const lvl = (x, warn, err) => x == null ? 0 : x >= err ? 2 : x >= warn ? 1 : 0;

export function lineColumns({ withPanel = false } = {}){
  return [
    ...(withPanel ? [{id: "panel", t: "Щит", w: 80, v: l => S.panels[l.panel]?.name,
      h: l => `<a data-go="${esc(toPanel(l.panel, l.number, "lines"))}">${esc(S.panels[l.panel]?.name)}</a>`}] : []),
    {id: "pl", t: "М.", title: "Место", n: 1, w: 38, d: 0, v: l => l.scheme.place},
    {id: "qf", t: "QF", w: 58, v: l => l.scheme.qf},
    {id: "num", t: "Линия", title: "Номер линии. Правка меняет его и в «Разбивке», и в связях щитов (фидер, G6 нижестоящего)",
      w: 78, v: l => l.number, agg: "count", edit: {path: "number"}},
    {id: "name", t: "Наименование", w: 260, v: l => l.name, nogroup: 1,
      h: l => { const c = l.feeds && S.panels[l.feeds]; return esc(l.name) + (c ? ` → <a data-go="${esc(toPanel(c.id))}">${esc(c.name)}</a>` : ""); },
      edit: {path: "name"}},
    {id: "p", t: "P, кВт", n: 1, w: 62, v: l => l.input.p_kw, agg: "sum", edit: {path: "input.p_kw"}},
    {id: "ph", t: "Фаза", w: 62, v: l => l.scheme.phase},
    {id: "node", t: "Узел", title: "Узел каскада над линией (вкладка «Структура»)", w: 48, v: l => l.node,
      tip: l => l.rcd_node ? `30 мА даёт узел ${l.rcd_node}` : ""},
    {id: "ic", t: "Iр, А", n: 1, w: 58, v: l => l.result.i_calc, agg: "sum"},
    {id: "in", t: "In, А", title: "In автомата (схема, иначе — по току). Правка — In ручной (Исходные данные R); пусто — снова по току",
      n: 1, w: 50, d: 0, v: l => lineIn(l), agg: "max", edit: {path: "input.in_manual", cur: l => l.input.in_manual ?? ""},
      cls: l => lineIn(l) != null && l.result.i_calc != null && lineIn(l) < l.result.i_calc ? "bad" : ""},
    {id: "ld", t: "Загр., %", title: "Загрузка автомата Iр / In", n: 1, w: 70, d: 0, v: load, agg: "max",
      bar: l => { const x = load(l); return x == null ? null : {f: x / 100, lvl: lvl(x, 80, 100)}; }},
    {id: "brk", t: "Автомат", w: 120, v: l => l.scheme.breaker,
      edit: {path: "scheme.breaker_set", ok: l => l.scheme.group_head || !l.scheme.qf, cur: l => l.scheme.breaker_set}},
    {id: "brs", t: "Автомат задан", title: "«авто» или текст (Однолинейка E)", w: 90, hidden: 1, v: l => l.scheme.breaker_set},
    {id: "cab", t: "Кабель", w: 170, v: l => cableShort(l.scheme.cable), tip: l => l.scheme.cable},
    {id: "code", t: "Марка", title: "Код марки (Исходные данные N)", w: 52, v: l => l.input.cable_code, edit: {path: "input.cable_code"},
      ...mm("input.cable_code")},
    {id: "cores", t: "Жил", n: 1, w: 40, d: 0, v: l => l.input.cores, edit: {path: "input.cores"}, ...mm("input.cores")},
    {id: "s", t: "S, мм²", title: "Сечение. Правка — Sмин (Исходные данные Q): не меньше; пусто — по расчёту", n: 1, w: 54, d: 1,
      v: l => l.result.s_mm2, edit: {path: "input.s_min", cur: l => l.input.s_min ?? ""}, ...mm("result.s_mm2")},
    {id: "smin", t: "Sмин", title: "Сечение ручное / мин. по коду, мм²", n: 1, w: 50, d: 1, hidden: 1, v: l => l.input.s_min, edit: {path: "input.s_min"}},
    {id: "inm", t: "In ручн.", title: "In ручной, А", n: 1, w: 56, d: 0, hidden: 1, v: l => l.input.in_manual, edit: {path: "input.in_manual"}},
    {id: "low", t: "Слаб.", title: "Слаботочная линия (DALI, KNX…): без тока, потерь и нагрузки щита. Правка: по коду / да / нет", w: 44,
      v: l => l.input.low_current ? "да" : "", edit: {path: "input.low_current_set", cur: l => l.input.low_current_set}},
    {id: "par", t: "∥", title: "Кабелей параллельно", n: 1, w: 34, d: 0, hidden: 1, v: l => l.input.parallel, edit: {path: "input.parallel"}},
    {id: "lp", t: "Lпл, м", title: "Длина по плану (правится)", n: 1, w: 66, d: 1, v: l => l.input.length_plan, agg: "sum", edit: {path: "input.length_plan"}},
    {id: "len", t: "L, м", title: "Длина итог (с опусками, лотком) — формула книги; править Lпл", n: 1, w: 60, d: 0, v: l => l.input.length, agg: "sum", ...mm("input.length")},
    {id: "tray", t: "Лоток", title: "По лотку, м (Исходные данные J; DLN — с плана)", n: 1, w: 50, d: 0, v: l => l.input.tray, agg: "sum", ...mm("input.tray"), edit: {path: "input.tray"}},
    {id: "pvc", t: "ПВХ", title: "В гофре ПВХ, м (Исходные данные K)", n: 1, w: 46, d: 0, v: l => l.input.pvc, agg: "sum", ...mm("input.pvc"), edit: {path: "input.pvc"}},
    {id: "pnd", t: "ПНД", title: "В гофре ПНД, м (Исходные данные L)", n: 1, w: 46, d: 0, v: l => l.input.pnd, agg: "sum", ...mm("input.pnd"), edit: {path: "input.pnd"}},
    {id: "du", t: "ΔU, %", n: 1, w: 54, v: l => l.result.du, agg: "max"},
    {id: "dus", t: "ΔUΣ, %", title: "ΔU от ВРУ (полоса — доля от допустимой с надбавкой за длину > 100 м)", n: 1, w: 66,
      v: duSum, agg: "max",
      cls: l => duSum(l) != null && duMax(l) != null && duSum(l) > duMax(l) + 1e-9 ? "bad" : "",
      bar: l => duSum(l) == null || !duMax(l) ? null : {f: duSum(l) / duMax(l), lvl: lvl(duSum(l) / duMax(l), 0.8, 1.0001)},
      tip: l => duMax(l) == null ? "" : (l.result.du_from ? `от ВРУ ${S.panels[l.result.du_from]?.name || ""}` + (duSum(l) !== l.result.du_sum ? ` (в книге ${num(l.result.du_sum, 2)} %)` : "") + "; " : "")
        + `допустимо ${num(duMax(l), 2)} %` + (l.result.du_extra
        ? ` = ${num(l.result.du_max, 1)} % + ${num(l.result.du_extra, 3)} % за длину ${num(l.result.du_path_len, 0)} м от ВРУ (> 100 м: 0,005 %/м, не больше 0,5 %)`
        : (l.result.du_path_len != null ? ` (длина от ВРУ ${num(l.result.du_path_len, 0)} м)` : ""))},
    {id: "dul", t: "L от ВРУ", title: "Длина проводки от ВРУ: линия + фидеры выше, м", n: 1, w: 62, d: 0, hidden: 1, v: l => l.result.du_path_len},
    {id: "dur", t: "Запас ΔU", title: "Допустимая ΔU − ΔUΣ, %", n: 1, w: 62, hidden: 1, v: duReserve, agg: "min",
      cls: l => (duReserve(l) ?? 1) < 0 ? "bad" : ""},
    {id: "kz", t: "Iкз, А", title: "Iкз min в конце линии", n: 1, w: 56, d: 0, v: l => l.result.ikz_min, agg: "min"},
    {id: "kzr", t: "Iкз/In", n: 1, w: 54, d: 1, v: l => l.result.ikz_ratio, agg: "min",
      cls: l => (l.result.ikz_check || "").startsWith("мало") ? "bad" : "", tip: l => l.result.ikz_check},
    {id: "cat", t: "Кат.", w: 40, v: l => l.result.category},
    {id: "grp", t: "Кс", title: "Группа для Кс", w: 44, hidden: 1, v: l => l.input.group_ks},
    {id: "bus", t: "Шина", w: 60, hidden: 1, v: l => l.scheme.bus, edit: {path: "scheme.bus"}},
    {id: "term", t: "Клемма", w: 56, hidden: 1, v: l => l.scheme.terminal},
    {id: "st", t: "Статус", title: "Наведите — пояснение: что значит и как исправить", w: 220, v: l => l.result.status, nogroup: 1,
      cls: l => lineLevel(l) >= 2 ? "lv2" : lineLevel(l) ? "lv1" : "",
      tip: l => [l.result.status, ...new Set(lineIssues(l).filter(i => i.hint).map(i => i.hint)),
                 ...lineDismissed(l).map(i => `Снято вручную (${i.dismissed}): ${i.text}`)].filter(Boolean).join("\n\n"),
      // ✕ и «снято» — до текста: длинный статус обрезается многоточием
      h: l => `${refLink(l.refs.scheme || l.refs.input, "✎")}`
        + (lineIssues(l).length ? ` <a class="dis" data-dispanel="${esc(l.panel)}" data-disline="${esc(l.number)}" title="Снять замечания линии вручную (проверено, согласен). Книга не меняется; вернуть — «Замечания» → «Снятые»">✕</a>` : "")
        + (lineDismissed(l).length ? ` <span class="mut" title="Снятые вручную — см. «Замечания» → «Снятые»">снято ${lineDismissed(l).length}</span>` : "")
        + ` ${esc(l.result.status)}`},
  ];
}

export const lineRef = l => ({panel: l.panel, line: l.number, name: S.panels[l.panel]?.name});
export const lineAttrs = l => `class="l${lineLevel(l)}"`;
