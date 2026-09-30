# -*- coding: utf-8 -*-
"""grsh_rollup.py - сборка категорий, пожарных и ИБП нижестоящих щитов в ГРЩ (v3.12, продолжение model_v312.py).

    py grsh_rollup.py <книга.xlsx> [--out другая.xlsx] [--no-links]     одна книга (все шаги)
    py grsh_rollup.py --folder "<папка объекта>" [--model]              весь объект: сначала щиты, потом книги с «Нижестоящие_щиты»

Порядок: сначала model_v312.py (категории, режимы, ИБП), затем этот скрипт. Идемпотентен (формулы перезаписываются).
Повторный запуск model_v312.py стирает сборку в «Режимы» (он переписывает те же формулы) - после него запустить этот скрипт снова.

Что делает (всё через Excel COM, невидимый экземпляр, DisplayAlerts=False, UpdateLinks=0; строки/столбцы не вставляются):

  1. «Питание_щита» T43:BD47 - экспорт для вышестоящего щита: 37 «составляющих» по щиту 1-3
     (Руст, Руст быт, Рр сил, Руст пост, Σ P·cos φ) для категорий I особая / I / II / III, пожарных,
     нагрузки на вводах (без шин ИБП/ДГУ; обычной и пожарной) и вход ИБП (P и P·cos φ, только для выбранного на «Режимах» щита).
     Слева, в J44:S47, остаются итоги по категориям (Руст, Рр) из model_v312.py.
  2. «Нагрузка_щитов» AJ - секция 3 отдельным значением (было: 3 считалась как 1); результаты секций 1 и 2 не меняются,
     если нет линий секции 3.
  3. Внешние связи: пути приводятся к файлам в той же папке (Excel сохраняет их относительными).
     В исходных книгах v3.11 адреса были записаны в «%D0%9E…»-виде, который Excel не расшифровывает (связь «битая», работал только кэш).
  4. «Нижестоящие_щиты» P:BG - для каждой строки-фидера связи «Вставить связь» с T45:BD45 нижестоящего щита
     (тем же механизмом, что D:J: относительные внешние ссылки с кэшем) и служебные столбцы (щит, источник, категория и пожарная
     линии-фидера, признак сборки).
  5. «Нагрузка_щитов» BB - признак «нижний щит учтён по категориям»; такие линии-фидеры не считаются повторно «одной
     категорией на всю линию» (нет двойного счёта): их мощность берётся по составу нижнего щита.
  6. «Режимы»: служебные суммы по источникам и категориям = собственные линии + нижние щиты; блок «на вводах по щитам» (строки 198-205).
  7. Лист «Категории_по_щитам» (создаётся, если в «Нижестоящие_щиты» есть связанные фидеры): итог по объекту и строка на щит.
"""
import argparse
import glob
import os
import re
import sys
import urllib.parse

import model_v312 as M
from model_v312 import (CAT, ExcelSession, NG, RAWCOLS, R, SOURCES, box, ngr, log, KC, setnf, sheet, style_head,
                        C_GREY, C_HEAD, C_INPUT, C_TAB, LAST)

NH = "Нижестоящие_щиты"
NH_R0, NH_R1 = 5, 34                       # строки фидеров
COMP = ["Руст, кВт", "Руст быт, кВт", "Рр сил (Σ P·Кс), кВт", "Руст пост, кВт", "Σ P·cos φ"]
CSRC = ["C", "D", "E", "F", "G"]           # служебные строки «Режимов»: Руст, быт, Рр сил, пост, Σ P·cos
EXP0 = 20                                  # первый столбец экспорта в «Питание_щита» (T)
NCOL = 37                                  # число столбцов экспорта
# порядок столбцов экспорта (n): 0-19 категории (по 5), 20-24 пожарные, 25-29 на вводах, 30-34 пожарные на вводах, 35-36 вход ИБП
N_CAT, N_FIRE, N_FEED, N_FEEDF, N_UPS = 0, 20, 25, 30, 35
NHC0 = EXP0 + 1                            # первый столбец связей в «Нижестоящие_щиты» (U)
# служебные столбцы «Нижестоящие_щиты»
NH_P, NH_Q, NH_S_, NH_R_, NH_T = 16, 17, 19, 18, 20     # щит, источник, пожарная, категория, признак сборки
NH_MAXC, NH_CHK = NHC0 + NCOL, NHC0 + NCOL + 1          # BF, BG

R_ON = R["svc"] + 54                       # строки «на вводах по щитам» в «Режимы»


def L(n):
    """номер столбца -> буквы"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def svc_rows():
    S = R["svc"]
    d = {"nf": {s: S + 2 + i for i, s in enumerate(SOURCES)}, "f": {s: S + 7 + i for i, s in enumerate(SOURCES)}}
    d["catk"] = {}
    for k in range(3):
        base = S + 19 + k * 11
        d["catk"][k] = dict(head=base, nf={j: base + 1 + i for i, j in enumerate(CAT)}, f={j: base + 5 + i for i, j in enumerate(CAT)},
                            nf_all=base + 9, f_all=base + 10)
    d["on"] = {k: dict(nf=R_ON + 2 + 2 * k, f=R_ON + 3 + 2 * k) for k in range(3)}
    return d


# ------------------------------------------------------------------------------------------------
# формулы-строительные блоки
# ------------------------------------------------------------------------------------------------
def nhr(colnum):
    return "%s!$%s$%d:$%s$%d" % (NH, L(colnum), NH_R0, L(colnum), NH_R1)


def nhd(n):
    return nhr(NHC0 + n)


def own(rc, sel, src=None, fire=None, cat=None, excl=False):
    """SUMIFS по «своим» линиям щита (без линий-фидеров, учтённых по составу нижнего щита: BB=0)"""
    a = [ngr(rc), ngr("C"), sel, ngr("BB"), "0"]
    if src:
        a += [ngr("AU"), '"%s"' % src]
    if excl:
        a += [ngr("AU"), '"<>ИБП"', ngr("AU"), '"<>ДГУ"']
    if fire:
        a += [ngr("AT"), '"%s"' % fire]
    if cat:
        a += [ngr("AS"), '"%s"' % cat]
    return "SUMIFS(%s)" % ",".join(a)


def rolled(n, sel, q=None, qexcl=False, s=None, r=None):
    """SUMIFS по столбцу n связей «Нижестоящие_щиты» для линий-фидеров щита sel"""
    a = [nhd(n), nhr(NH_P), sel, nhr(NH_T), "1"]
    if q:
        a += [nhr(NH_Q), '"%s"' % q]
    if qexcl:
        a += [nhr(NH_Q), '"<>ИБП"', nhr(NH_Q), '"<>ДГУ"']
    if s:
        a += [nhr(NH_S_), '"%s"' % s]
    if r:
        a += [nhr(NH_R_), '"%s"' % r]
    return "SUMIFS(%s)" % ",".join(a)


UPS_COMP = {0: N_UPS, 3: N_UPS, 4: N_UPS + 1}     # компоненты, в которые входит вход ИБП нижнего щита: Руст, Руст пост, Σ P·cos


def f_src_nf(c, sel, src=None, excl=False):
    """составляющая c нагрузки на источнике (без пожарных): свои линии + нижние щиты + вход их ИБП"""
    t = [own(RAWCOLS[c], sel, src=src, fire="нет", excl=excl),
         rolled(N_FEED + c, sel, q=src, qexcl=excl, s="нет")]
    if c in UPS_COMP:
        t.append(rolled(UPS_COMP[c], sel, q=src, qexcl=excl, s="нет"))
    return "+".join(t)


def f_src_f(c, sel, src=None, excl=False):
    """то же, пожарные"""
    t = [own(RAWCOLS[c], sel, src=src, fire="да", excl=excl),
         rolled(N_FEEDF + c, sel, q=src, qexcl=excl),
         rolled(N_FEED + c, sel, q=src, qexcl=excl, s="да")]
    if c in UPS_COMP:
        t.append(rolled(UPS_COMP[c], sel, q=src, qexcl=excl, s="да"))
    return "+".join(t)


def f_cat_nf(c, sel, j):
    return "%s+%s" % (own(RAWCOLS[c], sel, fire="нет", cat=j), rolled(N_CAT + CAT.index(j) * 5 + c, sel, s="нет"))


def f_cat_f(c, sel, j):
    t = [own(RAWCOLS[c], sel, fire="да", cat=j), rolled(N_FIRE + c, sel, r=j)]
    for i in range(4):
        t.append(rolled(N_CAT + i * 5 + c, sel, s="да", r=j))
    return "+".join(t)


# ------------------------------------------------------------------------------------------------
# 1. «Питание_щита»: экспорт для вышестоящего щита
# ------------------------------------------------------------------------------------------------
def export_labels():
    lab = []
    for j in CAT:
        lab += ["%s: %s" % (j, c) for c in COMP]
    lab += ["пожарные: %s" % c for c in COMP]
    lab += ["на вводах: %s" % c for c in COMP]
    lab += ["на вводах, пожарные: %s" % c for c in COMP]
    lab += ["ИБП: вход P, кВт", "ИБП: вход P·cos φ"]
    return lab


def step_export(wb):
    ws = sheet(wb, "Питание_щита")
    sv = svc_rows()
    labs = export_labels()
    ws.Range("T43").Value = ("v3.12: составляющие по категориям, пожарным и ИБП для сборки в ГРЩ "
                             "(Копировать T45:BD45 → «Нижестоящие_щиты» U:BE → Вставить связь; см. 4_Инструкции/Категории_режимы_ИБП.md, раздел 6)")
    ws.Range("T43").Font.Name = "Arial"
    ws.Range("T43").Font.Size = 9
    ws.Range("T43").Font.Bold = True
    for n in range(NCOL):
        c = EXP0 + n
        col = L(c)
        ws.Range("%s44" % col).Value = labs[n]
        ws.Range("%s44" % col).Font.Name = "Arial"
        ws.Range("%s44" % col).Font.Size = 9
        ws.Columns(c).ColumnWidth = 11
    style_head(ws.Range("%s44:%s44" % (L(EXP0), L(EXP0 + NCOL - 1))))
    ws.Rows(44).RowHeight = max(ws.Rows(44).RowHeight, 52)
    for k in range(3):
        r = 45 + k
        ck = sv["catk"][k]
        for n in range(NCOL):
            col = L(EXP0 + n)
            if n < N_FIRE:
                j = CAT[n // 5]
                ref = "Режимы!$%s$%d" % (CSRC[n % 5], ck["nf"][j])
            elif n < N_FEED:
                ref = "Режимы!$%s$%d" % (CSRC[n - N_FIRE], ck["f_all"])
            elif n < N_FEEDF:
                ref = "Режимы!$%s$%d" % (CSRC[n - N_FEED], sv["on"][k]["nf"])
            elif n < N_UPS:
                ref = "Режимы!$%s$%d" % (CSRC[n - N_FEEDF], sv["on"][k]["f"])
            elif n == N_UPS:
                ws.Range("%s%d" % (col, r)).Formula = '=IF($B%d="","",IF(Режимы!$C$7=%d,N(ИБП!$C$46),0))' % (r, k + 1)
                continue
            else:
                ws.Range("%s%d" % (col, r)).Formula = (
                    '=IF($B%d="","",IF(Режимы!$C$7=%d,N(ИБП!$C$46)*IF(N(ИБП!$C$39)>0,ИБП!$C$39,0.95),0))' % (r, k + 1))
                continue
            ws.Range("%s%d" % (col, r)).Formula = '=IF($B%d="","",IFERROR(%s*1,""))' % (r, ref)
    rng = ws.Range("%s45:%s47" % (L(EXP0), L(EXP0 + NCOL - 1)))
    rng.Font.Name = "Arial"
    rng.Font.Size = 10
    rng.HorizontalAlignment = -4108
    setnf(rng, "0.00")
    box(ws.Range("%s44:%s47" % (L(EXP0), L(EXP0 + NCOL - 1))))
    ws.Range("T48").Value = ("ИБП: вход считается на «ИБП» только для щита, выбранного на «Режимы» (C3); у остальных щитов книги 0. "
                             "Составляющие: Руст, Руст быт, Рр сил, Руст пост, Σ P·cos φ; Рр = Руст быт·Кс + Ко·Рр сил + Руст пост считает вышестоящий щит.")
    ws.Range("T48").Font.Size = 8
    ws.Range("T48").Font.Color = C_GREY


# ------------------------------------------------------------------------------------------------
# 2. «Нагрузка_щитов» AJ: секция 3
# ------------------------------------------------------------------------------------------------
AJ_FORMULA = ('=IF(A3="","",IFERROR(IF(INDEX(\'Исходные данные\'!$BB$1:$BB$1000,MATCH(D3,\'Исходные данные\'!$D$1:$D$1000,0))&""="2",2,'
              'IF(INDEX(\'Исходные данные\'!$BB$1:$BB$1000,MATCH(D3,\'Исходные данные\'!$D$1:$D$1000,0))&""="3",3,1)),1))')


def step_section3(wb):
    ws = sheet(wb, NG)
    ws.Range("AJ3:AJ%d" % LAST).Formula = AJ_FORMULA


# ------------------------------------------------------------------------------------------------
# 3. внешние связи: относительные пути в той же папке
# ------------------------------------------------------------------------------------------------
def real_name(src):
    nm = os.path.basename(src.replace("\\", "/"))
    prev = None
    while prev != nm:
        prev, nm = nm, urllib.parse.unquote(nm)
    return nm


def step_links(wb, folder):
    ls = wb.LinkSources(1)
    if not ls:
        return 0, 0
    fixed = 0
    for old in list(ls):
        nm = real_name(old)
        new = os.path.join(folder, nm)
        if os.path.normcase(os.path.normpath(old)) == os.path.normcase(os.path.normpath(new)):
            continue
        if not os.path.exists(new):
            log("  нет файла для связи:", old)
            continue
        wb.ChangeLink(old, new, 1)
        fixed += 1
    return fixed, len(ls)


# ------------------------------------------------------------------------------------------------
# 4-6. Сборка в книге с «Нижестоящие_щиты»
# ------------------------------------------------------------------------------------------------
def nh_link_rows(wsn):
    rows = []
    for r in range(NH_R0, NH_R1 + 1):
        c = wsn.Range("D%d" % r)
        if c.HasFormula and "[" in c.Formula:
            rows.append(r)
    return rows


def step_nh(wb, sep):
    ws = sheet(wb, NH)
    labs = export_labels()
    src_cols = ws.Range("D4")
    rows = nh_link_rows(ws)
    # оформление шапки и данных как у D:J
    ws.Range("D4").Copy(ws.Range("%s4:%s4" % (L(NH_P), L(NH_CHK))))
    ws.Range("D5:D%d" % NH_R1).Copy()
    ws.Range("%s%d:%s%d" % (L(NH_P), NH_R0, L(NH_CHK), NH_R1)).PasteSpecial(-4122)     # только форматы
    ws.Application.CutCopyMode = False
    ws.Range("%s3" % L(NH_P)).Value = "СЛУЖЕБНОЕ (v3.12): как линия-фидер учитывается в режимах"
    ws.Range("%s3" % L(NHC0)).Value = ("ИЗ ФАЙЛА НИЖЕСТОЯЩЕГО (v3.12, Вставить связь из «Питание_щита» T45:BD45): составляющие по категориям, "
                                       "пожарным, нагрузке на вводах и входу ИБП")
    ws.Range("%s3" % L(NH_MAXC)).Value = "ПРОВЕРКА КАТЕГОРИИ ФИДЕРА (v3.12)"
    for a in ("%s3" % L(NH_P), "%s3" % L(NHC0), "%s3" % L(NH_MAXC)):
        ws.Range(a).Font.Bold = True
        ws.Range(a).Font.Size = 9
    heads = {NH_P: "Щит (этой книги)", NH_Q: "Источник для режимов", NH_R_: "Категория линии", NH_S_: "Пожарная (линия)",
             NH_T: "Сборка по составу (1/0)"}
    for c, h in heads.items():
        ws.Range("%s4" % L(c)).Value = h
    for n in range(NCOL):
        ws.Range("%s4" % L(NHC0 + n)).Value = labs[n]
        ws.Range("%s4" % L(NHC0 + n)).Font.Size = 9
    ws.Range("%s4" % L(NH_MAXC)).Value = "Высшая категория нагрузки нижнего щита"
    ws.Range("%s4" % L(NH_CHK)).Value = "Категория фидера ≥ категории нагрузки"
    ws.Range("%s4:%s4" % (L(NH_P), L(NH_CHK))).WrapText = True
    ws.Range("%s4:%s4" % (L(NH_P), L(NH_CHK))).Interior.Color = C_HEAD
    ws.Range("%s4:%s4" % (L(NH_P), L(NH_T))).Font.Color = C_GREY
    for c in range(NH_P, NH_CHK + 1):
        ws.Columns(c).ColumnWidth = 11
    ws.Columns(NH_MAXC).ColumnWidth = 14
    ws.Columns(NH_CHK).ColumnWidth = 22
    # служебные P:T и проверка (формулы по всем строкам 5-34)
    for r in range(NH_R0, NH_R1 + 1):
        m = 'MATCH($B%d,%s!$D$3:$D$%d,0)' % (r, NG, LAST)
        for c, colng in ((NH_P, "C"), (NH_Q, "AU"), (NH_R_, "AS"), (NH_S_, "AT")):
            ws.Range("%s%d" % (L(c), r)).Formula = '=IF($B%d="","",IFERROR(INDEX(%s!$%s$3:$%s$%d,%s)&"",""))' % (r, NG, colng, colng, LAST, m)
        ws.Range("%s%d" % (L(NH_T), r)).Formula = '=IF($B%d="","",IF(AND(ISNUMBER($D%d),SUM($%s%d:$%s%d)>0),1,0))' % (
            r, r, L(NHC0), r, L(NHC0 + N_FEED - 1), r)
        # высшая категория нагрузки по Руст категорий (I особая, I, II, III)
        rc = [L(NHC0 + i * 5) for i in range(4)]
        ws.Range("%s%d" % (L(NH_MAXC), r)).Formula = (
            '=IF($%s%d<>1,"",IF(N($%s%d)>0,"I особая",IF(N($%s%d)>0,"I",IF(N($%s%d)>0,"II",IF(N($%s%d)>0,"III","")))))' % (
                L(NH_T), r, rc[0], r, rc[1], r, rc[2], r, rc[3], r))
        ws.Range("%s%d" % (L(NH_CHK), r)).Formula = (
            '=IF($%s%d="","",IF(IFERROR(MATCH($%s%d,{"III","II","I","I особая"},0),1)>=IFERROR(MATCH($%s%d,{"III","II","I","I особая"},0),1),"ОК",'
            '"фидер "&$%s%d&" ниже нагрузки "&$%s%d))' % (L(NH_MAXC), r, L(NH_R_), r, L(NH_MAXC), r, L(NH_R_), r, L(NH_MAXC), r))
    ws.Range("%s%d:%s%d" % (L(NH_P), NH_R0, L(NH_T), NH_R1)).Font.Color = C_GREY
    ws.Range("%s%d:%s%d" % (L(NH_P), NH_R0, L(NH_CHK), NH_R1)).HorizontalAlignment = -4108
    setnf(ws.Range("%s%d:%s%d" % (L(NHC0), NH_R0, L(NHC0 + NCOL - 1), NH_R1)), "0.00")
    # связи: тем же способом, что D:J (тот же внешний файл, «Питание_щита»!T45:BD45)
    n_ok = 0
    for r in rows:
        f0 = ws.Range("D%d" % r).Formula
        m = re.match(r"^(=.*!)([A-Z]+)(\d+)$", f0)
        if not m:
            log("  строка %d: не разобрана формула связи: %s" % (r, f0[:80]))
            continue
        pref, rownum = m.group(1), m.group(3)
        for n in range(NCOL):
            ws.Range("%s%d" % (L(NHC0 + n), r)).Formula = "%s%s%s" % (pref, L(EXP0 + n), rownum)
        n_ok += 1
    return n_ok, rows


def step_ng_flag(wb):
    ws = sheet(wb, NG)
    vals = ws.Range("BB3:BB%d" % LAST).Formula
    if any(v[0] not in ("", None) and not str(v[0]).startswith("=") for v in vals):
        raise SystemExit("«Нагрузка_щитов» BB занят вводом пользователя")
    c = ws.Range("BB2")
    c.Value = "Нижний щит учтён по категориям (1/0)"
    c.Font.Name = "Calibri"
    c.Font.Size = 11
    c.WrapText = True
    c.HorizontalAlignment = -4108
    box(c)
    ws.Columns("BB").ColumnWidth = 14
    ws.Range("BB3:BB%d" % LAST).Formula = (
        '=IF($A3="","",IFERROR(IF(INDEX(%s!$%s$%d:$%s$%d,MATCH($D3,%s!$B$%d:$B$%d,0))=1,1,0),0))' % (
            NH, L(NH_T), NH_R0, L(NH_T), NH_R1, NH, NH_R0, NH_R1))
    body = ws.Range("BB3:BB%d" % LAST)
    body.Font.Name = "Calibri"
    body.Font.Size = 11
    body.HorizontalAlignment = -4108


def step_modes_rollup(wb):
    ws = sheet(wb, "Режимы")
    sv = svc_rows()
    heads = ["Руст, кВт", "Руст быт", "Рр сил (P·Кс)", "Руст пост", "Σ P·cos φ", "Кс быт", "Рр, кВт", "cos φ", "Qр, квар", "Sр, кВА"]
    # источники: своя нагрузка + нижние щиты
    for s in SOURCES:
        for fire, key, fn in (("нет", "nf", f_src_nf), ("да", "f", f_src_f)):
            r = sv[key][s]
            for c in range(5):
                ws.Range("%s%d" % (CSRC[c], r)).Formula = '=IF($C$4="",0,%s)' % fn(c, "$C$4", src=s)
    # категории
    for k in range(3):
        d = sv["catk"][k]
        hs = "$C$%d" % d["head"]
        for j in CAT:
            for c in range(5):
                ws.Range("%s%d" % (CSRC[c], d["nf"][j])).Formula = '=IF(%s="",0,%s)' % (hs, f_cat_nf(c, hs, j))
                ws.Range("%s%d" % (CSRC[c], d["f"][j])).Formula = '=IF(%s="",0,%s)' % (hs, f_cat_f(c, hs, j))
    # на вводах по щитам (для экспорта вышестоящему)
    S0 = R_ON
    ws.Range("B%d" % S0).Value = "СЛУЖЕБНОЕ: нагрузка щитов 1–3 на вводах (без шин ИБП/ДГУ), со входом ИБП нижних щитов (для «Питание_щита» T45:BD47)"
    ws.Range("B%d" % S0).Font.Bold = True
    for i, h in enumerate(heads[:5]):
        ws.Range("%s%d" % (CSRC[i], S0 + 1)).Value = h
    style_head(ws.Range("B%d:G%d" % (S0 + 1, S0 + 1)))
    ws.Range("B%d" % (S0 + 1)).Value = "Щит / без пожарных / пожарные"
    for k in range(3):
        d = sv["catk"][k]
        hs = "$C$%d" % d["head"]
        for key, fn, fire in (("nf", f_src_nf, "без пожарных"), ("f", f_src_f, "пожарные")):
            r = sv["on"][k][key]
            ws.Range("B%d" % r).Formula = '="Щит %d "&IFERROR($C$%d&"","")&": на вводах, %s"' % (k + 1, d["head"], fire)
            for c in range(5):
                ws.Range("%s%d" % (CSRC[c], r)).Formula = '=IF(%s="",0,%s)' % (hs, fn(c, hs, src=None, excl=True))
    setnf(ws.Range("C%d:G%d" % (S0 + 2, S0 + 7)), "0.00")
    ws.Range("B%d:G%d" % (S0, S0 + 7)).Font.Color = C_GREY
    ws.Range("B%d:G%d" % (S0 + 1, S0 + 1)).Font.Color = 0
    ws.Range("B%d:G%d" % (S0, S0 + 7)).Font.Name = "Arial"
    ws.Range("B%d:G%d" % (S0, S0 + 7)).Font.Size = 10
    # пояснение под таблицей «по категориям»
    t = R["cat"] + 7
    ws.Range("B%d" % t).Value = (
        "Кс быт берётся по сумме Руст быт своего набора линий, поэтому сумма строк по Pр может отличаться от «Итого». Пожарные линии в Pуст, Pр без пожарных не входят. "
        "Нижестоящие щиты (лист «Нижестоящие_щиты», v3.12) входят по составу: категория и пожарность каждой их линии, а не одна категория на линию-фидер; "
        "по щитам — лист «Категории_по_щитам». «Линий» считает линию-фидер по её собственной категории.")


# ------------------------------------------------------------------------------------------------
# 7. Лист «Категории_по_щитам»
# ------------------------------------------------------------------------------------------------
SUMM = "Категории_по_щитам"


def step_summary(wb, feeder_rows):
    ws = sheet(wb, SUMM)
    if ws is None:
        ws = wb.Worksheets.Add(None, wb.Worksheets("ИБП"))
        ws.Name = SUMM
    ws.Tab.Color = C_TAB
    ws.Cells.Clear()
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 44
    for col in "CDEFGHIJ":
        ws.Columns(col).ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 38
    ws.Range("B1").Value = "КАТЕГОРИИ ПО ЩИТАМ (ГРЩ + нижестоящие щиты)"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 12
    ws.Range("B2").Value = ("Итог по объекту — из листа «Режимы» (раздел 2, Кс и Ко считаются на уровне этого щита); таблицы по щитам — по составу нижних щитов "
                            "(«Нижестоящие_щиты» U:BE). Щит выбирается на «Режимы» C3.")
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = C_GREY
    ws.Range("B3").Value = "Щит"
    ws.Range("C3").Formula = "=Режимы!$C$4"
    ws.Range("C3").Font.Bold = True
    sel = "$C$3"
    sv = svc_rows()

    def head(r, labels, c0="B"):
        for i, h in enumerate(labels):
            ws.Range("%s%d" % (chr(ord(c0) + i), r)).Value = h
        rr = ws.Range("%s%d:%s%d" % (c0, r, chr(ord(c0) + len(labels) - 1), r))
        style_head(rr)
        ws.Rows(r).RowHeight = 30

    # --- итог по объекту
    ws.Range("B5").Value = "1. ИТОГ ПО ОБЪЕКТУ (нормальный режим, пожарные отдельно и в максимум не входят)"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 11
    head(6, ["Категория", "Руст, кВт", "Рр, кВт", "Sр, кВА", "Пожарные Руст, кВт", "Пожарные Рр, кВт"])
    rc0 = R["cat"] + 2
    for i, j in enumerate(CAT):
        r = 7 + i
        ws.Range("B%d" % r).Value = j
        for col, src in (("C", "D"), ("D", "E"), ("E", "G"), ("F", "H"), ("G", "I")):
            ws.Range("%s%d" % (col, r)).Formula = "=Режимы!$%s$%d" % (src, rc0 + i)
    ws.Range("B11").Value = "Всего (сумма строк не равна «Итого»: Кс быт по набору)"
    for col in "CDEFG":
        ws.Range("%s11" % col).Formula = "=SUM(%s7:%s10)" % (col, col)
    ws.Range("B12").Value = "Итого по щиту (Режимы, раздел 2)"
    for col, src in (("C", "D"), ("D", "E"), ("E", "G"), ("F", "H"), ("G", "I")):
        ws.Range("%s12" % col).Formula = "=Режимы!$%s$%d" % (src, rc0 + 4)
    ws.Range("B13").Value = "Пожарные вне максимума (кВт)"
    ws.Range("C13").Formula = "=F12"
    ws.Range("D13").Formula = "=G12"
    ws.Range("B14").Value = "ИБП: вход, кВт (в итог не входит; нагрузка ИБП учтена в категориях на выходе)"
    ws.Range("C14").Formula = "=SUM(H18:H%d)" % (19 + NH_R1 - NH_R0)
    ws.Range("B12:G12").Font.Bold = True
    ws.Range("B13:D13").Font.Bold = True
    setnf(ws.Range("C7:G14"), "0.00")
    ws.Range("C7:G14").HorizontalAlignment = -4108
    box(ws.Range("B6:G14"))

    # --- служебный блок: свои линии (справа, N:S)
    ws.Range("N5").Value = "СЛУЖЕБНОЕ: свои линии щита (без нижестоящих)"
    ws.Range("N5").Font.Bold = True
    head(6, ["Категория"] + [c.replace(", кВт", "") for c in COMP], c0="N")
    own_rows = {}
    for i, j in enumerate(CAT + ["пожарные"]):
        r = 7 + i
        own_rows[j] = r
        ws.Range("N%d" % r).Value = j
        for c in range(5):
            col = chr(ord("O") + c)
            if j == "пожарные":
                ws.Range("%s%d" % (col, r)).Formula = "=" + own(RAWCOLS[c], sel, fire="да")
            else:
                ws.Range("%s%d" % (col, r)).Formula = "=" + own(RAWCOLS[c], sel, fire="нет", cat=j)
    setnf(ws.Range("O7:S11"), "0.00")
    ws.Range("N5:S11").Font.Color = C_GREY
    ws.Range("N6:S6").Font.Color = 0
    box(ws.Range("N6:S11"))

    # --- Руст по щитам
    ws.Range("B16").Value = "2. УСТАНОВЛЕННАЯ МОЩНОСТЬ ПО ЩИТАМ, кВт"
    ws.Range("B16").Font.Bold = True
    ws.Range("B16").Font.Size = 11
    head(17, ["Щит / линия-фидер"] + CAT + ["Пожарные", "ИБП: вход", "Итого (без ИБП)", "Как учтено"])
    ws.Range("B18").Value = "Собственные линии щита (без нижестоящих)"
    for i, j in enumerate(CAT):
        ws.Range("%s18" % "CDEF"[i]).Formula = "=O%d" % own_rows[j]
    ws.Range("G18").Formula = "=O%d" % own_rows["пожарные"]
    ws.Range("H18").Formula = "=IF(Режимы!$C$4=%s,N(ИБП!$C$46),0)" % sel
    ws.Range("I18").Formula = "=SUM(C18:G18)"
    ws.Range("J18").Value = "собственные линии"
    for i in range(NH_R1 - NH_R0 + 1):
        r = 19 + i
        nr = NH_R0 + i
        b = "%s!$B$%d" % (NH, nr)
        p, q, rr, s, t = ("%s!$%s$%d" % (NH, L(c), nr) for c in (NH_P, NH_Q, NH_R_, NH_S_, NH_T))
        ws.Range("B%d" % r).Formula = '=IF(OR(%s="",%s<>%s),"",%s)' % (b, p, sel, b)
        for ci, j in enumerate(CAT):
            comp = "%s!$%s$%d" % (NH, L(NHC0 + ci * 5), nr)
            f = ('=IF($B%d="","",IF(%s=1,IF(%s="да",0,%s),IF(AND(%s="нет",%s="%s"),N(%s!$D$%d),0)))' % (
                r, t, s, comp, s, rr, j, NH, nr))
            ws.Range("%s%d" % ("CDEF"[ci], r)).Formula = f
        cats = "+".join("%s!$%s$%d" % (NH, L(NHC0 + ci * 5), nr) for ci in range(4))
        ws.Range("G%d" % r).Formula = ('=IF($B%d="","",IF(%s=1,%s!$%s$%d+IF(%s="да",%s,0),IF(%s="да",N(%s!$D$%d),0)))' % (
            r, t, NH, L(NHC0 + N_FIRE), nr, s, cats, s, NH, nr))
        ws.Range("H%d" % r).Formula = '=IF($B%d="","",N(%s!$%s$%d))' % (r, NH, L(NHC0 + N_UPS), nr)
        ws.Range("I%d" % r).Formula = '=IF($B%d="","",SUM(C%d:G%d))' % (r, r, r)
        ws.Range("J%d" % r).Formula = ('=IF($B%d="","",IF(%s=1,"по составу нижнего щита","одной категорией линии («%s»)"))' % (r, t, "нет данных v3.12"))
    last = 19 + NH_R1 - NH_R0
    rt = last + 1
    ws.Range("B%d" % rt).Value = "Итого по строкам"
    for col in "CDEFGHI":
        ws.Range("%s%d" % (col, rt)).Formula = "=SUM(%s18:%s%d)" % (col, col, last)
    ws.Range("B%d" % (rt + 1)).Value = "По «Режимы» (раздел 2)"
    for col, src in (("C", 0), ("D", 1), ("E", 2), ("F", 3)):
        ws.Range("%s%d" % (col, rt + 1)).Formula = "=Режимы!$D$%d" % (rc0 + src)
    ws.Range("G%d" % (rt + 1)).Formula = "=Режимы!$H$%d" % (rc0 + 4)
    ws.Range("I%d" % (rt + 1)).Formula = "=SUM(C%d:G%d)" % (rt + 1, rt + 1)
    ws.Range("B%d" % (rt + 2)).Value = "Разница (должна быть 0)"
    for col in "CDEFGI":
        ws.Range("%s%d" % (col, rt + 2)).Formula = "=ROUND(%s%d-%s%d,6)" % (col, rt, col, rt + 1)
    setnf(ws.Range("C18:I%d" % (rt + 2)), "0.00")
    ws.Range("C18:I%d" % (rt + 2)).HorizontalAlignment = -4108
    ws.Range("B%d:I%d" % (rt, rt)).Font.Bold = True
    box(ws.Range("B17:J%d" % (rt + 2)))
    ws.Range("B%d" % (rt + 3)).Value = ("«Одной категорией линии»: нижний щит без данных v3.12 (нет связи с T45:BD45) — вся мощность фидера в категории его строки "
                                        "«Исходных данных» BN. Если у линии-фидера BO = да, вся мощность нижнего щита идёт в пожарные.")
    ws.Range("B%d" % (rt + 3)).Font.Size = 8
    ws.Range("B%d" % (rt + 3)).Font.Color = C_GREY

    # --- Рр по щитам (справочно)
    ra = rt + 5
    ws.Range("B%d" % ra).Value = "3. РАСЧЁТНАЯ МОЩНОСТЬ ПО ЩИТАМ, кВт (справочно: у каждого щита свой Кс быт и Ко; сумма строк ≠ Рр ГРЩ)"
    ws.Range("B%d" % ra).Font.Bold = True
    ws.Range("B%d" % ra).Font.Size = 11
    head(ra + 1, ["Щит / линия-фидер"] + CAT + ["Пожарные"])

    def rr_expr(b, sl, po):
        return "IF(%s=0,0,%s*%s)+Ко_сил*%s+%s" % (b, b, KC(b), sl, po)

    r0 = ra + 2
    ws.Range("B%d" % r0).Value = "Собственные линии щита (без нижестоящих)"
    for i, j in enumerate(CAT + ["пожарные"]):
        col = "CDEFG"[i]
        rr_ = own_rows[j]
        ws.Range("%s%d" % (col, r0)).Formula = "=" + rr_expr("$P$%d" % rr_, "$Q$%d" % rr_, "$R$%d" % rr_)
    for i in range(NH_R1 - NH_R0 + 1):
        r = r0 + 1 + i
        nr = NH_R0 + i
        t = "%s!$%s$%d" % (NH, L(NH_T), nr)
        s = "%s!$%s$%d" % (NH, L(NH_S_), nr)
        ws.Range("B%d" % r).Formula = "=B%d" % (19 + i)
        for ci in range(5):
            base = NHC0 + (N_CAT + ci * 5 if ci < 4 else N_FIRE)
            b = "%s!$%s$%d" % (NH, L(base + 1), nr)
            sl = "%s!$%s$%d" % (NH, L(base + 2), nr)
            po = "%s!$%s$%d" % (NH, L(base + 3), nr)
            ru = "%s!$%s$%d" % (NH, L(base), nr)
            f = '=IF(OR($B%d="",%s<>1),"",IF(%s>0,%s,0))' % (r, t, ru, rr_expr(b, sl, po))
            ws.Range("%s%d" % ("CDEFG"[ci], r)).Formula = f
    lr = r0 + NH_R1 - NH_R0 + 1
    ws.Range("B%d" % (lr + 1)).Value = "Сумма строк (арифметически)"
    for col in "CDEFG":
        ws.Range("%s%d" % (col, lr + 1)).Formula = "=SUM(%s%d:%s%d)" % (col, r0, col, lr)
    ws.Range("B%d" % (lr + 2)).Value = "ГРЩ в целом («Режимы», раздел 2, Кс быт и Ко по набору категории)"
    for col, src in (("C", 0), ("D", 1), ("E", 2), ("F", 3)):
        ws.Range("%s%d" % (col, lr + 2)).Formula = "=Режимы!$E$%d" % (rc0 + src)
    ws.Range("G%d" % (lr + 2)).Formula = "=Режимы!$I$%d" % (rc0 + 4)
    setnf(ws.Range("C%d:G%d" % (r0, lr + 2)), "0.00")
    ws.Range("C%d:G%d" % (r0, lr + 2)).HorizontalAlignment = -4108
    ws.Range("B%d:G%d" % (lr + 1, lr + 2)).Font.Bold = True
    box(ws.Range("B%d:G%d" % (ra + 1, lr + 2)))
    # Рр собственных линий: формула по служебному блоку (P=быт, Q=сил, R=пост)
    ws.Range("P6").Value = "Руст быт"
    ws.Range("Q6").Value = "Рр сил"
    ws.Range("R6").Value = "Руст пост"
    ws.Range("S6").Value = "Σ P·cos"
    ws.Rows(6).RowHeight = 30
    ws.Rows(1).RowHeight = 18
    return ws


# ------------------------------------------------------------------------------------------------
def count_errors(ws, addr):
    try:
        e = ws.Range(addr).SpecialCells(-4123, 16)
        return e.Count, e.Address
    except Exception:
        return 0, ""


def run(path, out=None, links=True, rollup=None):
    path = os.path.abspath(path)
    out = os.path.abspath(out) if out else path
    folder = os.path.dirname(out)
    with ExcelSession() as xl:
        wb = xl.Workbooks.Open(path, 0, False)
        need = ["Режимы", "ИБП", "Питание_щита", NG, NH]
        miss = [n for n in need if sheet(wb, n) is None]
        if miss:
            wb.Close(False)
            raise SystemExit("Сначала model_v312.py: нет листов %s" % ", ".join(miss))
        M.DEC = xl.International[2]
        sep = xl.International[4]
        calc0 = xl.Calculation
        try:
            xl.ScreenUpdating = False
            xl.Calculation = -4135
            if links:
                fx, tot = step_links(wb, folder)
                log("связи: исправлено %d из %d" % (fx, tot))
            step_ng_flag(wb)
            log("экспорт «Питание_щита» T:BD...")
            step_export(wb)
            log("секция 3...")
            step_section3(wb)
            wsn = sheet(wb, NH)
            rows = nh_link_rows(wsn)
            if rollup is None:
                rollup = bool(rows)
            log("нижестоящие щиты: связанных фидеров %d" % len(rows))
            n_ok, _ = step_nh(wb, sep)
            log("«Нижестоящие_щиты»: связи U:BE добавлены в %d строках" % n_ok)
            step_modes_rollup(wb)
            if rows:
                log("лист «Категории_по_щитам»...")
                step_summary(wb, rows)
            xl.Calculation = -4105
            xl.CalculateFull()
            bad = 0
            for shn, addr in (("Режимы", "B1:L210"), ("Питание_щита", "T44:BD47"), (NH, "P4:BG34"), (NG, "BB3:BB%d" % LAST),
                              (NG, "AJ3:AJ%d" % LAST), ("Категории_по_щитам", "B1:S100")):
                w = sheet(wb, shn)
                if w is None:
                    continue
                n, a = count_errors(w, addr)
                if n:
                    bad += n
                    log("ОШИБКИ в %s!%s: %d (%s)" % (shn, addr, n, a[:100]))
            log("ошибок формул в новых областях: %d" % bad)
            try:
                wb.Worksheets(1).Activate()
            except Exception:
                pass
            try:
                xl.ScreenUpdating = True
            except Exception:
                pass
            if out == path:
                wb.Save()
            else:
                wb.SaveAs(out, 51)
            if links:
                ls = wb.LinkSources(1)
                log("связей после сохранения: %d" % (len(ls) if ls else 0))
            log("сохранено:", out)
        finally:
            try:
                xl.Calculation = calc0 if calc0 in (-4105, -4135, 2) else -4105
            except Exception:
                pass
            wb.Close(False)
    return out


def refresh_links(path):
    """Обновить значения связей (у щитов - обратные связи K:N из ГРЩ) и сохранить. Нужно после ChangeLink: связь читает
    значения из файла, который на тот момент мог быть ещё без кэша."""
    with ExcelSession() as xl:
        wb = xl.Workbooks.Open(os.path.abspath(path), 0, False)
        try:
            ls = wb.LinkSources(1)
            for src in (ls or []):
                wb.UpdateLink(src, 1)
            xl.CalculateFull()
            wb.Save()
            log("  связи обновлены: %d" % len(ls or []))
        finally:
            wb.Close(False)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--folder", help="папка объекта: все «Однолинейка*.xlsx»")
    ap.add_argument("--model", action="store_true", help="с --folder: сначала model_v312.py для книг без листа «Режимы»")
    ap.add_argument("--no-links", action="store_true", help="не трогать внешние связи")
    ap.add_argument("--refresh-only", action="store_true", help="с --folder: только обновить значения связей у щитов")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if a.folder:
        files = sorted(glob.glob(os.path.join(a.folder, "Однолинейка*.xlsx")))
        tops = [f for f in files if "ГРЩ" in os.path.basename(f)]
        rest = [f for f in files if f not in tops]
        if a.refresh_only:
            for f in rest:
                log(os.path.basename(f))
                refresh_links(f)
            return
        for f in rest + tops:
            log("=== %s" % os.path.basename(f))
            if a.model:
                M.run(f)
            run(f, links=not a.no_links)
        if not a.no_links:
            log("=== обновление значений обратных связей у щитов")
            for f in rest:
                log(os.path.basename(f))
                refresh_links(f)
    elif a.book:
        run(a.book, a.out, links=not a.no_links)
    else:
        ap.error("нужна книга или --folder")


if __name__ == "__main__":
    main()
