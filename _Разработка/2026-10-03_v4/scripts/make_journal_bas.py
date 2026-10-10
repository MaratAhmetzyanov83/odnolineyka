# -*- coding: utf-8 -*-
"""make_journal_bas.py - собрать модуль VBA SX_Journal.bas (макрос «Выгрузить журнал» для книги v4).

    py make_journal_bas.py            -> 3_Макросы_Excel/SX_Journal.bas

Файл целиком в ASCII: русские строки кодируются как ^XXXX и собираются функцией U() при запуске
(тот же приём, что в SX_Groups.bas) — импорт в PERSONAL.XLSB работает при любой кодовой странице.
"""
import os
import re

OUT = os.path.join(os.path.dirname(__file__), "..", "..", "..", "3_Макросы_Excel", "SX_Journal.bas")


def enc(s):
    return "".join(ch if ord(ch) < 128 and ch not in "^~" else "^%04X" % ord(ch) for ch in s)


def u(s):
    return 'U("%s")' % enc(s)


VBA = r'''Attribute VB_Name = "SX_Journal"
Option Explicit
'==========================================================================
' SX_Journal - export of the cable journal of an Odnolineyka v4 workbook
'              into a light values-only file for AutoCAD Data Link.
'              Lives in PERSONAL.XLSB, works on ActiveWorkbook.
'
' Macro:
'   SX_ExportJournal - sheet "Kabelnyi_zhurnal" A2:M... (+ summary O4:Q...)
'                      -> <workbook name>_<journal>.xlsx next to the workbook
'                      (same file every time: the Data Link in AutoCAD stays valid).
' The workbook itself is not changed. Russian text is encoded (^XXXX), see U().
'==========================================================================

Private Function U(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, r As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "^" Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 1, 4)))
            i = i + 5
        Else
            r = r & ch
            i = i + 1
        End If
    Loop
    U = r
End Function

Public Sub SX_ExportJournal()
    Dim wb As Workbook, ws As Worksheet, nb As Workbook, t As Worksheet
    Dim src As Range, sm As Range, path As String, base As String, nrow As Long
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(@SHEET@)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox @NOSHEET@, vbExclamation
        Exit Sub
    End If
    If wb.Path = "" Then
        MsgBox @NOSAVE@, vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo fail
    wb.Application.Calculate
    Set src = ws.Range("A5").CurrentRegion
    Set sm = ws.Range("O4").CurrentRegion
    nrow = src.Rows.Count - 1
    Set nb = Workbooks.Add(xlWBATWorksheet)
    Set t = nb.Worksheets(1)
    t.Name = @T1@
    t.Range("A1").Resize(src.Rows.Count, src.Columns.Count).Value = src.Value
    t.Rows(1).Font.Bold = True
    t.Columns.AutoFit
    Set t = nb.Worksheets.Add(After:=nb.Worksheets(1))
    t.Name = @T2@
    t.Range("A1").Resize(sm.Rows.Count, sm.Columns.Count).Value = sm.Value
    t.Rows(1).Font.Bold = True
    t.Columns.AutoFit
    nb.Worksheets(1).Activate
    base = wb.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    path = wb.Path & Application.PathSeparator & base & @SUFFIX@
    nb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    nb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox @DONE@ & vbCrLf & path & vbCrLf & @ROWS@ & nrow, vbInformation
    Exit Sub
fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not nb Is Nothing Then nb.Close SaveChanges:=False
    MsgBox @ERR@ & Err.Description, vbCritical
End Sub
'''

REPL = {"@SHEET@": u("Кабельный_журнал"),
        "@NOSHEET@": u("В активной книге нет листа «Кабельный_журнал». Откройте книгу однолинейки v4."),
        "@NOSAVE@": u("Сначала сохраните книгу: файл журнала кладётся рядом с ней."),
        "@T1@": u("Журнал"), "@T2@": u("Сводка"), "@SUFFIX@": u("_журнал.xlsx"),
        "@DONE@": u("Журнал выгружен (только значения). В AutoCAD свяжите таблицу (Data Link) с этим файлом:"),
        "@ROWS@": u("Строк: "), "@ERR@": u("Выгрузка не удалась: ")}


def main():
    s = VBA
    for k, v in REPL.items():
        s = s.replace(k, v)
    assert all(ord(c) < 128 for c in s), "в модуле не-ASCII символы"
    with open(OUT, "w", encoding="ascii", newline="\r\n") as f:
        f.write(s)
    print("записан", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
