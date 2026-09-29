Attribute VB_Name = "SX_Groups"
Option Explicit
'==========================================================================
' SX_Groups  -  "Typical groups" (templates) for the Odnolineyka workbooks
'                (v3.11, xlsx).  Lives in PERSONAL.XLSB, works on ActiveWorkbook.
'
' Macros:
'   SX_InsertGroup  - fill the selected rows of the sheet "Odnolineyka" with a template
'   SX_SaveAsGroup  - save the input values of the selected row as a new template
'
' Rules of the workbook that this module obeys:
'   * only INPUT cells are written (white list of columns below);
'   * a cell that HasFormula is never touched;
'   * blue cells (RGB 157,195,230) and merged cells are never touched;
'   * no rows / columns are added to the project sheets;
'   * protection of a sheet is removed with the password "sx" and restored
'     with the same options; Application state is restored on any error.
'
' Where the inputs really live (checked in v3.11):
'   Odnolineyka (unlocked D,E,F,I,J,K,L,N,P,Y): bus, breaker, 30 mA, devices,
'       module, channel, mode.  Cable, name, power, phase are FORMULAS there
'       (columns Q,R,S,T,V come from "Razbivka_po_shchitam" and "Ishodnye dannye").
'   Ishodnye dannye (keyed by the line number in column D): voltage, phases,
'       cable mark, cores, section, breaker (manual), ceiling/floor, Ks group,
'       Ks, section of the board, PZ category, laying factor.
'   The line of the selected Odnolineyka row = its column R (= Razbivka H).
'
' CYRILLIC.  This file is pure ASCII on purpose.  Every Russian text is stored
' as an escaped string and decoded at run time by U():  "~X" is one letter
' A..ya (ChrW(&H410 + Asc(X) - 48)), "^XXXX" is any other character (ChrW of
' the hex code).  The result does not depend on the Windows code page, on the
' file encoding, or on the editor that saved the file.  A transliterated hint
' follows most of the lines as a comment.
'==========================================================================

Private Const SHEET_PWD As String = "sx"
Private Const FIRST_ROW As Long = 3            ' first data row of the project sheets
Private Const LAST_ROW As Long = 1000          ' last data row
Private Const COL_LINE As Long = 18            ' Odnolineyka!R  = line number (formula)
Private Const COL_RAZB_LINE As Long = 8        ' Razbivka_po_shchitam!H = line number (input)
Private Const TPL_KEY_ROW As Long = 2          ' template sheet: row with keys "Sheet!Col"
Private Const TPL_HDR_ROW As Long = 3          ' template sheet: row with titles
Private Const TPL_FIRST_ROW As Long = 4        ' template sheet: first template
Private Const TPL_VAL_COL As Long = 4          ' template sheet: first value column (D)
Private Const TPL_DV_ROWS As Long = 300        ' rows of the template sheet with drop-downs
Private Const BIG_SELECTION As Long = 100      ' ask before filling more rows than this
Private Const BLUE_COLOR As Long = 15123357    ' RGB(157,195,230) = the blue "do not touch" cells

' white lists of input columns (comma-delimited, both ends)
Private Const WL_ONE As String = ",D,E,F,I,J,K,L,N,Y,"
Private Const WL_GRP As String = ",D,E,F,I,J,K,L,"
Private Const WL_SRC As String = ",F,H,N,O,Q,R,S,AZ,BA,BB,BD,BH,"

Private Type ProtState
    Was As Boolean
    Sh As Worksheet
    DrawObj As Boolean
    Scen As Boolean
    FmtCells As Boolean
    FmtCols As Boolean
    FmtRows As Boolean
    InsCols As Boolean
    InsRows As Boolean
    InsLinks As Boolean
    DelCols As Boolean
    DelRows As Boolean
    Sorting As Boolean
    Filtering As Boolean
    Pivots As Boolean
End Type

' saved Application state / protection
Private g_calc As Long
Private g_upd As Boolean
Private g_evt As Boolean
Private g_saved As Boolean
Private g_pO As ProtState
Private g_pS As ProtState
Private g_warn As String

' column map of the template sheet (filled by LoadColumns)
Private nCols As Long
Private cSheet() As Long        ' 1 = Odnolineyka, 2 = Ishodnye dannye
Private cLetter() As String
Private cNum() As Long          ' column number on the target sheet
Private cTpl() As Long          ' column number on the template sheet

' plan of writes (filled by BuildPlan)
Private nPlan As Long
Private pSh() As Long
Private pTRow() As Long         ' target row
Private pTCol() As Long         ' target column
Private pSelRow() As Long       ' selected row of Odnolineyka this write belongs to
Private pVal() As Variant
Private pOld() As Variant
Private pOldEmpty() As Boolean
Private cntFormula As Long
Private cntSame As Long
Private cntBad As Long
Private cntBlue As Long
Private cntNoLine As Long
Private cntKept As Long
Private cntDV As Long
Private cntW1 As Long
Private cntW2 As Long
Private badText As String
Private badSeen As String
Private dvText As String

'--------------------------------------------------------------------------
' text helpers
'--------------------------------------------------------------------------
Private Function U(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, r As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "~" Then
            r = r & ChrW(&H410 + Asc(Mid$(s, i + 1, 1)) - 48)
            i = i + 2
        ElseIf ch = "^" Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 1, 4)))
            i = i + 5
        Else
            r = r & ch
            i = i + 1
        End If
    Loop
    U = r
End Function

Private Function ttl() As String
    ttl = U("~B~X~_~^~R~k~U ~S~`~c~_~_~k")   ' ru: Tipovye gruppy
End Function

Private Function nmOne() As String
    nmOne = U("~>~T~]~^~[~X~]~U~Y~Z~P")   ' ru: Odnolinejka
End Function

Private Function nmSrc() As String
    nmSrc = U("~8~a~e~^~T~]~k~U ~T~P~]~]~k~U")   ' ru: Ishodnye dannye
End Function

Private Function nmRazb() As String
    nmRazb = U("~@~P~W~Q~X~R~Z~P_~_~^_~i~X~b~P~\")   ' ru: Razbivka_po_schitam
End Function

Private Function nmTpl() As String
    nmTpl = U("~B~X~_~^~R~k~U_~S~`~c~_~_~k")   ' ru: Tipovye_gruppy
End Function

Private Function LS() As String
    LS = Application.International(xlListSeparator)
End Function

Private Function CellText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CellText = CStr(v)
End Function

Private Function IsGroupText(ByVal s As String) As Boolean
    IsGroupText = (LCase$(Left$(Trim$(s), 3)) = U("~^~T~]"))   ' ru: odn
End Function

Private Function IsGroupCol(ByVal letter As String) As Boolean
    IsGroupCol = (InStr(WL_GRP, "," & letter & ",") > 0)
End Function

Private Function IsNum(ByVal v As Variant, ByRef d As Double) As Boolean
    If VarType(v) = vbString Then Exit Function
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        d = CDbl(v)
        IsNum = True
    End If
End Function

Private Function ParseVal(ByVal s As String) As Variant
    ' "2.5" -> Double 2.5 (Val ignores the regional settings), anything else stays text
    If Len(s) > 0 And Not (s Like "*[!0-9.]*") And Left$(s, 1) <> "." And Right$(s, 1) <> "." Then
        ParseVal = Val(s)
    Else
        ParseVal = s
    End If
End Function

Private Sub PutValue(ByVal c As Range, ByVal v As Variant)
    ' text that Excel would take for a formula ("=", "+", "-") is stored with a text prefix
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            Select Case Left$(v, 1)
                Case "=", "+", "-", "@"
                    c.Value2 = "'" & v
                    Exit Sub
            End Select
        End If
    End If
    c.Value2 = v
End Sub

Private Function FindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

'--------------------------------------------------------------------------
' Application state and sheet protection
'--------------------------------------------------------------------------
Private Sub AppSave()
    If g_saved Then Exit Sub
    g_calc = Application.Calculation
    g_upd = Application.ScreenUpdating
    g_evt = Application.EnableEvents
    g_saved = True
End Sub

Private Sub AppQuiet()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub AppRestore()
    If Not g_saved Then Exit Sub
    On Error Resume Next
    Application.EnableEvents = g_evt
    Application.Calculation = g_calc
    Application.ScreenUpdating = g_upd
    g_saved = False
    Err.Clear
End Sub

Private Sub SaveProt(ByRef p As ProtState, ByVal ws As Worksheet)
    Set p.Sh = ws
    p.Was = False
    If ws.ProtectContents Then
        p.Was = True
        p.DrawObj = ws.ProtectDrawingObjects
        p.Scen = ws.ProtectScenarios
        With ws.Protection
            p.FmtCells = .AllowFormattingCells
            p.FmtCols = .AllowFormattingColumns
            p.FmtRows = .AllowFormattingRows
            p.InsCols = .AllowInsertingColumns
            p.InsRows = .AllowInsertingRows
            p.InsLinks = .AllowInsertingHyperlinks
            p.DelCols = .AllowDeletingColumns
            p.DelRows = .AllowDeletingRows
            p.Sorting = .AllowSorting
            p.Filtering = .AllowFiltering
            p.Pivots = .AllowUsingPivotTables
        End With
    End If
End Sub

Private Sub UnprotectSheet(ByRef p As ProtState, ByVal ws As Worksheet)
    Dim failed As Boolean
    SaveProt p, ws
    If Not p.Was Then Exit Sub
    On Error Resume Next
    ws.Unprotect SHEET_PWD
    If Err.Number <> 0 Then failed = True
    Err.Clear
    If ws.ProtectContents Then failed = True
    On Error GoTo 0
    If failed Then
        p.Was = False                   ' we did not remove it, nothing to restore
        Err.Raise vbObjectError + 1001, "SX_Groups", _
            U("~=~U ~c~T~P~[~^~a~l ~a~]~o~b~l ~W~P~i~X~b~c ~[~X~a~b~P ^00AB") & ws.Name & U("^00BB (~_~P~`~^~[~l ~]~U ^00ABsx^00BB?). ~=~X~g~U~S~^ ~]~U ~W~P~_~X~a~P~]~^.")   ' ru: Ne udalos snyat zaschitu lista ? ? (parol ne ?sx??). Nichego
    End If
End Sub

Private Sub ReProtect(ByRef p As ProtState)
    If Not p.Was Then Exit Sub
    p.Was = False
    On Error Resume Next
    If Not p.Sh.ProtectContents Then
        p.Sh.Protect Password:=SHEET_PWD, DrawingObjects:=p.DrawObj, Contents:=True, _
            Scenarios:=p.Scen, AllowFormattingCells:=p.FmtCells, _
            AllowFormattingColumns:=p.FmtCols, AllowFormattingRows:=p.FmtRows, _
            AllowInsertingColumns:=p.InsCols, AllowInsertingRows:=p.InsRows, _
            AllowInsertingHyperlinks:=p.InsLinks, AllowDeletingColumns:=p.DelCols, _
            AllowDeletingRows:=p.DelRows, AllowSorting:=p.Sorting, _
            AllowFiltering:=p.Filtering, AllowUsingPivotTables:=p.Pivots
    End If
    If p.Sh.ProtectContents = False Or Err.Number <> 0 Then
        g_warn = g_warn & vbLf & U("~2~=~8~<~0~=~8~5: ~]~U ~c~T~P~[~^~a~l ~R~U~`~]~c~b~l ~W~P~i~X~b~c ~[~X~a~b~P ^00AB") & p.Sh.Name & U("^00BB. ~2~Z~[~n~g~X~b~U ~R~`~c~g~]~c~n (~_~P~`~^~[~l sx).")   ' ru: VNIMANIE: ne udalos vernut zaschitu lista ? ?. Vklyuchite vr
    End If
    Err.Clear
End Sub

Private Sub Cleanup()
    ReProtect g_pS
    ReProtect g_pO
    AppRestore
End Sub

'--------------------------------------------------------------------------
' Context checks
'--------------------------------------------------------------------------
Private Function CheckContext(ByRef wb As Workbook, ByRef wsO As Worksheet, ByRef selRng As Range) As Boolean
    If Application.Workbooks.Count = 0 Then
        MsgBox U("~=~U~b ~^~b~Z~`~k~b~^~Y ~Z~]~X~S~X."), vbExclamation, ttl()   ' ru: Net otkrytoj knigi.
        Exit Function
    End If
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox U("~=~U~b ~P~Z~b~X~R~]~^~Y ~Z~]~X~S~X."), vbExclamation, ttl()   ' ru: Net aktivnoj knigi.
        Exit Function
    End If
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox U("~0~Z~b~X~R~U~] ~]~U ~[~X~a~b. ~?~U~`~U~Y~T~X~b~U ~]~P ~[~X~a~b ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB ~X ~R~k~T~U~[~X~b~U ~a~b~`~^~Z~X."), vbExclamation, ttl()   ' ru: Aktiven ne list. Perejdite na list ?Odnolinejka? i vydelite 
        Exit Function
    End If
    Set wsO = FindSheet(wb, nmOne())
    If wsO Is Nothing Then
        MsgBox U("~2 ~Z~]~X~S~U ^00AB") & wb.Name & U("^00BB ~]~U~b ~[~X~a~b~P ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB."), vbExclamation, ttl()   ' ru: V knige ? ? net lista ?Odnolinejka?.
        Exit Function
    End If
    If Not ActiveSheet Is wsO Then
        MsgBox U("~0~Z~b~X~R~U~] ~[~X~a~b ^00AB") & ActiveSheet.Name & U("^00BB. ~?~U~`~U~Y~T~X~b~U ~]~P ~[~X~a~b ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB ~X ~R~k~T~U~[~X~b~U ~a~b~`~^~Z~X."), vbExclamation, ttl()   ' ru: Aktiven list ? ?. Perejdite na list ?Odnolinejka? i vydelite
        Exit Function
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox U("~2~k~T~U~[~X~b~U ~o~g~U~Y~Z~X ~a~b~`~^~Z ~]~P ~[~X~a~b~U ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB."), vbExclamation, ttl()   ' ru: Vydelite yachejki strok na liste ?Odnolinejka?.
        Exit Function
    End If
    Set selRng = Selection
    CheckContext = True
End Function

Private Function CollectRows(ByVal selRng As Range, ByVal wsO As Worksheet, ByRef selRows() As Long) As Long
    Dim mark(FIRST_ROW To LAST_ROW) As Boolean
    Dim ar As Range, r1 As Long, r2 As Long, r As Long, n As Long
    For Each ar In selRng.Areas
        r1 = ar.Row
        r2 = ar.Row + ar.Rows.Count - 1
        If r1 < FIRST_ROW Then r1 = FIRST_ROW
        If r2 > LAST_ROW Then r2 = LAST_ROW
        For r = r1 To r2
            If Not mark(r) Then
                If Not wsO.Rows(r).Hidden Then
                    mark(r) = True
                    n = n + 1
                End If
            End If
        Next r
    Next ar
    If n = 0 Then Exit Function
    ReDim selRows(1 To n)
    n = 0
    For r = FIRST_ROW To LAST_ROW
        If mark(r) Then
            n = n + 1
            selRows(n) = r
        End If
    Next r
    CollectRows = n
End Function

'--------------------------------------------------------------------------
' Template sheet
'--------------------------------------------------------------------------
Private Function ColDefs() As Variant
    ' "sheet|column|title"   sheet: O = Odnolineyka, S = Ishodnye dannye
    Dim a(1 To 21) As String
    a(1) = U("O|D|~H~X~]~P (~_~c~a~b~^ = ~^~a~]~^~R~]~P~o)")   ' ru: O|D|Shina (pusto = osnovnaya)
    a(2) = U("O|E|~0~R~b~^~\~P~b: ~P~R~b~^ ~X~[~X ~b~U~Z~a~b")   ' ru: O|E|Avtomat: avto ili tekst
    a(3) = U("O|F|30 ~\~0: ~T~P / ~]~U~b")   ' ru: O|F|30 mA: da / net
    a(4) = U("O|I|~0~_~_~P~`~P~b 1")   ' ru: O|I|Apparat 1
    a(5) = U("O|J|~0~_~_~P~`~P~b 2")   ' ru: O|J|Apparat 2
    a(6) = U("O|K|~0~_~_~P~`~P~b 3")   ' ru: O|K|Apparat 3
    a(7) = U("O|L|~<~^~T~c~[~l")   ' ru: O|L|Modul
    a(8) = U("O|N|~:~P~]~P~[")   ' ru: O|N|Kanal
    a(9) = U("O|Y|~@~U~V~X~\")   ' ru: O|Y|Rezhim
    a(10) = U("S|F|~=~P~_~`~o~V~U~]~X~U, ~2")   ' ru: S|F|Napryazhenie, V
    a(11) = U("S|H|~:~^~[-~R~^ ~d~P~W")   ' ru: S|H|Kol-vo faz
    a(12) = U("S|N|~<~P~`~Z~P ~Z~P~Q~U~[~o")   ' ru: S|N|Marka kabelya
    a(13) = U("S|O|~:~^~[-~R~^ ~V~X~[")   ' ru: S|O|Kol-vo zhil
    a(14) = U("S|Q|~A~U~g~U~]~X~U (~`~c~g~]~^~U), ~\~\^00B2")   ' ru: S|Q|Sechenie (ruchnoe), mm?
    a(15) = U("S|R|~0~R~b~^~\~P~b (~`~c~g~]~^~U), ~0")   ' ru: S|R|Avtomat (ruchnoe), A
    a(16) = U("S|S|~?~^~b~^~[~^~Z 1 / ~?~^~[ 2")   ' ru: S|S|Potolok 1 / Pol 2
    a(17) = U("S|AZ|~3~`~c~_~_~P ~:~a")   ' ru: S|AZ|Gruppa Ks
    a(18) = U("S|BA|~:~a (~T~[~o ~A~X~[)")   ' ru: S|BA|Ks (dlya Sil)
    a(19) = U("S|BB|~A~U~Z~f~X~o")   ' ru: S|BB|Sekciya
    a(20) = U("S|BD|~:~P~b~U~S~^~`~X~o ~?~7")   ' ru: S|BD|Kategoriya PZ
    a(21) = U("S|BH|~: ~_~`~^~Z~[~P~T~Z~X ~[~X~]~X~X")   ' ru: S|BH|K prokladki linii
    ColDefs = a
End Function

Private Function StarterList() As Variant
    ' "name|description|mode|sheet:col=value|..."   (O = Odnolineyka, S = Ishodnye dannye)
    Dim a(1 To 12) As String
    a(1) = U("~@~^~W~U~b~Z~X 1~d (R)|~@~^~W~U~b~^~g~]~P~o ~S~`~c~_~_~P: ~P~R~b~^~\~P~b ~P~R~b~^, 230 ~2, ~2~2~3~]~S(~0)-LS 3~e2,5, ~_~^~[, ~Q~k~b. ~?~`~X~\~U~` ~>~a~b~`~^~R ~X ~Z~]~X~S~X ~6~4 (~Z~^~T R).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=2.5|S:S=2|S:AZ=~1~k~b|S:BD=~@~^~W~U~b~^~g~]~P~o ~a~U~b~l ~X ~Q~k~b~^~R~k~U ~_~`~X~Q~^~`~k")   ' ru: Rozetki 1f (R)|Rozetochnaya gruppa: avtomat avto, 230 V, VVG
    a(2) = U("~@~^~W~U~b~Z~P 3~d 400 ~2 (5 ~V~X~[)|~A~X~[~^~R~P~o ~`~^~W~U~b~Z~P: 3 ~d~P~W~k, 5 ~V~X~[ 2,5 ~\~\^00B2, 400 ~2. ~:~P~Z R.1.03.4 ~R ~_~`~X~\~U~`~U ~>~a~b~`~^~R.|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=400|S:H=3|S:N=~\~a|S:O=5|S:Q=2.5|S:S=2|S:AZ=~1~k~b|S:BD=~@~^~W~U~b~^~g~]~P~o ~a~U~b~l ~X ~Q~k~b~^~R~k~U ~_~`~X~Q~^~`~k")   ' ru: Rozetka 3f 400 V (5 zhil)|Silovaya rozetka: 3 fazy, 5 zhil 2
    a(3) = U("~A~R~U~b 1~d 3~e1,5 (L)|~>~a~R~U~i~U~]~X~U: ~a~R~^~Y ~P~R~b~^~\~P~b ~]~P ~[~X~]~X~n, 230 ~2, 3~e1,5. ~?~`~X~\~U~` ~>~a~b~`~^~R (~Z~^~T L).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=1.5|S:AZ=~1~k~b|S:BD=~@~P~Q~^~g~U~U ~^~a~R~U~i~U~]~X~U")   ' ru: Svet 1f 3h1,5 (L)|Osveschenie: svoj avtomat na liniyu, 230 V
    a(4) = U("~A~R~U~b + ~<~:-5-1, 5~e1,5 (~S~`~c~_~_~P)|~2~k~T~U~[~X~b~l ~R~a~U ~[~X~]~X~X ~S~`~c~_~_~k: ~P~R~b~^~\~P~b ~X ~<~:-5-1 ~_~^~_~P~T~c~b ~R ~_~U~`~R~c~n ~a~b~`~^~Z~c, ~Z~P~Q~U~[~l 5~e1,5 - ~R ~Z~P~V~T~c~n. ~?~`~X~\~U~` ~>~a~b~`~^~R.|~^~T~]~P ~S~`~c~_~_~P|O:E=~P~R~b~^|O:I=~<~:-5-1|S:F=230|S:H=1|S:N=~\~a|S:O=5|S:Q=1.5|S:AZ=~1~k~b|S:BD=~@~P~Q~^~g~U~U ~^~a~R~U~i~U~]~X~U")   ' ru: Svet + MK-5-1, 5h1,5 (gruppa)|Vydelit vse linii gruppy: avto
    a(5) = U("~A~R~U~b ~]~P ~\~^~T~c~[~U ZIOMB24V2 (~S~`~c~_~_~P)|~2~k~T~U~[~X~b~l ~R~a~U ~[~X~]~X~X ~S~`~c~_~_~k: ~P~R~b~^~\~P~b ~X ~\~^~T~c~[~l - ~R ~_~U~`~R~c~n ~a~b~`~^~Z~c, 3~e1,5 - ~R ~Z~P~V~T~c~n; ~Z~P~]~P~[~k ~]~P~W~]~P~g~P~n~b~a~o ~a~P~\~X. ~?~`~X~\~U~` ~>~a~b~`~^~R.|~^~T~]~P ~S~`~c~_~_~P|O:E=~P~R~b~^|O:L=ZIOMB24V2|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=1.5|S:AZ=~1~k~b|S:BD=~@~P~Q~^~g~U~U ~^~a~R~U~i~U~]~X~U")   ' ru: Svet na module ZIOMB24V2 (gruppa)|Vydelit vse linii gruppy: 
    a(6) = U("~:~^~]~T~X~f~X~^~]~U~` (K)|~2~]~c~b~`~U~]~]~X~Y ~X~[~X ~R~]~U~h~]~X~Y ~Q~[~^~Z: ~P~R~b~^, 230 ~2, 1,5 ~\~\^00B2, ~P~R~b~^~\~P~b 10 ~0, ~_~^~b~^~[~^~Z, ~a~X~[~^~R~P~o ~:~a 0,5. ~:~]~X~S~X ~6~4 (~Z~^~T K, 53 ~[~X~]~X~X).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=1.5|S:R=10|S:S=1|S:AZ=~A~X~[|S:BA=0.5|S:BD=~:~^~]~T~X~f~X~^~]~X~`~^~R~P~]~X~U ~X ~c~R~[~P~V~]~U~]~X~U")   ' ru: Kondicioner (K)|Vnutrennij ili vneshnij blok: avto, 230 V, 1
    a(7) = U("~2~U~]~b~X~[~o~f~X~o (~2, ~?, ~?~2)|~2~k~b~o~V~]~k~U ~X ~_~`~X~b~^~g~]~k~U ~c~a~b~P~]~^~R~Z~X: 230 ~2, 1,5 ~\~\^00B2, ~P~R~b~^~\~P~b 10 ~0, ~_~^~b~^~[~^~Z, ~a~X~[~^~R~P~o ~:~a 0,5. ~:~]~X~S~X ~6~4 (~2, ~?, ~?~2 - 33 ~[~X~]~X~X).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=1.5|S:R=10|S:S=1|S:AZ=~A~X~[|S:BA=0.5|S:BD=~2~U~]~b~X~[~o~f~X~o")   ' ru: Ventilyaciya (V, P, PV)|Vytyazhnye i pritochnye ustanovki: 2
    a(8) = U("~=~P~a~^~a (~=~2~:)|~=~P~a~^~a~k ~R~^~T~^~a~]~P~Q~V~U~]~X~o ~X ~T~`~U~]~P~V~P: 230 ~2, 2,5 ~\~\^00B2, ~P~R~b~^~\~P~b 16 ~0, ~_~^~b~^~[~^~Z, ~:~a 0,5 (~R ~I~A-~2~: ~Q~k~R~P~U~b 0,7). ~:~]~X~S~X ~6~4 (~=~2~: - 13 ~[~X~]~X~Y).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=2.5|S:R=16|S:S=1|S:AZ=~A~X~[|S:BA=0.5|S:BD=~=~P~a~^~a~k (~R~^~T~^~a~]~P~Q~V~U~]~X~U, ~T~`~U~]~P~V)")   ' ru: Nasos (NVK)|Nasosy vodosnabzheniya i drenazha: 230 V, 2,5 mm
    a(9) = U("~>~Q~^~S~`~U~R ~Z~`~^~R~[~X ~X ~R~^~T~^~a~b~^~Z~^~R (OV)|~3~`~U~n~i~X~Y ~Z~P~Q~U~[~l: 230 ~2, 2,5 ~\~\^00B2, ~P~R~b~^~\~P~b 16 ~0, ~_~^~b~^~[~^~Z, ~]~P~S~`~c~W~Z~P ~_~^~a~b~^~o~]~]~P~o (~?~^~a~b). ~:~]~X~S~X ~6~4 (OV - 9 ~[~X~]~X~Y).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:F=230|S:H=1|S:N=~\~a|S:O=3|S:Q=2.5|S:R=16|S:S=1|S:AZ=~?~^~a~b|S:BD=~>~Q~^~S~`~U~R ~R~^~T~^~a~b~^~Z~^~R ~X ~Z~`~^~R~[~X")   ' ru: Obogrev krovli i vodostokov (OV)|Greyuschij kabel: 230 V, 2,
    a(10) = U("~2~R~^~T / ~]~U ~`~X~a~^~R~P~b~l (~R~`~c~g~]~c~n)|~2~R~^~T~k, ~b~`~P~]~W~X~b, ~`~U~W~U~`~R: ~`~U~V~X~\ ~R~`~c~g~]~c~n, ~R ~`~P~a~g~q~b ~:~a ~]~U ~R~e~^~T~o~b (~=~U~b). ~:~]~X~S~X ~6~4, ~>~a~b~`~^~R, ~A~Z~^~[~Z~^~R~^.|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:Y=~R~`~c~g~]~c~n|S:AZ=~=~U~b")   ' ru: Vvod / ne risovat (vruchnuyu)|Vvody, tranzit, rezerv: rezhim
    a(11) = U("~?~`~X~R~^~T / ~Z~`~P~] (~M~:, SH)|~?~`~X~R~^~T~k ~h~b~^~` ~X ~R~^~`~^~b, ~m~[~U~Z~b~`~^~Z~`~P~]~k: ~a~R~^~Y ~P~R~b~^~\~P~b, ~Z~P~b~U~S~^~`~X~o ~?~7 - ~_~`~X~R~^~T~k. ~?~^ ~a~_~`~P~R~^~g~]~X~Z~c ~X ~A~Z~^~[~Z~^~R~^, ~T~P~]~]~k~e ~\~P~[~^ - ~Z~P~Q~U~[~l ~_~`~^~R~U~`~l~b~U.|~Z~P~V~T~P~o ~a~b~`~^~Z~P|O:E=~P~R~b~^|S:BD=~?~`~X~R~^~T~k (~h~b~^~`~k, ~R~^~`~^~b~P, ~Z~`~P~]~k)")   ' ru: Privod / kran (EK, SH)|Privody shtor i vorot, elektrokrany: 
    a(12) = U("~4~P~b~g~X~Z ~_~`~^~b~U~g~Z~X (WD)|~4~P~b~g~X~Z~X ~_~`~^~b~U~g~Z~X: ~Q~U~W ~a~R~^~U~S~^ ~P~R~b~^~\~P~b~P, ~S~`~c~_~_~P ~?~^~a~b. ~?~^ ~a~_~`~P~R~^~g~]~X~Z~c, ~R ~Z~]~X~S~P~e ~[~X~]~X~Y ~]~U~b. ~<~^~T~c~[~l ~c~Z~P~V~X~b~U ~a~P~\~X (~a~b~^~[~Q~U~f L).|~Z~P~V~T~P~o ~a~b~`~^~Z~P|S:AZ=~?~^~a~b|S:BD=~A~[~P~Q~^~b~^~g~]~k~U ~a~X~a~b~U~\~k, ~P~R~b~^~\~P~b~X~W~P~f~X~o, KNX")   ' ru: Datchik protechki (WD)|Datchiki protechki: bez svoego avtoma
    StarterList = a
End Function

Private Function DvFor(ByVal code As String, ByVal letter As String) As String
    Dim l As String
    l = LS()
    If code = "O" Then
        Select Case letter
            Case "E": DvFor = "=" & U("~A~_~X~a~^~Z_~P~R~b~^~\~P~b~^~R")   ' ru: Spisok_avtomatov
            Case "F": DvFor = U("~T~P") & l & U("~]~U~b")   ' ru: da net
            Case "I", "J", "K": DvFor = "=" & U("~A~_~X~a~^~Z_~P~_~_~P~`~P~b~^~R")   ' ru: Spisok_apparatov
            Case "L": DvFor = "=" & U("~A~_~X~a~^~Z_~\~^~T~c~[~U~Y")   ' ru: Spisok_modulej
            Case "N": DvFor = "+" & l & "=" & l & U("~`") & l & U("~Q/~Z")   ' ru: r b/k
            Case "Y": DvFor = U("~R~`~c~g~]~c~n")   ' ru: vruchnuyu
        End Select
    Else
        Select Case letter
            Case "H": DvFor = "1" & l & "3"
            Case "S", "BB": DvFor = "1" & l & "2"
            Case "AZ": DvFor = U("~1~k~b") & l & U("~A~X~[") & l & U("~?~^~a~b") & l & U("~=~U~b")   ' ru: Byt Sil Post Net
            Case "BD": DvFor = "=" & U("~:~P~b~U~S~^~`~X~X_~?~7")   ' ru: Kategorii_PZ
        End Select
    End If
End Function

Private Sub AddListDV(ByVal rng As Range, ByVal f As String)
    If Len(f) = 0 Then Exit Sub
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
        Operator:=xlBetween, Formula1:=f
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    Err.Clear
End Sub

Private Sub FormatTplRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastC As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC))
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Cells(r, 1).WrapText = True
    ws.Cells(r, 2).WrapText = True
End Sub

Private Function CreateTplSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, defs As Variant, st As Variant
    Dim i As Long, k As Long, p As Long, c As Long, r As Long, lastC As Long
    Dim parts() As String, f() As String, tok As String, p1 As Long, p2 As Long
    Dim code As String, letter As String, col As Long

    AppSave
    AppQuiet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nmTpl()
    On Error Resume Next
    ws.Tab.Color = RGB(255, 192, 0)
    On Error GoTo 0

    defs = ColDefs()
    lastC = TPL_VAL_COL + UBound(defs) - LBound(defs)

    ws.Cells(1, 1).Value2 = U("~B~8~?~>~2~K~5 ~3~@~C~?~?~K ~T~[~o ~\~P~Z~`~^~a~^~R SX_InsertGroup / SX_SaveAsGroup. ~>~T~]~P ~a~b~`~^~Z~P - ~^~T~X~] ~h~P~Q~[~^~]. ~?~c~a~b~P~o ~o~g~U~Y~Z~P = ~]~U ~b~`~^~S~P~b~l. ~A~b~`~^~Z~P 2 (~a~U~`.) - ~a~[~c~V~U~Q~]~P~o: ^00AB~[~X~a~b!~a~b~^~[~Q~U~f^00BB, ~]~U ~\~U~]~o~b~l. ^00AB~^~T~]~P ~S~`~c~_~_~P^00BB - ~P~R~b~^~\~P~b, ~P~_~_~P~`~P~b~k, ~\~^~T~c~[~l ~X ~h~X~]~P ~_~X~h~c~b~a~o ~b~^~[~l~Z~^ ~R ~_~U~`~R~c~n ~R~k~T~U~[~U~]~]~c~n ~a~b~`~^~Z~c.")   ' ru: TIPOVYE GRUPPY dlya makrosov SX_InsertGroup / SX_SaveAsGroup
    ws.Cells(2, 1).Value2 = U("~a~[~c~V~U~Q~]~P~o ~a~b~`~^~Z~P: ~[~X~a~b!~a~b~^~[~Q~U~f")   ' ru: sluzhebnaya stroka: list!stolbec
    ws.Cells(TPL_HDR_ROW, 1).Value2 = U("~8~\~o ~h~P~Q~[~^~]~P")   ' ru: Imya shablona
    ws.Cells(TPL_HDR_ROW, 2).Value2 = U("~>~_~X~a~P~]~X~U")   ' ru: Opisanie
    ws.Cells(TPL_HDR_ROW, 3).Value2 = U("~@~U~V~X~\ ~R~k~T~U~[~U~]~X~o")   ' ru: Rezhim vydeleniya

    For i = LBound(defs) To UBound(defs)
        f = Split(CStr(defs(i)), "|")
        c = TPL_VAL_COL + i - LBound(defs)
        If f(0) = "O" Then
            ws.Cells(TPL_KEY_ROW, c).Value2 = nmOne() & "!" & f(1)
        Else
            ws.Cells(TPL_KEY_ROW, c).Value2 = nmSrc() & "!" & f(1)
        End If
        ws.Cells(TPL_HDR_ROW, c).Value2 = f(2) & vbLf & "(" & f(1) & ")"
    Next i

    ws.Cells(1, 1).Font.Bold = True
    With ws.Range(ws.Cells(TPL_KEY_ROW, 1), ws.Cells(TPL_KEY_ROW, lastC))
        .Font.Italic = True
        .Font.Size = 8
        .Font.Color = RGB(128, 128, 128)
        .Interior.Color = RGB(242, 242, 242)
    End With
    With ws.Range(ws.Cells(TPL_HDR_ROW, 1), ws.Cells(TPL_HDR_ROW, lastC))
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(221, 235, 247)
    End With
    ' yellow = sheet "Odnolineyka", green = sheet "Ishodnye dannye"
    For i = LBound(defs) To UBound(defs)
        f = Split(CStr(defs(i)), "|")
        c = TPL_VAL_COL + i - LBound(defs)
        If f(0) = "O" Then
            ws.Cells(TPL_HDR_ROW, c).Interior.Color = RGB(255, 242, 204)
        Else
            ws.Cells(TPL_HDR_ROW, c).Interior.Color = RGB(226, 239, 218)
        End If
    Next i
    ws.Rows(TPL_HDR_ROW).RowHeight = 45
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 60
    ws.Columns(3).ColumnWidth = 15
    For c = TPL_VAL_COL To lastC
        ws.Columns(c).ColumnWidth = 14
    Next c

    st = StarterList()
    For i = LBound(st) To UBound(st)
        r = TPL_FIRST_ROW + i - LBound(st)
        parts = Split(CStr(st(i)), "|")
        PutValue ws.Cells(r, 1), parts(0)
        PutValue ws.Cells(r, 2), parts(1)
        PutValue ws.Cells(r, 3), parts(2)
        For k = 3 To UBound(parts)
            tok = parts(k)
            p1 = InStr(tok, ":")
            p2 = InStr(tok, "=")
            code = Left$(tok, p1 - 1)
            letter = Mid$(tok, p1 + 1, p2 - p1 - 1)
            col = 0
            For p = LBound(defs) To UBound(defs)
                f = Split(CStr(defs(p)), "|")
                If f(0) = code And f(1) = letter Then col = TPL_VAL_COL + p - LBound(defs)
            Next p
            If col > 0 Then PutValue ws.Cells(r, col), ParseVal(Mid$(tok, p2 + 1))
        Next k
        FormatTplRow ws, r, lastC
    Next i

    ' drop-downs for typing new templates by hand (warning only, never blocks)
    AddListDV ws.Range(ws.Cells(TPL_FIRST_ROW, 3), ws.Cells(TPL_FIRST_ROW + TPL_DV_ROWS, 3)), _
        U("~Z~P~V~T~P~o ~a~b~`~^~Z~P") & LS() & U("~^~T~]~P ~S~`~c~_~_~P")   ' ru: kazhdaya stroka odna gruppa
    For i = LBound(defs) To UBound(defs)
        f = Split(CStr(defs(i)), "|")
        c = TPL_VAL_COL + i - LBound(defs)
        AddListDV ws.Range(ws.Cells(TPL_FIRST_ROW, c), ws.Cells(TPL_FIRST_ROW + TPL_DV_ROWS, c)), DvFor(f(0), f(1))
    Next i

    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Cells(TPL_FIRST_ROW, TPL_VAL_COL).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    Err.Clear
    On Error GoTo 0
    Set CreateTplSheet = ws
End Function

Private Function GetTplSheet(ByVal wb As Workbook, ByRef created As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = FindSheet(wb, nmTpl())
    If ws Is Nothing Then
        Set ws = CreateTplSheet(wb)
        created = True
    End If
    Set GetTplSheet = ws
End Function

Private Function LoadColumns(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByRef bad As String) As Long
    Dim lastC As Long, c As Long, key As String, p As Long, sn As String, lt As String, sh As Long, cap As Long
    nCols = 0
    lastC = wsT.Cells(TPL_KEY_ROW, wsT.Columns.Count).End(xlToLeft).Column
    If lastC < TPL_VAL_COL Then Exit Function
    cap = lastC - TPL_VAL_COL + 1
    ReDim cSheet(1 To cap)
    ReDim cLetter(1 To cap)
    ReDim cNum(1 To cap)
    ReDim cTpl(1 To cap)
    For c = TPL_VAL_COL To lastC
        key = Trim$(CellText(wsT.Cells(TPL_KEY_ROW, c)))
        sh = 0
        p = InStr(key, "!")
        If p > 1 Then
            sn = Trim$(Left$(key, p - 1))
            lt = UCase$(Trim$(Mid$(key, p + 1)))
            If Len(lt) >= 1 And Len(lt) <= 2 And InStr(lt, ",") = 0 Then
                If StrComp(sn, nmOne(), vbTextCompare) = 0 Then
                    If InStr(WL_ONE, "," & lt & ",") > 0 Then sh = 1
                ElseIf StrComp(sn, nmSrc(), vbTextCompare) = 0 Then
                    If InStr(WL_SRC, "," & lt & ",") > 0 Then sh = 2
                End If
            End If
        End If
        If sh = 0 Then
            If Len(key) > 0 Then bad = bad & key & "; "
        ElseIf sh = 2 And wsS Is Nothing Then
            bad = bad & key & " (" & U("~]~U~b ~[~X~a~b~P") & "); "   ' ru: net lista
        Else
            nCols = nCols + 1
            cSheet(nCols) = sh
            cLetter(nCols) = lt
            cTpl(nCols) = c
            If sh = 1 Then
                cNum(nCols) = wsO.Range(lt & "1").Column
            Else
                cNum(nCols) = wsS.Range(lt & "1").Column
            End If
        End If
    Next c
    LoadColumns = nCols
End Function

Private Function LoadTemplates(ByVal wsT As Worksheet, ByRef tName() As String, ByRef tDesc() As String, _
                               ByRef tGroup() As Boolean, ByRef tRow() As Long) As Long
    Dim lastR As Long, r As Long, n As Long, s As String
    lastR = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    If lastR < TPL_FIRST_ROW Then Exit Function
    ReDim tName(1 To lastR - TPL_FIRST_ROW + 1)
    ReDim tDesc(1 To lastR - TPL_FIRST_ROW + 1)
    ReDim tGroup(1 To lastR - TPL_FIRST_ROW + 1)
    ReDim tRow(1 To lastR - TPL_FIRST_ROW + 1)
    For r = TPL_FIRST_ROW To lastR
        s = Trim$(CellText(wsT.Cells(r, 1)))
        If Len(s) > 0 Then
            n = n + 1
            tName(n) = s
            tDesc(n) = CellText(wsT.Cells(r, 2))
            tGroup(n) = IsGroupText(CellText(wsT.Cells(r, 3)))
            tRow(n) = r
        End If
    Next r
    LoadTemplates = n
End Function

Private Function ReadInput(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByRef v As Variant) As Boolean
    ' True if the cell holds a plain (non-formula, non-empty) value
    Dim cel As Range
    Set cel = ws.Cells(r, c)
    If cel.HasFormula Then Exit Function
    v = cel.Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        v = Trim$(CStr(v))
        If Len(v) = 0 Then Exit Function
    End If
    ReadInput = True
End Function

'--------------------------------------------------------------------------
' Line number -> row on "Ishodnye dannye"
'--------------------------------------------------------------------------
Private Function LineOfRow(ByVal wsO As Worksheet, ByVal wsR As Worksheet, ByVal r As Long) As String
    Dim v As Variant
    v = wsO.Cells(r, COL_LINE).Value2
    If Not IsError(v) Then
        If Not IsEmpty(v) Then LineOfRow = Trim$(CStr(v))
    End If
    If Len(LineOfRow) = 0 And Not wsR Is Nothing Then
        v = wsR.Cells(r, COL_RAZB_LINE).Value2
        If Not IsError(v) Then
            If Not IsEmpty(v) Then LineOfRow = Trim$(CStr(v))
        End If
    End If
End Function

Private Function BuildLineIndex(ByVal wsS As Worksheet) As Collection
    Dim a As Variant, i As Long, k As String, col As Collection
    Set col = New Collection
    a = wsS.Range("D" & FIRST_ROW & ":D" & LAST_ROW).Value2
    For i = 1 To UBound(a, 1)
        k = ""
        If Not IsError(a(i, 1)) Then
            If Not IsEmpty(a(i, 1)) Then k = Trim$(CStr(a(i, 1)))
        End If
        If Len(k) > 0 Then
            On Error Resume Next
            col.Add i + FIRST_ROW - 1, k       ' duplicate key -> first line wins
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    Set BuildLineIndex = col
End Function

Private Function LookupLine(ByVal idx As Collection, ByVal key As String) As Long
    Dim n As Variant
    On Error Resume Next
    n = idx(key)
    If Err.Number = 0 Then LookupLine = CLng(n)
    Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------
' Checks of a value / of a target cell
'--------------------------------------------------------------------------
Private Function InNamedFirstCol(ByVal wb As Workbook, ByVal nm As String, ByVal v As Variant) As Boolean
    ' True when v is in the first column of the named range (or the name does not exist: no check)
    Dim rg As Range, m As Variant
    InNamedFirstCol = True
    On Error Resume Next
    Set rg = wb.Names(nm).RefersToRange
    Err.Clear
    On Error GoTo 0
    If rg Is Nothing Then Exit Function
    m = Application.Match(v, rg.Columns(1), 0)
    InNamedFirstCol = Not IsError(m)
End Function

Private Function ExtraOK(ByVal wb As Workbook, ByVal sh As Long, ByVal letter As String, _
                         ByVal v As Variant, ByRef why As String) As Boolean
    ' checks for the columns that have no drop-down of their own
    Dim d As Double
    ExtraOK = True
    why = ""
    If sh <> 2 Then Exit Function
    Select Case letter
        Case "F", "BH"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d <= 0 Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~]~c~V~]~^ ~g~X~a~[~^ ~Q~^~[~l~h~U 0")   ' ru: nuzhno chislo bolshe 0
        Case "H"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d <> 1 And d <> 3 Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~d~P~W: 1 ~X~[~X 3")   ' ru: faz: 1 ili 3
        Case "S", "BB"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d <> 1 And d <> 2 Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~]~c~V~]~^ 1 ~X~[~X 2")   ' ru: nuzhno 1 ili 2
        Case "O"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d < 1 Or d > 10 Or d <> Int(d) Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~V~X~[: ~f~U~[~^~U 1-10")   ' ru: zhil: celoe 1-10
        Case "R"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d < 0 Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~]~c~V~]~^ ~g~X~a~[~^ (~0)")   ' ru: nuzhno chislo (A)
        Case "BA"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d <= 0 Or d > 1 Then
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~:~a: ~g~X~a~[~^ 0-1")   ' ru: Ks: chislo 0-1
        Case "Q"
            If Not IsNum(v, d) Then
                ExtraOK = False
            ElseIf d <= 0 Then
                ExtraOK = False
            ElseIf Not InNamedFirstCol(wb, U("I~T~^~__S"), v) Then   ' ru: Idop_S
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~a~U~g~U~]~X~o ~]~U~b ~R ~b~P~Q~[~X~f~U ^00AB~A~_~`~P~R~^~g~]~P~o^00BB")   ' ru: secheniya net v tablice ?Spravochnaya?
        Case "N"
            If VarType(v) <> vbString Then
                ExtraOK = False
            ElseIf Not InNamedFirstCol(wb, U("~:~P~Q~U~[~l_~^~Q~^~W~]~P~g~U~]~X~U"), v) Then   ' ru: Kabel_oboznachenie
                ExtraOK = False
            End If
            If Not ExtraOK Then why = U("~]~U~b ~b~P~Z~^~Y ~\~P~`~Z~X ~R ^00AB~A~_~`~P~R~^~g~]~P~o^00BB")   ' ru: net takoj marki v ?Spravochnaya?
    End Select
End Function

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) = vbString And VarType(b) = vbString Then
        SameValue = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
    ElseIf VarType(a) <> vbString And VarType(b) <> vbString Then
        If IsNumeric(a) And IsNumeric(b) Then SameValue = (CDbl(a) = CDbl(b))
    End If
End Function

Private Function Classify(ByVal wb As Workbook, ByVal tgt As Range, ByVal v As Variant, _
                          ByVal sh As Long, ByVal letter As String, _
                          ByRef oldV As Variant, ByRef why As String) As Long
    ' 0 = empty target (write), 1 = filled target (ask), 2 = formula, 3 = already equal,
    ' 4 = template value invalid, 5 = blue / merged cell
    Dim cur As Variant
    why = ""
    oldV = Empty
    If tgt.MergeCells Then
        Classify = 5
        Exit Function
    End If
    If tgt.HasFormula Then
        Classify = 2
        Exit Function
    End If
    If tgt.Interior.Color = BLUE_COLOR Then
        Classify = 5
        Exit Function
    End If
    If Not ExtraOK(wb, sh, letter, v, why) Then
        Classify = 4
        Exit Function
    End If
    cur = tgt.Value2
    oldV = cur
    If IsError(cur) Then
        Classify = 1
    ElseIf IsEmpty(cur) Then
        Classify = 0
    ElseIf VarType(cur) = vbString Then
        If Len(cur) = 0 Then
            Classify = 0
        ElseIf SameValue(cur, v) Then
            Classify = 3
        Else
            Classify = 1
        End If
    ElseIf SameValue(cur, v) Then
        Classify = 3
    Else
        Classify = 1
    End If
End Function

Private Function PassesValidation(ByVal c As Range) As Boolean
    ' Excel's own verdict (drop-down lists incl. dynamic ones); no validation = passes
    Dim t As Long
    PassesValidation = True
    On Error Resume Next
    t = c.Validation.Type
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    PassesValidation = CBool(c.Validation.Value)
    If Err.Number <> 0 Then
        Err.Clear
        PassesValidation = True
    End If
End Function

'--------------------------------------------------------------------------
' Plan and apply
'--------------------------------------------------------------------------
Private Sub ResetPlan(ByVal cap As Long)
    nPlan = 0
    ReDim pSh(1 To cap)
    ReDim pTRow(1 To cap)
    ReDim pTCol(1 To cap)
    ReDim pSelRow(1 To cap)
    ReDim pVal(1 To cap)
    ReDim pOld(1 To cap)
    ReDim pOldEmpty(1 To cap)
    cntFormula = 0
    cntSame = 0
    cntBad = 0
    cntBlue = 0
    cntNoLine = 0
    cntKept = 0
    cntDV = 0
    cntW1 = 0
    cntW2 = 0
    badText = ""
    badSeen = ""
    dvText = ""
End Sub

Private Sub BuildPlan(ByVal wb As Workbook, ByVal wsO As Worksheet, ByVal wsS As Worksheet, _
                      ByVal wsT As Worksheet, ByVal tRow As Long, ByVal grpMode As Boolean, _
                      ByRef selRows() As Long, ByVal nRows As Long)
    Dim j As Long, i As Long, r As Long, firstRow As Long, srcRow As Long, st As Long
    Dim vals() As Variant, has() As Boolean
    Dim idx As Collection, lineKey As String, needSrc As Boolean
    Dim wsR As Worksheet, tgt As Range, oldV As Variant, why As String, sName As String

    ResetPlan nRows * nCols
    ReDim vals(1 To nCols)
    ReDim has(1 To nCols)
    For j = 1 To nCols
        has(j) = ReadInput(wsT, tRow, cTpl(j), vals(j))
        If has(j) Then
            If cSheet(j) = 2 Then needSrc = True
        End If
    Next j

    Set wsR = FindSheet(wb, nmRazb())
    If needSrc And Not wsS Is Nothing Then Set idx = BuildLineIndex(wsS)
    firstRow = selRows(1)

    For i = 1 To nRows
        r = selRows(i)
        srcRow = 0
        If needSrc Then
            lineKey = LineOfRow(wsO, wsR, r)
            If Len(lineKey) > 0 And Not idx Is Nothing Then srcRow = LookupLine(idx, lineKey)
            If srcRow = 0 Then cntNoLine = cntNoLine + 1
        End If
        For j = 1 To nCols
            If has(j) Then
                Set tgt = Nothing
                If cSheet(j) = 1 Then
                    If Not (grpMode And IsGroupCol(cLetter(j)) And r <> firstRow) Then
                        Set tgt = wsO.Cells(r, cNum(j))
                    End If
                ElseIf srcRow > 0 Then
                    Set tgt = wsS.Cells(srcRow, cNum(j))
                End If
                If Not tgt Is Nothing Then
                    st = Classify(wb, tgt, vals(j), cSheet(j), cLetter(j), oldV, why)
                    Select Case st
                        Case 0, 1
                            nPlan = nPlan + 1
                            pSh(nPlan) = cSheet(j)
                            pTRow(nPlan) = tgt.Row
                            pTCol(nPlan) = tgt.Column
                            pSelRow(nPlan) = r
                            pVal(nPlan) = vals(j)
                            pOld(nPlan) = oldV
                            pOldEmpty(nPlan) = (st = 0)
                        Case 2
                            cntFormula = cntFormula + 1
                        Case 3
                            cntSame = cntSame + 1
                        Case 4
                            cntBad = cntBad + 1
                            If InStr(badSeen, "|" & cSheet(j) & cLetter(j) & "|") = 0 Then
                                badSeen = badSeen & "|" & cSheet(j) & cLetter(j) & "|"
                                If cSheet(j) = 1 Then sName = nmOne() Else sName = nmSrc()
                                If Len(badText) < 400 Then
                                    badText = badText & vbLf & "  " & sName & "!" & cLetter(j) & " = " & CStr(vals(j)) & ": " & why
                                End If
                            End If
                        Case 5
                            cntBlue = cntBlue + 1
                    End Select
                End If
            End If
        Next j
    Next i
End Sub

Private Function PlanConflicts() As Long
    Dim k As Long, n As Long
    For k = 1 To nPlan
        If Not pOldEmpty(k) Then n = n + 1
    Next k
    PlanConflicts = n
End Function

Private Function ApplyPlan(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal overwrite As Boolean) As Long
    ' returns the number of selected rows that received at least one value
    Dim k As Long, tgt As Range, hit(FIRST_ROW To LAST_ROW) As Boolean, n As Long, ws As Worksheet
    For k = 1 To nPlan
        If pOldEmpty(k) Or overwrite Then
            If pSh(k) = 1 Then Set ws = wsO Else Set ws = wsS
            Set tgt = ws.Cells(pTRow(k), pTCol(k))
            PutValue tgt, pVal(k)
            If PassesValidation(tgt) Then
                If pSh(k) = 1 Then cntW1 = cntW1 + 1 Else cntW2 = cntW2 + 1
                If Not hit(pSelRow(k)) Then
                    hit(pSelRow(k)) = True
                    n = n + 1
                End If
            Else
                ' not in the drop-down: put the old content back
                If pOldEmpty(k) Or IsEmpty(pOld(k)) Then
                    tgt.ClearContents
                ElseIf IsError(pOld(k)) Then
                    ' cannot restore an error value (never a formula: those are skipped)
                    tgt.ClearContents
                Else
                    PutValue tgt, pOld(k)
                End If
                cntDV = cntDV + 1
                If Len(dvText) < 300 Then dvText = dvText & vbLf & "  " & tgt.Address(False, False) & " = " & CStr(pVal(k))
            End If
        Else
            cntKept = cntKept + 1
        End If
    Next k
    ApplyPlan = n
End Function

'--------------------------------------------------------------------------
' Pick a template with InputBox
'--------------------------------------------------------------------------
Private Function PickTemplate(ByVal n As Long, ByRef tName() As String) As Long
    Dim i As Long, p As String, s As String, k As Long, cnt As Long, hit As Long, more As Long
    p = U("~2~R~U~T~X~b~U ~]~^~\~U~` ~h~P~Q~[~^~]~P (~X~[~X ~g~P~a~b~l ~]~P~W~R~P~]~X~o):") & vbLf & vbLf   ' ru: Vvedite nomer shablona (ili chast nazvaniya):
    For i = 1 To n
        s = i & ". " & tName(i) & vbLf
        If Len(p) + Len(s) > 850 Then
            more = n - i + 1
            Exit For
        End If
        p = p & s
    Next i
    If more > 0 Then p = p & U("... ~U~i~q ") & more & U(" - ~_~^~[~]~k~Y ~a~_~X~a~^~Z ~]~P ~[~X~a~b~U ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB")   ' ru: ... esche   - polnyj spisok na liste ?Tipovye_gruppy?
    s = Trim$(InputBox(p, ttl(), "1"))
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then
        k = CLng(Val(s))
        If k >= 1 And k <= n Then
            PickTemplate = k
        Else
            MsgBox U("~=~U~b ~h~P~Q~[~^~]~P ~a ~]~^~\~U~`~^~\ ") & s & ".", vbExclamation, ttl()   ' ru: Net shablona s nomerom 
        End If
        Exit Function
    End If
    For i = 1 To n
        If StrComp(tName(i), s, vbTextCompare) = 0 Then
            PickTemplate = i
            Exit Function
        End If
    Next i
    For i = 1 To n
        If InStr(1, tName(i), s, vbTextCompare) > 0 Then
            cnt = cnt + 1
            hit = i
        End If
    Next i
    If cnt = 1 Then
        PickTemplate = hit
    ElseIf cnt = 0 Then
        MsgBox U("~H~P~Q~[~^~] ^00AB") & s & U("^00BB ~]~U ~]~P~Y~T~U~]."), vbExclamation, ttl()   ' ru: Shablon ? ? ne najden.
    Else
        MsgBox U("^00AB") & s & U("^00BB ~_~^~T~e~^~T~X~b ~Z ~]~U~a~Z~^~[~l~Z~X~\ ~h~P~Q~[~^~]~P~\ (") & cnt & U("). ~2~R~U~T~X~b~U ~]~^~\~U~`."), vbExclamation, ttl()   ' ru: ? ? podhodit k neskolkim shablonam ( ). Vvedite nomer.
    End If
End Function

'==========================================================================
' PUBLIC MACRO 1: fill selected rows with a template
'==========================================================================
Public Sub SX_InsertGroup()
    Dim wb As Workbook, wsO As Worksheet, wsT As Worksheet, wsS As Worksheet
    Dim selRng As Range, created As Boolean, badKeys As String
    Dim selRows() As Long, nRows As Long
    Dim nTpl As Long, tName() As String, tDesc() As String, tGroup() As Boolean, tRow() As Long
    Dim pick As Long, rowsDone As Long, overwrite As Boolean, nConf As Long
    Dim finalMsg As String, msgIcon As Long, ans As VbMsgBoxResult
    Dim needS As Boolean, k As Long

    On Error GoTo EH
    g_warn = ""
    msgIcon = vbInformation
    If Not CheckContext(wb, wsO, selRng) Then GoTo Fin

    Set wsT = GetTplSheet(wb, created)
    If created Then
        wsO.Activate                       ' Worksheets.Add activates the new sheet
        AppRestore
        MsgBox U("~2 ~Z~]~X~S~U ~]~U ~Q~k~[~^ ~[~X~a~b~P ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB - ~a~^~W~T~P~] ~a~^ ~a~b~P~`~b~^~R~k~\~X ~h~P~Q~[~^~]~P~\~X. ~A~R~^~X ~h~P~Q~[~^~]~k ~T~^~Q~P~R~[~o~Y~b~U ~a~b~`~^~Z~P~\~X ~R~]~X~W.") _   ' ru: V knige ne bylo lista ?Tipovye_gruppy? - sozdan so startovym
            & vbLf & vbLf & U("~2~k~T~U~[~U~]~X~U ~]~P ~[~X~a~b~U ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB ~a~^~e~`~P~]~U~]~^, ~R~k~Q~U~`~X~b~U ~h~P~Q~[~^~]."), vbInformation, ttl()   ' ru: Vydelenie na liste ?Odnolinejka? sohraneno, vyberite shablon
    End If
    Set wsS = FindSheet(wb, nmSrc())

    If LoadColumns(wsT, wsO, wsS, badKeys) = 0 Then
        finalMsg = U("~=~P ~[~X~a~b~U ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB ~]~U~b ~`~P~a~_~^~W~]~P~]~]~k~e ~a~b~^~[~Q~f~^~R (~a~b~`~^~Z~P 2: ^00AB~[~X~a~b!~a~b~^~[~Q~U~f^00BB).")   ' ru: Na liste ?Tipovye_gruppy? net raspoznannyh stolbcov (stroka 
        msgIcon = vbExclamation
        GoTo Fin
    End If
    nTpl = LoadTemplates(wsT, tName, tDesc, tGroup, tRow)
    If nTpl = 0 Then
        finalMsg = U("~=~P ~[~X~a~b~U ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB ~]~U~b ~h~P~Q~[~^~]~^~R (~a 4-~Y ~a~b~`~^~Z~X, ~X~\~o ~R ~a~b~^~[~Q~f~U A).")   ' ru: Na liste ?Tipovye_gruppy? net shablonov (s 4-j stroki, imya 
        msgIcon = vbExclamation
        GoTo Fin
    End If

    pick = PickTemplate(nTpl, tName)
    If pick = 0 Then GoTo Fin

    nRows = CollectRows(selRng, wsO, selRows)
    If nRows = 0 Then
        finalMsg = U("~2~k~T~U~[~X~b~U ~a~b~`~^~Z~X ~T~P~]~]~k~e ~[~X~a~b~P ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB (~a~b~`~^~Z~X 3-1000, ~]~U ~a~Z~`~k~b~k~U).")   ' ru: Vydelite stroki dannyh lista ?Odnolinejka? (stroki 3-1000, n
        msgIcon = vbExclamation
        GoTo Fin
    End If
    If nRows > BIG_SELECTION Then
        ans = MsgBox(U("~2~k~T~U~[~U~]~^ ~a~b~`~^~Z: ") & nRows & U(". ~7~P~_~^~[~]~X~b~l ~h~P~Q~[~^~]~^~\ ^00AB") & tName(pick) & U("^00BB ~R~a~U?"), _   ' ru: Vydeleno strok:  . Zapolnit shablonom ? ? vse?
                     vbYesNo + vbQuestion + vbDefaultButton2, ttl())
        If ans <> vbYes Then GoTo Fin
    End If

    ' ---- plan (reads only, nothing is written yet) ----
    BuildPlan wb, wsO, wsS, wsT, tRow(pick), tGroup(pick), selRows, nRows
    If nPlan = 0 Then
        finalMsg = U("~7~P~_~^~[~]~U~]~^ 0 ~a~b~`~^~Z ~h~P~Q~[~^~]~^~\ ") & tName(pick) & "." & vbLf & vbLf & U("~=~U~g~U~S~^ ~W~P~_~X~a~k~R~P~b~l.") & SkipReport(badKeys)   ' ru: Zapolneno 0 strok shablonom  Nechego zapisyvat.
        GoTo Fin
    End If

    overwrite = False
    nConf = PlanConflicts()
    If nConf > 0 Then
        ans = MsgBox(U("~O~g~U~U~Z ~R~R~^~T~P, ~S~T~U ~c~V~U ~g~b~^-~b~^ ~U~a~b~l: ") & nConf & U(" (~X~W ") & nPlan & U(").") & vbLf & vbLf & _   ' ru: Yacheek vvoda, gde uzhe chto-to est:   (iz 
                     U("~?~U~`~U~W~P~_~X~a~P~b~l?") & vbLf & _   ' ru: Perezapisat?
                     U("~4~P - ~_~U~`~U~W~P~_~X~a~P~b~l; ~=~U~b - ~W~P~_~^~[~]~X~b~l ~b~^~[~l~Z~^ ~_~c~a~b~k~U; ~>~b~\~U~]~P - ~R~k~Y~b~X."), _   ' ru: Da - perezapisat; Net - zapolnit tolko pustye; Otmena - vyjt
                     vbYesNoCancel + vbQuestion + vbDefaultButton2, ttl())
        If ans = vbCancel Then GoTo Fin
        overwrite = (ans = vbYes)
    End If

    ' ---- write ----
    AppSave
    AppQuiet
    UnprotectSheet g_pO, wsO
    For k = 1 To nPlan
        If pSh(k) = 2 Then needS = True
    Next k
    If needS Then UnprotectSheet g_pS, wsS
    rowsDone = ApplyPlan(wsO, wsS, overwrite)

    finalMsg = U("~7~P~_~^~[~]~U~]~^ ") & rowsDone & U(" ~a~b~`~^~Z ~h~P~Q~[~^~]~^~\ ") & tName(pick) & "." & vbLf & vbLf & _   ' ru: Zapolneno   strok shablonom 
               U("~7~P~_~X~a~P~]~^ ~o~g~U~U~Z: ") & (cntW1 + cntW2) & U(" (~>~T~]~^~[~X~]~U~Y~Z~P: ") & cntW1 & U("; ~8~a~e~^~T~]~k~U ~T~P~]~]~k~U: ") & cntW2 & ")." & _   ' ru: Zapisano yacheek:   (Odnolinejka:  ; Ishodnye dannye: 
               SkipReport(badKeys)

Fin:
    On Error Resume Next
    Cleanup
    If Len(g_warn) > 0 Then
        finalMsg = finalMsg & vbLf & g_warn
        msgIcon = vbExclamation
    End If
    If Len(finalMsg) > 1000 Then finalMsg = Left$(finalMsg, 995) & "..."
    If Len(finalMsg) > 0 Then MsgBox finalMsg, msgIcon, ttl()
    Exit Sub
EH:
    finalMsg = U("~>~h~X~Q~Z~P ") & Err.Number & ": " & Err.Description & vbLf & vbLf & _   ' ru: Oshibka 
               U("~7~P~i~X~b~P ~[~X~a~b~^~R ~X ~]~P~a~b~`~^~Y~Z~X Excel ~R~^~a~a~b~P~]~^~R~[~U~]~k. ~?~`~^~R~U~`~l~b~U ~W~P~_~^~[~]~U~]~]~^~U.")   ' ru: Zaschita listov i nastrojki Excel vosstanovleny. Proverte za
    msgIcon = vbCritical
    Resume Fin
End Sub

Private Function SkipReport(ByVal badKeys As String) As String
    Dim s As String
    If cntFormula > 0 Then s = s & vbLf & U("~?~`~^~_~c~i~U~]~^ ~o~g~U~U~Z ~a ~d~^~`~\~c~[~P~\~X (~]~U ~b~`~^~S~P~U~\): ") & cntFormula   ' ru: Propuscheno yacheek s formulami (ne trogaem): 
    If cntKept > 0 Then s = s & vbLf & U("~>~a~b~P~R~[~U~]~^ ~Z~P~Z ~Q~k~[~^ (~c~V~U ~W~P~_~^~[~]~U~]~k): ") & cntKept   ' ru: Ostavleno kak bylo (uzhe zapolneny): 
    If cntSame > 0 Then s = s & vbLf & U("~C~V~U ~a~^~R~_~P~T~P~n~b ~a ~h~P~Q~[~^~]~^~\: ") & cntSame   ' ru: Uzhe sovpadayut s shablonom: 
    If cntBlue > 0 Then s = s & vbLf & U("~?~`~^~_~c~i~U~]~^ ~a~X~]~X~e/~^~Q~j~U~T~X~]~q~]~]~k~e ~o~g~U~U~Z: ") & cntBlue   ' ru: Propuscheno sinih/obedinennyh yacheek: 
    If cntNoLine > 0 Then s = s & vbLf & U("~A~b~`~^~Z ~Q~U~W ~]~^~\~U~`~P ~[~X~]~X~X (~X~[~X ~U~S~^ ~]~U~b ~R ^00AB~8~a~e~^~T~]~k~e ~T~P~]~]~k~e^00BB): ") & cntNoLine & U(" - ~W~]~P~g~U~]~X~o ~T~[~o ^00AB~8~a~e~^~T~]~k~e ~T~P~]~]~k~e^00BB ~b~P~\ ~]~U ~W~P~_~X~a~P~]~k.")   ' ru: Strok bez nomera linii (ili ego net v ?Ishodnyh dannyh?):   
    If cntBad > 0 Then s = s & vbLf & U("~7~]~P~g~U~]~X~o ~h~P~Q~[~^~]~P ~]~U ~_~`~^~h~[~X ~_~`~^~R~U~`~Z~c: ") & cntBad & badText   ' ru: Znacheniya shablona ne proshli proverku: 
    If cntDV > 0 Then s = s & vbLf & U("~=~U ~_~`~^~e~^~T~o~b ~R~k~_~P~T~P~n~i~X~Y ~a~_~X~a~^~Z (~R~^~W~R~`~P~i~U~]~^ ~Z~P~Z ~Q~k~[~^): ") & cntDV & dvText   ' ru: Ne prohodyat vypadayuschij spisok (vozvrascheno kak bylo): 
    If Len(badKeys) > 0 Then s = s & vbLf & U("~A~b~^~[~Q~f~k ~h~P~Q~[~^~]~P ~_~`~^~_~c~i~U~]~k (~]~U ~R~e~^~T~o~b ~R ~a~_~X~a~^~Z ~R~R~^~T~P): ") & badKeys   ' ru: Stolbcy shablona propuscheny (ne vhodyat v spisok vvoda): 
    SkipReport = s
End Function

'==========================================================================
' PUBLIC MACRO 2: save the selected row as a new template
'==========================================================================
Public Sub SX_SaveAsGroup()
    Dim wb As Workbook, wsO As Worksheet, wsT As Worksheet, wsS As Worksheet, wsR As Worksheet
    Dim selRng As Range, created As Boolean, badKeys As String
    Dim r As Long, j As Long, srcRow As Long, lineKey As String
    Dim vals() As Variant, has() As Boolean, nGot As Long
    Dim nm As String, descr As String, lastR As Long, newRow As Long, lastC As Long, i As Long
    Dim ans As VbMsgBoxResult, finalMsg As String, msgIcon As Long, idx As Collection, shield As String
    Dim v As Variant

    On Error GoTo EH
    g_warn = ""
    msgIcon = vbInformation
    If Not CheckContext(wb, wsO, selRng) Then GoTo Fin
    If selRng.Areas.Count > 1 Or selRng.Rows.Count > 1 Then
        finalMsg = U("~2~k~T~U~[~X~b~U ~>~4~=~C ~a~b~`~^~Z~c (~[~n~Q~c~n ~o~g~U~Y~Z~c ~R ~]~U~Y) ~]~P ~[~X~a~b~U ^00AB~>~T~]~^~[~X~]~U~Y~Z~P^00BB.")   ' ru: Vydelite ODNU stroku (lyubuyu yachejku v nej) na liste ?Odno
        msgIcon = vbExclamation
        GoTo Fin
    End If
    r = selRng.Row
    If r < FIRST_ROW Or r > LAST_ROW Then
        finalMsg = U("~2~k~T~U~[~X~b~U ~a~b~`~^~Z~c ~T~P~]~]~k~e (~a~b~`~^~Z~X 3-1000).")   ' ru: Vydelite stroku dannyh (stroki 3-1000).
        msgIcon = vbExclamation
        GoTo Fin
    End If

    Set wsT = GetTplSheet(wb, created)
    If created Then
        wsO.Activate
        AppRestore
    End If
    Set wsS = FindSheet(wb, nmSrc())
    Set wsR = FindSheet(wb, nmRazb())
    If LoadColumns(wsT, wsO, wsS, badKeys) = 0 Then
        finalMsg = U("~=~P ~[~X~a~b~U ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB ~]~U~b ~`~P~a~_~^~W~]~P~]~]~k~e ~a~b~^~[~Q~f~^~R (~a~b~`~^~Z~P 2: ^00AB~[~X~a~b!~a~b~^~[~Q~U~f^00BB).")   ' ru: Na liste ?Tipovye_gruppy? net raspoznannyh stolbcov (stroka 
        msgIcon = vbExclamation
        GoTo Fin
    End If

    ReDim vals(1 To nCols)
    ReDim has(1 To nCols)
    lineKey = LineOfRow(wsO, wsR, r)
    srcRow = 0
    If Len(lineKey) > 0 And Not wsS Is Nothing Then
        Set idx = BuildLineIndex(wsS)
        srcRow = LookupLine(idx, lineKey)
    End If
    For j = 1 To nCols
        If cSheet(j) = 1 Then
            has(j) = ReadInput(wsO, r, cNum(j), vals(j))
        ElseIf srcRow > 0 Then
            has(j) = ReadInput(wsS, srcRow, cNum(j), vals(j))
        End If
        If has(j) Then nGot = nGot + 1
    Next j
    If nGot = 0 Then
        finalMsg = U("~2 ~a~b~`~^~Z~U ") & r & U(" ~]~U~b ~W~]~P~g~U~]~X~Y ~R~R~^~T~P ~T~[~o ~h~P~Q~[~^~]~P (~_~c~a~b~^ ~X~[~X ~b~^~[~l~Z~^ ~d~^~`~\~c~[~k).")   ' ru: V stroke   net znachenij vvoda dlya shablona (pusto ili tolk
        msgIcon = vbExclamation
        GoTo Fin
    End If

    nm = Trim$(InputBox(U("~8~\~o ~]~^~R~^~S~^ ~h~P~Q~[~^~]~P (~W~]~P~g~U~]~X~Y ~R~R~^~T~P: ") & nGot & U("):"), ttl(), U("~=~^~R~P~o ~S~`~c~_~_~P")))   ' ru: Imya novogo shablona (znachenij vvoda:  Novaya gruppa
    If Len(nm) = 0 Then GoTo Fin

    lastR = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    If lastR < TPL_FIRST_ROW - 1 Then lastR = TPL_FIRST_ROW - 1
    newRow = 0
    For i = TPL_FIRST_ROW To lastR
        If StrComp(Trim$(CellText(wsT.Cells(i, 1))), nm, vbTextCompare) = 0 Then
            newRow = i
            Exit For
        End If
    Next i
    If newRow > 0 Then
        ans = MsgBox(U("~H~P~Q~[~^~] ^00AB") & nm & U("^00BB ~c~V~U ~U~a~b~l. ~7~P~\~U~]~X~b~l?"), vbYesNo + vbQuestion + vbDefaultButton2, ttl())   ' ru: Shablon ? ? uzhe est. Zamenit?
        If ans <> vbYes Then GoTo Fin
    Else
        newRow = lastR + 1
    End If

    AppSave
    AppQuiet
    lastC = TPL_VAL_COL
    For j = 1 To nCols
        If cTpl(j) > lastC Then lastC = cTpl(j)
    Next j
    wsT.Range(wsT.Cells(newRow, 1), wsT.Cells(newRow, lastC)).ClearContents
    shield = Trim$(CellText(wsO.Cells(r, 1)))
    descr = U("~8~W ~a~b~`~^~Z~X ") & r   ' ru: Iz stroki 
    If Len(shield) > 0 Then descr = descr & U(", ~i~X~b ") & shield   ' ru: , schit 
    If Len(lineKey) > 0 Then descr = descr & U(", ~[~X~]~X~o ") & lineKey   ' ru: , liniya 
    descr = descr & " (" & Format$(Date, "dd.mm.yyyy") & ")."
    PutValue wsT.Cells(newRow, 1), nm
    PutValue wsT.Cells(newRow, 2), descr
    PutValue wsT.Cells(newRow, 3), U("~Z~P~V~T~P~o ~a~b~`~^~Z~P")   ' ru: kazhdaya stroka
    For j = 1 To nCols
        If has(j) Then PutValue wsT.Cells(newRow, cTpl(j)), vals(j)
    Next j
    FormatTplRow wsT, newRow, lastC

    finalMsg = U("~H~P~Q~[~^~] ^00AB") & nm & U("^00BB ~a~^~e~`~P~]~q~] (~[~X~a~b ^00AB~B~X~_~^~R~k~U_~S~`~c~_~_~k^00BB, ~a~b~`~^~Z~P ") & newRow & U(", ~W~]~P~g~U~]~X~Y: ") & nGot & ")."   ' ru: Shablon ? ? sohranen (list ?Tipovye_gruppy?, stroka  , znach
    If Len(lineKey) = 0 Then
        finalMsg = finalMsg & vbLf & U("~C ~a~b~`~^~Z~X ~]~U~b ~]~^~\~U~`~P ~[~X~]~X~X - ~W~]~P~g~U~]~X~o ^00AB~8~a~e~^~T~]~k~e ~T~P~]~]~k~e^00BB ~]~U ~R~W~o~b~k.")   ' ru: U stroki net nomera linii - znacheniya ?Ishodnyh dannyh? ne 
    ElseIf srcRow = 0 Then
        finalMsg = finalMsg & vbLf & U("~;~X~]~X~X ^00AB") & lineKey & U("^00BB ~]~U~b ~R ^00AB~8~a~e~^~T~]~k~e ~T~P~]~]~k~e^00BB - ~R~W~o~b~k ~b~^~[~l~Z~^ ~W~]~P~g~U~]~X~o ^00AB~>~T~]~^~[~X~]~U~Y~Z~X^00BB.")   ' ru: Linii ? ? net v ?Ishodnyh dannyh? - vzyaty tolko znacheniya 
    End If
    finalMsg = finalMsg & vbLf & U("~@~U~V~X~\ - ^00AB~Z~P~V~T~P~o ~a~b~`~^~Z~P^00BB; ~T~[~o ~S~`~c~_~_ ~_~^~\~U~]~o~Y~b~U ~]~P ^00AB~^~T~]~P ~S~`~c~_~_~P^00BB. ~=~P~W~R~P~]~X~U ~X ~^~_~X~a~P~]~X~U ~\~^~V~]~^ ~_~`~P~R~X~b~l.")   ' ru: Rezhim - ?kazhdaya stroka?; dlya grupp pomenyajte na ?odna g

Fin:
    On Error Resume Next
    Cleanup
    If Len(g_warn) > 0 Then finalMsg = finalMsg & vbLf & g_warn
    If Len(finalMsg) > 1000 Then finalMsg = Left$(finalMsg, 995) & "..."
    If Len(finalMsg) > 0 Then MsgBox finalMsg, msgIcon, ttl()
    Exit Sub
EH:
    finalMsg = U("~>~h~X~Q~Z~P ") & Err.Number & ": " & Err.Description   ' ru: Oshibka 
    msgIcon = vbCritical
    Resume Fin
End Sub
