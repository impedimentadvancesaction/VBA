Attribute VB_Name = "mod_ReportGenerator"
Option Explicit

'==============================================================================
' REPORT GENERATOR
'------------------------------------------------------------------------------
' HOW IT WORKS
'   * Sheet2 (Working Out) rows are grouped by S2_COL4. Each unique value = one
'     report, producing TWO output files (Report Template 1 + Report Template 2)
'     saved as .xlsx in this workbook's folder (existing files overwritten).
'   * Within a group, each row's S2_COL3 value is the key looked up against
'     Sheet3's S3_KEY column. One Sheet5 data row is written per group row.
'   * Templates are NEVER modified: each output file is built on a COPY of the
'     template sheet in a brand-new workbook, so the templates in this workbook
'     always remain in their default state (no reset step required).
'   * Every report is fully validated BEFORE any file is created. On failure a
'     pop-up appears, that report is skipped, and generation continues.
'   * A final summary lists every report generated and every report cancelled.
'
' EVERYTHING you may need to change lives between the CONFIG banners and inside
' GetTypeConfig(). Nothing below the "ENGINE" banner requires editing.
'
' RUN: GenerateReports
'==============================================================================

'============================ CONFIG: GENERAL =================================
Private Const OUTPUT_FILE_EXT As String = ".xlsx"
Private Const A16_DATE_SUFFIX As String = "- 1 Hour"        ' text appended after the A16 date
Private Const S5_C11_DATE_FORMAT As String = "dd/mm/yyyy"   ' display format for Sheet5 S5_C11

'============================ CONFIG: SHEET NAMES =============================
Private Const SH_WORKING As String = "Working Out Sheet"    ' "Sheet2"
Private Const SH_DATA    As String = "Data Export"          ' "Sheet3"
Private Const SH_TPL1    As String = "Report Template 1"    ' "Sheet4"
Private Const SH_TPL2    As String = "Report Template 2"    ' "Sheet5"

'==================== CONFIG: SHEET2 (WORKING OUT SHEET) ======================
' Header captions sit in row S2_HDR_ROW; data starts on the next row.
Private Const S2_HDR_ROW As Long = 1
Private Const S2_COL2 As String = "ColName2"   ' -> Template 1 cell A19
Private Const S2_COL3 As String = "ColName3"   ' lookup key matched against Sheet3 S3_KEY
Private Const S2_COL4 As String = "ColName4"   ' grouping value: one report per unique value
Private Const S2_COL5 As String = "ColName5"   ' -> Template 2 column S5_C4
Private Const S2_COL6 As String = "ColName6"   ' report type

'======================= CONFIG: SHEET3 (DATA EXPORT) =========================
Private Const S3_HDR_ROW As Long = 1
Private Const S3_KEY  As String = "ColName1"   ' key column (matched by Sheet2 S2_COL3 values)
Private Const S3_RET1 As String = "ColName1"   ' -> Sheet5 S5_C2
Private Const S3_RET2 As String = "ColName2"   ' -> Sheet5 S5_C3
Private Const S3_RET3 As String = "ColName3"   ' -> Sheet5 S5_C6
Private Const S3_RET4 As String = "ColName4"   ' -> Sheet5 S5_C7
Private Const S3_RET5 As String = "ColName5"   ' -> Sheet5 S5_C8
Private Const S3_RET6 As String = "ColName6"   ' -> Sheet5 S5_C9  (TRUE -> "Y", else "N")
Private Const S3_RET7 As String = "ColName7"   ' -> Sheet5 S5_C10 (in list below -> "Y", else "N")
' Values in S3_RET7 that convert to "Y" (semicolon-separated, case-insensitive):
Private Const S3_RET7_YES_LIST As String = "VALUE1;VALUE2;VALUE3"

'=================== CONFIG: SHEET5 (REPORT TEMPLATE 2) =======================
' Row 1 holds the special value (preserved automatically because the whole
' sheet is copied). Headers sit in row S5_HDR_ROW; data is written from
' S5_DATA_ROW downward.
Private Const S5_HDR_ROW  As Long = 2
Private Const S5_DATA_ROW As Long = 3
Private Const S5_C1  As String = "ColName1"
Private Const S5_C2  As String = "ColName2"
Private Const S5_C3  As String = "ColName3"
Private Const S5_C4  As String = "ColName4"
Private Const S5_C5  As String = "ColName5"
Private Const S5_C6  As String = "ColName6"
Private Const S5_C7  As String = "ColName7"
Private Const S5_C8  As String = "ColName8"
Private Const S5_C9  As String = "ColName9"
Private Const S5_C10 As String = "ColName10"
Private Const S5_C11 As String = "ColName11"

'=================== CONFIG: SHEET4 (REPORT TEMPLATE 1) =======================
Private Const T1_CELL_TITLE As String = "A1"
Private Const T1_CELL_SUB   As String = "A4"
Private Const T1_CELL_DATE  As String = "A16"
Private Const T1_CELL_REF   As String = "A19"
Private Const T1_CELL_V1    As String = "A28"
Private Const T1_CELL_V2    As String = "A34"

'========================= CONFIG: REPORT TYPES ===============================
Private Const RTYPE1 As String = "ReportType1"
Private Const RTYPE2 As String = "ReportType2"
Private Const RTYPE3 As String = "ReportType3"
Private Const RTYPE4 As String = "ReportType4"

Private Type TypeCfg
    IsValid As Boolean
    WorkdayOffset As Long    ' working days (Mon-Fri) added to today for the A16 date
    A1Prefix As String       ' A1 = A1Prefix & " (" & <unique value> & ")"
    A4Line2 As String        ' second line of cell A4
    A28A34Prefix As String   ' A28 / A34 = prefix & <unique value>
    S5C5Literal As String    ' constant written down Sheet5 column S5_C5
    S5FilePrefix As String   ' Template 2 filename = prefix & <unique value>
End Type

' All per-type placeholders (VALUE / VALUE2 etc.) are set here, once per type,
' so the four types can diverge later without touching the engine.
Private Function GetTypeConfig(ByVal typeName As String) As TypeCfg
    Dim c As TypeCfg
    c.IsValid = True
    Select Case UCase$(Trim$(typeName))
        Case UCase$(RTYPE1)
            c.WorkdayOffset = 2
            c.A1Prefix = "VALUE"
            c.A4Line2 = "VALUE2"
            c.A28A34Prefix = "VALUE"
            c.S5C5Literal = "VALUE"
            c.S5FilePrefix = "VALUE"
        Case UCase$(RTYPE2)
            c.WorkdayOffset = 1
            c.A1Prefix = "VALUE"
            c.A4Line2 = "VALUE2"
            c.A28A34Prefix = "VALUE"
            c.S5C5Literal = "VALUE"
            c.S5FilePrefix = "VALUE"
        Case UCase$(RTYPE3)
            c.WorkdayOffset = 1
            c.A1Prefix = "VALUE"
            c.A4Line2 = "VALUE2"
            c.A28A34Prefix = "VALUE"
            c.S5C5Literal = "VALUE"
            c.S5FilePrefix = "VALUE"
        Case UCase$(RTYPE4)
            c.WorkdayOffset = 1
            c.A1Prefix = "VALUE"
            c.A4Line2 = "VALUE2"
            c.A28A34Prefix = "VALUE"
            c.S5C5Literal = "VALUE"
            c.S5FilePrefix = "VALUE"
        Case Else
            c.IsValid = False
    End Select
    GetTypeConfig = c
End Function

'==============================================================================
'                                   ENGINE
'==============================================================================

Public Sub GenerateReports()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    On Error GoTo Fatal

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 100, , _
            "Save this workbook first so the output files have a folder to be saved into."
    End If

    Dim ws2 As Worksheet, ws3 As Worksheet, ws5 As Worksheet
    Set ws2 = GetSheet(SH_WORKING)
    Set ws3 = GetSheet(SH_DATA)
    GetSheet SH_TPL1                     ' existence check only
    Set ws5 = GetSheet(SH_TPL2)

    ' --- Header maps (header caption -> column number) ------------------------
    Dim m2 As Object, m3 As Object, m5 As Object
    Set m2 = HeaderMap(ws2, S2_HDR_ROW)
    Set m3 = HeaderMap(ws3, S3_HDR_ROW)
    Set m5 = HeaderMap(ws5, S5_HDR_ROW)
    RequireHeaders m2, SH_WORKING, Array(S2_COL2, S2_COL3, S2_COL4, S2_COL5, S2_COL6)
    RequireHeaders m3, SH_DATA, Array(S3_KEY, S3_RET1, S3_RET2, S3_RET3, S3_RET4, S3_RET5, S3_RET6, S3_RET7)
    RequireHeaders m5, SH_TPL2, Array(S5_C1, S5_C2, S5_C3, S5_C4, S5_C5, S5_C6, S5_C7, S5_C8, S5_C9, S5_C10, S5_C11)

    ' --- Read both data sheets into memory once -------------------------------
    Dim last2 As Long, last3 As Long
    last2 = ws2.Cells(ws2.Rows.Count, m2(S2_COL4)).End(xlUp).Row
    last3 = ws3.Cells(ws3.Rows.Count, m3(S3_KEY)).End(xlUp).Row
    If last2 <= S2_HDR_ROW Then Err.Raise vbObjectError + 101, , "No data rows found on '" & SH_WORKING & "'."
    If last3 <= S3_HDR_ROW Then Err.Raise vbObjectError + 102, , "No data rows found on '" & SH_DATA & "'."

    Dim a2 As Variant, a3 As Variant
    a2 = ws2.Range(ws2.Cells(1, 1), ws2.Cells(last2, MaxMapValue(m2))).Value
    a3 = ws3.Range(ws3.Cells(1, 1), ws3.Cells(last3, MaxMapValue(m3))).Value

    ' --- Index Sheet3: key -> row number (first occurrence wins) --------------
    Dim idx3 As Object
    Set idx3 = CreateObject("Scripting.Dictionary")
    idx3.CompareMode = vbTextCompare
    Dim r As Long, k As String
    For r = S3_HDR_ROW + 1 To last3
        If Not IsError(a3(r, m3(S3_KEY))) Then
            k = Trim$(CStr(a3(r, m3(S3_KEY))))
            If Len(k) > 0 Then If Not idx3.Exists(k) Then idx3.Add k, r
        End If
    Next r

    ' --- Group Sheet2 rows by the report value (insertion order preserved) ----
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Dim x As String
    For r = S2_HDR_ROW + 1 To last2
        If Not IsError(a2(r, m2(S2_COL4))) Then
            x = Trim$(CStr(a2(r, m2(S2_COL4))))
            If Len(x) > 0 Then
                If Not groups.Exists(x) Then groups.Add x, New Collection
                groups(x).Add r
            End If
        End If
    Next r
    If groups.Count = 0 Then Err.Raise vbObjectError + 103, , "No report values found in '" & S2_COL4 & "' on '" & SH_WORKING & "'."

    ' --- Generate ---------------------------------------------------------------
    Dim okList As Collection, badList As Collection
    Set okList = New Collection
    Set badList = New Collection

    Dim key As Variant, failMsg As String
    For Each key In groups.Keys
        failMsg = BuildOneReport(CStr(key), groups(key), a2, a3, m2, m3, m5, idx3)
        If Len(failMsg) = 0 Then
            okList.Add CStr(key)
        Else
            badList.Add CStr(key) & "  -  " & failMsg
            MsgBox "Report for '" & key & "' was cancelled:" & vbNewLine & vbNewLine & failMsg, _
                   vbExclamation, "Report cancelled"
        End If
    Next key

    ' --- Summary ----------------------------------------------------------------
    Dim msg As String
    msg = "Report generation complete." & vbNewLine & vbNewLine
    msg = msg & "Generated (" & okList.Count & "):" & vbNewLine & JoinCollection(okList) & vbNewLine & vbNewLine
    msg = msg & "Cancelled (" & badList.Count & "):" & vbNewLine & JoinCollection(badList)
    MsgBox msg, vbInformation, "Report Generator"

Restore:
    Application.DisplayAlerts = True
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fatal:
    MsgBox "Report generation stopped:" & vbNewLine & vbNewLine & Err.Description, _
           vbCritical, "Report Generator"
    Resume Restore
End Sub

'------------------------------------------------------------------------------
' Builds both output files for one report value. Returns "" on success or a
' human-readable failure reason (report is skipped, nothing is left behind).
'------------------------------------------------------------------------------
Private Function BuildOneReport(ByVal x As String, grpRows As Collection, _
                                a2 As Variant, a3 As Variant, _
                                m2 As Object, m3 As Object, m5 As Object, _
                                idx3 As Object) As String
    Dim wbOut As Workbook
    On Error GoTo Failed

    Dim n As Long
    n = grpRows.Count

    Dim s2r() As Long, s3r() As Long
    ReDim s2r(1 To n)
    ReDim s3r(1 To n)

    Dim i As Long, itm As Variant
    i = 0
    For Each itm In grpRows
        i = i + 1
        s2r(i) = itm
    Next itm

    ' --- Report type: taken from the group's first row; a mixed group is an
    '     input error, so it is cancelled loudly rather than silently guessed.
    Dim typeName As String
    typeName = Trim$(CStr(a2(s2r(1), m2(S2_COL6))))
    For i = 2 To n
        If StrComp(Trim$(CStr(a2(s2r(i), m2(S2_COL6)))), typeName, vbTextCompare) <> 0 Then
            BuildOneReport = "Rows for this value have different '" & S2_COL6 & "' report types."
            Exit Function
        End If
    Next i

    Dim cfg As TypeCfg
    cfg = GetTypeConfig(typeName)
    If Not cfg.IsValid Then
        BuildOneReport = "Unknown report type '" & typeName & "' in '" & S2_COL6 & "'."
        Exit Function
    End If

    ' --- Validate every lookup BEFORE creating anything -----------------------
    Dim retCols As Variant
    retCols = Array(S3_RET1, S3_RET2, S3_RET3, S3_RET4, S3_RET5, S3_RET6, S3_RET7)

    Dim j As Long, k As String, v As Variant
    For i = 1 To n
        k = Trim$(CStr(a2(s2r(i), m2(S2_COL3))))
        If Len(k) = 0 Then
            BuildOneReport = "'" & S2_COL3 & "' is blank on row " & s2r(i) & " of '" & SH_WORKING & "'."
            Exit Function
        End If
        If Not idx3.Exists(k) Then
            BuildOneReport = "Ref '" & k & "' could not be found in '" & SH_DATA & "' column '" & S3_KEY & "'."
            Exit Function
        End If
        s3r(i) = idx3(k)
        For j = LBound(retCols) To UBound(retCols)
            v = a3(s3r(i), m3(retCols(j)))
            If IsError(v) Then
                BuildOneReport = "'" & SH_DATA & "' column '" & retCols(j) & "' contains an error value for ref '" & k & "'."
                Exit Function
            ElseIf Len(Trim$(CStr(v))) = 0 Then
                BuildOneReport = "'" & SH_DATA & "' column '" & retCols(j) & "' is blank for ref '" & k & "'."
                Exit Function
            End If
        Next j
    Next i

    ' --- Report date (weekends skipped) ---------------------------------------
    Dim repDate As Date, dateText As String
    repDate = AddWorkdays(Date, cfg.WorkdayOffset)
    dateText = FormatUKDate(repDate)

    ' ==================== OUTPUT 1: Report Template 1 =========================
    Dim a1Text As String
    a1Text = cfg.A1Prefix & " (" & x & ")"

    ThisWorkbook.Worksheets(SH_TPL1).Copy       ' new single-sheet workbook
    Set wbOut = ActiveWorkbook
    With wbOut.Worksheets(1)
        .Range(T1_CELL_TITLE).Value = a1Text
        .Range(T1_CELL_SUB).Value = ExtractBetweenParens(a1Text) & vbLf & cfg.A4Line2
        .Range(T1_CELL_REF).Value = a2(s2r(1), m2(S2_COL2))
        .Range(T1_CELL_DATE).Value = dateText & " " & A16_DATE_SUFFIX
        .Range(T1_CELL_V1).Value = cfg.A28A34Prefix & x
        .Range(T1_CELL_V2).Value = cfg.A28A34Prefix & x
    End With
    SaveAndClose wbOut, SanitizeFileName(a1Text)
    Set wbOut = Nothing

    ' ==================== OUTPUT 2: Report Template 2 =========================
    ' Build all values in memory first (logical columns 1-11), then write each
    ' logical column to its physical position in one range write.
    Dim data() As Variant
    ReDim data(1 To n, 1 To 11)
    For i = 1 To n
        data(i, 1) = x
        data(i, 2) = a3(s3r(i), m3(S3_RET1))
        data(i, 3) = a3(s3r(i), m3(S3_RET2))
        data(i, 4) = a2(s2r(i), m2(S2_COL5))
        data(i, 5) = cfg.S5C5Literal
        data(i, 6) = a3(s3r(i), m3(S3_RET3))
        data(i, 7) = a3(s3r(i), m3(S3_RET4))
        data(i, 8) = a3(s3r(i), m3(S3_RET5))
        data(i, 9) = YNFromTrue(a3(s3r(i), m3(S3_RET6)))
        data(i, 10) = YNFromList(a3(s3r(i), m3(S3_RET7)))
        data(i, 11) = repDate            ' genuine date serial; displayed via S5_C11_DATE_FORMAT
    Next i

    Dim tgtHdrs As Variant
    tgtHdrs = Array(S5_C1, S5_C2, S5_C3, S5_C4, S5_C5, S5_C6, S5_C7, S5_C8, S5_C9, S5_C10, S5_C11)

    ThisWorkbook.Worksheets(SH_TPL2).Copy       ' row 1 special value comes along
    Set wbOut = ActiveWorkbook
    Dim buf() As Variant, c As Long
    With wbOut.Worksheets(1)
        For c = 1 To 11
            ReDim buf(1 To n, 1 To 1)
            For i = 1 To n
                buf(i, 1) = data(i, c)
            Next i
            .Cells(S5_DATA_ROW, m5(tgtHdrs(c - 1))).Resize(n, 1).Value = buf
        Next c
        .Cells(S5_DATA_ROW, m5(S5_C11)).Resize(n, 1).NumberFormat = S5_C11_DATE_FORMAT
    End With
    SaveAndClose wbOut, SanitizeFileName(cfg.S5FilePrefix & x)
    Set wbOut = Nothing

    BuildOneReport = vbNullString
    Exit Function

Failed:
    BuildOneReport = "Unexpected error: " & Err.Description
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
End Function

'==============================================================================
'                                  HELPERS
'==============================================================================

' Adds n working days to a date, skipping Saturdays and Sundays.
Private Function AddWorkdays(ByVal startDate As Date, ByVal daysToAdd As Long) As Date
    Dim d As Date, added As Long
    d = startDate
    Do While added < daysToAdd
        d = d + 1
        If Weekday(d, vbMonday) <= 5 Then added = added + 1
    Loop
    AddWorkdays = d
End Function

' "MMMM DD, YYYY" with explicit English month names, immune to system locale.
Private Function FormatUKDate(ByVal d As Date) As String
    Static monthNames As Variant
    If IsEmpty(monthNames) Then
        monthNames = Array("January", "February", "March", "April", "May", "June", _
                           "July", "August", "September", "October", "November", "December")
    End If
    FormatUKDate = monthNames(Month(d) - 1) & " " & Format$(Day(d), "00") & ", " & Year(d)
End Function

' Trimmed text between the first ")" and the following "(" on line 1 of s.
Private Function ExtractBetweenParens(ByVal s As String) As String
    Dim line1 As String, p1 As Long, p2 As Long
    line1 = Split(Replace(s, vbCrLf, vbLf), vbLf)(0)
    p1 = InStr(1, line1, ")")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, line1, "(")
    If p2 = 0 Then Exit Function
    ExtractBetweenParens = Trim$(Mid$(line1, p1 + 1, p2 - p1 - 1))
End Function

' TRUE (boolean or text, any case) -> "Y", anything else -> "N".
Private Function YNFromTrue(ByVal v As Variant) As String
    If VarType(v) = vbBoolean Then
        YNFromTrue = IIf(v, "Y", "N")
    ElseIf StrComp(Trim$(CStr(v)), "TRUE", vbTextCompare) = 0 Then
        YNFromTrue = "Y"
    Else
        YNFromTrue = "N"
    End If
End Function

' Value in S3_RET7_YES_LIST (case-insensitive) -> "Y", anything else -> "N".
Private Function YNFromList(ByVal v As Variant) As String
    Dim items() As String, i As Long, s As String
    s = Trim$(CStr(v))
    items = Split(S3_RET7_YES_LIST, ";")
    YNFromList = "N"
    For i = LBound(items) To UBound(items)
        If StrComp(s, Trim$(items(i)), vbTextCompare) = 0 Then
            YNFromList = "Y"
            Exit Function
        End If
    Next i
End Function

' Strips characters Windows forbids in filenames.
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "-")
    Next i
    SanitizeFileName = Trim$(s)
End Function

' Saves a workbook as .xlsx in this workbook's folder (silent overwrite), closes it.
Private Sub SaveAndClose(wb As Workbook, ByVal baseName As String)
    wb.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & baseName & OUTPUT_FILE_EXT, _
              FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
End Sub

' Header caption -> column number for a worksheet's header row (case-insensitive).
Private Function HeaderMap(ws As Worksheet, ByVal hdrRow As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim lastCol As Long, c As Long, h As String
    lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(hdrRow, c).Value))
        If Len(h) > 0 Then If Not d.Exists(h) Then d.Add h, c
    Next c
    Set HeaderMap = d
End Function

' Raises a clear error if any required header caption is missing.
Private Sub RequireHeaders(m As Object, ByVal sheetName As String, req As Variant)
    Dim i As Long, missing As String
    For i = LBound(req) To UBound(req)
        If Not m.Exists(req(i)) Then missing = missing & vbNewLine & "  - " & req(i)
    Next i
    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 200, , _
            "The following headers were not found on '" & sheetName & "':" & missing & _
            vbNewLine & vbNewLine & "Update the CONFIG constants or the sheet headers."
    End If
End Sub

' Worksheet by name, with a friendly error if missing.
Private Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Err.Raise vbObjectError + 201, , _
            "Worksheet '" & nm & "' was not found. Update the sheet-name CONFIG constants."
    End If
End Function

Private Function MaxMapValue(m As Object) As Long
    Dim v As Variant
    For Each v In m.Items
        If v > MaxMapValue Then MaxMapValue = v
    Next v
End Function

Private Function JoinCollection(col As Collection) As String
    Dim itm As Variant, s As String
    For Each itm In col
        s = s & vbNewLine & "  - " & itm
    Next itm
    If Len(s) = 0 Then s = vbNewLine & "  (none)"
    JoinCollection = s
End Function