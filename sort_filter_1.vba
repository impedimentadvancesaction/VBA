Option Explicit

' Cleans the source sheet, splits records into category tabs, and randomly highlights rows in each tab.
' Then filters each generated tab to show only the rows highlighted in green.

'================== CONFIG (edit these) ==================
Private Const SOURCE_SHEET As String = "MainSheetName"

' Step 1 - columns to delete by their header name (row 1 AFTER rows 1-3 are removed)
Private Function ColumnsToDelete() As Variant
    ColumnsToDelete = Array("DummyColA", "DummyColB", "DummyColC")
End Function

' Step 1 - delete any row where this column equals either value
Private Const FILTER_COLUMN As String = "DummyCol"
Private Const DELETE_VALUE_1 As String = "DummyValue1"
Private Const DELETE_VALUE_2 As String = "DummyValue2"

' Step 2 - split into tabs by this column
Private Const GROUP_COLUMN As String = "Category"

' Step 3 - how many rows to highlight + the green used (also used by Step 4)
Private Const RANDOM_ROW_COUNT As Long = 10
Private Function GreenColour() As Long
    GreenColour = RGB(198, 239, 206)   ' soft "good" green
End Function
'========================================================

Private mNewSheets As Collection   ' tracks tabs created in Step 2

'==================== MAIN ====================
Public Sub RunSampFixer()
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Randomize

    Set mNewSheets = New Collection

    Step1_CleanSource
    Step2_SplitIntoTabs
    Step3_HighlightRandomRows
    Step4_FilterByColour

CleanUp:
    Application.CutCopyMode = False
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

'==================== STEP 1 ====================
Private Sub Step1_CleanSource()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SOURCE_SHEET)

    ' 1a. remove rows 1-3 (old row 4 becomes the header row)
    ws.Rows("1:3").Delete

    ' 1b. delete columns by name (resolve all indices, then delete in one shot)
    Dim cols As Variant, i As Long, idx As Long, delRange As Range
    cols = ColumnsToDelete()
    For i = LBound(cols) To UBound(cols)
        idx = GetColIndex(ws, CStr(cols(i)))
        If idx > 0 Then
            If delRange Is Nothing Then Set delRange = ws.Columns(idx) _
            Else Set delRange = Union(delRange, ws.Columns(idx))
        End If
    Next i
    If Not delRange Is Nothing Then delRange.Delete

    ' 1c. delete rows where FILTER_COLUMN is either target value
    '     (AutoFilter + bulk delete = the performant equivalent of looping)
    Dim fCol As Long
    fCol = GetColIndex(ws, FILTER_COLUMN)
    If fCol = 0 Then Exit Sub

    Dim lastRow As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, fCol).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .AutoFilter Field:=fCol, Criteria1:=DELETE_VALUE_1, _
                    Operator:=xlOr, Criteria2:=DELETE_VALUE_2
        Dim vis As Range
        On Error Resume Next
        Set vis = .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not vis Is Nothing Then vis.EntireRow.Delete
    End With
    ws.AutoFilterMode = False
End Sub

'==================== STEP 2 ====================
Private Sub Step2_SplitIntoTabs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SOURCE_SHEET)

    Dim gCol As Long
    gCol = GetColIndex(ws, GROUP_COLUMN)
    If gCol = 0 Then Exit Sub

    Dim lastRow As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, gCol).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' collect unique values
    Dim dict As Object, arr As Variant, r As Long, key As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    arr = ws.Range(ws.Cells(2, gCol), ws.Cells(lastRow, gCol)).Value
    For r = 1 To UBound(arr, 1)
        key = Trim(CStr(arr(r, 1)))
        If Len(key) > 0 Then If Not dict.Exists(key) Then dict.Add key, 1
    Next r

    ' one tab per value (filtered copy includes the header row automatically)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Dim srcRange As Range, k As Variant, newWs As Worksheet
    Set srcRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    For Each k In dict.Keys
        Set newWs = GetOrCreateSheet(SafeSheetName(CStr(k)))
        newWs.Cells.Clear
        srcRange.AutoFilter Field:=gCol, Criteria1:=CStr(k)
        srcRange.SpecialCells(xlCellTypeVisible).Copy newWs.Range("A1")
        ws.AutoFilterMode = False
        mNewSheets.Add newWs.Name
    Next k
    Application.CutCopyMode = False
End Sub

'==================== STEP 3 ====================
Private Sub Step3_HighlightRandomRows()
    ' Run random-row highlighting on each worksheet created during Step 2.
    Dim nm As Variant
    For Each nm In mNewSheets
        HighlightOneSheet ThisWorkbook.Worksheets(CStr(nm))
    Next nm
End Sub

Private Sub HighlightOneSheet(ws As Worksheet)
    ' Determine the used data bounds (row 1 is header).
    Dim lastRow As Long, lastCol As Long, dataRows As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    dataRows = lastRow - 1
    If dataRows <= 0 Then Exit Sub

    ' Choose how many rows to highlight (cap at available data rows).
    Dim pick As Long
    pick = RANDOM_ROW_COUNT
    If pick > dataRows Then pick = dataRows   ' fewer than 10 rows -> highlight all

    ' Build a unique set of random data-row offsets.
    Dim chosen As Object, idx As Long
    Set chosen = CreateObject("Scripting.Dictionary")
    Do While chosen.Count < pick
        idx = Int(Rnd() * dataRows) + 1
        If Not chosen.Exists(idx) Then chosen.Add idx, 1
    Loop

    ' Apply the green fill across each selected row.
    Dim k As Variant, tRow As Long
    For Each k In chosen.Keys
        tRow = 1 + CLng(k)   ' header is row 1, data starts row 2
        ws.Range(ws.Cells(tRow, 1), ws.Cells(tRow, lastCol)).Interior.Color = GreenColour()
    Next k
End Sub

'==================== STEP 4 ====================
Private Sub Step4_FilterByColour()
    ' Loop each generated sheet and keep only rows highlighted in Step 3 visible.
    Dim nm As Variant, ws As Worksheet, lastRow As Long, lastCol As Long
    For Each nm In mNewSheets
        Set ws = ThisWorkbook.Worksheets(CStr(nm))

        ' Recalculate the sheet bounds before applying the color filter.
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then

            ' Filter by fill color on column A (entire row is colored, so any column works).
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
                Field:=1, Criteria1:=GreenColour(), Operator:=xlFilterCellColor
        End If
    Next nm
End Sub

'==================== HELPERS ====================
Private Function GetColIndex(ws As Worksheet, headerName As String) As Long
    ' Find a header in row 1 using a case-insensitive exact match.
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim(CStr(ws.Cells(1, c).Value)), headerName, vbTextCompare) = 0 Then
            GetColIndex = c: Exit Function
        End If
    Next c
End Function

Private Function SafeSheetName(ByVal nm As String) As String
    ' Replace invalid worksheet-name characters and enforce Excel length rules.
    Dim bad As Variant, b As Variant
    bad = Array("\", "/", "?", "*", "[", "]", ":")
    For Each b In bad: nm = Replace(nm, b, "_"): Next b
    nm = Trim(nm)
    If Len(nm) = 0 Then nm = "Blank"
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    SafeSheetName = nm
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    ' Reuse an existing worksheet if present; otherwise create it at workbook end.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function