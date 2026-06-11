Option Explicit

'Prepares source data, splits it into tabs by a key column, and samples rows per tab.
'Then highlights sampled rows and filters each generated tab to show only those highlights.

'========================
' EDIT THESE VALUES ONLY
'========================
Private Const SOURCE_SHEET As String = "MainSheetName"

'Header used in step 1 to delete rows containing either value below
Private Const ROW_DELETE_CHECK_HEADER As String = "Dummy Status Header"

'Header used in step 2 to split data into new tabs
Private Const SPLIT_BY_HEADER As String = "Dummy Split Header"

Private Const SAMPLE_ROWS_PER_TAB As Long = 10

'RGB(198, 239, 206) - friendly green
Private Const FRIENDLY_GREEN As Long = 13561798

'Returns the list of source column headers that should be deleted during preparation.
Private Function ColumnsToRemove() As Variant
    ColumnsToRemove = Array( _
        "Dummy Column 1", _
        "Dummy Column 2", _
        "Dummy Column 3" _
    )
End Function

'Returns the list of cell values that mark rows to be deleted.
Private Function RowValuesToDelete() As Variant
    RowValuesToDelete = Array( _
        "Dummy Value 1", _
        "Dummy Value 2" _
    )
End Function

'========================
' MAIN ROUTINE
'========================
'Orchestrates the full end-to-end Samp Fixer workflow.
Public Sub RunSampFixerProcess()

    Dim wb As Workbook
    Dim sourceWs As Worksheet
    Dim createdSheets As Collection
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set sourceWs = wb.Worksheets(SOURCE_SHEET)
    Set createdSheets = New Collection

    With Application
        oldCalc = .Calculation
        oldScreenUpdating = .ScreenUpdating
        oldEnableEvents = .EnableEvents

        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    PrepareSampFixerData sourceWs
    CreateTabsByUniqueValue sourceWs, createdSheets
    HighlightRandomRows createdSheets
    FilterHighlightedRows createdSheets

CleanExit:
    With Application
        .Calculation = oldCalc
        .ScreenUpdating = oldScreenUpdating
        .EnableEvents = oldEnableEvents
    End With
    Exit Sub

CleanFail:
    MsgBox "Process failed: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub

'========================
' STEP 1
'========================
'Cleans the source sheet by removing top rows, selected columns, and rows with unwanted values.
Private Sub PrepareSampFixerData(ByVal ws As Worksheet)

    ws.Rows("1:3").Delete

    DeleteColumnsByHeader ws, ColumnsToRemove()
    DeleteRowsByValues ws, ROW_DELETE_CHECK_HEADER, RowValuesToDelete()

End Sub

'========================
' STEP 2
'========================
'Creates one new worksheet per unique split value and copies matching rows into each tab.
Private Sub CreateTabsByUniqueValue(ByVal sourceWs As Worksheet, ByVal createdSheets As Collection)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim splitCol As Long
    Dim data As Variant
    Dim groups As Object
    Dim rows As Collection
    Dim key As Variant
    Dim valueKey As String
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim outData As Variant
    Dim outWs As Worksheet
    Dim outName As String

    lastRow = LastUsedRow(sourceWs)
    lastCol = LastUsedColumn(sourceWs)

    If lastRow < 2 Then Exit Sub

    splitCol = HeaderColumn(sourceWs, SPLIT_BY_HEADER)

    data = sourceWs.Range(sourceWs.Cells(1, 1), sourceWs.Cells(lastRow, lastCol)).Value

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    'Iterates through each data row to group row indexes by split-column value.
    For r = 2 To UBound(data, 1)

        If Not IsError(data(r, splitCol)) Then
            valueKey = Trim$(CStr(data(r, splitCol)))

            If Len(valueKey) > 0 Then
                If groups.Exists(valueKey) Then
                    Set rows = groups(valueKey)
                Else
                    Set rows = New Collection
                    groups.Add valueKey, rows
                End If

                rows.Add r
            End If
        End If

    Next r

    'Builds and writes one output worksheet for each unique grouped value.
    For Each key In groups.Keys

        Set rows = groups(key)

        ReDim outData(1 To rows.Count + 1, 1 To lastCol)

        'Copies the source header row into the output array.
        For c = 1 To lastCol
            outData(1, c) = data(1, c)
        Next c

        'Copies each grouped source row into the output array.
        For r = 1 To rows.Count
            outRow = CLng(rows(r))

            'Copies every column value for the current grouped row.
            For c = 1 To lastCol
                outData(r + 1, c) = data(outRow, c)
            Next c
        Next r

        outName = GetUniqueSheetName(CStr(key), sourceWs.Parent)

        Set outWs = sourceWs.Parent.Worksheets.Add( _
            After:=sourceWs.Parent.Worksheets(sourceWs.Parent.Worksheets.Count) _
        )

        outWs.Name = outName
        outWs.Range("A1").Resize(UBound(outData, 1), UBound(outData, 2)).Value = outData
        outWs.Rows(1).Font.Bold = True
        outWs.Columns.AutoFit

        createdSheets.Add outWs

    Next key

End Sub

'========================
' STEP 3
'========================
'Randomly selects a sample of data rows in each created tab and highlights the full row.
Private Sub HighlightRandomRows(ByVal createdSheets As Collection)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sampleCount As Long
    Dim pickedRows As Object
    Dim pickedRow As Long
    Dim item As Variant

    Randomize

    'Processes each created worksheet to apply random-row highlighting.
    For Each ws In createdSheets

        lastRow = LastUsedRow(ws)
        lastCol = LastUsedColumn(ws)

        If lastRow > 1 Then

            sampleCount = SAMPLE_ROWS_PER_TAB

            If lastRow - 1 < sampleCount Then
                sampleCount = lastRow - 1
            End If

            Set pickedRows = CreateObject("Scripting.Dictionary")

            'Keeps selecting random row numbers until the sample size is reached.
            Do While pickedRows.Count < sampleCount
                pickedRow = CLng(Int((lastRow - 1) * Rnd) + 2)
                pickedRows(CStr(pickedRow)) = pickedRow
            Loop

            'Applies the highlight color across each randomly selected row.
            For Each item In pickedRows.Items
                ws.Range(ws.Cells(CLng(item), 1), ws.Cells(CLng(item), lastCol)).Interior.Color = FRIENDLY_GREEN
            Next item

        End If

    Next ws

End Sub

'========================
' STEP 4
'========================
'Applies a color filter so each created tab shows only highlighted sample rows.
Private Sub FilterHighlightedRows(ByVal createdSheets As Collection)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    'Applies the highlight-color filter on each created worksheet.
    For Each ws In createdSheets

        lastRow = LastUsedRow(ws)
        lastCol = LastUsedColumn(ws)

        If lastRow > 1 Then

            If ws.AutoFilterMode Then ws.AutoFilterMode = False

            Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

            dataRange.AutoFilter _
                Field:=1, _
                Criteria1:=FRIENDLY_GREEN, _
                Operator:=xlFilterCellColor

        End If

    Next ws

End Sub

'========================
' HELPERS
'========================
'Deletes columns whose header names match the configured removal list.
Private Sub DeleteColumnsByHeader(ByVal ws As Worksheet, ByVal headersToRemove As Variant)

    Dim deleteHeaders As Object
    Dim i As Long
    Dim lastCol As Long
    Dim headerText As String

    Set deleteHeaders = CreateObject("Scripting.Dictionary")
    deleteHeaders.CompareMode = vbTextCompare

    'Loads each configured header into a lookup dictionary for fast matching.
    For i = LBound(headersToRemove) To UBound(headersToRemove)
        deleteHeaders(Trim$(CStr(headersToRemove(i)))) = True
    Next i

    lastCol = LastUsedColumn(ws)

    'Scans header cells right-to-left and deletes any column in the removal list.
    For i = lastCol To 1 Step -1
        headerText = Trim$(CStr(ws.Cells(1, i).Value))

        If deleteHeaders.Exists(headerText) Then
            ws.Columns(i).Delete
        End If
    Next i

End Sub

'Deletes rows where the target column contains any configured value to remove.
Private Sub DeleteRowsByValues(ByVal ws As Worksheet, ByVal headerName As String, ByVal valuesToDelete As Variant)

    Dim checkCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim deleteRange As Range

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastRow < 2 Then Exit Sub

    checkCol = HeaderColumn(ws, headerName)

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    dataRange.AutoFilter _
        Field:=checkCol, _
        Criteria1:=valuesToDelete, _
        Operator:=xlFilterValues

    On Error Resume Next
    Set deleteRange = dataRange.Offset(1).Resize(dataRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not deleteRange Is Nothing Then
        deleteRange.EntireRow.Delete
    End If

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

End Sub

'Returns the column index for an exact header name in row 1.
Private Function HeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long

    Dim foundCell As Range

    Set foundCell = ws.Rows(1).Find( _
        What:=headerName, _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If foundCell Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Header not found: " & headerName
    End If

    HeaderColumn = foundCell.Column

End Function

'Finds the last used row on a worksheet, defaulting to row 1 if empty.
Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    )

    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If

End Function

'Finds the last used column based on row 1.
Private Function LastUsedColumn(ByVal ws As Worksheet) As Long

    LastUsedColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

End Function

'Builds a valid, unique worksheet name from a raw value.
Private Function GetUniqueSheetName(ByVal rawName As String, ByVal wb As Workbook) As String

    Dim baseName As String
    Dim proposedName As String
    Dim suffix As String
    Dim counter As Long

    baseName = CleanSheetName(rawName)

    If Len(baseName) = 0 Then baseName = "Blank"

    baseName = Left$(baseName, 31)
    proposedName = baseName
    counter = 2

    'Appends an incrementing suffix until an unused worksheet name is found.
    Do While SheetExists(proposedName, wb)
        suffix = " (" & counter & ")"
        proposedName = Left$(baseName, 31 - Len(suffix)) & suffix
        counter = counter + 1
    Loop

    GetUniqueSheetName = proposedName

End Function

'Replaces invalid worksheet-name characters and trims whitespace.
Private Function CleanSheetName(ByVal rawName As String) As String

    Dim badChars As Variant
    Dim i As Long

    CleanSheetName = Trim$(rawName)

    badChars = Array(":", "\", "/", "?", "*", "[", "]")

    'Replaces each invalid worksheet-name character with a hyphen.
    For i = LBound(badChars) To UBound(badChars)
        CleanSheetName = Replace$(CleanSheetName, badChars(i), "-")
    Next i

End Function

'Checks whether a worksheet name already exists in the workbook.
Private Function SheetExists(ByVal sheetName As String, ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function