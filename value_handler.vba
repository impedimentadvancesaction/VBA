Option Explicit

' -------------------------
' Configuration - EDIT ONLY
' -------------------------
Private Const sSheetData As String = "Sheet1"      ' source import data sheet
Private Const sSheetRefs As String = "Sheet2"      ' sheet with Ref, Name, Date rows
Private Const sSheetTemplate As String = "Sheet3"  ' report template sheet (visible)
Private Const sTemplateBackup As String = "TemplateBackup" ' hidden copy of template (must exist)
Private Const sSaveFolder As String = ""           ' "" = current workbook folder; or full path with trailing "\"

' Column header names on Sheet2 (Ref list)
Private Const hdrRef As String = "Ref"
Private Const hdrName As String = "Name"
Private Const hdrDate As String = "Date"

' Column header names on Sheet1 (source) - these are the 7 columns to copy in order
' Example: "Player","Game","Date","Score","Value","Cost","Count"
Private srcCols As Variant
' Column header names on Sheet3 (destination) - where the 7 columns should be placed (matching order)
Private dstCols As Variant

' Which of the 7 columns are numeric and must be summed on duplicate Date+Name
' Use 1-based indices into the 7 columns. For your spec: columns 3,4,5,6,7 are summed.
Private sumIndices As Variant

' Minimum distinct dates required per Name
Private Const MIN_DAYS As Long = 7

' -------------------------
' End Configuration
' -------------------------

' Initialize configuration arrays
Private Sub InitConfig()
    ' Edit these arrays to match your column header names
    srcCols = Array("ColumnName1", "ColumnName2", "ColumnName3", "ColumnName4", "ColumnName5", "ColumnName6", "ColumnName7")
    dstCols = Array("ColumnName1", "ColumnName2", "ColumnName3", "ColumnName4", "ColumnName5", "ColumnName6", "ColumnName7")
    sumIndices = Array(3, 4, 5, 6, 7) ' 1-based positions in the 7-column list
End Sub

' -------------------------
' Main routine
' -------------------------
Public Sub CreateReportsPerRef()
    Dim wb As Workbook, wbOut As Workbook
    Dim wsData As Worksheet, wsRefs As Worksheet, wsTemp As Worksheet, wsBackup As Worksheet
    Dim dictRefs As Object, dictNames As Object
    Dim lastRow As Long, r As Long
    Dim refKey As Variant
    Dim savePath As String
    Dim curFolder As String
    
    InitConfig
    
    Set wb = ThisWorkbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Validate sheets
    If Not SheetExists(sSheetData, wb) Or Not SheetExists(sSheetRefs, wb) Or Not SheetExists(sSheetTemplate, wb) Then
        MsgBox "One or more configured sheets do not exist. Check configuration.", vbCritical
        GoTo CleanExit
    End If
    If Not SheetExists(sTemplateBackup, wb) Then
        MsgBox "Template backup sheet '" & sTemplateBackup & "' not found. Create a hidden copy of the template named '" & sTemplateBackup & "'.", vbCritical
        GoTo CleanExit
    End If
    
    Set wsData = wb.Worksheets(sSheetData)
    Set wsRefs = wb.Worksheets(sSheetRefs)
    Set wsTemp = wb.Worksheets(sSheetTemplate)
    Set wsBackup = wb.Worksheets(sTemplateBackup)
    
    ' Determine save folder
    If Len(Trim(sSaveFolder)) > 0 Then
        savePath = sSaveFolder
        If Right(savePath, 1) <> "\" Then savePath = savePath & "\"
    Else
        curFolder = wb.Path
        If Len(curFolder) = 0 Then
            MsgBox "Current workbook has not been saved. Please save the workbook first or set sSaveFolder.", vbExclamation
            GoTo CleanExit
        End If
        savePath = curFolder & "\"
    End If
    
    ' Build dictionary of unique Refs and associated Names+Dates from Sheet2
    Set dictRefs = CreateObject("Scripting.Dictionary")
    lastRow = FindLastRowWithValue(wsRefs)
    If lastRow < 2 Then
        MsgBox "No data found on " & sSheetRefs, vbExclamation
        GoTo CleanExit
    End If
    
    Dim colRef As Long, colName As Long, colDate As Long
    colRef = FindHeaderColumn(wsRefs, hdrRef)
    colName = FindHeaderColumn(wsRefs, hdrName)
    colDate = FindHeaderColumn(wsRefs, hdrDate)
    If colRef = 0 Or colName = 0 Or colDate = 0 Then
        MsgBox "Could not find Ref/Name/Date headers on " & sSheetRefs, vbCritical
        GoTo CleanExit
    End If
    
    Dim keyRef As String, nm As String, dtVal As Variant
    For r = 2 To lastRow
        keyRef = Trim(CStr(wsRefs.Cells(r, colRef).Value))
        If Len(keyRef) > 0 Then
            nm = Trim(CStr(wsRefs.Cells(r, colName).Value))
            dtVal = wsRefs.Cells(r, colDate).Value
            If Not dictRefs.Exists(keyRef) Then
                dictRefs.Add keyRef, CreateObject("Scripting.Dictionary")
            End If
            ' For each Name under a Ref, store the threshold date (if multiple rows for same Name under same Ref, last one wins)
            If Len(nm) > 0 Then
                dictRefs(keyRef)(nm) = dtVal
            End If
        End If
    Next r
    
    If dictRefs.Count = 0 Then
        MsgBox "No Refs found on " & sSheetRefs, vbExclamation
        GoTo CleanExit
    End If
    
    ' For each Ref, create report
    Dim refDict As Object, nameKey As Variant
    For Each refKey In dictRefs.Keys
        Set refDict = dictRefs(refKey) ' dictionary of Name -> threshold date
        ' Reset template from backup
        RestoreTemplate wsTemp, wsBackup
        
        ' Build a temporary in-memory list of rows to paste to template
        ' We'll copy matching rows from wsData for each Name and paste into wsTemp starting at first data row after headers
        Dim dataHeaderRow As Long
        dataHeaderRow = FindHeaderRow(wsTemp) ' assumes template has headers; data will be appended below
        If dataHeaderRow = 0 Then dataHeaderRow = 1
        
        Dim pasteRow As Long
        pasteRow = dataHeaderRow + 1
        
        ' Copy matching rows from wsData
        Dim srcHeaderRow As Long
        srcHeaderRow = FindHeaderRow(wsData)
        If srcHeaderRow = 0 Then
            MsgBox "No headers found on data sheet " & sSheetData, vbCritical
            GoTo CleanExit
        End If
        
        ' Map source header columns to indices for the 7 configured srcCols
        Dim srcMap As Object
        Set srcMap = CreateObject("Scripting.Dictionary")
        Dim i As Long, colIndex As Long
        For i = LBound(srcCols) To UBound(srcCols)
            colIndex = FindHeaderColumn(wsData, srcCols(i))
            If colIndex = 0 Then
                MsgBox "Source column '" & srcCols(i) & "' not found on " & sSheetData, vbCritical
                GoTo CleanExit
            End If
            srcMap(i) = colIndex
        Next i
        
        ' Map destination columns on template (to preserve headers and positions)
        Dim dstMap As Object
        Set dstMap = CreateObject("Scripting.Dictionary")
        For i = LBound(dstCols) To UBound(dstCols)
            colIndex = FindHeaderColumn(wsTemp, dstCols(i))
            If colIndex = 0 Then
                MsgBox "Destination column '" & dstCols(i) & "' not found on " & sSheetTemplate, vbCritical
                GoTo CleanExit
            End If
            dstMap(i) = colIndex
        Next i
        
        ' For each Name under this Ref, find matching rows in wsData and copy the 7 columns
        Dim dataLastRow As Long
        dataLastRow = FindLastRowWithValue(wsData)
        Dim foundAny As Boolean
        foundAny = False
        
        For Each nameKey In refDict.Keys
            For r = srcHeaderRow + 1 To dataLastRow
                If Trim(CStr(wsData.Cells(r, srcMap(0)).Value)) = "" Then
                    ' If the first configured source column is the Name column, we used that; otherwise we still compare against the configured Name column
                End If
                ' Compare Name: we assume one of the srcCols corresponds to the Name field; user config must set that accordingly.
                ' We'll compare against the second configured column (srcCols(1)) if that is the Name; if not, user should configure accordingly.
                ' To keep flexible, compare against all 7 source columns for a match to the Name (safe fallback).
                Dim cellVal As String
                cellVal = ""
                Dim matchFound As Boolean
                matchFound = False
                For i = LBound(srcCols) To UBound(srcCols)
                    cellVal = Trim(CStr(wsData.Cells(r, srcMap(i)).Value))
                    If Len(cellVal) > 0 Then
                        If StrComp(cellVal, CStr(nameKey), vbTextCompare) = 0 Then
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next i
                If matchFound Then
                    foundAny = True
                    ' Copy the 7 columns preserving formats and data types
                    For i = LBound(srcCols) To UBound(srcCols)
                        wsTemp.Cells(pasteRow, dstMap(i)).Value = wsData.Cells(r, srcMap(i)).Value
                        ' Preserve number format
                        wsTemp.Cells(pasteRow, dstMap(i)).NumberFormat = wsData.Cells(r, srcMap(i)).NumberFormat
                    Next i
                    pasteRow = pasteRow + 1
                End If
            Next r
        Next nameKey
        
        If Not foundAny Then
            MsgBox "No data found for Ref '" & refKey & "' after matching Names. Skipping report.", vbInformation
            GoTo NextRef
        End If
        
        ' At this point wsTemp contains appended rows. Now perform date parsing, filtering, aggregation.
        Dim dataRange As Range
        Dim firstDataRow As Long, lastDataRowTemp As Long
        firstDataRow = dataHeaderRow + 1
        lastDataRowTemp = FindLastRowWithValue(wsTemp)
        If lastDataRowTemp < firstDataRow Then
            MsgBox "No data present on template after copying for Ref '" & refKey & "'. Skipping.", vbInformation
            GoTo NextRef
        End If
        
        ' Identify which destination column holds the Date and which holds the Name.
        ' We will attempt to find the Date column by header name "Date" on the template; if not found, assume dstCols(2) (3rd) is Date.
        Dim colDateOnTemp As Long, colNameOnTemp As Long
        colDateOnTemp = FindHeaderColumn(wsTemp, "Date")
        If colDateOnTemp = 0 Then
            ' fallback: assume the third configured destination column is the Date (common case)
            colDateOnTemp = dstMap(2) ' zero-based index 2 => third column
        End If
        colNameOnTemp = FindHeaderColumn(wsTemp, "Name")
        If colNameOnTemp = 0 Then
            ' fallback: assume second configured destination column is Name
            colNameOnTemp = dstMap(1)
        End If
        
        ' Build a dictionary keyed by Name|YYYY-MM-DD to aggregate rows
        Dim dictAgg As Object
        Set dictAgg = CreateObject("Scripting.Dictionary")
        Dim rowKey As String
        Dim rowDate As Date
        Dim v As Variant
        Dim colNum As Long
        Dim j As Long
        
        ' Loop through rows and filter out rows where Date < threshold for that Name
        ' We'll collect rows to keep into dictAgg
        For r = firstDataRow To lastDataRowTemp
            v = wsTemp.Cells(r, colNameOnTemp).Value
            If Len(Trim(CStr(v))) = 0 Then
                ' skip rows with no Name
                GoTo ContinueLoop1
            End If
            ' Parse date explicitly
            If IsDate(wsTemp.Cells(r, colDateOnTemp).Value) Then
                rowDate = CDate(wsTemp.Cells(r, colDateOnTemp).Value)
            Else
                ' Try to parse UK formatted text dd/mm/yyyy or dd-mm-yyyy or dd.mm.yyyy
                Dim txtDate As String
                txtDate = Trim(CStr(wsTemp.Cells(r, colDateOnTemp).Value))
                If txtDate = "" Then GoTo ContinueLoop1
                txtDate = Replace(txtDate, "-", "/")
                txtDate = Replace(txtDate, ".", "/")
                If IsDate(txtDate) Then
                    rowDate = CDate(txtDate)
                Else
                    ' invalid date - skip row
                    GoTo ContinueLoop1
                End If
            End If
            
            Dim nameVal As String
            nameVal = Trim(CStr(wsTemp.Cells(r, colNameOnTemp).Value))
            ' Get threshold date for this name from refDict
            Dim threshold As Variant
            If refDict.Exists(nameVal) Then
                threshold = refDict(nameVal)
                If Not IsDate(threshold) Then
                    ' try parse
                    Dim ttxt As String
                    ttxt = Trim(CStr(threshold))
                    ttxt = Replace(ttxt, "-", "/")
                    ttxt = Replace(ttxt, ".", "/")
                    If IsDate(ttxt) Then
                        threshold = CDate(ttxt)
                    Else
                        ' if threshold invalid, treat as very early date (keep all)
                        threshold = DateSerial(1900, 1, 1)
                    End If
                End If
            Else
                ' If name not found in refDict (shouldn't happen) treat threshold as very early
                threshold = DateSerial(1900, 1, 1)
            End If
            
            ' If rowDate < threshold then skip (delete)
            If CDate(rowDate) < CDate(threshold) Then
                GoTo ContinueLoop1
            End If
            
            ' Build key: Name|YYYY-MM-DD
            rowKey = nameVal & "|" & Format(rowDate, "yyyy-mm-dd")
            If Not dictAgg.Exists(rowKey) Then
                ' store an array of values for the 7 columns plus formats
                Dim arrRow As Variant
                ReDim arrRow(1 To 7)
                For j = 0 To 6
                    arrRow(j + 1) = wsTemp.Cells(r, dstMap(j)).Value
                Next j
                dictAgg.Add rowKey, arrRow
            Else
                ' aggregate numeric columns (sumIndices)
                Dim existing As Variant
                existing = dictAgg(rowKey)
                For j = LBound(sumIndices) To UBound(sumIndices)
                    colNum = sumIndices(j) ' 1-based index into the 7 columns
                    ' Ensure numeric addition: coerce to Double if possible
                    Dim v1 As Variant, v2 As Variant
                    v1 = existing(colNum)
                    v2 = wsTemp.Cells(r, dstMap(colNum - 1)).Value
                    If IsNumeric(v1) And IsNumeric(v2) Then
                        existing(colNum) = CDbl(NzZero(v1)) + CDbl(NzZero(v2))
                    ElseIf IsNumeric(v2) And Not IsNumeric(v1) Then
                        existing(colNum) = CDbl(NzZero(v2))
                    Else
                        ' if neither numeric, attempt to convert; if fails, keep existing
                        If IsNumeric(CDbl(NzZero(v2))) Then
                            existing(colNum) = CDbl(NzZero(v1)) + CDbl(NzZero(v2))
                        End If
                    End If
                Next j
                dictAgg(rowKey) = existing
            End If
ContinueLoop1:
        Next r
        
        ' If dictAgg empty => no data after filtering
        If dictAgg.Count = 0 Then
            MsgBox "No data remains for Ref '" & refKey & "' after date filtering. Skipping.", vbInformation
            GoTo NextRef
        End If
        
        ' Clear existing data rows on template and write aggregated rows back (preserve header row)
        ' Clear rows below header
        wsTemp.Range(wsTemp.Rows(firstDataRow), wsTemp.Rows(wsTemp.Rows.Count)).ClearContents
        ' Also clear shapes in data area? We preserve sheet-level shapes separately if needed.
        
        Dim outRow As Long
        outRow = firstDataRow
        Dim k As Variant
        For Each k In dictAgg.Keys
            Dim outArr As Variant
            outArr = dictAgg(k)
            For j = 1 To 7
                wsTemp.Cells(outRow, dstMap(j - 1)).Value = outArr(j)
                ' For numeric columns, set a general number format if not present
                If IsNumeric(outArr(j)) Then
                    ' preserve currency/number formats if source had them; otherwise leave as General
                End If
            Next j
            ' Ensure the Date column is a real date value
            Dim parts() As String
            parts = Split(k, "|")
            If UBound(parts) >= 1 Then
                wsTemp.Cells(outRow, colDateOnTemp).Value = CDate(parts(1))
                wsTemp.Cells(outRow, colDateOnTemp).NumberFormat = "dd/mm/yyyy"
            End If
            outRow = outRow + 1
        Next k
        
        ' Now count distinct dates per Name
        Dim dictCount As Object
        Set dictCount = CreateObject("Scripting.Dictionary")
        For r = firstDataRow To outRow - 1
            nameVal = Trim(CStr(wsTemp.Cells(r, colNameOnTemp).Value))
            If Not dictCount.Exists(nameVal) Then
                dictCount.Add nameVal, CreateObject("Scripting.Dictionary")
            End If
            dictCount(nameVal)(Format(wsTemp.Cells(r, colDateOnTemp).Value, "yyyy-mm-dd")) = 1
        Next r
        
        ' Evaluate MIN_DAYS rule
        Dim onlyOneName As Boolean
        onlyOneName = (dictCount.Count = 1)
        Dim namesToRedo As Object
        Set namesToRedo = CreateObject("Scripting.Dictionary")
        Dim missing As Long
        For Each nameKey In dictCount.Keys
            missing = MIN_DAYS - dictCount(nameKey).Count
            If missing > 0 Then
                namesToRedo.Add nameKey, missing
            End If
        Next nameKey
        
        ' If only one Name and missing > 0 => abort, delete added data and notify
        If onlyOneName And namesToRedo.Count > 0 Then
            Dim onlyName As String
            onlyName = namesToRedo.Keys(0)
            missing = namesToRedo(onlyName)
            Dim redoDate As Date
            redoDate = AddBusinessDays(Date, missing)
            ' Clear data rows we added (restore template)
            RestoreTemplate wsTemp, wsBackup
            MsgBox "Report for Ref '" & refKey & "' aborted. Name '" & onlyName & "' has only " & (MIN_DAYS - missing) & " distinct days." & vbCrLf & _
                   "Suggested re-do date: " & Format(redoDate, "dd/mm/yyyy"), vbExclamation
            GoTo NextRef
        End If
        
        ' Save report as FUN_<Ref>.xlsx
        Dim outFile As String
        outFile = savePath & "FUN_" & refKey & ".xlsx"
        ' Copy template sheet to new workbook and save
        wsTemp.Copy ' this creates a new workbook with the copied sheet as the only sheet
        Set wbOut = ActiveWorkbook
        ' Remove any code modules, keep values/formats as-is; save as xlsx
        Application.DisplayAlerts = False
        wbOut.SaveAs Filename:=outFile, FileFormat:=xlOpenXMLWorkbook
        wbOut.Close SaveChanges:=False
        Application.DisplayAlerts = True
        
        ' After saving, if multiple names and some need redoing, notify
        If Not onlyOneName And namesToRedo.Count > 0 Then
            Dim msg As String
            msg = "Report saved as " & outFile & vbCrLf & "The following Names have fewer than " & MIN_DAYS & " distinct days and need re-doing:" & vbCrLf
            For Each nameKey In namesToRedo.Keys
                missing = namesToRedo(nameKey)
                Dim rd As Date
                rd = AddBusinessDays(Date, missing)
                msg = msg & "- " & nameKey & " needs " & missing & " more day(s). Suggested re-do date: " & Format(rd, "dd/mm/yyyy") & vbCrLf
            Next nameKey
            MsgBox msg, vbInformation, "Partial Data - Needs Re-do"
        Else
            MsgBox "Report saved: " & outFile, vbInformation, "Report Complete"
        End If
        
NextRef:
        ' ensure template restored for next iteration
        RestoreTemplate wsTemp, wsBackup
        ' continue to next ref
    Next refKey

CleanExit:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

' -------------------------
' Helper functions
' -------------------------

' Check if sheet exists
Private Function SheetExists(sName As String, wb As Workbook) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sName) Is Nothing
    On Error GoTo 0
End Function

' Find header row (assumes headers are in a single row; returns row number or 0)
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, lastR As Long, lastC As Long
    lastR = 50 ' assume headers within first 50 rows
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For r = 1 To lastR
        If Application.WorksheetFunction.CountA(ws.Rows(r)) > 0 Then
            ' crude heuristic: first non-empty row with at least one cell that looks like a header
            FindHeaderRow = r
            Exit Function
        End If
    Next r
    FindHeaderRow = 0
End Function

' Find column number by header text (exact match) in the header row
Private Function FindHeaderColumn(ws As Worksheet, headerText As Variant) As Long
    Dim hr As Long, c As Long, lastC As Long
    Dim cellText As String, normCell As String, normHeader As String

    hr = FindHeaderRow(ws)
    If hr = 0 Then Exit Function
    lastC = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column

    ' Normalise headerText: coerce to string, replace CR/LF with space, remove common invisible chars, collapse spaces, trim
    normHeader = CStr(headerText)
    normHeader = Replace(normHeader, vbCrLf, " ")
    normHeader = Replace(normHeader, vbCr, " ")
    normHeader = Replace(normHeader, vbLf, " ")
    normHeader = Replace(normHeader, Chr(160), " ")   ' non-breaking space
    normHeader = Replace(normHeader, vbTab, " ")
    normHeader = Replace(normHeader, ChrW(&H200B), " ") ' zero-width space (if present)
    normHeader = Application.WorksheetFunction.Trim(normHeader) ' collapses multiple spaces
    normHeader = Trim(normHeader)

    For c = 1 To lastC
        cellText = CStr(ws.Cells(hr, c).Value)
        normCell = Replace(cellText, vbCrLf, " ")
        normCell = Replace(normCell, vbCr, " ")
        normCell = Replace(normCell, vbLf, " ")
        normCell = Replace(normCell, Chr(160), " ")
        normCell = Replace(normCell, vbTab, " ")
        normCell = Replace(normCell, ChrW(&H200B), " ")
        normCell = Application.WorksheetFunction.Trim(normCell)
        normCell = Trim(normCell)
        If StrComp(normCell, normHeader, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0
End Function

' Find last used row in a sheet
Private Function FindLastRowWithValue(ws As Worksheet) As Long
    On Error Resume Next
    FindLastRowWithValue = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If FindLastRowWithValue = 0 Then FindLastRowWithValue = 0
    On Error GoTo 0
End Function

' Restore template sheet from hidden backup (copies all formats, shapes, etc.)
Private Sub RestoreTemplate(wsTarget As Worksheet, wsBackup As Worksheet)
    Application.ScreenUpdating = False
    ' Clear target completely then copy used range and shapes from backup
    wsTarget.Cells.Clear
    wsBackup.Cells.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteAll
    ' Copy column widths
    Dim c As Long
    For c = 1 To wsBackup.UsedRange.Columns.Count
        wsTarget.Columns(c).ColumnWidth = wsBackup.Columns(c).ColumnWidth
    Next c
    ' Copy row heights
    Dim r As Long
    For r = 1 To wsBackup.UsedRange.Rows.Count
        wsTarget.Rows(r).RowHeight = wsBackup.Rows(r).RowHeight
    Next r
    ' Copy shapes
    Dim shp As Shape, pastedShp As Shape
    For Each shp In wsBackup.Shapes
        shp.Copy
        wsTarget.Paste
        Set pastedShp = wsTarget.Shapes(wsTarget.Shapes.Count)
        pastedShp.Top = shp.Top
        pastedShp.Left = shp.Left
        pastedShp.Width = shp.Width
        pastedShp.Height = shp.Height
    Next shp
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

' Add business days (skip weekends). n may be positive.
Private Function AddBusinessDays(startDate As Date, n As Long) As Date
    Dim d As Date
    d = startDate
    Dim added As Long
    added = 0
    Do While added < n
        d = d + 1
        If Weekday(d, vbMonday) <= 5 Then added = added + 1
    Loop
    AddBusinessDays = d
End Function

' NzZero: return 0 for empty or non-numeric
Private Function NzZero(v As Variant) As Double
    If IsError(v) Then NzZero = 0: Exit Function
    If IsNumeric(v) Then NzZero = CDbl(v) Else NzZero = 0
End Function
