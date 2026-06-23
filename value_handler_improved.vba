Option Explicit

'======================================================================
' Report generator
' ----------------------------------------------------------------------
' Builds one .xlsx report per Ref (listed on the Refs sheet) from a
' template, populated with the matching rows from the source data sheet.
'
' For each Ref the routine:
'   1. Resets the visible template from a hidden backup copy.
'   2. Copies the matching source rows into the template.
'   3. Drops rows dated before each Name's threshold date.
'   4. Aggregates duplicate Name+Date rows (summing the configured cols).
'   5. Checks each Name has at least MIN_DAYS distinct dates.
'   6. Saves the result as FUN_<Ref>.xlsx and reports any Names that
'      still need more days.
'
' Behaviour is identical to the original module, with two exceptions,
' both intentional and noted inline:
'   * The Name from the Refs sheet is now searched ONLY in the column
'     given by NAME_COL_INDEX (previously it matched against any of the
'     7 columns - see the Name match loop).  [Requested change]
'   * RestoreTemplate now deletes existing template shapes before
'     re-adding the backup's shapes, so shapes no longer accumulate on
'     every restore.                                          [Defect fix]
'
' Performance: the source rows for each Name are now found via an
' in-memory index built once, instead of re-scanning the whole data
' sheet for every Name.  The column maps are also built once rather than
' per Ref.  The matched rows, date parsing, aggregation and write-back
' are unchanged, so the output is the same.
'======================================================================


'======================================================================
' Configuration - EDIT ONLY
'======================================================================
Private Const sSheetData     As String = "Sheet1"          ' source import data sheet
Private Const sSheetRefs     As String = "Sheet2"          ' sheet with Ref, Name, Date rows
Private Const sSheetTemplate As String = "Sheet3"          ' report template sheet (visible)
Private Const sTemplateBackup As String = "TemplateBackup" ' hidden copy of template (must exist)
Private Const sSaveFolder    As String = ""                ' "" = current workbook folder; or full path with trailing "\"

' Column header names on the Refs sheet (Sheet2)
Private Const hdrRef  As String = "Ref"
Private Const hdrName As String = "Name"
Private Const hdrDate As String = "Date"

' Minimum distinct dates required per Name
Private Const MIN_DAYS As Long = 7

' Which of the 7 configured columns holds the Name (0-based position
' into srcCols / dstCols). The Name from the Refs sheet is searched ONLY
' in this column on the data sheet.
Private Const NAME_COL_INDEX As Long = 1   ' position 1 = the 2nd column ("ColumnName2")

' Which of the 7 configured columns holds the Date (0-based position into
' srcCols / dstCols). Used for the pre-flight date check and the template
' Date-column fallback.
Private Const DATE_COL_INDEX As Long = 2   ' position 2 = the 3rd column ("ColumnName3")

' How TEXT dates are written on each sheet (real Excel date cells are
' read directly and are unaffected by this). Set each to match its sheet.
'   DORDER_DMY = day/month/year  (UK, e.g. 31/12/2025)
'   DORDER_MDY = month/day/year  (US, e.g. 12/31/2025)
Private Const DORDER_DMY As Long = 1
Private Const DORDER_MDY As Long = 2
Private Const DATA_DATE_ORDER As Long = DORDER_MDY   ' Sheet1 = US mm/dd/yyyy
Private Const REFS_DATE_ORDER As Long = DORDER_DMY   ' Sheet2 = UK dd/mm/yyyy

' Number format applied to the Date column in the output report (Sheet3).
' This is an Excel cell number-format string (not a VBA Format() code).
Private Const OUTPUT_DATE_NUMFORMAT As String = "[$-en-US,1]mm/dd/yyyy"  ' US display

' VBA Format() pattern used for dates shown inside message boxes
' (e.g. suggested re-do dates). Plain pattern, no locale tag.
Private Const MSG_DATE_FORMAT As String = "mm/dd/yyyy"

' The 7 columns to copy, in order. These arrays are filled in InitConfig.
Private srcCols    As Variant  ' source header names on the data sheet
Private dstCols    As Variant  ' destination header names on the template (same order)
Private sumIndices As Variant  ' 1-based positions of the columns to sum on duplicate Date+Name
'======================================================================
' End Configuration
'======================================================================


' Populate the configuration arrays. Edit these to match your headers.
Private Sub InitConfig()
    ' The 7 source headers (data sheet), in the order they should be copied.
    srcCols = Array("ColumnName1", "ColumnName2", "ColumnName3", "ColumnName4", "ColumnName5", "ColumnName6", "ColumnName7")
    ' The 7 destination headers (template), in the same order.
    dstCols = Array("ColumnName1", "ColumnName2", "ColumnName3", "ColumnName4", "ColumnName5", "ColumnName6", "ColumnName7")
    ' Which of the 7 columns are numeric and get summed on duplicate Date+Name (1-based).
    sumIndices = Array(3, 4, 5, 6, 7)
End Sub


'======================================================================
' Main routine
'======================================================================
Public Sub CreateReportsPerRef()

    ' --- Workbook / sheet objects ---
    Dim wb As Workbook
    Dim wsData As Worksheet, wsRefs As Worksheet
    Dim wsTemp As Worksheet, wsBackup As Worksheet

    ' --- Save location ---
    Dim savePath As String, curFolder As String

    ' --- Refs structure: Ref -> (Name -> threshold date) ---
    Dim dictRefs As Object, refDict As Object
    Dim colRef As Long, colName As Long, colDate As Long
    Dim refKey As Variant, nameKey As Variant
    Dim keyRef As String, nm As String, dtVal As Variant

    ' --- Column maps (built once) ---
    Dim srcMap As Object, dstMap As Object       ' position (0..6) -> actual column number
    Dim srcHeaderRow As Long, dataLastRow As Long
    Dim dataHeaderRow As Long, firstDataRow As Long
    Dim colDateOnTemp As Long, colNameOnTemp As Long
    Dim colIndex As Long

    ' --- Name -> data rows index (built once) ---
    Dim nameIdx As Object, rowNum As Variant

    ' --- Per-Ref working values ---
    Dim pasteRow As Long, foundAny As Boolean
    Dim lastRow As Long, r As Long, i As Long, j As Long

    InitConfig

    Set wb = ThisWorkbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' --- Validate that all configured sheets exist ---
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

    ' --- Work out the save folder ---
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

    ' --- Build the Refs dictionary from the Refs sheet ---
    ' Structure: dictRefs(Ref) = dictionary of (Name -> threshold date).
    Set dictRefs = CreateObject("Scripting.Dictionary")
    lastRow = FindLastRowWithValue(wsRefs)
    If lastRow < 2 Then
        MsgBox "No data found on " & sSheetRefs, vbExclamation
        GoTo CleanExit
    End If

    colRef = FindHeaderColumn(wsRefs, hdrRef)
    colName = FindHeaderColumn(wsRefs, hdrName)
    colDate = FindHeaderColumn(wsRefs, hdrDate)
    If colRef = 0 Or colName = 0 Or colDate = 0 Then
        MsgBox "Could not find Ref/Name/Date headers on " & sSheetRefs, vbCritical
        GoTo CleanExit
    End If

    For r = 2 To lastRow
        keyRef = Trim(CStr(wsRefs.Cells(r, colRef).Value))
        If Len(keyRef) > 0 Then
            nm = Trim(CStr(wsRefs.Cells(r, colName).Value))
            dtVal = wsRefs.Cells(r, colDate).Value
            If Not dictRefs.Exists(keyRef) Then
                dictRefs.Add keyRef, CreateObject("Scripting.Dictionary")
            End If
            ' Store the threshold date per Name. If the same Name appears
            ' more than once under a Ref, the last row wins (as before).
            If Len(nm) > 0 Then
                dictRefs(keyRef)(nm) = dtVal
            End If
        End If
    Next r

    If dictRefs.Count = 0 Then
        MsgBox "No Refs found on " & sSheetRefs, vbExclamation
        GoTo CleanExit
    End If

    '------------------------------------------------------------------
    ' Build the lookups that do not change between Refs ONCE, up front.
    ' (The data sheet is constant, and the template is reset to the same
    ' backup for every Ref, so its header positions never change.)
    '------------------------------------------------------------------

    ' Reset the template so its headers are in their clean positions
    ' before we map the destination columns.
    RestoreTemplate wsTemp, wsBackup

    ' Locate the header row on the data sheet.
    srcHeaderRow = FindHeaderRow(wsData)
    If srcHeaderRow = 0 Then
        MsgBox "No headers found on data sheet " & sSheetData, vbCritical
        GoTo CleanExit
    End If
    dataLastRow = FindLastRowWithValue(wsData)

    ' Map each of the 7 source headers to its column number on the data sheet.
    Set srcMap = CreateObject("Scripting.Dictionary")
    For i = LBound(srcCols) To UBound(srcCols)
        colIndex = FindHeaderColumn(wsData, srcCols(i))
        If colIndex = 0 Then
            MsgBox "Source column '" & srcCols(i) & "' not found on " & sSheetData, vbCritical
            GoTo CleanExit
        End If
        srcMap(i) = colIndex
    Next i

    ' Map each of the 7 destination headers to its column number on the template.
    Set dstMap = CreateObject("Scripting.Dictionary")
    For i = LBound(dstCols) To UBound(dstCols)
        colIndex = FindHeaderColumn(wsTemp, dstCols(i))
        If colIndex = 0 Then
            MsgBox "Destination column '" & dstCols(i) & "' not found on " & sSheetTemplate, vbCritical
            GoTo CleanExit
        End If
        dstMap(i) = colIndex
    Next i

    ' Header row / first data row on the template, and the Date / Name
    ' columns. We look these up by header text and fall back to fixed
    ' positions exactly as the original did.
    dataHeaderRow = FindHeaderRow(wsTemp)
    If dataHeaderRow = 0 Then dataHeaderRow = 1
    firstDataRow = dataHeaderRow + 1

    colDateOnTemp = FindHeaderColumn(wsTemp, "Date")
    If colDateOnTemp = 0 Then colDateOnTemp = dstMap(DATE_COL_INDEX) ' fallback: configured Date column
    colNameOnTemp = FindHeaderColumn(wsTemp, "Name")
    If colNameOnTemp = 0 Then colNameOnTemp = dstMap(NAME_COL_INDEX) ' fallback: the configured Name column

    '------------------------------------------------------------------
    ' Pre-flight date check.
    ' Scan the Date column on the data sheet (US order) AND the Refs sheet
    ' (UK order) for any value that is not a valid date in its sheet's
    ' format, and report them up front. This means nothing is ever silently
    ' mis-parsed: you see exactly which cells are wrong before any report
    ' is built. (Real Excel date cells always pass - their format is just
    ' display.)
    '------------------------------------------------------------------
    Dim dateReport As String, badCount As Long, resp As VbMsgBoxResult
    dateReport = ""
    badCount = AuditDateColumn(wsData, srcMap(DATE_COL_INDEX), DATA_DATE_ORDER, dateReport)
    badCount = badCount + AuditDateColumn(wsRefs, colDate, REFS_DATE_ORDER, dateReport)
    If badCount > 0 Then
        resp = MsgBox(badCount & " cell(s) are not a valid date for their sheet" & vbCrLf & _
                      "(Sheet1 expects US mm/dd/yyyy, Sheet2 expects UK dd/mm/yyyy):" & vbCrLf & vbCrLf & _
                      dateReport & vbCrLf & _
                      "Rows with a bad date will be skipped, and a bad threshold date keeps all rows." & vbCrLf & _
                      "Continue anyway?", vbExclamation + vbYesNo, "Date check")
        If resp <> vbYes Then GoTo CleanExit
    End If

    ' Index the data rows by Name once. This replaces the original
    ' scan-the-whole-sheet-per-Name loop and is the main speed-up.
    Set nameIdx = BuildNameRowIndex(wsData, srcMap(NAME_COL_INDEX), srcHeaderRow + 1, dataLastRow)

    ' Collects each saved-report outcome so a single summary can be shown
    ' once at the end, instead of one pop-up per report.
    Dim summary As String
    summary = ""

    '------------------------------------------------------------------
    ' Produce one report per Ref.
    '------------------------------------------------------------------
    For Each refKey In dictRefs.Keys
        Set refDict = dictRefs(refKey)        ' Name -> threshold date for this Ref

        ' The template is already clean here: it is restored once before the
        ' loop, and again after every Ref at NextRef, so this iteration is
        ' already covered. No duplicate restore needed.
        pasteRow = firstDataRow

        '--------------------------------------------------------------
        ' Copy the matching source rows into the template.
        ' A row matches a Name only via the single Name column
        ' (NAME_COL_INDEX); the nameIdx gives us its rows directly.
        '--------------------------------------------------------------
        foundAny = False
        For Each nameKey In refDict.Keys
            If nameIdx.Exists(CStr(nameKey)) Then
                For Each rowNum In nameIdx(CStr(nameKey))
                    foundAny = True
                    ' Copy the 7 columns, preserving each cell's number format.
                    For i = LBound(srcCols) To UBound(srcCols)
                        wsTemp.Cells(pasteRow, dstMap(i)).Value = wsData.Cells(rowNum, srcMap(i)).Value
                        wsTemp.Cells(pasteRow, dstMap(i)).NumberFormat = wsData.Cells(rowNum, srcMap(i)).NumberFormat
                    Next i
                    pasteRow = pasteRow + 1
                Next rowNum
            End If
        Next nameKey

        If Not foundAny Then
            MsgBox "No data found for Ref '" & refKey & "' after matching Names. Skipping report.", vbInformation
            GoTo NextRef
        End If

        '--------------------------------------------------------------
        ' Filter by threshold date and aggregate duplicate Name+Date rows.
        '--------------------------------------------------------------
        Dim lastDataRowTemp As Long
        lastDataRowTemp = FindLastRowWithValue(wsTemp)
        If lastDataRowTemp < firstDataRow Then
            MsgBox "No data present on template after copying for Ref '" & refKey & "'. Skipping.", vbInformation
            GoTo NextRef
        End If

        ' dictAgg: key "Name|yyyy-mm-dd" -> array of the 7 column values.
        Dim dictAgg As Object
        Set dictAgg = CreateObject("Scripting.Dictionary")
        Dim rowKey As String, rowDate As Date
        Dim nameVal As String, threshold As Date
        Dim arrRow As Variant, existing As Variant
        Dim colNum As Long

        For r = firstDataRow To lastDataRowTemp

            ' Skip rows with no Name.
            If Len(Trim(CStr(wsTemp.Cells(r, colNameOnTemp).Value))) = 0 Then GoTo ContinueLoop1

            ' Parse the row date. It was copied from the data sheet, so it
            ' follows the data-sheet order (US) when stored as text; real
            ' date cells are read directly. No reliance on regional settings.
            ' Unparseable -> skip (already listed by the pre-flight check).
            If Not TryParseDate(wsTemp.Cells(r, colDateOnTemp).Value, DATA_DATE_ORDER, rowDate) Then GoTo ContinueLoop1

            nameVal = Trim(CStr(wsTemp.Cells(r, colNameOnTemp).Value))

            ' Threshold date for this Name (from the Refs sheet, UK order).
            ' If it is missing or unparseable, treat it as a very early date
            ' so the row is kept.
            If refDict.Exists(nameVal) Then
                If Not TryParseDate(refDict(nameVal), REFS_DATE_ORDER, threshold) Then
                    threshold = DateSerial(1900, 1, 1)
                End If
            Else
                threshold = DateSerial(1900, 1, 1)
            End If

            ' Drop rows dated before the threshold (both are date-only).
            If rowDate < threshold Then GoTo ContinueLoop1

            ' Aggregate by Name + date.
            rowKey = nameVal & "|" & Format(rowDate, "yyyy-mm-dd")
            If Not dictAgg.Exists(rowKey) Then
                ' First time we have seen this Name+date: store the 7 values.
                ReDim arrRow(1 To 7)
                For j = 0 To 6
                    arrRow(j + 1) = wsTemp.Cells(r, dstMap(j)).Value
                Next j
                dictAgg.Add rowKey, arrRow
            Else
                ' Seen before: add the summed columns together. NzZero
                ' turns blanks / non-numerics into 0, so this reproduces
                ' all the cases the original handled explicitly.
                existing = dictAgg(rowKey)
                For j = LBound(sumIndices) To UBound(sumIndices)
                    colNum = sumIndices(j)   ' 1-based position within the 7 columns
                    existing(colNum) = NzZero(existing(colNum)) + _
                                       NzZero(wsTemp.Cells(r, dstMap(colNum - 1)).Value)
                Next j
                dictAgg(rowKey) = existing
            End If
ContinueLoop1:
        Next r

        If dictAgg.Count = 0 Then
            MsgBox "No data remains for Ref '" & refKey & "' after date filtering. Skipping.", vbInformation
            GoTo NextRef
        End If

        '--------------------------------------------------------------
        ' Write the aggregated rows back to the template.
        '--------------------------------------------------------------
        ' Clear contents below the header (formats stay, as before).
        wsTemp.Range(wsTemp.Rows(firstDataRow), wsTemp.Rows(wsTemp.Rows.Count)).ClearContents

        Dim outRow As Long, outArr As Variant, k As Variant
        Dim parts() As String, iso() As String
        outRow = firstDataRow
        For Each k In dictAgg.Keys
            outArr = dictAgg(k)
            For j = 1 To 7
                wsTemp.Cells(outRow, dstMap(j - 1)).Value = outArr(j)
            Next j
            ' Write the Date column back as a real date value. The key holds
            ' the date as yyyy-mm-dd, so build it with DateSerial (no CDate,
            ' no locale dependence) and display it with the configured
            ' output number format (US for Sheet3).
            parts = Split(k, "|")
            If UBound(parts) >= 1 Then
                iso = Split(parts(1), "-")
                wsTemp.Cells(outRow, colDateOnTemp).Value = DateSerial(CLng(iso(0)), CLng(iso(1)), CLng(iso(2)))
                wsTemp.Cells(outRow, colDateOnTemp).NumberFormat = OUTPUT_DATE_NUMFORMAT
            End If
            outRow = outRow + 1
        Next k

        '--------------------------------------------------------------
        ' Count distinct dates per Name and apply the MIN_DAYS rule.
        '--------------------------------------------------------------
        Dim dictCount As Object
        Set dictCount = CreateObject("Scripting.Dictionary")
        For r = firstDataRow To outRow - 1
            nameVal = Trim(CStr(wsTemp.Cells(r, colNameOnTemp).Value))
            If Not dictCount.Exists(nameVal) Then
                dictCount.Add nameVal, CreateObject("Scripting.Dictionary")
            End If
            dictCount(nameVal)(Format(wsTemp.Cells(r, colDateOnTemp).Value, "yyyy-mm-dd")) = 1
        Next r

        Dim onlyOneName As Boolean
        onlyOneName = (dictCount.Count = 1)

        ' namesToRedo: Name -> number of additional distinct days needed.
        Dim namesToRedo As Object
        Set namesToRedo = CreateObject("Scripting.Dictionary")
        Dim missing As Long
        For Each nameKey In dictCount.Keys
            missing = MIN_DAYS - dictCount(nameKey).Count
            If missing > 0 Then
                namesToRedo.Add nameKey, missing
            End If
        Next nameKey

        ' Single Name short of MIN_DAYS => abort this Ref, clear the data,
        ' and suggest a re-do date.
        If onlyOneName And namesToRedo.Count > 0 Then
            Dim onlyName As String, redoDate As Date
            onlyName = namesToRedo.Keys(0)
            missing = namesToRedo(onlyName)
            redoDate = AddBusinessDays(Date, missing)
            RestoreTemplate wsTemp, wsBackup   ' discard the rows we added
            MsgBox "Report for Ref '" & refKey & "' aborted. Name '" & onlyName & "' has only " & (MIN_DAYS - missing) & " distinct days." & vbCrLf & _
                   "Suggested re-do date: " & Format(redoDate, MSG_DATE_FORMAT), vbExclamation
            GoTo NextRef
        End If

        '--------------------------------------------------------------
        ' Save the report as FUN_<Ref>.xlsx.
        '--------------------------------------------------------------
        Dim outFile As String
        outFile = savePath & "FUN_" & refKey & ".xlsx"
        SaveReport wsTemp, outFile

        ' Record the outcome for the end-of-run summary instead of showing a
        ' pop-up for every saved report.
        If Not onlyOneName And namesToRedo.Count > 0 Then
            Dim rd As Date
            summary = summary & "Report saved as " & outFile & vbCrLf & _
                      "  The following Names have fewer than " & MIN_DAYS & " distinct days and need re-doing:" & vbCrLf
            For Each nameKey In namesToRedo.Keys
                missing = namesToRedo(nameKey)
                rd = AddBusinessDays(Date, missing)
                summary = summary & "  - " & nameKey & " needs " & missing & " more day(s). Suggested re-do date: " & Format(rd, MSG_DATE_FORMAT) & vbCrLf
            Next nameKey
            summary = summary & vbCrLf
        Else
            summary = summary & "Report saved: " & outFile & vbCrLf
        End If

NextRef:
        ' Leave the template clean for the next Ref.
        RestoreTemplate wsTemp, wsBackup
    Next refKey

    ' Single summary for the whole run, shown only after every report has
    ' been processed (not once per saved report).
    If Len(summary) > 0 Then
        MsgBox summary, vbInformation, "Reports complete"
    End If

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


'======================================================================
' Helper procedures
'======================================================================

' Return True if a sheet with the given name exists in the workbook.
Private Function SheetExists(sName As String, wb As Workbook) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sName) Is Nothing
    On Error GoTo 0
End Function


' Build a lookup of data-sheet rows keyed by Name, so a Name's rows can
' be fetched directly instead of scanning the whole sheet for each Name.
' Names are compared case-insensitively, matching the original
' StrComp(..., vbTextCompare) behaviour. Each value is a Collection of
' the actual sheet row numbers, in ascending order.
Private Function BuildNameRowIndex(ws As Worksheet, nameCol As Long, _
                                   firstRow As Long, lastRow As Long) As Object
    Dim idx As Object, vals As Variant
    Dim i As Long, key As String

    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare        ' case-insensitive keys

    If lastRow < firstRow Then
        Set BuildNameRowIndex = idx
        Exit Function
    End If

    ' Read the whole Name column in one go (fast), then index it.
    vals = ws.Range(ws.Cells(firstRow, nameCol), ws.Cells(lastRow, nameCol)).Value

    If IsArray(vals) Then
        For i = 1 To UBound(vals, 1)
            key = Trim(CStr(vals(i, 1)))
            If Len(key) > 0 Then
                If Not idx.Exists(key) Then idx.Add key, New Collection
                idx(key).Add firstRow + i - 1     ' real sheet row
            End If
        Next i
    Else
        ' Single-row range: .Value is a scalar, not an array.
        key = Trim(CStr(vals))
        If Len(key) > 0 Then
            idx.Add key, New Collection
            idx(key).Add firstRow
        End If
    End If

    Set BuildNameRowIndex = idx
End Function


' Find the header row: the first non-empty row within the first 50 rows.
' (Simple heuristic, kept as-is so column lookups behave as before.)
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, lastR As Long
    lastR = 50
    For r = 1 To lastR
        If Application.WorksheetFunction.CountA(ws.Rows(r)) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
    FindHeaderRow = 0
End Function


' Find a column by header text in the header row. The comparison ignores
' line breaks, tabs, non-breaking / zero-width spaces, collapses runs of
' spaces and is case-insensitive.
Private Function FindHeaderColumn(ws As Worksheet, headerText As Variant) As Long
    Dim hr As Long, c As Long, lastC As Long
    Dim cellText As String, normCell As String, normHeader As String

    hr = FindHeaderRow(ws)
    If hr = 0 Then Exit Function
    lastC = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column

    ' Normalise the header we are looking for.
    normHeader = CStr(headerText)
    normHeader = Replace(normHeader, vbCrLf, " ")
    normHeader = Replace(normHeader, vbCr, " ")
    normHeader = Replace(normHeader, vbLf, " ")
    normHeader = Replace(normHeader, Chr(160), " ")        ' non-breaking space
    normHeader = Replace(normHeader, vbTab, " ")
    normHeader = Replace(normHeader, ChrW(&H200B), " ")    ' zero-width space
    normHeader = Application.WorksheetFunction.Trim(normHeader)  ' collapse multiple spaces
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


' Find the last used row in a sheet (0 if the sheet is empty).
Private Function FindLastRowWithValue(ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not foundCell Is Nothing Then FindLastRowWithValue = foundCell.Row
End Function


' Reset the visible template from the hidden backup, including cell
' contents, formats, column widths, row heights and shapes.
Private Sub RestoreTemplate(wsTarget As Worksheet, wsBackup As Worksheet)
    Dim c As Long, r As Long, si As Long
    Dim shp As Shape, pastedShp As Shape
    Dim prevActive As Object

    Application.ScreenUpdating = False

    ' Clear the target completely, then copy the backup's cells back in.
    wsTarget.Cells.Clear
    wsBackup.Cells.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteAll

    ' Match column widths to the backup.
    For c = 1 To wsBackup.UsedRange.Columns.Count
        wsTarget.Columns(c).ColumnWidth = wsBackup.Columns(c).ColumnWidth
    Next c

    ' Match row heights to the backup.
    For r = 1 To wsBackup.UsedRange.Rows.Count
        wsTarget.Rows(r).RowHeight = wsBackup.Rows(r).RowHeight
    Next r

    ' Cell copy does not carry shapes, so handle them separately.
    ' Remove any shapes already on the target first, otherwise they would
    ' accumulate every time the template is restored.
    For si = wsTarget.Shapes.Count To 1 Step -1
        wsTarget.Shapes(si).Delete
    Next si

    ' Worksheet.Paste needs the destination sheet to be active, so make it
    ' active for the duration and restore the previous selection after.
    Set prevActive = ActiveSheet
    wsTarget.Activate
    For Each shp In wsBackup.Shapes
        shp.Copy
        wsTarget.Paste
        Set pastedShp = wsTarget.Shapes(wsTarget.Shapes.Count)
        pastedShp.Top = shp.Top
        pastedShp.Left = shp.Left
        pastedShp.Width = shp.Width
        pastedShp.Height = shp.Height
    Next shp
    On Error Resume Next
    prevActive.Activate
    On Error GoTo 0

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub


' Copy the populated template into a new single-sheet workbook and save
' it as a macro-free .xlsx, then close it.
Private Sub SaveReport(wsTemp As Worksheet, ByVal outFile As String)
    Dim wbOut As Workbook
    wsTemp.Copy                       ' creates a new workbook holding only this sheet
    Set wbOut = ActiveWorkbook
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=outFile, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub


' Add n business days to a date, skipping Saturdays and Sundays.
Private Function AddBusinessDays(startDate As Date, n As Long) As Date
    Dim d As Date, added As Long
    d = startDate
    Do While added < n
        d = d + 1
        If Weekday(d, vbMonday) <= 5 Then added = added + 1
    Loop
    AddBusinessDays = d
End Function


' Return a value as a Double; blanks, errors and non-numerics become 0.
Private Function NzZero(v As Variant) As Double
    If IsError(v) Then NzZero = 0: Exit Function
    If IsNumeric(v) Then NzZero = CDbl(v) Else NzZero = 0
End Function


' Parse a date WITHOUT relying on the machine's regional settings.
' Returns True and sets outDate on success; False for blanks or anything
' that is not a valid date in the requested order.
'
' order = DORDER_DMY -> UK order dd/mm/yyyy (used for the Refs sheet)
' order = DORDER_MDY -> US order mm/dd/yyyy (used for the data sheet)
'
' Accepts:
'   * A genuine Excel date cell (stored as a serial) - used as-is, since a
'     real date serial is locale-independent; the display format does not
'     matter, only text dates carry an order ambiguity.
'   * Text with "/", "-" or "." separators, in the requested order.
'     2-digit years follow Excel's convention (00-29 -> 2000-2029,
'     30-99 -> 1930-1999). Any time component is dropped.
'
' Impossible dates such as 31/02/2025 (UK) or 02/31/2025 (US) are
' rejected: the value is rebuilt with DateSerial and then checked to
' confirm the day, month and year survived unchanged.
Private Function TryParseDate(ByVal v As Variant, ByVal order As Long, _
                              ByRef outDate As Date) As Boolean
    Dim s As String
    Dim p() As String
    Dim d As Long, m As Long, y As Long
    Dim chk As Date

    TryParseDate = False

    ' A real date cell comes through as a Date variant - trust it directly.
    ' This is the unambiguous case: the display format (US or UK) and the
    ' 'order' argument are irrelevant for a genuine date serial.
    If VarType(v) = vbDate Then
        outDate = Int(CDate(v))            ' strip any time component
        TryParseDate = True
        Exit Function
    End If

    ' Otherwise treat it as text and parse explicitly in the given order.
    s = Trim(CStr(v))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, "-", "/")
    s = Replace(s, ".", "/")
    s = Replace(s, " ", "")                ' drop stray spaces
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function   ' need exactly three components (year last)

    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function

    If order = DORDER_DMY Then
        d = CLng(p(0)): m = CLng(p(1))     ' UK: day / month / year
    Else
        m = CLng(p(0)): d = CLng(p(1))     ' US: month / day / year
    End If
    y = CLng(p(2))

    ' Expand 2-digit years (Excel convention).
    If y < 100 Then
        If y < 30 Then y = 2000 + y Else y = 1900 + y
    End If

    ' Range checks before constructing the date.
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    If y < 1900 Or y > 9999 Then Exit Function

    ' Build and verify the round-trip to reject impossible dates. This also
    ' catches text that contradicts the configured order (e.g. 13/04/2025
    ' in a US column -> month 13 -> rejected rather than silently swapped).
    chk = DateSerial(y, m, d)
    If Day(chk) <> d Or Month(chk) <> m Or Year(chk) <> y Then Exit Function

    outDate = chk
    TryParseDate = True
End Function


' Scan one sheet's date column for values that are not valid dates in the
' given order (DORDER_DMY for UK, DORDER_MDY for US). Appends up to 15
' "Sheet!cell = 'value'" lines to 'report' and returns the problem count.
Private Function AuditDateColumn(ws As Worksheet, dateCol As Long, _
                                 ByVal order As Long, ByRef report As String) As Long
    Dim hdr As Long, firstRow As Long, lastRow As Long, r As Long
    Dim count As Long
    Dim v As Variant, d As Date
    Const MAX_LIST As Long = 15

    hdr = FindHeaderRow(ws)
    If hdr = 0 Then Exit Function
    firstRow = hdr + 1
    lastRow = FindLastRowWithValue(ws)

    For r = firstRow To lastRow
        v = ws.Cells(r, dateCol).Value
        If Len(Trim(CStr(v))) > 0 Then        ' ignore blank cells
            If Not TryParseDate(v, order, d) Then
                count = count + 1
                If count <= MAX_LIST Then
                    report = report & "  " & ws.Name & "!" & _
                             ws.Cells(r, dateCol).Address(False, False) & _
                             "  =  '" & CStr(v) & "'" & vbCrLf
                End If
            End If
        End If
    Next r

    If count > MAX_LIST Then
        report = report & "  ... and " & (count - MAX_LIST) & " more on " & ws.Name & vbCrLf
    End If
    AuditDateColumn = count
End Function
