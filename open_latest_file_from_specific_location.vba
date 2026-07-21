Sub ImportTodayFilePreserveAll_PickNewest()
    ' Declare paths, workbook/sheet references, date patterns, and file-selection helpers.
    Dim folderPath As String
    Dim f As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim todayUK1 As String, todayUK2 As String, todayUK3 As String, todayUK4 As String
    Dim bestFile As String
    Dim bestDate As Date
    Dim fiDate As Date
    Dim extList As Variant
    Dim i As Long

    ' Enable centralized error handling and reduce UI overhead while importing.
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Configure source folder and normalize trailing backslash.
    folderPath = "C:\Path\To\Your\Folder\"  ' <-- set your folder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Build accepted UK-style date patterns expected in source file names.
    todayUK1 = Format(Date, "dd-mm-yyyy")
    todayUK2 = Format(Date, "ddmmyyyy")
    todayUK3 = Format(Date, "dd_mm_yyyy")
    todayUK4 = Format(Date, "dd.mm.yyyy")

    ' Restrict search to supported Excel/text file extensions.
    extList = Array("*.xlsx", "*.xlsm", "*.xls", "*.csv") ' restrict extensions

    ' Use active sheet as import target and clear existing content first.
    Set wsDest = ActiveSheet
    wsDest.Cells.Clear

    ' Initialize newest-file tracking state.
    bestFile = ""
    bestDate = #1/1/1900#

    ' Loop through allowed extensions and all files of that type
    For i = LBound(extList) To UBound(extList)
        f = Dir(folderPath & extList(i), vbNormal)
        Do While f <> ""
            ' Keep only files whose names include today's date in supported formats.
            If InStr(1, f, todayUK1, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK2, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK3, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK4, vbTextCompare) > 0 Then

                ' Select the most recently modified matching file.
                fiDate = FileDateTime(folderPath & f)
                If fiDate > bestDate Then
                    bestDate = fiDate
                    bestFile = f
                End If
            End If
            f = Dir()
        Loop
    Next i

    ' Exit early if no dated file for today was found.
    If bestFile = "" Then
        MsgBox "No matching file found for today in " & folderPath, vbExclamation, "Not Found"
        GoTo CleanExit
    End If

    ' Open the newest matching file and read from its first worksheet.
    Set wbSrc = Workbooks.Open(folderPath & bestFile, ReadOnly:=True)
    Set wsSrc = wbSrc.Worksheets(1)

    ' Stop if source sheet has no data to import.
    If Application.WorksheetFunction.CountA(wsSrc.Cells) = 0 Then
        MsgBox "Source sheet is empty: " & bestFile, vbInformation, "Nothing to Import"
        GoTo CloseSource
    End If

    ' Copy full used range including values, formulas, and formatting.
    wsSrc.UsedRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
    ' Replicate source column widths and row heights for layout parity.
    Dim c As Long, r As Long
    For c = 1 To wsSrc.UsedRange.Columns.Count
        wsDest.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
    Next c
    For r = 1 To wsSrc.UsedRange.Rows.Count
        wsDest.Rows(r).RowHeight = wsSrc.Rows(r).RowHeight
    Next r

    ' Copy and reposition shapes so visual elements match the source sheet.
    Dim shp As Shape, pastedShp As Shape
    For Each shp In wsSrc.Shapes
        shp.Copy
        wsDest.Paste
        Set pastedShp = wsDest.Shapes(wsDest.Shapes.Count)
        pastedShp.Top = shp.Top
        pastedShp.Left = shp.Left
        pastedShp.Width = shp.Width
        pastedShp.Height = shp.Height
    Next shp

    MsgBox "Imported from file:" & vbCrLf & bestFile, vbInformation, "Import Complete"

CloseSource:
    ' Close source workbook without saving changes.
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False

CleanExit:
    ' Reset Excel state before exiting.
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' Show error details, attempt safe close, then continue cleanup path.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    If Not wbSrc Is Nothing Then On Error Resume Next: wbSrc.Close SaveChanges:=False
    Resume CleanExit
End Sub
