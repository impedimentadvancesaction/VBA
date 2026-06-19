Sub ImportTodayFilePreserveAll_PickNewest()
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

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    folderPath = "C:\Path\To\Your\Folder\"  ' <-- set your folder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    todayUK1 = Format(Date, "dd-mm-yyyy")
    todayUK2 = Format(Date, "ddmmyyyy")
    todayUK3 = Format(Date, "dd_mm_yyyy")
    todayUK4 = Format(Date, "dd.mm.yyyy")

    extList = Array("*.xlsx", "*.xlsm", "*.xls", "*.csv") ' restrict extensions

    Set wsDest = ActiveSheet
    wsDest.Cells.Clear

    bestFile = ""
    bestDate = #1/1/1900#

    ' Loop through allowed extensions and all files of that type
    For i = LBound(extList) To UBound(extList)
        f = Dir(folderPath & extList(i), vbNormal)
        Do While f <> ""
            If InStr(1, f, todayUK1, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK2, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK3, vbTextCompare) > 0 _
               Or InStr(1, f, todayUK4, vbTextCompare) > 0 Then

                fiDate = FileDateTime(folderPath & f)
                If fiDate > bestDate Then
                    bestDate = fiDate
                    bestFile = f
                End If
            End If
            f = Dir()
        Loop
    Next i

    If bestFile = "" Then
        MsgBox "No matching file found for today in " & folderPath, vbExclamation, "Not Found"
        GoTo CleanExit
    End If

    Set wbSrc = Workbooks.Open(folderPath & bestFile, ReadOnly:=True)
    Set wsSrc = wbSrc.Worksheets(1)

    If Application.WorksheetFunction.CountA(wsSrc.Cells) = 0 Then
        MsgBox "Source sheet is empty: " & bestFile, vbInformation, "Nothing to Import"
        GoTo CloseSource
    End If

    wsSrc.UsedRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
    ' copy column widths
    Dim c As Long, r As Long
    For c = 1 To wsSrc.UsedRange.Columns.Count
        wsDest.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
    Next c
    For r = 1 To wsSrc.UsedRange.Rows.Count
        wsDest.Rows(r).RowHeight = wsSrc.Rows(r).RowHeight
    Next r

    ' copy shapes
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
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False

CleanExit:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    If Not wbSrc Is Nothing Then On Error Resume Next: wbSrc.Close SaveChanges:=False
    Resume CleanExit
End Sub
