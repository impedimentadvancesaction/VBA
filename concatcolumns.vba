Option Explicit

Public Sub ConcatenatePairedColumns()

    Const HEADER_ROW As Long = 1
    Const FIRST_HEADER As String = "ColName"
    Const SECOND_HEADER As String = "ColName2"

    Dim ws As Worksheet
    Dim firstCol As Long
    Dim secondCol As Long
    Dim outputCol As Long
    Dim lastRow As Long

    Dim firstData As Variant
    Dim secondData As Variant
    Dim outputData() As Variant

    Dim firstParts As Variant
    Dim secondParts As Variant

    Dim rowIndex As Long
    Dim partIndex As Long
    Dim outputIndex As Long
    Dim outputCount As Long
    Dim pairCount As Long

    Dim firstValue As String
    Dim secondValue As String

    Set ws = ActiveSheet

    firstCol = FindHeaderColumn(ws, FIRST_HEADER, HEADER_ROW)
    secondCol = FindHeaderColumn(ws, SECOND_HEADER, HEADER_ROW)

    If firstCol = 0 Or secondCol = 0 Then
        MsgBox "One or both column headers could not be found.", _
               vbExclamation, "Missing Header"
        Exit Sub
    End If

    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, firstCol).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, secondCol).End(xlUp).Row)

    If lastRow <= HEADER_ROW Then Exit Sub

    firstData = ws.Range( _
        ws.Cells(HEADER_ROW + 1, firstCol), _
        ws.Cells(lastRow, firstCol)).Value2

    secondData = ws.Range( _
        ws.Cells(HEADER_ROW + 1, secondCol), _
        ws.Cells(lastRow, secondCol)).Value2

    'First pass: determine the required output size.
    For rowIndex = 1 To UBound(firstData, 1)

        firstValue = Trim$(CStr(firstData(rowIndex, 1)))
        secondValue = Trim$(CStr(secondData(rowIndex, 1)))

        If Len(firstValue) > 0 And Len(secondValue) > 0 Then

            firstParts = Split(firstValue, ",")
            secondParts = Split(secondValue, ",")

            pairCount = Application.Min( _
                UBound(firstParts) - LBound(firstParts) + 1, _
                UBound(secondParts) - LBound(secondParts) + 1)

            outputCount = outputCount + pairCount

        End If

    Next rowIndex

    If outputCount = 0 Then Exit Sub

    ReDim outputData(1 To outputCount, 1 To 1)

    'Second pass: concatenate matching values.
    For rowIndex = 1 To UBound(firstData, 1)

        firstValue = Trim$(CStr(firstData(rowIndex, 1)))
        secondValue = Trim$(CStr(secondData(rowIndex, 1)))

        If Len(firstValue) > 0 And Len(secondValue) > 0 Then

            firstParts = Split(firstValue, ",")
            secondParts = Split(secondValue, ",")

            pairCount = Application.Min( _
                UBound(firstParts) - LBound(firstParts) + 1, _
                UBound(secondParts) - LBound(secondParts) + 1)

            For partIndex = 0 To pairCount - 1

                outputIndex = outputIndex + 1

                outputData(outputIndex, 1) = _
                    Trim$(CStr(firstParts(partIndex))) & _
                    Trim$(CStr(secondParts(partIndex)))

            Next partIndex

        End If

    Next rowIndex

    'Use the next available worksheet column.
    outputCol = ws.Cells(HEADER_ROW, ws.Columns.Count) _
                  .End(xlToLeft).Column + 1

    'Concatenate the source column names without a space.
    ws.Cells(HEADER_ROW, outputCol).Value = _
        FIRST_HEADER & SECOND_HEADER

    'Write all results in one operation.
    ws.Cells(HEADER_ROW + 1, outputCol) _
      .Resize(outputCount, 1).Value = outputData

End Sub

Private Function FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal headerRow As Long) As Long

    Dim headerCell As Range

    Set headerCell = ws.Rows(headerRow).Find( _
        What:=headerName, _
        After:=ws.Cells(headerRow, ws.Columns.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If Not headerCell Is Nothing Then
        FindHeaderColumn = headerCell.Column
    End If

End Function
