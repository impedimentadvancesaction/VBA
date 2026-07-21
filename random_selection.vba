Sub SelectRandomRows()

    ' Declare source/destination sheets, counters, and random-selection storage.
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastSrcRow As Long
    Dim lastDstRow As Long
    Dim pickedRows() As Long
    Dim randRow As Long
    Dim i As Integer
    Dim j As Integer
    Dim alreadyPicked As Boolean
    Const NUM_PICKS As Integer = 5

    ' Use active sheet as source and fixed Results sheet as destination.
    Set wsSrc = ActiveSheet
    Set wsDst = ThisWorkbook.Worksheets("Results")

    ' Assume row 1 is a header; data starts at row 2
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' Validate there are enough data rows to pick unique records.
    If lastSrcRow - 1 < NUM_PICKS Then
        MsgBox "Not enough data rows to select " & NUM_PICKS & " unique rows.", vbExclamation
        Exit Sub
    End If

    ' Prepare array that stores selected source row numbers.
    ReDim pickedRows(1 To NUM_PICKS)

    ' Pick 5 unique random row numbers
    Randomize
    i = 0
    Do While i < NUM_PICKS
        randRow = CLng(Int((lastSrcRow - 1) * Rnd() + 2)) ' rows 2 to lastSrcRow

        ' Reject duplicates by checking against already selected rows.
        alreadyPicked = False
        For j = 1 To i
            If pickedRows(j) = randRow Then
                alreadyPicked = True
                Exit For
            End If
        Next j

        ' Accept unique pick and store it.
        If Not alreadyPicked Then
            i = i + 1
            pickedRows(i) = randRow
        End If
    Loop

    ' Find next available row in Results
    lastDstRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row
    If lastDstRow > 1 Or wsDst.Cells(1, 1).Value <> "" Then
        lastDstRow = lastDstRow + 1
    End If

    ' Copy each picked row to Results
    For i = 1 To NUM_PICKS
        wsSrc.Rows(pickedRows(i)).Copy Destination:=wsDst.Rows(lastDstRow)
        lastDstRow = lastDstRow + 1
    Next i

    ' Confirm completion to the user.
    MsgBox NUM_PICKS & " random rows copied to the Results sheet.", vbInformation

End Sub