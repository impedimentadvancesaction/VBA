Option Explicit

' Cleans and reshapes data from the "Samp Fixer" sheet into separate tabs by a chosen category.
' Then highlights random rows in each new tab and filters each tab to show only highlighted records.

' =======================================================================
' MAIN SUBROUTINE
' =======================================================================
Public Sub RunSampFixerProcess()
    ' Turn off application features for performance speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets("Samp Fixer")
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "The tab 'Samp Fixer' could not be found.", vbCritical
        GoTo Cleanup
    End If

    ' ==========================================
    ' --- USER CONFIGURATION START ---
    ' ==========================================
    
    ' 1. Columns to delete by name
    Dim colsToDelete As Variant
    colsToDelete = Array("DummyCol1", "DummyCol2", "DummyCol3")
    
    ' 2. Column name and the two values that should trigger an entire row deletion
    Dim filterColStep1 As String: filterColStep1 = "TargetColumnName1"
    Dim delValue1 As String: delValue1 = "ValueToDeleteA"
    Dim delValue2 As String: delValue2 = "ValueToDeleteB"
    
    ' 3. Column name containing the unique values to create tabs for
    Dim splitColStep2 As String: splitColStep2 = "TargetColumnName2"
    
    ' ==========================================
    ' --- USER CONFIGURATION END ---
    ' ==========================================

    ' Execute Steps in Order
    Call Step1_CleanData(wsMain, colsToDelete, filterColStep1, delValue1, delValue2)
    Call Step2_SplitDataToTabs(wsMain, splitColStep2)
    Call Step3_HighlightRandom10()
    Call Step4_FilterByColor()

Cleanup:
    ' Restore application features
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "Process Complete!", vbInformation
End Sub

' =======================================================================
' STEP 1: Remove rows 1-3, delete named columns, delete rows based on 2 values
' =======================================================================
Private Sub Step1_CleanData(ws As Worksheet, colsToDelete As Variant, filterCol As String, val1 As String, val2 As String)
    ' Cleans the source by removing fixed rows/columns and deleting rows that match either filter value.
    Dim lastRow As Long, lastCol As Long
    Dim c As Long, colIndex As Long

    ' Remove top 3 rows
    ws.Rows("1:3").Delete Shift:=xlUp
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Loop backwards to delete columns safely
    For c = lastCol To 1 Step -1
        If Not IsError(Application.Match(ws.Cells(1, c).Value, colsToDelete, 0)) Then
            ws.Columns(c).Delete
        End If
    Next c

    ' Recalculate last column after deletions
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find column index for row deletion
    colIndex = 0
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = filterCol Then
            colIndex = c
            Exit For
        End If
    Next c

    ' If column found, use AutoFilter to delete rows highly performantly
    If colIndex > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        If lastRow > 1 Then
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
                Field:=colIndex, Criteria1:=val1, Operator:=xlOr, Criteria2:=val2
                
            On Error Resume Next
            ' Delete visible rows (offset by 1 to protect headers)
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            On Error GoTo 0
            
            ws.AutoFilterMode = False
        End If
    End If
End Sub

' =======================================================================
' STEP 2: Create a new tab for each unique value and copy relevant rows
' =======================================================================
Private Sub Step2_SplitDataToTabs(ws As Worksheet, splitColName As String)
    ' Splits the cleaned dataset into separate worksheets based on unique values in the split column.
    Dim lastRow As Long, lastCol As Long
    Dim colIndex As Long, c As Long
    Dim dict As Object, v As Variant
    Dim arrData As Variant
    Dim newWs As Worksheet, rngData As Range

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find column index for splitting data
    colIndex = 0
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = splitColName Then
            colIndex = c
            Exit For
        End If
    Next c

    If colIndex = 0 Then Exit Sub ' Column not found

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    Set rngData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Extract unique values using a Dictionary (fastest method)
    Set dict = CreateObject("Scripting.Dictionary")
    arrData = ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)).Value

    If IsArray(arrData) Then
        For c = 1 To UBound(arrData, 1)
            If Not IsEmpty(arrData(c, 1)) Then dict(arrData(c, 1)) = 1
        Next c
    ElseIf Not IsEmpty(arrData) Then
        dict(arrData) = 1
    End If

    ' Filter and copy to new sheets
    For Each v In dict.Keys
        ' Check if sheet exists, delete if it does to avoid errors
        On Error Resume Next
        Set newWs = ThisWorkbook.Sheets(Left(CStr(v), 31))
        If Not newWs Is Nothing Then newWs.Delete
        On Error GoTo 0
        
        ' Add new sheet
        Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newWs.Name = Left(CStr(v), 31) ' Max length for Excel sheet names is 31 chars

        ' Filter master sheet and copy visible cells to new sheet
        rngData.AutoFilter Field:=colIndex, Criteria1:=CStr(v)
        rngData.SpecialCells(xlCellTypeVisible).Copy Destination:=newWs.Range("A1")
    Next v

    ws.AutoFilterMode = False
End Sub

' =======================================================================
' STEP 3: Choose 10 random rows per new tab and highlight them green
' =======================================================================
Private Sub Step3_HighlightRandom10()
    ' Highlights up to 10 unique random data rows (or all rows when fewer than 10) on each generated sheet.
    Dim ws As Worksheet
    Dim lastRow As Long, count As Long, randomRow As Long
    Dim dict As Object
    
    Randomize ' Seed the random number generator
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Samp Fixer" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If lastRow > 1 Then
                Set dict = CreateObject("Scripting.Dictionary")
                count = 0
                
                ' If sheet has 10 or fewer data rows, highlight them all
                If (lastRow - 1) <= 10 Then
                    ws.Range(ws.Rows(2), ws.Rows(lastRow)).Interior.Color = RGB(144, 238, 144) ' Light Green
                Else
                    ' Randomly pick 10 unique rows
                    Do While count < 10
                        randomRow = Int((lastRow - 2 + 1) * Rnd + 2) ' Random int between 2 and lastRow
                        If Not dict.Exists(randomRow) Then
                            dict.Add randomRow, 1
                            ws.Rows(randomRow).Interior.Color = RGB(144, 238, 144)
                            count = count + 1
                        End If
                    Loop
                End If
            End If
        End If
    Next ws
End Sub

' =======================================================================
' STEP 4: Filter data in new tabs to only show the highlighted rows
' =======================================================================
Private Sub Step4_FilterByColor()
    ' Applies a color filter on each generated sheet so only green-highlighted rows remain visible.
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Samp Fixer" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            If lastRow > 1 Then
                ' Apply color filter to Column 1 (affects entire visible range)
                ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
                    Field:=1, _
                    Criteria1:=RGB(144, 238, 144), _
                    Operator:=xlFilterCellColor
            End If
        End If
    Next ws
End Sub