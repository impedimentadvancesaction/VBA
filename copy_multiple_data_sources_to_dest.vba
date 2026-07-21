Option Explicit

' Copies data from multiple source workbooks/sheets into target sheets
' in this workbook based on a simple mapping table.
Public Sub CopyDataFromMultipleWorkbooks()
	' Variable declarations for source/destination handling and copy sizing.
	Dim sourceMap As Variant
	Dim i As Long
	Dim srcWb As Workbook
	Dim srcWs As Worksheet
	Dim dstWs As Worksheet
	Dim srcPath As String
	Dim srcSheetName As String
	Dim dstSheetName As String
	Dim appendMode As Boolean
	Dim firstCell As Range
	Dim lastCell As Range
	Dim dataArr As Variant
	Dim rowsCount As Long
	Dim colsCount As Long
	Dim pasteRow As Long
    
	' Route unexpected errors to central cleanup.
	On Error GoTo CleanFail
    
	' Disable UI/calculation features for faster, cleaner execution.
	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.DisplayAlerts = False
	Application.Calculation = xlCalculationManual

	' Format per item:
	' Array("source file path", "source tab name", "destination tab name", appendMode)
	' appendMode = True  -> append below existing data in destination
	' appendMode = False -> clear destination and replace from A1
	sourceMap = Array( _
		Array("C:\Data\Source1.xlsx", "Input", "Target_1", False), _
		Array("C:\Data\Source2.xlsx", "RawData", "Target_2", False), _
		Array("C:\Data\Source3.xlsx", "Export", "Target_3", False) _
	)

	' Process each mapping entry one workbook/sheet pair at a time.
	For i = LBound(sourceMap) To UBound(sourceMap)
		' Read mapping values for this iteration.
		srcPath = CStr(sourceMap(i)(0))
		srcSheetName = CStr(sourceMap(i)(1))
		dstSheetName = CStr(sourceMap(i)(2))
		appendMode = CBool(sourceMap(i)(3))

		' Validate source file path before trying to open.
		If Len(Dir$(srcPath)) = 0 Then
			Err.Raise vbObjectError + 1000, , "Source file not found: " & srcPath
		End If

		' Open source workbook read-only and resolve source/destination sheets.
		Set srcWb = Workbooks.Open(Filename:=srcPath, UpdateLinks:=0, ReadOnly:=True)
		Set srcWs = srcWb.Worksheets(srcSheetName)
		Set dstWs = ThisWorkbook.Worksheets(dstSheetName)

		' Skip copy if source sheet is fully empty.
		Set firstCell = srcWs.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlNext)
		If Not firstCell Is Nothing Then
			' Detect used range bounds and load data into memory.
			Set lastCell = srcWs.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
			colsCount = srcWs.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

			dataArr = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastCell.Row, colsCount)).Value2
			rowsCount = UBound(dataArr, 1)
			colsCount = UBound(dataArr, 2)

			' Choose destination start row based on append/replace mode.
			If appendMode Then
				pasteRow = NextAvailableRow(dstWs)
			Else
				' Clear only values so existing formatting remains.
				dstWs.Cells.ClearContents
				pasteRow = 1
			End If

			' Write all values in a single bulk operation.
			dstWs.Cells(pasteRow, 1).Resize(rowsCount, colsCount).Value2 = dataArr
		End If

		' Close source workbook without saving changes.
		srcWb.Close SaveChanges:=False
		Set srcWb = Nothing
	Next i

CleanExit:
	' Always restore application settings before exit.
	On Error Resume Next
	If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False

	Application.Calculation = xlCalculationAutomatic
	Application.DisplayAlerts = True
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	Exit Sub

CleanFail:
	' Show error message, then run shared cleanup path.
	MsgBox "Copy failed: " & Err.Description, vbExclamation
	Resume CleanExit
End Sub

Private Function NextAvailableRow(ByVal ws As Worksheet) As Long
	' Returns next empty row based on last used row in the worksheet.
	Dim lastUsed As Range

	Set lastUsed = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	If lastUsed Is Nothing Then
		NextAvailableRow = 1
	Else
		NextAvailableRow = lastUsed.Row + 1
	End If
End Function