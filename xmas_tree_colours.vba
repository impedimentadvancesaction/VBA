' Flashes A1:C1 with random Christmas tree colours every 0.5 seconds.
' Run StartXmasLights to begin and StopXmasLights to end.

Dim nextRun As Date          ' Time of the next scheduled update
Dim running As Boolean       ' True while the lights are flashing
Const INTERVAL As Double = 0.5 / 86400   ' 0.5 seconds expressed as a fraction of a day

' Starts the flashing effect
Sub StartXmasLights()
    running = True
    Randomize
    UpdateColours
End Sub

' Stops the flashing effect and cancels the pending update
Sub StopXmasLights()
    running = False
    On Error Resume Next     ' Ignore the error if nothing is scheduled
    Application.OnTime nextRun, "UpdateColours", , False
    On Error GoTo 0
End Sub

' Gives each of A1:C1 a random tree colour, then schedules itself again
Sub UpdateColours()
    Dim colours(0 To 4) As Long
    Dim cell As Range

    If Not running Then Exit Sub

    colours(0) = RGB(0, 100, 0)     ' tree green
    colours(1) = RGB(34, 139, 34)   ' forest green
    colours(2) = RGB(139, 69, 19)   ' trunk brown
    colours(3) = RGB(220, 20, 60)   ' bauble red
    colours(4) = RGB(255, 215, 0)   ' star gold

    ' Pick a random colour for each cell independently
    For Each cell In ActiveSheet.Range("A1:C1").Cells
        cell.Interior.Color = colours(Int(Rnd * 5))
    Next cell

    ' Schedule the next update 0.5 seconds from now
    nextRun = Now + INTERVAL
    Application.OnTime nextRun, "UpdateColours"
End Sub
