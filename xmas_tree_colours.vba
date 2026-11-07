Sub XmasTreeColours()
    Dim colours(0 To 4) As Long
    Dim cell As Range

    colours(0) = RGB(0, 100, 0)     ' tree green
    colours(1) = RGB(34, 139, 34)   ' forest green
    colours(2) = RGB(139, 69, 19)   ' trunk brown
    colours(3) = RGB(220, 20, 60)   ' bauble red
    colours(4) = RGB(255, 215, 0)   ' star gold

    Randomize

    For Each cell In ActiveSheet.Range("A1:C1").Cells
        cell.Interior.Color = colours(Int(Rnd * 5))
    Next cell
End Sub
