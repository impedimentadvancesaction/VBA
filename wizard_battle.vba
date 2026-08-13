Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Const WIZ1_NAME As String = "Wizard1"
Const WIZ2_NAME As String = "Wizard2"
Const MAX_HP As Long = 100

Sub RunMagicBattle()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim shp1 As Shape, shp2 As Shape
    Set shp1 = ws.Shapes(WIZ1_NAME)
    Set shp2 = ws.Shapes(WIZ2_NAME)

    Randomize
    Dim hp1 As Long, hp2 As Long
    hp1 = MAX_HP: hp2 = MAX_HP

    CreateHealthBar ws, shp1, "HP1_Back", "HP1_Fill"
    CreateHealthBar ws, shp2, "HP2_Back", "HP2_Fill"

    Dim turn As Integer: turn = 1
    Dim dmg As Long

    Do While hp1 > 0 And hp2 > 0
        Dim attacker As Shape, defender As Shape
        If turn = 1 Then
            Set attacker = shp1: Set defender = shp2
        Else
            Set attacker = shp2: Set defender = shp1
        End If

        dmg = Int(Rnd * 16) + 10 '10-25

        CastSpell ws, attacker, defender, IIf(turn = 1, RGB(80, 160, 255), RGB(255, 90, 90))
        ShakeShape defender, 6, 12
        ShowFloatingDamage ws, defender, dmg

        If turn = 1 Then
            hp2 = hp2 - dmg: If hp2 < 0 Then hp2 = 0
            UpdateHealthBar ws, "HP2_Fill", hp2
        Else
            hp1 = hp1 - dmg: If hp1 < 0 Then hp1 = 0
            UpdateHealthBar ws, "HP1_Fill", hp1
        End If

        Sleep 400
        turn = 3 - turn
    Loop

    ShowVictoryBanner ws, IIf(hp1 > 0, shp1.Name, shp2.Name)
End Sub

Sub CreateHealthBar(ws As Worksheet, overShp As Shape, backName As String, fillName As String)
    Dim barW As Single, barH As Single, barTop As Single, barLeft As Single
    barW = overShp.Width: barH = 12
    barTop = overShp.Top - 20: barLeft = overShp.Left

    On Error Resume Next
    ws.Shapes(backName).Delete
    ws.Shapes(fillName).Delete
    On Error GoTo 0

    With ws.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barW, barH)
        .Name = backName
        .Fill.ForeColor.RGB = RGB(60, 60, 60)
        .Line.Visible = msoFalse
    End With
    With ws.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barW, barH)
        .Name = fillName
        .Fill.ForeColor.RGB = RGB(50, 220, 90)
        .Line.Visible = msoFalse
    End With
End Sub

Sub UpdateHealthBar(ws As Worksheet, fillName As String, hp As Long)
    Dim shp As Shape: Set shp = ws.Shapes(fillName)
    Dim fullW As Single: fullW = ws.Shapes(Replace(fillName, "Fill", "Back")).Width
    shp.Width = fullW * (hp / MAX_HP)
    If hp < MAX_HP * 0.3 Then
        shp.Fill.ForeColor.RGB = RGB(220, 50, 50)
    ElseIf hp < MAX_HP * 0.6 Then
        shp.Fill.ForeColor.RGB = RGB(230, 200, 40)
    End If
End Sub

Sub CastSpell(ws As Worksheet, fromShp As Shape, toShp As Shape, spellColor As Long)
    Dim orb As Shape
    Dim startX As Single, startY As Single, endX As Single, endY As Single
    startX = fromShp.Left + fromShp.Width / 2
    startY = fromShp.Top + fromShp.Height / 2
    endX = toShp.Left + toShp.Width / 2
    endY = toShp.Top + toShp.Height / 2

    Set orb = ws.Shapes.AddShape(msoShapeOval, startX, startY, 18, 18)
    orb.Fill.ForeColor.RGB = spellColor
    orb.Line.Visible = msoFalse
    orb.Shadow.Visible = True

    Dim steps As Integer: steps = 20
    Dim i As Integer
    For i = 1 To steps
        orb.Left = startX + (endX - startX) * (i / steps) - orb.Width / 2
        orb.Top = startY + (endY - startY) * (i / steps) - orb.Height / 2
        orb.Width = 14 + 6 * Sin(i * 0.8)
        orb.Height = orb.Width
        DoEvents
        Sleep 15
    Next i

    orb.Delete
    FlashImpact ws, toShp, spellColor
End Sub

Sub FlashImpact(ws As Worksheet, targetShp As Shape, flashColor As Long)
    Dim ov As Shape
    Set ov = ws.Shapes.AddShape(msoShapeOval, targetShp.Left, targetShp.Top, targetShp.Width, targetShp.Height)
    ov.Fill.ForeColor.RGB = flashColor
    ov.Fill.Transparency = 0.4
    ov.Line.Visible = msoFalse
    DoEvents
    Sleep 100
    ov.Delete
End Sub

Sub ShakeShape(shp As Shape, intensity As Integer, iterations As Integer)
    Dim origLeft As Single, origTop As Single
    origLeft = shp.Left: origTop = shp.Top
    Dim i As Integer
    For i = 1 To iterations
        shp.Left = origLeft + Rnd * intensity - intensity / 2
        shp.Top = origTop + Rnd * intensity - intensity / 2
        DoEvents
        Sleep 15
    Next i
    shp.Left = origLeft: shp.Top = origTop
End Sub

Sub ShowFloatingDamage(ws As Worksheet, targetShp As Shape, dmg As Long)
    Dim txt As Shape
    Set txt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        targetShp.Left + targetShp.Width / 2 - 15, targetShp.Top, 40, 20)
    With txt.TextFrame2.TextRange
        .Text = "-" & dmg
        .Font.Size = 16
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 60, 60)
    End With
    txt.Fill.Visible = msoFalse
    txt.Line.Visible = msoFalse

    Dim i As Integer
    For i = 1 To 15
        txt.Top = txt.Top - 2
        DoEvents
        Sleep 25
    Next i
    txt.Delete
End Sub

Sub ShowVictoryBanner(ws As Worksheet, winnerName As String)
    Dim banner As Shape
    Set banner = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        ws.Range("A1").Left, ws.Range("A1").Top + 100, 300, 50)
    With banner.TextFrame2.TextRange
        .Text = winnerName & " WINS!"
        .Font.Size = 28
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 215, 0)
    End With
    banner.Fill.Visible = msoFalse
    banner.Line.Visible = msoFalse
End Sub
