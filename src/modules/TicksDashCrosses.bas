' Extracted from: TicksDashCrosses.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "TicksDashCrosses"
Sub TextBulletsTicks()
'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If Err Then
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------

If ActiveWindow.Selection.TextRange2.ParagraphFormat.Bullet.Character = 252 Then
    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.HangingPunctuation = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Bullet.Visible = False
    End With
Else
    With ActiveWindow.Selection.TextRange.ParagraphFormat.Bullet
        .Character = 252
        .Visible = True
        .Font.Name = "Wingdings"
        .Font.color = RGB(0, 128, 0)
    End With
End If
End Sub

Sub TextBulletsCrosses()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If Err Then
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------

If ActiveWindow.Selection.TextRange2.ParagraphFormat.Bullet.Character = 215 Then
    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.HangingPunctuation = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Bullet.Visible = False
    End With
Else
    With ActiveWindow.Selection.TextRange.ParagraphFormat.Bullet
        .Character = 215
        .Visible = True
        .Font.Name = "Arial"
        .Font.color = RGB(255, 0, 0)
    End With
End If
End Sub
Sub TextBulletsDashes()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If Err Then
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------

If ActiveWindow.Selection.TextRange2.ParagraphFormat.Bullet.Character = 8722 Then
    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.HangingPunctuation = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Bullet.Visible = False
    End With
Else
    With ActiveWindow.Selection.TextRange.ParagraphFormat.Bullet
        .Character = 8722
        .Visible = True
        .Font.Name = "Arial"
        .Font.color = RGB(59, 154, 220)
    End With
End If
End Sub


