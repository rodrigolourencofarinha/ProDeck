' Extracted from: BulletPointsLEvel.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "BulletPointsLEvel"
Sub Bullet_Point_1()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If Err Then
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -------------------------------------------------

If ActiveWindow.Selection.TextRange2.ParagraphFormat.LeftIndent = 14.2 Then

    With ActiveWindow.Selection.TextRange2
            .ParagraphFormat.IndentLevel = 1
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Bullet.Visible = False
    End With
    
Else

    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.Bullet.Type = msoBulletUnnumbered
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.Bullet.Character = 8226
        ' Before
        .ParagraphFormat.LeftIndent = 14.2
        ' Hanging
        .ParagraphFormat.FirstLineIndent = -14.2
    End With
    
End If

End Sub
Sub Bullet_Point_2()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If Err Then
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -------------------------------------------------

If ActiveWindow.Selection.TextRange2.ParagraphFormat.LeftIndent = 28.45 Then

    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.HangingPunctuation = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Bullet.Visible = False
    End With
    
Else

    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 3
        .ParagraphFormat.Bullet.Type = msoBulletUnnumbered
        .ParagraphFormat.Bullet.Character = 8722
        ' Before
        .ParagraphFormat.LeftIndent = 28.45
        ' Hanging
        .ParagraphFormat.FirstLineIndent = -14.2
    End With
    
End If

End Sub
Sub Bullet_Point_3()

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

If ActiveWindow.Selection.TextRange2.ParagraphFormat.LeftIndent = 42.5 Then

    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 1
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.HangingPunctuation = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Bullet.Visible = False
    End With
    
Else

    With ActiveWindow.Selection.TextRange2
        .ParagraphFormat.IndentLevel = 3
        .ParagraphFormat.Bullet.Type = msoBulletUnnumbered
        .ParagraphFormat.Bullet.Character = 8227
        ' Before
        .ParagraphFormat.LeftIndent = 42.5
        ' Hanging
        .ParagraphFormat.FirstLineIndent = -14.2
    End With
    
End If

End Sub
