Attribute VB_Name = "Footnotes"
Sub Notes()
On Error Resume Next
Dim sld As slide
Dim slide As slide
Dim shp As Shape
Dim note As Shape
Dim ttlTop As Single
Dim ttlBottom As Single
Dim ttlLeft As Single
Dim ttlRight As Single
Dim ttlHeight As Single
Dim ttlWidth As Single
Dim X As Long

sngName = "PRODECK SLIDE NOTES"

If Not ActiveWindow.Selection.SlideRange.Shapes.HasTitle Then
    ttlTop = cm2Points(0.6)
    ttlLeft = cm2Points(1)
    ttlWidth = cm2Points(32)
    ttlHeight = cm2Points(2.5)
    Else
    ttlTop = ActiveWindow.Selection.SlideRange.Shapes.title.Top
    ttlLeft = ActiveWindow.Selection.SlideRange.Shapes.title.Left
    ttlWidth = ActiveWindow.Selection.SlideRange.Shapes.title.Width
    ttlHeight = ActiveWindow.Selection.SlideRange.Shapes.title.Height
End If


Set sld = ActiveWindow.Selection.SlideRange(1)

ttlBottom = Application.ActivePresentation.PageSetup.SlideHeight

For Each shp In sld.Shapes
    If shp.Name = "PRODECK SLIDE NOTES" Then
        shp.Left = ttlLeft
        shp.Top = ttlBottom - cm2Points(2.5)
        Exit Sub
    ElseIf shp.Name = "PRODECK SLIDE NOTES - USER DEFINED POSITION" Then
        shp.Left = ttlLeft
        shp.Top = ttlBottom - cm2Points(2.5)
        Exit Sub
    Else
    End If
Next shp


Set note = sld.Shapes.AddTextbox(msoShapeRectangle, _
    Left:=ttlLeft, Top:=ttlBottom - cm2Points(2.5), Width:=ttlWidth, Height:=cm2Points(0.86))
note.Name = "PRODECK SLIDE NOTES"
note.TextFrame.VerticalAnchor = msoAnchorBottom
note.TextFrame.AutoSize = ppAutoSizeShapeToFitText
note.TextFrame.WordWrap = msoFalse
note.TextFrame2.MarginLeft = 0
note.TextFrame2.MarginRight = 0
note.TextFrame2.MarginTop = 0
note.TextFrame2.MarginBottom = 0
With note.TextFrame2.TextRange.ParagraphFormat
    .Bullet.Type = msoBulletNumbered
    .Bullet.Style = msoBulletArabicParenBoth
    ' Before
    .LeftIndent = 14.2
    ' Hanging
    .FirstLineIndent = -14.2
End With
With note.TextFrame.TextRange
    .text = "Notes:" & vbCrLf
    .ParagraphFormat.Alignment = ppAlignLeft
    With .Font
        .size = 10
        .color.RGB = RGB(0, 0, 0)
    End With
 note.TextFrame2.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Type = msoBulletNone

    note.Left = ttlLeft

    For Each sld In ActiveWindow.Presentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = "PRODECK SLIDE NOTES - USER DEFINED POSITION" Then
                note.Left = shp.Left
                note.Top = shp.Top
            End If
        Next shp
    Next sld

End With


note.Select

End Sub
Sub Update_Notes()
On Error Resume Next
Dim X As Long
Dim shp As Shape
Dim sld As slide
Dim Shpl As Shape
Dim sngLastY As Single
Dim sngLastX As Single
Dim L As Long
Dim B As Long

Set Shp1 = ActiveWindow.Selection.ShapeRange(1)
sngName = Shp1.Name

  
If sngName = "PRODECK SLIDE NOTES" Then
ElseIf sngName = "PRODECK SLIDE NOTES - USER DEFINED POSITION" Then
Else
Exit Sub
End If

For Each sld In ActiveWindow.Presentation.Slides
    For Each shp In sld.Shapes
        If shp.Name = "PRODECK SLIDE NOTES - USER DEFINED POSITION" Then
            shp.Name = "PRODECK SLIDE NOTES"
        End If
    Next shp
Next sld

Shp1.Name = "PRODECK SLIDE NOTES - USER DEFINED POSITION"

L = Shp1.Left
B = Shp1.Top + Shp1.Height

For Each sld In ActiveWindow.Presentation.Slides
    For Each shp In sld.Shapes
        If shp.Name = "PRODECK SLIDE NOTES" Then
            shp.Left = L
            shp.Top = B - shp.Height
        End If
    Next shp
Next sld
End Sub

Function cm2Points(inVal As Single) As Single
'convert cm to points
cm2Points = inVal * 28.346
End Function






