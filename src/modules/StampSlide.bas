' Extracted from: Stampslide.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "Stampslide"
Function StampS_Function(stamp As String, red As Long, green As Long, blue As Long)
Dim sld As slide
Dim stmp As Shape
Dim sldHeight As Single
Dim sldWidth As Single
Dim sldRight As Single

sldHeight = ActivePresentation.PageSetup.SlideHeight
sldWidth = ActivePresentation.PageSetup.SlideWidth
sldRight = sldWidth

Set sld = ActiveWindow.Presentation.Slides(ActiveWindow.View.slide.SlideNumber)

For Each sld In ActiveWindow.Selection.SlideRange
    
    For Each shp In sld.Shapes
        If shp.Name = "PRODECK SLIDE LABEL" Then
            shp.Delete
        End If
    Next shp
    
    Set stmp = sld.Shapes.AddShape(msoShapeDiagonalStripe, _
        Left:=sldRight - sldHeight / 5, Top:=0, Width:=sldHeight / 5, Height:=sldHeight / 5)
    stmp.Flip msoFlipHorizontal
    stmp.Fill.ForeColor.RGB = RGB(red, green, blue)
    stmp.line.Visible = msoFalse
    stmp.TextFrame.HorizontalAnchor = msoAnchorCenter
    stmp.TextFrame.VerticalAnchor = msoAnchorMiddle
    stmp.TextFrame.AutoSize = ppAutoSizeNone
    stmp.TextFrame.WordWrap = msoFalse
    stmp.TextFrame2.MarginLeft = 0
    stmp.TextFrame2.MarginRight = 0
    stmp.TextFrame2.MarginTop = 0
    stmp.TextFrame2.MarginBottom = 0
    stmp.TextFrame2.ThreeD.RotationZ = 45
    With stmp.TextFrame.TextRange
        .text = stamp
        .ParagraphFormat.Alignment = ppAlignCenter
        With .Font
            .size = 14
            .Name = "Arial"
            .color.RGB = RGB(255, 255, 255)
        End With
    stmp.Name = "PRODECK SLIDE LABEL"
    End With

Next sld

End Function
Sub StampS_Delete()

Dim shp As Shape
Dim sld As slide

sngName = "PRODECK SLIDE LABEL"
    For Each sld In ActiveWindow.Presentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = sngName Then
                shp.Delete
            End If
        Next shp
    Next sld
End Sub
Sub StampS_Front()
Dim X As Long
Dim shp As Shape
Dim sld As slide

For Each sld In ActiveWindow.Presentation.Slides
    For X = 1 To ActiveWindow.Presentation.Slides.Count
        For Each shp In ActiveWindow.Presentation.Slides(X).Shapes
            If shp.Name = "PRODECK SLIDE LABEL" Then
                shp.ZOrder msoBringToFront
            End If
        Next
    Next
Next

End Sub
Function cm2Points(inVal As Single) As Single
'convert cm to points
cm2Points = inVal * 28.346
End Function
Sub StampS_New()
StampS_Function "NEW", 0, 176, 80
End Sub
Sub StampS_Updated()
StampS_Function "UPDATED", 46, 117, 182
End Sub
Sub StampS_Draft()
StampS_Function "DRAFT", 191, 144, 0
End Sub
Sub StampS_Preliminary()
StampS_Function "PRELIMINARY", 255, 153, 0
End Sub
Sub StampS_Appendix()
StampS_Function "TO APPENDIX", 255, 51, 153
End Sub
Sub StampS_Remove()
StampS_Function "TO BE REMOVED", 255, 0, 0
End Sub









