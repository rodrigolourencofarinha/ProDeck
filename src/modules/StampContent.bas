Attribute VB_Name = "StampContent"
Function StampC_Function(stamp As String)
Dim sld As slide
Dim stmp As Shape
Dim ttlTop As Single
Dim ttlBottom As Single
Dim ttlLeft As Single
Dim ttlRight As Single
Dim ttlHeight As Single
Dim ttlWidth As Single

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


ttlRight = ttlLeft + ttlWidth
ttlBottom = ttlTop + ttlHeight

Set sld = ActiveWindow.Presentation.Slides(ActiveWindow.View.slide.SlideNumber)


Set stmp = sld.Shapes.AddTextbox(msoShapeRectangle, _
    Left:=ttlRight - cm2Points(4), Top:=ttlBottom + cm2Points(0.6), Width:=cm2Points(4), Height:=cm2Points(0.8))
stmp.Fill.ForeColor.RGB = RGB(255, 255, 255)
stmp.line.Weight = 1
stmp.line.ForeColor.RGB = RGB(0, 0, 0)
stmp.TextFrame.HorizontalAnchor = msoAnchorCenter
stmp.TextFrame.VerticalAnchor = msoAnchorMiddle
stmp.TextFrame.AutoSize = ppAutoSizeNone
stmp.Height = cm2Points(0.7)
stmp.Width = cm2Points(4)
stmp.TextFrame2.MarginLeft = 3.714285714286
stmp.TextFrame2.MarginRight = 3.714285714286
stmp.TextFrame2.MarginTop = 3.714285714286
stmp.TextFrame2.MarginBottom = 3.714285714286
With stmp.TextFrame.TextRange
    .text = stamp
    .ParagraphFormat.Alignment = ppAlignCenter
    With .Font
        .size = 14
        .Name = "Arial"
        .color.RGB = RGB(0, 0, 0)
    End With
End With

End Function
Function cm2Points(inVal As Single) As Single
'convert cm to points
cm2Points = inVal * 28.346
End Function
Sub StampC_Draft()
StampC_Function ("DRAFT")
End Sub
Sub StampC_Preliminary()
StampC_Function ("PRELIMINARY")
End Sub
Sub StampC_Illustrative()
StampC_Function ("ILLUSTRATIVE")
End Sub
Sub StampC_Not_Exhaustive()
StampC_Function ("NOT EXHAUSTIVE")
End Sub
Sub StampC_Discussion()
StampC_Function ("FOR DISCUSSION")
End Sub



