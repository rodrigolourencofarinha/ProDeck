Attribute VB_Name = "Headings"
Function Heading(headingNumber As Double)

On Error Resume Next
Dim sld As slide
Dim slide As slide
Dim shp As Shape
Dim head As Shape
Dim ttlTop As Double
Dim ttlBottom As Double
Dim ttlLeft As Double
Dim ttlRight As Double
Dim ttlHeight As Double
Dim ttlWidth As Double
Dim size As Double
Dim X As Long
Dim line As Shape
Dim n As Long
Dim Y As Long
Dim sPosition As Double
Dim Top As Double


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

n = headingNumber

size = (ttlWidth + cm2Points(1) * (1 - n)) / n

Top = ttlTop + ttlHeight + cm2Points(1)
 
For Y = 1 To n
    
    sPosition = ttlLeft + (size + cm2Points(1)) * (Y - 1)
    
    Set head = sld.Shapes.AddTextbox(msoShapeRectangle, _
        Left:=sPosition, Top:=Top, Width:=ttlWidth, Height:=27.53559)
    head.TextFrame.VerticalAnchor = msoAnchorBottom
    head.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    head.TextFrame.WordWrap = msoFalse
    head.TextFrame2.MarginLeft = 0
    head.TextFrame2.MarginRight = 0
    head.TextFrame2.MarginTop = 0
    head.TextFrame2.MarginBottom = 0
    With head.TextFrame.TextRange
        .text = "Heading"
        .ParagraphFormat.Alignment = ppAlignLeft
        With .Font
            .Bold = msoTrue
            .size = 20
            .color.RGB = RGB(0, 0, 0)
        End With
    End With
    
    Set line = sld.Shapes.AddLine(sPosition, Top + cm2Points(1), sPosition + size, Top + cm2Points(1))
    
    line.line.ForeColor.RGB = RGB(0, 0, 0)
    line.line.DashStyle = msoLineSolid
    line.line.Weight = 1
    
Next Y

End Function
Function cm2Points(inVal As Single) As Single
'convert cm to points
cm2Points = inVal * 28.346
End Function
Sub HeadingOne()
Heading (1)
End Sub
Sub HeadingTwo()
Heading (2)
End Sub
Sub HeadingThree()
Heading (3)
End Sub
Sub HeadingFour()
Heading (4)
End Sub
Sub HeadingFive()
Heading (5)
End Sub
Sub HeadingSix()
Heading (6)
End Sub

