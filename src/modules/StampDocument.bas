' Extracted from: StampDocument.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "StampDocument"
Function StampD_Function(stamp As String)
On Error Resume Next
Dim sld As slide
Dim slide As slide
Dim shp As Shape
Dim stmp As Shape
Dim ttlTop As Single
Dim ttlBottom As Single
Dim ttlLeft As Single
Dim ttlRight As Single
Dim ttlHeight As Single
Dim ttlWidth As Single
Dim X As Long

sngName = "PRODECK DOCUMENT STAMP"
For Each sld In ActiveWindow.Presentation.Slides
    For Each shp In sld.Shapes
        If shp.Name = sngName Then
            shp.Delete
        End If
    Next shp
Next sld

On Error Resume Next
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


If Err.Number <> 0 Then
Else
    ttlTop = ActiveWindow.Selection.SlideRange.Shapes.title.Top
    ttlLeft = ActiveWindow.Selection.SlideRange.Shapes.title.Left
    ttlWidth = ActiveWindow.Selection.SlideRange.Shapes.title.Width
    ttlHeight = ActiveWindow.Selection.SlideRange.Shapes.title.Height
End If


ttlRight = ttlLeft + ttlWidth
ttlBottom = ttlTop + ttlHeight

For Each slide In ActiveWindow.Presentation.Slides
    Set stmp = slide.Shapes.AddTextbox(msoShapeRectangle, _
        Left:=ttlRight - cm2Points(6), Top:=0, Width:=cm2Points(6), Height:=cm2Points(0.7))
    stmp.Name = "PRODECK DOCUMENT STAMP"
    stmp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    stmp.TextFrame.VerticalAnchor = msoAnchorMiddle
    stmp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    stmp.TextFrame.WordWrap = msoFalse
    stmp.TextFrame2.MarginLeft = cm2Points(0.1)
    stmp.TextFrame2.MarginRight = cm2Points(0.1)
    stmp.TextFrame2.MarginTop = 0
    stmp.TextFrame2.MarginBottom = 0
    With stmp.TextFrame.TextRange
        .text = stamp
        .ParagraphFormat.Alignment = ppAlignRight
        With .Font
            .size = 16
            .Name = "Arial"
            .color.RGB = RGB(255, 0, 0)
        End With
    stmp.Left = ttlRight - stmp.Width
    End With
Next slide

For Each sld In ActiveWindow.Presentation.Slides
    For X = 1 To ActiveWindow.Presentation.Slides.Count
        For Each shp In ActiveWindow.Presentation.Slides(X).Shapes
            If shp.Name = "PRODECK SLIDE LABEL" Then
                shp.ZOrder msoBringToFront
            End If
        Next
    Next
Next


End Function
Sub StampD_Delete()
On Error Resume Next
Dim shp As Shape
Dim sld As slide

sngName = "PRODECK DOCUMENT STAMP"
    For Each sld In ActiveWindow.Presentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = sngName Then
                shp.Delete
            End If
        Next shp
    Next sld
End Sub
Sub StampD_Update()

    Dim X As Long
    Dim shp As Shape
    Dim sld As slide
    Dim Shpl As Shape
    Dim sngLastY As Single
    Dim sngLastX As Single
    Dim Shp1 As Shape
    
    Set Shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    Shp1.Copy
    sngName = Shp1.Name
    
    If sngName = "PRODECK DOCUMENT STAMP" Then
    Else
    Exit Sub
    End If
    
    
    sngLeft = Shp1.Left
    sngTop = Shp1.Top
    sngwidth = Shp1.Width
    sngheight = Shp1.Height
    sngForeColor = Shp1.Fill.ForeColor.RGB
    sngVerticalAchor = Shp1.TextFrame.VerticalAnchor
    sngAutoSize = Shp1.TextFrame.AutoSize
    sngWordWrap = Shp1.TextFrame.WordWrap
    sngMarginLeft = Shp1.TextFrame2.MarginLeft
    sngMarginRight = Shp1.TextFrame2.MarginRight
    sngMarginTop = Shp1.TextFrame2.MarginTop
    sngMarginBottom = Shp1.TextFrame2.MarginBottom
    sngText = Shp1.TextFrame.TextRange.text
    sngParagrahFormatAlignment = Shp1.TextFrame.TextRange.ParagraphFormat.Alignment
    sngFontSize = Shp1.TextFrame.TextRange.Font.size
    sngFontName = Shp1.TextFrame.TextRange.Font.Name
    sngFontColor = Shp1.TextFrame.TextRange.Font.color.RGB
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        For X = 1 To ActiveWindow.Presentation.Slides.Count
            For Each shp In ActiveWindow.Presentation.Slides(X).Shapes
                If shp.Name = "PRODECK DOCUMENT STAMP" Then
                    shp.Delete
                End If
            Next
                Set stmp = ActiveWindow.Presentation.Slides(X).Shapes.AddTextbox(msoShapeRectangle, _
                    Left:=sngLeft, Top:=sngTop, Width:=sngwidth, Height:=sngheight)
                stmp.Name = "PRODECK DOCUMENT STAMP"
                stmp.Fill.ForeColor.RGB = sngForeColor
                stmp.TextFrame.VerticalAnchor = sngVerticalAchor
                stmp.TextFrame.AutoSize = sngAutoSize
                stmp.TextFrame.WordWrap = sngWordWrap
                stmp.TextFrame2.MarginLeft = sngMarginLeft
                stmp.TextFrame2.MarginRight = sngMarginRight
                stmp.TextFrame2.MarginTop = sngMarginTop
                stmp.TextFrame2.MarginBottom = sngMarginBottom
                With stmp.TextFrame.TextRange
                    .text = sngText
                    .ParagraphFormat.Alignment = sngParagrahFormatAlignment
                    With .Font
                        .size = sngFontSize
                        .Name = sngFontName
                        .color.RGB = sngFontColor
                    End With
                stmp.Left = sngLeft
                End With
        Next
    Else
        MsgBox "No shape selected."
    End If
    

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
Sub StampD_Draft()
StampD_Function ("DRAFT MATERIAL")
End Sub
Sub StampD_Preliminary()
StampD_Function ("PRELIMINARY MATERIAL")
End Sub
Sub StampD_Confidential()
StampD_Function ("CONFIDENTIAL MATERIAL")
End Sub
Sub StampD_Distibute()
StampD_Function ("DO NOT DISTRIBUTE")
End Sub
Sub StampD_Internal()
StampD_Function ("FOR INTERNAL USE ONLY")
End Sub









