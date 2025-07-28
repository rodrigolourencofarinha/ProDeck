' Extracted from: SelectObjects.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "SelectObjects"
Sub Select_Similar_Fill_Color()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
      MsgBox "You must have one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

Dim myFillColor As Variant
Dim shp As Shape
Dim mySlide As Integer
mySlide = ActiveWindow.View.slide.SlideIndex

If ActiveWindow.Selection.HasChildShapeRange Then

    If ActiveWindow.Selection.ChildShapeRange.Fill.Visible = msoFalse Then
    
        MsgBox "Shape with no fill.", vbCritical
    
    Else

        myFillColor = ActiveWindow.Selection.ChildShapeRange.Fill.ForeColor.RGB
        
        For Each shp In ActiveWindow.Selection.ShapeRange.GroupItems
        
            If shp.Fill.Visible = msoFalse Then
                Resume Next
            ElseIf shp.Fill.ForeColor.RGB = myFillColor Then
                shp.Select Replace:=False
            End If
            
        Next
    
    End If

Else

    If ActiveWindow.Selection.ShapeRange.Fill.Visible = msoFalse Then
    
        MsgBox "Shape with no fill.", vbCritical
        
    Else
    
        myFillColor = ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB
        
        For Each shp In ActiveWindow.Presentation.Slides(mySlide).Shapes.Range
        
            If shp.Fill.Visible = msoFalse Then
                Resume Next
            ElseIf shp.Fill.ForeColor.RGB = myFillColor Then
                shp.Select Replace:=False
            End If
            
        Next
        
    End If
    
End If

End Sub

Sub Select_Similar_Line_Color()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
      MsgBox "You must have one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

Dim myLineColor As Variant
Dim shp As Shape
Dim mySlide As Integer
mySlide = ActiveWindow.View.slide.SlideIndex

If ActiveWindow.Selection.HasChildShapeRange Then
    If ActiveWindow.Selection.ChildShapeRange.Fill.Visible = msoFalse Then
        MsgBox "Shape with no fill.", vbCritical
    Else
        myLineColor = ActiveWindow.Selection.ChildShapeRange.line.ForeColor.RGB
        For Each shp In ActiveWindow.Selection.ShapeRange.GroupItems
            If shp.line.Visible = msoFalse Then
            Resume Next
            ElseIf shp.line.ForeColor.RGB = myLineColor Then
            shp.Select Replace:=False
            End If
        Next
    End If
    
Else

    If ActiveWindow.Selection.ShapeRange.line.Visible = msoFalse Then
        MsgBox "Shape with no outline.", vbCritical
        Exit Sub
    Else
        myLineColor = ActiveWindow.Selection.ShapeRange.line.ForeColor.RGB
        For Each shp In ActiveWindow.Presentation.Slides(mySlide).Shapes.Range
            If shp.line.Visible = msoFalse Then
            Resume Next
            ElseIf shp.line.ForeColor.RGB = myLineColor Then
            shp.Select Replace:=False
            End If
        Next
        
    End If
    
End If
End Sub
Sub Select_Similar_Font_Color()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
      MsgBox "You must have one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

Dim myFontColor As Variant
Dim shp As Shape
Dim mySlide As Integer
mySlide = ActiveWindow.View.slide.SlideIndex


If ActiveWindow.Selection.HasChildShapeRange Then

    If Not ActiveWindow.Selection.ChildShapeRange.HasTextFrame Then
        MsgBox "Object has no text frame", vbCritical
        Exit Sub
    Else
        myFontColor = ActiveWindow.Selection.ChildShapeRange.TextFrame.TextRange.Font.color.RGB
        For Each shp In ActiveWindow.Selection.ShapeRange.GroupItems
            If Not shp.HasTextFrame Then
                Resume Next
            ElseIf _
            shp.Type = mso3DModel Or _
            shp.Type = msoCallout Or _
            shp.Type = msoSmartArt Or _
            shp.Type = msoCanvas Or _
            shp.Type = msoComment Or _
            shp.Type = msoContentApp Or _
            shp.Type = msoDiagram Or _
            shp.Type = msoEmbeddedOLEObject Or _
            shp.Type = msoFormControl Or _
            shp.Type = msoGraphic Or _
            shp.Type = msoIgxGraphic Or _
            shp.Type = msoInk Or _
            shp.Type = msoInkComment Or _
            shp.Type = msoLine Or _
            shp.Type = msoLinked3DModel Or _
            shp.Type = msoLinkedGraphic Or _
            shp.Type = msoLinkedOLEObject Or _
            shp.Type = msoLinkedPicture Or _
            shp.Type = msoMedia Or _
            shp.Type = msoOLEControlObject Or _
            shp.Type = msoTable Or _
            shp.Type = msoScriptAnchor Or _
            shp.Type = msoSlicer Or _
            shp.Type = msoWebVideo Then
            Resume Next
                ElseIf shp.TextFrame.TextRange.Font.color.RGB = myFontColor Then
                shp.Select Replace:=False
            End If
        Next
    End If
Else
    If Not ActiveWindow.Selection.ShapeRange.HasTextFrame Then
        MsgBox "Object has no text frame.", vbCritical
        Exit Sub
    Else
        myFontColor = ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.color.RGB
        For Each shp In ActiveWindow.Presentation.Slides(mySlide).Shapes.Range
            If Not shp.HasTextFrame Then
                Resume Next
            ElseIf _
            shp.Type = mso3DModel Or _
            shp.Type = msoCallout Or _
            shp.Type = msoSmartArt Or _
            shp.Type = msoCanvas Or _
            shp.Type = msoComment Or _
            shp.Type = msoContentApp Or _
            shp.Type = msoDiagram Or _
            shp.Type = msoEmbeddedOLEObject Or _
            shp.Type = msoFormControl Or _
            shp.Type = msoGraphic Or _
            shp.Type = msoIgxGraphic Or _
            shp.Type = msoInk Or _
            shp.Type = msoInkComment Or _
            shp.Type = msoLine Or _
            shp.Type = msoLinked3DModel Or _
            shp.Type = msoLinkedGraphic Or _
            shp.Type = msoLinkedOLEObject Or _
            shp.Type = msoLinkedPicture Or _
            shp.Type = msoMedia Or _
            shp.Type = msoOLEControlObject Or _
            shp.Type = msoTable Or _
            shp.Type = msoScriptAnchor Or _
            shp.Type = msoSlicer Or _
            shp.Type = msoWebVideo Then
            Resume Next
                ElseIf shp.TextFrame.TextRange.Font.color.RGB = myFontColor Then
                shp.Select Replace:=False
            End If
        Next
    End If
End If
End Sub

