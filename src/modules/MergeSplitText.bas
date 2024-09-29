Attribute VB_Name = "MergeSplitText"
Sub Merge_Text_Boxes()
Dim oFirstShape As Shape
Dim ShapeOne As Shape
Dim oSh As Shape
Dim X As Long
Dim L As Long
Dim rayShapes() As Shape

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count < 2 Then
      MsgBox "You must have at least two shapes selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)

For L = 1 To ActiveWindow.Selection.ShapeRange.Count
    Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L

' make sure selected shapes are sorted by top value
Call SortByTop(rayShapes)

Set ShapeOne = rayShapes(1)

If ShapeOne.Type = msoGroup Then

    MsgBox "Select an object and not a group.", vbCritical
    Exit Sub
    
Else
    
    ShapeOne.PickUp
    
    Set oFirstShape = ActiveWindow.Selection.SlideRange.Shapes.AddTextbox(msoTextOrientationHorizontal, _
    Left:=ShapeOne.Left, Top:=ShapeOne.Top, Width:=ShapeOne.Width, Height:=ShapeOne.Height)
    
    oFirstShape.Apply
    
    oFirstShape.TextFrame.TextRange.text = _
        ShapeOne.TextFrame.TextRange.text & vbCrLf
            
    For X = 2 To ActiveWindow.Selection.ShapeRange.Count
    
        If rayShapes(X).Type = msoGroup Then
            MsgBox "Select an object and not a group.", vbCritical
            Exit Sub
        Else
            oFirstShape.TextFrame.TextRange.text = _
                oFirstShape.TextFrame.TextRange.text _
                & rayShapes(X).TextFrame.TextRange.text
        End If
        
        If X < ActiveWindow.Selection.ShapeRange.Count Then
            oFirstShape.TextFrame.TextRange.text = _
                oFirstShape.TextFrame.TextRange.text _
                & vbCrLf
        End If
        
    Next
    
    For X = ActiveWindow.Selection.ShapeRange.Count To 2 Step -1
        rayShapes(X).Delete
    Next
    
    ShapeOne.Delete
    
    Set oFirstShape = Nothing
    
    Set oSh = Nothing

End If
    
End Sub
Sub Split_Text_Boxes()

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

Dim oShp As Shape
Dim oSld As slide
Dim L As Long
Dim otr As TextRange2
Dim X As Integer

If ActiveWindow.Selection.ShapeRange.Type = msoGroup Then
    MsgBox "Select an object and not a group.", vbCritical
Else
    Set oShp = ActiveWindow.Selection.ShapeRange(1)
    oShp.PickUp
    
    Set oSld = oShp.Parent
    
        If oShp.HasTextFrame Then
        
            If oShp.TextFrame2.HasText Then
            
                With oShp.TextFrame2.TextRange
                
                    For L = 1 To .Paragraphs.Count
                    
                    Set otr = .Paragraphs(L)
                    
                        With oSld.Shapes.AddTextbox(msoTextOrientationHorizontal, oShp.Left, oShp.Top + X, 400, 15)
                        .Apply
                        .TextFrame2.AutoSize = oShp.TextFrame2.AutoSize
                        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                        .TextFrame2.WordWrap = msoFalse
                        .TextFrame2.VerticalAnchor = oShp.TextFrame2.VerticalAnchor
                        .TextFrame2.MarginBottom = oShp.TextFrame2.MarginBottom
                        .TextFrame2.MarginTop = oShp.TextFrame2.MarginTop
                        .TextFrame2.MarginLeft = oShp.TextFrame2.MarginLeft
                        .TextFrame2.MarginRight = oShp.TextFrame2.MarginRight
                        
                            With .TextFrame2.TextRange
                            .text = Replace(otr.text, vbCr, "")
                            .Font.size = oShp.TextFrame2.TextRange.Font.size
                            .Font.Bold = oShp.TextFrame2.TextRange.Font.Bold
                            .Font.Italic = oShp.TextFrame2.TextRange.Font.Italic
                            .Font.Fill.ForeColor = oShp.TextFrame2.TextRange.Font.Fill.ForeColor
                            .ParagraphFormat.Alignment = msoAlignLeft
                            End With
                            
                        If .TextFrame2.HasText = msoFalse Then
                            .Delete
                        End If
                        
                        X = X + .Height
                        
                    End With
                    
                Next L
                
            End With
                
        End If
        
    End If
    
    oShp.Delete
    Exit Sub
    
End If

End Sub



