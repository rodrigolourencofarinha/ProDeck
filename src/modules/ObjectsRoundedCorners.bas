Attribute VB_Name = "ObjectsRoundedCorners"
Sub ObjectsCopyRoundedCorner()
Dim SlideShape  As PowerPoint.Shape

        
If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
    MsgBox "No shapes selected."
Else

On Error Resume Next
Dim ShapeRadius As Single
If Application.ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 0 Then

ShapeRadius = ActiveWindow.Selection.ShapeRange(1).Adjustments(1) / (1 / (ActiveWindow.Selection.ShapeRange(1).Height + ActiveWindow.Selection.ShapeRange(1).Width))

If ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 1 Then
    ShapeRadius2 = ActiveWindow.Selection.ShapeRange(1).Adjustments(2) / (1 / (ActiveWindow.Selection.ShapeRange(1).Height + ActiveWindow.Selection.ShapeRange(1).Width))
End If

For Each SlideShape In ActiveWindow.Selection.ShapeRange
    With SlideShape

        .AutoShapeType = ActiveWindow.Selection.ShapeRange(1).AutoShapeType
        .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
        If ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 1 Then
            .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
        End If
    End With
Next

End If

End If

End Sub
Sub ObjectsCopyShapeTypeAndAdjustments()

Dim SlideShape  As PowerPoint.Shape

If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
    MsgBox "No shapes selected."
Else

Dim AdjustmentsCount As Long
Dim ShapeCount  As Long

For ShapeCount = 2 To ActiveWindow.Selection.ShapeRange.Count
    
    ActiveWindow.Selection.ShapeRange(ShapeCount).AutoShapeType = ActiveWindow.Selection.ShapeRange(1).AutoShapeType
    
    For AdjustmentsCount = 1 To ActiveWindow.Selection.ShapeRange(1).Adjustments.Count
        
        ActiveWindow.Selection.ShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = ActiveWindow.Selection.ShapeRange(1).Adjustments(AdjustmentsCount)
        
    Next AdjustmentsCount
    
Next ShapeCount

End If

End Sub

