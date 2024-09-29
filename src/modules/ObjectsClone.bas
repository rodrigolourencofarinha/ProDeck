Attribute VB_Name = "ObjectsClone"
Sub ObjectsCloneRight()
      
Set myDocument = Application.ActiveWindow

Dim OldTop, OldLeft As Double

If Not myDocument.Selection.Type = ppSelectionShapes Then
    MsgBox "No shapes selected.", vbCritical
    
Else

If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        
        OldTop = ActiveWindow.Selection.ShapeRange.Top
        OldLeft = ActiveWindow.Selection.ShapeRange.Left
        
        Set SlideShape = ActiveWindow.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = OldTop
            .Left = OldLeft + SlideShape.Width
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = ActiveWindow.Selection.ShapeRange.Group
        
        OldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.Left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = OldTop
            .Left = OldLeft + SlideShape.Width
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
    
End If
    
End Sub

Sub ObjectsCloneDown()

Set myDocument = Application.ActiveWindow
Dim OldTop, OldLeft As Double

If Not myDocument.Selection.Type = ppSelectionShapes Then
  MsgBox "No shapes selected.", vbCritical

Else

If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        
        OldTop = ActiveWindow.Selection.ShapeRange.Top
        OldLeft = ActiveWindow.Selection.ShapeRange.Left
        
        Set SlideShape = ActiveWindow.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = OldTop + SlideShape.Height
            .Left = OldLeft
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = ActiveWindow.Selection.ShapeRange.Group
        
        OldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.Left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = OldTop + SlideShape.Height
            .Left = OldLeft
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
End If
    
End Sub
