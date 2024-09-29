Attribute VB_Name = "ConnectShapes"
Sub ConnectRectangleShapes(ShapeDirection As String)

Dim Left1, Right1, Top1, Bottom1, Left2, Right2, Top2, Bottom2 As Double
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        
        If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
        
        Left1 = ActiveWindow.Selection.ChildShapeRange(1).Left
        Right1 = Left1 + ActiveWindow.Selection.ChildShapeRange(1).Width
        Top1 = ActiveWindow.Selection.ChildShapeRange(1).Top
        Bottom1 = Top1 + ActiveWindow.Selection.ChildShapeRange(1).Height
        
        Left2 = ActiveWindow.Selection.ChildShapeRange(2).Left
        Right2 = Left2 + ActiveWindow.Selection.ChildShapeRange(2).Width
        Top2 = ActiveWindow.Selection.ChildShapeRange(2).Top
        Bottom2 = Top2 + ActiveWindow.Selection.ChildShapeRange(2).Height
        
        
        Select Case ShapeDirection
        
        Case "RightToLeft"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right1, Y1:=Top1)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
        Case "LeftToRight"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right2, Y1:=Top2)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
         Case "BottomToTop"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left1, Y1:=Bottom1)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
         Case "TopToBottom"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left2, Y1:=Bottom2)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
            
        End Select
        
        Else
        MsgBox "Select two shapes.", vbCritical
        End If
        

    Else
    
        If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).Left
        Right1 = Left1 + ActiveWindow.Selection.ShapeRange(1).Width
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Bottom1 = Top1 + ActiveWindow.Selection.ShapeRange(1).Height
        
        Left2 = ActiveWindow.Selection.ShapeRange(2).Left
        Right2 = Left2 + ActiveWindow.Selection.ShapeRange(2).Width
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        Bottom2 = Top2 + ActiveWindow.Selection.ShapeRange(2).Height
        
        
        Select Case ShapeDirection
        
        Case "RightToLeft"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right1, Y1:=Top1)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
        Case "LeftToRight"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right2, Y1:=Top2)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
         Case "BottomToTop"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left1, Y1:=Bottom1)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
         Case "TopToBottom"
            With ActiveWindow.Selection.SlideRange.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left2, Y1:=Bottom2)
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
                .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
                '.ConvertToShape
                .ConvertToShape.line.Visible = msoFalse
            End With
            
            
        End Select
        
        Else
        MsgBox "Select two shapes.", vbCritical
        End If
        
    End If
    End If
    
End Sub
