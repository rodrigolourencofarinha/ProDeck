' Extracted from: RectifyLines.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "RectifyLines"
Sub RectifyLines()

Dim LineShape   As Shape

If ActiveWindow.Selection.Type = ppSelectionShapes Then

    If ActiveWindow.Selection.HasChildShapeRange Then
    
        For Each LineShape In ActiveWindow.Selection.ChildShapeRange
            
            With LineShape
                
                If .Fill.Type = -2 And .AutoShapeType = -2 Then
                    
                    If .Width > .Height Then
                        .Height = 0
                    Else
                        .Width = 0
                    End If
                End If
            End With
            
        Next
        
    Else
    
        For Each LineShape In ActiveWindow.Selection.ShapeRange
        
            With LineShape
                
                If .Fill.Type = -2 And .AutoShapeType = -2 Then
                    
                    If .Width > .Height Then
                        .Height = 0
                    Else
                        .Width = 0
                    End If
                End If
            End With
            
        Next
    
    End If
        
Else

    MsgBox "No shape selected."
    
End If

        
End Sub
