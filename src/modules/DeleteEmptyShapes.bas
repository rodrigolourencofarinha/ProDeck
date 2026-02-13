Attribute VB_Name = "DeleteEmptyShapes"
Option Explicit
Sub Delete_Empty_Shapes()

Dim oSl As slide
Dim oSh As Shape
Dim oShg As Shape
Dim i As Integer

On Error Resume Next

For Each oSl In ActiveWindow.Presentation.Slides

    For i = oSl.Shapes.Count To 1 Step -1
    
        Set oSh = oSl.Shapes(i)
        
            If oSh.Type = msoTextBox Or oSh.Type = msoAutoShape Then
            
                If oSh.Fill.Visible = msoFalse And oSh.line.Visible = msoFalse And (Not oSh.TextFrame.HasText) Then
                oSh.Delete
                End If
            
            End If
            
        Set oSh = Nothing
        
    Next i
    
Next oSl

End Sub
