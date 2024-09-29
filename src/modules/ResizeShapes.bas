Attribute VB_Name = "ResizeShapes"
Sub Resize_Shapes()

Dim sngNewWidth As Single
Dim sngNewHeight As Single
Dim oSh As Shape

If ActiveWindow.Selection.HasChildShapeRange Then

    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box --------------------------------------------------


    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ChildShapeRange
        sngNewWidth = .Item(1).Width
        sngNewHeight = .Item(1).Height
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ChildShapeRange
        oSh.Width = sngNewWidth
        oSh.Height = sngNewHeight
    Next

Else

    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box --------------------------------------------------


    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ShapeRange
        sngNewWidth = .Item(1).Width
        sngNewHeight = .Item(1).Height
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ShapeRange
        oSh.Width = sngNewWidth
        oSh.Height = sngNewHeight
    Next

End If


End Sub
Sub Resize_Height()

Dim sngNewHeight As Single
Dim oSh As Shape

If ActiveWindow.Selection.HasChildShapeRange Then
    
    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box --------------------------------------------------
   
    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ChildShapeRange
        sngNewHeight = .Item(1).Height
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ChildShapeRange
        oSh.Height = sngNewHeight
    Next

Else


    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box --------------------------------------------------
   
    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ShapeRange
        sngNewHeight = .Item(1).Height
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ShapeRange
        oSh.Height = sngNewHeight
    Next

End If


End Sub
Sub Resize_Width()

Dim sngNewWidth As Single
Dim oSh As Shape

If ActiveWindow.Selection.HasChildShapeRange Then

    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box -----------------------------------------------
   
    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ChildShapeRange
        sngNewWidth = .Item(1).Width
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ChildShapeRange
        oSh.Width = sngNewWidth
    Next

Else

    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ShapeRange.Count < 2 Then
          MsgBox "You must have at least two objects selected.", vbCritical
            Exit Sub
        End If
        If Err <> 0 Then
            Exit Sub
        End If
    'END: Error message message box --------------------------------------------------
   
    ' Start with the height/width of first shape in selection
    With ActiveWindow.Selection.ShapeRange
        sngNewWidth = .Item(1).Width
    End With
    ' now that we know the height/width of smallest shape
    For Each oSh In ActiveWindow.Selection.ShapeRange
        oSh.Width = sngNewWidth
    Next


End If

End Sub
