' Extracted from: StackObjects.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "StackObjects"

Sub Stack_Top()

Dim Shp1 As Shape
Dim shp As Shape
Dim X As Long
Dim M As Long
Dim sngLastY As Single
Dim rayShapes() As Shape

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
'START: Error message message box -----------------------------------------------

    ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ChildShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ChildShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByTop(rayShapes)
    
    Set Shp1 = rayShapes(1)
    sngLastY = Shp1.Top + Shp1.Height
    For X = 2 To UBound(rayShapes)
        Set shp = rayShapes(X)
        With shp
            .Top = sngLastY
            sngLastY = .Top + .Height
        End With
    Next X

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
'START: Error message message box -----------------------------------------------

    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByTop(rayShapes)
    
    Set Shp1 = rayShapes(1)
    sngLastY = Shp1.Top + Shp1.Height
    For X = 2 To UBound(rayShapes)
        Set shp = rayShapes(X)
        With shp
            .Top = sngLastY
            sngLastY = .Top + .Height
        End With
    Next X

End If

End Sub
Sub Stack_Left()

Dim Shp1 As Shape
Dim shp As Shape
Dim X As Long
Dim M As Long
Dim sngLastY As Single
Dim rayShapes() As Shape

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
'START: Error message message box -----------------------------------------------

    ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ChildShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ChildShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByLeft(rayShapes)
    
    Set Shp1 = rayShapes(1)
    sngLastY = Shp1.Left + Shp1.Width
    For X = 2 To UBound(rayShapes)
        Set shp = rayShapes(X)
        With shp
            .Left = sngLastY
            sngLastY = .Left + .Width
        End With
    Next X

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
'START: Error message message box -----------------------------------------------

    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByLeft(rayShapes)
    
    Set Shp1 = rayShapes(1)
    sngLastY = Shp1.Left + Shp1.Width
    For X = 2 To UBound(rayShapes)
        Set shp = rayShapes(X)
        With shp
            .Left = sngLastY
            sngLastY = .Left + .Width
        End With
    Next X

End If

End Sub
Sub Stack_Bottom()
Dim Shp1 As Shape
Dim shp As Shape
Dim X As Long
Dim M As Long
Dim B As Long
Dim sngLastY As Single
Dim rayShapes() As Shape

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
'START: Error message message box -----------------------------------------------


    ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ChildShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ChildShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByTop(rayShapes)
    
    B = UBound(rayShapes)
    
    Set Shp1 = rayShapes(B)
    sngLastY = Shp1.Top
    For X = (UBound(rayShapes) - 1) To 1 Step -1
        Set shp = rayShapes(X)
        With shp
            .Top = sngLastY + .Height
            sngLastY = .Top
        End With
    Next X


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
'START: Error message message box -----------------------------------------------


    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByTop(rayShapes)
    
    B = UBound(rayShapes)
    
    Set Shp1 = rayShapes(B)
    sngLastY = Shp1.Top
    For X = (UBound(rayShapes) - 1) To 1 Step -1
        Set shp = rayShapes(X)
        With shp
            .Top = sngLastY - .Height
            sngLastY = .Top
        End With
    Next X
    
End If
End Sub
Sub Stack_Right()
Dim Shp1 As Shape
Dim shp As Shape
Dim X As Long
Dim M As Long
Dim B As Long
Dim sngLastY As Single
Dim rayShapes() As Shape

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
'START: Error message message box -----------------------------------------------


    ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ChildShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ChildShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByLeft(rayShapes)
    
    B = UBound(rayShapes)
    
    Set Shp1 = rayShapes(B)
    sngLastY = Shp1.Left
    For X = (UBound(rayShapes) - 1) To 1 Step -1
        Set shp = rayShapes(X)
        With shp
            .Left = sngLastY - .Width
            sngLastY = .Left
        End With
    Next X

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
'START: Error message message box -----------------------------------------------


    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    For M = 1 To ActiveWindow.Selection.ShapeRange.Count
    Set rayShapes(M) = ActiveWindow.Selection.ShapeRange(M)
    Next
    ' make sure selected shapes are sorted by top value
    Call SortByLeft(rayShapes)
    
    B = UBound(rayShapes)
    
    Set Shp1 = rayShapes(B)
    sngLastY = Shp1.Left
    For X = (UBound(rayShapes) - 1) To 1 Step -1
        Set shp = rayShapes(X)
        With shp
            .Left = sngLastY - .Width
            sngLastY = .Left
        End With
    Next X
    
End If
End Sub
