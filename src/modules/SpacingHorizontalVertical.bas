Attribute VB_Name = "SpacingHorizontalVertical"
Sub Increase_Horizontal_Spacing()
Dim sngGap As Single
Dim rayShapes() As Shape
Dim L As Double
Dim M As Double
Dim A As Double

If ActiveWindow.Selection.HasChildShapeRange Then

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
sngGap = cm2Points(0.1)  ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ChildShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ChildShapeRange(L)
Next L
' make sure selected shapes are sorted by Left value
Call SortByLeft(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left + sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left + sngGap * A
A = A + 1
Next M

End If
Else

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L
' make sure selected shapes are sorted by Left value
Call SortByLeft(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left + sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left + sngGap * A
A = A + 1
Next M

End If

End If

End Sub
Sub Decrease_Horizontal_Spacing()
Dim sngGap As Single
Dim rayShapes() As Shape
Dim L As Double
Dim M As Double
Dim A As Double

If ActiveWindow.Selection.HasChildShapeRange Then

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ChildShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ChildShapeRange(L)
Next L
' make sure selected shapes are sorted by Left value
Call SortByLeft(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left - sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left - sngGap * A
A = A + 1
Next M

End If

Else

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L
' make sure selected shapes are sorted by Left value
Call SortByLeft(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left - sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Left = rayShapes(L).Left + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Left = rayShapes(M).Left - sngGap * A
A = A + 1
Next M

End If

End If

End Sub
Sub Increase_Vertical_Spacing()
Dim sngGap As Single
Dim rayShapes() As Shape
Dim L As Double
Dim M As Double
Dim A As Double

If ActiveWindow.Selection.HasChildShapeRange Then

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------


ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ChildShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ChildShapeRange(L)
Next L
' make sure selected shapes are sorted by top value
Call SortByTop(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top + sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top + sngGap * A
A = A + 1
Next M

End If

Else

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------


ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L
' make sure selected shapes are sorted by top value
Call SortByTop(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top + sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top - sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top + sngGap * A
A = A + 1
Next M

End If

End If



End Sub
Sub Decrease_Vertical_Spacing()
Dim sngGap As Single
Dim rayShapes() As Shape
Dim L As Double
Dim M As Double
Dim A As Double

If ActiveWindow.Selection.HasChildShapeRange Then


'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ChildShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------


ReDim rayShapes(1 To ActiveWindow.Selection.ChildShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ChildShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ChildShapeRange(L)
Next L
' make sure selected shapes are sorted by top value
Call SortByTop(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top - sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top - sngGap * A
A = A + 1
Next M

End If

Else

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
       MsgBox "You must have at least two objects.", vbCritical
         Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------


ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
sngGap = cm2Points(0.1) ' Original: 2 cm gap
For L = 1 To ActiveWindow.Selection.ShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L
' make sure selected shapes are sorted by top value
Call SortByTop(rayShapes)

'MsgBox ((UBound(rayShapes) / 2) + 0.5)

If UBound(rayShapes) Mod 2 = 0 Then
A = ((UBound(rayShapes) / 2))
For L = 1 To ((UBound(rayShapes) / 2))
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 2) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top - sngGap * A
A = A + 1
Next M

Else
A = ((UBound(rayShapes) / 2) - 0.5)
For L = 1 To ((UBound(rayShapes) / 2) - 0.5)
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L).Top + sngGap * A
A = A - 1
Next L

A = 1
For M = ((UBound(rayShapes) / 2) + 1.5) To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(M).Top = rayShapes(M).Top - sngGap * A
A = A + 1
Next M

End If


End If


End Sub



Public Sub SortByTop(Arrayin As Variant)
' sort the shapes based on their top value
Dim b_Cont As Boolean
Dim lngCount As Double
Dim vSwap As Shape
Do
    b_Cont = False
    For lngCount = LBound(Arrayin) To UBound(Arrayin) - 1
        Debug.Print Arrayin(lngCount).Name
        If Arrayin(lngCount).Top > Arrayin(lngCount + 1).Top Then
            Set vSwap = Arrayin(lngCount)
            Set Arrayin(lngCount) = Arrayin(lngCount + 1)
            Set Arrayin(lngCount + 1) = vSwap
            b_Cont = True
        End If
    Next lngCount
Loop Until Not b_Cont
'release objects
Set vSwap = Nothing
End Sub

Sub SortByLeft(Arrayin As Variant)
' sort the shapes based on their left value
Dim b_Cont As Boolean
Dim lngCount As Double
Dim vSwap As Shape
Do
    b_Cont = False
    For lngCount = LBound(Arrayin) To UBound(Arrayin) - 1
        Debug.Print Arrayin(lngCount).Name
        If Arrayin(lngCount).Left > Arrayin(lngCount + 1).Left Then
            Set vSwap = Arrayin(lngCount)
            Set Arrayin(lngCount) = Arrayin(lngCount + 1)
            Set Arrayin(lngCount + 1) = vSwap
            b_Cont = True
        End If
    Next lngCount
Loop Until Not b_Cont
'release objects
Set vSwap = Nothing
End Sub




Function cm2Points(inVal As Single) As Single
'convert cm to points
cm2Points = inVal * 28.346
End Function




