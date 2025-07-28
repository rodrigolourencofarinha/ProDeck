Attribute VB_Name = "SwapPositions"
Function Swap_Positions(Position As String)
   
Dim oshpR As ShapeRange
Dim oShp As Shape
Dim C1 As Single
Dim M1 As Single
Dim C2 As Single
Dim M2 As Single
Dim T1 As Single
Dim T2 As Single
Dim L1 As Single
Dim L2 As Single
Dim B1 As Single
Dim B2 As Single
Dim R1 As Single
Dim R2 As Single

If ActiveWindow.Selection.HasChildShapeRange Then

    'START: Error message message box -----------------------------------------------
    On Error Resume Next
        Err.Clear
       If ActiveWindow.Selection.ChildShapeRange.Count <> 2 Then
          MsgBox "You must have two objects selected.", vbCritical
          Exit Function
        End If
        If Err <> 0 Then
            Exit Function
        End If
    'END: Error message message box -------------------------------------------------
    
    Set oshpR = ActiveWindow.Selection.ChildShapeRange
       C1 = oshpR(1).Left + (oshpR(1).Width) / 2
       C2 = oshpR(2).Left + (oshpR(2).Width) / 2
       M1 = oshpR(1).Top + (oshpR(1).Height) / 2
       M2 = oshpR(2).Top + (oshpR(2).Height) / 2
       T1 = oshpR(1).Top
       T2 = oshpR(2).Top
       L1 = oshpR(1).Left
       L2 = oshpR(2).Left
       B1 = oshpR(1).Top + oshpR(1).Height
       B2 = oshpR(2).Top + oshpR(2).Height
       R1 = oshpR(1).Left + oshpR(1).Width
       R2 = oshpR(2).Left + oshpR(2).Width
    
       
    Select Case Position
        Case "TopLeft"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "TopCenter"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "TopRight"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
        Case "MiddleLeft"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "Center"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "MiddleRight"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
        Case "BottomLeft"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "BottomCenter"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "BottomRight"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
    End Select

Else
       
'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
      MsgBox "You must have two objects selected.", vbCritical
        Exit Function
    End If
    If Err <> 0 Then
        Exit Function
    End If
'END: Error message message box -------------------------------------------------

   Set oshpR = ActiveWindow.Selection.ShapeRange
   C1 = oshpR(1).Left + (oshpR(1).Width) / 2
   C2 = oshpR(2).Left + (oshpR(2).Width) / 2
   M1 = oshpR(1).Top + (oshpR(1).Height) / 2
   M2 = oshpR(2).Top + (oshpR(2).Height) / 2
   T1 = oshpR(1).Top
   T2 = oshpR(2).Top
   L1 = oshpR(1).Left
   L2 = oshpR(2).Left
   B1 = oshpR(1).Top + oshpR(1).Height
   B2 = oshpR(2).Top + oshpR(2).Height
   R1 = oshpR(1).Left + oshpR(1).Width
   R2 = oshpR(2).Left + oshpR(2).Width

   
    Select Case Position
        Case "TopLeft"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "TopCenter"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "TopRight"
            oshpR(1).Top = T2
            oshpR(2).Top = T1
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
        Case "MiddleLeft"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "Center"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "MiddleRight"
            oshpR(1).Top = M2 - (oshpR(1).Height) / 2
            oshpR(2).Top = M1 - (oshpR(2).Height) / 2
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
        Case "BottomLeft"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = L2
            oshpR(2).Left = L1
        Case "BottomCenter"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = C2 - oshpR(1).Width / 2
            oshpR(2).Left = C1 - oshpR(2).Width / 2
        Case "BottomRight"
            oshpR(1).Top = B2 - oshpR(1).Height
            oshpR(2).Top = B1 - oshpR(2).Height
            oshpR(1).Left = R2 - oshpR(1).Width
            oshpR(2).Left = R1 - oshpR(2).Width
    End Select
    
End If

End Function
Sub Swap_TopLeft()
Swap_Positions ("TopLeft")
End Sub
Sub Swap_TopCenter()
Swap_Positions ("TopCenter")
End Sub
Sub Swap_TopRight()
Swap_Positions ("TopRight")
End Sub
Sub Swap_MiddleLeft()
Swap_Positions ("MiddleLeft")
End Sub
Sub Swap_Center()
Swap_Positions ("Center")
End Sub
Sub Swap_MiddleRight()
Swap_Positions ("MiddleRight")
End Sub
Sub Swap_BottomLeft()
Swap_Positions ("BottomLeft")
End Sub
Sub Swap_BottomCenter()
Swap_Positions ("BottomCenter")
End Sub
Sub Swap_BottomRight()
Swap_Positions ("BottomRight")
End Sub

Sub Swap_Text_NoFormat()

Dim text1, text2 As String

If ActiveWindow.Selection.HasChildShapeRange Then
    If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
        If ActiveWindow.Selection.ChildShapeRange(1).HasTextFrame And ActiveWindow.Selection.ChildShapeRange(2).HasTextFrame Then
            text1 = ActiveWindow.Selection.ChildShapeRange(1).TextFrame.TextRange.text
            text2 = ActiveWindow.Selection.ChildShapeRange(2).TextFrame.TextRange.text
            ActiveWindow.Selection.ChildShapeRange(1).TextFrame.TextRange.text = text2
            ActiveWindow.Selection.ChildShapeRange(2).TextFrame.TextRange.text = text1
        Else
            MsgBox "Select two shapes that (can) have text."
        End If
    Else
        MsgBox "Select two shapes to swap their text."
    End If
Else
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        If ActiveWindow.Selection.ShapeRange(1).HasTextFrame And ActiveWindow.Selection.ShapeRange(2).HasTextFrame Then
            text1 = ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.text
            text2 = ActiveWindow.Selection.ShapeRange(2).TextFrame.TextRange.text
            ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.text = text2
            ActiveWindow.Selection.ShapeRange(2).TextFrame.TextRange.text = text1
        Else
            MsgBox "Select two shapes that (can) have text."
        End If
    Else
        MsgBox "Select two shapes to swap their text."
    End If
End If
End Sub
