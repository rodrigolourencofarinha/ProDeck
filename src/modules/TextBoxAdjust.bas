' Extracted from: TextBoxAdjust.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "TextBoxAdjust"
Sub Wrap_Text()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        MsgBox "You must have at least one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

If ActiveWindow.Selection.HasChildShapeRange Then

    With ActiveWindow.Selection.ChildShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame.WordWrap = msoFalse Then
            .TextFrame.WordWrap = msoTrue
            Else
            .TextFrame.WordWrap = msoFalse
            End If
        'END: Shape with group ---------------------------------------------------
   
   End With
   
Else

    With ActiveWindow.Selection.ShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame.WordWrap = msoFalse Then
            .TextFrame.WordWrap = msoTrue
            Else
            .TextFrame.WordWrap = msoFalse
            End If
        'END: Shape with group ---------------------------------------------------
        
   End With
   
End If

End Sub

Sub Shape_Fit_Text()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        MsgBox "You must have at least one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

If ActiveWindow.Selection.HasChildShapeRange Then

    With ActiveWindow.Selection.ChildShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame2.AutoSize = ppAutoSizeNone Or msoAutoSizeTextToFitShape Then
            .TextFrame2.AutoSize = ppAutoSizeShapeToFitText
            Else
            .TextFrame2.AutoSize = ppAutoSizeNone
            End If
        'END: Shape with group ---------------------------------------------------
   
   End With
   
Else

    With ActiveWindow.Selection.ShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame2.AutoSize = ppAutoSizeNone Or msoAutoSizeTextToFitShape Then
            .TextFrame2.AutoSize = ppAutoSizeShapeToFitText
            Else
            .TextFrame2.AutoSize = ppAutoSizeNone
            End If
        'END: Shape with group ---------------------------------------------------
        
   End With
   
End If

End Sub
Sub Text_Overflow()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        MsgBox "You must have at least one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------

If ActiveWindow.Selection.HasChildShapeRange Then

    With ActiveWindow.Selection.ChildShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame2.AutoSize = ppAutoSizeNone Or msoAutoSizeShapeToFitText Then
            .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
            Else
            .TextFrame2.AutoSize = ppAutoSizeNone
            End If
        'END: Shape with group ---------------------------------------------------
   
   End With
   
Else

    With ActiveWindow.Selection.ShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame2.AutoSize = ppAutoSizeNone Or msoAutoSizeShapeToFitText Then
            .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
            Else
            .TextFrame2.AutoSize = ppAutoSizeNone
            End If
        'END: Shape with group ---------------------------------------------------
        
   End With
   
End If

End Sub

Sub Do_Not_Autofit()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        MsgBox "You must have at least one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box --------------------------------------------------


If ActiveWindow.Selection.HasChildShapeRange Then
    ActiveWindow.Selection.ChildShapeRange.TextFrame2.AutoSize = ppAutoSizeNone
Else
    ActiveWindow.Selection.ShapeRange.TextFrame2.AutoSize = ppAutoSizeNone
End If
    
End Sub
