' Extracted from: MarginToggle.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "MarginToggle"
Sub Margin_Toggle()

Dim shp As Shape

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
'END: Error message message box -------------------------------------------------

If ActiveWindow.Selection.HasChildShapeRange Then

    With ActiveWindow.Selection.ChildShapeRange
    
        'START: Found shape with group -------------------------------------------
            If .TextFrame2.MarginLeft = 0 Then
                With .TextFrame2
                    .MarginLeft = 7.14285714285
                    .MarginRight = 7.142857142857
                    .MarginTop = 3.714285714286
                    .MarginBottom = 3.714285714286
                End With
            Else
                With .TextFrame2
                    .MarginLeft = 0
                    .MarginRight = 0
                    .MarginTop = 0
                    .MarginBottom = 0
                End With
            End If
        'END: Shape with group ---------------------------------------------------
   
   End With

Else

    'START: Found shape without group -------------------------------------------
        If ActiveWindow.Selection.ShapeRange().TextFrame2.MarginLeft = 0 Then
            With ActiveWindow.Selection.ShapeRange().TextFrame2
                .MarginLeft = 7.14285714285
                .MarginRight = 7.142857142857
                .MarginTop = 3.714285714286
                .MarginBottom = 3.714285714286
            End With
        Else
            With ActiveWindow.Selection.ShapeRange().TextFrame2
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With
        End If
    'END: Shape without group ---------------------------------------------------

End If

End Sub
