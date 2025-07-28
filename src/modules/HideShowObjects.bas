' Extracted from: HideShowObjects.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "HideShowObjects"
Sub Hide_Objects()

If ActiveWindow.Selection.HasChildShapeRange Then
'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ChildShapeRange.Count < 1 Then
      MsgBox "You must have at least one object selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

With ActiveWindow.Selection.ChildShapeRange()
    .Visible = False
End With

Else
'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count < 1 Then
      MsgBox "You must have at least one object selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

With ActiveWindow.Selection.ShapeRange()
    .Visible = False
End With

End If

End Sub
Sub Show_All_Objects()

'Loop through selected slides
Dim sld As slide
Dim shp As Shape

For Each sld In ActiveWindow.Selection.SlideRange

    For Each shp In sld.Shapes
    
        'START: Execution code ----------------------------------------------
        shp.Visible = True
        'END: Execution code ------------------------------------------------
        
    Next shp
    
Next sld
    
End Sub
