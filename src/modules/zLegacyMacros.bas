Attribute VB_Name = "zLegacyMacros"
Sub Grid_Shapes()
'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
      MsgBox "Select a shape covering the area you want to fill with a grid.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'START: Error message message box -----------------------------------------------

Dim oSh As Shape
Dim oShN As Shape
Dim oSld As slide
Dim sngwidth As Single  ' width/height of a grid rect
Dim sngheight As Single
Dim lCols As Long
Dim lRows As Long
Dim X As Long   ' which col across are we making
Dim Y As Long   ' which row down are we making
Dim sngLeft As Single   ' where to draw current rectangle
Dim sngTop As Single
Dim sTemp As Variant
Dim shptype As Long
    If ActiveWindow.Selection.ShapeRange.Type = msoGroup Then
    MsgBox "Select an object and not a group.", vbCritical
    Exit Sub
    End If
    ' get rows/cols from user
    sTemp = InputBox("Step 1 of 2: How many columns?", "Create Grid of Shapes")
        If (StrPtr(sTemp) = 0) Then
        Exit Sub
        End If
        If Not IsNumeric(sTemp) Then
        MsgBox "Columns must be a positive integer.", vbCritical
        Exit Sub
        End If
        If CLng(sTemp) > 0 Then
            lCols = CLng(sTemp)
            sTemp = InputBox("Step 2 of 2: How many Rows?", "Create Grid of Shapes")
            If (StrPtr(sTemp) = 0) Then
            Exit Sub
            End If
            If Not IsNumeric(sTemp) Then
            MsgBox "Rows must be a positive integer.", vbCritical
            Exit Sub
            End If
            If CLng(sTemp) > 0 Or IsNumeric(sTemp) Then
                lRows = CLng(sTemp)
            Else
                MsgBox "Rows must be a positive integer.", vbCritical
                Exit Sub
            End If
         Else
            MsgBox "Columns must be a positive integer.", vbCritical
            Exit Sub
         End If
    Set oSh = ActiveWindow.Selection.ShapeRange(1)
    Set oSld = oSh.Parent
    sngwidth = oSh.Width / lCols
    sngheight = oSh.Height / lRows
    shptype = oSh.AutoShapeType
    oSh.PickUp
    For X = 0 To lCols - 1
        For Y = 0 To lRows - 1
          ' with osld.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
            Set oShN = oSld.Shapes.AddShape(shptype, oSh.Left + X * sngwidth, oSh.Top + Y * sngheight, sngwidth, sngheight)
            With oShN
                .Apply
                Call .Tags.Add("Grid", "YES")
                .TextFrame2.TextRange.text = oSh.TextFrame2.TextRange.text
                .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
            End With
        Next
    Next
    oSh.Delete
End Sub


