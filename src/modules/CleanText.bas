' Extracted from: CleanText.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "CleanText"
Sub Remove_Line_Break()
On Error Resume Next
Dim X As Long

If ActiveWindow.Selection.Type = ppSelectionText Then
    
    With ActiveWindow.Selection.TextRange2
        .text = Replace(.text, Chr(13), " ")
        .text = Replace(.text, Chr(10), " ")
        .text = Replace(.text, Chr(11), " ")
        Do While InStr(.text, "  ") > 0
            .Replace FindWhat:="  ", ReplaceWhat:=" "
        Loop
    End With

ElseIf ActiveWindow.Selection.HasChildShapeRange Then

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ChildShapeRange.Count < 1 Then
      MsgBox "You must have one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------
    
    For X = 1 To ActiveWindow.Selection.ChildShapeRange.Count
        If ActiveWindow.Selection.ChildShapeRange(X).HasTextFrame Then
            With ActiveWindow.Selection.ChildShapeRange(X).TextFrame.TextRange
                .text = Replace(.text, Chr(13), " ")
                .text = Replace(.text, Chr(10), " ")
                .text = Replace(.text, Chr(11), " ")
                Do While InStr(.text, "  ") > 0
                    .Replace FindWhat:="  ", ReplaceWhat:=" "
                Loop
            End With
        Else
        End If
    Next

Else


'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count < 1 Then
      MsgBox "You must have one shape selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------
    
    For X = 1 To ActiveWindow.Selection.ShapeRange.Count
        If ActiveWindow.Selection.ShapeRange(X).HasTextFrame Then
            With ActiveWindow.Selection.ShapeRange(X).TextFrame.TextRange
                .text = Replace(.text, Chr(13), " ")
                .text = Replace(.text, Chr(10), " ")
                .text = Replace(.text, Chr(11), " ")
                Do While InStr(.text, "  ") > 0
                    .Replace FindWhat:="  ", ReplaceWhat:=" "
                Loop
            End With
        Else
        End If
    Next

End If

End Sub
Sub Remove_Spaces()
Dim shpTextRng As TextRange
Dim sld As slide
Dim shp As Shape
Dim X As Long
Dim I As Long
Dim J As Long

For Each sld In ActiveWindow.Selection.SlideRange

    For Each shp In sld.Shapes
    
        With shp
            Select Case .Type
            
                Case Is = msoGroup
                
                    For X = 1 To .GroupItems.Count
                    
                        If .GroupItems(X).HasTextFrame Then
                        
                            If .GroupItems(X).TextFrame.HasText Then
                            
                                Set shpTextRng = shp.GroupItems(X).TextFrame.TextRange
                                
                                Do While InStr(shpTextRng.text, "  ") > 0
                                    shpTextRng.Replace FindWhat:="  ", ReplaceWhat:=" "
                                Loop
                                
                            End If
                            
                        End If
                        
                    Next X
                    
                Case Is = msoTable
                
                        For I = 1 To .Table.Rows.Count
                        
                            For J = 1 To .Table.Columns.Count
                            
                                Set shpTextRng = shp.Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange
                                Do While InStr(shpTextRng.text, "  ") > 0
                                    shpTextRng.text = Replace(shpTextRng.text, "  ", " ")
                                Loop
                                
                            Next J
                            
                        Next I
            End Select
            
        End With
        
    Next shp
    
Next sld

For Each sld In ActiveWindow.Selection.SlideRange

    For Each shp In sld.Shapes
    
        If shp.HasTextFrame Then
            Set shpTextRng = shp.TextFrame.TextRange
            Do While InStr(shpTextRng.text, "  ") > 0
                shpTextRng.Replace FindWhat:="  ", ReplaceWhat:=" "
            Loop
        End If
        
    Next shp
    
Next sld

End Sub


