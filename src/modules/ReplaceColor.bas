' Extracted from: ReplaceColor.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "ReplaceColor"


Sub DemoShowColorMessage(ByVal color As Long)

Dim myColor As ColorPickerUtils.PickColor
myColor = ColorPickerUtils.GetRGBFromLong(color)

MsgBox "You chose a color: " & vbCrLf & _
        "Long value = " & color & vbCrLf & _
        "RGB value = (" & myColor.red & _
        ", " & myColor.green & ", " & myColor.blue & ")"
            
End Sub
Sub Replace_Fill_Colors()

Dim oSh As Shape
Dim oSl As slide
Dim X As Long
Dim Y As Long
Dim I As Long
Dim J As Long
Dim MainColor As Long, SecondColor As Long

On Error Resume Next

'START: Activate form ---------------------------------------------------------
'Chose MainColor
With ColorPickerForm
    .Label1.Caption = "Step 1 of 2: Pick a color to replace"
End With

MainColor = ColorPicker

If continue = False Then
Exit Sub
End If
    
'Choose SecondColors
With ColorPickerForm
    .Label1.Caption = "Step 2 of 2: Pick the new color"
End With
SecondColor = ColorPicker

If continue = False Then
Exit Sub
End If

'END: Activate form ---------------------------------------------------------

For Each oSl In ActiveWindow.Selection.SlideRange

    For Each oSh In oSl.Shapes
    
        With oSh
        
            Select Case .Type
            Case Is = msoGroup
            
                For X = 1 To .GroupItems.Count
                
                    With .GroupItems(X)
                    
                        ' Fill
                        If .Fill.ForeColor.RGB = MainColor Then
                            .Fill.ForeColor.RGB = SecondColor
                        End If
                        
                    End With
                    
                Next X
                
            Case Is = msoTable
            
                   For I = 1 To .Table.Rows.Count
                   
                       For J = 1 To .Table.Columns.Count
                       
                            If .Table.Rows.Item(I).Cells(J).Shape.Fill.ForeColor.RGB = MainColor Then
                                .Table.Rows.Item(I).Cells(J).Shape.Fill.ForeColor.RGB = SecondColor
                            End If
                            
                       Next J
                       
                   Next I
                   
            Case Else
            
               ' Fill
               If .Fill.ForeColor.RGB = MainColor Then
                   .Fill.ForeColor.RGB = SecondColor
               End If
               
            End Select
        
        End With ' oSh
        
    Next oSh
    
Next oSl
    
End Sub
Sub Replace_Line_Colors()

Dim oSh As Shape
Dim oSl As slide
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim MainColor As Long, SecondColor As Long

On Error Resume Next

'START: Activate form ---------------------------------------------------------
'Chose MainColor
With ColorPickerForm
    .Label1.Caption = "Step 1 of 2: Pick a color to replace"
End With
MainColor = ColorPicker

If continue = False Then
Exit Sub
End If

'Choose SecondColors
With ColorPickerForm
    .Label1.Caption = "Step 2 of 2: Pick the new color"
End With
SecondColor = ColorPicker

If continue = False Then
Exit Sub
End If

'END: Activate form ---------------------------------------------------------

For Each oSl In ActiveWindow.Selection.SlideRange

    For Each oSh In oSl.Shapes
    
        With oSh
        
            Select Case .Type
            Case Is = msoGroup
            
                For X = 1 To .GroupItems.Count
                
                    With .GroupItems(X)
                        ' Line
                        If .line.Visible Then
                            If .line.ForeColor.RGB = MainColor Then
                                .line.ForeColor.RGB = SecondColor
                            End If
                        End If
                    End With
                    
                Next X
                
         Case Else
         
            ' Line
            If .line.Visible Then
                If .line.ForeColor.RGB = MainColor Then
                    .line.ForeColor.RGB = SecondColor
                End If
            End If
            
        End Select
        
        End With ' oSh
        
    Next oSh
    
Next oSl

End Sub
Sub Replace_Font_Colors()

Dim oSh As Shape
Dim oSl As slide
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim W As Long
Dim I As Long
Dim J As Long
Dim MainColor As Long, SecondColor As Long

On Error Resume Next

    'START: Activate form ---------------------------------------------------------
    'Chose MainColor
    With ColorPickerForm
        .Label1.Caption = "Step 1 of 2: Pick a color to replace"
    End With
    MainColor = ColorPicker
    
    If continue = False Then
    Exit Sub
    End If
    
    'Choose SecondColors
    With ColorPickerForm
        .Label1.Caption = "Step 2 of 2: Pick the new color"
    End With
    SecondColor = ColorPicker
    
    If continue = False Then
    Exit Sub
    End If
    
    'END: Activate form ---------------------------------------------------------

    For Each oSl In ActiveWindow.Selection.SlideRange
    
        For Each oSh In oSl.Shapes
        
            With oSh
            
                Select Case .Type
                
                Case Is = msoGroup
                
                    For X = 1 To .GroupItems.Count
                    
                        With .GroupItems(X)
                        
                            ' Text
                            If .HasTextFrame Then
                            
                                If .TextFrame.HasText Then
                                
                                    For Y = 1 To .TextFrame.TextRange.Runs.Count
                                    
                                        If .TextFrame.TextRange.Runs(Y).Font.color.RGB = MainColor Then
                                            .TextFrame.TextRange.Runs(Y).Font.color.RGB = SecondColor
                                        End If
                                        
                                    Next
                                    
                                End If
                                
                            End If
                            
                        End With
                        
                    Next X
                    
                Case Is = msoTable
                
                       For I = 1 To .Table.Rows.Count
                       
                           For J = 1 To .Table.Columns.Count
                           
                            ' Text
                            If .Table.Rows.Item(I).Cells(J).Shape.HasTextFrame Then
                            
                                If .Table.Rows.Item(I).Cells(J).Shape.TextFrame.HasText Then
                                
                                    For W = 1 To .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs.Count
                                    
                                        If .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs(W).Font.color.RGB = MainColor Then
                                            .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs(W).Font.color.RGB = SecondColor
                                        End If
                                        
                                    Next W
                                    
                                End If
                                
                            End If
                            
                           Next J
                           
                       Next I
                       
                 Case Else
                 
                    ' Text
                    If .HasTextFrame Then
                    
                        If .TextFrame.HasText Then
                        
                            For Z = 1 To .TextFrame.TextRange.Runs.Count
                            
                                If .TextFrame.TextRange.Runs(Z).Font.color.RGB = MainColor Then
                                    .TextFrame.TextRange.Runs(Z).Font.color.RGB = SecondColor
                                End If
                                
                            Next
                            
                        End If
                        
                    End If
                    
                End Select
            
            End With ' oSh
            
        Next oSh
        
    Next oSl

End Sub
Sub Replace_All_Colors()

Dim oSh As Shape
Dim oSl As slide
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim W As Long
Dim I As Long
Dim J As Long
Dim MainColor As Long, SecondColor As Long

On Error Resume Next

'START: Activate form ---------------------------------------------------------
'Chose MainColor
With ColorPickerForm
    .Label1.Caption = "Step 1 of 2: Pick a color to replace"
End With
MainColor = ColorPicker

If continue = False Then
Exit Sub
End If

'Choose SecondColors
With ColorPickerForm
    .Label1.Caption = "Step 2 of 2: Pick the new color"
End With
SecondColor = ColorPicker

If continue = False Then
Exit Sub
End If

'END: Activate form ---------------------------------------------------------

For Each oSl In ActiveWindow.Selection.SlideRange

    For Each oSh In oSl.Shapes
    
        With oSh
        
            Select Case .Type
            Case Is = msoGroup
            
                For X = 1 To .GroupItems.Count
                
                    With .GroupItems(X)
                        
                        ' Fill
                        If .Fill.ForeColor.RGB = MainColor Then
                            .Fill.ForeColor.RGB = SecondColor
                        End If
                        
                        ' Line
                        If .line.Visible Then
                            If .line.ForeColor.RGB = MainColor Then
                                .line.ForeColor.RGB = SecondColor
                            End If
                        End If
                        
                        ' Text
                        If .HasTextFrame Then
                            If .TextFrame.HasText Then
                                For Y = 1 To .TextFrame.TextRange.Runs.Count
                                    If .TextFrame.TextRange.Runs(Y).Font.color.RGB = MainColor Then
                                        .TextFrame.TextRange.Runs(Y).Font.color.RGB = SecondColor
                                    End If
                                Next
                            End If
                        End If
    
                    End With
                    
                Next X
                
            Case Is = msoTable
            
                   For I = 1 To .Table.Rows.Count
                   
                       For J = 1 To .Table.Columns.Count
                       
                        ' Text
                        If .Table.Rows.Item(I).Cells(J).Shape.HasTextFrame Then
                            If .Table.Rows.Item(I).Cells(J).Shape.TextFrame.HasText Then
                                For W = 1 To .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs.Count
                                    If .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs(W).Font.color.RGB = MainColor Then
                                        .Table.Rows.Item(I).Cells(J).Shape.TextFrame.TextRange.Runs(W).Font.color.RGB = SecondColor
                                    End If
                                Next W
                            End If
                        End If
                        
                        ' Fill
                        If .Table.Rows.Item(I).Cells(J).Shape.Fill.ForeColor.RGB = MainColor Then
                            .Table.Rows.Item(I).Cells(J).Shape.Fill.ForeColor.RGB = SecondColor
                        End If
                        
                       Next J
                       
                   Next I
                   
         Case Else
         
            ' Fill
            If .Fill.ForeColor.RGB = MainColor Then
                .Fill.ForeColor.RGB = SecondColor
            End If
            
            ' Line
            If .line.Visible Then
                If .line.ForeColor.RGB = MainColor Then
                    .line.ForeColor.RGB = SecondColor
                End If
            End If
            
            ' Text
            If .HasTextFrame Then
                If .TextFrame.HasText Then
                    For Z = 1 To .TextFrame.TextRange.Runs.Count
                        If .TextFrame.TextRange.Runs(Z).Font.color.RGB = MainColor Then
                            .TextFrame.TextRange.Runs(Z).Font.color.RGB = SecondColor
                        End If
                    Next
                End If
            End If
            
        End Select
        
        End With ' oSh
        
    Next oSh
    
Next oSl

End Sub

