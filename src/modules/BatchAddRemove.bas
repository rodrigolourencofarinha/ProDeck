' Extracted from: BatchAddRemove.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "BatchAddRemove"
Sub Delete_Similar_Objects()

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
      MsgBox "You must have one object selected.", vbCritical
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -----------------------------------------------

Dim shp As Shape
Dim Shp1 As Shape
Dim sld As slide
Dim sngLastY As Single
Dim sngLastX As Single
Dim sngheight As Single
Dim sngwidth As Single
Dim mbResult As Integer
Dim sngName As String

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete all objects in similar position in this presentation?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------
 
Select Case mbResult

    Case vbOK
    
        Set Shp1 = ActiveWindow.Selection.ShapeRange(1)
        
        sngLastY = Shp1.Top
        sngLastX = Shp1.Left
        sngheight = Shp1.Height
        sngwidth = Shp1.Width
        sngName = Shp1.Name
        
            For Each sld In ActiveWindow.Presentation.Slides
            
                For Each shp In sld.Shapes
                
                If InStr(1, sngName, "PRODECK OBJECTS ALL SLIDES") = 1 Then
                    If shp.Name = sngName Then
                        shp.Delete
                    End If
                Else
                    If shp.Top = sngLastY And shp.Left = sngLastX And shp.Height = sngheight And shp.Width = sngwidth Then
                        shp.Delete
                    End If
                End If
                            

                Next shp
                
            Next sld
            
    Case vbCancel
    ' Do nothing and allow the macro to run
    
End Select

End Sub
Sub Paste_Objects_All_Slides()
On Error Resume Next
Dim sld As slide
Dim oSh As Shape
Dim X As Long

RandomNumber = Round(Rnd() * 1000000, 0)

Set ObjectTemp = ActiveWindow.Selection.SlideRange(1).Shapes.Paste

ObjectTemp.Copy
ObjectTemp.Delete

If Err Then

    MsgBox "Clipboard is empty or contains data which may not be pasted here.", vbCritical
    Err.Clear
    Exit Sub
    
End If

    For X = 1 To ActiveWindow.Selection.SlideRange.Count
        
        DoEvents
        Set ObjectAllSlides = ActiveWindow.Selection.SlideRange(X).Shapes.Paste
        
        With ObjectAllSlides
            .Name = "PRODECK OBJECTS ALL SLIDES" + Str(RandomNumber)
        End With
        
        

    Next X
    
End Sub
Sub Update_Objects_Position()

Dim CrossSlideShapeId As String
Dim X As Long
Dim shp As Shape
Dim Shpl As Shape
Dim sngLastY As Single
Dim sngLastX As Single

Set Shp1 = ActiveWindow.Selection.ShapeRange(1)
sngLastY = Shp1.Top
sngLastX = Shp1.Left
sngName = Shp1.Name

If InStr(1, sngName, "PRODECK OBJECTS ALL SLIDES") = 1 Then
Else
    Exit Sub
End If

If ActiveWindow.Selection.Type = ppSelectionShapes Then

    For X = 1 To ActiveWindow.Presentation.Slides.Count
    
        For Each shp In ActiveWindow.Presentation.Slides(X).Shapes
        
            If shp.Name = sngName Then
                shp.Top = sngLastY
                shp.Left = sngLastX
            End If
            
        Next
        
    Next
    
Else

    MsgBox "No shape selected."
    
End If

End Sub




