Attribute VB_Name = "ExportNotes"
Option Explicit
Sub Export_Notes_Text()
    
    Dim oSlides As Slides
    Dim oSl As slide
    Dim oSh As Shape
    Dim strNotesText As String
    Dim strFileName As String
    Dim intFileNum As Integer
    Dim lngReturn As Long
    Dim sPath As String
    Dim sName As String
    Dim mbResult As Integer
    Dim mbContinue As Integer
    Dim DotPosition As Integer
    

    'Strip extension from filename
    DotPosition = InStrRev(ActivePresentation.Name, ".")
    If DotPosition > 0 Then
        sName = Left(ActivePresentation.Name, DotPosition - 1)
    Else
        sName = ActivePresentation.Name
    End If
    
    sPath = Environ("USERPROFILE") & "\Desktop\"
    
    ' Get a filename to store the collected text
    strFileName = sPath & "SpeakerNotes_from_" & sName & ".txt"

    'START: Message box prompt ------------------------------------------------------
    mbResult = MsgBox("SpeakerNotes_from_" & sName & ".txt " _
    & "will be saved under " & Environ("USERPROFILE") & "\Desktop." _
    & vbCrLf & vbCrLf _
    & "Do you wish to continue?", _
     vbOKCancel + vbQuestion, "Export Speaker Notes")
    'END: Message box prompt --------------------------------------------------------
     
        Select Case mbResult
        
           Case vbOK
           
                ' File exists?
                Dim strFileExists As String
                strFileExists = Dir(strFileName)
                
                If strFileExists <> "" Then
                   mbContinue = MsgBox("The file SpeakerNotes_from_" & sName & ".txt already exists." & vbCrLf & "Do you want to replace the existing file?", vbOKCancel + vbExclamation, "Confirm Save")
                Else
                
                End If
        
                ' is the path valid?  crude but effective test:  try to create the file.
                intFileNum = FreeFile()
                On Error Resume Next
                Open strFileName For Output As intFileNum
                If Err.Number <> 0 Then     ' we have a problem
                
                    sPath = Environ("USERPROFILE") & "\OneDrive\Desktop\"
                    strFileName = sPath & "SpeakerNotes_from_" & sName & ".txt"
                    
                    
                    ' File exists?
                    strFileExists = Dir(strFileName)
                    If strFileExists <> "" Then
                    mbContinue = MsgBox("The file SpeakerNotes_from_" & sName & ".txt already exists." & vbCrLf & "Do you want to replace the existing file?", vbOKCancel + vbExclamation, "Confirm Save")
                    Else
                        ' EXECUTE EXTRACT COMMENT ================================================================
                    
                        intFileNum = FreeFile()
                        On Error Resume Next
                        Open strFileName For Output As intFileNum
                        If Err.Number <> 0 Then     ' we have a problem
                            MsgBox "Couldn't create the file: " & strFileName & vbCrLf _
                                & "Please try again."
                            Exit Sub
                        End If
                        Close #intFileNum  ' temporarily

                    
                        ' Get the notes text
                        Set oSlides = ActiveWindow.Presentation.Slides
                        For Each oSl In oSlides
                            If Not NotesText(oSl) = "" Then
                                strNotesText = strNotesText & "======================================" & vbCrLf
                                strNotesText = strNotesText & "Slide: " & oSl.SlideIndex & vbCrLf
                                strNotesText = strNotesText & "Title: " & SlideTitle(oSl) & vbCrLf
                                strNotesText = strNotesText & NotesText(oSl) & vbCrLf
                            End If
                        Next oSl
                        ' now write the text to file
                        Open strFileName For Output As intFileNum
                        Print #intFileNum, strNotesText
                        Close #intFileNum
                        
                        Exit Sub
                        
                        
                        ' EXECUTE EXTRACT COMMENT ================================================================
                    End If
                    
                    intFileNum = FreeFile()
                    On Error Resume Next
                    Open strFileName For Output As intFileNum
                    If Err.Number <> 0 Then     ' we have a problem
                        MsgBox "Couldn't create the file: " & strFileName & vbCrLf _
                            & "Please try again."
                        Exit Sub
                    End If
                    Close #intFileNum  ' temporarily
                End If
                
                Select Case mbContinue
                
                    Case vbOK
                    
                        ' Get the notes text
                        Set oSlides = ActiveWindow.Presentation.Slides
                        For Each oSl In oSlides
                            If Not NotesText(oSl) = "" Then
                                strNotesText = strNotesText & "======================================" & vbCrLf
                                strNotesText = strNotesText & "Slide: " & oSl.SlideIndex & vbCrLf
                                strNotesText = strNotesText & "Title: " & SlideTitle(oSl) & vbCrLf
                                strNotesText = strNotesText & NotesText(oSl) & vbCrLf
                            End If
                        Next oSl
                
                        ' now write the text to file
                        Open strFileName For Output As intFileNum
                        Print #intFileNum, strNotesText
                        Close #intFileNum
                        
                    Case vbCancel
                        Exit Sub
                    End Select
                    
           Case vbCancel
               ' Do nothing and allow the macro to run
               
               Exit Sub
               
        End Select
        
End Sub
Function SlideTitle(oSl As slide) As String

Dim oSh As Shape

For Each oSh In oSl.Shapes

    If oSh.Type = msoPlaceholder Then
    
        If oSh.PlaceholderFormat.Type = ppPlaceholderTitle _
            Or oSh.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
            
            If Len(oSh.TextFrame.TextRange.text) > 0 Then
                SlideTitle = oSh.TextFrame.TextRange.text
            Else
                SlideTitle = "Slide " & CStr(oSl.SlideIndex)
            End If
            
            Exit Function
            
        End If
        
    End If
    
Next

End Function
Function NotesText(oSl As slide) As String

Dim oSh As Shape

For Each oSh In oSl.NotesPage.Shapes

    If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
    
        If oSh.HasTextFrame Then
        
            If oSh.TextFrame.HasText Then
                NotesText = oSh.TextFrame.TextRange.text
            End If
            
        End If
        
    End If
    
Next oSh

End Function











