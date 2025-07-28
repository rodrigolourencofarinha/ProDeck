Attribute VB_Name = "ExportComments"
Option Explicit
Sub Export_Comments_Text()

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
Dim oCom As Comment
Dim DotPosition As Integer
Dim RepliesCount As Integer


'Strip extension from filename
DotPosition = InStrRev(ActivePresentation.Name, ".")

If DotPosition > 0 Then
    sName = Left(ActivePresentation.Name, DotPosition - 1)
Else
    sName = ActivePresentation.Name
End If

sPath = Environ("USERPROFILE") & "\Desktop\"

' Get a filename to store the collected text
strFileName = sPath & "Comments_from_" & sName & ".txt"

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Comments_from_" & sName & ".txt " _
& "will be saved under " & Environ("USERPROFILE") & "\Desktop." _
& vbCrLf & vbCrLf _
& "Do you wish to continue?", _
 vbOKCancel + vbQuestion, "Export Comments")
'END: Message box prompt --------------------------------------------------------
 
Select Case mbResult

   Case vbOK
   
        ' File exists?
        Dim strFileExists As String
        strFileExists = Dir(strFileName)
        If strFileExists <> "" Then
           mbContinue = MsgBox("The file Comments_from_" & sName & ".txt already exists." & vbCrLf & "Do you want to replace the existing file?", vbOKCancel + vbExclamation, "Confirm Save")
        Else
        End If

        ' is the path valid?  crude but effective test:  try to create the file.
        intFileNum = FreeFile()
        On Error Resume Next
        Open strFileName For Output As intFileNum
        If Err.Number <> 0 Then     ' we have a problem
        
            sPath = Environ("USERPROFILE") & "\OneDrive\Desktop\"
            strFileName = sPath & "Comments_from_" & sName & ".txt"
            
            ' File exists?
            strFileExists = Dir(strFileName)
            If strFileExists <> "" Then
            mbContinue = MsgBox("The file Comments_from_" & sName & ".txt already exists." & vbCrLf & "Do you want to replace the existing file?", vbOKCancel + vbExclamation, "Confirm Save")
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
                If Not oSl.Comments.Count = 0 Then
                    strNotesText = strNotesText & "======================================" & vbCrLf
                    strNotesText = strNotesText & "Slide: " & oSl.SlideIndex & vbCrLf
                    strNotesText = strNotesText & "--------------" & vbCrLf
    
                        For Each oCom In oSl.Comments
                            strNotesText = strNotesText & oCom.Author & vbCrLf
                            strNotesText = strNotesText & oCom.DateTime & vbCrLf
                            strNotesText = strNotesText & oCom.text & vbCrLf

                            
                            For RepliesCount = oCom.Replies.Count To 1 Step -1
                                strNotesText = strNotesText & "*** Reply ***" & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).Author & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).DateTime & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).text & vbCrLf
                            Next
                            
                            strNotesText = strNotesText & "--------------" & vbCrLf
                            
                        Next oCom
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
                If Not oSl.Comments.Count = 0 Then
                    strNotesText = strNotesText & "======================================" & vbCrLf
                    strNotesText = strNotesText & "Slide: " & oSl.SlideIndex & vbCrLf
                    strNotesText = strNotesText & "--------------" & vbCrLf
    
                        For Each oCom In oSl.Comments
                            strNotesText = strNotesText & oCom.Author & vbCrLf
                            strNotesText = strNotesText & oCom.DateTime & vbCrLf
                            strNotesText = strNotesText & oCom.text & vbCrLf

                            
                            For RepliesCount = oCom.Replies.Count To 1 Step -1
                                strNotesText = strNotesText & "*** Reply ***" & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).Author & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).DateTime & vbCrLf
                                strNotesText = strNotesText & oCom.Replies(RepliesCount).text & vbCrLf
                            Next
                            
                            strNotesText = strNotesText & "--------------" & vbCrLf
                            
                        Next oCom
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








