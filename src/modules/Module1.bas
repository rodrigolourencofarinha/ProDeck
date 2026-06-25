Attribute VB_Name = "Module1"
Option Explicit

Public Sub Save_Image_Only_Copy()

    Dim p As Presentation, q As Presentation, t As Slide
    Dim i As Long, n As Long, w As Single, h As Single, ew As Long, eh As Long
    Dim folder As String, stem As String, outFile As String, tmp As String, img As String, stepName As String
    Dim errNum As Long, errText As String

    On Error GoTo Fail

    Set p = ActivePresentation
    If p.Path = "" Then MsgBox "Save this presentation before creating an image-only copy.", vbExclamation, "Save Image-Only Copy": Exit Sub

    n = p.Slides.Count
    If n = 0 Then MsgBox "This presentation does not contain any slides.", vbExclamation, "Save Image-Only Copy": Exit Sub

    folder = p.Path
    stem = BaseName(p.Name) & "_image_only"
    outFile = UniquePptx(folder, stem)
    tmp = JoinPath(folder, BaseName(FileNameOnly(outFile)) & "__prodeck_tmp")

    stepName = "Preparing the temporary export folder"
    ResetFolder tmp

    w = p.PageSetup.SlideWidth
    h = p.PageSetup.SlideHeight
    ew = CLng(w * 2)
    eh = CLng(h * 2)

    ProgressForm.Show
    SetProgress 0

    stepName = "Exporting slides as PNG files"
    For i = 1 To n
        img = JoinPath(tmp, "slide_" & Format$(i, "0000") & ".png")
        p.Slides(i).Export img, "PNG", ew, eh
        SetProgress (i / n) * 50
    Next i

    stepName = "Creating the image-only presentation"
    Set q = Presentations.Add
    q.PageSetup.SlideWidth = w
    q.PageSetup.SlideHeight = h

    stepName = "Rebuilding the presentation with image slides"
    For i = 1 To n
        img = JoinPath(tmp, "slide_" & Format$(i, "0000") & ".png")
        If i = 1 Then
            Set t = q.Slides(1)
            Do While t.Shapes.Count > 0
                t.Shapes(1).Delete
            Loop
        Else
            Set t = q.Slides.Add(q.Slides.Count + 1, ppLayoutBlank)
        End If
        t.Shapes.AddPicture img, msoFalse, msoTrue, 0, 0, w, h
        SetProgress 50 + (i / n) * 50
    Next i

    stepName = "Saving the image-only PPTX"
    q.SaveAs outFile, ppSaveAsOpenXMLPresentation
    q.Close
    Set q = Nothing

    stepName = "Cleaning up temporary files"
    ClearFolder tmp
    ProgressForm.Hide

    MsgBox "Image-only copy saved to:" & vbCrLf & outFile, vbInformation, "Save Image-Only Copy"
    Exit Sub

Fail:
    errNum = Err.Number
    errText = Err.Description

    On Error Resume Next
    If Not q Is Nothing Then q.Close
    ClearFolder tmp
    ProgressForm.Hide
    MsgBox ErrTextFor(stepName, errNum, errText), vbCritical, "Save Image-Only Copy"

End Sub

Private Function UniquePptx(ByVal folder As String, ByVal stem As String) As String

    Dim p As String, i As Long

    p = JoinPath(folder, stem & ".pptx")
    i = 2

    Do While Len(Dir(p)) > 0
        p = JoinPath(folder, stem & " (" & i & ").pptx")
        i = i + 1
    Loop

    UniquePptx = p

End Function

Private Sub ResetFolder(ByVal p As String)

    If FolderExists(p) Then ClearFolder p
    MkDir p

End Sub

Private Sub ClearFolder(ByVal p As String)

    Dim f As String

    If Len(p) = 0 Then Exit Sub
    If Not FolderExists(p) Then Exit Sub

    f = Dir(JoinPath(p, "*"))
    Do While f <> ""
        If f <> "." And f <> ".." Then Kill JoinPath(p, f)
        f = Dir
    Loop

    RmDir p

End Sub

Private Function FolderExists(ByVal p As String) As Boolean

    On Error Resume Next
    FolderExists = (GetAttr(p) And vbDirectory) = vbDirectory
    If Err.Number <> 0 Then Err.Clear: FolderExists = False
    On Error GoTo 0

End Function

Private Function JoinPath(ByVal a As String, ByVal b As String) As String

    Dim s As String

    s = PathSep(a)
    If Right$(a, 1) = s Then JoinPath = a & b Else JoinPath = a & s & b

End Function

Private Function PathSep(ByVal p As String) As String

    If InStr(p, "/") > 0 Then
        PathSep = "/"
    ElseIf InStr(p, ":") > 0 And InStr(p, "\") = 0 Then
        PathSep = ":"
    Else
        PathSep = "\"
    End If

End Function

Private Function FileNameOnly(ByVal p As String) As String

    Dim s As String, i As Long

    s = PathSep(p)
    i = InStrRev(p, s)

    If i > 0 Then FileNameOnly = Mid$(p, i + 1) Else FileNameOnly = p

End Function

Private Function BaseName(ByVal p As String) As String

    Dim i As Long

    i = InStrRev(p, ".")
    If i > 0 Then BaseName = Left$(p, i - 1) Else BaseName = p

End Function

Private Function ErrTextFor(ByVal stepName As String, ByVal errNum As Long, ByVal errText As String) As String

    ErrTextFor = "ProDeck could not save the image-only copy."
    If Len(stepName) > 0 Then ErrTextFor = ErrTextFor & vbCrLf & vbCrLf & "Step: " & stepName
    ErrTextFor = ErrTextFor & vbCrLf & "Error: " & errNum
    If Len(errText) > 0 Then ErrTextFor = ErrTextFor & vbCrLf & errText

End Function
