' Extracted from: StickyNotes.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "StickyNotes"
Sub GenerateStickyNote()
Dim sldHeight As Single
Dim sldWidth As Single
Dim stnHeight As Single
Dim stnWidth As Single
Dim stnLeft As Single
Dim stnTop As Single
Dim StickyNote As Shape

RandomNumber = Round(Rnd() * 1000000, 0)

Dim NumberOfStickies As Long
NumberOfStickies = 0

For ShapeNumber = 1 To ActiveWindow.Selection.SlideRange.Shapes.Count
    
    If InStr(1, ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
        NumberOfStickies = NumberOfStickies + 1
    End If
    
Next


sldHeight = ActivePresentation.PageSetup.SlideHeight
sldWidth = ActivePresentation.PageSetup.SlideWidth
stnHeight = sldHeight * 0.16
stnWidth = sldWidth * 0.13
stnLeft = sldWidth - stnWidth - 5
stnTop = (stnHeight) * (NumberOfStickies) + 5 * (NumberOfStickies + 1)

Set StickyNote = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, stnLeft, stnTop, stnWidth, stnHeight)

With StickyNote
    .line.Visible = False
    .Fill.ForeColor.RGB = RGB(255, 244, 148)
    .Fill.Transparency = 0
    .Name = "StickyNote" + Str(RandomNumber)
    .Shadow.OffsetX = 0.066
    .Shadow.OffsetY = 1
    .Shadow.size = 100
    .Shadow.Blur = 4
    .Shadow.Transparency = 0.7
    .Shadow.Visible = True
    .AutoShapeType = msoShapeFoldedCorner
    .TextFrame2.MarginLeft = 3.714285714286
    .TextFrame2.MarginRight = 3.714285714286
    .TextFrame2.MarginTop = 3.714285714286
    .TextFrame2.MarginBottom = 3.714285714286
    
    With .TextFrame2
        .MarginBottom = 0.13 * 28.346
        .MarginLeft = 0.25 * 28.346
        .MarginRight = 0.25 * 28.346
        .MarginTop = 0.13 * 28.346
        .VerticalAnchor = msoAnchorTop
        .AutoSize = msoAutoSizeShapeToFitText
        
        With .TextRange
            .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
            '.text = "Note: "
            With .Font
                .size = 11
                .Name = "Arial"
                .Bold = msoFalse
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With
        
    End With
    .Tags.Add "PRODECK STICKYNOTE", NumberOfStickies
    .Select
End With

    
End Sub
Sub ConvertCommentsToStickyNotes()
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To ActiveWindow.Selection.SlideRange.Shapes.Count
        
        If InStr(1, ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Dim CommentsCount As Long
    Dim RepliesCount As Long
    
    For CommentsCount = ActiveWindow.Selection.SlideRange.Comments.Count To 1 Step -1
        
        Set StickyNote = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.SlideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
        
        With StickyNote
            .line.Visible = False
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            .Fill.Transparency = 0.1
            .Name = "StickyNote" + Str(RandomNumber)
            .Left = ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Left
            .Top = ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Top
            .Tags.Add "PRODECK STICKYNOTE", NumberOfStickies
            
            With .TextFrame
                .MarginBottom = 2
                .MarginLeft = 2
                .MarginRight = 2
                .MarginTop = 2
                .VerticalAnchor = msoAnchorTop
                .AutoSize = ppAutoSizeShapeToFitText
                
                With .TextRange
                    .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                    .text = ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Author & " (" & ActiveWindow.Selection.SlideRange.Comments(CommentsCount).AuthorInitials & "):" & vbNewLine & ActiveWindow.Selection.SlideRange.Comments(CommentsCount).text
                    With .Font
                        .size = 10
                        .color.RGB = RGB(0, 0, 0)
                    End With
                    
                    For RepliesCount = ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Replies.Count To 1 Step -1
                        
                        .text = .text & vbNewLine & vbNewLine & ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).Author & " (" & ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).AuthorInitials & "):" & vbNewLine & ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).text
                        
                    Next
                    
                End With
                
            End With
        End With
        
        ActiveWindow.Selection.SlideRange.Comments(CommentsCount).Delete
        NumberOfStickies = NumberOfStickies + 1
    Next
    
End Sub
Sub MoveStickyNotesOffSlide()
    
    For ShapeNumber = 1 To ActiveWindow.Selection.SlideRange.Shapes.Count
        
        If InStr(1, ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            
            ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "PRODECK OLD POSITION TOP", CStr(ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Top)
            ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "PRODECK OLD POSITION LEFT", CStr(ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Left)
            
            
            With ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            End If
        
    Next
    
End Sub
Sub MoveStickyNotesOnSlide()
    
    For ShapeNumber = 1 To ActiveWindow.Selection.SlideRange.Shapes.Count
        On Error Resume Next
        If InStr(1, ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Top = CLng(ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Tags("PRODECK OLD POSITION TOP"))
            ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Left = CLng(ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Tags("PRODECK OLD POSITION LEFT"))
            
        End If
        On Error GoTo 0
    Next
    
End Sub
Sub DeleteStickyNotesOnSlide()
    
    For ShapeNumber = ActiveWindow.Selection.SlideRange.Shapes.Count To 1 Step -1
        
        If InStr(1, ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            ActiveWindow.Selection.SlideRange.Shapes(ShapeNumber).Delete
        End If
        
    Next
End Sub
Sub DeleteStickyNotesOnAllSlides()

    Dim PresentationSlide As slide
    For Each PresentationSlide In ActiveWindow.Presentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
                PresentationSlide.Shapes(ShapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub
Sub MoveStickyNotesOnAllSlides()
    Dim PresentationSlide As slide
    
    For Each PresentationSlide In ActiveWindow.Presentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            On Error Resume Next
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            PresentationSlide.Shapes(ShapeNumber).Top = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("PRODECK OLD POSITION TOP"))
            PresentationSlide.Shapes(ShapeNumber).Left = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("PRODECK OLD POSITION LEFT"))
            End If
            On Error GoTo 0
        Next
        
    Next
    
End Sub
Sub MoveStickyNotesOffAllSlides()
    Dim PresentationSlide As slide
    
    For Each PresentationSlide In ActiveWindow.Presentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
                
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "PRODECK OLD POSITION TOP", CStr(PresentationSlide.Shapes(ShapeNumber).Top)
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "PRODECK OLD POSITION LEFT", CStr(PresentationSlide.Shapes(ShapeNumber).Left)
            
            
            With PresentationSlide.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            
            End If
            
        Next
        
    Next
    
End Sub



