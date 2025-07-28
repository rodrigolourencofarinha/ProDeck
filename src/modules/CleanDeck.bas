Attribute VB_Name = "CleanDeck"
Sub Remove_All_Comments()
Dim oSl As slide
Dim oCom As Comment
Dim X As Long
Dim mbResult As Integer

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete all comments in this presentation?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------
 
Select Case mbResult

   Case vbOK
       With ActivePresentation
           For Each oSl In .Slides
               For X = oSl.Comments.Count To 1 Step -1
                   oSl.Comments(X).Delete
               Next
           Next
       End With
       
   Case vbCancel
   'Do nothing and allow the macro to run
    
End Select

End Sub
Sub Remove_All_Notes()
Dim objSlide As slide
Dim objShape As Shape
Dim Counter As Long
Dim mbResult As Integer

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete all speaker notes in this presentation?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------
 
Select Case mbResult

   Case vbOK
       For Each objSlide In ActiveWindow.Presentation.Slides
           For Each objShape In objSlide.NotesPage.Shapes
               If objShape.TextFrame.HasText Then
                   objShape.TextFrame.TextRange = ""
               End If
           Next
       Next
       
   Case vbCancel
       ' Do nothing and allow the macro to run
       
End Select

End Sub
Sub Remove_All_Animations()
Dim sld As slide
Dim sldM As slide
Dim X As Long
Dim Y As Integer
Dim W As Integer
Dim Z As Integer
Dim Counter As Long
Dim mbResult As Integer

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete animations from all slides?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------

Select Case mbResult
   Case vbOK
   
       'Loop Through Each Slide in ActivePresentation
         For Each sld In ActiveWindow.Presentation.Slides
           'Loop through each animation on slide
             For X = sld.TimeLine.MainSequence.Count To 1 Step -1
               'Remove Each Animation
                 sld.TimeLine.MainSequence.Item(X).Delete
             Next X
         Next sld
         
         
        For Z = 1 To ActivePresentation.Designs.Count
           
           On Error Resume Next
           For W = ActivePresentation.Designs(Z).SlideMaster.TimeLine.MainSequence.Count To 1 Step -1
            'Remove Each Animation
              ActivePresentation.Designs(Z).SlideMaster.TimeLine.MainSequence.Item(W).Delete
           Next W
           
        Next Z
         
        For Z = 1 To ActivePresentation.Designs.Count
           
           For W = ActivePresentation.Designs(Z).SlideMaster.CustomLayouts.Count To 1 Step -1
            
               For Y = ActivePresentation.Designs(Z).SlideMaster.CustomLayouts(W).TimeLine.MainSequence.Count To 1 Step -1
                 'Remove Each Animation
                   ActivePresentation.Designs(Z).SlideMaster.CustomLayouts(W).TimeLine.MainSequence.Item(Y).Delete
                   
               Next Y
               
           Next W
           
        Next Z
        
        ActiveWindow.ViewType = ppViewSlide
        ActiveWindow.ViewType = ppViewNormal
         
   
       ' Do nothing and allow the macro to run
End Select

End Sub
Sub Remove_Transitions()
Dim sld As slide
Dim Counter As Long
Dim mbResult As Integer

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete transitions from all slides?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------
 
 Select Case mbResult
    Case vbOK
        'Loop Through Each Slide in ActivePresentation
            For Each sld In ActiveWindow.Presentation.Slides
                'Remove Each Animation
                sld.SlideShowTransition.AdvanceOnTime = False
                sld.SlideShowTransition.AdvanceOnClick = True
                sld.SlideShowTransition.EntryEffect = ppEffectNone
            Next sld
         
        For Z = 1 To ActivePresentation.Designs.Count
        
            On Error Resume Next
            ActivePresentation.Designs(Z).SlideMaster.SlideShowTransition.AdvanceOnTime = False
            ActivePresentation.Designs(Z).SlideMaster.SlideShowTransition.AdvanceOnClick = True
            ActivePresentation.Designs(Z).SlideMaster.SlideShowTransition.EntryEffect = ppEffectNone
           
           For W = ActivePresentation.Designs(Z).SlideMaster.CustomLayouts.Count To 1 Step -1
            
                On Error Resume Next
                ActivePresentation.Designs(Z).SlideMaster.CustomLayouts(W).SlideShowTransition.AdvanceOnTime = False
                ActivePresentation.Designs(Z).SlideMaster.CustomLayouts(W).SlideShowTransition.AdvanceOnClick = True
                ActivePresentation.Designs(Z).SlideMaster.CustomLayouts(W).SlideShowTransition.EntryEffect = ppEffectNon
               
           Next W
           
        Next Z
        
        ActiveWindow.ViewType = ppViewSlide
        ActiveWindow.ViewType = ppViewNormal
        
    Case vbCancel
        ' Do nothing and allow the macro to run
End Select
End Sub
Sub Remove_Unused_Layouts()
Dim I As Integer
Dim J As Integer
Dim Counter As Long
Dim oPres As Presentation
Dim mbResult As Integer

'START: Message box prompt ------------------------------------------------------
mbResult = MsgBox("Delete all unused layouts from this presentation?", _
 vbOKCancel + vbExclamation, "Confirm Action")
'END: Message box prompt --------------------------------------------------------
 
Select Case mbResult

   Case vbOK
       Set oPres = ActivePresentation
       On Error Resume Next
       With oPres
           For I = 1 To .Designs.Count
               For J = .Designs(I).SlideMaster.CustomLayouts.Count To 1 Step -1
                   .Designs(I).SlideMaster.CustomLayouts(J).Delete
                       Next
           Next I
       End With
       
   Case vbCancel
        ' Do nothing and allow the macro to run
        
End Select

End Sub
Sub CompressPicture()

Dim sld As slide
Dim shp As Shape
Dim sldOne As slide
Dim W As Long


If ActiveWindow.Selection.Type = ppSelectionShapes Then
    If ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
        Application.CommandBars.ExecuteMso "PicturesCompress"
        Exit Sub
    End If
End If

ActiveWindow.ViewType = ppViewSlide
ActiveWindow.ViewType = ppViewNormal

Set sldOne = ActiveWindow.View.slide

W = sldOne.SlideIndex

On Error Resume Next
For Each shp In sldOne.Shapes

    If shp.Type = msoPicture Then
        ActiveWindow.View.GotoSlide sld.SlideIndex
        shp.Select
        Application.CommandBars.ExecuteMso "PicturesCompress"
        Exit Sub
    Else
    End If
    
Next shp


For Each sld In ActiveWindow.Presentation.Slides

    For Each shp In sld.Shapes
    
        If shp.Type = msoPicture Then
            ActiveWindow.View.GotoSlide sld.SlideIndex
            shp.Select
            Application.CommandBars.ExecuteMso "PicturesCompress"
            Exit Sub
        Else
        End If
        
    Next shp
    
Next sld

End Sub





