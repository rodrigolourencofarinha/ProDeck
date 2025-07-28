' Extracted from: SpellcheckLanguage.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "SpellcheckLanguage"

Function ChangeSpellCheckLanguage(TargetLanguageID As String)

    Dim TargetLanguage As String
    
    If TargetLanguageID = msoLanguageIDBrazilianPortuguese Then
        TargetLanguage = "Portuguese (Brazil)"
    ElseIf TargetLanguageID = msoLanguageIDEnglishUS Then
        TargetLanguage = "English (US)"
    Else
        TargetLanguage = "Spanish (Spain)"
    End If
    
    
    ActivePresentation.DefaultLanguageID = TargetLanguageID
    
    
    Dim PresentationSlide As PowerPoint.slide
    Dim SlideShape  As PowerPoint.Shape
    Dim SlideSmartArtNode As SmartArtNode
    Dim GroupCount  As Integer
    
    
    #If Mac Then
    'Mac does not (yet) support property .HasHandoutMaster
        
    On Error Resume Next
    For Each SlideShape In ActivePresentation.HandoutMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    On Error GoTo 0

    #Else
    
    If ActivePresentation.HasHandoutMaster Then
    For Each SlideShape In ActivePresentation.HandoutMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #End If
               
    If ActivePresentation.HasTitleMaster Then
    For Each SlideShape In ActivePresentation.TitleMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #If Mac Then
    'Mac does not (yet) support property .HasNotesMaster
        
    On Error Resume Next
    For Each SlideShape In ActivePresentation.NotesMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    On Error GoTo 0
        
    #Else
    
    If ActivePresentation.HasNotesMaster Then
    For Each SlideShape In ActivePresentation.NotesMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #End If
    
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
    
        For Each SlideShape In PresentationSlide.Shapes
            ChangeShapeSpellCheckLanguage SlideShape, TargetLanguageID
        Next SlideShape
    Next PresentationSlide
        
    For Each SlideShape In ActivePresentation.SlideMaster.Shapes
        ChangeShapeSpellCheckLanguage SlideShape, TargetLanguageID
    Next
    
    ActivePresentation.DefaultLanguageID = TargetLanguageID
    'MsgBox "Changed spellcheck language to " + TargetLanguage + " on all slides."
    
End Function

Sub ChangeShapeSpellCheckLanguage(SlideShape, TargetLanguageID)
On Error Resume Next
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ChangeShapeSpellCheckLanguage SlideShapeChild, TargetLanguageID
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
                       
        End If
        
        If SlideShape.HasTable Then
            For TableRow = 1 To SlideShape.Table.Rows.Count
                    For TableColumn = 1 To SlideShape.Table.Columns.Count
                        SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame2.TextRange.LanguageID = TargetLanguageID
                    Next
            Next
        End If
        
        If SlideShape.HasSmartArt Then
            
            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.Count
                
                For Each SlideSmartArtNode In SlideShape.SmartArt.AllNodes
                    SlideSmartArtNode.TextFrame2.TextRange.LanguageID = TargetLanguageID
                 Next
                            
            Next
            
        End If
        
    End If
    
End Sub
Sub Language_BR()
    ChangeSpellCheckLanguage (msoLanguageIDBrazilianPortuguese)
End Sub
Sub Language_US()
    ChangeSpellCheckLanguage (msoLanguageIDEnglishUS)
End Sub
Sub Language_ES()
    ChangeSpellCheckLanguage (msoLanguageIDSpanish)
End Sub
Sub Language_FR()
    ChangeSpellCheckLanguage (msoLanguageIDFrench)
End Sub
Sub Language_DE()
    ChangeSpellCheckLanguage (msoLanguageIDGerman)
End Sub
