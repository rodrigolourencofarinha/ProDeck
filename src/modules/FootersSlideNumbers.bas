' Extracted from: FootersSlideNumbers.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "FootersSlideNumbers"
Sub Fix_Footers()
Dim sld As slide
Dim txt As String


For Each sld In ActivePresentation.Slides
On Error Resume Next
    
    txt = sld.HeadersFooters.Footer.text
    sld.DisplayMasterShapes = msoTrue
    sld.HeadersFooters.Footer.Visible = msoFalse
    sld.HeadersFooters.Footer.Visible = msoTrue
    sld.HeadersFooters.Footer.text = txt
    txt = ""
Next

End Sub

Sub Fix_Slide_Numbers()
Dim sld As slide

For Each sld In ActivePresentation.Slides
On Error Resume Next
    sld.DisplayMasterShapes = msoTrue
    sld.HeadersFooters.SlideNumber.Visible = msoFalse
    sld.HeadersFooters.SlideNumber.Visible = msoTrue
Next

On Error Resume Next
If ActivePresentation.HasTitleMaster Then
    With ActivePresentation.TitleMaster.HeadersFooters
    .SlideNumber.Visible = msoTrue
End With
End If

On Error Resume Next
With ActivePresentation.SlideMaster.HeadersFooters
    .SlideNumber.Visible = msoTrue
End With

End Sub

