Attribute VB_Name = "SlideUtils"
' Various slide utilities that warrant their own function
Option Explicit

' Insert a new slide in the active presentation one slide ahead
' of the current slide
Function InsertSlide() As slide
    ' make a new slide
    Dim curSlide As slide
    Set curSlide = CurrentSlide("InsertSlide: Could not insert slide")
    
    If curSlide Is Nothing Then
        Set InsertSlide = Nothing
        Exit Function
    End If
    
    Set InsertSlide = ActiveWindow.Presentation.Slides.AddSlide(curSlide.SlideIndex + 1, curSlide.CustomLayout)
    ActiveWindow.View.GotoSlide InsertSlide.SlideIndex
End Function


' get the current slide
' if any error occurs, such as PPT being in an
' uncooperative view, then "Nothing" is returned
' and an error message box is displayed with an optional message
' if no message box is desired, pass in Message:="" empty string as input
Function CurrentSlide(Optional Message As String = "Could not get current slide") As slide
    On Error GoTo SlideError
       Set CurrentSlide = ActiveWindow.View.slide
    On Error GoTo 0
    Exit Function
    
SlideError:
    Set CurrentSlide = Nothing
    If Not Message = "" Then
        MsgBox Message
    End If
    
End Function


