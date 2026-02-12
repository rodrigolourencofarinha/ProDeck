Attribute VB_Name = "Controls"
Sub Control_Text_Box()
On Error Resume Next
Application.CommandBars.ExecuteMso ("TextBoxInsert")

End Sub


Sub PasteValue()
On Error Resume Next
Application.CommandBars.ExecuteMso ("PasteTextOnly")

End Sub
