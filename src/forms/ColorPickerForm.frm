Attribute VB_Name = "ColorPickerForm"
Attribute VB_Base = "0{115FAA20-D722-4A7E-A871-33C7E4298C59}{40192C3E-139D-41A7-A518-6E2B443D546F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'*************************************************************
' Attributes
Private SelectedColor As ColorPickerUtils.PickColor
Private InitColor As ColorPickerUtils.PickColor
Private MyStandardColors As New Collection
Private IsUpdatingUI As Boolean

'*************************************************************
' Initilialize

Public Sub resetForm()
    SelectedColor.red = 0
    SelectedColor.green = 0
    SelectedColor.blue = 0
    updateColor
    setStandardColors
    setThemeColors
End Sub


'*************************************************************
' Public functions

Public Function GetSelectedColor() As Long
    If SelectedColor.red = -1 Then
        GetSelectedColor = -1
    Else
        GetSelectedColor = RGB(SelectedColor.red, SelectedColor.green, SelectedColor.blue)
    End If
End Function

Public Sub SetSelectedColor(ByVal color As Long)
    SelectedColor = GetRGBFromLong(color)
    updateColor
End Sub

Public Sub SetInitColor(ByVal color As Long)
    SelectedColor = GetRGBFromLong(color)
    InitColor = GetRGBFromLong(color)
    updateColor
End Sub

Private Sub GreenBarOld_Change()

End Sub

'*************************************************************
' TextBox functions

Private Sub RedBox_Change()
    RedBox.text = setColor(RedBox.text, SelectedColor.red)
    updateColor
End Sub

Private Sub GreenBox_Change()
    GreenBox.text = setColor(GreenBox.text, SelectedColor.green)
    updateColor
End Sub

Private Sub BlueBox_Change()
    BlueBox.text = setColor(BlueBox.text, SelectedColor.blue)
    updateColor
End Sub

Private Sub LongBox_Change()
    On Error Resume Next
        SelectedColor = GetRGBFromLong(LongBox.text)
        updateColor
    On Error GoTo 0
End Sub

Private Sub HexBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then HexBox_Change
End Sub

Private Sub HexBox_Change()
    If IsUpdatingUI Then Exit Sub

    Dim s As String
    s = UCase$(Trim$(HexBox.value))

    ' Allow user to type "#FFAABB"
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)

    ' Only parse when complete
    If Len(s) <> 6 Then Exit Sub
    If Not s Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]" Then Exit Sub

    On Error Resume Next
        SelectedColor = GetRGBFromHex(s)
        updateColor
    On Error GoTo 0
End Sub

'*************************************************************
' Scrollbar functions

Private Sub RedBar_Change()
    SelectedColor.red = RedBar.value
    updateColor
End Sub

Private Sub GreenBar_Change()
    SelectedColor.green = GreenBar.value
    updateColor
End Sub

Private Sub BlueBar_Change()
    SelectedColor.blue = BlueBar.value
    updateColor
End Sub

'Private Sub RedBar_Scroll()
'    SelectedColor.red = RedBar.value
'    updateColor
'End Sub

'Private Sub GreenBar_Scroll()
'    SelectedColor.green = GreenBar.value
'    updateColor
'End Sub

'Private Sub BlueBar_Scroll()
'    SelectedColor.blue = BlueBar.value
'    updateColor
'End Sub


'*************************************************************
' Button functions

Private Sub OKButton_Click()
    continue = True
    ColorPickerForm.Hide
End Sub

Private Sub CancelButton_Click()
    continue = False
    ColorPickerForm.Hide
    Exit Sub
End Sub
Private Sub CommandButtonEsc_Click()
    continue = False
    ColorPickerForm.Hide
    Exit Sub
End Sub


Private Sub ScrollBar1_Change()
    SelectedColor.green = GreenBar.value
    updateColor
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then
    Cancel = True
    continue = False
    ColorPickerForm.Hide
End If
End Sub



'*************************************************************
' Helper functions

' set the color label background color
Private Sub updateColor()
    IsUpdatingUI = True

    With SelectedColor
        ColorLabel.BackColor = RGB(.red, .green, .blue)
        LongBox.value = RGB(.red, .green, .blue)

        If Not (Me.ActiveControl Is HexBox) Then
            HexBox.value = GetHexFromRGB(SelectedColor)
        End If

        RedBox.value = .red
        RedBar.value = .red
        GreenBox.value = .green
        GreenBar.value = .green
        BlueBox.value = .blue
        BlueBar.value = .blue
    End With

    IsUpdatingUI = False
End Sub

' set the color to the value parsed from text, with limits of
' 0-255
Private Function setColor(ByRef text As String, ByRef color As Integer) As Integer
    On Error Resume Next
        If text = "" Then
            color = 0
        Else
            color = CInt(text)
            If color < 0 Then
                color = 0
            ElseIf color > 255 Then
                color = 255
            End If
        End If
    On Error GoTo 0
    setColor = color
End Function


'*************************************************************
' Color Array Functions

' set the theme color boxes
Private Sub setThemeColors()
    
    ' Create file path string
    Dim curSlide As slide
    Set curSlide = SlideUtils.CurrentSlide("")
    If curSlide Is Nothing Then
        ThemeColor1.BackColor = MyStandardColors(1)
        ThemeColor2.BackColor = MyStandardColors(2)
        ThemeColor3.BackColor = MyStandardColors(3)
        ThemeColor4.BackColor = MyStandardColors(4)
        ThemeColor5.BackColor = MyStandardColors(5)
        ThemeColor6.BackColor = MyStandardColors(6)
        ThemeColor7.BackColor = MyStandardColors(7)
        ThemeColor8.BackColor = MyStandardColors(8)
        ThemeColor9.BackColor = MyStandardColors(9)
        ThemeColor10.BackColor = MyStandardColors(10)
        ThemeColor11.BackColor = MyStandardColors(11)
        ThemeColor12.BackColor = MyStandardColors(12)
    Else
        With ActiveWindow.View.slide.ThemeColorScheme
            ThemeColor1.BackColor = .Colors(1)
            ThemeColor2.BackColor = .Colors(2)
            ThemeColor3.BackColor = .Colors(3)
            ThemeColor4.BackColor = .Colors(4)
            ThemeColor5.BackColor = .Colors(5)
            ThemeColor6.BackColor = .Colors(6)
            ThemeColor7.BackColor = .Colors(7)
            ThemeColor8.BackColor = .Colors(8)
            ThemeColor9.BackColor = .Colors(9)
            ThemeColor10.BackColor = .Colors(10)
            ThemeColor11.BackColor = .Colors(11)
            ThemeColor12.BackColor = .Colors(12)
        End With
    End If
    


End Sub

' set the theme color boxes
Private Sub setStandardColors()
    
        MyStandardColors.Add RGB(128, 0, 0)
        MyStandardColors.Add RGB(255, 0, 0)
        MyStandardColors.Add RGB(255, 128, 0)
        MyStandardColors.Add RGB(255, 255, 0)
        MyStandardColors.Add RGB(0, 128, 0)
        MyStandardColors.Add RGB(0, 255, 0)
        MyStandardColors.Add RGB(0, 0, 128)
        MyStandardColors.Add RGB(0, 0, 255)
        MyStandardColors.Add RGB(0, 255, 255)
        MyStandardColors.Add RGB(255, 0, 255)
        MyStandardColors.Add RGB(100, 100, 100)
        MyStandardColors.Add RGB(200, 200, 200)
    
        StandardColor1.BackColor = MyStandardColors(1)
        StandardColor2.BackColor = MyStandardColors(2)
        StandardColor3.BackColor = MyStandardColors(3)
        StandardColor4.BackColor = MyStandardColors(4)
        StandardColor5.BackColor = MyStandardColors(5)
        StandardColor6.BackColor = MyStandardColors(6)
        StandardColor7.BackColor = MyStandardColors(7)
        StandardColor8.BackColor = MyStandardColors(8)
        StandardColor9.BackColor = MyStandardColors(9)
        StandardColor10.BackColor = MyStandardColors(10)
        StandardColor11.BackColor = MyStandardColors(11)
        StandardColor12.BackColor = MyStandardColors(12)
End Sub

Private Sub setColorFromTheme(ByVal ind As MsoThemeColorSchemeIndex)
    Dim curSlide As slide
    Set curSlide = SlideUtils.CurrentSlide("")
    If curSlide Is Nothing Then
        ind = ((ind - 1) Mod MyStandardColors.Count) + 1
        SelectedColor = GetRGBFromLong(MyStandardColors(ind))
    Else
        SelectedColor = GetRGBFromLong(curSlide.ThemeColorScheme.Colors(ind))
    End If
    updateColor
End Sub

Private Sub setColorFromStandard(ByVal ind As Integer)
    SelectedColor = GetRGBFromLong(MyStandardColors(ind))
    updateColor
End Sub


Private Sub ThemeColor1_Click()
    setColorFromTheme 1
End Sub

Private Sub ThemeColor2_Click()
    setColorFromTheme 2
End Sub

Private Sub ThemeColor3_Click()
    setColorFromTheme 3
End Sub

Private Sub ThemeColor4_Click()
    setColorFromTheme 4
End Sub

Private Sub ThemeColor5_Click()
    setColorFromTheme 5
End Sub

Private Sub ThemeColor6_Click()
    setColorFromTheme 6
End Sub

Private Sub ThemeColor7_Click()
    setColorFromTheme 7
End Sub

Private Sub ThemeColor8_Click()
    setColorFromTheme 8
End Sub

Private Sub ThemeColor9_Click()
    setColorFromTheme 9
End Sub

Private Sub ThemeColor10_Click()
    setColorFromTheme 10
End Sub

Private Sub ThemeColor11_Click()
    setColorFromTheme 11
End Sub

Private Sub ThemeColor12_Click()
    setColorFromTheme 12
End Sub

Private Sub StandardColor1_Click()
    setColorFromStandard 1
End Sub

Private Sub StandardColor2_Click()
    setColorFromStandard 2
End Sub

Private Sub StandardColor3_Click()
    setColorFromStandard 3
End Sub

Private Sub StandardColor4_Click()
    setColorFromStandard 4
End Sub

Private Sub StandardColor5_Click()
    setColorFromStandard 5
End Sub

Private Sub StandardColor6_Click()
    setColorFromStandard 6
End Sub

Private Sub StandardColor7_Click()
    setColorFromStandard 7
End Sub

Private Sub StandardColor8_Click()
    setColorFromStandard 8
End Sub

Private Sub StandardColor9_Click()
    setColorFromStandard 9
End Sub

Private Sub StandardColor10_Click()
    setColorFromStandard 10
End Sub

Private Sub StandardColor11_Click()
    setColorFromStandard 11
End Sub

Private Sub StandardColor12_Click()
    setColorFromStandard 12
End Sub

