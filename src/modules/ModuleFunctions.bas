' Extracted from: ModuleFunctions.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "ModuleFunctions"
Private Declare PtrSafe Function WindowsColorDialog Lib "comdlg32.dll" Alias "ChooseColorA" (pcc As CHOOSECOLOR_TYPE) As LongPtr
    
Private Type CHOOSECOLOR_TYPE
    lStructSize As LongPtr
    hwndOwner   As LongPtr
    hInstance   As LongPtr
    rgbResult   As LongPtr
    lpCustColors As LongPtr
    flags       As LongPtr
    lCustData   As LongPtr
    lpfnHook    As LongPtr
    lpTemplateName As String
End Type

Private Const CC_ANYCOLOR = &H100
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80
Sub SetProgress(PercentageCompleted As Single)

    ProgressForm.ProgressBar.Width = PercentageCompleted * 2
    ProgressForm.ProgressLabel.Caption = Round(PercentageCompleted, 0) & "% completed"
    DoEvents
    
End Sub
Function ColorDialog(StandardColor As Variant) As Variant
    
Dim ChooseColorType As CHOOSECOLOR_TYPE
Dim ReturnColor As Variant

Static PredefinedColors(16)  As Long
     
If ActivePresentation.ExtraColors.Count > 0 Then
    For ExtraColorCount = 1 To ActivePresentation.ExtraColors.Count
        PredefinedColors(ExtraColorCount - 1) = ActivePresentation.ExtraColors(ExtraColorCount)
    Next
End If

With ActivePresentation.SlideMaster.Theme
    PredefinedColors(10) = .ThemeColorScheme(msoThemeAccent1).RGB
    PredefinedColors(11) = .ThemeColorScheme(msoThemeAccent2).RGB
    PredefinedColors(12) = .ThemeColorScheme(msoThemeAccent3).RGB
    PredefinedColors(13) = .ThemeColorScheme(msoThemeAccent4).RGB
    PredefinedColors(14) = .ThemeColorScheme(msoThemeAccent5).RGB
    PredefinedColors(15) = .ThemeColorScheme(msoThemeAccent6).RGB
End With

With ChooseColorType
    .lStructSize = Len(ChooseColorType)
    .flags = CC_RGBINIT Or CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN
    .rgbResult = StandardColor
    .lpCustColors = VarPtr(PredefinedColors(0))
End With

ReturnColor = WindowsColorDialog(ChooseColorType)

If Not ReturnColor = 0 Then
    ColorDialog = ChooseColorType.rgbResult
Else
    ColorDialog = StandardColor
End If
    
End Function
