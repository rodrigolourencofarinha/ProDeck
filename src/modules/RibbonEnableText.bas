Attribute VB_Name = "RibbonEnableText"
Option Explicit

Public Function RibbonEnableText_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableText_IsEnabled = True

    Select Case controlId
        Case "AdditionalBullets"
            returnedVal = Ribbon_CanUseTextSelection(False, False, True)
        Case "Bullet_Point_1"
            returnedVal = Ribbon_CanUseTextSelection(False, False, True)
        Case "Bullet_Point_2"
            returnedVal = Ribbon_CanUseTextSelection(False, False, True)
        Case "Bullet_Point_3"
            returnedVal = Ribbon_CanUseTextSelection(False, False, True)
        Case "Margin_Toggle"
            returnedVal = Ribbon_CanUseTextSelection(False, True, False)
        Case "Margin_Toggle2"
            returnedVal = Ribbon_CanUseTextSelection(True, True, False)
        Case "Wrap_Text"
            returnedVal = Ribbon_CanUseTextSelection(True, True, False)
        Case "Do_Not_Autofit"
            returnedVal = Ribbon_CanUseTextSelection(True, True, False)
        Case "Shape_Fit_Text"
            returnedVal = Ribbon_CanUseTextSelection(True, True, False)
        Case "Text_Overflow"
            returnedVal = Ribbon_CanUseTextSelection(True, True, False)
        Case Else
            RibbonEnableText_IsEnabled = False
    End Select
End Function
