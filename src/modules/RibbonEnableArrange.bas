Attribute VB_Name = "RibbonEnableArrange"
Option Explicit

Public Function RibbonEnableArrange_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableArrange_IsEnabled = True

    Select Case controlId
        Case "Increase_Horizontal_Spacing"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Decrease_Horizontal_Spacing"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Increase_Vertical_Spacing"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Decrease_Vertical_Spacing"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Swap_Positions"
            returnedVal = Ribbon_SelectedShapeOrChildCountEquals(2)
        Case "Swap_Text"
            returnedVal = Ribbon_CanSwapTextSelection()
        Case "Stack_Top"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Stack_Left"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Stack_Bottom"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Stack_Right"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Objects_Corner"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case Else
            RibbonEnableArrange_IsEnabled = False
    End Select
End Function
