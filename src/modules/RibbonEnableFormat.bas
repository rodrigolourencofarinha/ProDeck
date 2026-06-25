Attribute VB_Name = "RibbonEnableFormat"
Option Explicit

Public Function RibbonEnableFormat_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableFormat_IsEnabled = True

    Select Case controlId
        Case "Select_Objects"
            returnedVal = Ribbon_CanUseSingleShapeOrChildInEditView()
        Case "ConnectRectanglesMenu"
            returnedVal = Ribbon_SelectedShapeOrChildCountEquals(2)
        Case "CloneSelectionRight"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(1)
        Case "CloneSelectionDown"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(1)
        Case "TableConfig"
            returnedVal = Ribbon_CanUseTableSelection()
        Case "Replace_Colors"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case Else
            RibbonEnableFormat_IsEnabled = False
    End Select
End Function
