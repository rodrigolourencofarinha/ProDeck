Attribute VB_Name = "RibbonEnableSizeOrganize"
Option Explicit

Public Function RibbonEnableSizeOrganize_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableSizeOrganize_IsEnabled = True

    Select Case controlId
        Case "RectifyLines"
            returnedVal = Ribbon_CanUseRectifyLineSelection()
        Case "Resize_Height"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Resize_Width"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "Resize_Shapes"
            returnedVal = Ribbon_SelectedShapeOrChildCountAtLeast(2)
        Case "LockUnlock"
            returnedVal = Ribbon_CanUseLockUnlock()
        Case "Hide_Objects"
            returnedVal = Ribbon_CanUseShapeSelectionInEditView()
        Case "Show_All_Objects"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case "Merge_Text_Boxes"
            returnedVal = Ribbon_CanMergeTextSelection()
        Case "Split_Text_Boxes"
            returnedVal = Ribbon_CanSplitTextSelection()
        Case Else
            RibbonEnableSizeOrganize_IsEnabled = False
    End Select
End Function

Private Function Ribbon_CanUseLockUnlock() As Boolean
    On Error Resume Next

    Ribbon_CanUseLockUnlock = ActiveWindow.Selection.Type = ppSelectionShapes
End Function
