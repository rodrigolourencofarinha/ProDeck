Attribute VB_Name = "RibbonEnableInsert"
Option Explicit

Public Function RibbonEnableInsert_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableInsert_IsEnabled = True

    Select Case controlId
        Case "Delete_Similar_Objects"
            returnedVal = Ribbon_CanUseSingleShapeInEditView()
        Case "Update_Objects_Position"
            returnedVal = Ribbon_CanUseSingleShapeInEditView()
        Case "Grid_Shapes"
            returnedVal = Ribbon_CanUseGridShapeSelection()
        Case "Insert_Sticky_Note"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case "TextAndOthers"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case "Paste_Objects_All_Slides"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case Else
            RibbonEnableInsert_IsEnabled = False
    End Select
End Function
