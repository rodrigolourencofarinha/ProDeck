Attribute VB_Name = "RibbonImages"
Option Explicit

Public Sub Ribbon_GetImage(ByVal controlId As String, ByRef returnedVal)

#If Mac Then

    Select Case controlId
        Case "Margin_Toggle"
            returnedVal = "TableCellCustomMarginsDialog"
        Case "Margin_Toggle2"
            returnedVal = "TableCellCustomMarginsDialog"
        Case "Select_Objects"
            returnedVal = "SelectObjects"
        Case "Stack_Objects"
            returnedVal = "TraceDependents"
        Case "Swap_Center"
            returnedVal = "CircularReferences"
        Case "Objects_Corner"
            returnedVal = "DrawingCanvasExpand"
        Case "Merge_Text_Boxes"
            returnedVal = "MergeCenter"
        Case "Split_Text_Boxes"
            returnedVal = "MergeOrSplitCells"
        Case "CleanText"
            returnedVal = "PointEraserMedium"
        Case "About"
            returnedVal = "Help"
    End Select

#Else

    Select Case controlId
        Case "Margin_Toggle"
            returnedVal = "TextBoxMargins"
        Case "Margin_Toggle2"
            returnedVal = "TextBoxMargins"
        Case "Select_Objects"
            returnedVal = "SelectedEngagementGoTo"
        Case "Stack_Objects"
            returnedVal = "SnapToAlignmentBox"
        Case "Swap_Center"
            returnedVal = "MailMergeMatchFields"
        Case "Objects_Corner"
            returnedVal = "CornerRounding"
        Case "Merge_Text_Boxes"
            returnedVal = "ContentControlsGroup"
        Case "Split_Text_Boxes"
            returnedVal = "ContentControlsUngroup"
        Case "CleanText"
            returnedVal = "GroupRemoveHiddenInformation"
        Case "About"
            returnedVal = "ResultsPaneAccessibilityMoreInfo"
    End Select

#End If

End Sub
