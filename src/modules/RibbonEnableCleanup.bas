Attribute VB_Name = "RibbonEnableCleanup"
Option Explicit

Public Function RibbonEnableCleanup_IsEnabled(ByVal controlId As String, ByRef returnedVal As Boolean) As Boolean
    RibbonEnableCleanup_IsEnabled = True

    Select Case controlId
        Case "Remove_Line_Break"
            returnedVal = Ribbon_CanRemoveLineBreakSelection()
        Case "Remove_Spaces"
            returnedVal = Ribbon_CanUseCurrentSlide()
        Case "CleanDeck"
            returnedVal = Ribbon_CanUseCurrentView()
        Case "StampD_Update"
            returnedVal = Ribbon_CanUpdateDocumentStamp()
        Case "StampsMenu"
            returnedVal = Ribbon_CanUseStampsMenu()
        Case Else
            RibbonEnableCleanup_IsEnabled = False
    End Select
End Function

Private Function Ribbon_CanUpdateDocumentStamp() As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.ShapeRange.Count = 1 Then
            If ActiveWindow.Selection.ShapeRange(1).Name = "PRODECK DOCUMENT STAMP" Then
                Ribbon_CanUpdateDocumentStamp = True
            Else
                Ribbon_CanUpdateDocumentStamp = False
            End If
        Else
            Ribbon_CanUpdateDocumentStamp = False
        End If
    Else
        Ribbon_CanUpdateDocumentStamp = False
    End If
End Function

Private Function Ribbon_CanUseStampsMenu() As Boolean
    On Error Resume Next

    If ActiveWindow.Application.Visible = msoFalse Then
        Ribbon_CanUseStampsMenu = False
    ElseIf ActiveWindow.Presentation.Slides.Count = 0 Then
        Ribbon_CanUseStampsMenu = False
    ElseIf Ribbon_IsBlockedView() Then
        Ribbon_CanUseStampsMenu = False
    ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
        Ribbon_CanUseStampsMenu = False
    Else
        Ribbon_CanUseStampsMenu = True
    End If
End Function
