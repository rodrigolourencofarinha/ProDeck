Attribute VB_Name = "EnableButtons"
Option Explicit
Public Rib As IRibbonUI
Public xmlID As String
Private Const AUTO_TAB_APP_NAME As String = "ProDeck"
Private Const AUTO_TAB_SECTION As String = "Preferences"
Private Const AUTO_TAB_KEY As String = "AutoTabMonitoringEnabled"
Private Const AUTO_TAB_DEFAULT As String = "True"
Private Const PRODECK_TAB_ID As String = "ProDeck"
Private Const PRODECK_TAB_KEYTIP As String = "%X"
Private mAutoTabMonitoringEnabled As Boolean
Private mAutoTabMonitoringLoaded As Boolean
Private mAutoTabRecoveryNoticeShown As Boolean
Private mAutoTabEvent As EventClassModule
'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)

    TrapEvents  'instantiate the event handler

    Set Rib = ribbon

End Sub

Sub EnabledBtInfo(control As IRibbonControl, ByRef returnedVal)

    returnedVal = Ribbon_IsEnabled(control.Id)

    Call RefreshRibbon(control.Id)

End Sub
Public Sub IsMac(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = True
#Else
    returnedVal = False
#End If
End Sub
Public Sub IsWindows(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub
Sub RibbonObjectGetImage(control As IRibbonControl, ByRef returnedVal)

    Ribbon_GetImage control.Id, returnedVal

End Sub
Sub RefreshRibbon(Id As String)

xmlID = Id

If Rib Is Nothing Then
Else
    Rib.Invalidate
End If

End Sub

' ===== Flattened from RibbonEnableArrange.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonEnableCleanup.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonEnableFormat.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonEnableInsert.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonEnableSizeOrganize.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonEnableState.bas for Windows-compatible PPTM packaging =====

Public Function Ribbon_IsEnabled(ByVal controlId As String) As Boolean
    Dim returnedVal As Boolean

    If RibbonEnableInsert_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableText_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableFormat_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableArrange_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableSizeOrganize_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableCleanup_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    Ribbon_IsEnabled = False
End Function

' ===== Flattened from RibbonEnableText.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonImages.bas for Windows-compatible PPTM packaging =====

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

' ===== Flattened from RibbonSelectionGuards.bas for Windows-compatible PPTM packaging =====

Public Function Ribbon_IsBlockedView() As Boolean
    On Error Resume Next

    Select Case ActiveWindow.ActivePane.ViewType
        Case ppViewHandoutMaster, ppViewMasterThumbnails, ppViewNotesMaster, ppViewNotesPage, _
             ppViewPrintPreview, ppViewSlideMaster, ppViewSlideSorter, ppViewTitleMaster
            Ribbon_IsBlockedView = True
        Case Else
            Ribbon_IsBlockedView = False
    End Select
End Function

Public Function Ribbon_CanUseCurrentView() As Boolean
    Ribbon_CanUseCurrentView = Not Ribbon_IsBlockedView()
End Function

Public Function Ribbon_HasSlideSelection() As Boolean
    On Error Resume Next

    Ribbon_HasSlideSelection = (ActiveWindow.Selection.SlideRange.Count <> 0)
End Function

Public Function Ribbon_CanUseCurrentSlide() As Boolean
    If Ribbon_IsBlockedView() Then
        Ribbon_CanUseCurrentSlide = False
    Else
        Ribbon_CanUseCurrentSlide = Ribbon_HasSlideSelection()
    End If
End Function

Public Function Ribbon_CanUseSingleShapeInEditView() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then Exit Function

    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        Ribbon_CanUseSingleShapeInEditView = ActiveWindow.Selection.Type = ppSelectionShapes
    Else
        Ribbon_CanUseSingleShapeInEditView = False
    End If
End Function

Public Function Ribbon_CanUseSingleShapeOrChildInEditView() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then Exit Function

    If ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.Count = 1 Then
            Ribbon_CanUseSingleShapeOrChildInEditView = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_CanUseSingleShapeOrChildInEditView = False
        End If
    Else
        If ActiveWindow.Selection.ShapeRange.Count = 1 Then
            Ribbon_CanUseSingleShapeOrChildInEditView = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_CanUseSingleShapeOrChildInEditView = False
        End If
    End If
End Function

Public Function Ribbon_CanUseShapeSelectionInEditView() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then
        Ribbon_CanUseShapeSelectionInEditView = False
    Else
        Ribbon_CanUseShapeSelectionInEditView = ActiveWindow.Selection.Type = ppSelectionShapes
    End If
End Function

Public Function Ribbon_SelectedShapeOrChildCountAtLeast(ByVal minimumCount As Long) As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.Count >= minimumCount Then
            Ribbon_SelectedShapeOrChildCountAtLeast = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_SelectedShapeOrChildCountAtLeast = False
        End If
    Else
        If ActiveWindow.Selection.ShapeRange.Count >= minimumCount Then
            Ribbon_SelectedShapeOrChildCountAtLeast = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_SelectedShapeOrChildCountAtLeast = False
        End If
    End If
End Function

Public Function Ribbon_SelectedShapeOrChildCountEquals(ByVal exactCount As Long) As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.Count = exactCount Then
            Ribbon_SelectedShapeOrChildCountEquals = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_SelectedShapeOrChildCountEquals = False
        End If
    Else
        If ActiveWindow.Selection.ShapeRange.Count = exactCount Then
            Ribbon_SelectedShapeOrChildCountEquals = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_SelectedShapeOrChildCountEquals = False
        End If
    End If
End Function

Public Function Ribbon_CanUseTextSelection(ByVal excludeShapeTable As Boolean, ByVal excludeTextTable As Boolean, ByVal excludeSmartArt As Boolean) As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.Type = ppSelectionNone Then
        Ribbon_CanUseTextSelection = False
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Ribbon_CanUseTextSelection = Not Ribbon_IsBlockedTextShapeType(excludeShapeTable, excludeSmartArt)
    ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
        Ribbon_CanUseTextSelection = Not Ribbon_IsBlockedTextShapeType(excludeTextTable, excludeSmartArt)
    Else
        Ribbon_CanUseTextSelection = False
    End If
End Function

Public Function Ribbon_CanUseGridShapeSelection() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then Exit Function

    If ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Fill.Visible = msoTrue Then
        If Ribbon_IsBlockedGridShapeType() Then
            Ribbon_CanUseGridShapeSelection = False
        Else
            Ribbon_CanUseGridShapeSelection = True
        End If
    Else
        Ribbon_CanUseGridShapeSelection = False
    End If
End Function

Public Function Ribbon_CanSwapTextSelection() As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
            If ActiveWindow.Selection.ChildShapeRange(1).HasTextFrame And ActiveWindow.Selection.ChildShapeRange(2).HasTextFrame Then
                Ribbon_CanSwapTextSelection = True
            Else
                Ribbon_CanSwapTextSelection = False
            End If
        Else
            Ribbon_CanSwapTextSelection = False
        End If
    Else
        If ActiveWindow.Selection.ShapeRange.Count = 2 Then
            If ActiveWindow.Selection.ShapeRange(1).HasTextFrame And ActiveWindow.Selection.ShapeRange(2).HasTextFrame Then
                Ribbon_CanSwapTextSelection = True
            Else
                Ribbon_CanSwapTextSelection = False
            End If
        Else
            Ribbon_CanSwapTextSelection = False
        End If
    End If
End Function

Public Function Ribbon_CanUseTableSelection() As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        If ActiveWindow.Selection.ShapeRange.HasTable Then
            Ribbon_CanUseTableSelection = ActiveWindow.Selection.Type = ppSelectionShapes
        Else
            Ribbon_CanUseTableSelection = False
        End If
    Else
        Ribbon_CanUseTableSelection = False
    End If
End Function

Public Function Ribbon_CanUseRectifyLineSelection() As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.HasChildShapeRange Then
            If ActiveWindow.Selection.ChildShapeRange.Count > 0 Then
                If ActiveWindow.Selection.ChildShapeRange.Connector = msoTrue Then
                    Ribbon_CanUseRectifyLineSelection = True
                Else
                    Ribbon_CanUseRectifyLineSelection = False
                End If
            Else
                Ribbon_CanUseRectifyLineSelection = False
            End If
        Else
            If ActiveWindow.Selection.ShapeRange.Connector = msoTrue Then
                Ribbon_CanUseRectifyLineSelection = True
            Else
                Ribbon_CanUseRectifyLineSelection = False
            End If
        End If
    Else
        Ribbon_CanUseRectifyLineSelection = False
    End If
End Function

Public Function Ribbon_CanMergeTextSelection() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then Exit Function

    If ActiveWindow.Selection.ShapeRange.Count >= 2 And ActiveWindow.Selection.ShapeRange.HasTextFrame Then
        Ribbon_CanMergeTextSelection = ActiveWindow.Selection.Type = ppSelectionShapes
    Else
        Ribbon_CanMergeTextSelection = False
    End If
End Function

Public Function Ribbon_CanSplitTextSelection() As Boolean
    On Error Resume Next

    If Ribbon_IsBlockedView() Then Exit Function

    If ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.HasTextFrame Then
        Ribbon_CanSplitTextSelection = ActiveWindow.Selection.Type = ppSelectionShapes
    Else
        Ribbon_CanSplitTextSelection = False
    End If
End Function

Public Function Ribbon_CanRemoveLineBreakSelection() As Boolean
    On Error Resume Next

    If ActiveWindow.Selection.Type = ppSelectionNone Then
        Ribbon_CanRemoveLineBreakSelection = False
    ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
        Ribbon_CanRemoveLineBreakSelection = True
    ElseIf ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.HasTextFrame Then
            If ActiveWindow.Selection.ChildShapeRange.Count >= 1 Then
                Ribbon_CanRemoveLineBreakSelection = ActiveWindow.Selection.Type = ppSelectionShapes
            Else
                Ribbon_CanRemoveLineBreakSelection = False
            End If
        Else
            Ribbon_CanRemoveLineBreakSelection = False
        End If
    Else
        If ActiveWindow.Selection.ShapeRange.HasTextFrame Then
            If ActiveWindow.Selection.ShapeRange.Count >= 1 Then
                Ribbon_CanRemoveLineBreakSelection = ActiveWindow.Selection.Type = ppSelectionShapes
            Else
                Ribbon_CanRemoveLineBreakSelection = False
            End If
        Else
            Ribbon_CanRemoveLineBreakSelection = False
        End If
    End If
End Function

Private Function Ribbon_IsBlockedTextShapeType(ByVal excludeTable As Boolean, ByVal excludeSmartArt As Boolean) As Boolean
    On Error Resume Next

    Select Case ActiveWindow.Selection.ShapeRange.Type
        Case msoChart, msoCallout, msoCanvas, msoComment, msoDiagram, msoEmbeddedOLEObject, _
             msoFormControl, msoLine, msoLinkedOLEObject, msoLinkedPicture, msoMedia, _
             msoOLEControlObject, msoPicture, msoScriptAnchor, mso3DModel, msoContentApp, _
             msoGraphic, msoInk, msoLinked3DModel, msoLinkedGraphic, msoSlicer, _
             msoInkComment, msoWebVideo
            Ribbon_IsBlockedTextShapeType = True
        Case msoTable
            Ribbon_IsBlockedTextShapeType = excludeTable
        Case msoSmartArt
            Ribbon_IsBlockedTextShapeType = excludeSmartArt
        Case Else
            Ribbon_IsBlockedTextShapeType = False
    End Select
End Function

Private Function Ribbon_IsBlockedGridShapeType() As Boolean
    On Error Resume Next

    Select Case ActiveWindow.Selection.ShapeRange.Type
        Case msoChart, msoCallout, msoCanvas, msoComment, msoDiagram, msoEmbeddedOLEObject, _
             msoFormControl, msoLine, msoLinkedOLEObject, msoLinkedPicture, msoMedia, _
             msoOLEControlObject, msoPicture, msoScriptAnchor, mso3DModel, msoContentApp, _
             msoGraphic, msoInk, msoLinked3DModel, msoLinkedGraphic, msoSlicer, _
             msoInkComment, msoSmartArt, msoGroup, msoTable, msoWebVideo
            Ribbon_IsBlockedGridShapeType = True
        Case Else
            Ribbon_IsBlockedGridShapeType = False
    End Select
End Function

' ===== Auto-tab recovery implementation flattened for Windows-compatible PPTM packaging =====

Public Sub AutoTab_InvalRibbon()
    On Error Resume Next
    If MyRibbon Is Nothing Then Set MyRibbon = Rib
    If Rib Is Nothing Then Set Rib = MyRibbon
    If MyRibbon Is Nothing Then Exit Sub
    Err.Clear
    MyRibbon.Invalidate
    If Err.Number <> 0 Then
        Err.Clear
        Set MyRibbon = Nothing
        If Not Rib Is Nothing Then
            Rib.Invalidate
            If Err.Number = 0 Then
                Set MyRibbon = Rib
            Else
                Err.Clear
                Set Rib = Nothing
            End If
        End If
    End If
End Sub

Public Sub AutoTab_ArmEventHandler()
    On Error Resume Next
    If mAutoTabEvent Is Nothing Then Set mAutoTabEvent = New EventClassModule
    Set mAutoTabEvent.App = Application
    If Err.Number <> 0 Then
        Err.Clear
        Set mAutoTabEvent = Nothing
        Set mAutoTabEvent = New EventClassModule
        Set mAutoTabEvent.App = Application
    End If
End Sub

Public Sub AutoTab_Activate()
    AutoTab_InvalRibbon
    If Not AutoTab_TryActivate(True) Then AutoTab_ShowRecoveryMessage
End Sub

Public Sub AutoTab_MaybeActivate()
    If AutoTab_IsEnabled() Then
        AutoTab_ArmEventHandler
        Call AutoTab_TryActivate(True)
    End If
End Sub

Public Sub AutoTab_Initialize()
    mAutoTabMonitoringLoaded = False
    AutoTab_LoadSetting
End Sub

Public Function AutoTab_IsEnabled() As Boolean
    AutoTab_LoadSetting
    AutoTab_IsEnabled = mAutoTabMonitoringEnabled
End Function

Public Sub AutoTab_Toggle(ByVal pressed As Boolean)
    Dim promptText As String
    promptText = IIf(pressed, "Enable", "Disable") & " automatic return to the ProDeck tab when selecting shapes or text?"
    If MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton1, "ProDeck") = vbYes Then
        mAutoTabMonitoringEnabled = pressed
        mAutoTabMonitoringLoaded = True
        AutoTab_SaveSetting pressed
        If pressed Then
            AutoTab_ArmEventHandler
            mAutoTabRecoveryNoticeShown = False
            AutoTab_InvalRibbon
            If Not AutoTab_TryActivate(True) Then AutoTab_ShowRecoveryMessage
        End If
    End If
    AutoTab_InvalRibbon
End Sub

Public Sub AutoTab_Repair()
    mAutoTabMonitoringEnabled = True
    mAutoTabMonitoringLoaded = True
    mAutoTabRecoveryNoticeShown = False
    AutoTab_SaveSetting True
    AutoTab_ArmEventHandler
    AutoTab_InvalRibbon
    If Not AutoTab_TryActivate(True) Then AutoTab_ShowRecoveryMessage
End Sub

Public Function AutoTab_TryActivate(Optional ByVal allowKeytipFallback As Boolean = True) As Boolean
    On Error Resume Next
    If MyRibbon Is Nothing Then Set MyRibbon = Rib
    If AutoTab_TryRibbon(MyRibbon) Then
        AutoTab_TryActivate = True
        Exit Function
    End If
    If AutoTab_TryRibbon(Rib) Then
        Set MyRibbon = Rib
        AutoTab_TryActivate = True
        Exit Function
    End If
    Set MyRibbon = Nothing
    Set Rib = Nothing
    If allowKeytipFallback Then AutoTab_TryActivate = AutoTab_TryKeytip()
End Function

Private Function AutoTab_TryRibbon(ByVal ribbon As IRibbonUI) As Boolean
    On Error GoTo ActivationFailed
    If ribbon Is Nothing Then Exit Function
    ribbon.ActivateTab PRODECK_TAB_ID
    AutoTab_TryRibbon = True
    Exit Function
ActivationFailed:
    Err.Clear
End Function

Private Function AutoTab_TryKeytip() As Boolean
#If Mac Then
    AutoTab_TryKeytip = False
#Else
    On Error Resume Next
    Err.Clear
    SendKeys PRODECK_TAB_KEYTIP
    AutoTab_TryKeytip = (Err.Number = 0)
    Err.Clear
#End If
End Function

Private Sub AutoTab_ShowRecoveryMessage()
    If mAutoTabRecoveryNoticeShown Then Exit Sub
    mAutoTabRecoveryNoticeShown = True
    MsgBox "PowerPoint lost the live ProDeck ribbon connection. ProDeck re-armed selection monitoring, but PowerPoint may need the add-in to be reloaded before automatic tab activation can resume.", vbExclamation + vbOKOnly, "ProDeck"
End Sub

Private Sub AutoTab_LoadSetting()
    Dim savedValue As String
    If mAutoTabMonitoringLoaded Then Exit Sub
    On Error Resume Next
    savedValue = GetSetting(AUTO_TAB_APP_NAME, AUTO_TAB_SECTION, AUTO_TAB_KEY, AUTO_TAB_DEFAULT)
    If Err.Number <> 0 Then
        Err.Clear
        savedValue = AUTO_TAB_DEFAULT
    End If
    savedValue = LCase$(Trim$(savedValue))
    mAutoTabMonitoringEnabled = Not (savedValue = "false" Or savedValue = "0" Or savedValue = "no" Or savedValue = "off")
    mAutoTabMonitoringLoaded = True
End Sub

Private Sub AutoTab_SaveSetting(ByVal isEnabled As Boolean)
    On Error Resume Next
    SaveSetting AUTO_TAB_APP_NAME, AUTO_TAB_SECTION, AUTO_TAB_KEY, CStr(isEnabled)
    Err.Clear
End Sub
