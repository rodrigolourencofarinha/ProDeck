Attribute VB_Name = "RibbonSelectionGuards"
Option Explicit

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
