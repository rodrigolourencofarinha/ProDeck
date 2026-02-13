Attribute VB_Name = "EnableButtons"
Option Explicit
Public Rib As IRibbonUI
Public xmlID As String
'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)

    TrapEvents  'instantiate the event handler
    
    Set Rib = ribbon
    
End Sub

Sub EnabledBtInfo(control As IRibbonControl, ByRef returnedVal)

    'Check the ActiveWindow.Selection.ShapeRange
    Select Case control.Id
    
        Case "Delete_Similar_Objects"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
                
        Case "Update_Objects_Position"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
                
        Case "Grid_Shapes"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Fill.Visible = msoTrue Then
                             If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGroup Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                             returnedVal = False
                             ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                             returnedVal = False
                             Else
                             returnedVal = True
                             End If
                         Else
                         returnedVal = False
                    End If
                End If
                
        Case "AdditionalBullets"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
            
        Case "Bullet_Point_1"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Bullet_Point_2"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Bullet_Point_3"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Margin_Toggle"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Margin_Toggle2"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Wrap_Text"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Do_Not_Autofit"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Shape_Fit_Text"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Text_Overflow"
            On Error Resume Next
            If ActiveWindow.Selection.Type = ppSelectionNone Then
            returnedVal = False
            ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                If ActiveWindow.Selection.ShapeRange.Type = msoChart Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCallout Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoCanvas Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoDiagram Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoEmbeddedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoFormControl Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLine Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedOLEObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoMedia Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoOLEControlObject Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoScriptAnchor Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = mso3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoContentApp Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInk Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinked3DModel Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoLinkedGraphic Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSlicer Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoInkComment Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoTable Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.ShapeRange.Type = msoWebVideo Then
                returnedVal = False
                Else
                returnedVal = True
                End If
            Else
            returnedVal = False
            End If
        
        Case "Select_Objects"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.HasChildShapeRange Then
                        If ActiveWindow.Selection.ChildShapeRange.Count = 1 Then
                        returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                        Else
                        returnedVal = False
                        End If
                    Else
                        If ActiveWindow.Selection.ShapeRange.Count = 1 Then
                        returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                        Else
                        returnedVal = False
                        End If
                    End If
                End If
        
        Case "Increase_Horizontal_Spacing"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Decrease_Horizontal_Spacing"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Increase_Vertical_Spacing"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Decrease_Vertical_Spacing"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Swap_Positions"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
                
        Case "Swap_Text"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
                        If ActiveWindow.Selection.ChildShapeRange(1).HasTextFrame And ActiveWindow.Selection.ChildShapeRange(2).HasTextFrame Then
                            returnedVal = True
                        Else
                            returnedVal = False
                        End If
                    Else
                        returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
                        If ActiveWindow.Selection.ShapeRange(1).HasTextFrame And ActiveWindow.Selection.ShapeRange(2).HasTextFrame Then
                            returnedVal = True
                        Else
                            returnedVal = False
                        End If
                    Else
                        returnedVal = False
                    End If
                End If
        
        Case "ConnectRectanglesMenu"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count = 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "RectifyLines"
                On Error Resume Next
                If ActiveWindow.Selection.Type = ppSelectionShapes Then
                    If ActiveWindow.Selection.HasChildShapeRange Then
                        If ActiveWindow.Selection.ChildShapeRange.Count > 0 Then
                            If ActiveWindow.Selection.ChildShapeRange.Connector = msoTrue Then
                                returnedVal = True
                            Else
                                returnedVal = False
                            End If
                        Else
                            returnedVal = False
                        End If
                    Else
                        If ActiveWindow.Selection.ShapeRange.Connector = msoTrue Then
                            returnedVal = True
                        Else
                            returnedVal = False
                        End If
                    End If
                Else
                    returnedVal = False
                End If
        
        Case "Stack_Top"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Stack_Left"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Stack_Bottom"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Stack_Right"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Objects_Corner"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        Case "CloneSelectionRight"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "CloneSelectionDown"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 1 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "TableConfig"
                On Error Resume Next
                If ActiveWindow.Selection.ShapeRange.Count = 1 Then
                    If ActiveWindow.Selection.ShapeRange.HasTable Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    returnedVal = False
                End If
        
        Case "Hide_Objects"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                End If
        
        Case "Resize_Height"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Resize_Width"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Resize_Shapes"
                On Error Resume Next
                If ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "LockUnlock"
            On Error Resume Next
            returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
        
        Case "Merge_Text_Boxes"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.ShapeRange.Count >= 2 And ActiveWindow.Selection.ShapeRange.HasTextFrame Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
                
                
        
        Case "Split_Text_Boxes"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                    If ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.HasTextFrame Then
                    returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Remove_Line_Break"
                On Error Resume Next
                If ActiveWindow.Selection.Type = ppSelectionNone Then
                    returnedVal = False
                ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
                    returnedVal = True
                ElseIf ActiveWindow.Selection.HasChildShapeRange Then
                    If ActiveWindow.Selection.ChildShapeRange.HasTextFrame Then
                        If ActiveWindow.Selection.ChildShapeRange.Count >= 1 Then
                        returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                        Else
                        returnedVal = False
                        End If
                    Else
                    returnedVal = False
                    End If
                Else
                    If ActiveWindow.Selection.ShapeRange.HasTextFrame Then
                        If ActiveWindow.Selection.ShapeRange.Count >= 1 Then
                        returnedVal = ActiveWindow.Selection.Type = ppSelectionShapes
                        Else
                        returnedVal = False
                        End If
                    Else
                    returnedVal = False
                    End If
                End If
        
        Case "Insert_Sticky_Note"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If
                
      Case "TextAndOthers"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If

        
        Case "Paste_Objects_All_Slides"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If
        
        Case "Show_All_Objects"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If
        
        Case "Remove_Spaces"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If
        
        Case "CleanDeck"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                Else
                returnedVal = True
                End If
        
        Case "Replace_Colors"
                On Error Resume Next
                If ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                returnedVal = False
                Else
                returnedVal = True
                End If

                
        
        Case "StampD_Update"
                On Error Resume Next
                If ActiveWindow.Selection.Type = ppSelectionShapes Then
                    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
                        If ActiveWindow.Selection.ShapeRange(1).Name = "PRODECK DOCUMENT STAMP" Then
                        returnedVal = True
                        Else
                        returnedVal = False
                        End If
                    Else
                    returnedVal = False
                    End If
                Else
                returnedVal = False
                End If
        
        Case "StampsMenu"
                On Error Resume Next
                If ActiveWindow.Application.Visible = msoFalse Then
                    returnedVal = False
                ElseIf ActiveWindow.Presentation.Slides.Count = 0 Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewHandoutMaster Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewMasterThumbnails Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesMaster Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewNotesPage Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewPrintPreview Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideMaster Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewSlideSorter Then
                    returnedVal = False
                ElseIf ActiveWindow.ActivePane.ViewType = ppViewTitleMaster Then
                    returnedVal = False
                ElseIf ActiveWindow.Selection.SlideRange.Count = 0 Then
                    returnedVal = False
                Else
                    returnedVal = True
                End If
        
        End Select
    
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


#If Mac Then
    
    Select Case control.Id
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
    
    Select Case control.Id
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
Sub RefreshRibbon(Id As String)

xmlID = Id

If Rib Is Nothing Then
Else
    Rib.Invalidate
End If

End Sub






