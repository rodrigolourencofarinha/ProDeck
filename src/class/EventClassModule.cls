VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As PowerPoint.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    InvalRibbon   'Calls procedure in the code module to invalidate the ribbon when the selection changes
    On Error Resume Next
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.ShapeRange.Type = msoTable Then
            Exit Sub
        ElseIf ActiveWindow.Selection.ShapeRange.Type = msoSmartArt Then
            Exit Sub
        'ElseIf ActiveWindow.Selection.ShapeRange.Type = msoPicture Then
            'Exit Sub
        ElseIf ActiveWindow.Selection.ShapeRange.AnimationSettings.Animate = msoTrue Then
            Exit Sub
        Else
            On Error Resume Next
            MyRibbon.ActivateTab ("Strator")
        End If
    Else
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
            Exit Sub
        Else
            If ActiveWindow.Selection.Type = ppSelectionText Then
                On Error Resume Next
                MyRibbon.ActivateTab ("Strator")
            Else
                Exit Sub
            End If
        End If
    End If
End Sub





