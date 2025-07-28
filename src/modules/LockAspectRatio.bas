' Extracted from: LockAspectRatio.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "LockAspectRatio"
Option Explicit
Public MyRibbon As IRibbonUI
Dim X As New EventClassModule
Public IsPressed As Boolean
'Procedure to invalidate the ribbon called from the event class
Sub InvalRibbon()
    MyRibbon.Invalidate
End Sub
'Callback for customUI.onLoad
Sub OnLoad(ribbon As IRibbonUI)
    Set MyRibbon = ribbon
    Set X.App = PowerPoint.Application
    'This option allows the tab to be activated on launch
    On Error Resume Next
    MyRibbon.ActivateTab ("Strator")
End Sub
'Callback for all Style Buttons getPressed
Sub getPressedStyleBtn(control As IRibbonControl, ByRef returnedVal)

    Select Case control.Id
    
    Case "LockUnlock"
    
        On Error Resume Next
        
        If ActiveWindow.Selection.HasChildShapeRange Then
            returnedVal = ActiveWindow.Selection.ChildShapeRange.LockAspectRatio = msoTrue
        ElseIf ActiveWindow.Selection.ShapeRange.Count > 0 Then
            returnedVal = ActiveWindow.Selection.ShapeRange.LockAspectRatio
        Else
            returnedVal = "False"
        End If
        
    End Select
    
End Sub
'Callback for LockTrue onAction
Sub Lock_and_Unlock(control As IRibbonControl, pressed As Boolean)

'START: Error message message box -----------------------------------------------
On Error Resume Next
    Err.Clear
   If ActiveWindow.Selection.ShapeRange.Count = 0 Then
      MsgBox "You must have at least one object selected.", vbCritical
      Rib.InvalidateControl (control.Id)
        Exit Sub
    End If
    If Err <> 0 Then
        Exit Sub
    End If
'END: Error message message box -------------------------------------------------
    
If pressed Then

    Select Case control.Id
    
        Case "LockUnlock"
            ActiveWindow.Selection.ShapeRange().LockAspectRatio = msoTrue
            ActiveWindow.Selection.ChildShapeRange().LockAspectRatio = msoTrue
            
    End Select
    
Else

    ActiveWindow.Selection.ShapeRange().LockAspectRatio = msoFalse
    ActiveWindow.Selection.ChildShapeRange().LockAspectRatio = msoFalse
    Rib.InvalidateControl (control.Id)
    
End If
    
End Sub

