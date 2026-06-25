Attribute VB_Name = "LockAspectRatio"
Option Explicit
Public MyRibbon As IRibbonUI
Private X As EventClassModule
Public IsPressed As Boolean
Sub InvalRibbon()
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
Sub ArmAutoTabEventHandler()
    On Error Resume Next
    If X Is Nothing Then Set X = New EventClassModule
    Set X.App = Application
    If Err.Number <> 0 Then
        Err.Clear
        Set X = Nothing
        Set X = New EventClassModule
        Set X.App = Application
    End If
End Sub
'Callback for customUI.onLoad
Sub OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Set MyRibbon = ribbon
    Set Rib = ribbon
    ArmAutoTabEventHandler
    If Not TrapFlag Then TrapEvents
    InitializeAutoTabMonitoringSetting
    MaybeActivateProDeckTab
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



