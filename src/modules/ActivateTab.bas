Attribute VB_Name = "ActivateTab"
Option Explicit

Sub ActivateProDeckTab(Optional control As IRibbonControl)
    AutoTab_Activate
End Sub

Sub MaybeActivateProDeckTab()
    AutoTab_MaybeActivate
End Sub

Sub InitializeAutoTabMonitoringSetting()
    AutoTab_Initialize
End Sub

Public Function IsAutoTabMonitoringEnabled() As Boolean
    IsAutoTabMonitoringEnabled = AutoTab_IsEnabled()
End Function

Sub ToggleAutoTabMonitoring(control As IRibbonControl, pressed As Boolean)
    AutoTab_Toggle pressed
End Sub

Sub GetAutoTabMonitoringPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = AutoTab_IsEnabled()
    If returnedVal Then AutoTab_ArmEventHandler
End Sub

Sub RepairAutoTabMonitoring(Optional control As IRibbonControl)
    AutoTab_Repair
End Sub
