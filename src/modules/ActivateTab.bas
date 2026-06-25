Attribute VB_Name = "ActivateTab"
Option Explicit

Private Const AUTO_TAB_APP_NAME As String = "ProDeck"
Private Const AUTO_TAB_SECTION As String = "Preferences"
Private Const AUTO_TAB_KEY As String = "AutoTabMonitoringEnabled"
Private Const AUTO_TAB_DEFAULT As String = "True"

Private mAutoTabMonitoringEnabled As Boolean
Private mAutoTabMonitoringLoaded As Boolean

Sub ActivateProDeckTab(Optional control As IRibbonControl)
    On Error Resume Next
    If MyRibbon Is Nothing Then Set MyRibbon = Rib
    If Not MyRibbon Is Nothing Then MyRibbon.ActivateTab "ProDeck"
End Sub

Sub MaybeActivateProDeckTab()
    If IsAutoTabMonitoringEnabled() Then ActivateProDeckTab
End Sub

Sub InitializeAutoTabMonitoringSetting()
    mAutoTabMonitoringLoaded = False
    EnsureAutoTabMonitoringSettingLoaded
End Sub

Public Function IsAutoTabMonitoringEnabled() As Boolean
    EnsureAutoTabMonitoringSettingLoaded
    IsAutoTabMonitoringEnabled = mAutoTabMonitoringEnabled
End Function

Sub ToggleAutoTabMonitoring(control As IRibbonControl, pressed As Boolean)
    Dim promptText As String

    promptText = IIf(pressed, "Enable", "Disable") & " automatic return to the ProDeck tab when selecting shapes or text?"

    If MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton1, "ProDeck") = vbYes Then
        mAutoTabMonitoringEnabled = pressed
        mAutoTabMonitoringLoaded = True
        SaveAutoTabMonitoringSetting pressed
        If pressed Then
            ArmAutoTabEventHandler
            ActivateProDeckTab
        End If
    End If

    InvalRibbon
End Sub

Sub GetAutoTabMonitoringPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IsAutoTabMonitoringEnabled()
End Sub

Private Sub EnsureAutoTabMonitoringSettingLoaded()
    Dim savedValue As String

    If mAutoTabMonitoringLoaded Then
        Exit Sub
    End If

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

Private Sub SaveAutoTabMonitoringSetting(ByVal isEnabled As Boolean)
    On Error Resume Next
    SaveSetting AUTO_TAB_APP_NAME, AUTO_TAB_SECTION, AUTO_TAB_KEY, CStr(isEnabled)
    Err.Clear
End Sub
