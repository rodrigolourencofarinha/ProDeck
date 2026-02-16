Attribute VB_Name = "ActivateTab"
Option Explicit

Private Const AUTO_TAB_APP_NAME As String = "ProDeck"
Private Const AUTO_TAB_SECTION As String = "Preferences"
Private Const AUTO_TAB_KEY As String = "AutoTabMonitoringEnabled"
Private Const AUTO_TAB_DEFAULT As String = "True"

Private mAutoTabMonitoringEnabled As Boolean
Private mAutoTabMonitoringLoaded As Boolean

Sub ActivateProDeckTab()
    On Error Resume Next
    MyRibbon.ActivateTab ("ProDeck")
End Sub

Sub MaybeActivateProDeckTab()
    If Not IsAutoTabMonitoringEnabled() Then
        Exit Sub
    End If

    ActivateProDeckTab
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
    Dim userAnswer As VbMsgBoxResult

    If pressed Then
        promptText = "Enable automatic return to the ProDeck tab when selecting shapes or text?"
    Else
        promptText = "Disable automatic return to the ProDeck tab when selecting shapes or text?"
    End If

    userAnswer = MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton1, "ProDeck")

    If userAnswer = vbYes Then
        SetAutoTabMonitoringEnabled pressed

        If pressed Then
            ActivateProDeckTab
        End If
    End If

    On Error Resume Next
    MyRibbon.InvalidateControl "AutoTabMonitoring"
End Sub

Sub GetAutoTabMonitoringPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IsAutoTabMonitoringEnabled()
End Sub

Private Sub SetAutoTabMonitoringEnabled(ByVal isEnabled As Boolean)
    mAutoTabMonitoringEnabled = isEnabled
    mAutoTabMonitoringLoaded = True
    SaveSetting AUTO_TAB_APP_NAME, AUTO_TAB_SECTION, AUTO_TAB_KEY, CStr(isEnabled)
End Sub

Private Sub EnsureAutoTabMonitoringSettingLoaded()
    Dim savedValue As String

    If mAutoTabMonitoringLoaded Then
        Exit Sub
    End If

    savedValue = GetSetting(AUTO_TAB_APP_NAME, AUTO_TAB_SECTION, AUTO_TAB_KEY, AUTO_TAB_DEFAULT)
    mAutoTabMonitoringEnabled = (LCase$(Trim$(savedValue)) = "true")
    mAutoTabMonitoringLoaded = True
End Sub

