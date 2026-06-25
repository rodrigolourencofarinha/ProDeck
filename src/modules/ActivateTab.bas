Attribute VB_Name = "ActivateTab"
Option Explicit

Private Const AUTO_TAB_APP_NAME As String = "ProDeck"
Private Const AUTO_TAB_SECTION As String = "Preferences"
Private Const AUTO_TAB_KEY As String = "AutoTabMonitoringEnabled"
Private Const AUTO_TAB_DEFAULT As String = "True"
Private Const PRODECK_TAB_ID As String = "ProDeck"
Private Const PRODECK_TAB_KEYTIP As String = "%X"

Private mAutoTabMonitoringEnabled As Boolean
Private mAutoTabMonitoringLoaded As Boolean
Private mAutoTabRecoveryNoticeShown As Boolean

Sub ActivateProDeckTab(Optional control As IRibbonControl)
    InvalRibbon
    If Not TryActivateProDeckTab(True) Then ShowAutoTabRecoveryMessage
End Sub

Sub MaybeActivateProDeckTab()
    If IsAutoTabMonitoringEnabled() Then
        ArmAutoTabEventHandler
        Call TryActivateProDeckTab(True)
    End If
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
            mAutoTabRecoveryNoticeShown = False
            InvalRibbon
            If Not TryActivateProDeckTab(True) Then ShowAutoTabRecoveryMessage
        End If
    End If

    InvalRibbon
End Sub

Sub GetAutoTabMonitoringPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IsAutoTabMonitoringEnabled()
    If returnedVal Then ArmAutoTabEventHandler
End Sub

Sub RepairAutoTabMonitoring(Optional control As IRibbonControl)
    mAutoTabMonitoringEnabled = True
    mAutoTabMonitoringLoaded = True
    mAutoTabRecoveryNoticeShown = False
    SaveAutoTabMonitoringSetting True
    ArmAutoTabEventHandler
    InvalRibbon
    If Not TryActivateProDeckTab(True) Then ShowAutoTabRecoveryMessage
End Sub

Public Function TryActivateProDeckTab(Optional ByVal allowKeytipFallback As Boolean = True) As Boolean
    On Error Resume Next

    If MyRibbon Is Nothing Then Set MyRibbon = Rib

    If TryActivateRibbonObject(MyRibbon) Then
        TryActivateProDeckTab = True
        Exit Function
    End If

    If TryActivateRibbonObject(Rib) Then
        Set MyRibbon = Rib
        TryActivateProDeckTab = True
        Exit Function
    End If

    Set MyRibbon = Nothing
    Set Rib = Nothing

    If allowKeytipFallback Then
        TryActivateProDeckTab = TryActivateProDeckTabWithKeytip()
    End If
End Function

Private Function TryActivateRibbonObject(ByVal ribbon As IRibbonUI) As Boolean
    On Error GoTo ActivationFailed

    If ribbon Is Nothing Then Exit Function
    ribbon.ActivateTab PRODECK_TAB_ID
    TryActivateRibbonObject = True
    Exit Function

ActivationFailed:
    Err.Clear
End Function

Private Function TryActivateProDeckTabWithKeytip() As Boolean
#If Mac Then
    TryActivateProDeckTabWithKeytip = False
#Else
    On Error Resume Next
    Err.Clear
    SendKeys PRODECK_TAB_KEYTIP
    TryActivateProDeckTabWithKeytip = (Err.Number = 0)
    Err.Clear
#End If
End Function

Private Sub ShowAutoTabRecoveryMessage()
    If mAutoTabRecoveryNoticeShown Then Exit Sub
    mAutoTabRecoveryNoticeShown = True
    MsgBox "PowerPoint lost the live ProDeck ribbon connection. ProDeck re-armed selection monitoring, but PowerPoint may need the add-in to be reloaded before automatic tab activation can resume.", vbExclamation + vbOKOnly, "ProDeck"
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
