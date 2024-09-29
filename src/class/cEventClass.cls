VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEventClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents PPTEvent As Application
Attribute PPTEvent.VB_VarHelpID = -1
Private Sub PPTEvent_WindowSelectionChange(ByVal Sel As Selection)
    'Force refresh of the "btInfo" button:
    RefreshRibbon "btInfo"
End Sub
