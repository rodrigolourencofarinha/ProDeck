Attribute VB_Name = "cEventClass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Public WithEvents PPTEvent As Application
Attribute PPTEvent.VB_VarHelpID = -1
Private Sub PPTEvent_WindowSelectionChange(ByVal Sel As Selection)
    'Force refresh of the "btInfo" button:
    RefreshRibbon "btInfo"
End Sub
