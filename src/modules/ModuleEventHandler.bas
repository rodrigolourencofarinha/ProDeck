' Extracted from: ModuleEventHandler.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "ModuleEventHandler"
Option Explicit
Public cPPTObject As New cEventClass
Public TrapFlag As Boolean
Sub TrapEvents()

'Creates an instance of the application event handler
If TrapFlag = True Then
   MsgBox "Relax, my friend, the EventHandler is already active.", vbInformation + vbOKOnly, "PowerPoint Event Handler Example"
   Exit Sub
End If

Set cPPTObject.PPTEvent = Application
TrapFlag = True
   
End Sub
Sub ReleaseTrap()

If TrapFlag = True Then
   Set cPPTObject.PPTEvent = Nothing
   Set cPPTObject = Nothing
   TrapFlag = False
End If

End Sub
