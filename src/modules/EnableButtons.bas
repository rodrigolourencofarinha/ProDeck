Attribute VB_Name = "EnableButtons"
Option Explicit
Public Rib As IRibbonUI
Public xmlID As String
'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)

    TrapEvents  'instantiate the event handler

    Set Rib = ribbon

End Sub

Sub EnabledBtInfo(control As IRibbonControl, ByRef returnedVal)

    returnedVal = Ribbon_IsEnabled(control.Id)

    Call RefreshRibbon(control.Id)

End Sub
Public Sub IsMac(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = True
#Else
    returnedVal = False
#End If
End Sub
Public Sub IsWindows(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub
Sub RibbonObjectGetImage(control As IRibbonControl, ByRef returnedVal)

    Ribbon_GetImage control.Id, returnedVal

End Sub
Sub RefreshRibbon(Id As String)

xmlID = Id

If Rib Is Nothing Then
Else
    Rib.Invalidate
End If

End Sub
