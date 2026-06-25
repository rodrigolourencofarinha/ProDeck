Attribute VB_Name = "RibbonEnableState"
Option Explicit

Public Function Ribbon_IsEnabled(ByVal controlId As String) As Boolean
    Dim returnedVal As Boolean

    If RibbonEnableInsert_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableText_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableFormat_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableArrange_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableSizeOrganize_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    If RibbonEnableCleanup_IsEnabled(controlId, returnedVal) Then
        Ribbon_IsEnabled = returnedVal
        Exit Function
    End If

    Ribbon_IsEnabled = False
End Function
