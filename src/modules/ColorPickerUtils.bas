Attribute VB_Name = "ColorPickerUtils"
Option Explicit

Public Type PickColor
    red As Integer
    green As Integer
    blue As Integer
End Type

Public continue As Boolean

Private pColor As Long
' Launch a color picker form in one of three ways:
' 1) no arguments: initial color is black
' 2) 1 Long argument: intial color set with Long color value
' 3) 3 Long arguments: initial color set with RGB(red,green,blue)
Public Function ColorPicker(Optional ByVal red As Long = -1, _
        Optional ByVal green As Long = -1, Optional ByVal blue As Long = -1) As Long
        
Load ColorPickerForm
ColorPickerForm.resetForm

If Not red = -1 Then

    If (green = -1 And blue = -1) Then
        ColorPickerForm.SetInitColor red
    Else
        ColorPickerForm.SetInitColor RGB(red, green, blue)
    End If
    
End If

ColorPickerForm.Show
ColorPicker = ColorPickerForm.GetSelectedColor
pColor = ColorPickerForm.GetSelectedColor
Unload ColorPickerForm

End Function
Public Function SelectedColor() As Long

    SelectedColor = pColor
    
End Function
' get separate R-G-B values from a color stored as a Long
Public Function GetRGBFromLong(ByVal color As Long) As PickColor

    Dim newColor As PickColor
    newColor.red = color Mod 256
    newColor.green = color \ 256 Mod 256
    newColor.blue = color \ (65536) Mod 256
    GetRGBFromLong = newColor
    
End Function
Function GetHexFromRGB(color As PickColor) As String

Dim value As Long

value = color.blue + color.green * CLng(256) + color.red * CLng(65536)

GetHexFromRGB = Hex(value)

Do While Len(GetHexFromRGB) < 6
    GetHexFromRGB = "0" & GetHexFromRGB
Loop

End Function
Function GetRGBFromHex(hexString As String) As PickColor

    GetRGBFromHex = GetRGBFromLong(GetLongFromHex(hexString))
    
End Function
Public Function Hex2Dec(ByVal h As String) As Long
    h = Replace$(h, "#", "")
    h = Replace$(h, "&H", "")
    h = Replace$(h, "&h", "")

    If Len(h) = 0 Then
        Hex2Dec = 0
    Else
        ' Works on Mac/Windows for typical color hex (<= 8 hex digits)
        Hex2Dec = CLng("&H" & h)
    End If
End Function
Function GetLongFromHex(hexColor As String) As Long

Dim r As String
Dim g As String
Dim B As String

hexColor = VBA.Replace(hexColor, "#", "")
hexColor = VBA.Right$("000000" & hexColor, 6)

r = Left(hexColor, 2)
g = Mid(hexColor, 3, 2)
B = Right(hexColor, 2)

GetLongFromHex = Hex2Dec(B & g & r)

End Function

