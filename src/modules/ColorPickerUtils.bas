' Extracted from: ColorPickerUtils.bas
' Source: ProDeck_v1_6_2.pptm

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
Function Hex2Dec(h)

Dim L As Long: L = Len(h)

If L < 16 Then               ' CDec results in Overflow error for hex numbers above 16 ^ 8

    Hex2Dec = CDec("&h0" & h)
    
    If Hex2Dec < 0 Then Hex2Dec = Hex2Dec + 4294967296# ' 2 ^ 32
    
ElseIf L < 25 Then

    Hex2Dec = Hex2Dec(Left$(h, L - 9)) * 68719476736# + CDec("&h" & Right$(h, 9)) ' 16 ^ 9 = 68719476736
    
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

