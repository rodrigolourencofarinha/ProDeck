Attribute VB_Name = "ZTest"
Sub AAAPrint()
Dim shp As Shape
Dim sTemp As Variant

sTemp = ActiveWindow.Selection.ShapeRange(1).Left
MsgBox sTemp
sTemp = ActiveWindow.Selection.ShapeRange(1).Top
MsgBox sTemp
sTemp = ActiveWindow.Selection.ShapeRange(1).Height
MsgBox sTemp


End Sub






