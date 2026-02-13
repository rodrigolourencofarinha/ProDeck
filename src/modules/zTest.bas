Attribute VB_Name = "ZTest"
Option Explicit
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

Public Sub TestExecuteMso()
    Dim ids As Variant, i As Long, ok As Boolean
    
    ids = Array( _
        "IconsInsertGallery", _
        "InsertOnlinePictures", _
        "OnlinePicturesInsert", _
        "PicturesInsertOnline", _
        "StockImagesInsertGallery", _
        "InsertStockImages", _
        "InsertIcons", _
        "InsertM365Picture", _
        "IconInsertFromFile" _
    )
    
    Debug.Print "---- Testing ExecuteMso ----"
    For i = LBound(ids) To UBound(ids)
        ok = TryExecuteMso(CStr(ids(i)))
        Debug.Print ids(i) & vbTab & IIf(ok, "OK", "NO")
    Next i
End Sub

Private Function TryExecuteMso(ByVal idMso As String) As Boolean
    On Error GoTo EH
    Application.CommandBars.ExecuteMso idMso
    TryExecuteMso = True
    Exit Function
EH:
    TryExecuteMso = False
End Function






