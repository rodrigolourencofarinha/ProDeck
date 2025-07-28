' Extracted from: WatermarkAddRemove.bas
' Source: ProDeck_v1_6_2.pptm

Attribute VB_Name = "WatermarkAddRemove"
Sub Watermark_Add()
On Error Resume Next
Dim Watermark As Shape
Const PI = 3.14159265358979

sngName = "PRODECK WATERMARK"
For Each sld In ActiveWindow.Presentation.Slides
    For Each shp In sld.Shapes
        If shp.Name = sngName Then
            shp.Delete
        End If
    Next shp
Next sld


   
   WatermarkText = InputBox("Please input watermark text", "Watermark", "CONFIDENTIAL")
   If WatermarkText = "" Then
   Exit Sub
   End If
   PredefinedColor = RGB(204, 0, 0)
   WatermarkTextColor = ColorDialog(PredefinedColor)
   
   ProgressForm.Show
   
   For Each PresentationSlide In ActiveWindow.Presentation.Slides
   With PresentationSlide
   
        SetProgress (PresentationSlide.SlideNumber / ActiveWindow.Presentation.Slides.Count * 100)
        
        Set Watermark = .Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=400, Height:=100)
        Watermark.Width = Sqr(Application.ActivePresentation.PageSetup.SlideWidth * Application.ActivePresentation.PageSetup.SlideWidth + Application.ActivePresentation.PageSetup.SlideHeight * Application.ActivePresentation.PageSetup.SlideHeight)
        Watermark.TextFrame.TextRange.text = WatermarkText
        Watermark.TextFrame.TextRange.Font.size = 100
        Watermark.TextFrame.HorizontalAnchor = msoAnchorCenter
        Watermark.Rotation = -Atn(Application.ActivePresentation.PageSetup.SlideHeight / Application.ActivePresentation.PageSetup.SlideWidth) * 180 / PI
        Watermark.Left = (Application.ActivePresentation.PageSetup.SlideWidth - Watermark.Width) / 2
        Watermark.Top = (Application.ActivePresentation.PageSetup.SlideHeight - Watermark.Height) / 2
        Watermark.Name = "PRODECK WATERMARK"
        
        Watermark.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = WatermarkTextColor
        Watermark.TextFrame2.TextRange.Characters.Font.Fill.Transparency = 0.9
        
    End With
    Next PresentationSlide
    
    ProgressForm.Hide
    
    'ConvertSlidesToPictures
   
End Sub

Sub Watermark_Delete()
On Error Resume Next
Dim shp As Shape
Dim sld As slide
 

For Each sld In ActiveWindow.Presentation.Slides
    For Each shp In sld.Shapes
        If shp.Name = "PRODECK WATERMARK" Then
            shp.Delete
        End If
    Next shp
Next sld

End Sub




