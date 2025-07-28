' Extracted from: TableToShapes.bas
' Source: ProDeck_v1_6_3.pptm

Attribute VB_Name = "TableToShapes"
Sub ConvertTableToShapes()

Dim NewShape As Shape

          
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
    MsgBox "Please select a table."
    
    ElseIf ActiveWindow.Selection.ShapeRange.HasTable Then
    
    TableTop = ActiveWindow.Selection.ShapeRange.Top
    TableLeft = ActiveWindow.Selection.ShapeRange.Left
    

    ProgressForm.Show
    
    For RowsCount = 1 To ActiveWindow.Selection.ShapeRange.Table.Rows.Count
    
    SetProgress (RowsCount / ActiveWindow.Selection.ShapeRange.Table.Rows.Count * 100)
    
        For ColsCount = 1 To ActiveWindow.Selection.ShapeRange.Table.Columns.Count
                
            Set NewShape = ActiveWindow.Selection.SlideRange.Shapes.AddTextbox(msoShapeRectangle, Left:=TableLeft, Top:=TableTop, Width:=ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width, Height:=ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height)

            With NewShape
                .TextFrame.TextRange.text = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.text
                .TextFrame.MarginBottom = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginBottom
                .TextFrame.MarginLeft = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginLeft
                .TextFrame.MarginRight = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginRight
                .TextFrame.MarginTop = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginTop
                .Fill.ForeColor.RGB = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.Fill.ForeColor.RGB
                .line.ForeColor.RGB = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders(ppBorderBottom).ForeColor.RGB
            End With
            
            With NewShape.TextFrame2.TextRange
            .Font.size = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Font.size
            .Font.Bold = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Font.Bold
            .Font.Italic = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Font.Italic
            .Font.Fill.ForeColor = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Font.Fill.ForeColor
            .ParagraphFormat.Alignment = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.ParagraphFormat.Alignment
            End With
        
            
            With NewShape
            .TextFrame2.AutoSize = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.AutoSize
            .TextFrame2.WordWrap = msoTrue
            .TextFrame2.AutoSize = msoAutoSizeNone
            .TextFrame2.VerticalAnchor = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.VerticalAnchor
            .TextFrame2.MarginBottom = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.MarginBottom
            .TextFrame2.MarginTop = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.MarginTop
            .TextFrame2.MarginLeft = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.MarginLeft
            .TextFrame2.MarginRight = ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.MarginRight
            .Left = TableLeft
            .Top = TableTop
            .Width = ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            .Height = ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
            End With
            
            TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            
        Next ColsCount
        
        TableLeft = Application.ActiveWindow.Selection.ShapeRange.Left
        TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
        
    Next RowsCount
    
    ProgressForm.Hide
    
    Application.ActiveWindow.Selection.ShapeRange.Delete
    
    Else
    
    MsgBox "No table selected."
    
    End If
       
End Sub


