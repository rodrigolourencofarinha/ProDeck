Attribute VB_Name = "TableDistributeWithGaps"
Sub TableColumnGapsEven()
    TableColumnGaps "even", 5
End Sub
Sub TableColumnGapsOdd()
    TableColumnGaps "odd", 5
End Sub
Sub TableRowGapsEven()
    TableRowGaps "even", 5
End Sub
Sub TableRowGapsOdd()
    TableRowGaps "odd", 5
End Sub
Sub TableDistributeColumnsWithGaps()
   
Dim TotalWidth As Double
Dim NumberOfColumnsToDistribute As Long
TotalWidth = 0
NumberOfColumnsToDistribute = 0
 
If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
    MsgBox "No table or cells selected."
Else

If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
With Application.ActiveWindow.Selection.ShapeRange.Table
    
    TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
    
    For ColsCount = 1 To .Columns.Count
    
        For RowsCount = 1 To .Rows.Count
            
            If .Cell(RowsCount, ColsCount).Selected Then
                
            If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
            
            TotalWidth = TotalWidth + .Columns(ColsCount).Width
            NumberOfColumnsToDistribute = NumberOfColumnsToDistribute + 1
            Exit For
                
            End If
                
            
            End If
            
        Next RowsCount
        Next ColsCount
        
        
        If NumberOfColumnsToDistribute > 0 Then
        For ColsCount = 1 To .Columns.Count
        
            For RowsCount = 1 To .Rows.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                .Columns(ColsCount).Width = TotalWidth / NumberOfColumnsToDistribute
                Exit For
                    
                End If
                    
                
                End If
                
            Next RowsCount
        Next ColsCount
        End If
        
    End With
    
    Else
    
    MsgBox "No table or cells selected.", vbCritical
    
    End If
    
    End If

End Sub
Sub TableDistributeRowsWithGaps()
    
    Dim TotalHeight As Double
    Dim NumberOfRowsToDistribute As Long
    TotalHeight = 0
    NumberOfRowsToDistribute = 0
     
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected.", vbCritical
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
        
        For RowsCount = 1 To .Rows.Count
           
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                TotalHeight = TotalHeight + .Rows(RowsCount).Height
                NumberOfRowsToDistribute = NumberOfRowsToDistribute + 1
                Exit For
                    
                End If
                    
                
                End If
                
            Next ColsCount
        Next RowsCount
        
        
        If NumberOfRowsToDistribute > 0 Then
        
        For RowsCount = 1 To .Rows.Count
        
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                .Rows(RowsCount).Height = TotalHeight / NumberOfRowsToDistribute
                Exit For
                    
                End If
                    
                
                End If
                
            Next ColsCount
        Next RowsCount
        End If
        
    End With
    
    Else
    
    MsgBox "No table or cells selected.", vbCritical
    
    End If
    
    End If

End Sub
Sub TableColumnGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected.", vbCritical
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even" Then
                
                If MsgBox("Existing column gaps found in table, do you want to remove those first?", vbYesNo + vbExclamation) = vbYes Then
                    TableColumnRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                NumberOfColumns = .Columns.Count
                Dim ColumnWidthArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns + 1
                    ReDim ColumnWidthArray(0)
                    
                    For ColumnCount = 1 To NumberOfColumns
                        ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                        ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                        
                        If ColumnCount = NumberOfColumns Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = GapSize
                        End If
                        
                    Next ColumnCount
                    
                Else
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns - 1
                    
                    For ColumnCount = 1 To NumberOfColumns
                        
                        If Not ColumnCount = 1 Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                            ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                            
                        Else
                            ReDim ColumnWidthArray(1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        End If
                        
                    Next ColumnCount
                    
                End If
                
                For ColumnCount = NumberOfColumns To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedColumn = .Columns.Add(ColumnCount)
                        
                        For CellCount = 1 To AddedColumn.Cells.Count
                            AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                            AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                            AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                            
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If ColumnCount = NumberOfColumns Then
                            
                            Set AddedColumn = .Columns.Add
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                                
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not ColumnCount = 1 Then
                            
                            Set AddedColumn = .Columns.Add(ColumnCount)
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                                
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next ColumnCount
                
                For ColumnCount = 1 To NumberOfNewColumns
                    
                    .Columns(ColumnCount).Width = ColumnWidthArray(ColumnCount - 1)
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table.", vbCritical
            
        End If
        
    End If
    
End Sub
Sub TableColumnIncreaseGaps()
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected.", vbCritical
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width + 1
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
Sub TableColumnDecreaseGaps()
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If ((ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Columns(ColumnCount).Width - 1) >= 0)) Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width - 1
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
Sub TableColumnRemoveGaps()
     
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA COLUMNGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = .Columns.Count To 1 Step -1
                    
                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Delete
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
Sub TableRowGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even" Then
                
                If MsgBox("Existing row gaps found in table, do you want to remove those first?", vbYesNo + vbExclamation) = vbYes Then
                    TableRowRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                NumberOfRows = .Rows.Count
                Dim RowHeightArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows + 1
                    ReDim RowHeightArray(0)
                    
                    For RowCount = 1 To NumberOfRows
                        ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                        RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                        RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                        
                        If RowCount = NumberOfRows Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = GapSize
                        End If
                        
                    Next RowCount
                    
                Else
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows - 1
                    
                    For RowCount = 1 To NumberOfRows
                        
                        If Not RowCount = 1 Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                            RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                            
                        Else
                            ReDim RowHeightArray(1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                        End If
                        
                    Next RowCount
                    
                End If
                
                For RowCount = NumberOfRows To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedRow = .Rows.Add(RowCount)
                        
                        For CellCount = 1 To AddedRow.Cells.Count
                            AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                            AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                            AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                            AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                            
                            AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                            AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                            AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                            AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If RowCount = NumberOfRows Then
                            
                            Set AddedRow = .Rows.Add
                            
                            For CellCount = 1 To AddedRow.Cells.Count
                                AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                                
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not RowCount = 1 Then
                            
                            Set AddedRow = .Rows.Add(RowCount)
                            
                            For CellCount = 1 To AddedRow.Cells.Count
                                AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.size = 1
                                
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next RowCount
                
                For RowCount = 1 To NumberOfNewRows
                    
                    .Rows(RowCount).Height = RowHeightArray(RowCount - 1)
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
Sub TableRowIncreaseGaps()
          
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height + 1
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
Sub TableRowDecreaseGaps()
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If ((RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Rows(RowCount).Height - 1) >= 0)) Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height - 1
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowRemoveGaps()
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA ROWGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = .Rows.Count To 1 Step -1
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Delete
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub


