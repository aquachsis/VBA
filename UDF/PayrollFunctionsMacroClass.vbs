'UDF User Definied Function
'Press Ctrl + Shift + A to tooltip!


Function AddTwoCells(cell1, cell2)
    AddTwoCells = cell1 + cell2
End Function

Function MultiplyTwoCells(cell1, cell2)
    MultiplyTwoCells = cell1 * cell2
End Function

Function LastRow()
    'This gives last row for column 1
    LastRow = ActiveSheet.Cells(Rows.Count,1).End(xlUp).Row
End Function

Function LastRowByColumn(SelectedCell As Range)
    'TODO As Range otherwise it might thing it's a string?
    'This gives last row of selected cell or column
    ColumnNumber = SelectedCell.Column
    LastRowByColumn = ActiveSheet.Cells(Rows.Count,ColumnNumber).End(xlUp).Row
End Function

Function VlookupWithMessage(lookup_value, table_array, column_index, error_message)
    VlookupWithMessage = _
        Application.WorksheetFunction.IfError( _
            Application.VLOOKUP(lookup_value, table_array, column_index, 0), _
        error_message)
End Function
