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
