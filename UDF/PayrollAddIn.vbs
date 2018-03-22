Function JoinText(InRange, Delim)
'Concatenate with a delimiter
  Dim Cell As Range
  For Each Cell In InRange
    JoinText = JoinText & Cell.Value & Delim
  Next Cell
  JoinText = Left(JoinText, Len(JoinText) - Len(Delim))
End Function

Function LastRowByColumn(SelectedCell As Range)
    'TODO As Range otherwise it might thing it's a string?
    'This gives last row of selected cell or column
    ColumnNumber = SelectedCell.Column
    LastRowByColumn = ActiveSheet.Cells(Rows.Count,ColumnNumber).End(xlUp).Row
End Function

Function VlookupSimple(lookup_value, table_array)
'Vlookup that only returns an exact match, and returns value from last column
  column_index = table_array.Columns.Count
  VlookupSimple = Application.VLOOKUP(lookup_value, table_array, column_index, 0)
End Function

Function VlookupWithMessage(lookup_value, table_array, column_index, error_message)
'Vlookup that only returns an exact match, and wrapped in a iferror
    VlookupWithMessage = _
        Application.WorksheetFunction.IfError( _
            Application.VLOOKUP(lookup_value, table_array, column_index, 0), _
        error_message)
End Function
