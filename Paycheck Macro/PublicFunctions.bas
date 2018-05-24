Attribute VB_Name = "PublicFunctions"

Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function

Public Function FindLastColumn()
    FindLastColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
End Function
