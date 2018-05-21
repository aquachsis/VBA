Sub ExtractSingleEmployee()

Dim LastRow As Long
Dim EmployeeNumberColumn As Long
Dim DirArray As Variant

EmployeeNumberColumn = 5
LastRow = Cells(Rows.Count, EmployeeNumberColumn).End(xlUp).Row

DirArray = Range("E1:E" & LastRow).Value

Range("A1:A" & LastRow) = DirArray
End Sub
