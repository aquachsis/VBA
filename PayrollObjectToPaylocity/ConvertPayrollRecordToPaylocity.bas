Public LastRow As Long

Sub FormatTimeClock()
    'Payroll Number and Week Ending, both uncessary'
    Columns("A:B").EntireColumn.Delete

    LastRow = FindLastRow(1)
    Call FillData("UID (EE# + Rate)", "I", "=CONCATENATE($A2,""|"",$B2)")
    Call FillData("Regular", "J", "=IF($E2=0,"""",$E2)")
    Call FillData("Overtime", "K", "=IF($F2=0,"""",$F2)")
    Call FillData("Double", "L", "=IF($G2=0,"""",$G2)")
    Call FillData("Penalty Hour", "M", "=IF($H2=0,"""",$H2)")
End Sub


Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function

Public Sub FillData(Header, ColumnLetter, Data)
    Range(ColumnLetter & "1") = Header
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    ' Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
    '     Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
    Columns.AutoFit
End Sub
