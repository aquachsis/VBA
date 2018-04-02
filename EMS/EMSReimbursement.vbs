Sub ExpenseReimbursement()
    Dim UpcomingReimbursementFile As String
    Dim UpcomingReimbursementName As String
    MsgBox "Select the " & vbNewLine & _
        "Expense Pending Reimbursement - Verification.csv file."
    UpcomingReimbursementFile = Application.GetOpenFilename

    Dim HistoricalReimbursedFile As String
    Dim HistoricalReimbursedName As String
    MsgBox "Select the " & vbNewLine & _
        "Expense Reimbursed - Historical.csv file."
    HistoricalReimbursedFile = Application.GetOpenFilename

    'Get the main workbook started
    Workbooks.Open Filename:= UpcomingReimbursementFile
    UpcomingReimbursementName = ActiveWorkbook.Name
    Sheets(1).Name = "Upcoming"

    'Copy the historical worksheet to the main workbook
    Workbooks.Open Filename:= HistoricalReimbursedFile
    HistoricalReimbursedName = ActiveWorkbook.Name
    Sheets(1).Name = "Historical"
    Sheets("Historical").Copy _
        Before:=Workbooks(UpcomingReimbursementName).Sheets(1)
    Workbooks(HistoricalReimbursedName).Close SaveChanges:=False

    Sheets.Add().Name = "Summary"

    'Starting comparison

    Worksheets("Upcoming").Activate
    Dim LastRow As Long
    LastRow = Worksheets("Upcoming").Cells(Rows.Count,3).End(xlUp).Row

    'Creating a unique ID
    Range("I1") = "Unique ID"
    Range("I2:I" & LastRow).Formula = _
        "=TEXTJOIN(""|"",,RC[-4]:RC[-2])&""|""&TEXT(RC[-1],""mm/dd/yyyy"")"
    Range("I2:I" & LastRow).Value = Range("I2:I" & LastRow).Value

    Range("J1") = "Duplicate? (Excludes Mileage)"
    Range("J2:J" & LastRow).Formula = _
        "=IF(COUNTIFS(C[-4],""<>""&""Mileage"",C[-1],RC[-1])>1,""1"","""")"
    Range("J2:J" & LastRow).Value = Range("J2:J" & LastRow).Value

    Range("A1").CurrentRegion.Sort _
        Key1:=Range("J1"), Order1:=xlAscending, _
        Key2:=Range("I1"), Order2:=xlAscending, _
        Header:=xlYes


    Range("K1").Formula = "=COUNT(J:J)"
    Dim DuplicateCount As Long
    DuplicateCount = Range("K1")

    If DuplicateCount = 0 Then
        MsgBox "No duplicates."
    Else
        MsgBox "Number of duplicates: " & DuplicateCount
    End If

'historical uid'
Worksheets("Historical").Activate
LastRow = Worksheets("Historical").Cells(Rows.Count,3).End(xlUp).Row

'Creating a unique ID
'TODO make this is subroutine since col. are same between verification &
'historical
Range("I1") = "Unique ID"
Range("I2:I" & LastRow).Formula = _
    "=TEXTJOIN(""|"",,RC[-4]:RC[-2])&""|""&TEXT(RC[-1],""mm/dd/yyyy"")"
Range("I2:I" & LastRow).Value = Range("I2:I" & LastRow).Value

'TODO ignore mileage on historical as well'


    ' ActiveWorkbook.SaveAs _
    '   "Expense Reimbursement - Checking For Duplicates.xlsx", _
    '   FileFormat:=xlOpenXMLWorkbook, _
    '   CreateBackup:=False
End Sub








Sub EMSHyperlink()
Dim i as Long
Dim url as String
i = 2
Do While Range("F" & i).Value <> ""
    url = "https://na7.salesforce.com/" & Range("G" & i)
    ActiveSheet.Hyperlinks.Add Range("F" & i), url
    i = i + 1
Loop
End Sub
