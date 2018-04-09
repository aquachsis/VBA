Public LastRow As Long
Public CalcColumn As String
Public CalcHeader As String
Public CalcFirstCell As String
Public CalcLastCell As String
Public CalcFormula As String
Public UIDFormula As String
' Public DuplicateCount As Long

Sub EMS()
    Call OpenExpenseFiles()
    UIDFormula = _
        "=TEXTJOIN(""|"",FALSE,RC[5]:RC[7])&""|""&TEXT(RC[8],""mm/dd/yyyy"")"
    Call FindDuplicatesInUpcoming
    Call FindDuplicatesAgainstHistorical
    Call CreateSummarySheet
    Call SaveAsXLSX
    MsgBox "Macro completed."
End Sub

Sub OpenExpenseFiles()
    Dim UpcomingReimbursementFile As String
    Dim UpcomingReimbursementName As String
    Dim HistoricalReimbursedFile As String
    Dim HistoricalReimbursedName As String

    MsgBox "Select the upcoming reimbursement file."
    UpcomingReimbursementFile = Application.GetOpenFilename
    MsgBox "Select the historical reimbursed file."
    HistoricalReimbursedFile = Application.GetOpenFilename

    'Get the main workbook started
    Workbooks.Open Filename:=UpcomingReimbursementFile
    UpcomingReimbursementName = ActiveWorkbook.Name
    Sheets(1).Name = "Upcoming"

    'Copy the historical worksheet to the main workbook
    Workbooks.Open Filename:=HistoricalReimbursedFile
    HistoricalReimbursedName = ActiveWorkbook.Name
    Sheets(1).Name = "Historical"
    Sheets("Historical").Copy _
        Before:=Workbooks(UpcomingReimbursementName).Sheets(1)
    Workbooks(HistoricalReimbursedName).Close SaveChanges:=False
End Sub

Sub InsertCalculations(ColumnLetter, HeaderTitle, CalcFormula)
    CalcColumn = ColumnLetter
    LastRow = Cells(Rows.Count, 4).End(xlUp).Row 'col4 is ems ID, never blank
    CalcHeader = CalcColumn & 1 'row 1 is where the header starts
    CalcFirstCell = CalcColumn & 2 'row 2 is where data starts
    CalcLastCell = CalcColumn & LastRow
    Range(CalcHeader) = HeaderTitle
    Set CalcRange = Range(CalcFirstCell & ":" & CalcLastCell)
    CalcRange.FormulaR1C1 = CalcFormula
    CalcRange.Value = CalcRange.Value
End Sub

Sub SortForReadability()
    'Sort by duplicates, then UID, for readability.
    Range("A1").CurrentRegion.Sort _
        Key1:=Range("J1"), Order1:=xlAscending, _
        Key2:=Range("A1"), Order2:=xlAscending, _
        Header:=xlYes
End Sub

Sub FindDuplicatesInUpcoming()
    Dim DuplicateFormula As String
    DuplicateFormula = _
        "=IF(COUNTIFS(C[-3],""<>""&""Mileage"",C[-9],RC[-9])>1,1,"""")"

    Worksheets("Upcoming").Activate
    Range("A1").EntireColumn.Insert
    Call InsertCalculations("A", "UniqueID", UIDFormula)
    Call InsertCalculations("J", "Duplicate? (Excludes Mileage)", DuplicateFormula)
    Call SortForReadability
End Sub

Sub FindDuplicatesAgainstHistorical()
    Worksheets("Historical").Activate
    Range("A1").EntireColumn.Insert
    Call InsertCalculations("A", "UniqueID", UIDFormula)

    'TODO exclude mileage'
    CalcFormula = "=IF(RC[-3]=""Mileage"","""",IFNA(VLOOKUP(RC[-9],Upcoming!C[-9]:C[-6],4,0),""""))"
    Call InsertCalculations("J", "Duplicate Reimbursement?", CalcFormula)
    Call SortForReadability
End Sub

Sub CreateSummarySheet()
    Sheets.Add().Name = "Summary"

    'TODO only if there are duplicates'
    'Find Last Row with Duplicates, in Upcoming Sheet
    LastRow = Worksheets("Upcoming").Cells(Rows.Count, Range("J1").Column).End(xlUp).Row
    Worksheets("Upcoming").Rows("1:" & LastRow).Copy _
        Destination:=Worksheets("Summary").Range("A1")

    NextAvailableRow = LastRow + 1

    'TODO only if there are duplicates'
    'Find Last Row with Duplicates, in Historical Sheet
    LastRow = Worksheets("Historical").Cells(Rows.Count, Range("J1").Column).End(xlUp).Row
    Worksheets("Historical").Rows("1:" & LastRow).Copy _
        Destination:=Worksheets("Summary").Range("A" & NextAvailableRow)

    CalcFormula = _
      "=IF(RC[-1]=1,""Duplicate Within Upcoming Reimbursements"",""Trying to reimburse expense again."")"
    Call InsertCalculations("K", "Notes", CalcFormula)
End Sub

Sub SaveAsXLSX()
    OutfileNameXLSX = "Expense Reimbursement - Checking For Duplicates.xlsx"
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        OutfileNameXLSX, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
End Sub

'TODO add this for the end, based off summary sheet'
' Sub MessageDuplicate(DuplicateCount)
'     If DuplicateCount = 0 Then
'         MsgBox "No duplicates."
'     Else
'         MsgBox "Number of duplicates: " & DuplicateCount
'     End If
' End Sub
