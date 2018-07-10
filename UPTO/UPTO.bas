Public CheckDate As String
Public LastRow As Long
Public MainWbName As String
Public SourceFilename As String
Public SourceWbName As String
Public TransmissionDate As String

Sub createUnusedPTORecords()
    Call Alerts(False)

    MainWbName = ActiveWorkbook.Name
    CheckDate = InputBox( _
        "Check date we're processing?","Check Date","m/d/yyyy" _
    )
    TransmissionDate = InputBox( _
        "What was the transmission date?","Transmission Date","m/d/yyyy" _
    )

    Call LoadFile("UPTO")
    Call LoadFile("Employees - All SIS Employees")
    LastRow = Worksheets("UPTO").Cells(Rows.Count, 1).End(xlUp).Row + 1

    Worksheets("Main").Activate
    Call FillData("Employee", "A", "=VLOOKUP(UPTO!$A1,'Employees - All SIS Employees'!$A:$B,2,0)")
    Call FillData("Cash Out", "B", "TRUE")
    Call FillData("Notes_from_Payroll__c", "C", "Pay out unused accrual balance.")
    Call FillData("Hours - Paid", "D", "='UPTO'!$D1")
    Call FillData("Hours - Total Paid", "E", "='UPTO'!$D1")
    Call FillData("Transmission Date", "F", TransmissionDate)
    Call FillData("Use PTO","G", "TRUE")
    Call FillData("Approval Status", "H", "Approved")
    Call FillData("Record Type ID", "I", "012A0000000GjjL")
    Call FillData("Update Pending Balance", "J", "TRUE")
    Call FillData("Date", "K", "=VLOOKUP(UPTO!$A1,'UPTO Tab'!$C:$M,11,0)")
        Columns("K").NumberFormat = "m/d/yyyy"
    Call FillData("Processed by Payroll", "L", "TRUE")
    Call FillData("Check Date", "M", "=VLOOKUP(UPTO!$A1,'UPTO Tab'!$C:$M,11,0)")
        Columns("M").NumberFormat = "m/d/yyyy"

    Call Alerts(True)

    Worksheets("Main").Activate
    ActiveWorkbook.SaveAs _
    ThisWorkbook.Path & Application.PathSeparator & _
    "UPTO - Main Workbook - Check Date " & _
    Format(CheckDate, "mmddyyyy") & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ActiveWorkbook.SaveAs _
    ThisWorkbook.Path & Application.PathSeparator & _
    "UPTO Update - Check Date " & _
    Format(CheckDate, "mmddyyyy") & ".csv", _
    FileFormat:=xlCSVMSDOS, CreateBackup:=False
End Sub

Public Sub Alerts(trueFalse)
    Application.DisplayAlerts = trueFalse
    Application.ScreenUpdating = trueFalse
    Application.EnableEvents = trueFalse
End Sub

Public Sub FillData(Header, ColumnLetter, Formula)
    Range(ColumnLetter & "1") = Header
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow) = Formula
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub

Public Sub LoadFile(Filename)
    MsgBox "Select the " & Filename & " file."
    SourceFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=SourceFilename
    SourceWbName = ActiveWorkbook.Name
    Workbooks(SourceWbName).Sheets(1).Name = Filename
    Workbooks(SourceWbName).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)
End Sub
