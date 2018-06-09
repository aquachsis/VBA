Option Explicit
Public BenefitPolicyID As String
Public CheckDate As String
Public CompanyID As String
Public ContactIDFilename As String
Public ContactIDVLookup As String
Public ContactIDWbName As String
Public ElectionStatus As String
Public EmployeeID As String
Public LastRow As Long
Public MainWbName As String
Public PayPeriodStart As String
Public RecordTypeID As String

Public Sub CreateBenefitElection()
    BenefitPolicyID = "a2wA0000002TWI7"
    CompanyID = "12345"
    ContactIDVLookup = "=VLOOKUP('Employee Number'!$A2,'Contact ID'!$A:$B,2,0)"
    ElectionStatus = "Accepted"
    EmployeeID = "='Employee Number'!$A2"
    PayPeriodStart = InputBox("First date of the pay period?","Effective Date","m/d/yyyy")
    CheckDate = InputBox("Check date we're processing?","Check Date","m/d/yyyy")
    MainWbName = ActiveWorkbook.Name
    RecordTypeID = "012A0000000ra7r"

    'Find Last Row'
    Worksheets("Employee Number").Activate
    LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    'Copy Contact ID to main workbook'
    MsgBox "Select the Contact ID (Employees - All SIS Employees) report."
    ContactIDFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=ContactIDFilename
    ContactIDWbName = ActiveWorkbook.Name
    Workbooks(ContactIDWbName).Sheets(1).Name = "Contact ID"
    Workbooks(ContactIDWbName).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)

    'Fill Salesforce sheet
    Worksheets("Salesforce").Activate
    Call FillInData("A", ContactIDVLookup)
    Call FillInData("B", PayPeriodStart)
    Call FillInData("C", RecordTypeID)
    Call FillInData("D", BenefitPolicyID)
    Call FillInData("E", ElectionStatus)

    'Fill Paylocity sheet
    Worksheets("Paylocity").Activate
    Call FillInData("A", CompanyID)
    Call FillInData("B", EmployeeID)
    Call FillInData("C", "D")
    Call FillInData("D", "LICWA")
    Call FillInData("E", CheckDate)
    Call FillInData("F", "22.75")
    Call FillInData("G", "W")
    Call FillInData("H", "91")
    Call FillInData("I", "0")
    Call FillInData("J", "Flat")
    Call FillInData("K", PayPeriodStart)

    ' TODO Save Salesforce sheet as CSV and Palocity sheet (on its own as XLSX)
    'also SAVE AS this template as ORCA - Set Up - Check Date MMDDYYYY'
End Sub

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub
