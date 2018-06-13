Attribute VB_Name = "DirectDepositAudit"

Public LastRow as Long

Sub a_AuditDirectDeposits()
'Run this one. It's the main one.
    Call LoadFiles
    Call FormatPaylocityData
    Call FormatSalesforceData
End Sub

Sub LoadFiles()
Dim MainWorkbookName As String
Dim SalesforceReportPath As String
Dim SalesforceReportName As String
Dim PaylocityReportPath As String
Dim PaylocityReportName As String

'Gathers all the data that we need.
MainWorkbookName = ActiveWorkbook.Name

Sheets(1).Name = "Main"
With Sheets
    .Add().Name = "Salesforce"
    .Add().Name = "Paylocity"
End With

MsgBox "Select the Salesforce report."
SalesforceReportPath = Application.GetOpenFilename
MsgBox "Select the Paylocity report."
PaylocityReportPath = Application.GetOpenFilename

'Copy over the Salesforce information to the main workbook'
Workbooks.Open Filename:= SalesforceReportPath
SalesforceReportName = ActiveWorkbook.Name
Workbooks(SalesforceReportName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MainWorkbookName).Worksheets("Salesforce").Range("A1")
Workbooks(SalesforceReportName).Close SaveChanges:=False

'Copy over the Paylocity information to the main workbook'
Workbooks.Open Filename:= PaylocityReportPath
PaylocityReportName = ActiveWorkbook.Name
Workbooks(PaylocityReportName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MainWorkbookName).Worksheets("Paylocity").Range("A1")
Workbooks(PaylocityReportName).Close SaveChanges:=False
End Sub

Sub FormatPaylocityData()
Worksheets("Paylocity").Activate
Call UnformatSheet()

Call FindLastRow("B")
'Fills in employee number.
'Paylocity only shows employee number once, even if there are multiple accounts.
Range("A1:B" & LastRow).SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
Range("A1:B" & LastRow).Value = Range("A1:B" & LastRow).Value

Columns(1).Insert
Range("A1") = "Employee ID | Routing | Account | Type | Order"
Call FillInData("A","=CONCATENATE($C2,""|"",$F2,""|"",$G2,""|"",$H2,""|"",$D2)")
End Sub

Sub FormatSalesforceData()
Worksheets("Salesforce").Activate
Call UnformatSheet()
Call FindLastRow("B")
Range("B" & LastRow).Value = Range("B" & LastRow).Value

Columns(1).Insert
Range("A1") = "Employee ID | Routing | Account | Type | Order"
Call FillInData("A","=CONCATENATE($B2,""|"",$G2,""|"",$H2,""|"",$J2,""|"",$I2)")
End Sub


Public Sub UnformatSheet()
    ActiveWindow.DisplayGridlines = True
    With ActiveSheet
        .AutoFilterMode = False
        .Cells.WrapText = False
        .Cells.UnMerge
        .Cells.EntireRow.Hidden = False
        .Cells.EntireColumn.Hidden = False
    End With
    Do While IsEmpty(Range("A1").Value)
        Rows(1).EntireRow.Delete
    Loop
End Sub

Public Sub FindLastRow(ColumnLetter)
    LastRow = Range(ColumnLetter & Rows.Count).End(xlUp).Row
End Sub

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    ' Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
    '     Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub
