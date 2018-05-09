Public LastRow as Long

Sub AuditDirectDeposits()
'########################################################################'
'##Compares the direct deposit reports between Salesforce and Paylocity##'
'########################################################################'


Dim MainWorkbookName As String
Dim SalesforceReportPath As String
Dim SalesforceReportName As String
Dim PaylocityReportPath As String
Dim PaylocityReportName As String


'##################################'
'Gathers all the data that we need.'
'##################################'
MainWorkbookName = ActiveWorkbook.Name

Sheets(1).Name = "Main"
With Sheets
    .Add().Name = "Salesforce Full Data"
    .Add().Name = "Paylocity Full Data"
End With

MsgBox "Select the Salesforce report."
SalesforceReportPath = Application.GetOpenFilename
MsgBox "Select the Paylocity report."
PaylocityReportPath = Application.GetOpenFilename

'Copy over the Salesforce information to the main workbook'
Workbooks.Open Filename:= SalesforceReportPath
SalesforceReportName = ActiveWorkbook.Name
Workbooks(SalesforceReportName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MainWorkbookName).Worksheets("Salesforce Full Data").Range("A1")
Workbooks(SalesforceReportName).Close SaveChanges:=False

'Copy over the Paylocity information to the main workbook'
Workbooks.Open Filename:= PaylocityReportPath
PaylocityReportName = ActiveWorkbook.Name
Workbooks(PaylocityReportName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MainWorkbookName).Worksheets("Paylocity Full Data").Range("A1")
Workbooks(PaylocityReportName).Close SaveChanges:=False


'###################################'
'Formatting the data from Paylocity.'
'###################################'
Worksheets("Paylocity Full Data").Activate
Call UnformatSheet()

'Paylocity has blank rows above headers, delete those first'
Do While IsEmpty(Range("A1").Value)
    Rows(1).EntireRow.Delete
Loop

Call FindLastRow("B")
Range("A1:B" & LastRow).SpecialCells(xlCellTypeBlanks).Select
Selection.FormulaR1C1 = "=R[-1]C"
End Sub

' TODO up to the point where I start moving columns around

Sub UnformatSheet()
ActiveWindow.DisplayGridlines = True
    With ActiveSheet
        .AutoFilterMode = False
        .Cells.WrapText = False
        .Cells.UnMerge
        .Cells.EntireRow.Hidden = False
        .Cells.EntireColumn.Hidden = False
    End With
End Sub

Sub FindLastRow(ColumnLetter)
    LastRow = Range(ColumnLetter & Rows.Count).End(xlUp).Row
End Sub
