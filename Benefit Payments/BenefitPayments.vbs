Sub BenefitPayments()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Takes the Total Deduction file and create an uploadable'
'file to Salesforce for the benefit payments.           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim MyWorkbookName As String
Dim TotalDeductionsName As String
Dim BenefitIDName As String
Dim lngLastRow As Long
Dim BenefitHeaders As Variant
Dim UploadHeaders As Variant
Dim CheckDate As String
Dim OutfileNameCSV As String
Dim OutfileNameXLSX As String

MyWorkbookName = ActiveWorkbook.Name

MsgBox "Select the Total Deductions file."
TotalDeductionsFile = Application.GetOpenFilename
MsgBox "Select the Benefit IDs file."
BenefitIDFile = Application.GetOpenFilename

'Copy over the Total Deductions
Workbooks.Open Filename:= TotalDeductionsFile
TotalDeductionsName = ActiveWorkbook.Name
Workbooks(TotalDeductionsName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MyWorkbookName).Worksheets("Benefit Payments").Range("A1")
Workbooks(TotalDeductionsName).Close SaveChanges:=False

'Copy over the Benefit IDs
Workbooks.Open Filename:= BenefitIDFile
BenefitIDName = ActiveWorkbook.Name
Workbooks(BenefitIDName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MyWorkbookName).Worksheets("Election ID").Range("A1")
Workbooks(BenefitIDName).Close SaveChanges:=False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Format Benefit Payments sheet to find just benefit codes'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Benefit Payments").Activate

BenefitHeaders = Array( _
    "Employee Number", _
    "Employee", _
    "Check Date", _
    "Code", _
    "Description", _
    "Amount", _
    "Process ID", _
    "A Benefit Code?" _
)

Range("A1").EntireRow.Insert
Range("A1:H1") = BenefitHeaders

lngLastRow = Range("A" & Rows.Count).End(xlUp).Row
Range("H2:H" & lngLastRow).FormulaR1C1 = "=COUNTIF(BenefitList,RC[-4])"
Range("$A$1:H" & lngLastRow).AutoFilter Field:=8, Criteria1:="1"
Cells.Copy Destination:= Worksheets("Upload Template").Range("A1")
Application.CutCopyMode = False

''''''''''''''''''''''''
'Format Upload Template'
''''''''''''''''''''''''
Sheets("Upload Template").Activate
lngLastRow = Range("A" & Rows.Count).End(xlUp).Row

Range("A1").EntireColumn.Insert
Range("A1:A" & lngLastRow).FormulaR1C1 = "=CONCATENATE(RC[1],"" - "",RC[4])"

Range("A1").EntireColumn.Insert
Range("A1") = "Benefit ID"
Range("A2:A" & lngLastRow).FormulaR1C1 = "=IFNA(VLOOKUP(RC[1],'Election ID'!C:C[1],2,0),RC[1])"

Range("A1:B" & lngLastRow).Value = Range("A1:B" & lngLastRow).Value

'Insert 5 columns for the upload file'
Range("A1:E1").EntireColumn.Insert
Range("A2:A" & lngLastRow).Value = Range("F2:F" & lngLastRow).Value
Range("B2:B" & lngLastRow).Value = Range("M2:M" & lngLastRow).Value
Range("C2:C" & lngLastRow).Value = Range("J2:J" & lngLastRow).Value
Range("D2:D" & lngLastRow).FormulaR1C1 = _
    "=IF(RC[-2]>0,""Payment"",""SIS Reimbursing Employee"")"
Range("E2:E" & lngLastRow).FormulaR1C1 = _
    "=CONCATENATE(""Payment through transmission check date "",TEXT(RC[-1],""mm/dd/yyyy.""))"

UploadHeaders = Array( _
    "Benefit Election", _
    "Amount", _
    "Check Date", _
    "Transaction Type", _
    "Check" _
)

Range("A1:E1") = UploadHeaders
Columns("F:Z").EntireColumn.Delete

Range("A1").CurrentRegion.Sort _
    Key1:=Range("A1"), Order1:=xlAscending, _
    Header:=xlYes

Cells.EntireColumn.AutoFit


''''''''''''''''''''''
'Save the upload file'
''''''''''''''''''''''
CheckDate = Format(Range("C2"), "mmddyyyy")
OutfileNameCSV = "Benefit Payments - Check Date " & CheckDate & " - Upload.csv"
OutfileNameXLSX = "Benefit Payments - Check Date " & CheckDate & ".xlsx"
ActiveWorkbook.SaveAs _
    ThisWorkbook.Path & _
    Application.PathSeparator & _
    OutfileNameCSV, _
    FileFormat:=xlCSV, _
    CreateBackup:=False

ActiveWorkbook.SaveAs _
    ThisWorkbook.Path & _
    Application.PathSeparator & _
    OutfileNameXLSX, _
    FileFormat:=xlOpenXMLWorkbook, _
    CreateBackup:=False

End Sub
