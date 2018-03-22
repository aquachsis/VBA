Sub CES()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Takes the CES report from Paylocity and formats it in a way'
'where it provides the summarized information that the BLS '
'needs.                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox _
  "Choose the CES Raw Data File" & _
  vbNewLine & vbNewLine & _
  "Does not work on Mac because of RemoveDuplicates"

Dim strStatisticsReportLocation As String
strStatisticsReportLocation = Application.GetOpenFilename
Workbooks.Open Filename:= strStatisticsReportLocation

'Renames existing sheet for better readability
Sheets(1).Name = "Raw Data"

'Add all sheets that we'll need
With Sheets
    .Add().Name = "EE Count"
    .Add().Name = "Dollar"
    .Add().Name = "Hours"
    .Add().Name = "Summary"
End With

'Delete Raw Data's footer
Worksheets("Raw Data").Activate
Dim lngFooterRow As Long
lngFooterRow = ActiveSheet.Cells(Rows.Count,2).End(xlUp).Row
Rows(lngFooterRow & ":" & Rows.Count).EntireRow.Delete

'Employee Count
Worksheets("EE Count").Activate
Sheets("Raw Data").Columns("A:F").Copy _
    Destination:=Sheets("EE Count").Columns("A:F")
Sheets("EE Count").Columns("A:F").RemoveDuplicates Columns:=Array(3), Header:=xlYes
' TODO make this into an array
Range("H1") = "Employee Count"
Range("H2") = "All"
Range("I2").Formula = "=COUNT(C:C)"
Range("H3") = "Non-Supervisory"
Range("I3").Formula = "=COUNTIF(A:A,""N"")"
Range("H4") = "Women"
Range("I4").Formula = "=COUNTIF(F:F,""F"")"

'Dollar
Worksheets("Dollar").Activate
Sheets("Raw Data").Columns("A:F").Copy _
    Destination:=Sheets("Dollar").Columns("A:F")
    Sheets("Raw Data").Columns("K:K").Copy _
        Destination:=Sheets("Dollar").Columns("G:G")
' TODO make this into an array
Range("I1") = "Payroll, Excluding Commissions"
Range("I2") = "All"
Range("J2").Formula = "=SUM(G:G)"
Range("J2").Style = "Currency"
Range("I3") = "Non-Supervisory"
Range("J3").Formula = "=SUMIF(A:A,""N"",G:G)"
Range("J3").Style = "Currency"


'Hours
Worksheets("Hours").Activate
Sheets("Raw Data").Columns("A:F").Copy _
    Destination:=Sheets("Hours").Columns("A:F")
    Sheets("Raw Data").Columns("J:J").Copy _
        Destination:=Sheets("Hours").Columns("G:G")
' TODO make this into an array
Range("I1") = "Hours, Including Overtime"
Range("I2") = "All"
Range("J2").Formula = "=SUM(G:G)"
Range("J2").Style = "Comma"
Range("I3") = "Non-Supervisory"
Range("J3").Formula = "=SUMIF(A:A,""N"",G:G)"
Range("J3").Style = "Comma"


'Summary
' TODO make this into an array
Worksheets("Summary").Activate
Dim SummaryArray As Variant
    SummaryArray = Array( _
        "", _
        "Employee Count", _
        "Women Workers", _
        "Payroll, Excluding Commissions (Whole dollars)", _
        "Commissions Paid at Least Once a Month (Whole dollars)", _
        "Hours, Including Overtime (Whole hours)" _
    )
Range("A1:F1") = SummaryArray
Range("A2") = "All Workers"
Range("B2").Formula = Worksheets("EE Count").Range("I2").Value
Range("C2").Formula = Worksheets("EE Count").Range("I4").Value
Range("D2").Formula = Worksheets("Dollar").Range("J2").Value
Range("F2").Formula = Worksheets("Hours").Range("J2").Value
Range("A3") = "Nonsupervisory Workers"
Range("B3").Formula = Worksheets("EE Count").Range("I3").Value
Range("D3").Formula = Worksheets("Dollar").Range("J3").Value
Range("F3").Formula = Worksheets("Hours").Range("J3").Value

Range("A1:F3").Select
With Selection
.ColumnWidth = 20
.NumberFormat = "#,###"
.HorizontalAlignment = xlLeft
.VerticalAlignment = xlBottom
.WrapText = True
End With

ActiveWorkbook.Save

End Sub
