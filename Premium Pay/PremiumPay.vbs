Sub A_Premium_Pay()
Call DisableEvents
Call CopyRawData
Call UpdateRaw
Call FindAndDeleteDuplicatesMac
Call UpdateSummary
Call UpdateUpload
Call DiscrepancyTab
Call EnableEvents
CheckDate = InputBox("Enter Check Date as MMDDYYYY", "Check Date", "mmddyyyy")
ActiveWorkbook.SaveAs _
  ThisWorkbook.Path & _
  Application.PathSeparator & _
  "Premium Pay - Check Date " & _
  CheckDate & _
  ".xlsx", _
  FileFormat:=xlOpenXMLWorkbook, _
  CreateBackup:=False
End Sub


Sub CopyRawData()
InputFile = "rawdata.csv"
OutputFile = ActiveWorkbook.Name

Workbooks.Open Filename:=ThisWorkbook.Path & Application.PathSeparator & InputFile
'csv will only have one tab, so I can refer to worksheet by index number instead of name'
Workbooks(InputFile).Worksheets(1).Cells.Copy _
    Destination:=Workbooks(OutputFile).Worksheets("Raw Data").Range("A1")
Application.CutCopyMode = False 'Clears clipboard
Workbooks(InputFile).Close SaveChanges:=False
End Sub


Sub UpdateRaw()
Sheets("Raw Data").Select
Range("A1").EntireColumn.Insert
Dim LastRow As Long
  LastRow = Worksheets("Raw Data").Range("C" & Rows.Count).End(xlUp).Row
Range("A1") = "Employee Number | Hourly Pay Rate"
Range("A2:A" & LastRow).Formula = "=CONCATENATE(RC[4],""|"",RC[14])"
Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
Worksheets("Raw Data").Range("A1:A" & LastRow).Copy _
      Destination:=Worksheets("Upload").Range("A1")

Range("S1") = "Job Code Pay Rate"
Range("S2:S" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-2],Rates!C[-18]:C[-17],2,0)"
Range("S2:S" & LastRow).Value = Range("S2:S" & LastRow).Value
End Sub


Sub FindAndDeleteDuplicatesMac()
' Find duplicates
Dim LastRow As Long
  LastRow = Worksheets("Raw Data").Range("C" & Rows.Count).End(xlUp).Row
Sheets("Upload").Select
Dim matchFoundIndex As Long
Dim iCntr As Long
  For iCntr = 1 To LastRow
  If Cells(iCntr, 1) <> "" Then
  matchFoundIndex = WorksheetFunction.Match(Cells(iCntr, 1), Range("A1:A" & LastRow), 0)
    If iCntr <> matchFoundIndex Then
    Cells(iCntr, 2) = "Duplicate!"
    End If
  End If
Next

' Delete duplicates
' TODO fix this part, the macro doesn't know what to do if there are no duplicates
Dim delRange As Range
With ThisWorkbook.Sheets("Upload")
    .AutoFilterMode = False
    With .Range("A1:B" & LastRow)
        .AutoFilter Field:=2, Criteria1:="<>"
        Set delRange = .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    End With
    .AutoFilterMode = False
End With
If Not delRange Is Nothing Then delRange.Delete
End Sub


Sub UpdateSummary()
Dim LastRow As Long
  LastRow = Worksheets("Upload").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Summary").Select
Range("B2").Formula = "=VLOOKUP($A2,'Raw Data'!$A:$E,5,0)"
Range("C2").Formula = "=VLOOKUP($A2,'Raw Data'!$A:$E,4,0)"
Range("D2").Formula = "=SUMIF('Raw Data'!$A:$A,Summary!$A2,'Raw Data'!K:K)"
Range("E2").Formula = "=SUMIF('Raw Data'!$A:$A,Summary!$A2,'Raw Data'!L:L)"
Range("F2").Formula = "=SUMIF('Raw Data'!$A:$A,Summary!$A2,'Raw Data'!M:M)"
Range("H2").Formula = "=VLOOKUP($A2,'Raw Data'!$A:$O,15,0)"
Range("I2").Formula = "=VLOOKUP($A2,'Raw Data'!$A:$S,19,0)"
Range("A2:M2").Select
Selection.AutoFill Destination:=Range("A2:M" & LastRow)
Range("A1:B" & LastRow).Value = Range("A1:B" & LastRow).Value
End Sub

Sub UpdateUpload()
Dim LastRow As Long
  LastRow = Worksheets("Summary").Range("A" & Rows.Count).End(xlUp).Row
Worksheets("Summary").Range("B1:B" & LastRow).Copy _
      Destination:=Worksheets("Upload").Range("A1")

Call FindAndDeleteDuplicatesMac

Dim LastRowUpload As Long
  LastRow = Worksheets("Upload").Range("A" & Rows.Count).End(xlUp).Row
Range("B1") = "Earning or Deduction"
Range("C1") = "Code"
Range("D1") = "Number of Hours"
Range("E1") = "Gross Earnings"
Range("F1") = "Pay Rate"
Range("B2:C" & LastRow) = "E"
Range("E2:E" & LastRow).Formula = "=TEXT(SUMIF(Summary!C[-3],Upload!RC[-4],Summary!C[8]),""0.00"")"
Range("A1:E" & LastRow).Value = Range("A1:E" & LastRow).Value
Sheets("Upload").Name = "Discrep To Upload"

Cells.Select
Selection.AutoFilter Field:=5, Criteria1:="0"
ActiveSheet.AutoFilter.Range.Offset(1, 0).Delete
ActiveSheet.AutoFilterMode = False
End Sub

Sub DiscrepancyTab()
Sheets("Discrepancy").Select
Initials = InputBox("Your Initials", "Initials")
TransmissionDate = InputBox("Transmission Date as mm/dd/yyyy", "Transmission Date", "mm/dd/yyyy")
CheckDate = InputBox("Check Date as mm/dd/yyyy", "Check Date", "mm/dd/yyyy")

Range("A2") = Initials
Range("B2") = TransmissionDate
Worksheets("Discrep To Upload").Range("A:A").Copy _
      Destination:=Worksheets("Discrepancy").Range("C1")
Range("H2").Formula = "=CONCATENATE(VLOOKUP(RC[-5],'Discrep To Upload'!C[-7]:C[-3],5,0), "" E"")"
Range("P2") = CheckDate
Range("S2") = CheckDate
Dim LastRow As Long
  LastRow = Worksheets("Discrep To Upload").Range("A" & Rows.Count).End(xlUp).Row
Range("A2:B" & LastRow).FillDown
Range("D2:T2").Select
Selection.AutoFill Destination:=Range("D2:T" & LastRow)
Range("P2:P" & LastRow).FillDown
Range("S2:S" & LastRow).FillDown
Range("A1:T" & LastRow).Value = Range("A1:T" & LastRow).Value
End Sub


Sub DisableEvents()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
End Sub


Sub EnableEvents()
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub
