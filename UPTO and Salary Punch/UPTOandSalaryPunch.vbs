'Button1 is BigMac()'
'Button2 is SalPunch2()'

Public DSTOnOff As Single
Public DSTFormula As String

Sub BigMac()
Call SalPunch1
Call UPTO
End Sub



Sub SalPunch1()
' 1st part of SAL PUNCH
' Keyboard Shortcut: Ctrl+Shift+D


''''''''''''''''''''''''''''''''''''''''''''
'ARE WE IN DAYLIGHT SAVINGS TIME?          '
' 1 = Yes, we're currently observing DST   '
' 0 = No, we're not currently observing DST'
''''''''''''''''''''''''''''''''''''''''''''
DSTOnOff = 1 '1 or 0 depending if we're observing daylight savings

If DSTOnOff = 1 Then
    DSTFormula = "=CONCATENATE(TEXT(RC[3],""MM/DD/YYYY""),""T07:00:00.000Z"")"
Else
    DSTFormula = "=CONCATENATE(TEXT(RC[3],""MM/DD/YYYY""),""T08:00:00.000Z"")"
End If


Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Sheets("Instructions").Select

'Defines first three variables for SalPunch

Dim PayPeriodStart As Variant
PayPeriodStart = InputBox("Pay period start in this format: MM/DD/YYYY", _
  "Pay Period Start", "mm/dd/yyyy")
Dim PayPeriodEnd As Variant
PayPeriodEnd = InputBox("Pay period end in this format: MM/DD/YYYY", _
  "Pay Period End", "mm/dd/yyyy")

Range("O9") = PayPeriodStart
Range("O11") = PayPeriodEnd

'Defines 2 varaibles used by 2nd macro.
Dim TransmissionDate As Variant
TransmissionDate = InputBox("Transmission date in this format: MM/DD/YYYY", _
  "Transmission Date", "mm/dd/yyyy")
Dim CheckDate As Variant
CheckDate = InputBox("Check date in this format: MM/DD/YYYY", _
  "Check Date", "mm/dd/yyyy")

Range("O3") = TransmissionDate
Range("O7") = CheckDate

Range("O5").Select
   ActiveCell.FormulaR1C1 = "=TEXT(R[2]C,""MMDDYYYY"")"
   Range("O5").Select
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False

Workbooks.Open Filename:=ThisWorkbook.Path & "\Paylocity Salary Hours.csv"


Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Paylo Sal Hrs").Select
Windows("Paylocity Salary Hours.CSV").Activate
Sheets("Paylocity Salary Hours").Select
Cells.Select
Selection.Copy
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Paylo Sal Hrs").Select
Range("A1").Select
ActiveSheet.Paste
Windows("Paylocity Salary Hours.CSV").Activate
ActiveWindow.Close

Workbooks.Open Filename:=ThisWorkbook.Path & "\All SIS Employees.csv"


Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("SIS All Emp").Select
Windows("All SIS Employees.csv").Activate
Sheets("All SIS Employees").Select
Cells.Select
Selection.Copy
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("SIS All Emp").Select
Range("A1").Select
ActiveSheet.Paste
Sheets("Refresh PTO Balance").Select
Range("A1").Select
ActiveSheet.Paste
Windows("All SIS Employees.csv").Activate
ActiveWindow.Close

Sheets("Paylo Sal Hrs").Select
Columns("D:D").Select
Selection.Cut
Columns("A:A").Select
Selection.Insert Shift:=xlToRight
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B1").FormulaR1C1 = "Total"

Dim Lastrow As Long
Application.ScreenUpdating = False
Lastrow = Range("A" & Rows.Count).End(xlUp).Row


Cells.Find(What:="E-REGHours", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1).Select

    ActiveCell.EntireColumn.Select

    Selection.Cut
Columns("C:C").Select
Selection.Insert Shift:=xlToRight

Cells.Find(What:="E-PTOHours", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1).Select

ActiveCell.EntireColumn.Select

Selection.Cut
Columns("D:D").Select
Selection.Insert Shift:=xlToRight
Range("B2").FormulaR1C1 = "=SUM(RC[1]:RC[2])"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & Lastrow)

Columns("B:B").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("C:C").Select
Range(Selection, Selection.End(xlToRight)).Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Range("C2").Select

Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B2").Select
ActiveCell.FormulaR1C1 = _
    "=VLOOKUP(RC[-1],'SIS All Emp'!C[-1]:C,2,0)"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & Lastrow)
Range("B1").FormulaR1C1 = "Employee"
Columns("B:B").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("A:A").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

Range("B1") = "Record Type ID"
Range("B2") = "012A00000019lrE"

Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & Lastrow)
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C1") = "In"
Range("D1") = "Out"

Sheets("Instructions").Select
Range("O9").Select
Selection.Copy
Sheets("Paylo Sal Hrs").Select
Range("F2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C2").FormulaR1C1 = DSTFormula

Range("C2").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("F2").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & Lastrow)

Sheets("Instructions").Select
Range("O11").Select
Selection.Copy
Sheets("Paylo Sal Hrs").Select
Range("F2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("D2").FormulaR1C1 = DSTFormula
Range("D2").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("F2").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & Lastrow)

Dim sFilename As String
sFilename = "Salary Hours - Salesforce Upload.csv" 'You can give a nem to save
Workbooks.Add
'Saving the Workbook
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & sFilename, FileFormat:=xlCSV, CreateBackup:=False


Windows("Salary Hours - Salesforce Upload.csv").Activate
Range("G17").Select
Windows("UPTO and Salary Punch - Macro"). _
    Activate
Columns("A:E").Select
Selection.Copy
Windows("Salary Hours - Salesforce Upload.csv").Activate
Range("A1").Select
ActiveSheet.Paste
ActiveWorkbook.Save
Windows("UPTO and Salary Punch - Macro"). _
    Activate
Sheets("Instructions").Select

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.DisplayAlerts = True

End Sub

Sub SalPunch2()
'
' Macro2 Macro
' 2nd part of SAL PUNCH
' Keyboard Shortcut: Ctrl+t

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False



Workbooks.Open Filename:=ThisWorkbook.Path & "\Salary Hours - Salesforce Upload - Success.csv"

Application.ScreenUpdating = True
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Sal Hrs Success").Select
Windows("Salary Hours - Salesforce Upload - Success.csv").Activate
Cells.Select
Selection.Copy
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Sal Hrs Success").Select
Range("A1").Select
ActiveSheet.Paste
Windows("Salary Hours - Salesforce Upload - Success.csv").Activate
ActiveWindow.Close

Sheets("Sal Hrs Success").Select
Range("D7").Select
Cells.Find(What:="ID", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    , SearchFormat:=False).Activate

ActiveCell.EntireColumn.Select

Selection.Copy
Columns("A:A").Select
ActiveSheet.Paste

Dim Lastrow As Long
Application.ScreenUpdating = False
Lastrow = Range("A" & Rows.Count).End(xlUp).Row

Columns("B:B").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft
Range("B1").Select
ActiveCell.FormulaR1C1 = "Check Date"
Range("B2").Select
Sheets("Instructions").Select
Range("O7").Select
Selection.Copy
Sheets("Sal Hrs Success").Select
Range("B2").Select
ActiveSheet.Paste


Range("B2").Select
Selection.Copy
Range("B3").Select
ActiveSheet.Paste
Range("B2:B3").Select
Selection.AutoFill Destination:=Range("B2:B" & Lastrow)

Dim sFilename As String
sFilename = "Lock Payroll Records.csv" 'You can give a nem to save
Workbooks.Add
'Saving the Workbook
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & sFilename, FileFormat:=xlCSV, CreateBackup:=False


Windows("Lock Payroll Records.csv").Activate
Range("G17").Select
Windows("UPTO and Salary Punch - Macro").Activate
Columns("A:B").Select
Selection.Copy
Windows("Lock Payroll Records.csv").Activate
Range("A1").Select
ActiveSheet.Paste
ActiveWorkbook.Save

Windows("UPTO and Salary Punch - Macro").Activate
Sheets("Refresh PTO Balance").Select

 'This part creates the Refresh PTO Balance Upload file
Dim LR As Long
LR = Range("B" & Rows.Count).End(xlUp).Row

Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Columns("B:C").Select
Selection.Delete Shift:=xlToLeft
Range("B1").Select
ActiveCell.FormulaR1C1 = "Refresh PTO Accrued"
Range("B2").Select
ActiveCell.FormulaR1C1 = "TRUE"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & LR)


Dim Filename As String
Filename = "Refresh All Employee PTO Balance.csv" 'You can give a name to save
Workbooks.Add
'Saving the Workbook
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & Filename, FileFormat:=xlCSV, CreateBackup:=False

Windows("Refresh All Employee PTO Balance.csv").Activate
Windows("UPTO and Salary Punch - Macro").Activate
Sheets("Refresh PTO Balance").Select
Columns("A:B").Select
Selection.Copy
Windows("Refresh All Employee PTO Balance.csv").Activate
Range("A1").Select
ActiveSheet.Paste
ActiveWorkbook.Save

Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Instructions").Select

End Sub
Sub UPTO()
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+D
'

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Instructions").Select
Range("O5").Select
Dim CheckDate
CheckDate = ActiveCell.Value


Workbooks.Open Filename:=ThisWorkbook.Path & "\UPTO - Paylocity.csv"


Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("UPTO").Select
Windows("UPTO - Paylocity.csv").Activate
Cells.Select
Selection.Copy
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("UPTO").Select
Range("A1").Select
ActiveSheet.Paste
Windows("UPTO - Paylocity.csv").Activate
ActiveWindow.Close

Workbooks.Open Filename:=ThisWorkbook.Path & "\All SIS Employees.csv"


Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("SIS All Emp").Select
Windows("All SIS Employees.csv").Activate
Cells.Select
Selection.Copy
Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("SIS All Emp").Select
Range("A1").Select
ActiveSheet.Paste
Windows("All SIS Employees.csv").Activate
ActiveWindow.Close

Sheets("UPTO").Select
Range("A1").Select
Rows(1).Insert Shift:=xlShiftDown

[A:A].Select
  With Selection
      .NumberFormat = "General"
      .Value = .Value
  End With

Sheets("SIS All Emp").Select

[A:A].Select
  With Selection
      .NumberFormat = "General"
      .Value = .Value
  End With



Sheets("UPTO Template").Select
Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Sheets("UPTO").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("UPTO Template").Select
Range("A2").Select
ActiveSheet.Paste

Dim Lastrow As Long
Application.ScreenUpdating = False
Lastrow = Range("A" & Rows.Count).End(xlUp).Row

Range("B2:B" & Lastrow).Formula = "=VLOOKUP(RC[-1],'SIS All Emp'!C[-1]:C[2],2,0)"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & Lastrow)
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & Lastrow)

Sheets("UPTO").Select
Columns("D:D").Select
Selection.Copy
Sheets("UPTO Template").Select
Columns("E:E").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
Range("E1").Select
ActiveCell.FormulaR1C1 = "Hours - Paid"


Columns("F:F").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
Range("F1").Select
ActiveCell.FormulaR1C1 = "Hours - Total Paid"


Range("G2").Select

Sheets("Instructions").Select

Range("O3").Select
Selection.Copy
Sheets("UPTO Template").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("G3").Select

Range("H2:J2").Select
Selection.AutoFill Destination:=Range("H2:J" & Lastrow)

Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K" & Lastrow)

Range("L2").Select
ActiveCell.FormulaR1C1 = ""
Range("L2").Select
ActiveCell.FormulaR1C1 = _
    "=VLOOKUP(RC[-11],'TCL Vac Hrs - PASTE DATA HERE'!C[-9]:C[1],11,0)"
Range("L2").Select
Selection.AutoFill Destination:=Range("L2:L" & Lastrow)

Columns("L:L").Select
Selection.NumberFormat = "m/d/yyyy"
Range("M2").Select
Selection.AutoFill Destination:=Range("M2:M" & Lastrow)

Range("L2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
ActiveWindow.SmallScroll Down:=-15
Range("N2").Select
ActiveSheet.Paste
Range("N2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("N:N").Select
Application.CutCopyMode = False
Selection.NumberFormat = "m/d/yyyy"
Range("O10").Select

Range("G2").Select
Selection.Copy
Range("G3").Select
ActiveSheet.Paste
Range("G2:G3").Select
Application.CutCopyMode = False
Selection.AutoFill Destination:=Range("G2:G" & Lastrow)

Sheets("UPTO Template").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("A:A").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Columns("K:K").Select
ActiveWorkbook.Worksheets("UPTO Template").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("UPTO Template").Sort.SortFields.Add Key:=Range( _
    "K1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("UPTO Template").Sort
    .SetRange Columns("A:M")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With



Dim sFilename As String
sFilename = "UPTO Update " & CheckDate & ".csv" 'You can give a name to save
Workbooks.Add
'Saving the Workbook
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & sFilename, FileFormat:=xlCSV, CreateBackup:=False

Columns("A:M").Select
Selection.Copy
Windows("UPTO Update " & CheckDate & ".csv").Activate
Range("A1").Select
ActiveSheet.Paste
Range("A1").Select
Application.CutCopyMode = False
ActiveWorkbook.Save

Windows("UPTO Update " & CheckDate & ".csv").Activate
Range("G17").Select
Windows("UPTO and Salary Punch - Macro").Activate
Columns("A:M").Select
Selection.Copy
Windows("UPTO Update " & CheckDate & ".csv").Activate
Range("A1").Select
ActiveSheet.Paste
ActiveWorkbook.Save

Windows("UPTO and Salary Punch - Macro.xlsm").Activate
Sheets("Instructions").Select

Windows("UPTO Update " & CheckDate & ".csv").Activate
Range("A1").Select

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.DisplayAlerts = True

End Sub
