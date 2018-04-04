Sub BenefitPayments()
' Keyboard Shortcut: Ctrl + D
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

'this opens up the two files and copies them into the macro

Workbooks.Open Filename:=Application.GetOpenFilename
Windows("Benefit Payments - Report.csv").Activate
Cells.Select
Selection.Copy
Windows("Benefit Payments - After - Macro.xlsm").Activate
Sheets("Benefit Payments").Select
Range("A1").Select
ActiveSheet.Paste
Windows("Benefit Payments - Report.csv").Activate
ActiveWindow.Close

Workbooks.Open Filename:=ThisWorkbook.Path & "\Benefit - Election ID.csv"
Windows("Benefit - Election ID.csv").Activate
Cells.Select
Selection.Copy
Windows("Benefit Payments - After - Macro.xlsm").Activate
Sheets("Election ID").Select
Range("A1").Select
ActiveSheet.Paste
Windows("Benefit - Election ID.csv").Activate
ActiveWindow.Close

 'start of formuals and formatting within macro

Windows("Benefit Payments - After - Macro.xlsm").Activate
Sheets("Benefit Payments").Select
Range("A1").EntireRow.Insert

Dim Lastrow As Long
Lastrow = Range("A" & Rows.Count).End(xlUp).Row

Range("H2").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(BenefitList,RC[-4])"
Range("H2").Select
Selection.AutoFill Destination:=Range("H2:H" & Lastrow)
Rows("1:1").Select
Selection.AutoFilter
ActiveSheet.Range("$A$1:H" & Lastrow).AutoFilter Field:=8, Criteria1:="1"



Range("A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select

Selection.Copy
Sheets("Upload Template").Select
Cells.Select
ActiveSheet.Paste


'This defines last row for Upload Template tab

Rows("1:1").Select
Application.CutCopyMode = False
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Dim LR As Long
LR = Range("A" & Rows.Count).End(xlUp).Row

Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1").Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],"" - "",RC[4])"
Range("A1").Select
Selection.AutoFill Destination:=Range("A1:A" & LR)
Columns("A:A").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

 'This section grab benefit election id
Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A2").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],'Election ID'!C:C[1],2,0)"
Range("A2").Select
Selection.AutoFill Destination:=Range("A2:A" & LR)
Columns("A:A").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

 'this section being to format the report to start creating the upload file
Columns("H:H").Select
Application.CutCopyMode = False
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Columns("F:F").Select
Selection.Cut
Columns("C:C").Select
Selection.Insert Shift:=xlToRight
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("E:E").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1").Select
ActiveCell.FormulaR1C1 = "Benefit Election"
Range("B1").Select
ActiveCell.FormulaR1C1 = "Amount"
Range("C1").Select
ActiveCell.FormulaR1C1 = "Check Date"
Range("D1").Select
ActiveCell.FormulaR1C1 = "Transaction Type"
Range("E1").Select
ActiveCell.FormulaR1C1 = "Check"
Columns("G:G").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft

'This creates the CONCATENATE and pulls records for #NA
Range("D2").Select
ActiveCell.FormulaR1C1 = _
    "=CONCATENATE(""Payment through transmission check date "",TEXT(RC[-1],""mm/dd/yyyy.""))"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & LR)
Columns("D:D").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("A:A").Select
Application.CutCopyMode = False
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A2").Select
ActiveCell.FormulaR1C1 = "=IF(ISNA(RC[1]),RC[6],RC[1])"
Range("A2").Select
Selection.AutoFill Destination:=Range("A2:A" & LR)
Columns("A:A").Select
Selection.Copy
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'This continue to format and includes an If statment stating that if amount is greater than 0 enter "Payment" if not enter "SIS Reimbusing Employee"
Range("E1:F1").Select
Range("F1").Activate
Application.CutCopyMode = False
Selection.ClearContents
Range("E1").Select
ActiveCell.FormulaR1C1 = "Check"
Columns("E:E").Select
Selection.Cut
Columns("G:G").Select
Selection.Insert Shift:=xlToRight
Range("E1").Select
ActiveCell.FormulaR1C1 = "Transaction Type"
Range("E2").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-2]>0,""Payment"",""SIS Reimbursing Employee"")"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & LR)
Columns("E:E").Select
Selection.Copy
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("B1").Select
Application.CutCopyMode = False
Selection.Copy
Range("A1").Select
ActiveSheet.Paste
Columns("B:B").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Columns("F:F").Select
Selection.Delete Shift:=xlToLeft
Columns("A:E").Select
Range("E1").Activate
Selection.Columns.AutoFit

Sheets("Upload Template").Select
Range("C2").Select
Sheets("Upload Template").Select
Range("C2").Select
Selection.Copy
Sheets("Instructions").Select
Range("F2").Select
ActiveSheet.Paste
Range("G2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""mmddyyyy"")"
Range("F2:G2").Select
Range("G2").Activate
With Selection.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
End With
Range("F3").Select


'This sets the check date as a variable
Dim CheckDate As Variant
CheckDate = Range("G2")

Sheets("Upload Template").Select

Columns("A:E").Select
ActiveWorkbook.Worksheets("Upload Template").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Upload Template").Sort.SortFields.Add _
    Key:=Range("A2:A" & LR), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Upload Template").Sort
    .SetRange Range("A1:E" & LR)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Dim sFilename As String
sFilename = "Benefit Payments - Check Date " & CheckDate & " - Upload.csv"  'You can give a nem to save
Workbooks.Add
'Saving the Workbook
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & sFilename, FileFormat:=xlCSV, CreateBackup:=False


Windows("Benefit Payments - After - Macro").Activate
Sheets("Upload Template").Select
Columns("A:E").Select
Selection.Copy

Windows("Benefit Payments - Check Date " & CheckDate & " - Upload.csv").Activate
Range("A1").Select
ActiveSheet.Paste
ActiveWorkbook.Save

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.DisplayAlerts = True

End Sub
