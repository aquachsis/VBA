Sub FormatTimeClock()
'
' FormatTimeClock Macro
' Format time clock information from the paroll object to match with Ultipro.
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
   Sheets("Formulas").Select
    Columns("K:N").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K1").Select
    Application.CutCopyMode = False
    Range("K1").Select

    Columns("A:J").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select

'begin sort sequence

  Columns("A:A").Select
    Range("A:A").Activate
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Add Key:=Range( _
        "A:A"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formulas").Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Columns("B:B").Select
    Range("B:B").Activate
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Add Key:=Range( _
        "B:B"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formulas").Sort
        .SetRange Range("B:B")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Columns("C:C").Select
    Range("C:C").Activate
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Add Key:=Range( _
        "C:C"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formulas").Sort
        .SetRange Range("C:C")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Columns("D:D").Select
    Range("D:D").Activate
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Add Key:=Range( _
        "D:D"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formulas").Sort
        .SetRange Range("D:D")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


'begin replacing the empty cells
    Columns("A:D").Select
    Range("A1").Activate
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select

'begin adding first column and first row
    Columns("A:A").Select
    Range("A1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select




'begin single column steps
   Dim Rng As Range, LR As Long, i As Long
    LR = Range("B" & Rows.Count).End(xlUp).Row
    For i = 2 To LR
        Set Rng = Range("B" & i, "E" & i) 'Change range to suit needs
        Range("A" & Rows.Count).End(xlUp)(2).Resize(Rng.Count) = Application.WorksheetFunction.Transpose(Rng)
    Next i

    Range("A1").Select

        ActiveCell.Columns("A:A").EntireColumn.Select
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formulas").Sort.SortFields.Add Key:=ActiveCell.Range _
        ("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formulas").Sort
        .SetRange ActiveCell.Columns("A:A").EntireColumn
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Columns("B:B").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Range("A1").Select


'comma delimit
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    Range("A1").Select

'concatenate
    Range("D1").Select
    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-2])"
    ActiveCell.Offset(1, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Range("A1").Select


'delete first blank row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

'sum if and paste value
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C[-1],RC[-1],C[-2])"
    Range("D1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'remove duplicates
    ActiveCell.Cells.Select
    ActiveSheet.Range("A:E").RemoveDuplicates Columns:=4, Header:= _
        xlNo
    ActiveCell.Select


'delete useless columns
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select

'add boilerplate info
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "='Paste Data Here'!R2C3+1"
        Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2").Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D2").Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Range("A1").Select

 'delete blank columns and rows
    Columns("F:F").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    Columns("A:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.DisplayAlerts = False

       'Delimit employee number from pay rate
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|"


    Columns("B:B").Select
    Selection.NumberFormat = "General"
    Range("A1").Select

Application.DisplayAlerts = True

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "E"
    Range("B2").Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Columns("C:C").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Range("F1").Select

End Sub
