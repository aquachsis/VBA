Attribute VB_Name = "PublicSubs"

Public Sub FirstRowDelete()
    Do While IsEmpty(Cells(1,1).Value)
       Rows(1).EntireRow.Delete
   Loop
End Sub

Public Sub CopyToSheet(SheetName)
    Workbooks(MainWbName).Sheets.Add.Name = SheetName
    LastRow = PublicFunctions.FindLastRow(1)
    LastColumn = PublicFunctions.FindLastColumn
    ' TODO change to cell ranges
    ' Workbooks(MainWbName).Worksheets(SheetName).Range("A1:H" & LastRow).Value = _
    ' Workbooks(RawDataWbName).Worksheets(1).Range("A1:H" & LastRow).Value
    Workbooks(MainWbName).Worksheets(SheetName).Range( _
        Cells(1, 1), Cells(LastRow, LastColumn)).Value = _
    Workbooks(RawDataWbName).Worksheets(1).Range( _
        Cells(1, 1), Cells(LastRow, LastColumn)).Value
End Sub

Public Sub CreateUID(UIDFormula)
    Workbooks(MainWbName).Worksheets(RawDataWksName).Activate
    LastRow = PublicFunctions.FindLastRow(1)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "UID"
    Range("A2:A" & LastRow).FormulaR1C1 = UIDFormula
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
End Sub

Public Sub Unformat()
    'Show Gridlines
    ActiveWindow.DisplayGridlines = True
    'Unfilter
    ActiveSheet.AutoFilterMode = False
    'Unmerge
    ActiveSheet.Cells.UnMerge
End Sub

Public Sub InsertFormula(ColumnLetter, Formula)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).FormulaR1C1 = Formula
End Sub

Public Sub VLOOKUP (SheetTableArray, ColumnLetter, SearchRow)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Formula ="=VLOOKUP($A2,'" & SheetTableArray & "'!$A:$K," & SearchRow &",FALSE)"
End Sub
