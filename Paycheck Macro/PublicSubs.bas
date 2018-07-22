'Attribute VB_Name = "PublicSubs" TODO add back later

Public Sub FirstRowDelete()
    Do While IsEmpty(Cells(1,1).Value)
       Rows(1).EntireRow.Delete
   Loop
End Sub

Public Sub CopyToSheet (SheetName)
    Workbooks(MainWbName).Sheets.Add.Name = SheetName
    LastRow = PublicFunctions.FindLastRow(1)
    LastColumn = PublicFunctions.FindLastColumn
    ' TODO instead of using column AM, find a way for incorporate FindLastColumn
    ' to dynamically update depending on the report
    Workbooks(MainWbName).Worksheets(SheetName).Range("A1:AM" & LastRow).Value = _
    Workbooks(RawDataWbName).Worksheets(1).Range("A1:AM" & LastRow).Value
End Sub

Public Sub CreateUID2 (UIDFormula)
    LastRow = PublicFunctions.FindLastRow(1)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "UID"
    Range("A2:A" & LastRow).FormulaR1C1 = UIDFormula
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
End Sub

Public Sub CreateUID (UIDFormula)
    Workbooks(MainWbName).Worksheets(RawDataWksName).Activate
    LastRow = PublicFunctions.FindLastRow(1)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "UID"
    Range("A2:A" & LastRow).FormulaR1C1 = UIDFormula
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
End Sub

Public Sub Unformat()
    ActiveWindow.DisplayGridlines = True
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Cells.UnMerge
End Sub

Public Sub CreateSheetAndEmptyArray(SheetName)
    Erase DestArray
    Workbooks(MainWbName).Sheets.Add.Name = SheetName
    Workbooks(MainWbName).Worksheets(SheetName).Activate
    DestArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))
    DestRowNum = 1
End Sub

Public Sub CreateNewArray()
    For ColNum = 1 To LastColumn
        DestArray(DestRowNum, ColNum) = SrcArray(RowNum, ColNum)
    Next ColNum
    DestRowNum = DestRowNum + 1
End Sub

Public Sub InsertFormula(Header, ColumnLetter, Formula)
    Range(ColumnLetter & "1") = Header
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow) = Formula
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub

Public Sub VLOOKUP (SheetTableArray, ColumnLetter, colIndex)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Formula ="=VLOOKUP($A2,'" & SheetTableArray & "'!$A:$K," & colIndex &",FALSE)"
End Sub

Public Sub LookupByEE (SheetTableArray, ColumnLetter, colIndex)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Formula = _
    "=VLOOKUP($B2,'" & SheetTableArray & "'!$A:$K," & colIndex &",FALSE)"
End Sub

Public Sub LookupByEEandCK (SheetTableArray, ColumnLetter, colIndex)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Formula = _
    "=VLOOKUP($B2 &""|""& $X2,'" & SheetTableArray & "'!$A:$M," & colIndex & ",FALSE)"
End Sub
