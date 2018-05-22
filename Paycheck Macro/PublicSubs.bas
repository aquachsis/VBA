Public Sub FirstRowDelete()
    Do While IsEmpty(Cells(1,1).Value)
       Rows(1).EntireRow.Delete
   Loop
End Sub

Public Sub CopyToSheet(SheetName)
    Workbooks(MainWbName).Sheets.Add.Name = SheetName
    Workbooks(RawDataWbName).Worksheets(1).Cells.Copy Destination:= _
    Workbooks(MainWbName).Worksheets(SheetName).Range("A1")
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
