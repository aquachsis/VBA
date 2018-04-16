Option Explicit

Sub ConvertRangeToColumn()
'https://www.extendoffice.com/documents/excel/4233-excel-stack-columns.html
Dim rngSrc As Range
Dim rngDest As Range
Dim rngFirstColumn As Range
Dim Rng As Range
Dim columnIndex As Long
Dim columnNumber As Long

Set rngSrc = Application.Selection
Set rngSrc = Application.InputBox("Source Ranges:", "Source", rngSrc.Address, Type:=8)
Set rngDest = Application.InputBox("Convert to (single cell):", "Destination", Type:=8)
Set rngFirstColumn = rngSrc.Columns(1) 'To repeat the first column

Application.ScreenUpdating = False

columnIndex = 0 'keeps track of which row in destination to paste into next
For Each Rng In rngSrc.Columns
    columnNumber = Rng.Column 'keeps track of which column in source we're on in the loop

    If columnNumber <> 1 Then
        'The employee ID / first column'
        rngFirstColumn.Range(Cells(2,1),Cells(rngFirstColumn.Rows.Count,1)).Copy
        rngDest.Offset(columnIndex, 0).PasteSpecial Paste:=xlPasteAll

        'The header / first row'
        Rng.Cells(1, 1).Copy
        Range(rngDest.Offset(columnIndex, 1), rngDest.Offset(columnIndex + Rng.Rows.Count - 2, 1)).PasteSpecial Paste:=xlPasteAll
        'The amounts / data values'
        Rng.Range(Cells(2,1),Cells(Rng.Rows.Count,1)).Copy
        rngDest.Offset(columnIndex, 2).PasteSpecial Paste:=xlPasteAll

        'Increments to track where to paste to next
        columnIndex = columnIndex + Rng.Rows.Count - 1
    End If


Next

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub
