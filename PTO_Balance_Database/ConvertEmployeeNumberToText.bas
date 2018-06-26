Public FolderPath As String
Public FileName As String
Public LastRow As Long

Sub ConvertEmployeeNumberToText()
    Call ApplicationOff

    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    'PTO group 7 is only XLSX'
    'Dir only works on Windows, not Macs'
    FileName = Dir(FolderPath & "*.XLSX")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)
        Call ConvertNumberToText
        Wb.Close SaveChanges:=True
        FileName = Dir()
    Loop

    Call ApplicationOn
    MsgBox "Completed"
End Sub

Sub ApplicationOff()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
End Sub

Sub ApplicationOn()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub


Sub ConvertNumberToText()
    LastRow = ActiveSheet.Cells(Rows.Count, 20).End(xlUp).Row
    Columns("C:C").Insert
    Range("C1") = "Employee Number"
    Range("C2:C" & LastRow) = "=IF(ISBLANK($D2),"""",TEXT($D2,0))"
    Range("C2:C" & LastRow).Copy
    Range("C2:C" & LastRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D").Delete
End Sub
