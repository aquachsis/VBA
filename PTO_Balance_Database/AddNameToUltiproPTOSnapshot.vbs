Public FolderPath As String
Public FileName As String
Public LastRowColumnA As Long
Public LastRowColumnB As Long

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

Sub RemoveTitleAndFooter()
ActiveSheet.Cells.UnMerge

'Ultipro has a title in row 1, if B2 is blank then that's the title'
'Delete the title'
Do While IsEmpty(Range("B1").Value)
    Rows(1).EntireRow.Delete
Loop


'Delete the footer'
LastRowColumnA = Cells(Rows.Count, 1).End(xlUp).Row
LastRowColumnB = Cells(Rows.Count, 2).End(xlUp).Row

If LastRowColumnA = LastRowColumnB then
Else
    Rows(LastRowColumnA).EntireRow.Delete
End If
End Sub

Sub AddEmployeeName()
    Columns ("A").TextToColumns
    Columns("B:C").Insert
    Range("B1") = "First Name"
    Range("C1") = "Last Name"
    Range("B2:B" & LastRowColumnB).FormulaR1C1 = _
        "=VLOOKUP(RC[-1],[formatPTO.xlsm]Names!C1:C3,2,0)"
    Range("C2:C" & LastRowColumnB).FormulaR1C1 = _
        "=VLOOKUP(RC[-2],[formatPTO.xlsm]Names!C1:C3,3,0)"
    Range("B2:C" & LastRowColumnB).Value = Range("B2:C" & LastRowColumnB).Value
End Sub

Sub FormatPTOFiles()
    Call ApplicationOff

    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    'PTO groups 2, 3, and 4 are only XLSX'
    'Dir only works on Windows, not Macs'
    FileName = Dir(FolderPath & "*.XLSX")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)
        Call RemoveTitleAndFooter
        Call AddEmployeeName
        Wb.Close SaveChanges:=True
        FileName = Dir()
    Loop

    Call ApplicationOn
    MsgBox "Completed"
End Sub
