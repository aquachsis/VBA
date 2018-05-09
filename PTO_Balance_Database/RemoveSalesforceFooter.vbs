Public FolderPath As String
Public FileName As String
Public LastRowColumnA As Long
Public LastColumn As Long
Public LastRowOfLastColumn As Long

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

Sub RemoveFooter()
LastRowColumnA = Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
LastRowOfLastColumn = Cells(Rows.Count, LastColumn).End(xlUp).Row

If LastRowColumnA = LastRowOfLastColumn then
Else
    Rows(LastRowOfLastColumn + 1 & ":" & LastRowColumnA).EntireRow.Delete
End If
End Sub


Sub FormatPTOFiles()
    Call ApplicationOff

    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    'Dir only works on Windows, not Macs'
    FileName = Dir(FolderPath & "*.XLSX")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)
        Call RemoveFooter
        Wb.Close SaveChanges:=True
        FileName = Dir()
    Loop

    Call ApplicationOn
    MsgBox "Completed"
End Sub
