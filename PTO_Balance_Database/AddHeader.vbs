Public FolderPath As String
Public FileName As String

Sub FormatPTOFiles()
    Call ApplicationOff

    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    'PTO groups 2, 3, and 4 are only XLSX'
    'Dir only works on Windows, not Macs'
    FileName = Dir(FolderPath & "*.XLSX")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)

        ' TODO Call Actions Here
        Call AddHeader

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

Sub AddHeader()
    Range("D1") = "Transmission Date"
End Sub
