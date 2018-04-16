Public FolderPath As String
Public FileName As String
Public FileNameNoExt As String
Public LastRow As Long
Public LastColumn As Long
Public TransmissionHeaderColumn As Long
Public TransmissionDate As String

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

Sub SaveAsXLSX()
    WorkbookName = ActiveWorkbook.Name
    If InStr(WorkbookName, ".") > 0 Then
       WorkbookName = Left(WorkbookName, InStr(WorkbookName, ".") - 1)
    End If
    OutfileNameXLSX = WorkbookName
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        "New" & _
        Application.PathSeparator & _
        OutfileNameXLSX, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
End Sub

Sub AddTransmissionDate()
    'First 46 characters of all file names are the same
    TransmissionDate = Mid(FileName, 46, Len(FileName))
    If InStr(TransmissionDate, ".") > 0 Then
       TransmissionDate = Left(TransmissionDate, InStr(TransmissionDate, ".") - 1)
    End If

    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
     'Salesforce reports have footers in column 1, get last row from col 2
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    TransmissionHeaderColumn = LastColumn + 1
    Cells(1, TransmissionHeaderColumn) = "Transmission Date"
    Range( _
        Cells(2, TransmissionHeaderColumn), _
        Cells(LastRow,TransmissionHeaderColumn)) _
        = TransmissionDate
    Range( _
        Cells(2, TransmissionHeaderColumn), _
        Cells(LastRow,TransmissionHeaderColumn)).NumberFormat = "yyyy-mm-dd"
End Sub

Sub FormatPTOFiles()
    Call ApplicationOff
    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    ' FileName = Dir(FolderPath & "*.XLSX")
    ' FileName = Dir(FolderPath & "*.XLS")
    FileName = Dir(FolderPath & "*.CSV")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)

        Sheets(1).Name = "PTO Balance"
        Call AddTransmissionDate
        Call SaveAsXLSX
        Wb.Close SaveChanges:=False
        FileName = Dir()
    Loop
    Call ApplicationOn
    MsgBox "Completed"
End Sub
