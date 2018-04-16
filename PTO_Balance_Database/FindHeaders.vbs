Public FolderPath As String
Public FileName As String
Public FileNameNoExt As String
Public LastRow As Long
Public LastColumn As Long
Public TransmissionHeaderColumn As Long
Public TransmissionDate As String
Public MainWorkbook As String
Public RowCount As Long

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


Sub FormatUltipro()
    Cells.Unmerge
    Do While IsEmpty(Range("B1").Value)
        Rows(1).EntireRow.Delete
    Loop
End Sub

Sub CopyOverHeader()
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    Worksheets(1).Range(Cells(1, 1),Cells(1,LastColumn)).Copy Destination:= _
        Workbooks(MainWorkbook).Worksheets(1).Range("B" & RowCount)
    Workbooks(MainWorkbook).Worksheets(1).Range("A" & RowCount) = FileName
    RowCount = RowCount + 1
End Sub


Sub FindHeaders()
    Call ApplicationOff
    MainWorkbook = ActiveWorkbook.Name
    RowCount = 1
    Dim Wb As Workbook
    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    ' FileName = Dir(FolderPath & "*.CSV")
    FileName = Dir(FolderPath & "*.XLSX")
    Do While FileName <> ""
        Set Wb = Workbooks.Open(FolderPath & FileName)
        Sheets(1).Name = "PTO Balance"
        Call FormatUltipro
        Call CopyOverHeader
        Wb.Close SaveChanges:=False
        FileName = Dir()
    Loop
    Call ApplicationOn
    MsgBox "Completed"
End Sub
