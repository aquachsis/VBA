Public MainWbName As String
Public RawDataWbName As String
Public FolderPath As String
Public FileName As String
Public FirstRow As Long
Public FirstColumn As Long
Public HeaderRow As Long
Public LastColumn As Long
Public LastRow As Long
Public DestArray() As Variant
Public SrcArray() As Variant
Public Wb As Workbook

Sub GlobalConstants()
    HeaderRow = 1
    FirstRow = 1
    FirstColumn = 1
End Sub
Sub CompileMonthlyReport()
    Dim UnecessaryColumns As Range
    Dim FirstUncessaryColumn As Long
    Call GlobalConstants

    FolderPath = ThisWorkbook.Path & Application.PathSeparator
    FileName = Dir(FolderPath & "*.csv")
    Workbooks.Open(FolderPath & FileName)
    MainWbName = ActiveWorkbook.Name
    LastColumn = FindLastColumn(1)
    LastRow = FindLastRow(1)
    FirstUncessaryColumn = 10
    Set UnecessaryColumns = _
        Range(Cells(HeaderRow, FirstUncessaryColumn), Cells(LastRow, LastColumn))
    UnecessaryColumns.Delete

    Call EditWeeklyReports()
    Call AddFormatting()
    Call SaveAsXLSX()

    MsgBox "Completed"
End Sub
Sub EditWeeklyReports()
    Dim ColNum As Long
    Dim RowNum As Long
    FileName = Dir(FolderPath & "*.xlsx")
    Do While FileName <> ""
        Workbooks.Open(FolderPath & FileName)
        RawDataWbName = ActiveWorkbook.Name
        Workbooks(RawDataWbName).Worksheets(1).Activate
        Rows("1:10").Delete 'Uncessary rows
        LastColumn = FindLastColumn(1)
        LastRow = FindLastRow(1)
        Erase SrcArray
        SrcArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))
        Erase DestArray
        DestArray = Range(Cells(FirstRow,FirstColumn + LastColumn), Cells(LastRow, LastColumn + LastColumn))
        For RowNum = 2 To UBound(SrcArray)
            If SrcArray(RowNum, 5) <> "" Then
                DestArray(RowNum, 1) = SrcArray(RowNum, 8)
                DestArray(RowNum, 2) = SrcArray(RowNum, 7)
                DestArray(RowNum, 3) = SrcArray(RowNum, 9)
                DestArray(RowNum, 4) = SrcArray(RowNum, 10)
                DestArray(RowNum, 5) = SrcArray(RowNum, 8)
                DestArray(RowNum, 6) = SrcArray(RowNum, 3)
                DestArray(RowNum, 7) = SrcArray(RowNum, 5)
                DestArray(RowNum, 8) = SrcArray(RowNum, 6)
                DestArray(RowNum, 9) = SrcArray(RowNum, 2)
            Else
            End If
        Next RowNum

        Workbooks(MainWbName).Worksheets(1).Activate
        LastRow = FindLastRow(9)
        Workbooks(MainWbName).Worksheets(1).Cells(LastRow + 1, 1).Resize(UBound(SrcArray), LastColumn) = DestArray
        Workbooks(RawDataWbName).Close SaveChanges:=False 'close source file
        FileName = Dir()
    Loop
End Sub
Sub AddFormatting()
    LastColumn = FindLastColumn(1)
    Range(Cells(FirstRow, FirstColumn), Cells(FirstRow, LastColumn)).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.7
    End With
    Cells.Sort _
        Key1:=Range("F1"), Order1:=xlAscending, _
        Header:=xlYes

        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    Cells.Columns.AutoFit
End Sub
Sub SaveAsXLSX()
    SaveAsFilename = InputBox("What do you want to file name to be?","File name","ALINE YYYY ")
    OutfileNameXLSX = SaveAsFilename & ".xlsx"
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        OutfileNameXLSX, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
End Sub
Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function
Public Function FindLastColumn(RowNumber)
    FindLastColumn = ActiveSheet.Cells(RowNumber, Columns.Count).End(xlToLeft).Column
End Function
