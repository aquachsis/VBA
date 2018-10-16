'Used to review direct deposit that came through onboarding'
Public LastRow As Long
Public MainWbName As String
Public SourceFilename As String
Dim EmployeeNumberColumn As Integer
Dim RoutingNumberColumn As Integer
Dim AccountNumberColumn As Integer
Dim AccountTypeColumn As Integer
Dim OrderColumn As Integer
Dim AmountColumn As Integer
Dim UIDColumn As Integer
Dim AmountConverted As String
Dim NotesColumn As Integer

Sub ReviewDirectDeposit()
    MainWbName = ActiveWorkbook.Name
    Sheets(1).Name = "Main"
    Call LoadFile("Salesforce")
    Call LoadFile("Paylocity")
    Call FormatPaylocityData
    Call FormatSalesforceData
    Call CreateUploadFile
    MsgBox "Macro completed."
End Sub

Public Sub FindLastRow(ColumnLetter)
    LastRow = Range(ColumnLetter & Rows.Count).End(xlUp).Row
End Sub

Sub FormatPaylocityData()
    Worksheets("Paylocity").Activate
    Call UnformatSheet
    Call FindLastRow("B") 'Column A includes footers, so use B'

    'Fills in omitted employee name and numbers, Paylocity groups them
    Range("A1:B" & LastRow).SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    Range("A1:B" & LastRow).Value = Range("A1:B" & LastRow).Value

    EmployeeNumberColumn = 3
    RoutingNumberColumn = 6
    AccountNumberColumn = 7
    AccountTypeColumn = 8
    OrderColumn = 4
    AmountColumn = 10
    UIDColumn = 1

    'Creates UID: EE#|ABA|ACCT|Type|Order|AMT'
    Columns(1).Insert
    Cells(1, UIDColumn) = "UID"
    For i = 2 To LastRow
        Cells(i, UIDColumn) = _
            Cells(i, EmployeeNumberColumn) & "|" & _
            Cells(i, RoutingNumberColumn) & "|" & _
            Cells(i, AccountNumberColumn) & "|" & _
            Cells(i, AccountTypeColumn) & "|" & _
            Cells(i, OrderColumn) & "|" & _
            Cells(i, AmountColumn)
    Next i
End Sub

Sub FormatSalesforceData()
    Worksheets("Salesforce").Activate
    Call UnformatSheet
    Call FindLastRow("B") 'Column A includes footers, so use B'

    EmployeeNumberColumn = 3
    RoutingNumberColumn = 6
    AccountNumberColumn = 7
    AccountTypeColumn = 9
    OrderColumn = 8
    AmountColumn = 11
    UIDColumn = 12
    PrenoteColumn = 13
    CorrectInPaylocityColumn = 14

    Cells(1, UIDColumn) = "UID"
        For i = 2 To LastRow
        If Cells(i, OrderColumn) = "99" Then
            AmountConverted = "100"
        Else
            AmountConverted = Cells(i, AmountColumn)
        End If

        Cells(i, UIDColumn) = _
            Cells(i, EmployeeNumberColumn) & "|" & _
            Cells(i, RoutingNumberColumn) & "|" & _
            Cells(i, AccountNumberColumn) & "|" & _
            Cells(i, AccountTypeColumn) & "|" & _
            Cells(i, OrderColumn) & "|" & _
            AmountConverted
    Next i

    Cells(1, PrenoteColumn) = "Skipped Prenote?"
    Call FillInData("M", "=IF(ISNA(VLOOKUP($L2,Paylocity!$A:$I,9,0)),""NA"",VLOOKUP($L2,Paylocity!$A:$I,9,0))")
    For i = 2 To LastRow
        If Cells(i, PrenoteColumn) = "0" Then
            Cells(i, PrenoteColumn).Value = "Good. Not skipped."
        Else
            Cells(i, PrenoteColumn).Value = "Bad. Cannot find or skipped."
            Cells(i, PrenoteColumn).Interior.ColorIndex = 3 '3 = Red
            End If
        Next i

    Cells(1, CorrectInPaylocityColumn) = "Correct in Paylocity?"
    Call FillInData("N", "=IF(ISNA(VLOOKUP($L2,Paylocity!$A:$A,1,0)),""No"",""Yes"")")
    For i = 2 To LastRow
        If Cells(i, CorrectInPaylocityColumn) = "No" Then
            Cells(i, CorrectInPaylocityColumn).Interior.ColorIndex = 3 '3 = Red
        Else
        End If
    Next i

End Sub

Sub LoadFile(SourceName)
    MsgBox "Select the " & SourceName & " report."
    SourceFilename = Application.GetOpenFilename
    Workbooks.Open Filename:=SourceFilename
    SourceFilename = ActiveWorkbook.Name
    Workbooks(SourceFilename).Sheets(1).Name = SourceName
    Workbooks(SourceFilename).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)
End Sub

Sub UnformatSheet()
    ActiveWindow.DisplayGridlines = True
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    With ActiveSheet
        .AutoFilterMode = False
        .Cells.WrapText = False
        .Cells.UnMerge
        .Cells.EntireRow.Hidden = False
        .Cells.EntireColumn.Hidden = False
    End With
    Do While IsEmpty(Range("A1").Value)
        Rows(1).EntireRow.Delete
    Loop
End Sub

Sub CreateUploadFile()
    Dim HeaderRow As Integer
    Dim IDColumn As Integer
    Dim ReviewedInPaylocityColumn As Integer
    Dim UploadCounter As Long
    HeaderRow = 1
    IDColumn = 1
    ReviewedInPaylocityColumn = 2
    UploadCounter = 2

    Sheets.Add.Name = "Upload"
    Cells(HeaderRow, IDColumn) = "id"
    Cells(HeaderRow, ReviewedInPaylocityColumn) = "Reviewed in Paylocity"

    Worksheets("Salesforce").Activate
    Call FindLastRow("B") 'Column A includes footers, so use B'
    For i = 2 To LastRow
        If Cells(i, 14) = "Yes" AND Cells(i, 13) = "Good. Not skipped." Then
            Worksheets("Upload").Cells(UploadCounter, IDColumn).Value = Worksheets("Salesforce").Cells(i, 2).Value
            Worksheets("Upload").Cells(UploadCounter, ReviewedInPaylocityColumn).Value = "TRUE"
            UploadCounter = UploadCounter + 1
        Else
        End If
    Next i

    Worksheets("Upload").SaveAs ThisWorkbook.Path & Application.PathSeparator & "ddReviewedUpload.csv", CreateBackup:=False, FileFormat:=xlCSV
End Sub

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
    Columns.AutoFit
End Sub
