'TODO uncomment
'Attribute VB_Name = "SplitPayReports"
Public RowNum As Long
Public ColNum As Long
Public DestArray() As Variant
Public SrcArray() As Variant
Public CodeColumn As Long
Public DestRowNum As Long
Public FirstRow As Long
Public FirstColumn As Long
Sub Deductions()
    Call ImportRawData.ImportRawData("Deductions/Expenses")

'##### Create UID TODO make into a function
    LastRow = PublicFunctions.FindLastRow(1)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "UID"
    Range("A2:A" & LastRow).FormulaR1C1 = "=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])"
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
    Columns("B:G").Delete

'##### Defines source array size and set constants.
    FirstRow = 1
    FirstColumn = 1
    LastRow = PublicFunctions.FindLastRow(1) 'Uncessary, but makes block clear?
    LastColumn = PublicFunctions.FindLastColumn
    SrcArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))

    CodeColumn = 2

'##### Create deduction sheet from source file.
    Call CreateSheetAndEmptyArray("Deductions")
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) <> "EXP" OR RowNum = 1 Then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Deductions").Range("A1").Resize(UBound(DestArray), 3) = DestArray

'##### Create expense sheet from source file.
    Call CreateSheetAndEmptyArray("Expenses")
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) = "EXP" Or RowNum = 1 Then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Expenses").Range("A1").Resize(UBound(DestArray), 3) = DestArray

'##### Close the source file.
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub Earnings()
    Call ImportRawData.ImportRawData("Earnings/Memos")

    'Copy deductions.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="<>Memo"
    RawDataWksName = "Earnings"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    'Copy expenses.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="Memo"
    RawDataWksName = "Memos"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub Taxes()
    Call ImportRawData.ImportRawData("Taxes")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Taxes"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub
