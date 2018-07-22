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
    'v--This creates UID--v'
    LastRow = PublicFunctions.FindLastRow(1)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "UID"
    Range("A2:A" & LastRow).FormulaR1C1 = "=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])"
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
    Columns("B:G").Delete
    '^--This creates UID--^'

    ' SrcArray = Range("A1:C" & LastRow)
    FirstRow = 1
    FirstColumn = 1
    LastRow = PublicFunctions.FindLastRow(1) 'Uncessary, but makes block clear?
    LastColumn = PublicFunctions.FindLastColumn
    SrcArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))

    'v--splits expenses from deductions--v'
    Workbooks(MainWbName).Sheets.Add.Name = "Deductions"
    Workbooks(MainWbName).Worksheets("Deductions").Activate
    DestArray = Range("A1:C" & LastRow)
    CodeColumn = 2
    DestRowNum = 1

    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) <> "EXP" then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Deductions").Range("A1").Resize(UBound(DestArray), 3) = DestArray
    Erase DestArray 'Clears array after setting it into cells, so we can reuse it later'
    '^--splits expenses from deductions--^'

    'v--splits expenses from deductions--v'
    Workbooks(MainWbName).Sheets.Add.Name = "Expenses"
    Workbooks(MainWbName).Worksheets("Expenses").Activate
    DestArray = Range("A1:C" & LastRow)
    DestRowNum = 1
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) = "EXP" then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Expenses").Range("A1").Resize(UBound(DestArray), 3) = DestArray
    '^--splits expenses from deductions--^'

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
