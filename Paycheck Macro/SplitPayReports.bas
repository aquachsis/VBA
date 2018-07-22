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
    Call PublicSubs.CreateUID2("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")
    Columns("B:G").Delete 'TODO might be able to add to CreateUID2'

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
    Workbooks(MainWbName).Worksheets("Deductions").Range("A1").Resize(UBound(DestArray), LastColumn) = DestArray

'##### Create expense sheet from source file.
    Call CreateSheetAndEmptyArray("Expenses")
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) = "EXP" Or RowNum = 1 Then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Expenses").Range("A1").Resize(UBound(DestArray), LastColumn) = DestArray

    Workbooks(RawDataWbName).Close SaveChanges:=False 'close source file
End Sub

Sub Earnings()
    Call ImportRawData.ImportRawData("Earnings/Memos")
    Call PublicSubs.CreateUID2("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")
    Columns("B:G").Delete 'TODO might be able to add to CreateUID2'

'##### Defines source array size and set constants.
    FirstRow = 1
    FirstColumn = 1
    LastRow = PublicFunctions.FindLastRow(1) 'Uncessary, but makes block clear?
    LastColumn = PublicFunctions.FindLastColumn
    SrcArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))
    CodeColumn = 2

'##### Create earnings sheet from source file.
    Call CreateSheetAndEmptyArray("Earnings")
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) <> "Memo" OR RowNum = 1 Then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Earnings").Range("A1").Resize(UBound(DestArray), LastColumn) = DestArray

'##### Create memos sheet from source file.
    Call CreateSheetAndEmptyArray("Memos")
    For RowNum = 1 To UBound(SrcArray)
        If SrcArray(RowNum, CodeColumn) = "Memo" OR RowNum = 1 Then
            Call CreateNewArray()
        Else
        End If
    Next RowNum
    Workbooks(MainWbName).Worksheets("Memos").Range("A1").Resize(UBound(DestArray), LastColumn) = DestArray

    Workbooks(RawDataWbName).Close SaveChanges:=False 'close source file
End Sub

Sub Taxes()
    Call ImportRawData.ImportRawData("Taxes")
    Call PublicSubs.CreateUID2("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")
    Columns("B:G").Delete 'TODO might be able to add to CreateUID2'

'##### Defines source array size and set constants.
    FirstRow = 1
    FirstColumn = 1
    LastRow = PublicFunctions.FindLastRow(1) 'Uncessary, but makes block clear?
    LastColumn = PublicFunctions.FindLastColumn
    SrcArray = Range(Cells(FirstRow,FirstColumn), Cells(LastRow, LastColumn))
    CodeColumn = 2

'##### Create memos sheet from source file.
    Call CreateSheetAndEmptyArray("Taxes")
    For RowNum = 1 To UBound(SrcArray)
        Call CreateNewArray()
    Next RowNum
    Workbooks(MainWbName).Worksheets("Taxes").Range("A1").Resize(UBound(DestArray), LastColumn) = DestArray

    Workbooks(RawDataWbName).Close SaveChanges:=False 'close source file
End Sub
