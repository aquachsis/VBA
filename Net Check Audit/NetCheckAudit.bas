
Option Explicit
Public LastRow As Long
Public MainWbName As String
Public SourceFilename As String
Public SourceWbName As String

Sub NetCheckAudit()
'Run this Subroutine, it's the main one.
    Call OpenNetCheckAudit
    Call LoadEINotIssued
    Worksheets("Main").Activate
    Call Unformat
    Call DeleteSubtotals
    Call DetermineDisposableIncome
    Call DetermineCategory
    Call AddFormatting
    Call SortData
    'TODO replacements should have $0 tax, and ded = net'
    Worksheets("Main").Activate
    MsgBox "Complete."
End Sub

Sub OpenNetCheckAudit()
    MsgBox "Select the Pre Process Net Check Audit report."
    SourceFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=SourceFilename
    Sheets(1).Name = "Main"
    MainWbName = ActiveWorkbook.Name
End Sub

Sub Unformat()
    ActiveWindow.DisplayGridlines = True

    With Cells
        .WrapText = False
        .MergeCells = False
    End With

    Do While IsEmpty(Cells(1, 1).Value)
       Rows(1).EntireRow.Delete
   Loop
End Sub

Sub DeleteSubtotals()
    LastRow = FindLastRow(1)
    Dim i As Long
    For i = LastRow To 2 Step -1
        If Not Application.IsNumber(Cells(i, 5)) = True Then
            Rows(i).EntireRow.Delete
        End If
    Next i
End Sub

Sub DetermineDisposableIncome()
    LastRow = FindLastRow(1)
    Range("K1") = "50% of Disposable Income"
    Call FillInData("K", "=($G2-$H2)/2")
End Sub

Sub DetermineCategory()
    Range("L1") = "Category"
    'To prevent overlap and priotize in this order:
    'Net under 5. Net less than 50% of disposable income. Gross over $2500.
    Call FillInData( _
        "L", _
        "=IF(ISNA(VLOOKUP($A2,'EI Not Issued'!$C:$C,1,0))," & _
            "IF($A2=17296,""Peter Pan, delete check.""," & _
            "IF($J2<5,""Net Under $5""," & _
            "IF($J2<$K2,""Net Less Than 50% of Disposable Income""," & _
            "IF($G2>2500,""Gross Over $2500"",""""))))," & _
        """Remove check, final check processed but not issued."")" _
    )
End Sub

Sub AddFormatting()
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
    End With

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Columns.AutoFit
    Rows.AutoFit
    ' Worksheets(1).Columns("A:Z").AutoFi
    Columns(1).Value = Columns(1).Value
End Sub

Sub SortData()
'Sort by category, batch, then employee number
    Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlDescending, _
        Key2:=Range("C1"), Order2:=xlAscending, _
        Key3:=Range("A1"), Order3:=xlAscending, _
        Header:=xlYes
End Sub

Sub LoadEINotIssued()
    MsgBox "Select the EI - Processed Not Issued report."
    SourceFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=SourceFilename
    Sheets(1).Name = "Main"
    SourceWbName = ActiveWorkbook.Name
    Workbooks(SourceWbName).Sheets(1).Name = "EI Not Issued"
    Workbooks(SourceWbName).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)
End Sub

Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub
