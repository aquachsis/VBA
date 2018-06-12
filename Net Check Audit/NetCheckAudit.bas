Attribute VB_Name = "NetCheckAudit"

Option Explicit
Public LastRow As Long
Public NetCheckAuditFilename As String

Sub NetCheckAudit()
'Run this Subroutine, it's the main one.
    Call OpenNetCheckAudit
    Call Unformat
    Call DeleteSubtotals
    Call DetermineDisposableIncome
    Call DetermineCategory
    Call AddFormatting
    Call SortData
End Sub

Sub OpenNetCheckAudit()
    MsgBox "Select the Pre Process Net Check Audit report."
    NetCheckAuditFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=NetCheckAuditFilename
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
    Call FillInData("K","=($G2-$H2)/2")
End Sub

Sub DetermineCategory()
    Range("L1") = "Category"
    'To prevent overlap and priotize in this order:
    'Net under 5. Net less than 50% of disposable income. Gross over $2500.
    Call FillInData( _
        "L", _
        "=IF($J2<5,""Net Under $5"",IF($J2<$K2,""Net Less Than 50% of Disposable Income"",IF($G2>2500,""Gross Over $2500"","""")))" _
    )
End Sub

Sub AddFormatting()
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 12
    End With

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Columns.AutoFit
    Rows.AutoFit
    ' Worksheets(1).Columns("A:Z").AutoFit
End Sub

Sub SortData()
'Sort by category, batch, then employee number
    Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlDescending, _
        Key2:=Range("C1"), Order2:=xlAscending, _
        Key3:=Range("A1"), Order3:=xlAscending, _
        Header:=xlYes
End Sub

Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    ' Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
    '     Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub
