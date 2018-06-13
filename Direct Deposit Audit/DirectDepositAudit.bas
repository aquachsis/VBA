Attribute VB_Name = "DirectDepositAudit"

Public LastRow as Long
Public MainWbName As String
Public RawDataFilename As String
Public RawDataWbName As String

Sub a_AuditDirectDeposits()
'Run this one. It's the main one.
    MainWbName = ActiveWorkbook.Name
    Sheets(1).Name = "Main"
    Call LoadFile("Salesforce")
    Call LoadFile("Paylocity")
    Call FormatPaylocityData
    Call FormatSalesforceData
    Call CompareData
End Sub

Sub FormatPaylocityData()
    Worksheets("Paylocity").Activate
    Call UnformatSheet()

    Call FindLastRow("B")
    'Fills in employee number.
    'Paylocity only shows employee number once, even if there are multiple accounts.
    Range("A1:B" & LastRow).SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    Range("A1:B" & LastRow).Value = Range("A1:B" & LastRow).Value

    Columns(1).Insert
    Range("A1") = "Paylocity: Employee ID | Routing | Account | Type | Order"
    Call FillInData("A","=CONCATENATE($C2,""|"",$F2,""|"",$G2,""|"",$H2,""|"",$D2)")

End Sub

Sub FormatSalesforceData()
    Worksheets("Salesforce").Activate
    Call UnformatSheet()
    Call FindLastRow("B")
    Range("B" & LastRow).Value = Range("B" & LastRow).Value

    Columns(1).Insert
    Range("A1") = "Salesforce: Employee ID | Routing | Account | Type | Order"
    Call FillInData("A","=CONCATENATE($B2,""|"",$G2,""|"",$H2,""|"",$J2,""|"",$I2)")
End Sub

Sub CompareData()
    Worksheets("Paylocity").Activate
    Columns(2).Insert
    Range("B1") = "In Salesforce?"
    Call FillInData("B","=IF(ISNA(VLOOKUP($A2,Salesforce!$A:$A,1,0)),""No"",""Yes"")")

    Worksheets("Salesforce").Activate
    Columns(2).Insert
    Range("B1") = "In Paylocity?"
    Call FillInData("B","=IF(ISNA(VLOOKUP($A2,Paylocity!$A:$A,1,0)),""No"",""Yes"")")
End Sub

Public Sub LoadFile(Filename)
    MsgBox "Select the " & Filename & " report."
    RawDataFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=RawDataFilename
    RawDataWbName = ActiveWorkbook.Name
    Workbooks(RawDataWbName).Sheets(1).Name = Filename
    Workbooks(RawDataWbName).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)
End Sub

Public Sub UnformatSheet()
    ActiveWindow.DisplayGridlines = True
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

Public Sub FindLastRow(ColumnLetter)
    LastRow = Range(ColumnLetter & Rows.Count).End(xlUp).Row
End Sub

Public Sub FillInData(ColumnLetter, Data)
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = Data
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
    Columns.AutoFit
End Sub
