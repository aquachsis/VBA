Attribute VB_Name = "DirectDepositAudit"

Public LastRow as Long
Public MainSheetRowCounter As Long
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
    Call SaveAsXLSX
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
    Call FindLastRow("A")
    Columns(2).Insert
    Range("B1") = "In Salesforce?"
    LastRow = LastRow - 1 'Paylocity report has one blank cell at the bottom'
    Call FillInData("B","=IF(ISNA(VLOOKUP($A2,Salesforce!$A:$A,1,0)),""No"",""Yes"")")

    MainSheetRowCounter = 1
    For i = 2 To LastRow
        If Cells(i, 2) = "No" Then
            Sheets("Main").Cells(MainSheetRowCounter, 1) = _
                Sheets("Paylocity").Cells(i, 1)
            Sheets("Main").Cells(MainSheetRowCounter, 2) = _
                "In Paylocity but not in Salesforce"
            MainSheetRowCounter = MainSheetRowCounter + 1
        End If
    Next i

    Worksheets("Salesforce").Activate
    Call FindLastRow("A")
    Columns(2).Insert
    Range("B1") = "In Paylocity?"
    Call FillInData("B","=IF(ISNA(VLOOKUP($A2,Paylocity!$A:$A,1,0)),""No"",""Yes"")")
    For i = 2 To LastRow
        If Cells(i, 2) = "No" Then
            Sheets("Main").Cells(MainSheetRowCounter, 1) = _
                Sheets("Salesforce").Cells(i, 1)
            Sheets("Main").Cells(MainSheetRowCounter, 2) = _
                "In Salesforce but not in Paylocity"

            'Checks to see if it's in Salesforce but we haven't entered
            'into Paylocity, i.e. new hires
            If Cells(i, 21) = "0" Then
            Sheets("Main").Cells(MainSheetRowCounter, 3) = _
                "We haven't entered into Paylocity yet, i.e. new hire, or recently entered direct deposit."
            End If

            MainSheetRowCounter = MainSheetRowCounter + 1
        End If
    Next i

    Worksheets("Main").Activate
    Rows(1).Insert
    Columns(1).Insert
    Call FindLastRow("B")
    Range("A1") = "Employee ID"
    Call FillInData("A","=LEFT($B2,5)")
    Range("B1") = "Employee ID | Routing | Account | Type | Order"
    Range("C1") = "Error Type"
    Range("D1") = "Notes"
    Range("A1:D1").Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.7
    End With
    Range("A1").CurrentRegion.Sort _
        Key1:=Range("A1"), Order1:=xlAscending, _
        Header:=xlYes

        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True

    Columns.AutoFit
End Sub

Sub SaveAsXLSX()
    OutfileNameXLSX = "Direct Deposit Audit " & Format(Date, "mmddyyyy") &".xlsx"
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        OutfileNameXLSX, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
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
