'TODO Attribute VB_Name = "AddressWithholding"

Sub Main()
    Call ImportGlobalConstants.Constants 'TODO remove, testing'
    ActiveSheet.Name = "Main" 'TODO remove, testing'

    Call Import
    Call RemoveUncessaryData
    Call ManipulateData
End Sub

Sub Import()
    Call ImportRawData.ImportRawData("Address and Withholding")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Add and WH"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub RemoveUncessaryData()
    ' TODO we need to add the name back in
    Call PublicSubs.Unformat()
    Rows(1).EntireRow.Delete
    ' Delete uncessary columns
    Columns(22).Delete
    Range(Columns(2), Columns(18)).Delete
End Sub

Sub ManipulateData()
    LastRow = PublicFunctions.FindLastRow(1) 'TODO remove, testing'

    'Creates UID then delete the original columns
    Call InsertFormula("UID", "A", "=TEXTJOIN(""|"",FALSE,$B2:$C2)")
    Range(Columns(2), Columns(3)).Delete

    LastRow = PublicFunctions.FindLastRow(1)

    'Creates address field'
    Call InsertFormula("Address", "B", "=TEXTJOIN(""|"",FALSE,RC[3]:RC[7])")
    Columns("E:I").EntireColumn.Delete 'Delete separate address fields'

    'SPLIT FEDERAL WITHHOLDING STATUS AND ALLOWANCES'
    Call InsertFormula("FITW Status", "F","=LEFT($E2,FIND("" "",$E2)-1)")
    Call InsertFormula("FITW Allowance","G", _
        "=MID(RIGHT($E2,LEN($E2)-FIND("" "",$E2)),2,LEN(RIGHT($E2,LEN($E2)-FIND("" "",$E2)))-7)")
    Columns("E:E").Delete 'Delete original concatenated FITW Status & Allowance'

    'Splits the federal withhold amount (default, flat amounts, or percentages)
    Call InsertFormula("FITW Withholding Type", "H","=LEFT($G2,FIND("" "",$G2)-1)")
    Call InsertFormula("FITW Withholding Amount", "I", "=RIGHT($G2,LEN($G2)-FIND("" "",$G2))")
    Columns("G:G").Delete

    Call InsertFormula("FITW Withholding Amount", "I", "")
    For i = 2 To LastRow
        If Cells(i, 7) = "D" OR Cells(i, 7) = "B" Then
            'Default or Blocked
            Cells(i, 9) = "0"
        ElseIf Cells(i, 7) = "F" OR Cells(i, 7) = "AF" Then
            'Flat Dollar Amount or Additional Flat Amount
            Cells(i, 9) = "=MID(RC[-1],2,LEN(RC[-1])-7)"
        ElseIf Cells(i, 7) = "P" OR Cells(i, 7) = "AP" Then
            'Flat Percentage or Additional Percentage
            Cells(i, 9) = "=MID(RC[-1],7,LEN(RC[-1])-7)"
        ElseIf Cells(i, 7) = "AFAP" OR Cells(i, 7) = "FDFP" Then
            'Additional Flat Plus Additional %, or Flat Dollar Plus Fixed %
            Cells(i, 9) = "=MID(RC[-1],2,LEN(RC[-1])-2)"
        Else
        End If
    Next i
    Range("I:I").Value = Range("I:I").Value
    Columns("H:H").Delete

    'SPLIT STATE WITHHOLDING STATUS AND ALLOWANCES'
    Call InsertFormula("SITW Status", "K","=LEFT($J2,FIND("" "",$J2)-1)")
    Call InsertFormula("SITW Allowances", "L", _
        "=MID(RIGHT($J2,LEN($J2)-FIND("" "",$J2)),2,LEN(RIGHT($J2,LEN($J2)-FIND("" "",$J2)))-7)")
    Columns("J:J").Delete 'Delete original concatenated SITW Status & Allowance'

    'Splits the state withhold amount (default, flat amounts, or percentages)
    Call InsertFormula("SITW Withholding Type", "M","=LEFT($L2,FIND("" "",$L2)-1)")
    Call InsertFormula("SITW Withholding Amount", "N", "=RIGHT($L2,LEN($L2)-FIND("" "",$L2))")
    Columns("L:L").Delete

    Call InsertFormula("SITW Withholding Amount", "N", "")
    For i = 2 To LastRow
        If Cells(i, 12) = "D" OR Cells(i, 12) = "B" Then
            'Default or Blocked
            Cells(i, 14) = "0"
        ElseIf Cells(i, 12) = "F" OR Cells(i, 12) = "AF" Then
            'Flat Dollar Amount or Additional Flat Amount
            Cells(i, 14) = "=MID(RC[-1],2,LEN(RC[-1])-7)"
        ElseIf Cells(i, 12) = "P" OR Cells(i, 12) = "AP" Then
            'Flat Percentage or Additional Percentage
            Cells(i, 14) = "=MID(RC[-1],7,LEN(RC[-1])-7)"
        ElseIf Cells(i, 12) = "AFAP" OR Cells(i, 12) = "FDFP" Then
            'Additional Flat Plus Additional %, or Flat Dollar Plus Fixed %
            Cells(i, 14) = "=MID(RC[-1],2,LEN(RC[-1])-2)"
        Else
        End If
    Next i
    Range("N:N").Value = Range("N:N").Value
    Columns("M:M").Delete

    For i = 2 To LastRow
        If Cells(i, 10) = "N/A" Then
            Range("K" & i & ":M" & i) = "N/A"
        Else
        End If
    Next i

    Range("A1").Value = "UID"
    Range("B1").Value = "Address"
    Range("C1").Value = "Period Begin"
    Range("D1").Value = "Period Date"
    Range("E1").Value = "Fed Status"
    Range("F1").Value = "Fed Allowance"
    Range("G1").Value = "Fed Type"
    Range("H1").Value = "Fed Amount"
    Range("I1").Value = "State"
    Range("J1").Value = "State Status"
    Range("K1").Value = "State Allowance"
    Range("L1").Value = "State Type"
    Range("M1").Value = "State Amount"
End Sub


Public Sub FillFormula(Header, ColumnLetter, Formula)
    Range(ColumnLetter & "1") = Header
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow) = Formula
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub

Public Sub InsertFormula(Header, ColumnLetter, Formula)
    Columns(ColumnLetter & ":" & ColumnLetter).Insert
    Range(ColumnLetter & "1") = Header
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow) = Formula
    Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value = _
        Range(ColumnLetter & "2:" & ColumnLetter & LastRow).Value
End Sub
