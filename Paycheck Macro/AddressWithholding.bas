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

    'Creates address field'
    LastRow = PublicFunctions.FindLastRow(1)
    Call InsertFormula("Address", "B", "=TEXTJOIN(""|"",FALSE,RC[3]:RC[7])")
    Columns("E:I").EntireColumn.Delete

    'Splits Federal Status and Allowances'
    Columns("F:F").Insert
    Columns("E:E").TextToColumns _
        Destination:=Range("E1"), _
        DataType:=xlDelimited, _
        Space:=True
    Call InsertFormula("Federal Allowances", "G", "=MID($F2,2,LEN($F2)-7)")
    Columns("F:F").Delete

'TODO THIS IS WHERE I AM WORKING AT

    'Splits the federal withhold amount (default, flat amounts, or percentages)
    Columns("H:I").Insert
    Columns("G:G").TextToColumns _
        Destination:=Range("G1"), _
        DataType:=xlDelimited, _
        Space:=True

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

    'Splits state withholding status'
    Columns("K:K").Insert
    Columns("J:J").TextToColumns _
        Destination:=Range("J1"), _
        DataType:=xlDelimited, _
        Space:=True
    Call InsertFormula("", "L", "=MID(RC[-1],2,LEN(RC[-1])-7)")
    Columns("K:K").Delete

    Columns("L:L").TextToColumns _
        Destination:=Range("L1"), _
        DataType:=xlDelimited, _
        Space:=True

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
    Range("C1").Value = "Begin Date"
    Range("D1").Value = "End Date"
    Range("E1").Value = "FITW Election Status"
    Range("F1").Value = "FITW Exemptions"
    Range("G1").Value = "Fed Amount Type"
    Range("H1").Value = "Fed Amount"
    Range("I1").Value = "State"
    Range("J1").Value = "SITW Filing Status"
    Range("K1").Value = "SITW Exemptions"
    Range("L1").Value = "State Amount Type"
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
