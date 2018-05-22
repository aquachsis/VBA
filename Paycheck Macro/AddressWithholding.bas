Sub Main()
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
    'Creates UID then delete the original columns
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[1]:RC[2])")
    Range(Columns(2), Columns(3)).Delete

    'Creates address field'
    LastRow = PublicFunctions.FindLastRow(1)
    Range("B1").EntireColumn.Insert
    Range("B1").Value = "Address"
    Range("B2:B" & LastRow).FormulaR1C1 = "=TEXTJOIN(""|"",FALSE,RC[3]:RC[7])"
    Range("B2:B" & LastRow).Value = Range("B2:B" & LastRow).Value

    'TODO Split federal withholding information'
    'TODO Split state withholding information'
End Sub
