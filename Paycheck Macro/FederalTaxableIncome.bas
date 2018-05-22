Sub Main()
    Call Import
    Call ManipulateData
End Sub

Sub Import()
    Call ImportRawData.ImportRawData("Federal Taxable Income")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Fed Taxable Inc"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",,RC[1],RC[6],RC[8],RC[9])")
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub ManipulateData()
End Sub
