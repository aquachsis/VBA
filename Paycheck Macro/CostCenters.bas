Attribute VB_Name = "CostCenters"

Sub Main()
    Call Import
    Call ManipulateData
End Sub

Sub Import()
    Call ImportRawData.ImportRawData("Cost Centers")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Cost Centers"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub ManipulateData()
End Sub
