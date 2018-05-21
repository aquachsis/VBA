Sub ImportRawData(ReportName)
    MsgBox "Select the " & ReportName & " report / raw data."
    RawDataFilename = Application.GetOpenFilename
    Workbooks.Open Filename:=RawDataFilename
    RawDataWbName = ActiveWorkbook.Name
    Call PublicSubs.FirstRowDelete
End Sub
