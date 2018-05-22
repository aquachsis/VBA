Attribute VB_Name = "a_Main"

Sub a_Main()
    Call ImportGlobalConstants.Constants

    Activesheet.Name = "Main"
    Call ImportRawData
    Call ConsolidateSheets
    MsgBox "Completed."
End Sub

Sub ConsolidateSheets()
    Worksheets("Main").Activate

    'This is the UID'
    Worksheets("Fed Taxable Inc").Columns(1).Copy Destination:= _
    Worksheets("Main").Range("A1")

End Sub

Sub ImportRawData()
    ' Call SplitPayReports.Deductions
    ' Call SplitPayReports.Earnings
    ' Call SplitPayReports.Taxes
    ' Call DirectDeposits.DirectDeposits
    Call FederalTaxableIncome.Main
    ' Call AddressWithholding.Main
    Call CostCenters.Main
End Sub
