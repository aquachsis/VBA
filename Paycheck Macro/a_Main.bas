Sub a_Main()
    Call ImportGlobalConstants.Constants
    Call ImportRawData
    Call ConsolidateSheets
    MsgBox "Completed."
End Sub

Sub ConsolidateSheets()
End Sub

Sub ImportRawData()
    MainWbName = ActiveWorkbook.Name

    ' Call SplitPayReports.Deductions
    ' Call SplitPayReports.Earnings
    ' Call SplitPayReports.Taxes
    ' Call DirectDeposits.DirectDeposits
    ' Call FederalTaxableIncome.Main
    ' Call AddressWithholding.Main
    Call CostCenters.Main
End Sub
