Sub a_Main()
    Call ImportGlobalConstants.Constants
    MainWbName = ActiveWorkbook.Name
    Call SplitPayReports.Deductions
    Call SplitPayReports.Earnings
    Call SplitPayReports.Taxes
    Call DirectDeposits.DirectDeposits
    Call FederalTaxableIncome.Import
    Call AddressWithholding.Import
    MsgBox "Completed."
End Sub
