Sub a_Main()
    Call ImportGlobalConstants.Constants
    MainWbName = ActiveWorkbook.Name
    '
    ' Call SplitPayReports.Deductions
    ' Call SplitPayReports.Earnings
    ' Call SplitPayReports.Taxes
    ' Call DirectDeposits.DirectDeposits
    Call FederalTaxableIncome.Main
    ' Call AddressWithholding.Main
    MsgBox "Completed."
End Sub
