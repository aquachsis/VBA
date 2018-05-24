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

    'UID'
    Worksheets("Fed Taxable Inc").Columns(1).Copy Destination:= _
    Worksheets("Main").Range("A1")

    'Define LastRow'
    LastRow = PublicFunctions.FindLastRow(1)

    Call PublicSubs.VLOOKUP("Fed Taxable Inc","C",4)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","H",2)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","K",8)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","M",9)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","O",10)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","S",5)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","U",11)
End Sub

Sub ImportRawData()
    Call SplitPayReports.Deductions
    Call SplitPayReports.Earnings
    Call SplitPayReports.Taxes
    Call DirectDeposits.DirectDeposits
    Call FederalTaxableIncome.Main
    Call AddressWithholding.Main
    Call CostCenters.Main
End Sub
