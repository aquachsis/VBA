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

    'Tab, Main Tab's Column, Source Column
    Range("A1") = "UID"
    Range("B1") = "Employee Number"
    Range("C1") = "Address"
    Range("D1") = "Check Date"
    Range("E1") = "Deductions [nested object]"
    Range("F1") = "Department"
    Range("G1") = "Division"
    Range("H1") = "Earnings [nested object]"
    Range("I1") = "Expenses [nested object]"
    Range("J1") = "Federal Filing Status"
    Range("K1") = "Federal Taxable Income"
    Range("L1") = "Gross Earnings"
    Range("M1") = "Memos [nested object]"
    Range("N1") = "Net Pay"
    Range("O1") = "Pay Distribution [nested object]"
    Range("P1") = "Pay Period Beginning"
    Range("Q1") = "Pay Period Ending"
    Range("R1") = "Process ID"
    Range("S1") = "PTO"
    Range("T1") = "Rate"
    Range("U1") = "State Filing Status"
    Range("V1") = "Taxes [nested object]"
    Range("W1") = "Void"
    Range("X1") = "Voucher / Check No"
    Range("Y1") = "Working State"

    Call PublicSubs.VLOOKUP("Fed Taxable Inc","B",2)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","D",4)
    Call PublicSubs.LookupByEE("Cost Centers","F",4)
    Call PublicSubs.LookupByEE("Cost Centers","G",3)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","K",8)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","L",9)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","N",10)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","R",5)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","T",11)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","X",7)
    Call PublicSubs.LookupByEE("Cost Centers","Y",5)
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
