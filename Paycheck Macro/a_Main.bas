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

    'TODO main sure there are no duplicates in UID'
    Worksheets("Fed Taxable Inc").Columns(1).Copy Destination:= _
        Worksheets("Main").Range("A1") 'This is the UID'

    'Define LastRow'
    LastRow = PublicFunctions.FindLastRow(1)

    'TODO rewrite functions to take header name as arugement'
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

    'These are first because other columns are dependent on them
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","B",2) 'employee number
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","X",7) 'voucher/check number

    'Function arguments are: tab to lookup from, main tab's column, src's column
    Call PublicSubs.LookupByEEandCK("Add and WH", "C",2)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","D",4)
        Columns("D").NumberFormat = "yyyy-mm-dd"
    Call PublicSubs.LookupByEE("Cost Centers","F",4)
    Call PublicSubs.LookupByEE("Cost Centers","G",3)
    Call PublicSubs.LookupByEEandCK("Add and WH", "J",5)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","K",8)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","L",9)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","N",10)
    Call PublicSubs.LookupByEEandCK("Add and WH", "P",3)
        Columns("P").NumberFormat = "yyyy-mm-dd"
    Call PublicSubs.LookupByEEandCK("Add and WH", "Q",4)
        Columns("Q").NumberFormat = "yyyy-mm-dd"
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","R",5)
    Call PublicSubs.VLOOKUP("Fed Taxable Inc","T",11)
    Call PublicSubs.LookupByEEandCK("Add and WH", "U",10)
    Range("W2:W" & LastRow).Formula = "=($N2<0)"
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
