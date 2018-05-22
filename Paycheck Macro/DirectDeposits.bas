Attribute VB_Name = "DirectDeposits"

Sub DirectDeposits()
    Call ImportRawData.ImportRawData("Direct Deposit")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Direct Deposit"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",,RC[1],RC[9],RC[7],RC[8])")
    Call CombineDirectDeposit
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub CombineDirectDeposit()
    'TODO how to handle employees that have live check and DD split'
    LastRow = PublicFunctions.FindLastRow(1)
    Range("B1").EntireColumn.Insert
    Range("B1").Value = "Direct Deposit"
    ' Check Type | Account Type | Account Number | Amount
    Range("B2:B" & LastRow).FormulaR1C1 = _
        "=TEXTJOIN(""|"",FALSE,RC[10],RC[11],RIGHT(RC[12],4),IF(RC[10]=""Voucher"",RC[15],RC[14]))"
    Range("B2:B" & LastRow).Value = Range("B2:B" & LastRow).Value
End Sub
