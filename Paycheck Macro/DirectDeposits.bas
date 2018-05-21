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
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "Direct Deposit"
    ' Check Type | Account Type | Account Number | Amount
    Range("A2:A" & LastRow).FormulaR1C1 = _
        "=TEXTJOIN(""|"",FALSE,RC[10],RC[11],RIGHT(RC[12],4),IF(RC[10]=""Voucher"",RC[15],RC[14]))"
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value
End Sub
