Sub Deductions()
    Call ImportRawData.ImportRawData("Deductions/Expenses")

    'Copy deductions.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="<>EXP"
    RawDataWksName = "Deductions"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Sub CreateUID()
        'do stuff here'
    End Sub

    'Copy expenses.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="EXP"
    RawDataWksName = "Expenses"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub Earnings()
    Call ImportRawData.ImportRawData("Earnings/Memos")

    'Copy deductions.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="<>Memo"
    RawDataWksName = "Earnings"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    'Copy expenses.
    Workbooks(RawDataWbName).Worksheets(1).Activate
    Range("A:H").AutoFilter _
        Field:=7, _
        Criteria1:="Memo"
    RawDataWksName = "Memos"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub Taxes()
    Call ImportRawData.ImportRawData("Taxes")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Taxes"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[3]:RC[6])")

    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub
