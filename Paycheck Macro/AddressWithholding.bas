Attribute VB_Name = "AddressWithholding"

Sub Main()
    Call Import
    Call RemoveUncessaryData
    Call ManipulateData
End Sub

Sub Import()
    Call ImportRawData.ImportRawData("Address and Withholding")
    Workbooks(RawDataWbName).Worksheets(1).Activate
    RawDataWksName = "Add and WH"
    Call PublicSubs.CopyToSheet(RawDataWksName)
    Workbooks(RawDataWbName).Close SaveChanges:=False
End Sub

Sub RemoveUncessaryData()
    ' TODO we need to add the name back in
    Call PublicSubs.Unformat()
    Rows(1).EntireRow.Delete
    ' Delete uncessary columns
    Columns(22).Delete
    Range(Columns(2), Columns(18)).Delete
End Sub

Sub ManipulateData()
    'Creates UID then delete the original columns
    Call PublicSubs.CreateUID("=TEXTJOIN(""|"",FALSE,RC[1]:RC[2])")
    Range(Columns(2), Columns(3)).Delete

    'Creates address field'
    LastRow = PublicFunctions.FindLastRow(1)
    Range("B1").EntireColumn.Insert
    Range("B1").Value = "Address"
    Range("B2:B" & LastRow).FormulaR1C1 = "=TEXTJOIN(""|"",FALSE,RC[3]:RC[7])"
    Range("B2:B" & LastRow).Value = Range("B2:B" & LastRow).Value

    Columns("E:I").EntireColumn.Delete
          Columns("F:F").Insert

 Rows("1:1").Clear
          Columns("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
            Space:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1))
          Columns("G:G").Insert


        Range("G2:G" & LastRow).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-7)"
        Range("G:G").Value = Range("G:G").Value
        Columns("F:F").Delete

        Columns("H:H").Insert
        Columns("G:G").TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
          Space:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1))
          Columns("I:I"). Insert
          Range("I2:I" & LastRow).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-7)"
          Range("A:L").AutoFilter _
              Field:=7, _
              Criteria1:=Array("P", "AP"), _
              Operator:=xlFilterValues
          Range("I:I").SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=MID(RC[-1],7,LEN(RC[-1])-7)"
          ActiveSheet.AutoFilterMode = False
          Range("A:L").AutoFilter _
              Field:=7, _
              Criteria1:=Array("AFAP", "FDFP"), _
              Operator:=xlFilterValues
          Range("I:I").SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-2)"
          ActiveSheet.AutoFilterMode = False
          Range("I:I").Value = Range("I:I").Value
          Columns("H:H").Delete

Columns("K:K").Insert
Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
  Space:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1))
  Columns("L:L"). Insert
  Range("L2:L" & LastRow).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-7)"
Range("L:L").Value = Range("L:L").Value
  Columns("K:K").Delete

Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
  Space:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1))
  Range("N2:N" & LastRow).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-7)"
  Range("A:N").AutoFilter _
      Field:=12, _
      Criteria1:=Array("P", "AP"), _
      Operator:=xlFilterValues
Range("N:N").SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=MID(RC[-1],7,LEN(RC[-1])-7)"
ActiveSheet.AutoFilterMode = False
Range("A:N").AutoFilter _
    Field:=7, _
    Criteria1:=Array("AFAP", "FDFP"), _
    Operator:=xlFilterValues
Range("N:N").SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=MID(RC[-1],2,LEN(RC[-1])-2)"
    ActiveSheet.AutoFilterMode = False
Range("N:N").Value = Range("N:N").Value
  Columns("M:M").Delete


  Range("A:M").AutoFilter _
      Field:=10, _
      Criteria1:=("N/A"), _
      Operator:=xlFilterValues
Range("K:M").SpecialCells(xlCellTypeVisible).Value = "N/A"
ActiveSheet.AutoFilterMode = False

  Range("A1").Value = "UID"
  Range("B1").Value = "Address"
  Range("C1").Value = "Begin Date"
  Range("D1").Value = "End Date"
  Range("E1").Value = "FITW Election Status"
  Range ("F1").Value = "FITW Exemptions"
  Range("G1").Value = "Fed Amount Type"
  Range ("H1").Value = "Fed Amount"
  Range ("I1").Value = "State"
  Range("J1").Value = "SITW Filing Status"
  Range("K1").Value = "SITW Exemptions"
  Range("L1").Value = "State Amount Type"
  Range ("M1").Value = "State Amount"

End Sub
