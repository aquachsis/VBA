Public CheckDate As String
Public EffectiveEndDate As Date
Public MainWbName As String
Public SourceFilename As String
Public SourceWbName As String

Sub ORCA()
'Compares the active ORCA Elections and Active ORCA cards in Salesforce
'We stop elections if there is no active cards and
'We start elections for active cards without an existing election
Dim ActiveCardsLR As Long
Dim BenefitElectionsLR As Long
Dim i As Long
Dim StartCount As Long
Dim StopCount As Long

MainWbName = ActiveWorkbook.Name
EffectiveEndDate = InputBox("FIRST day of the pay period?", "First Day of Pay Period", "m/d/yyyy")
EffectiveEndDate = DateAdd("d", -1, EffectiveEndDate)
CheckDate = InputBox("Check date we're processing?","Check Date","m/d/yyyy")

Call LoadFile("Active Cards")
ActiveCardsLR = Sheets("Active Cards").Range("B" & Rows.Count).End(xlUp).Row
Call LoadFile("Benefit Elections")
BenefitElectionsLR = Sheets("Benefit Elections").Range("B" & Rows.Count).End(xlUp).Row

'Working on starting elections
Sheets("Active Cards").Activate
Range("K1").Value = "Has Election?"
Range("K2:K" & ActiveCardsLR).Formula = _
  "=IF(ISNA(VLOOKUP(RC[-8],'Benefit Elections'!C[-8],1,0)),""Start Election"",""Has Election"")"
Range("L1").Value = "Cards for this Employee"
Range("L2:L" & ActiveCardsLR).Formula = "=COUNTIF(C[-9],RC[-9])"
Range("M1").Value = "Sum of Column L"
Range("M2").Formula = "=SUM(C[-1])+1"

If Range("M2").Value <> ActiveCardsLR Then
MsgBox "MACRO INTERRUPTED:" & vbNewLine & _
"Duplicates In Active Cards" & vbNewLine & _
"Remove the duplicate in Active Cards file then rerun macro." & vbNewLine & _
"Inform HR there is a duplicate card."
Range("A1:Z100").Value = "MACRO FAILED BECAUSE OF DUPLICATE. REMOVE DUPLICATES ON RAW FILE"
End
End If

StartCount = 1
For i = 1 To ActiveCardsLR
  If Range("K" & i).Value = "Start Election" Then
    StartCount = StartCount + 1
    Range("A" & i).Copy Destination:=Sheets("Start Elections").Range("A" & StartCount)
    Range("F" & i).Copy Destination:=Sheets("Start Elections").Range("B" & StartCount)
    Sheets("Start Elections").Range("C" & StartCount).Value = "012A0000000ra7r"
    Sheets("Start Elections").Range("D" & StartCount).Value = "a2wA0000002IAuS"
    Sheets("Start Elections").Range("E" & StartCount).Value = "Accepted"
  End If
Next i

'Working on stopping elections
Sheets("Benefit Elections").Activate
Range("F1").Value = "Has Card?"
Range("F2:F" & BenefitElectionsLR).Formula = _
  "=IF(ISNA(VLOOKUP(RC[-3],'Active Cards'!C[-3],1,0)),""Stop Election"",""Has Card"")"
Range("G1").Value = "Elections for this Employee"
Range("G2:G" & BenefitElectionsLR).Formula = "=COUNTIF(C[-4],RC[-4])"
Range("H1").Value = "Sum of Column G"
Range("H2").Formula = "=SUM(C[-1])+1"

If Range("H2").Value <> BenefitElectionsLR Then
  MsgBox "MACRO INTERRUPTED:" & vbNewLine & _
  "Duplicates In Elections" & vbNewLine & _
  "Remove the duplicate in Benefit Elections file then rerun macro." & vbNewLine & _
  "We must stop the duplicate election in Salesforce."
  Range("A1:Z100").Value = "MACRO FAILED BECAUSE OF DUPLICATE. REMOVE DUPLICATES ON RAW FILE"
  End
End If

StopCount = 1
For i = 1 To ActiveCardsLR
  If Range("F" & i).Value = "Stop Election" Then
    StopCount = StopCount + 1
    Range("A" & i).Copy Destination:=Sheets("Stop Elections").Range("A" & StopCount)
    Sheets("Stop Elections").Range("B" & StopCount).Value = EffectiveEndDate
  End If
Next i

'Display results summary and save files.
MsgBox "When asked to save as a macro free workbook, choose yes."
Call SaveFiles
MsgBox "Macro Completed. " & vbNewLine & vbNewLine & _
"We will start " & StartCount - 1 & " elections. " & vbNewLine & _
"We will stop " & StopCount - 1 & " elections."

End Sub

Public Sub SaveFiles()
    Worksheets("Start Elections").Activate
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & Application.PathSeparator & _
        "ORCA - Start Elections - Check Date " & _
        Format(CheckDate, "mmddyyyy") & ".csv", _
        FileFormat:=xlCSVMSDOS, CreateBackup:=False

    Worksheets("Stop Elections").Activate
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & Application.PathSeparator & _
        "ORCA - Stop Elections - Check Date " & _
        Format(CheckDate, "mmddyyyy") & ".csv", _
        FileFormat:=xlCSVMSDOS, CreateBackup:=False

    Worksheets("Instructions").Activate
    ActiveWorkbook.SaveAs _
        ThisWorkbook.Path & Application.PathSeparator & _
        "ORCA - Main Workbook - Check Date " & _
        Format(CheckDate, "mmddyyyy") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Public Sub LoadFile(Filename)
    MsgBox "Select the " & Filename & " report."
    SourceFilename = Application.GetOpenFilename
    Workbooks.Open FileName:=SourceFilename
    SourceWbName = ActiveWorkbook.Name
    Workbooks(SourceWbName).Sheets(1).Name = Filename
    Workbooks(SourceWbName).Sheets(1).Move After:= _
        Workbooks(MainWbName).Sheets(Workbooks(MainWbName).Worksheets.Count)
End Sub
