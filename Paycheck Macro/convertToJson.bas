Public columnCount As Long
Public excelRange As Range
Public jsonCollection As New Collection
Public jsonDict As New Dictionary
Public jsonDictAmount As New Dictionary
Public jsonDictCode As New Dictionary
Public nestedDeductions As New Collection
Public nestedEarnings As New Collection
Public nestedExpenses As New Collection
Public nestedMemos As New Collection
Public nestedPayDistributions As New Collection
Public nestedTaxes As New Collection
Public rowCount As Long
Public UIDValue As String

Sub Notes()
'https://www.linkedin.com/pulse/convert-excel-json-using-vba-thomas-clouse/
'http://excelerator.solutions/2017/08/24/convert-excel-to-json-using-vba/
'https://github.com/VBA-tools/VBA-Dictionary
'https://github.com/VBA-tools/VBA-JSON'

'TODO what if there are multiple of same code for the same UID? e.g. one UID has 2 ORCAs
End Sub

Sub convertPay()
    Worksheets("Deductions").Activate
    Set excelRange = Cells(1, 1).CurrentRegion
    'If UID = UID+1 Then //Add Dict //Else Add to Dict, Append Collection, Clear Dic

    For rowCount = 2 To excelRange.Rows.Count
        UIDValue = Cells(rowCount, 1)
        If Cells(rowCount, 1) = Cells(rowCount + 1, 1) Then
            Call addToDict
        Else
            Call addToDict
            nestedDeductions.Add jsonDictCode, UIDValue
            Set jsonDictCode = Nothing
        End If
    Next rowCount

    Worksheets("Main").Activate
    Set excelRange = Cells(1, 1).CurrentRegion
    For rowCount = 2 To excelRange.Rows.Count
        UIDValue = Cells(rowCount, 1)
        jsonDict("UID") = UIDValue 'TODO what if I don't add this to JSON?
        jsonDict("Deduction") = nestedDeductions(UIDValue)
            For columnCount = 2 To excelRange.Columns.Count
                jsonDict(Cells(1, columnCount)) = Cells(rowCount, columnCount)
            Next columnCount
        jsonCollection.Add jsonDict, UIDValue
        Set jsonDict = Nothing
    Next rowCount

    MsgBox JsonConverter.ConvertToJson(jsonCollection, Whitespace:=2)

    Set nestedDeductions = Nothing
    Set jsonCollection = Nothing
End Sub

Public Sub addToDict()
    jsonDictAmount("Amount") = Cells(rowCount, 3)
    jsonDictCode(Cells(rowCount, 2)) = jsonDictAmount
    Set jsonDictAmount = Nothing
End Sub
