Attribute VB_Name = "convertToJson"

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
Public rowStart As Long
Public UIDColumn As Long
Public UIDValue As String
Public UIDValueNextRow As String

Sub Notes()
'https://www.linkedin.com/pulse/convert-excel-json-using-vba-thomas-clouse/
'http://excelerator.solutions/2017/08/24/convert-excel-to-json-using-vba/
'https://github.com/VBA-tools/VBA-Dictionary
'https://github.com/VBA-tools/VBA-JSON'

'TODO what if there are multiple of same code for the same UID? e.g. one UID has 2 ORCAs
'TODO earnings can have two XRG with two different rates, how to handle?
'TODO expenses probably doesn't need to be nested since it's just one expense amount

End Sub

'TODO set columns as parameters, reusability'
Sub addToDict()
    Dim amountColumn As Long
    Dim codeColumn As Long
    amountColumn = 2
    codeColumn = 3
    jsonDictAmount("Amount") = Cells(rowCount, codeColumn)
    jsonDictCode(Cells(rowCount, amountColumn)) = jsonDictAmount
    Set jsonDictAmount = Nothing
End Sub

Sub convertPay()
    UIDColumn = 1
    rowStart = 2
    Worksheets("Deductions").Activate
    Set excelRange = Cells(1, 1).CurrentRegion
    'rowCount starts on 2 because 1 is the header'
    For rowCount = rowStart To excelRange.Rows.Count
        UIDValue = Cells(rowCount, UIDColumn)
        UIDValueNextRow =  Cells(rowCount + 1, UIDColumn)
        If UIDValue = UIDValueNextRow Then
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
        UIDValue = Cells(rowCount, UIDColumn)
        jsonDict("UID") = UIDValue
        jsonDict("Deductions") = nestedDeductions(UIDValue)
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
