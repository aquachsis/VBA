Public columnCount As Long
Public excelRange As Range
Public jsonMainColl As New Collection
Public jsonDeductionColl As New Collection
Public jsonDict As New Dictionary
Public jsonDictAmount As New Dictionary
Public jsonDictCode As New Dictionary
Public rowCount As Long
Public UIDValue As String

Sub convertPay()
    Worksheets("Deduction").Activate
    Set excelRange = Cells(1, 1).CurrentRegion
    'If UID = UID+1 Then //Add Dict //Else Add to Dict, Append Collection, Clear Dic
    For rowCount = 2 To excelRange.Rows.Count
        UIDValue = Cells(rowCount, 1)
        If Cells(rowCount, 1) = Cells(rowCount + 1, 1) Then
            Call addToDict
        Else
            Call addToDict
            jsonDeductionColl.Add jsonDictCode, UIDValue
            Set jsonDictCode = Nothing
        End If
    Next rowCount

    ' MsgBox JsonConverter.ConvertToJson(jsonDeductionColl("12345"), Whitespace:=2)

    Worksheets("Main").Activate
    Set excelRange = Cells(1, 1).CurrentRegion
    For rowCount = 2 To excelRange.Rows.Count
        UIDValue = Cells(rowCount, 1)
        jsonDict("UID") = UIDValue
        jsonDict("Net") = Cells(rowCount, 2)
        jsonDict("Deduction") = jsonDeductionColl(UIDValue)
        jsonMainColl.Add jsonDict, UIDValue
        Set jsonDict = Nothing
    Next rowCount

    MsgBox JsonConverter.ConvertToJson(jsonMainColl, Whitespace:=2)

    Set jsonDeductionColl = Nothing
    Set jsonMainColl = Nothing
End Sub

Public Sub addToDict()
    jsonDictAmount("Amount") = Cells(rowCount, 3)
    jsonDictCode(Cells(rowCount, 2)) = jsonDictAmount
    Set jsonDictAmount = Nothing
End Sub


Sub accessDictAndColl()
    Dim UIDValue As String
    UIDValue = Cells(2, 1)
    jsonDictCode("Code") = Cells(2, 2)
    jsonDictCode("Amount") = Cells(2, 3)
    jsonMainColl.Add jsonDictCode, UIDValue
    MsgBox JsonConverter.ConvertToJson(jsonMainColl("12345"), Whitespace:=2)

    'This accesses the decution information by employee number'
    Worksheets("Deduction").Activate
    UIDValue = Cells(4, 1)
    jsonDictCode("Code") = Cells(4, 2)
    jsonDictCode("Amount") = Cells(4, 3)
    jsonMainColl.Add jsonDictCode, UIDValue
    MsgBox JsonConverter.ConvertToJson(jsonMainColl("99999"), Whitespace:=2)

    Set jsonMainColl = Nothing
End Sub
