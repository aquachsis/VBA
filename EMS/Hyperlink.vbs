Sub EMSHyperlink()
Dim i as Long
Dim url as String
i = 2
Do While Range("F" & i).Value <> ""
    url = "https://na7.salesforce.com/" & Range("G" & i)
    ActiveSheet.Hyperlinks.Add Range("F" & i), url
    i = i + 1
Loop
End Sub
