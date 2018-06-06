Sub NetCheckAudit()
    Call Unformat
    Call DeleteSubtotals
End Sub

Sub Unformat()
    ActiveWindow.DisplayGridlines = True

    With Cells
        .WrapText = False
        .MergeCells = False
    End With

    Do While IsEmpty(Cells(1, 1).Value)
       Rows(1).EntireRow.Delete
   Loop
End Sub

Sub DeleteSubtotals()
    Dim LastRow As Long
    LastRow = FindLastRow(1)

    For i = LastRow To 2 Step -1
        If Not Application.IsNumber(Cells(i, 5)) = True Then
            Rows(i).EntireRow.Delete
        End If
    Next i
End Sub


Public Function FindLastRow(ColumnNumber)
    FindLastRow = ActiveSheet.Cells(Rows.Count, ColumnNumber).End(xlUp).Row
End Function
