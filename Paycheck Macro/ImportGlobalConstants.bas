Attribute VB_Name = "ImportGlobalConstants"

Option Explicit

Public LastRow As Long
Public LastColumn As Long
Public MainWbName As String
Public RawDataWbName As String
Public RawDataFilename As String
Public RawDataWksName As String
Public UIDFormula As String

Public Sub Constants()
    MainWbName = ActiveWorkbook.Name
End Sub
