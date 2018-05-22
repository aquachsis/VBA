Attribute VB_Name = "ImportGlobalConstants"

Option Explicit

Public MainWbName As String
Public RawDataWbName As String
Public RawDataFilename As String
Public RawDataWksName As String
Public UIDFormula As String

Public Sub Constants()
    MainWbName = ActiveWorkbook.Name
End Sub
