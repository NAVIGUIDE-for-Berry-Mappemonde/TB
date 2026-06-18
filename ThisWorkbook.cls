Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    InitialiserGestionSiAbsent
    On Error GoTo 0
End Sub
