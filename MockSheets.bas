Attribute VB_Name = "MockSheets"
Option Explicit

Public Sub CreateMockSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim sheetName As Variant

    sheetNames = Array("CustomerDB", "PODB", "Customer List", "CRDB", "InventoryDB")

    Application.ScreenUpdating = False

    For Each sheetName In sheetNames
        ' Check if sheet already exists
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(sheetName)
        End If
        Set ws = Nothing
    Next sheetName

    Application.ScreenUpdating = True

    Debug.Print "[MockSheets] Created 5 mock sheets"
End Sub
