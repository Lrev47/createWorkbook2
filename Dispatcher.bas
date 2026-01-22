Attribute VB_Name = "Dispatcher"
Option Explicit

' ============================================================================
' DISPATCHER - Routes dropdown changes to the right TX module
' ============================================================================
' When user selects an order type, this clears the form and calls the
' appropriate TX_*.Build() to set up the new form layout.
' ============================================================================

Public Sub HandleOrderTypeChange(ByVal orderType As String)
    ' Triggered by Worksheet_Change when user picks a new order type
    Debug.Print "[Dispatcher] HandleOrderTypeChange: Received '" & orderType & "'"
    Dim ws As Worksheet
    Dim moduleName As String

    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Clear everything first - prevents stale data from previous order type
    ClearFormArea ws

    ' Route to the matching TX module based on selection
    moduleName = GetModuleName(orderType)

    If moduleName <> "" Then
        On Error Resume Next
        Application.Run moduleName & ".Build"
        If Err.Number <> 0 Then
            Debug.Print "Error calling " & moduleName & ".Build: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Info Box needs to match the new order type's fields
    EntryPoint.BuildInfoBox orderType

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Debug.Print "[Dispatcher] Loaded form for " & orderType
End Sub

Public Sub ClearFormArea(Optional ws As Worksheet = Nothing)
    ' Wipe form completely before building new layout

    ' Get Sheet1 if not provided
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets("Sheet1")
    End If

    ' Full range: header row + form fields + 300-row bulk entry area
    With ws.Range("B5:F311")
        .UnMerge
        .ClearContents
        .ClearFormats
        .Validation.Delete
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(255, 255, 255)  ' Reset to white
    End With

    ' Restore header row styling
    ws.Range("B5:F5").Interior.Color = RGB(100, 120, 150)

    ' Hidden area below form stays dark navy (matches Main)
    ws.Range("B40:F311").Interior.Color = RGB(0, 0, 51)

    Debug.Print "[Dispatcher] Cleared and reset form area B5:F311"
End Sub

Private Function GetModuleName(ByVal orderType As String) As String
    ' Simple mapping from dropdown value to module name
    Select Case orderType
        Case "New Usage"
            GetModuleName = "TX_NewUsage"
        Case "Return"
            GetModuleName = "TX_Return"
        Case "Swap"
            GetModuleName = "TX_Swap"
        Case Else
            GetModuleName = ""
            Debug.Print "[Dispatcher] Unknown order type - " & orderType
    End Select
End Function
