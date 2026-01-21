Attribute VB_Name = "TX_Swap"
Option Explicit

' Order Type: Swap

Public Sub Build()
    ' Swap form: customer info header + 300-row Dealer ID entry area
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' B5 header label
    With ws.Range("B5")
        .Value = "Customer Information"
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Visual separation between labels and inputs
    ws.Range("B6:B11").Interior.Color = RGB(220, 220, 220)
    ws.Range("C6:F11").Interior.Color = RGB(255, 255, 255)

    ' Wider merge accommodates long URLs and contact info
    ws.Range("C6:F6").Merge
    ws.Range("C7:F7").Merge
    ws.Range("C8:F8").Merge
    ws.Range("C9:F9").Merge
    ws.Range("C10:F10").Merge
    ws.Range("C11:F11").Merge

    ' Column B labels
    ws.Range("B6").Value = "V Simple Link"
    ws.Range("B7").Value = "On Site Contact"
    ws.Range("B8").Value = "Phone"
    ws.Range("B9").Value = "Email"
    ws.Range("B10").Value = "Sales Rep"
    ws.Range("B11").Value = "Customer Damage?"

    ' Customer Damage flag affects billing - important for swap transactions
    ws.Range("C11").Validation.Delete
    ws.Range("C11").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Yes,No"
    ws.Range("C11").Validation.InCellDropdown = True

    ' Consistent with corporate template styling
    ws.Range("B6:B11").Font.Name = "Bookman Old Style"

    ' Subtle borders for clean look
    With ws.Range("B5:B11").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    ' Horizontal lines from B5:F5 to B11:F11
    With ws.Range("B5:F11").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With
    ' Bottom edge of last row
    With ws.Range("B11:F11").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With

    ' Swap table: returning equipment on left, replacement placeholder (D grayed out)
    ws.Range("B12").Value = "Returning Dealer ID"
    ws.Range("C12").Value = "UC#"
    ws.Range("D12").Value = "Replacement ID"
    ws.Range("E12:F12").Merge
    ws.Range("E12").Value = "Swap Date"

    ' Header row styling matches other section headers
    With ws.Range("B12:F12")
        .Interior.Color = RGB(100, 120, 150)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Replacement ID column is placeholder - filled in later by operations team
    ws.Range("D12").Interior.Color = RGB(50, 50, 50)

    ' =========================================================================
    ' DATA ENTRY AREA - 300 rows for Dealer ID entry
    ' =========================================================================

    ' First row always visible so user knows where to start
    With ws.Range("B13:F13")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Bookman Old Style"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Hidden rows reveal as data is entered - keeps form compact
    With ws.Range("B14:F312")
        .Interior.Color = RGB(0, 0, 51)
        .Font.Name = "Bookman Old Style"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Visually separate the Replacement ID column
    With ws.Range("D13:D312").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 51)
        .Weight = xlThin
    End With
    With ws.Range("D13:D312").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 51)
        .Weight = xlThin
    End With

    ' Wider date column for date picker visibility
    Dim i As Long
    For i = 13 To 312
        ws.Range("E" & i & ":F" & i).Merge
    Next i

    ' Date validation prevents common entry errors
    With ws.Range("E13:E312").Validation
        .Delete  ' Clear any existing validation
        .Add Type:=xlValidateDate, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="1/1/1900", _
             Formula2:="12/31/2099"
        .IgnoreBlank = True
        .ErrorTitle = "Invalid Date"
        .ErrorMessage = "Please enter a valid date in short date format."
        .ShowInput = False
        .ShowError = True
    End With

    ' =========================================================================
    ' CONDITIONAL FORMATTING - Zebra stripes reveal rows as data is entered
    ' =========================================================================

    ' Apply to hidden rows only (row 13 is always visible)
    Dim dataRange As Range
    Set dataRange = ws.Range("B14:F312")

    ' Start fresh
    dataRange.FormatConditions.Delete

    ' Alternating colors make it easier to track across wide rows
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B14<>"""",MOD(ROW(),2)=0)"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(240, 240, 240)

    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B14<>"""",MOD(ROW(),2)=1)"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)

    ' "Next row" indicator shows user where to enter next Dealer ID
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B13<>"""",$B14="""")"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)

    ' VBA limitation: FormatCondition.Borders doesn't support full styling

    Debug.Print "[TX_Swap] Build called - row 13 white, rows 14-312 dark navy with zebra stripes + next-row indicator"
End Sub

' ============================================================================
' VALIDATION & DATA FUNCTIONS
' ============================================================================

Public Function Validate() As Boolean
    ' Validates: V Simple URL present, customer lookup succeeded
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Debug.Print "[TX_Swap] Validate: Checking V Simple Link..."
    Debug.Print "[TX_Swap] Validate: V Simple Link = " & ws.Range("C6").Value
    ' V Simple Link required for tracking
    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter a V Simple Link.", vbExclamation, "Validation Error"
        Validate = False
        Exit Function
    End If

    ' Need valid URL structure for ID extraction
    If InStr(ws.Range("C6").Value, "/") = 0 Then
        MsgBox "V Simple Link must be a valid URL with an ID." & vbCrLf & _
               "Example: https://vsimple.com/deals/12345", vbExclamation, "Invalid URL"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_Swap] Validate: Customer Name lookup = " & ws.Range("I6").Value
    ' Customer lookup is driven by Dealer ID - if it fails, ID may be wrong
    If IsError(ws.Range("I6").Value) Then
        MsgBox "Customer Name lookup failed. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    ' Empty means Dealer ID not found in CRDB
    If Trim(CStr(ws.Range("I6").Value)) = "" Then
        MsgBox "Customer Name lookup returned empty. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_Swap] Validate: PASSED"
    Validate = True
End Function

Public Function GetTemplate() As String
    ' Swap only has one template variant
    GetTemplate = TPL_SWAP
    Debug.Print "[TX_Swap] GetTemplate: Selected " & GetTemplate
End Function

Public Function GetFileName() As String
    ' Equipment type determines filename format:
    ' BATT -> "Power Swap", CHGR -> "Charger Swap", else -> quantity + model
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim customerName As String
    Dim customerNumber As String
    Dim vSimpleId As String
    Dim equipType As String
    Dim model As String
    Dim qty As Long
    Dim transactionType As String

    ' Pull from Info Box - these are computed from Dealer ID lookups
    customerName = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I6")))
    customerNumber = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I7")))
    vSimpleId = PathHelper.ExtractVSimpleId(PathHelper.SafeCellValue(ws.Range("C6")))
    equipType = UCase(Trim(PathHelper.SafeCellValue(ws.Range("I11"))))
    model = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I12")))
    qty = Val(PathHelper.SafeCellValue(ws.Range("I13")))

    Debug.Print "[TX_Swap] GetFileName: customerName=" & customerName
    Debug.Print "[TX_Swap] GetFileName: customerNumber=" & customerNumber
    Debug.Print "[TX_Swap] GetFileName: vSimpleId=" & vSimpleId
    Debug.Print "[TX_Swap] GetFileName: equipType=" & equipType
    Debug.Print "[TX_Swap] GetFileName: model=" & model
    Debug.Print "[TX_Swap] GetFileName: qty=" & qty

    ' Validate required components
    If customerName = "" Then
        MsgBox "Customer Name not found.", vbExclamation, "Missing Data"
        GetFileName = ""
        Exit Function
    End If

    If vSimpleId = "" Then
        MsgBox "Could not extract VSimple ID from link.", vbExclamation, "Invalid URL"
        GetFileName = ""
        Exit Function
    End If

    ' Equipment type drives the transaction type in filename
    Select Case equipType
        Case "BATT"
            transactionType = "Power Swap"
        Case "CHGR"
            transactionType = "Charger Swap"
        Case Else
            ' Trucks get quantity + model format
            If qty = 0 Then qty = 1
            If model = "" Then model = "Equipment"
            transactionType = CStr(qty) & " " & model
    End Select

    Debug.Print "[TX_Swap] GetFileName: transactionType=" & transactionType

    ' Final filename format matches our naming convention
    GetFileName = customerName & "_" & transactionType & "_" & customerNumber & "_" & vSimpleId & "_UW.xlsm"
    Debug.Print "[TX_Swap] GetFileName: Result = " & GetFileName
    Exit Function

ErrorHandler:
    Debug.Print "[TX_Swap] GetFileName error: " & Err.Description
    MsgBox "Error generating filename: " & Err.Description, vbCritical, "Filename Error"
    GetFileName = ""
End Function

Public Function GetCustomerName() As String
    ' I6 is computed from Dealer ID -> CRDB -> CustomerDB chain
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    GetCustomerName = Trim(CStr(ws.Range("I6").Value))
End Function

Public Sub MapToWorkbook(wb As Workbook)
    ' Copy form values to the Swap template - similar structure to Return
    Dim srcWs As Worksheet
    Dim destWs As Worksheet

    Set srcWs = ThisWorkbook.Worksheets("Sheet1")
    Set destWs = wb.Worksheets("Overview")

    Debug.Print "[TX_Swap] MapToWorkbook: Starting field mapping..."
    Debug.Print "[TX_Swap] MapToWorkbook: Customer # -> C4 = " & srcWs.Range("I7").Value

    ' Customer Information
    destWs.Range("C4").Value = srcWs.Range("I7").Value    ' Customer # (from Info box)
    destWs.Range("C11").Value = srcWs.Range("C10").Value  ' Sales Rep
    destWs.Range("C12").Value = srcWs.Range("C7").Value   ' On Site Contact
    destWs.Range("C13").Value = srcWs.Range("C8").Value   ' Phone
    destWs.Range("C14").Value = srcWs.Range("C9").Value   ' Email
    destWs.Range("C15").Value = srcWs.Range("C11").Value  ' Customer Damage

    ' Embed V Simple link on header for easy access to deal
    Dim vSimpleUrl As String
    vSimpleUrl = Trim(srcWs.Range("C6").Value)

    If Len(vSimpleUrl) > 0 Then
        Call AddHyperlinkPreserveStyle(destWs, "B2", vSimpleUrl)
        Debug.Print "[TX_Swap] MapToWorkbook: V Simple hyperlink added to B2"
    End If

    ' Map bulk data to Swap sheet
    Dim swapWs As Worksheet
    Set swapWs = wb.Worksheets("Swap")

    Dim i As Long
    For i = 13 To 313
        If srcWs.Range("B" & i).Value <> "" Then
            swapWs.Range("D" & i).Value = srcWs.Range("B" & i).Value   ' Returning Dealer ID
            swapWs.Range("O" & i).Value = srcWs.Range("D" & i).Value   ' Replacement ID
            swapWs.Range("J" & i).Value = srcWs.Range("E" & i).Value   ' Swap Date
        End If
    Next i

    Debug.Print "[TX_Swap] MapToWorkbook: Bulk data mapped to Swap sheet"
    Debug.Print "[TX_Swap] MapToWorkbook: Complete - 7 header fields + bulk data mapped"
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Sub AddHyperlinkPreserveStyle(ws As Worksheet, cellAddr As String, url As String)
    ' Excel's Hyperlinks.Add changes styling - this preserves the original look
    Dim cell As Range
    Set cell = ws.Range(cellAddr)

    ' Store original formatting
    Dim originalText As String
    Dim originalFontColor As Long
    Dim originalFontName As String
    Dim originalFontSize As Single
    Dim originalFontBold As Boolean
    Dim originalUnderline As XlUnderlineStyle

    originalText = cell.Value
    originalFontColor = cell.Font.Color
    originalFontName = cell.Font.Name
    originalFontSize = cell.Font.Size
    originalFontBold = cell.Font.Bold
    originalUnderline = cell.Font.Underline

    ' Add hyperlink
    ws.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:=originalText

    ' Restore original formatting (hyperlinks change font color and add underline)
    With cell.Font
        .Color = originalFontColor
        .Name = originalFontName
        .Size = originalFontSize
        .Bold = originalFontBold
        .Underline = originalUnderline
    End With
End Sub
