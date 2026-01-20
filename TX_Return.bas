Attribute VB_Name = "TX_Return"
Option Explicit

' Order Type: Return

Public Sub Build()
    ' Return form: simple header + 300-row bulk entry area for serial numbers
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
    ws.Range("B6:B10").Interior.Color = RGB(220, 220, 220)
    ws.Range("C6:F10").Interior.Color = RGB(255, 255, 255)

    ' Wider merge for input cells - handles long URLs and contact info
    ws.Range("C6:F6").Merge
    ws.Range("C7:F7").Merge
    ws.Range("C8:F8").Merge
    ws.Range("C9:F9").Merge
    ws.Range("C10:F10").Merge

    ' Column B labels
    ws.Range("B6").Value = "V Simple Link"
    ws.Range("B7").Value = "On Site Contact"
    ws.Range("B8").Value = "Phone"
    ws.Range("B9").Value = "Email"
    ws.Range("B10").Value = "Sales Rep"

    ' Font matches our corporate template style
    ws.Range("B6:B10").Font.Name = "Bookman Old Style"

    ' Subtle borders for clean look
    With ws.Range("B5:B10").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    ' Row separators make scanning easier
    With ws.Range("B5:F10").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With
    ' Bottom edge of last row
    With ws.Range("B10:F10").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With

    ' Equipment table headers - Dealer ID and UC# auto-populate from CRDB
    ws.Range("B11").Value = "Serial Number"
    ws.Range("C11").Value = "Dealer ID"
    ws.Range("D11").Value = "UC#"
    ws.Range("E11:F11").Merge
    ws.Range("E11").Value = "Return Date"

    ' Header row styling matches section headers
    With ws.Range("B11:F11")
        .Interior.Color = RGB(100, 120, 150)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' =========================================================================
    ' DATA ENTRY AREA - Set up 300 rows for bulk serial number entry
    ' =========================================================================

    ' First row always visible so user knows where to start
    With ws.Range("B12:F12")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Bookman Old Style"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Hidden rows reveal as data is entered - keeps form compact
    With ws.Range("B13:F311")
        .Interior.Color = RGB(0, 0, 51)
        .Font.Name = "Bookman Old Style"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Wider Return Date column for date picker visibility
    Dim i As Long
    For i = 12 To 311
        ws.Range("E" & i & ":F" & i).Merge
    Next i

    ' Date validation prevents typos - common issue with manual entry
    With ws.Range("E12:E311").Validation
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
    ' INDEX/MATCH FORMULAS - Auto-populate from CRDB based on Serial Number
    ' =========================================================================

    ' Dealer ID auto-fills when user enters serial - saves manual lookup
    ws.Range("C12:C311").Formula = "=IF(B12="""","""",IFERROR(INDEX(CRDB!W:W,MATCH(B12,CRDB!X:X,0)),""""))"

    ' UC# also auto-fills - both come from CRDB
    ws.Range("D12:D311").Formula = "=IF(B12="""","""",IFERROR(INDEX(CRDB!I:I,MATCH(B12,CRDB!X:X,0)),""""))"

    ' Hidden helper column for equipment type classification
    ' Drives the Model/Quantity display logic in Info Box
    With ws.Range("R12:R311")
        .Formula = "=IF(B12="""","""",IFERROR(INDEX(CRDB!S:S,MATCH(B12,CRDB!X:X,0)),""""))"
        .Interior.Color = RGB(0, 0, 51)  ' Dark navy background (matches sheet)
        .Font.Color = RGB(0, 0, 51)       ' Dark navy font (invisible)
    End With

    ' Count unique truck models - if >1, we show "Various Equipment"
    With ws.Range("R9")
        .Formula = "=IFERROR(COUNTA(UNIQUE(FILTER(R12:R311,(R12:R311<>""BATT"")*(R12:R311<>""CHGR"")*(R12:R311<>"""")))),0)"
        .Interior.Color = RGB(0, 0, 51)
        .Font.Color = RGB(0, 0, 51)
    End With

    ' Count unique power equipment - if >1, we show "Various Power"
    With ws.Range("R10")
        .Formula = "=IFERROR(COUNTA(UNIQUE(FILTER(R12:R311,((R12:R311=""BATT"")+(R12:R311=""CHGR""))*(R12:R311<>"""")))),0)"
        .Interior.Color = RGB(0, 0, 51)
        .Font.Color = RGB(0, 0, 51)
    End With

    ' Classify first item as POWER or TRUCK - drives filename format
    With ws.Range("R11")
        .Formula = "=IF(R12="""","""",IFERROR(IF(OR(R12=""BATT"",R12=""CHGR""),""POWER"",""TRUCK""),""""))"
        .Interior.Color = RGB(0, 0, 51)
        .Font.Color = RGB(0, 0, 51)
    End With

    ' =========================================================================
    ' CONDITIONAL FORMATTING - Zebra stripes reveal rows as data is entered
    ' =========================================================================

    ' Apply to hidden rows only (row 12 is always visible)
    Dim dataRange As Range
    Set dataRange = ws.Range("B13:F311")

    ' Start fresh with no formatting rules
    dataRange.FormatConditions.Delete

    ' Zebra stripes make it easier to track across wide rows
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B13<>"""",MOD(ROW(),2)=1)"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(240, 240, 240)

    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B13<>"""",MOD(ROW(),2)=0)"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)

    ' "Next row" indicator - shows user where to enter next serial
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND($B12<>"""",$B13="""")"
    dataRange.FormatConditions(dataRange.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)

    ' VBA limitation: FormatCondition.Borders doesn't support full styling

    Debug.Print "[TX_Return] Build called - row 12 white, rows 13-311 dark navy with conditional formatting"
End Sub

' ============================================================================
' VALIDATION & DATA FUNCTIONS
' ============================================================================

Public Function Validate() As Boolean
    ' Catches: missing URL, failed customer lookups from serial numbers
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Debug.Print "[TX_Return] Validate: Checking V Simple Link..."
    Debug.Print "[TX_Return] Validate: V Simple Link = " & ws.Range("C6").Value
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

    Debug.Print "[TX_Return] Validate: Customer Name lookup = " & ws.Range("I6").Value
    ' Customer lookup is driven by serial number - if it fails, serial may be wrong
    If IsError(ws.Range("I6").Value) Then
        MsgBox "Customer Name lookup failed. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    ' Empty means serial number not found in CRDB
    If Trim(CStr(ws.Range("I6").Value)) = "" Then
        MsgBox "Customer Name lookup returned empty. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_Return] Validate: PASSED"
    Validate = True
End Function

Public Function GetTemplate() As String
    ' Return only has one template variant (unlike New Usage with 8)
    GetTemplate = TPL_RETURN
    Debug.Print "[TX_Return] GetTemplate: Selected " & GetTemplate
End Function

Public Function GetFileName() As String
    ' Filename format depends on whether we have mixed or single equipment type
    ' Mixed equipment gets shorter name; single model includes quantity and model
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim customerName As String
    Dim customerNumber As String
    Dim model As String
    Dim quantity As String
    Dim vSimpleId As String
    Dim isVarious As Boolean

    ' Pull from Info Box - these are computed from serial number lookups
    customerName = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I6")))
    customerNumber = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I7")))
    model = PathHelper.SafeCellValue(ws.Range("I11"))
    quantity = PathHelper.SafeCellValue(ws.Range("I12"))
    vSimpleId = PathHelper.ExtractVSimpleId(PathHelper.SafeCellValue(ws.Range("C6")))

    Debug.Print "[TX_Return] GetFileName: customerName=" & customerName
    Debug.Print "[TX_Return] GetFileName: customerNumber=" & customerNumber
    Debug.Print "[TX_Return] GetFileName: model=" & model
    Debug.Print "[TX_Return] GetFileName: quantity=" & quantity
    Debug.Print "[TX_Return] GetFileName: vSimpleId=" & vSimpleId

    ' Validate required components
    If customerName = "" Then
        MsgBox "Customer Name not found. Please verify Customer # lookup.", _
               vbExclamation, "Missing Data"
        GetFileName = ""
        Exit Function
    End If

    If vSimpleId = "" Then
        MsgBox "Could not extract VSimple ID from link. Please check the URL format.", _
               vbExclamation, "Invalid URL"
        GetFileName = ""
        Exit Function
    End If

    ' "Various" means mixed equipment - use shorter filename format
    isVarious = (model = "Various Power" Or model = "Various Equipment" Or model = "")

    ' Two filename formats based on equipment mix
    If isVarious Then
        ' Various: CustomerName_Return_CustomerNumber_VSimpleID_UW.xlsm
        GetFileName = customerName & "_Return_" & customerNumber & "_" & vSimpleId & "_UW.xlsm"
    Else
        ' Single model: CustomerName_Return_(Quantity)_Model_CustomerNumber_VSimpleID_UW.xlsm
        model = PathHelper.SanitizeName(model)
        If quantity = "" Then quantity = "1"
        GetFileName = customerName & "_Return_(" & quantity & ")_" & model & "_" & customerNumber & "_" & vSimpleId & "_UW.xlsm"
    End If

    Debug.Print "[TX_Return] GetFileName: Result = " & GetFileName
    Exit Function

ErrorHandler:
    Debug.Print "[TX_Return] GetFileName error: " & Err.Description
    MsgBox "Error generating filename: " & Err.Description, vbCritical, "Filename Error"
    GetFileName = ""
End Function

Public Function GetCustomerName() As String
    ' I6 is computed from serial number -> CRDB -> CustomerDB chain
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    GetCustomerName = Trim(CStr(ws.Range("I6").Value))
End Function

Public Sub MapToWorkbook(wb As Workbook)
    ' Copy form values to the Return template
    Dim srcWs As Worksheet
    Dim destOverview As Worksheet
    Dim destReturn As Worksheet

    Set srcWs = ThisWorkbook.Worksheets("Sheet1")
    Set destOverview = wb.Worksheets("Overview")
    Set destReturn = wb.Worksheets("Return")

    Debug.Print "[TX_Return] MapToWorkbook: Starting field mapping..."

    ' =========================================================================
    ' OVERVIEW SHEET MAPPINGS
    ' =========================================================================
    destOverview.Range("C11").Value = srcWs.Range("C10").Value   ' Sales Rep
    destOverview.Range("C12").Value = srcWs.Range("C7").Value    ' On Site Contact
    destOverview.Range("C13").Value = srcWs.Range("C8").Value    ' Phone
    destOverview.Range("C14").Value = srcWs.Range("C9").Value    ' Email

    Debug.Print "[TX_Return] MapToWorkbook: Overview mappings complete"

    ' =========================================================================
    ' RETURN SHEET MAPPINGS - Dealer IDs (values only, not formulas)
    ' =========================================================================
    destReturn.Range("D13:D313").Value = srcWs.Range("C12:C312").Value

    Debug.Print "[TX_Return] MapToWorkbook: Return sheet Dealer IDs mapped (values only)"

    ' Embed V Simple link on header
    Dim vSimpleUrl As String
    vSimpleUrl = Trim(srcWs.Range("C6").Value)

    If Len(vSimpleUrl) > 0 Then
        Call AddHyperlinkPreserveStyle(destOverview, "B2", vSimpleUrl)
        Debug.Print "[TX_Return] MapToWorkbook: V Simple hyperlink added to B2"
    End If

    Debug.Print "[TX_Return] MapToWorkbook: Complete"
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
