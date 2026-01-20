Attribute VB_Name = "TX_NewUsage"
Option Explicit

' Order Type: New Usage

Public Sub Build()
    ' I designed the form layout to match our existing New Usage templates
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

    ' Label column gets a subtle grey to visually separate from input area
    ws.Range("B6:B39").Interior.Color = RGB(220, 220, 220)

    ' Input cells stay white for clear data entry
    ws.Range("C6:F39").Interior.Color = RGB(255, 255, 255)

    ' Merge C:F for each input row - gives more room for long values
    Dim mergeRows As Variant
    Dim mr As Variant
    mergeRows = Array(6, 7, 8, 9, 10, 11, 13, 14, 15, 17, 18, 19, 20, 21, 23, 24, 25, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 38, 39)
    For Each mr In mergeRows
        ws.Range("C" & mr & ":F" & mr).Merge
    Next mr

    ' Left-align with indent looks cleaner than centered text
    ws.Range("C6:F39").HorizontalAlignment = xlLeft
    ws.Range("C6:F39").IndentLevel = 1

    ' Amount fields get accounting format for currency display
    Dim amountCells As Variant
    amountCells = Array("C28", "C30", "C32", "C34", "C35")
    Dim ac As Variant
    For Each ac In amountCells
        With ws.Range(CStr(ac))
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Font.Name = "Bookman Old Style"
            .Value = 0
        End With
    Next ac

    ' Date validation prevents typos - common issue with manual entry
    With ws.Range("C13")
        .NumberFormat = "m/d/yyyy"
        .Validation.Delete
        .Validation.Add Type:=xlValidateDate, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlGreater, _
            Formula1:="1/1/1900"
        .Validation.ErrorMessage = "Please enter a valid date."
    End With

    ' URC is deprecated for new deals - grey it out so users skip it
    With ws.Range("C15")
        .Value = "NA"
        .Interior.Color = RGB(220, 220, 220)
    End With

    ' Column B labels
    ws.Range("B6").Value = "V simple Link"
    ws.Range("B7").Value = "Customer #"
    ws.Range("B8").Value = "On Site Contact"
    ws.Range("B9").Value = "Phone"
    ws.Range("B10").Value = "Email"
    ws.Range("B11").Value = "New Customer"
    ws.Range("B12").Value = "Order Information"
    ws.Range("B13").Value = "Closed/Won Date in CRM"
    ws.Range("B14").Value = "CRM Opportunity Number"
    ws.Range("B15").Value = "URC"
    ws.Range("B16").Value = "Equipment Information"
    ws.Range("B17").Value = "Stock Equipment"
    ws.Range("B18").Value = "Truck PO"
    ws.Range("B19").Value = "Battery PO"
    ws.Range("B20").Value = "Charger PO"
    ws.Range("B21").Value = "Non-Raymond PO"
    ws.Range("B22").Value = "Agreement Information"
    ws.Range("B23").Value = "Term"
    ws.Range("B24").Value = "Margin"
    ws.Range("B25").Value = "Freight Included"
    ws.Range("B26").Value = "Maintenance Information"
    ws.Range("B27").Value = "Type of Maintenance"
    ws.Range("B28").Value = "Maint Amount"
    ws.Range("B29").Value = "Battery Watering"
    ws.Range("B30").Value = "Battery Watering Amount"
    ws.Range("B31").Value = "Battery Maintenance"
    ws.Range("B32").Value = "Battery Maintenance Amount"
    ws.Range("B33").Value = "Charger Maintenance"
    ws.Range("B34").Value = "Charger Maintenance Amount"
    ws.Range("B35").Formula = "=IF(C27=""SM"",""SM Rate"","""")"
    ws.Range("B36").Formula = "=IF(C27=""SM"",""SM Frequency"","""")"
    ws.Range("B37").Value = "Return Information"
    ws.Range("B38").Value = "Has Return"
    ws.Range("B39").Value = "Return V Simple Link"

    ' Bookman Old Style matches our corporate template style
    ws.Range("B6:B39").Font.Name = "Bookman Old Style"

    ' Subtle border separates labels from input area
    With ws.Range("B5:B39").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    ' Row separators make the form easier to scan
    With ws.Range("B5:F39").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With
    ' Bottom edge of last row
    With ws.Range("B39:F39").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(140, 140, 140)
        .Weight = xlThin
    End With

    ' Section headers break up the long form into logical groups
    Dim headerRows As Variant
    Dim r As Variant
    headerRows = Array(12, 16, 22, 26, 37)
    For Each r In headerRows
        With ws.Range("B" & r & ":F" & r)
            .Interior.Color = RGB(100, 120, 150)
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Calibri"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next r

    ' Dropdowns prevent typos and ensure data consistency
    With ws.Range("C11")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Yes,No"
    End With

    With ws.Range("C17")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Yes,No"
    End With

    With ws.Range("C24")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Full,Reduced,Enhanced Reduced,Full +"
    End With

    With ws.Range("C25")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Yes,No"
    End With

    With ws.Range("C27")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="CFPM,SM"
    End With

    With ws.Range("C29")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Bi Weekly,Monthly"
    End With

    With ws.Range("C31")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Semi Annual,Quarterly"
    End With

    With ws.Range("C33")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Semi Annual,Quarterly"
    End With

    With ws.Range("C38")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Yes,No"
    End With

    Debug.Print "TX_NewUsage.Build called"
End Sub

' ============================================================================
' VALIDATION & DATA FUNCTIONS
' ============================================================================

Public Function Validate() As Boolean
    ' Catches: empty URL, missing customer #, failed lookups
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Debug.Print "[TX_NewUsage] Validate: Checking V Simple Link..."
    Debug.Print "[TX_NewUsage] Validate: V Simple Link = " & ws.Range("C6").Value
    ' V Simple Link is required - we extract the ID for the filename
    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter a V Simple Link.", vbExclamation, "Validation Error"
        Validate = False
        Exit Function
    End If

    ' Must be a valid URL structure for ID extraction to work
    If InStr(ws.Range("C6").Value, "/") = 0 Then
        MsgBox "V Simple Link must be a valid URL with an ID." & vbCrLf & _
               "Example: https://vsimple.com/deals/12345", vbExclamation, "Invalid URL"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_NewUsage] Validate: Checking Customer #..."
    Debug.Print "[TX_NewUsage] Validate: Customer # = " & ws.Range("C7").Value
    ' Customer # drives all the lookups - can't proceed without it
    If Trim(ws.Range("C7").Value) = "" Then
        MsgBox "Please enter a Customer #.", vbExclamation, "Validation Error"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_NewUsage] Validate: Customer Name lookup = " & ws.Range("I6").Value
    ' Gracefully handle #N/A, #REF! etc. - common when customer # is wrong
    If IsError(ws.Range("I6").Value) Then
        MsgBox "Customer Name lookup failed. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    ' Empty result means no match found in CustomerDB
    If Trim(CStr(ws.Range("I6").Value)) = "" Then
        MsgBox "Customer Name lookup returned empty. Please verify the Customer # is correct.", _
               vbExclamation, "Lookup Error"
        Validate = False
        Exit Function
    End If

    Debug.Print "[TX_NewUsage] Validate: PASSED"
    Validate = True
End Function

Public Function GetTemplate() As String
    ' My template selection logic - 8 variants from 3 boolean flags (Kehe, Return, Stock)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim hasReturn As Boolean
    Dim isStock As Boolean
    Dim isKehe As Boolean

    ' Pull flags from form - these drive which template variant we use
    hasReturn = (UCase(Trim(ws.Range("C38").Value)) = "YES")
    isStock = (UCase(Trim(ws.Range("C17").Value)) = "YES")

    ' Kehe is auto-detected from customer name - they have special template requirements
    isKehe = IsKeheCustomer()

    Debug.Print "[TX_NewUsage] GetTemplate: hasReturn=" & hasReturn & ", isStock=" & isStock & ", isKehe=" & isKehe
    ' Binary decision tree: 2^3 = 8 possible template combinations
    If hasReturn And isStock And isKehe Then
        GetTemplate = TPL_NEWUSAGE_RETURN_STOCK_KEHE
    ElseIf hasReturn And isStock Then
        GetTemplate = TPL_NEWUSAGE_RETURN_STOCK
    ElseIf hasReturn And isKehe Then
        GetTemplate = TPL_NEWUSAGE_RETURN_KEHE
    ElseIf isStock And isKehe Then
        GetTemplate = TPL_NEWUSAGE_STOCK_KEHE
    ElseIf hasReturn Then
        GetTemplate = TPL_NEWUSAGE_RETURN
    ElseIf isStock Then
        GetTemplate = TPL_NEWUSAGE_STOCK
    ElseIf isKehe Then
        GetTemplate = TPL_NEWUSAGE_KEHE
    Else
        GetTemplate = TPL_NEWUSAGE
    End If
    Debug.Print "[TX_NewUsage] GetTemplate: Selected " & GetTemplate
End Function

Public Function GetFileName() As String
    ' Builds standardized filename: CustomerName_NewUsage_City_Quantity_Model_OppNumber_VSimpleId_UW.xlsm
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim customerName As String
    Dim city As String
    Dim quantity As String
    Dim model As String
    Dim oppNumber As String
    Dim vSimpleId As String

    ' SafeCellValue handles #N/A etc. gracefully - prevents runtime errors
    customerName = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I6")))
    city = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I7")))
    quantity = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I11")))
    model = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("I10")))
    oppNumber = PathHelper.SanitizeName(PathHelper.SafeCellValue(ws.Range("C14")))
    vSimpleId = PathHelper.ExtractVSimpleId(PathHelper.SafeCellValue(ws.Range("C6")))

    Debug.Print "[TX_NewUsage] GetFileName: customerName=" & customerName
    Debug.Print "[TX_NewUsage] GetFileName: city=" & city
    Debug.Print "[TX_NewUsage] GetFileName: vSimpleId=" & vSimpleId

    ' These are required - fail fast with clear message
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

    ' Optional components get placeholders - filename must always be complete
    If city = "" Then city = "NoCity"
    If model = "" Then model = "NoModel"
    If quantity = "" Then quantity = "0"
    If oppNumber = "" Then oppNumber = "NoOpp"

    GetFileName = customerName & "_NewUsage_" & city & "_" & quantity & "_" & _
                  model & "_" & oppNumber & "_" & vSimpleId & "_UW.xlsm"
    Debug.Print "[TX_NewUsage] GetFileName: Result = " & GetFileName
    Exit Function

ErrorHandler:
    Debug.Print "GetFileName error: " & Err.Description
    MsgBox "Error generating filename: " & Err.Description, vbCritical, "Filename Error"
    GetFileName = ""
End Function

Public Function GetCustomerName() As String
    ' I6 is populated by the INDEX/MATCH formula - just read it
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    GetCustomerName = Trim(CStr(ws.Range("I6").Value))
End Function

Public Sub MapToWorkbook(wb As Workbook)
    Dim srcWs As Worksheet
    Dim destWs As Worksheet

    Set srcWs = ThisWorkbook.Worksheets("Sheet1")
    Set destWs = wb.Worksheets("Overview")

    Debug.Print "[TX_NewUsage] MapToWorkbook: Starting field mapping..."
    Debug.Print "[TX_NewUsage] MapToWorkbook: Customer # -> C4 = " & srcWs.Range("C7").Value
    ' Map form fields to template cells - ordering matches Overview sheet layout
    destWs.Range("C4").Value = srcWs.Range("C7").Value    ' Customer #
    destWs.Range("C12").Value = srcWs.Range("C8").Value   ' On Site Contact
    destWs.Range("C13").Value = srcWs.Range("C9").Value   ' Phone
    destWs.Range("C14").Value = srcWs.Range("C10").Value  ' Email
    destWs.Range("C16").Value = srcWs.Range("C11").Value  ' New Customer

    ' Order Information
    destWs.Range("C18").Value = srcWs.Range("C13").Value  ' Closed/Won Date
    destWs.Range("C19").Value = srcWs.Range("C14").Value  ' CRM Opportunity #

    ' Equipment Information
    destWs.Range("C22").Value = srcWs.Range("C17").Value  ' Stock Equipment
    destWs.Range("C24").Value = srcWs.Range("C18").Value  ' Truck PO
    destWs.Range("C25").Value = srcWs.Range("C21").Value  ' Non-Raymond PO
    destWs.Range("C26").Value = srcWs.Range("C19").Value  ' Battery PO
    destWs.Range("C27").Value = srcWs.Range("C20").Value  ' Charger PO

    ' Agreement Information
    destWs.Range("C29").Value = srcWs.Range("C23").Value  ' Term
    destWs.Range("C30").Value = srcWs.Range("C24").Value  ' Margin
    destWs.Range("C31").Value = srcWs.Range("C25").Value  ' Freight Included

    ' Maintenance Information
    destWs.Range("C35").Value = srcWs.Range("C27").Value  ' Type of Maint
    destWs.Range("C36").Value = srcWs.Range("C28").Value  ' Maint Amount
    destWs.Range("C39").Value = srcWs.Range("C31").Value  ' Battery Maint (now row 31)
    destWs.Range("C40").Value = srcWs.Range("C32").Value  ' Battery Maint Amount (now row 32)
    destWs.Range("C41").Value = srcWs.Range("C29").Value  ' Battery Watering (now row 29)
    destWs.Range("C42").Value = srcWs.Range("C30").Value  ' Battery Watering Amount (now row 30)
    destWs.Range("C43").Value = srcWs.Range("C33").Value  ' Charger Maintenance
    destWs.Range("C44").Value = srcWs.Range("C34").Value  ' Charger Maint Amount

    ' Embed V Simple links as hidden hyperlinks on the headers
    ' Users can click through to the deal without searching

    Dim vSimpleUrl As String
    vSimpleUrl = Trim(srcWs.Range("C6").Value)

    If Len(vSimpleUrl) > 0 Then
        Call AddHyperlinkPreserveStyle(destWs, "B2", vSimpleUrl)
        Debug.Print "[TX_NewUsage] MapToWorkbook: V Simple hyperlink added to B2"
    End If

    ' If this is a HasReturn deal, add the return deal link too
    Dim hasReturn As Boolean
    hasReturn = (UCase(Trim(srcWs.Range("C38").Value)) = "YES")

    If hasReturn Then
        Dim returnVSimpleUrl As String
        returnVSimpleUrl = Trim(srcWs.Range("C39").Value)

        If Len(returnVSimpleUrl) > 0 Then
            Call AddHyperlinkPreserveStyle(destWs, "H2", returnVSimpleUrl)
        End If
    End If

    Debug.Print "[TX_NewUsage] MapToWorkbook: Complete - 30 fields mapped"
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function IsKeheCustomer() As Boolean
    ' KeHe has special template requirements - auto-detect to save user a step
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim custName As String
    custName = Trim(ws.Range("I6").Value)

    IsKeheCustomer = (custName = "KeHe Distributors")
    Debug.Print "[TX_NewUsage] IsKeheCustomer: " & custName & " = " & IsKeheCustomer
End Function

Private Sub AddHyperlinkPreserveStyle(ws As Worksheet, cellAddr As String, url As String)
    ' Excel's Hyperlinks.Add changes font color to blue + underline - this restores original
    Dim cell As Range
    Set cell = ws.Range(cellAddr)

    ' Capture current style before adding hyperlink
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

    ' Add the hyperlink (this changes styling)
    ws.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:=originalText

    ' Restore the original look - hyperlink still works when clicked
    With cell.Font
        .Color = originalFontColor
        .Name = originalFontName
        .Size = originalFontSize
        .Bold = originalFontBold
        .Underline = originalUnderline
    End With
End Sub
