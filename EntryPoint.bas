Attribute VB_Name = "EntryPoint"
Option Explicit

' ============================================================================
' ENTRY POINT - Main UI controller for Usage Workbook
' ============================================================================
' I designed this as the central orchestrator - handles UI setup, button events,
' and coordinates the transaction workflow: validate -> copy template -> map data -> save
' ============================================================================

Public Sub Main()
    ' Main entry - I set up the workbook UI and event handlers from scratch here
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    ' Defensive: get or create Sheet1 - handles both fresh workbooks and existing ones
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(1)
        ws.Name = "Sheet1"
    End If
    On Error GoTo 0

    ' Activate the sheet
    ws.Activate

    ' Had to add this - sheet protection from previous runs was blocking my setup code
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' Remove gridlines
    ActiveWindow.DisplayGridlines = False

    ' Dark navy background - matches our corporate branding and makes the white form pop
    ws.Cells.Interior.Color = RGB(0, 0, 51)

    ' Set column widths
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:D").ColumnWidth = 15
    ws.Columns("E:F").ColumnWidth = 11
    ws.Columns("H").ColumnWidth = 20

    ' Header title - styled to match our existing Excel templates
    With ws.Range("B2:D2")
        .Merge
        .Value = "Usage Workbook"
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(0, 32, 96)
        .Font.Name = "Cambria"
        .Font.Size = 18
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(100, 180, 255)
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(128, 128, 128)
        .Borders(xlEdgeRight).Weight = xlThin
    End With

    ' Button area - kept white to make the Submit button stand out
    With ws.Range("E2:F2")
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(100, 180, 255)
        .Borders(xlEdgeTop).Weight = xlThick
    End With

    ' H2 white fill for Clear button area (matches header)
    With ws.Range("H2")
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(100, 180, 255)
        .Borders(xlEdgeTop).Weight = xlThick
    End With

    ' White box from B4 to F39
    ws.Range("B4:F39").Interior.Color = RGB(255, 255, 255)

    ' B4 label
    With ws.Range("B4")
        .Value = "Order Type"
        .Font.Name = "Cambria"
        .Font.Bold = True
        .Font.Color = RGB(0, 32, 96)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Order Type dropdown - this drives the entire form layout via Dispatcher
    With ws.Range("C4:F4")
        .Merge
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="New Usage,Return,Swap"
        .Validation.InCellDropdown = True
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Name = "Cambria"
        .Font.Color = RGB(0, 32, 96)
    End With

    ' Top border for B4:F4 row
    With ws.Range("B4:F4").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RGB(100, 180, 255)
        .Weight = xlThick
    End With

    ' Blue-grey fill for B5:F5
    ws.Range("B5:F5").Interior.Color = RGB(100, 120, 150)

    ' Submit button - I'm using a Shape here instead of Form Control for better styling
    Dim shp As Shape

    ' Fixed: delete any existing button first to prevent duplicates after RESET
    On Error Resume Next
    ws.Shapes("btnSubmit").Delete
    On Error GoTo 0

    ' Add rounded rectangle shape in E2:F2 area
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("E2").Left + 5, _
        ws.Range("E2").Top + 3, _
        ws.Range("E2:F2").Width - 10, _
        ws.Range("E2").Height - 6)

    With shp
        .Name = "btnSubmit"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Submit"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 32, 96)
        .TextFrame2.TextRange.Font.Name = "Cambria"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "EntryPoint.OnSubmit"
    End With

    ' Create Clear button as Shape in H2
    On Error Resume Next
    ws.Shapes("btnClear").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("H2").Left + 5, _
        ws.Range("H2").Top + 3, _
        ws.Range("H2").Width - 10, _
        ws.Range("H2").Height - 6)

    With shp
        .Name = "btnClear"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Clear"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 32, 96)
        .TextFrame2.TextRange.Font.Name = "Cambria"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "EntryPoint.OnClear"
    End With

    ' Build secondary info box (H5:I12)
    BuildInfoBox

    ' Build links box (H15:I16)
    BuildLinksBox

    Application.ScreenUpdating = True

    ' Fixed: RESET sometimes leaves events disabled, which breaks the dropdown
    Application.EnableEvents = True

    ' Had to inject the event handler programmatically since VBA doesn't auto-wire sheet events
    InstallChangeHandler

    Debug.Print "[EntryPoint] Main complete - Sheet1 configured"
End Sub

Private Sub InstallChangeHandler()
    ' VBA quirk: can't bind sheet events from a standard module, so I inject the handler code directly
    Const vbext_ct_Document As Long = 100
    Dim vbComp As Object
    Dim codeMod As Object
    Dim eventCode As String
    Dim lineNum As Long

    On Error GoTo ErrorHandler

    ' Find Sheet1's code module
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_Document And vbComp.Name = "Sheet1" Then
            Set codeMod = vbComp.CodeModule
            Exit For
        End If
    Next vbComp

    If codeMod Is Nothing Then
        Debug.Print "[EntryPoint] Could not find Sheet1 code module"
        Exit Sub
    End If

    ' Check if event handler already exists
    On Error Resume Next
    lineNum = codeMod.ProcStartLine("Worksheet_Change", 0)
    On Error GoTo ErrorHandler

    If lineNum > 0 Then
        Debug.Print "[EntryPoint] Worksheet_Change handler already exists"
        Exit Sub
    End If

    ' Build the event handler code
    eventCode = vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    ' Respond to Order Type dropdown changes in C4" & vbCrLf & _
        "    If Not Intersect(Target, Me.Range(""C4"")) Is Nothing Then" & vbCrLf & _
        "        If Target.Value <> """" Then" & vbCrLf & _
        "            Dispatcher.HandleOrderTypeChange Target.Value" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf

    ' Add at end of module
    codeMod.InsertLines codeMod.CountOfLines + 1, eventCode

    Debug.Print "[EntryPoint] Installed Worksheet_Change event handler"
    Exit Sub

ErrorHandler:
    Debug.Print "[EntryPoint] Error installing event handler: " & Err.Description
End Sub

Public Sub BuildInfoBox(Optional orderType As String = "")
    ' I designed this info panel to dynamically adjust based on order type
    ' Shows customer info, lookups, and computed values (H=Label, I=Value)
    Debug.Print "[EntryPoint] BuildInfoBox: Creating info panel H5:I12 for orderType=" & orderType
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Clean slate - previous order type may have had different row count
    With ws.Range("H5:L13")
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = RGB(0, 0, 51)
        .Borders.LineStyle = xlNone
    End With

    ' Set column widths for H-I
    ws.Columns("H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 20

    ' Header row (H5:I5)
    With ws.Range("H5:I5")
        .Merge
        .Value = "Workbook Info"
        .Interior.Color = RGB(100, 120, 150)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Bookman Old Style"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Each order type needs different fields - I sized the box to fit exactly
    Dim lastRow As Long
    Select Case orderType
        Case "Swap"
            lastRow = 13 ' Swap has the most: Equipment Type, Model, Quantity
        Case "Return"
            lastRow = 12 ' Return needs Model + Quantity
        Case Else
            lastRow = 11 ' New Usage is the baseline
    End Select

    ' Label column styling (H6:H{lastRow})
    With ws.Range("H6:H" & lastRow)
        .Interior.Color = RGB(220, 220, 220)
        .Font.Name = "Bookman Old Style"
    End With

    ' Value column styling (I6:I{lastRow})
    With ws.Range("I6:I" & lastRow)
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Bookman Old Style"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Border around entire box
    With ws.Range("H5:I" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    ' Structured differently per transaction - business requirements drove these field choices
    Select Case orderType
        Case "Swap"
            ' Swap needs equipment type to determine filename format
            ws.Range("H6").Value = "Customer Name"
            ws.Range("H7").Value = "Customer #"
            ws.Range("H8").Value = "City"
            ws.Range("H9").Value = "State"
            ws.Range("H10").Value = "VSimple Id"
            ws.Range("H11").Value = "Equipment Type"
            ws.Range("H12").Value = "Model"
            ws.Range("H13").Value = "Quantity"

        Case "Return"
            ' Return derives customer from serial number lookup - different source than New Usage
            ws.Range("H6").Value = "Customer Name"
            ws.Range("H7").Value = "Customer #"
            ws.Range("H8").Value = "City"
            ws.Range("H9").Value = "State"
            ws.Range("H10").Value = "VSimple Id"
            ws.Range("H11").Value = "Model"
            ws.Range("H12").Value = "Quantity"

        Case Else
            ' New Usage - user enters Customer # directly, lookups fill the rest
            ws.Range("H6").Value = "Customer Name"
            ws.Range("H7").Value = "City"
            ws.Range("H8").Value = "State"
            ws.Range("H9").Value = "VSimple Id"
            ws.Range("H10").Value = "Model"
            ws.Range("H11").Value = "Quantity"
    End Select

    ' Set up INDEX/MATCH formulas - these auto-populate as user enters data
    Dim dbSheet As Worksheet
    On Error Resume Next
    Set dbSheet = ThisWorkbook.Worksheets("CustomerDB")
    On Error GoTo 0

    Select Case orderType
        Case "Return"
            ' Return workflow: serial number -> CRDB lookup -> customer info
            ' Customer # derived from CRDB via Serial Number
            ws.Range("I7").Formula = "=IFERROR(IF(B12="""","""",INDEX(CRDB!C:C,MATCH(B12,CRDB!X:X,0))),"""")"

            ' Chain lookup: Customer # -> CustomerDB -> Name/City/State
            ws.Range("I6").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!M:M,MATCH(I7,CustomerDB!A:A,0))),"""")"
            ws.Range("I8").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!E:E,MATCH(I7,CustomerDB!A:A,0))),"""")"
            ws.Range("I9").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!F:F,MATCH(I7,CustomerDB!A:A,0))),"""")"

            ' Extract VSimple ID from URL - needed for filename
            ws.Range("I10").Formula = "=IFERROR(IF(C6<>"""",MID(C6,FIND(""*"",SUBSTITUTE(C6,""/"",""*"",LEN(C6)-LEN(SUBSTITUTE(C6,""/"",""""))))+1,100),""""),"""")"

            ' Model display logic: if mixed equipment types, show "Various" instead of first item
            ws.Range("I11").Formula = "=IFERROR(IF(R11=""POWER"",IF(R10>1,""Various Power"",INDEX(CRDB!T:T,MATCH(B12,CRDB!X:X,0))),IF(R10>1,""Various Equipment"",IF(R9>1,""Various Equipment"",INDEX(CRDB!T:T,MATCH(B12,CRDB!X:X,0))))),"""")"

            ' Quantity only shows for single-model returns, hidden for "Various"
            ws.Range("I12").Formula = "=IFERROR(IF(OR(I11=""Various Power"",I11=""Various Equipment"",R12=""""),"""",COUNTIF(R12:R311,R12)),"""")"

        Case "Swap"
            ' Swap workflow: Dealer ID -> CRDB lookup -> customer info
            ' Dealer ID is the key field for Swap transactions
            ws.Range("I7").Formula = "=IFERROR(IF(B13="""","""",INDEX(CRDB!C:C,MATCH(B13,CRDB!W:W,0))),"""")"

            ' Chain lookup for customer details
            ws.Range("I6").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!M:M,MATCH(I7,CustomerDB!A:A,0))),"""")"
            ws.Range("I8").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!E:E,MATCH(I7,CustomerDB!A:A,0))),"""")"
            ws.Range("I9").Formula = "=IFERROR(IF(I7="""","""",INDEX(CustomerDB!F:F,MATCH(I7,CustomerDB!A:A,0))),"""")"

            ' Extract VSimple ID from URL for filename construction
            ws.Range("I10").Formula = "=IFERROR(IF(C6<>"""",MID(C6,FIND(""*"",SUBSTITUTE(C6,""/"",""*"",LEN(C6)-LEN(SUBSTITUTE(C6,""/"",""""))))+1,100),""""),"""")"

            ' Equipment Type drives the filename format (BATT=Power Swap, CHGR=Charger Swap)
            ws.Range("I11").Formula = "=IFERROR(IF(B13="""","""",INDEX(CRDB!S:S,MATCH(B13,CRDB!W:W,0))),"""")"

            ' Model lookup - only shown for non-battery/charger equipment
            ws.Range("I12").Formula = "=IFERROR(IF(B13="""","""",INDEX(CRDB!T:T,MATCH(B13,CRDB!W:W,0))),"""")"

            ' Count all entered Dealer IDs for the quantity
            ws.Range("I13").Formula = "=IFERROR(IF(B13="""","""",COUNTA(B13:B312)),"""")"

        Case Else
            ' New Usage: direct entry of Customer # in C7 triggers lookups
            If Not dbSheet Is Nothing Then
                ' CustomerDB lookups: A=Customer#, M=Name, E=City, F=State
                ws.Range("I6").Formula = "=IFERROR(IF(C7="""","""",INDEX(CustomerDB!M:M,MATCH(C7,CustomerDB!A:A,0))),"""")"
                ws.Range("I7").Formula = "=IFERROR(IF(C7="""","""",INDEX(CustomerDB!E:E,MATCH(C7,CustomerDB!A:A,0))),"""")"
                ws.Range("I8").Formula = "=IFERROR(IF(C7="""","""",INDEX(CustomerDB!F:F,MATCH(C7,CustomerDB!A:A,0))),"""")"
            End If
            ' Parse VSimple ID from URL for tracking
            ws.Range("I9").Formula = "=IFERROR(IF(C6<>"""",MID(C6,FIND(""*"",SUBSTITUTE(C6,""/"",""*"",LEN(C6)-LEN(SUBSTITUTE(C6,""/"",""""))))+1,100),""""),"""")"
            ' Model comes from PODB based on Truck PO number
            ws.Range("I10").Formula = "=IFERROR(IF(C18="""","""",INDEX(PODB!D:D,MATCH(C18,PODB!I:I,0))),"""")"
            ' Quantity in parens - matches our naming convention
            ws.Range("I11").Formula = "=IFERROR(IF(C18="""","""",""(""&COUNTIFS(PODB!I:I,C18,PODB!D:D,I10)&"")""),"""")"
    End Select
End Sub

Private Sub BuildLinksBox()
    ' Built this to output a markdown link for pasting into our ticketing system
    ' After submit, I16 contains something like [UsageWorkbook](https://...)
    Debug.Print "[EntryPoint] BuildLinksBox: Creating links panel H15:I16"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Reset in case previous layout was different
    With ws.Range("H15:I18")
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = RGB(0, 0, 51)
        .Borders.LineStyle = xlNone
    End With

    ' Clean up any leftover buttons from earlier versions
    On Error Resume Next
    ws.Shapes("btnCopyLink").Delete
    On Error GoTo 0

    ' Header row (H15:I15)
    With ws.Range("H15:I15")
        .Merge
        .Value = "Links"
        .Interior.Color = RGB(100, 120, 150)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Bookman Old Style"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Label column styling (H16)
    With ws.Range("H16")
        .Interior.Color = RGB(220, 220, 220)
        .Font.Name = "Bookman Old Style"
        .Value = "Workbook"
    End With

    ' Value column styling (I16)
    With ws.Range("I16")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Bookman Old Style"
    End With

    ' Border around entire box
    With ws.Range("H15:I16").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With
End Sub

Public Sub OnSubmit()
    ' Main workflow trigger - validates form, then runs ProcessTransaction
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim orderType As String
    orderType = ws.Range("C4").Value
    Debug.Print "[EntryPoint] OnSubmit: Order type = " & orderType

    If orderType = "" Then
        MsgBox "Please select an Order Type first.", vbExclamation
        Exit Sub
    End If

    ' Each transaction type has its own validation and filename logic
    Dim fileName As String

    Select Case orderType
        Case "New Usage"
            Debug.Print "[EntryPoint] OnSubmit: Validation " & IIf(TX_NewUsage.Validate(), "PASSED", "FAILED")
            If TX_NewUsage.Validate() Then
                fileName = TX_NewUsage.GetFileName()
                Debug.Print "[EntryPoint] OnSubmit: Filename = " & fileName

                ' Bail early if validation passed but filename couldn't be built
                If fileName = "" Then Exit Sub

                ProcessTransaction orderType, _
                                   TX_NewUsage.GetTemplate(), _
                                   fileName, _
                                   TX_NewUsage.GetCustomerName()
            End If

        Case "Return"
            Debug.Print "[EntryPoint] OnSubmit: Validation " & IIf(TX_Return.Validate(), "PASSED", "FAILED")
            If TX_Return.Validate() Then
                fileName = TX_Return.GetFileName()
                Debug.Print "[EntryPoint] OnSubmit: Filename = " & fileName

                ' Bail early if validation passed but filename couldn't be built
                If fileName = "" Then Exit Sub

                ProcessTransaction orderType, _
                                   TX_Return.GetTemplate(), _
                                   fileName, _
                                   TX_Return.GetCustomerName()
            End If

        Case "Swap"
            Debug.Print "[EntryPoint] OnSubmit: Validation " & IIf(TX_Swap.Validate(), "PASSED", "FAILED")
            If TX_Swap.Validate() Then
                fileName = TX_Swap.GetFileName()
                Debug.Print "[EntryPoint] OnSubmit: Filename = " & fileName

                ' Bail early if validation passed but filename couldn't be built
                If fileName = "" Then Exit Sub

                ProcessTransaction orderType, _
                                   TX_Swap.GetTemplate(), _
                                   fileName, _
                                   TX_Swap.GetCustomerName()
            End If

        Case Else
            MsgBox "Unknown order type: " & orderType, vbInformation
    End Select
End Sub

Public Sub OnClear()
    ' Reset form to blank state - lets user start fresh without running RESET
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Debug.Print "[EntryPoint] Clear button clicked"

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next

    ' Wipe customer info - covers all transaction types
    ws.Range("C6:F10").ClearContents

    ' Clear bulk data entry area (Return/Swap can have 300 rows of serials)
    ws.Range("B12:F311").ClearContents

    ' Clear the markdown link from previous export
    ws.Range("I16").ClearContents

    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Debug.Print "[EntryPoint] Form cleared"
End Sub

Private Sub ProcessTransaction(orderType As String, templateName As String, fileName As String, customerName As String)
    ' My transaction flow: validate -> copy template -> map data -> save
    ' This is the core workflow that all three order types share
    On Error GoTo ErrorHandler
    Debug.Print "[EntryPoint] ProcessTransaction: Starting for " & orderType

    Dim destPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim city As String
    Dim customerNumber As String

    Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' Source cells differ by order type - New Usage has direct entry, others derive from lookups
    Select Case orderType
        Case "New Usage"
            city = PathHelper.SafeCellValue(ws.Range("I7"))
            customerNumber = PathHelper.SafeCellValue(ws.Range("C7"))
        Case "Return", "Swap"
            city = PathHelper.SafeCellValue(ws.Range("I8"))
            customerNumber = PathHelper.SafeCellValue(ws.Range("I7"))
    End Select

    ' Build nested folder structure: Year/Customer/Location/OrderType/
    Debug.Print "[EntryPoint] ProcessTransaction: Step 1 - Building path..."
    destPath = PathHelper.BuildExportPath(customerName, city, customerNumber, orderType, fileName)
    Debug.Print "[EntryPoint] ProcessTransaction: Path = " & destPath
    If destPath = "" Then
        MsgBox "Could not create destination path. Please check folder permissions.", _
               vbCritical, "Path Error"
        Exit Sub
    End If

    ' Prevent accidental overwrites - give user options if file exists
    Debug.Print "[EntryPoint] ProcessTransaction: Step 2 - Checking existing file..."
    If Not FileHelper.HandleExistingFile(destPath) Then
        Exit Sub  ' User chose to open existing or cancel
    End If

    ' FileCopy is faster than Open/SaveAs and avoids file lock issues
    Debug.Print "[EntryPoint] ProcessTransaction: Step 3 - Copying template..."
    Set wb = FileHelper.CopyTemplate(templateName, destPath)
    If wb Is Nothing Then Exit Sub  ' Error already shown by CopyTemplate

    ' Each TX module knows its own field mappings
    Debug.Print "[EntryPoint] ProcessTransaction: Step 4 - Mapping data..."
    Select Case orderType
        Case "New Usage"
            TX_NewUsage.MapToWorkbook wb
        Case "Return"
            TX_Return.MapToWorkbook wb
        Case "Swap"
            TX_Swap.MapToWorkbook wb
    End Select

    ' Save and close before generating link (ensures file exists on SharePoint)
    Debug.Print "[EntryPoint] ProcessTransaction: Step 5 - Saving..."
    FileHelper.SaveAndClose wb

    ' Generate markdown link for pasting into ticketing system
    Dim sharePointUrl As String
    sharePointUrl = SharePointHelper.GetSharePointUrl(destPath)
    If sharePointUrl <> "" Then
        ws.Range("I16").Value = "[UsageWorkbook](" & sharePointUrl & ")"
        Debug.Print "[EntryPoint] Markdown link: [UsageWorkbook](" & sharePointUrl & ")"
    Else
        ' Fallback to local path if OneDrive sync not found
        ws.Range("I16").Value = destPath
        Debug.Print "[EntryPoint] Could not generate SharePoint URL, using local path"
    End If

    ' Success message
    MsgBox "Workbook created successfully!" & vbCrLf & vbCrLf & destPath, _
           vbInformation, "Success"

    Debug.Print "[EntryPoint] ProcessTransaction: SUCCESS - " & destPath
    Exit Sub

ErrorHandler:
    MsgBox "Error processing transaction:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Processing Error"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
End Sub
