Attribute VB_Name = "FileHelper"
Option Explicit

' ============================================================================
' FILE HELPER - Safe file operations with proper error handling
' ============================================================================
' Validates templates exist, handles the "file already exists" case,
' and wraps FileCopy/Save with meaningful error messages.
' ============================================================================

' Fail fast if template is missing - better than cryptic error later
Public Function TemplateExists(templateName As String) As Boolean
    Dim fullPath As String
    fullPath = Config.GetTemplatePath(templateName)
    TemplateExists = (Dir(fullPath) <> "")
    Debug.Print "[FileHelper] TemplateExists: " & templateName & " = " & TemplateExists
End Function

' Give user options when file already exists - prevents accidental overwrites
Public Function HandleExistingFile(fullPath As String) As Boolean
    ' If file doesn't exist, proceed normally
    If Dir(fullPath) = "" Then
        HandleExistingFile = True
        Exit Function
    End If

    Debug.Print "[FileHelper] HandleExistingFile: File exists at " & fullPath
    ' Offer choices: open existing file, open containing folder, or cancel
    Dim choice As VbMsgBoxResult
    choice = MsgBox("This workbook already exists:" & vbCrLf & vbCrLf & _
                    fullPath & vbCrLf & vbCrLf & _
                    "Yes = Open the file" & vbCrLf & _
                    "No = Open the folder" & vbCrLf & _
                    "Cancel = Do nothing", _
                    vbYesNoCancel + vbQuestion, "File Already Exists")

    Debug.Print "[FileHelper] HandleExistingFile: User chose " & choice
    Select Case choice
        Case vbYes
            ' Open the existing workbook
            On Error Resume Next
            Workbooks.Open fullPath
            If Err.Number <> 0 Then
                MsgBox "Could not open file: " & Err.Description, vbExclamation
            End If
            On Error GoTo 0

        Case vbNo
            ' Open Explorer with file pre-selected
            On Error Resume Next
            Shell "explorer.exe /select,""" & fullPath & """", vbNormalFocus
            On Error GoTo 0
    End Select

    HandleExistingFile = False  ' User made a choice, don't create new file
End Function

' FileCopy is faster than Open/SaveAs and avoids file lock issues
Public Function CopyTemplate(templateName As String, destPath As String) As Workbook
    On Error GoTo ErrorHandler

    Dim srcPath As String
    srcPath = Config.GetTemplatePath(templateName)
    Debug.Print "[FileHelper] CopyTemplate: Source = " & srcPath
    Debug.Print "[FileHelper] CopyTemplate: Dest = " & destPath

    ' Fail fast if template is missing
    If Dir(srcPath) = "" Then
        MsgBox "Template not found:" & vbCrLf & vbCrLf & srcPath, _
               vbCritical, "Template Missing"
        Set CopyTemplate = Nothing
        Exit Function
    End If

    ' Security: ensure we're not writing outside the allowed folder
    If Not PathHelper.IsPathSafe(destPath, Config.GetExportRoot()) Then
        MsgBox "Invalid destination path. Operation blocked for security.", _
               vbCritical, "Security Error"
        Set CopyTemplate = Nothing
        Exit Function
    End If

    ' Use FileCopy (faster than Open/SaveAs, no file lock issues)
    Debug.Print "[FileHelper] CopyTemplate: FileCopy executing..."
    FileCopy srcPath, destPath

    ' Open the copied workbook for data population
    Set CopyTemplate = Workbooks.Open(destPath)
    Debug.Print "[FileHelper] CopyTemplate: Success, workbook opened"
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 70
            MsgBox "Permission denied." & vbCrLf & vbCrLf & _
                   "The file may be open in another program, or you may lack " & _
                   "write access to this location.", vbCritical, "Access Denied"
        Case 76
            MsgBox "Path not found." & vbCrLf & vbCrLf & _
                   "Please verify your network connection and that the " & _
                   "destination folder exists.", vbCritical, "Path Error"
        Case 53
            MsgBox "Template file not found." & vbCrLf & vbCrLf & _
                   srcPath, vbCritical, "File Not Found"
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, _
                   vbCritical, "File Copy Error"
    End Select
    Set CopyTemplate = Nothing
End Function

' Save first, then close - ensures data is written before releasing handle
Public Sub SaveAndClose(wb As Workbook)
    On Error Resume Next

    If Not wb Is Nothing Then
        Debug.Print "[FileHelper] SaveAndClose: Saving " & wb.Name
        wb.Save
        wb.Close SaveChanges:=False  ' Already saved above
        Debug.Print "[FileHelper] SaveAndClose: Complete"
    End If

    On Error GoTo 0
End Sub

' Get file extension from filename
Public Function GetFileExtension(fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetFileExtension = Mid(fileName, dotPos)
    Else
        GetFileExtension = ""
    End If
End Function

