Attribute VB_Name = "PathHelper"
Option Explicit

' ============================================================================
' PATH HELPER - Security-focused path operations
' ============================================================================
' I built this to prevent path injection attacks and handle common edge cases.
' All user input MUST go through SanitizeName before use in paths.
' ============================================================================

' Security: sanitize all user input before using in paths
Public Function SanitizeName(name As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Const INVALID_CHARS As String = "\/:*?""<>|"

    result = Trim(name)

    ' Fixed: VBA's Trim() doesn't catch non-breaking spaces (char 160) from copy-paste
    ' Had to handle this explicitly since users paste from web/Excel constantly
    result = Replace(result, Chr(160), " ")
    result = Trim(result)

    ' Prevents directory traversal attacks (e.g., "../../etc/passwd")
    result = Replace(result, "..", "")
    result = Replace(result, "./", "")
    result = Replace(result, ".\", "")

    ' Strip characters that Windows doesn't allow in filenames
    For i = 1 To Len(INVALID_CHARS)
        c = Mid(INVALID_CHARS, i, 1)
        result = Replace(result, c, "")
    Next i

    ' Windows reserved names cause weird errors - block them
    Dim reserved As Variant
    reserved = Array("CON", "PRN", "AUX", "NUL", _
                     "COM1", "COM2", "COM3", "COM4", "COM5", _
                     "COM6", "COM7", "COM8", "COM9", _
                     "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", _
                     "LPT6", "LPT7", "LPT8", "LPT9")
    Dim r As Variant
    For Each r In reserved
        If UCase(result) = r Then
            result = result & "_safe"
            Exit For
        End If
    Next r

    ' Windows has ~260 char path limit - leave room for folder structure
    If Len(result) > 100 Then result = Left(result, 100)

    ' Clean up any trailing spaces from our processing
    result = Trim(result)

    Debug.Print "[PathHelper] SanitizeName: '" & name & "' -> '" & result & "'"
    SanitizeName = result
End Function

' Prevents directory traversal - path must stay under allowed root
Public Function IsPathSafe(fullPath As String, allowedRoot As String) As Boolean
    ' Normalize for comparison (handle mixed slashes)
    Dim normalizedPath As String
    Dim normalizedRoot As String

    normalizedPath = LCase(Replace(fullPath, "/", "\"))
    normalizedRoot = LCase(Replace(allowedRoot, "/", "\"))

    ' Ensure root ends with backslash
    If Right(normalizedRoot, 1) <> "\" Then normalizedRoot = normalizedRoot & "\"

    ' Path must start with root
    IsPathSafe = (Left(normalizedPath, Len(normalizedRoot)) = normalizedRoot)
    Debug.Print "[PathHelper] IsPathSafe: " & fullPath & " | allowed: " & IsPathSafe
End Function

' Creates folder if missing - used for building export path structure
Public Function EnsureFolder(folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    EnsureFolder = True
    Exit Function

ErrorHandler:
    EnsureFolder = False
    Debug.Print "EnsureFolder error for '" & folderPath & "': " & Err.Description
End Function

' Constructs the full export path following our folder convention
Public Function BuildExportPath(customerName As String, city As String, customerNumber As String, orderType As String, fileName As String) As String
    ' Structure: Transactions\{Year}\{CustomerName}\{City}-{CustomerNumber}\{OrderType}\
    Dim root As String
    Dim yearFolder As String
    Dim custFolder As String
    Dim locationFolder As String
    Dim typeFolder As String

    root = Config.GetExportRoot()
    yearFolder = CStr(Year(Date))
    custFolder = SanitizeName(customerName)
    typeFolder = SanitizeName(orderType)

    ' Customer name is required - can't build path without it
    If custFolder = "" Then
        Debug.Print "BuildExportPath: Empty customer name"
        BuildExportPath = ""
        Exit Function
    End If

    ' Use placeholders for optional fields so path is always complete
    Dim sanitizedCity As String
    Dim sanitizedCustNum As String
    sanitizedCity = SanitizeName(city)
    sanitizedCustNum = SanitizeName(customerNumber)
    If sanitizedCity = "" Then sanitizedCity = "NoCity"
    If sanitizedCustNum = "" Then sanitizedCustNum = "NoCust"
    locationFolder = sanitizedCity & "-" & sanitizedCustNum

    ' Create the nested folder hierarchy (MkDir one level at a time)
    If Not EnsureFolder(root & yearFolder) Then GoTo FolderError
    If Not EnsureFolder(root & yearFolder & "\" & custFolder) Then GoTo FolderError
    If Not EnsureFolder(root & yearFolder & "\" & custFolder & "\" & locationFolder) Then GoTo FolderError
    If Not EnsureFolder(root & yearFolder & "\" & custFolder & "\" & locationFolder & "\" & typeFolder) Then GoTo FolderError

    BuildExportPath = root & yearFolder & "\" & custFolder & "\" & locationFolder & "\" & typeFolder & "\" & fileName
    Debug.Print "[PathHelper] BuildExportPath: " & BuildExportPath
    Exit Function

FolderError:
    Debug.Print "BuildExportPath: Failed to create folder structure"
    BuildExportPath = ""
End Function

' Parses VSimple ID from URL - needed for filename construction
Public Function ExtractVSimpleId(url As String) As String
    ' Grabs everything after the last "/"
    ' Example: "https://vsimple.com/customer/12345" -> "12345"
    On Error Resume Next

    Dim cleanUrl As String
    cleanUrl = Trim(url)

    ' Handle empty input
    If Len(cleanUrl) = 0 Then
        ExtractVSimpleId = ""
        Exit Function
    End If

    ' Remove trailing slash if present
    If Right(cleanUrl, 1) = "/" Then
        cleanUrl = Left(cleanUrl, Len(cleanUrl) - 1)
    End If

    ' Extract text after last "/"
    Dim lastSlash As Long
    lastSlash = InStrRev(cleanUrl, "/")
    If lastSlash > 0 And lastSlash < Len(cleanUrl) Then
        ExtractVSimpleId = Mid(cleanUrl, lastSlash + 1)
    Else
        ExtractVSimpleId = ""
    End If
    Debug.Print "[PathHelper] ExtractVSimpleId: '" & url & "' -> '" & ExtractVSimpleId & "'"
End Function

' Gracefully handle #N/A, #REF! etc. - returns empty string instead of runtime error
Public Function SafeCellValue(rng As Range) As String
    ' Common when lookups fail or formulas reference deleted cells
    On Error Resume Next

    If rng Is Nothing Then
        SafeCellValue = ""
        Exit Function
    End If

    ' Check for Excel errors
    If IsError(rng.Value) Then
        Debug.Print "[PathHelper] SafeCellValue: Cell error detected, returning empty"
        SafeCellValue = ""
        Exit Function
    End If

    SafeCellValue = Trim(CStr(rng.Value))
End Function

