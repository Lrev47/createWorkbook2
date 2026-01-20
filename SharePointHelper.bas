Attribute VB_Name = "SharePointHelper"
Option Explicit

Private Const HKEY_CURRENT_USER = &H80000001

' Reads OneDrive sync metadata from registry to build SharePoint URL
' This lets us generate shareable links for files synced via OneDrive

Public Function GetSharePointUrl(localPath As String) As String
    Dim oReg As Object
    Dim arrSubKeys() As Variant
    Dim subKey As Variant
    Dim basePath As String
    Dim providerPath As String
    Dim mountPoint As String
    Dim urlNamespace As String
    Dim relativePath As String
    Dim sharePointUrl As String

    On Error GoTo ErrorHandler

    Debug.Print "[SharePointHelper] Looking for match: " & localPath

    ' WMI provides registry access - enumerate OneDrive sync providers
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    basePath = "SOFTWARE\SyncEngines\Providers\OneDrive"

    ' Each synced SharePoint library has its own provider entry
    oReg.EnumKey HKEY_CURRENT_USER, basePath, arrSubKeys

    If IsEmpty(arrSubKeys) Then
        Debug.Print "[SharePointHelper] No OneDrive providers found in registry"
        GetSharePointUrl = ""
        Exit Function
    End If

    ' Find the provider whose MountPoint matches our local path
    For Each subKey In arrSubKeys
        providerPath = basePath & "\" & subKey

        ' Read MountPoint value
        oReg.GetStringValue HKEY_CURRENT_USER, providerPath, "MountPoint", mountPoint

        If Len(mountPoint) > 0 Then
            Debug.Print "[SharePointHelper] Checking " & subKey & " MountPoint: " & mountPoint

            ' Check if our file lives under this provider's sync folder
            If InStr(1, localPath, mountPoint, vbTextCompare) = 1 Then
                ' Found the right provider - now build the SharePoint URL
                oReg.GetStringValue HKEY_CURRENT_USER, providerPath, "UrlNamespace", urlNamespace
                Debug.Print "[SharePointHelper] MATCH! URLNamespace: " & urlNamespace

                ' Get the path relative to the sync root
                relativePath = Mid(localPath, Len(mountPoint) + 2)

                ' Construct URL: base namespace + relative path
                sharePointUrl = urlNamespace & relativePath

                ' Fix slashes and encode spaces for URL
                sharePointUrl = Replace(sharePointUrl, "\", "/")
                sharePointUrl = Replace(sharePointUrl, " ", "%20")

                ' ?web=1 forces browser view instead of download prompt
                GetSharePointUrl = sharePointUrl & "?web=1"
                Debug.Print "[SharePointHelper] SharePoint URL: " & sharePointUrl & "?web=1"
                Exit Function
            End If
        End If
    Next subKey

    Debug.Print "[SharePointHelper] No matching OneDrive provider found for path"
    GetSharePointUrl = ""
    Exit Function

ErrorHandler:
    Debug.Print "[SharePointHelper] Error: " & Err.Description
    GetSharePointUrl = ""
End Function
