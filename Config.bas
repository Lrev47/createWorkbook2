Attribute VB_Name = "Config"
Option Explicit

'====================================================
' All template paths defined here for easy maintenance
'====================================================

' OneDrive sync folder path suffix (after USERPROFILE)
Private Const ONEDRIVE_SUFFIX As String = "\Associated Material Handling Industries, Inc\Fleet Assets - Documents\"
Private Const TPL_FOLDER As String = "Templates\UsageWorkbook\"

' 10 template variants - 8 for New Usage (3 boolean flags = 2^3) + Return + Swap
Public Const TPL_NEWUSAGE As String = "UW_NewUsage.xlsm"
Public Const TPL_NEWUSAGE_KEHE As String = "UW_NewUsage_Kehe.xlsm"
Public Const TPL_NEWUSAGE_RETURN As String = "UW_NewUsage_HasReturn.xlsm"
Public Const TPL_NEWUSAGE_RETURN_KEHE As String = "UW_NewUsage_HasReturn_Kehe.xlsm"
Public Const TPL_NEWUSAGE_STOCK As String = "UW_NewUsage_Stock.xlsm"
Public Const TPL_NEWUSAGE_STOCK_KEHE As String = "UW_NewUsage_Stock_Kehe.xlsm"
Public Const TPL_NEWUSAGE_RETURN_STOCK As String = "UW_NewUsage_HasReturn_Stock.xlsm"
Public Const TPL_NEWUSAGE_RETURN_STOCK_KEHE As String = "UW_NewUsage_HasReturn_Stock_Kehe.xlsm"
Public Const TPL_RETURN As String = "UW_Return.xlsm"
Public Const TPL_SWAP As String = "UW_Swap.xlsm"

'====================================================
' Path construction functions
'====================================================

Public Function GetBasePath() As String
    ' Builds path using USERPROFILE env var - works on any user's machine
    GetBasePath = Environ("USERPROFILE") & ONEDRIVE_SUFFIX
    Debug.Print "[Config] GetBasePath: " & GetBasePath
End Function

Public Function GetTemplatePath(templateName As String) As String
    ' Full path to a specific template file
    GetTemplatePath = GetBasePath() & TPL_FOLDER & templateName
    Debug.Print "[Config] GetTemplatePath: " & templateName & " -> " & GetTemplatePath
End Function

Public Function GetExportRoot() As String
    ' All generated workbooks go under this folder (organized by year/customer)
    GetExportRoot = GetBasePath() & "Transactions\"
    Debug.Print "[Config] GetExportRoot: " & GetExportRoot
End Function
