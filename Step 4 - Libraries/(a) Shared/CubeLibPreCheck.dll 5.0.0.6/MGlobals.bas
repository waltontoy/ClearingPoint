Attribute VB_Name = "MGlobals"
Option Explicit

Public Enum CheckResult
    cpiInitialized = 0
    cpiSuccess = 1
    cpiRetry = 2
    cpiCancel = 3
    cpiSetPath = 4
End Enum

Public Const KEY_ENCRYPT = ""

Public Declare Function PathFileExists _
    Lib "shlwapi.dll" Alias "PathFileExistsA" ( _
    ByVal pszPath As String) As Long

