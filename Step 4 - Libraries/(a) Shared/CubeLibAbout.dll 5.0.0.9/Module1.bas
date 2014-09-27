Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Sub pp_errorstr Lib "Keylib32.DLL" (ByVal number As Long, ByVal buffer As String)

Public Const G_CONST_APPLICATION_NAME = "ClearingPoint"
    Public Const G_CONST_COPYRIGHT_COMPANY = "Cubepoint, Inc."
    Public Const G_CONST_COPYRIGHT_YEAR_START = "2001"
    Public Const G_CONST_COPYRIGHT_YEAR_END = "2006"
    Public Const G_TECH_SUPPORT_URL = "http://www.cubepoint.be/"
