VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CThread"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Function GetThreadEXEName(Optional ByVal StripPath As Boolean = False, _
                                 Optional ByVal StripFileExt As Boolean = False) As String
    If InIDE() = False Then
        Const MAX_PATH As Long = 260
        Dim Buffer As String
        Buffer = String(MAX_PATH, vbNullChar)
        Buffer = Left$(Buffer, GetModuleFileName(0, StrPtr(Buffer), MAX_PATH + 1))
        
        If StripPath Then
            Buffer = Right$(Buffer, Len(Buffer) - InStrRev(Buffer, "\"))
        End If
        
        If StripFileExt Then
            GetThreadEXEName = Left$(Buffer, InStrRev(Buffer, ".") - 1)
        Else
            GetThreadEXEName = Buffer
        End If
    Else
        GetThreadEXEName = App.EXEName
    End If
End Function

Public Function GetThreadQualifiedEXEPath() As String
    If InIDE() = False Then
        Const MAX_PATH As Long = 260
        Dim Buffer As String
        Buffer = String(MAX_PATH, vbNullChar)
        Buffer = Left$(Buffer, GetModuleFileName(0, StrPtr(Buffer), MAX_PATH + 1))
        
        GetThreadQualifiedEXEPath = Left$(Buffer, InStrRev(Buffer, "\"))
    Else
        GetThreadQualifiedEXEPath = App.Path
    End If
End Function


Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
    If B = True Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function
