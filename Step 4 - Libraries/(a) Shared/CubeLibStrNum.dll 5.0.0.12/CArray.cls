VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function IsAlreadyInArrayStr(ByRef Arr() As String, ByVal strToFind As String) As Boolean
    Dim i As Long
        
    IsAlreadyInArrayStr = False
    
    For i = 1 To UBound(Arr)
        If Arr(i) = strToFind Then
            IsAlreadyInArrayStr = True
            Exit For
        End If
    
    Next
End Function
