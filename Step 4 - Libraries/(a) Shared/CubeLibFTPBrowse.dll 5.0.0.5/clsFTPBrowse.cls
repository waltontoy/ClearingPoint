VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTPBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Private rstOffline As ADODB.Recordset

Public Sub Browse(strUsername As String, strPassword As String, _
                  strFTPAddress As String, strClient As String, _
                  strSubClient As String, strFolder() As String)
                  
    frmBrowse.MyLoad strUsername, strPassword, _
                  strFTPAddress, strClient, _
                  strSubClient, strFolder()
End Sub




