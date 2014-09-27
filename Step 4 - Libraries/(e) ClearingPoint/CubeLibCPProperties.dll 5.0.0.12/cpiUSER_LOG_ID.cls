VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiUSER_LOG_ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_USER_NO As Long ' 1
Private mvarFIELD_LOGID_DESCRIPTION As String ' 2
Private mvarFIELD_DTYPE As Byte ' 3
Private mvarFIELD_COMM As String ' 4
Private mvarFIELD_PRINT As String ' 5

'  1
Public Property Let FIELD_USER_NO(ByVal vData As Long)
   mvarFIELD_USER_NO = vData
End Property

Public Property Get FIELD_USER_NO() As Long
   FIELD_USER_NO = mvarFIELD_USER_NO
End Property

'  2
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property

Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

'  3
Public Property Let FIELD_DTYPE(ByVal vData As Byte)
   mvarFIELD_DTYPE = vData
End Property

Public Property Get FIELD_DTYPE() As Byte
   FIELD_DTYPE = mvarFIELD_DTYPE
End Property

'  4
Public Property Let FIELD_COMM(ByVal vData As String)
   mvarFIELD_COMM = vData
End Property

Public Property Get FIELD_COMM() As String
   FIELD_COMM = mvarFIELD_COMM
End Property

'  5
Public Property Let FIELD_PRINT(ByVal vData As String)
   mvarFIELD_PRINT = vData
End Property

Public Property Get FIELD_PRINT() As String
   FIELD_PRINT = mvarFIELD_PRINT
End Property


