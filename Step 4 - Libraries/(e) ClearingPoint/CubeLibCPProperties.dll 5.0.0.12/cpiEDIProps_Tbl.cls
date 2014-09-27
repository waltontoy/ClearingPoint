VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEDIProps_Tbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_EDIPROP_ID As Long ' 1
Private mvarFIELD_EDIPROP_Type As Long ' 2
Private mvarFIELD_EDIPROP_QueueName As String ' 3
Private mvarFIELD_EDIPROP_Host As String ' 4
Private mvarFIELD_EDIPROP_Port As String ' 5
Private mvarFIELD_EDIPROP_UserName As String ' 6
Private mvarFIELD_EDIPROP_UserPassword As String ' 7
Private mvarFIELD_EDIPROP_SlowPollCycles As String ' 8
Private mvarFIELD_EDIPROP_FastPollCycles As String ' 9
Private mvarFIELD_EDIPROP_TimeOut As Long ' 10

'  1
Public Property Let FIELD_EDIPROP_ID(ByVal vData As Long)
   mvarFIELD_EDIPROP_ID = vData
End Property

Public Property Get FIELD_EDIPROP_ID() As Long
   FIELD_EDIPROP_ID = mvarFIELD_EDIPROP_ID
End Property

'  2
Public Property Let FIELD_EDIPROP_Type(ByVal vData As Long)
   mvarFIELD_EDIPROP_Type = vData
End Property

Public Property Get FIELD_EDIPROP_Type() As Long
   FIELD_EDIPROP_Type = mvarFIELD_EDIPROP_Type
End Property

'  3
Public Property Let FIELD_EDIPROP_QueueName(ByVal vData As String)
   mvarFIELD_EDIPROP_QueueName = vData
End Property

Public Property Get FIELD_EDIPROP_QueueName() As String
   FIELD_EDIPROP_QueueName = mvarFIELD_EDIPROP_QueueName
End Property

'  4
Public Property Let FIELD_EDIPROP_Host(ByVal vData As String)
   mvarFIELD_EDIPROP_Host = vData
End Property

Public Property Get FIELD_EDIPROP_Host() As String
   FIELD_EDIPROP_Host = mvarFIELD_EDIPROP_Host
End Property

'  5
Public Property Let FIELD_EDIPROP_Port(ByVal vData As String)
   mvarFIELD_EDIPROP_Port = vData
End Property

Public Property Get FIELD_EDIPROP_Port() As String
   FIELD_EDIPROP_Port = mvarFIELD_EDIPROP_Port
End Property

'  6
Public Property Let FIELD_EDIPROP_UserName(ByVal vData As String)
   mvarFIELD_EDIPROP_UserName = vData
End Property

Public Property Get FIELD_EDIPROP_UserName() As String
   FIELD_EDIPROP_UserName = mvarFIELD_EDIPROP_UserName
End Property

'  7
Public Property Let FIELD_EDIPROP_UserPassword(ByVal vData As String)
   mvarFIELD_EDIPROP_UserPassword = vData
End Property

Public Property Get FIELD_EDIPROP_UserPassword() As String
   FIELD_EDIPROP_UserPassword = mvarFIELD_EDIPROP_UserPassword
End Property

'  8
Public Property Let FIELD_EDIPROP_SlowPollCycles(ByVal vData As String)
   mvarFIELD_EDIPROP_SlowPollCycles = vData
End Property

Public Property Get FIELD_EDIPROP_SlowPollCycles() As String
   FIELD_EDIPROP_SlowPollCycles = mvarFIELD_EDIPROP_SlowPollCycles
End Property

'  9
Public Property Let FIELD_EDIPROP_FastPollCycles(ByVal vData As String)
   mvarFIELD_EDIPROP_FastPollCycles = vData
End Property

Public Property Get FIELD_EDIPROP_FastPollCycles() As String
   FIELD_EDIPROP_FastPollCycles = mvarFIELD_EDIPROP_FastPollCycles
End Property

'  10
Public Property Let FIELD_EDIPROP_TimeOut(ByVal vData As Long)
   mvarFIELD_EDIPROP_TimeOut = vData
End Property

Public Property Get FIELD_EDIPROP_TimeOut() As Long
   FIELD_EDIPROP_TimeOut = mvarFIELD_EDIPROP_TimeOut
End Property


