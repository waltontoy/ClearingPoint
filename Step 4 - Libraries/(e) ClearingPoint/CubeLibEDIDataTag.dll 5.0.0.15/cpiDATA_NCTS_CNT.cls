VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiDATA_NCTS_CNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_DATA_NCTS_CNT_ID As Long ' 1
Private mvarFIELD_DATA_NCTS_MSG_ID As Long ' 2
Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 3
Private mvarFIELD_DATA_NCTS_CNT_ParentID As Long ' 4
Private mvarFIELD_DATA_NCTS_CNT_Seq1 As String ' 5
Private mvarFIELD_DATA_NCTS_CNT_Seq2 As String ' 6
Private mvarFIELD_DATA_NCTS_CNT_Seq3 As String ' 7
Private mvarFIELD_DATA_NCTS_CNT_Instance As Long ' 8

'  1
Public Property Let FIELD_DATA_NCTS_CNT_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_CNT_ID = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_ID() As Long
   FIELD_DATA_NCTS_CNT_ID = mvarFIELD_DATA_NCTS_CNT_ID
End Property

'  2
Public Property Let FIELD_DATA_NCTS_MSG_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_MSG_ID = vData
End Property

Public Property Get FIELD_DATA_NCTS_MSG_ID() As Long
   FIELD_DATA_NCTS_MSG_ID = mvarFIELD_DATA_NCTS_MSG_ID
End Property

'  3
Public Property Let FIELD_NCTS_IEM_TMS_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ID() As Long
   FIELD_NCTS_IEM_TMS_ID = mvarFIELD_NCTS_IEM_TMS_ID
End Property

'  4
Public Property Let FIELD_DATA_NCTS_CNT_ParentID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_CNT_ParentID = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_ParentID() As Long
   FIELD_DATA_NCTS_CNT_ParentID = mvarFIELD_DATA_NCTS_CNT_ParentID
End Property

'  5
Public Property Let FIELD_DATA_NCTS_CNT_Seq1(ByVal vData As String)
   mvarFIELD_DATA_NCTS_CNT_Seq1 = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_Seq1() As String
   FIELD_DATA_NCTS_CNT_Seq1 = mvarFIELD_DATA_NCTS_CNT_Seq1
End Property

'  6
Public Property Let FIELD_DATA_NCTS_CNT_Seq2(ByVal vData As String)
   mvarFIELD_DATA_NCTS_CNT_Seq2 = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_Seq2() As String
   FIELD_DATA_NCTS_CNT_Seq2 = mvarFIELD_DATA_NCTS_CNT_Seq2
End Property

'  7
Public Property Let FIELD_DATA_NCTS_CNT_Seq3(ByVal vData As String)
   mvarFIELD_DATA_NCTS_CNT_Seq3 = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_Seq3() As String
   FIELD_DATA_NCTS_CNT_Seq3 = mvarFIELD_DATA_NCTS_CNT_Seq3
End Property

'  8
Public Property Let FIELD_DATA_NCTS_CNT_Instance(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_CNT_Instance = vData
End Property

Public Property Get FIELD_DATA_NCTS_CNT_Instance() As Long
   FIELD_DATA_NCTS_CNT_Instance = mvarFIELD_DATA_NCTS_CNT_Instance
End Property


