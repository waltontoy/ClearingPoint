VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiNCTS_IEM_MAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_IEM_MAP_ID As Long ' 1
Private mvarFIELD_NCTS_IEM_MAP_Source As String ' 2
Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 3
Private mvarFIELD_NCTS_IEM_ID As Long ' 4
Private mvarFIELD_NCTS_IEM_MAP_StartPosition As Long ' 5
Private mvarFIELD_NCTS_IEM_MAP_Length As Long ' 6
Private mvarFIELD_EDI_TMS_SEG_ID As Long ' 7
Private mvarFIELD_NCTS_IEM_MAP_ParentID As Long ' 8
Private mvarFIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL As Long ' 9
Private mvarFIELD_NCTS_IEM_MCN_ID As Long ' 10

'  1
Public Property Let FIELD_NCTS_IEM_MAP_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MAP_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_ID() As Long
   FIELD_NCTS_IEM_MAP_ID = mvarFIELD_NCTS_IEM_MAP_ID
End Property

'  2
Public Property Let FIELD_NCTS_IEM_MAP_Source(ByVal vData As String)
   mvarFIELD_NCTS_IEM_MAP_Source = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_Source() As String
   FIELD_NCTS_IEM_MAP_Source = mvarFIELD_NCTS_IEM_MAP_Source
End Property

'  3
Public Property Let FIELD_NCTS_IEM_TMS_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ID() As Long
   FIELD_NCTS_IEM_TMS_ID = mvarFIELD_NCTS_IEM_TMS_ID
End Property

'  4
Public Property Let FIELD_NCTS_IEM_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_ID() As Long
   FIELD_NCTS_IEM_ID = mvarFIELD_NCTS_IEM_ID
End Property

'  5
Public Property Let FIELD_NCTS_IEM_MAP_StartPosition(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MAP_StartPosition = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_StartPosition() As Long
   FIELD_NCTS_IEM_MAP_StartPosition = mvarFIELD_NCTS_IEM_MAP_StartPosition
End Property

'  6
Public Property Let FIELD_NCTS_IEM_MAP_Length(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MAP_Length = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_Length() As Long
   FIELD_NCTS_IEM_MAP_Length = mvarFIELD_NCTS_IEM_MAP_Length
End Property

'  7
Public Property Let FIELD_EDI_TMS_SEG_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_SEG_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_ID() As Long
   FIELD_EDI_TMS_SEG_ID = mvarFIELD_EDI_TMS_SEG_ID
End Property

'  8
Public Property Let FIELD_NCTS_IEM_MAP_ParentID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MAP_ParentID = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_ParentID() As Long
   FIELD_NCTS_IEM_MAP_ParentID = mvarFIELD_NCTS_IEM_MAP_ParentID
End Property

'  9
Public Property Let FIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL = vData
End Property

Public Property Get FIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL() As Long
   FIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL = mvarFIELD_NCTS_IEM_MAP_EDI_ITM_ORDINAL
End Property

'  10
Public Property Let FIELD_NCTS_IEM_MCN_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MCN_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_MCN_ID() As Long
   FIELD_NCTS_IEM_MCN_ID = mvarFIELD_NCTS_IEM_MCN_ID
End Property


