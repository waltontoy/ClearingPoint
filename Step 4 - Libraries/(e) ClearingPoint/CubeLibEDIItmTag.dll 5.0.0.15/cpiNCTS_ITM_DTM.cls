VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiNCTS_ITM_DTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_ITM_DTM_ID As Long ' 1
Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 2
Private mvarFIELD_EDI_TMS_ITM_ID As Long ' 3
Private mvarFIELD_NCTS_ITM_DTM_Description As String ' 4
Private mvarFIELD_NCTS_ITM_DTM_Value As String ' 5
Private mvarFIELD_NCTS_ITM_DTM_Codelist As String ' 6
Private mvarFIELD_NCTS_ITM_DTM_DataType As String ' 7
Private mvarFIELD_NCTS_ITM_DTM_Usage As String ' 8

'  1
Public Property Let FIELD_NCTS_ITM_DTM_ID(ByVal vData As Long)
   mvarFIELD_NCTS_ITM_DTM_ID = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_ID() As Long
   FIELD_NCTS_ITM_DTM_ID = mvarFIELD_NCTS_ITM_DTM_ID
End Property

'  2
Public Property Let FIELD_NCTS_IEM_TMS_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ID() As Long
   FIELD_NCTS_IEM_TMS_ID = mvarFIELD_NCTS_IEM_TMS_ID
End Property

'  3
Public Property Let FIELD_EDI_TMS_ITM_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ITM_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_ID() As Long
   FIELD_EDI_TMS_ITM_ID = mvarFIELD_EDI_TMS_ITM_ID
End Property

'  4
Public Property Let FIELD_NCTS_ITM_DTM_Description(ByVal vData As String)
   mvarFIELD_NCTS_ITM_DTM_Description = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_Description() As String
   FIELD_NCTS_ITM_DTM_Description = mvarFIELD_NCTS_ITM_DTM_Description
End Property

'  5
Public Property Let FIELD_NCTS_ITM_DTM_Value(ByVal vData As String)
   mvarFIELD_NCTS_ITM_DTM_Value = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_Value() As String
   FIELD_NCTS_ITM_DTM_Value = mvarFIELD_NCTS_ITM_DTM_Value
End Property

'  6
Public Property Let FIELD_NCTS_ITM_DTM_Codelist(ByVal vData As String)
   mvarFIELD_NCTS_ITM_DTM_Codelist = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_Codelist() As String
   FIELD_NCTS_ITM_DTM_Codelist = mvarFIELD_NCTS_ITM_DTM_Codelist
End Property

'  7
Public Property Let FIELD_NCTS_ITM_DTM_DataType(ByVal vData As String)
   mvarFIELD_NCTS_ITM_DTM_DataType = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_DataType() As String
   FIELD_NCTS_ITM_DTM_DataType = mvarFIELD_NCTS_ITM_DTM_DataType
End Property

'  8
Public Property Let FIELD_NCTS_ITM_DTM_Usage(ByVal vData As String)
   mvarFIELD_NCTS_ITM_DTM_Usage = vData
End Property

Public Property Get FIELD_NCTS_ITM_DTM_Usage() As String
   FIELD_NCTS_ITM_DTM_Usage = mvarFIELD_NCTS_ITM_DTM_Usage
End Property


