VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiNCTS_ITM_UNH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_ITM_UNH_ID As Long ' 1
Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 2
Private mvarFIELD_EDI_TMS_ITM_ID As Long ' 3
Private mvarFIELD_NCTS_ITM_UNH_Description As String ' 4
Private mvarFIELD_NCTS_ITM_UNH_Value As String ' 5
Private mvarFIELD_NCTS_ITM_UNH_Codelist As String ' 6
Private mvarFIELD_NCTS_ITM_UNH_DataType As String ' 7
Private mvarFIELD_NCTS_ITM_UNH_Usage As String ' 8

'  1
Public Property Let FIELD_NCTS_ITM_UNH_ID(ByVal vData As Long)
   mvarFIELD_NCTS_ITM_UNH_ID = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_ID() As Long
   FIELD_NCTS_ITM_UNH_ID = mvarFIELD_NCTS_ITM_UNH_ID
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
Public Property Let FIELD_NCTS_ITM_UNH_Description(ByVal vData As String)
   mvarFIELD_NCTS_ITM_UNH_Description = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_Description() As String
   FIELD_NCTS_ITM_UNH_Description = mvarFIELD_NCTS_ITM_UNH_Description
End Property

'  5
Public Property Let FIELD_NCTS_ITM_UNH_Value(ByVal vData As String)
   mvarFIELD_NCTS_ITM_UNH_Value = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_Value() As String
   FIELD_NCTS_ITM_UNH_Value = mvarFIELD_NCTS_ITM_UNH_Value
End Property

'  6
Public Property Let FIELD_NCTS_ITM_UNH_Codelist(ByVal vData As String)
   mvarFIELD_NCTS_ITM_UNH_Codelist = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_Codelist() As String
   FIELD_NCTS_ITM_UNH_Codelist = mvarFIELD_NCTS_ITM_UNH_Codelist
End Property

'  7
Public Property Let FIELD_NCTS_ITM_UNH_DataType(ByVal vData As String)
   mvarFIELD_NCTS_ITM_UNH_DataType = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_DataType() As String
   FIELD_NCTS_ITM_UNH_DataType = mvarFIELD_NCTS_ITM_UNH_DataType
End Property

'  8
Public Property Let FIELD_NCTS_ITM_UNH_Usage(ByVal vData As String)
   mvarFIELD_NCTS_ITM_UNH_Usage = vData
End Property

Public Property Get FIELD_NCTS_ITM_UNH_Usage() As String
   FIELD_NCTS_ITM_UNH_Usage = mvarFIELD_NCTS_ITM_UNH_Usage
End Property

