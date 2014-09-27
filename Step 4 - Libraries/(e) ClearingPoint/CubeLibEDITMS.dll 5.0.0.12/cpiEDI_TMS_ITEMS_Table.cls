VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEDI_TMS_ITEMS_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_EDI_TMS_ITM_ID As Long ' 1
Private mvarFIELD_EDI_TMS_SEG_ID As Long ' 2
Private mvarFIELD_EDI_TMS_ITM_RemarksSegmentTag As String ' 3
Private mvarFIELD_EDI_TMS_ITM_Ordinal As Long ' 4
Private mvarFIELD_EDI_TMS_GRP_ID As Long ' 5
Private mvarFIELD_EDI_TMS_ITM_RemarksGroupTag As String ' 6
Private mvarFIELD_EDI_TMS_ITM_Tag As String ' 7
Private mvarFIELD_EDI_TMS_ITM_Description As String ' 8
Private mvarFIELD_EDI_TMS_ITM_DataType As String ' 9
Private mvarFIELD_EDI_TMS_ITM_Usage As String ' 10

'  1
Public Property Let FIELD_EDI_TMS_ITM_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ITM_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_ID() As Long
   FIELD_EDI_TMS_ITM_ID = mvarFIELD_EDI_TMS_ITM_ID
End Property

'  2
Public Property Let FIELD_EDI_TMS_SEG_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_SEG_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_ID() As Long
   FIELD_EDI_TMS_SEG_ID = mvarFIELD_EDI_TMS_SEG_ID
End Property

'  3
Public Property Let FIELD_EDI_TMS_ITM_RemarksSegmentTag(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_RemarksSegmentTag = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_RemarksSegmentTag() As String
   FIELD_EDI_TMS_ITM_RemarksSegmentTag = mvarFIELD_EDI_TMS_ITM_RemarksSegmentTag
End Property

'  4
Public Property Let FIELD_EDI_TMS_ITM_Ordinal(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ITM_Ordinal = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_Ordinal() As Long
   FIELD_EDI_TMS_ITM_Ordinal = mvarFIELD_EDI_TMS_ITM_Ordinal
End Property

'  5
Public Property Let FIELD_EDI_TMS_GRP_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_GRP_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_GRP_ID() As Long
   FIELD_EDI_TMS_GRP_ID = mvarFIELD_EDI_TMS_GRP_ID
End Property

'  6
Public Property Let FIELD_EDI_TMS_ITM_RemarksGroupTag(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_RemarksGroupTag = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_RemarksGroupTag() As String
   FIELD_EDI_TMS_ITM_RemarksGroupTag = mvarFIELD_EDI_TMS_ITM_RemarksGroupTag
End Property

'  7
Public Property Let FIELD_EDI_TMS_ITM_Tag(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_Tag = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_Tag() As String
   FIELD_EDI_TMS_ITM_Tag = mvarFIELD_EDI_TMS_ITM_Tag
End Property

'  8
Public Property Let FIELD_EDI_TMS_ITM_Description(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_Description = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_Description() As String
   FIELD_EDI_TMS_ITM_Description = mvarFIELD_EDI_TMS_ITM_Description
End Property

'  9
Public Property Let FIELD_EDI_TMS_ITM_DataType(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_DataType = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_DataType() As String
   FIELD_EDI_TMS_ITM_DataType = mvarFIELD_EDI_TMS_ITM_DataType
End Property

'  10
Public Property Let FIELD_EDI_TMS_ITM_Usage(ByVal vData As String)
   mvarFIELD_EDI_TMS_ITM_Usage = vData
End Property

Public Property Get FIELD_EDI_TMS_ITM_Usage() As String
   FIELD_EDI_TMS_ITM_Usage = mvarFIELD_EDI_TMS_ITM_Usage
End Property


