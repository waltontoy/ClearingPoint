VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEDI_TMS_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_EDI_TMS_ID As Long ' 1
Private mvarFIELD_EDI_TMS_COR_ID As Long ' 2
Private mvarFIELD_EDI_TMS_SEG_ID As Long ' 3
Private mvarFIELD_EDI_TMS_RemarksSegmentTag As String ' 4
Private mvarFIELD_EDI_TMS_ParentID As Long ' 5
Private mvarFIELD_EDI_TMS_Sequence As Long ' 6
Private mvarFIELD_EDI_TMS_Level As Long ' 7
Private mvarFIELD_EDI_TMS_Occurrence As Long ' 8
Private mvarFIELD_EDI_TMS_Usage As String ' 9

'  1
Public Property Let FIELD_EDI_TMS_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_ID() As Long
   FIELD_EDI_TMS_ID = mvarFIELD_EDI_TMS_ID
End Property

'  2
Public Property Let FIELD_EDI_TMS_COR_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_COR_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_COR_ID() As Long
   FIELD_EDI_TMS_COR_ID = mvarFIELD_EDI_TMS_COR_ID
End Property

'  3
Public Property Let FIELD_EDI_TMS_SEG_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_SEG_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_ID() As Long
   FIELD_EDI_TMS_SEG_ID = mvarFIELD_EDI_TMS_SEG_ID
End Property

'  4
Public Property Let FIELD_EDI_TMS_RemarksSegmentTag(ByVal vData As String)
   mvarFIELD_EDI_TMS_RemarksSegmentTag = vData
End Property

Public Property Get FIELD_EDI_TMS_RemarksSegmentTag() As String
   FIELD_EDI_TMS_RemarksSegmentTag = mvarFIELD_EDI_TMS_RemarksSegmentTag
End Property

'  5
Public Property Let FIELD_EDI_TMS_ParentID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ParentID = vData
End Property

Public Property Get FIELD_EDI_TMS_ParentID() As Long
   FIELD_EDI_TMS_ParentID = mvarFIELD_EDI_TMS_ParentID
End Property

'  6
Public Property Let FIELD_EDI_TMS_Sequence(ByVal vData As Long)
   mvarFIELD_EDI_TMS_Sequence = vData
End Property

Public Property Get FIELD_EDI_TMS_Sequence() As Long
   FIELD_EDI_TMS_Sequence = mvarFIELD_EDI_TMS_Sequence
End Property

'  7
Public Property Let FIELD_EDI_TMS_Level(ByVal vData As Long)
   mvarFIELD_EDI_TMS_Level = vData
End Property

Public Property Get FIELD_EDI_TMS_Level() As Long
   FIELD_EDI_TMS_Level = mvarFIELD_EDI_TMS_Level
End Property

'  8
Public Property Let FIELD_EDI_TMS_Occurrence(ByVal vData As Long)
   mvarFIELD_EDI_TMS_Occurrence = vData
End Property

Public Property Get FIELD_EDI_TMS_Occurrence() As Long
   FIELD_EDI_TMS_Occurrence = mvarFIELD_EDI_TMS_Occurrence
End Property

'  9
Public Property Let FIELD_EDI_TMS_Usage(ByVal vData As String)
   mvarFIELD_EDI_TMS_Usage = vData
End Property

Public Property Get FIELD_EDI_TMS_Usage() As String
   FIELD_EDI_TMS_Usage = mvarFIELD_EDI_TMS_Usage
End Property


