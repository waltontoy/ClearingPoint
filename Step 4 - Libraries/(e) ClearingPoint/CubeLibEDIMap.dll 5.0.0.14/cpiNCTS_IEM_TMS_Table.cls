VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiNCTS_IEM_TMS_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 1
Private mvarFIELD_NCTS_IEM_ID As Long ' 2
Private mvarFIELD_NCTS_IEM_TMS_RemarksIEMName As String ' 3
Private mvarFIELD_EDI_TMS_ID As Long ' 4
Private mvarFIELD_NCTS_IEM_TMS_RemarksSegmentTag As String ' 5
Private mvarFIELD_NCTS_IEM_TMS_ParentID As Long ' 6
Private mvarFIELD_NCTS_IEM_TMS_RemarksTMSSequence As Long ' 7
Private mvarFIELD_NCTS_IEM_TMS_Ordinal As Long ' 8
Private mvarFIELD_NCTS_IEM_TMS_Occurrence As Long ' 9
Private mvarFIELD_NCTS_IEM_TMS_Usage As String ' 10
Private mvarFIELD_NCTS_IEM_TMS_RemarksQualifier As String ' 11

'  1
Public Property Let FIELD_NCTS_IEM_TMS_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ID() As Long
   FIELD_NCTS_IEM_TMS_ID = mvarFIELD_NCTS_IEM_TMS_ID
End Property

'  2
Public Property Let FIELD_NCTS_IEM_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_ID() As Long
   FIELD_NCTS_IEM_ID = mvarFIELD_NCTS_IEM_ID
End Property

'  3
Public Property Let FIELD_NCTS_IEM_TMS_RemarksIEMName(ByVal vData As String)
   mvarFIELD_NCTS_IEM_TMS_RemarksIEMName = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_RemarksIEMName() As String
   FIELD_NCTS_IEM_TMS_RemarksIEMName = mvarFIELD_NCTS_IEM_TMS_RemarksIEMName
End Property

'  4
Public Property Let FIELD_EDI_TMS_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_ID() As Long
   FIELD_EDI_TMS_ID = mvarFIELD_EDI_TMS_ID
End Property

'  5
Public Property Let FIELD_NCTS_IEM_TMS_RemarksSegmentTag(ByVal vData As String)
   mvarFIELD_NCTS_IEM_TMS_RemarksSegmentTag = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_RemarksSegmentTag() As String
   FIELD_NCTS_IEM_TMS_RemarksSegmentTag = mvarFIELD_NCTS_IEM_TMS_RemarksSegmentTag
End Property

'  6
Public Property Let FIELD_NCTS_IEM_TMS_ParentID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ParentID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ParentID() As Long
   FIELD_NCTS_IEM_TMS_ParentID = mvarFIELD_NCTS_IEM_TMS_ParentID
End Property

'  7
Public Property Let FIELD_NCTS_IEM_TMS_RemarksTMSSequence(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_RemarksTMSSequence = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_RemarksTMSSequence() As Long
   FIELD_NCTS_IEM_TMS_RemarksTMSSequence = mvarFIELD_NCTS_IEM_TMS_RemarksTMSSequence
End Property

'  8
Public Property Let FIELD_NCTS_IEM_TMS_Ordinal(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_Ordinal = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_Ordinal() As Long
   FIELD_NCTS_IEM_TMS_Ordinal = mvarFIELD_NCTS_IEM_TMS_Ordinal
End Property

'  9
Public Property Let FIELD_NCTS_IEM_TMS_Occurrence(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_Occurrence = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_Occurrence() As Long
   FIELD_NCTS_IEM_TMS_Occurrence = mvarFIELD_NCTS_IEM_TMS_Occurrence
End Property

'  10
Public Property Let FIELD_NCTS_IEM_TMS_Usage(ByVal vData As String)
   mvarFIELD_NCTS_IEM_TMS_Usage = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_Usage() As String
   FIELD_NCTS_IEM_TMS_Usage = mvarFIELD_NCTS_IEM_TMS_Usage
End Property

'  11
Public Property Let FIELD_NCTS_IEM_TMS_RemarksQualifier(ByVal vData As String)
   mvarFIELD_NCTS_IEM_TMS_RemarksQualifier = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_RemarksQualifier() As String
   FIELD_NCTS_IEM_TMS_RemarksQualifier = mvarFIELD_NCTS_IEM_TMS_RemarksQualifier
End Property


