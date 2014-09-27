VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnumerations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum NCTS_IEM_IDs
    NCTS_IEM_ID_IE04 = 1
    NCTS_IEM_ID_IE07 = 2
    NCTS_IEM_ID_IE13 = 3
    NCTS_IEM_ID_IE14 = 4
    NCTS_IEM_ID_IE15 = 5
    NCTS_IEM_ID_IE19 = 6
    NCTS_IEM_ID_IE21 = 7
    NCTS_IEM_ID_IE23 = 8
    NCTS_IEM_ID_IE29 = 9
    NCTS_IEM_ID_IE43 = 10
    NCTS_IEM_ID_IE44 = 11
    NCTS_IEM_ID_IE45 = 12
    NCTS_IEM_ID_IE51 = 13
    NCTS_IEM_ID_IE54 = 14
    NCTS_IEM_ID_IE100 = 15
    NCTS_IEM_ID_IE09 = 16
    NCTS_IEM_ID_IE25 = 17
    NCTS_IEM_ID_IE28 = 18
    NCTS_IEM_ID_IE60 = 19
    NCTS_IEM_ID_IE05 = 20
    NCTS_IEM_ID_IE08 = 21
    NCTS_IEM_ID_IE16 = 22
    NCTS_IEM_ID_IE58 = 23
    NCTS_IEM_ID_IE62 = 24
    NCTS_IEM_ID_IE906 = 25
    NCTS_IEM_ID_IE907 = 26
    NCTS_IEM_ID_IE904 = 27
    NCTS_IEM_ID_IE905 = 28
    NCTS_IEM_ID_CODEM = 29
    NCTS_IEM_ID_IE917 = 30
    NCTS_IEM_ID_IE34 = 32 'RACHELLE FOR IE34 091205
End Enum

Public Enum OutputTypes
    OutType_File = 1
    OutType_Database
End Enum

Public Enum MessageSources
    MsgSource_Empty = 1
    MsgSource_File
    MsgSource_Database
End Enum

Public Enum ENCTSMessageType
    EMsg_IE04 = 1
    EMsg_IE05 = 2
    EMsg_IE07 = 3
    EMsg_IE08 = 4
    EMsg_IE09 = 5
    EMsg_IE13 = 6
    EMsg_IE14 = 7
    EMsg_IE15 = 8
    EMsg_IE16 = 9
    EMsg_IE19 = 10
    EMsg_IE21 = 11
    EMsg_IE23 = 12
    EMsg_IE25 = 13
    EMsg_IE28 = 14
    EMsg_IE29 = 15
    EMsg_IE43 = 16
    EMsg_IE44 = 17
    EMsg_IE45 = 18
    EMsg_IE51 = 19
    EMsg_IE54 = 20
    EMsg_IE58 = 21
    EMsg_IE60 = 22
    EMsg_IE62 = 23
    EMsg_IE100 = 24
    EMsg_IE904 = 25
    EMsg_IE905 = 26
    EMsg_IE906 = 27
    EMsg_IE907 = 28
    EMsg_CODEM = 29
    EMsg_IE917 = 30
    EMsg_IE55 = 31  'IAN 06-04-2005 for IE55 support
    EMsg_IE34 = 32 'RACHELLE FOR IE34 091205
End Enum

Public Enum ESegmentType
    ESegment_Type_BGM = 1
    ESegment_Type_CNT = 2
    ESegment_Type_CST = 3
    ESegment_Type_DOC = 4
    ESegment_Type_DTM = 5
    ESegment_Type_FTX = 6
    ESegment_Type_GIR = 7
    ESegment_Type_GIS = 8
    ESegment_Type_LOC = 9
    ESegment_Type_MEA = 10
    ESegment_Type_NAD = 11
    ESegment_Type_PAC = 12
    ESegment_Type_PCI = 13
    ESegment_Type_RFF = 14
    ESegment_Type_SEL = 15
    ESegment_Type_TDT = 16
    ESegment_Type_TOD = 17
    ESegment_Type_TPL = 18
    ESegment_Type_UNB = 19
    ESegment_Type_UNH = 20
    ESegment_Type_UNS = 21
    ESegment_Type_UNT = 22
    ESegment_Type_UNZ = 23
End Enum

Public Enum EMessageStatusType
    EMsgStatusType_Document = 1
    EMsgStatusType_Sent = 2
    EMsgStatusType_Received = 3
End Enum

Public Enum ECommitType
    ECommitSave = 1
    ECommitSend = 2
    ECommitReceive = 3
End Enum

Public Enum IEMessageSources
    IEMessageSource_Constant = 1
    IEMessageSource_File
End Enum

Public Enum MessageStatusTypes
    MsgStatus_Document = 1
    MsgStatus_QueuedOnly
    MsgStatus_Queued
    MsgStatus_Received
    MsgStatus_Sent
End Enum

Public Enum InstanceIndexes
    InstanceIndex_Level0 = 0
    InstanceIndex_Level1 = 1
    InstanceIndex_Level2 = 2
    InstanceIndex_Level3 = 3
    InstanceIndex_Level4 = 4
End Enum

Public Enum SegmentKeyIndexes
    SegmentKeyIndex_Prefix = 0
    SegmentKeyIndex_NCTS_IEM_TMS_ID
    SegmentKeyIndex_SegmentInstance
End Enum
