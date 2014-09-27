VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDataNctsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiDataNctsMessage
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: DataNctsMessage Properties, mapped from DataNctsMessage table
' Methods: 0 (0/0)
'
' Properties: 45 (0/45)
' >P1 - Public Property Let User_ID(ByVal vData As Long)
'       Public Property Get User_ID() As Long
' >P2 - Public Property Let User_Name(ByVal vData As String)
'       Public Property Get User_Name() As String
' >P3 - Public Property Let User_Password(ByVal vData As String)
'       Public Property Get User_Password() As String
' >P4 - Public Property Let User_Rights(ByVal vData As Long)
'       Public Property Get User_Rights() As Long
' >P5 - Public Property Let User_StateSettings(ByVal vData As Long)
'       Public Property Get User_StateSettings() As Long
' >P6 - Public Property Let User_MainDimensionPosition(ByVal vData As String)
'       Public Property Get User_MainDimensionPosition() As String
' >P7 - Public Property Let User_TreeWidth(ByVal vData As Long)
'       Public Property Get User_TreeWidth() As Long
' >P8 - Public Property Let User_Level(ByVal vData As Long)
'       Public Property Get User_Level() As Long
' >P9 - Public Property Let User_ButtonGroup(ByVal vData As Long)
'       Public Property Get User_ButtonGroup() As Long
' >P10 - Public Property Let User_ButtonBarWidth(ByVal vData As Long)
'        Public Property Get User_ButtonBarWidth() As Long
'
' Events: 0 (0/0)
'
' API Methods: 0 (0/0)
'
' Constants: 0 (0/0)
'
' Enumerations: 0 (0/0)
'
' Variables: 45 (45/0)
' <V1 - Private mvarUser_ID As Long
' <V2 - Private mvarUser_Name As String
' <V3 - Private mvarUser_Password As String
' <V4 - Private mvarUser_Rights As Long
' <V5 - Private mvarUser_StateSettings As Long
' <V6 - Private mvarUser_MainDimensionPosition As String
' <V7 - Private mvarUser_TreeWidth As Long
' <V8 - Private mvarUser_Level As Long
' <V9 - Private mvarUser_ButtonGroup As Long
' <V10 - Private mvarUser_ButtonBarWidth As Long
' <V11 - Private mvarADMINISTRATOR_RIGHTS As Boolean
' <V12 - Private mvarMAINTAIN_TABLES As Boolean
'
' Objects: 0 (0/0)
' ************************************************************************************************
' ************************************************************************************************

Option Explicit

Public Key As String ' >V1
Public Index As Long ' >V2

Private mvarFIELD_DATA_NCTS_MSG_ID As Long ' <V3
Private mvarFIELD_NCTS_IEM_ID As Long ' <V4
Private mvarFIELD_DATA_NCTS_ID As Long ' <V5
Private mvarFIELD_DATA_NCTS_MSG_Date As Date ' <V6
Private mvarFIELD_DATA_NCTS_MSG_StatusType As String ' <V7
Private mvarFIELD_User_ID As Long ' <V8
Private mvarFIELD_DATA_NCTS_MSG_Message As String ' <V9
Private mvarFIELD_DATA_NCTS_MSG_Reference As String ' <V10

' Code: >P1
' Name: FIELD_DATA_NCTS_MSG_ID (Let/Get)
' Description: set/return FIELD_DATA_NCTS_MSG_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_MSG_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_MSG_ID = vData
End Property
Public Property Get FIELD_DATA_NCTS_MSG_ID() As Long
   FIELD_DATA_NCTS_MSG_ID = mvarFIELD_DATA_NCTS_MSG_ID
End Property

' Code: >P2
' Name: FIELD_NCTS_IEM_ID (Let/Get)
' Description: set/return FIELD_NCTS_IEM_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_NCTS_IEM_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_ID = vData
End Property
Public Property Get FIELD_NCTS_IEM_ID() As Long
   FIELD_NCTS_IEM_ID = mvarFIELD_NCTS_IEM_ID
End Property

' Code: >P3
' Name: FIELD_DATA_NCTS_ID (Let/Get)
' Description: set/return FIELD_DATA_NCTS_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_ID = vData
End Property
Public Property Get FIELD_DATA_NCTS_ID() As Long
   FIELD_DATA_NCTS_ID = mvarFIELD_DATA_NCTS_ID
End Property

' Code: >P4
' Name: FIELD_DATA_NCTS_MSG_Date (Let/Get)
' Description: set/return FIELD_DATA_NCTS_MSG_Date
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_MSG_Date(ByVal vData As Date)
   mvarFIELD_DATA_NCTS_MSG_Date = vData
End Property
Public Property Get FIELD_DATA_NCTS_MSG_Date() As Date
   FIELD_DATA_NCTS_MSG_Date = mvarFIELD_DATA_NCTS_MSG_Date
End Property

' Code: >P5
' Name: FIELD_DATA_NCTS_MSG_StatusType (Let/Get)
' Description: set/return FIELD_DATA_NCTS_MSG_StatusType
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_MSG_StatusType(ByVal vData As String)
   mvarFIELD_DATA_NCTS_MSG_StatusType = vData
End Property
Public Property Get FIELD_DATA_NCTS_MSG_StatusType() As String
   FIELD_DATA_NCTS_MSG_StatusType = mvarFIELD_DATA_NCTS_MSG_StatusType
End Property

' Code: >P6
' Name: FIELD_User_ID (Let/Get)
' Description: set/return FIELD_User_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_User_ID(ByVal vData As Long)
   mvarFIELD_User_ID = vData
End Property
Public Property Get FIELD_User_ID() As Long
   FIELD_User_ID = mvarFIELD_User_ID
End Property

' Code: >P7
' Name: FIELD_DATA_NCTS_MSG_Message (Let/Get)
' Description: set/return FIELD_DATA_NCTS_MSG_Message
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_MSG_Message(ByVal vData As String)
   mvarFIELD_DATA_NCTS_MSG_Message = vData
End Property
Public Property Get FIELD_DATA_NCTS_MSG_Message() As String
   FIELD_DATA_NCTS_MSG_Message = mvarFIELD_DATA_NCTS_MSG_Message
End Property

' Code: >P8
' Name: FIELD_DATA_NCTS_MSG_Reference (Let/Get)
' Description: set/return FIELD_DATA_NCTS_MSG_Reference
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATA_NCTS_MSG_Reference(ByVal vData As String)
   mvarFIELD_DATA_NCTS_MSG_Reference = vData
End Property
Public Property Get FIELD_DATA_NCTS_MSG_Reference() As String
   FIELD_DATA_NCTS_MSG_Reference = mvarFIELD_DATA_NCTS_MSG_Reference
End Property
