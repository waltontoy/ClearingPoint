VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDataNctsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiDataNctsTable
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: DataNctsTable Properties, mapped from DataNctsTable table
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

Private mvarFIELD_DATA_NCTS_ID As Long ' <V3
Private mvarFIELD_DATA_NCTS_MSG_ID As Long ' <V4
Private mvarFIELD_CODE As String ' <V5
Private mvarFIELD_LOGID_DESCRIPTION As String ' <V6
Private mvarFIELD_TYPE As String ' <V7
Private mvarFIELD_COMM As String ' <V8
Private mvarFIELD_USER_NO As Long ' <V9
Private mvarFIELD_LAST_MODIFIED_BY As String ' <V10
Private mvarFIELD_PRINT As String ' <V11
Private mvarFIELD_DOCUMENT_NAME As String ' <V12
Private mvarFIELD_DATE_CREATED As Date ' <V13
Private mvarFIELD_DATE_LAST_MODIFIED As Date ' <V14
Private mvarFIELD_DATE_REQUESTED As Date ' <V15
Private mvarFIELD_DATE_SEND As Date ' <V16
Private mvarFIELD_HEADER As Integer ' <V17
Private mvarFIELD_DETAIL As Integer ' <V18
Private mvarFIELD_DOCUMENT_COUNTER As Integer ' <V19
Private mvarFIELD_REMARKS As String ' <V20
Private mvarFIELD_TREE_ID As String ' <V21
Private mvarFIELD_SUPPLIER_NAME As String ' <V22
Private mvarFIELD_DTYPE As Byte ' <V23
Private mvarFIELD_Error_String As String ' <V24
Private mvarFIELD_Error_HD As String ' <V25
Private mvarFIELD_LOGID As String ' <V26
Private mvarFIELD_USERNAME As String ' <V27
Private mvarFIELD_DATE_PRINTED As Date ' <V28
Private mvarFIELD_PRINTED_BY As String ' <V29
Private mvarFIELD_MRN As String ' <V30
Private mvarFIELD_REMOTE_ID As Long ' <V30

' Code: >P1
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

' Code: >P2
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

' Code: >P3
' Name: FIELD_CODE (Let/Get)
' Description: set/return FIELD_CODE
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property
Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

' Code: >P4
' Name: FIELD_LOGID_DESCRIPTION (Let/Get)
' Description: set/return FIELD_LOGID_DESCRIPTION
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property
Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

' Code: >P5
' Name: FIELD_TYPE (Let/Get)
' Description: set/return FIELD_TYPE
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_TYPE(ByVal vData As String)
   mvarFIELD_TYPE = vData
End Property
Public Property Get FIELD_TYPE() As String
   FIELD_TYPE = mvarFIELD_TYPE
End Property

' Code: >P6
' Name: FIELD_COMM (Let/Get)
' Description: set/return FIELD_COMM
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_COMM(ByVal vData As String)
   mvarFIELD_COMM = vData
End Property
Public Property Get FIELD_COMM() As String
   FIELD_COMM = mvarFIELD_COMM
End Property

' Code: >P7
' Name: FIELD_USER_NO (Let/Get)
' Description: set/return FIELD_USER_NO
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_USER_NO(ByVal vData As Long)
   mvarFIELD_USER_NO = vData
End Property
Public Property Get FIELD_USER_NO() As Long
   FIELD_USER_NO = mvarFIELD_USER_NO
End Property

' Code: >P8
' Name: FIELD_LAST_MODIFIED_BY (Let/Get)
' Description: set/return FIELD_LAST_MODIFIED_BY
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LAST_MODIFIED_BY(ByVal vData As String)
   mvarFIELD_LAST_MODIFIED_BY = vData
End Property
Public Property Get FIELD_LAST_MODIFIED_BY() As String
   FIELD_LAST_MODIFIED_BY = mvarFIELD_LAST_MODIFIED_BY
End Property

' Code: >P9
' Name: FIELD_PRINT (Let/Get)
' Description: set/return FIELD_PRINT
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_PRINT(ByVal vData As String)
   mvarFIELD_PRINT = vData
End Property
Public Property Get FIELD_PRINT() As String
   FIELD_PRINT = mvarFIELD_PRINT
End Property

' Code: >P10
' Name: FIELD_DOCUMENT_NAME (Let/Get)
' Description: set/return FIELD_DOCUMENT_NAME
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DOCUMENT_NAME(ByVal vData As String)
   mvarFIELD_DOCUMENT_NAME = vData
End Property
Public Property Get FIELD_DOCUMENT_NAME() As String
   FIELD_DOCUMENT_NAME = mvarFIELD_DOCUMENT_NAME
End Property

' Code: >P11
' Name: FIELD_DATE_CREATED (Let/Get)
' Description: set/return FIELD_DATE_CREATED
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_CREATED(ByVal vData As Date)
   mvarFIELD_DATE_CREATED = vData
End Property
Public Property Get FIELD_DATE_CREATED() As Date
   FIELD_DATE_CREATED = mvarFIELD_DATE_CREATED
End Property

' Code: >P12
' Name: FIELD_DATE_LAST_MODIFIED (Let/Get)
' Description: set/return FIELD_DATE_LAST_MODIFIED
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_LAST_MODIFIED(ByVal vData As Date)
   mvarFIELD_DATE_LAST_MODIFIED = vData
End Property
Public Property Get FIELD_DATE_LAST_MODIFIED() As Date
   FIELD_DATE_LAST_MODIFIED = mvarFIELD_DATE_LAST_MODIFIED
End Property

' Code: >P13
' Name: FIELD_DATE_REQUESTED (Let/Get)
' Description: set/return FIELD_DATE_REQUESTED
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_REQUESTED(ByVal vData As Date)
   mvarFIELD_DATE_REQUESTED = vData
End Property
Public Property Get FIELD_DATE_REQUESTED() As Date
   FIELD_DATE_REQUESTED = mvarFIELD_DATE_REQUESTED
End Property

' Code: >P14
' Name: FIELD_DATE_SEND (Let/Get)
' Description: set/return FIELD_DATE_SEND
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_SEND(ByVal vData As Date)
   mvarFIELD_DATE_SEND = vData
End Property
Public Property Get FIELD_DATE_SEND() As Date
   FIELD_DATE_SEND = mvarFIELD_DATE_SEND
End Property

' Code: >P15
' Name: FIELD_HEADER (Let/Get)
' Description: set/return FIELD_HEADER
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_HEADER(ByVal vData As Integer)
   mvarFIELD_HEADER = vData
End Property
Public Property Get FIELD_HEADER() As Integer
   FIELD_HEADER = mvarFIELD_HEADER
End Property

' Code: >P16
' Name: FIELD_DETAIL (Let/Get)
' Description: set/return FIELD_DETAIL
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DETAIL(ByVal vData As Integer)
   mvarFIELD_DETAIL = vData
End Property
Public Property Get FIELD_DETAIL() As Integer
   FIELD_DETAIL = mvarFIELD_DETAIL
End Property

' Code: >P17
' Name: FIELD_DOCUMENT_COUNTER (Let/Get)
' Description: set/return FIELD_DOCUMENT_COUNTER
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DOCUMENT_COUNTER(ByVal vData As Integer)
   mvarFIELD_DOCUMENT_COUNTER = vData
End Property
Public Property Get FIELD_DOCUMENT_COUNTER() As Integer
   FIELD_DOCUMENT_COUNTER = mvarFIELD_DOCUMENT_COUNTER
End Property

' Code: >P18
' Name: FIELD_REMARKS (Let/Get)
' Description: set/return FIELD_REMARKS
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_REMARKS(ByVal vData As String)
   mvarFIELD_REMARKS = vData
End Property
Public Property Get FIELD_REMARKS() As String
   FIELD_REMARKS = mvarFIELD_REMARKS
End Property

' Code: >P19
' Name: FIELD_TREE_ID (Let/Get)
' Description: set/return FIELD_TREE_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_TREE_ID(ByVal vData As String)
   mvarFIELD_TREE_ID = vData
End Property
Public Property Get FIELD_TREE_ID() As String
   FIELD_TREE_ID = mvarFIELD_TREE_ID
End Property

' Code: >P20
' Name: FIELD_SUPPLIER_NAME (Let/Get)
' Description: set/return FIELD_SUPPLIER_NAME
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_SUPPLIER_NAME(ByVal vData As String)
   mvarFIELD_SUPPLIER_NAME = vData
End Property
Public Property Get FIELD_SUPPLIER_NAME() As String
   FIELD_SUPPLIER_NAME = mvarFIELD_SUPPLIER_NAME
End Property

' Code: >P21
' Name: FIELD_DTYPE (Let/Get)
' Description: set/return FIELD_DTYPE
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DTYPE(ByVal vData As Byte)
   mvarFIELD_DTYPE = vData
End Property
Public Property Get FIELD_DTYPE() As Byte
   FIELD_DTYPE = mvarFIELD_DTYPE
End Property

' Code: >P22
' Name: FIELD_Error_String (Let/Get)
' Description: set/return FIELD_Error_String
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Error_String(ByVal vData As String)
   mvarFIELD_Error_String = vData
End Property
Public Property Get FIELD_Error_String() As String
   FIELD_Error_String = mvarFIELD_Error_String
End Property

' Code: >P23
' Name: FIELD_Error_HD (Let/Get)
' Description: set/return FIELD_Error_HD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Error_HD(ByVal vData As String)
   mvarFIELD_Error_HD = vData
End Property
Public Property Get FIELD_Error_HD() As String
   FIELD_Error_HD = mvarFIELD_Error_HD
End Property

' Code: >P24
' Name: FIELD_LOGID (Let/Get)
' Description: set/return FIELD_LOGID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LOGID(ByVal vData As String)
   mvarFIELD_LOGID = vData
End Property
Public Property Get FIELD_LOGID() As String
   FIELD_LOGID = mvarFIELD_LOGID
End Property

' Code: >P25
' Name: FIELD_USERNAME (Let/Get)
' Description: set/return FIELD_USERNAME
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_USERNAME(ByVal vData As String)
   mvarFIELD_USERNAME = vData
End Property
Public Property Get FIELD_USERNAME() As String
   FIELD_USERNAME = mvarFIELD_USERNAME
End Property

' Code: >P26
' Name: FIELD_DATE_PRINTED (Let/Get)
' Description: set/return FIELD_DATE_PRINTED
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_PRINTED(ByVal vData As Date)
   mvarFIELD_DATE_PRINTED = vData
End Property
Public Property Get FIELD_DATE_PRINTED() As Date
   FIELD_DATE_PRINTED = mvarFIELD_DATE_PRINTED
End Property

' Code: >P27
' Name: FIELD_PRINTED_BY (Let/Get)
' Description: set/return FIELD_PRINTED_BY
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_PRINTED_BY(ByVal vData As String)
   mvarFIELD_PRINTED_BY = vData
End Property
Public Property Get FIELD_PRINTED_BY() As String
   FIELD_PRINTED_BY = mvarFIELD_PRINTED_BY
End Property

' Code: >P28
' Name: FIELD_MRN (Let/Get)
' Description: set/return FIELD_MRN
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_MRN(ByVal vData As String)
   mvarFIELD_MRN = vData
End Property
Public Property Get FIELD_MRN() As String
   FIELD_MRN = mvarFIELD_MRN
End Property
Public Property Let FIELD_REMOTE_ID(ByVal vData As Long)
   mvarFIELD_REMOTE_ID = vData
End Property
Public Property Get FIELD_REMOTE_ID() As Long
   FIELD_REMOTE_ID = mvarFIELD_REMOTE_ID
End Property

