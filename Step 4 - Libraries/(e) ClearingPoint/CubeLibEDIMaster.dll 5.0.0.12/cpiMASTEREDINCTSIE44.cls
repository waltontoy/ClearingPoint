VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiMASTEREDINCTSIE44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiMasterEdiNctsIE44
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: IE44 Properties, mapped from MasterEdiNctsIE44 table
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
' >P11 - Public Property Let ADMINISTRATOR_RIGHTS(ByVal vData As Boolean)
'        Public Property Get ADMINISTRATOR_RIGHTS() As Boolean
' >P12 - Public Property Let MAINTAIN_TABLES(ByVal vData As Boolean)
'        Public Property Get MAINTAIN_TABLES() As Boolean
' >P13 - Public Property Let ALL_LOGICAL_IDS(ByVal vData As Boolean)
'        Public Property Get ALL_LOGICAL_IDS() As Boolean
' >P14 - Public Property Let SHOW_ALL_SENT(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_SENT() As Boolean
' >P15 - Public Property Let SHOW_ALL_WITH_ERRORS(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_WITH_ERRORS() As Boolean
' >P16 - Public Property Let SHOW_ALL_WAITING(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_WAITING() As Boolean
' >P17 - Public Property Let SHOW_ALL_DELETED(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_DELETED() As Boolean
' >P18 - Public Property Let CLEAN_UP_DELETED(ByVal vData As Boolean)
'        Public Property Get CLEAN_UP_DELETED() As Boolean
' >P19 - Public Property Let EVERY_FIELD(ByVal vData As Integer)
'        Public Property Get EVERY_FIELD() As Integer
' >P20 - Public Property Let DAYS_OR_ITEMS(ByVal vData As String)
'        Public Property Get DAYS_OR_ITEMS() As String
' >P21 - Public Property Let DELETE_OTHER_USERS_ITEMS(ByVal vData As Boolean)
'        Public Property Get DELETE_OTHER_USERS_ITEMS() As Boolean
' >P22 - Public Property Let DELETE_SENT_IN_N_DAYS(ByVal vData As Integer)
'        Public Property Get DELETE_SENT_IN_N_DAYS() As Integer
' >P23 - Public Property Let WITH_SECURITY(ByVal vData As Boolean)
'        Public Property Get WITH_SECURITY() As Boolean
' >P24 - Public Property Let REFRESH_IN_SECONDS(ByVal vData As Long)
'        Public Property Get REFRESH_IN_SECONDS() As Long
' >P25 - Public Property Let RELATE_L1_TO_S1(ByVal vData As Boolean)
'        Public Property Get RELATE_L1_TO_S1() As Boolean
' >P26 - Public Property Let SUPPLIER_BOX(ByVal vData As String)
'        Public Property Get SUPPLIER_BOX() As String
' >P27 - Public Property Let LAST_USED_PRINTER(ByVal vData As String)
'        Public Property Get LAST_USED_PRINTER() As String
' >P28 - Public Property Let LAST_DV_PRINTER(ByVal vData As String)
'        Public Property Get LAST_DV_PRINTER() As String
' >P29 - Public Property Let LOGID_DESCRIPTION(ByVal vData As String)
'        Public Property Get LOGID_DESCRIPTION() As String
' >P30 - Public Property Let FIXED_USER(ByVal vData As Boolean)
'        Public Property Get FIXED_USER() As Boolean
' >P31 - Public Property Let LOGGED(ByVal vData As Boolean)
'        Public Property Get LOGGED() As Boolean
' >P32 - Public Property Let SDICOUNT(ByVal vData As Long)
'        Public Property Get SDICOUNT() As Long
' >P33 - Public Property Let SDECOUNT(ByVal vData As Long)
'        Public Property Get SDECOUNT() As Long
' >P34 - Public Property Let SDTCOUNT(ByVal vData As Long)
'        Public Property Get SDTCOUNT() As Long
' >P35 - Public Property Let DECOUNT(ByVal vData As Long)
'        Public Property Get DECOUNT() As Long
' >P36 - Public Property Let SHOW_ONLY_DOCTYPE(ByVal vData As Boolean)
'        Public Property Get SHOW_ONLY_DOCTYPE() As Boolean
' >P37 - Public Property Let SHOW_ONLY_VATNUM(ByVal vData As Boolean)
'        Public Property Get SHOW_ONLY_VATNUM() As Boolean
' >P38 - Public Property Let SHOW_ONLY_CTRYCODE(ByVal vData As Boolean)
'        Public Property Get SHOW_ONLY_CTRYCODE() As Boolean
' >P39 - Public Property Let SAVE_NEW_TARICCTRY(ByVal vData As Long)
'        Public Property Get SAVE_NEW_TARICCTRY() As Long
' >P40 - Public Property Let SHOW_ALL_TOBEPRINTED(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_TOBEPRINTED() As Boolean
' >P41 - Public Property Let SHOW_ALL_DRAFTS(ByVal vData As Boolean)
'        Public Property Get SHOW_ALL_DRAFTS() As Boolean
' >P42 - Public Property Let SDI2COUNT(ByVal vData As Long)
'        Public Property Get SDI2COUNT() As Long
' >P43 - Public Property Let SDE2COUNT(ByVal vData As Long)
'        Public Property Get SDE2COUNT() As Long
' >P44 - Public Property Let SDT2COUNT(ByVal vData As Long)
'        Public Property Get SDT2COUNT() As Long
' >P45 - Public Property Let LANGUAGE_FIELD(ByVal vData As Byte)
'        Public Property Get LANGUAGE_FIELD() As Byte

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
' <V13 - Private mvarALL_LOGICAL_IDS As Boolean
' <V14 - Private mvarSHOW_ALL_SENT As Boolean
' <V15 - Private mvarSHOW_ALL_WITH_ERRORS As Boolean
' <V16 - Private mvarSHOW_ALL_WAITING As Boolean
' <V17 - Private mvarSHOW_ALL_DELETED As Boolean
' <V18 - Private mvarCLEAN_UP_DELETED As Boolean
' <V19 - Private mvarEVERY_FIELD As Integer
' <V20 - Private mvarDAYS_OR_ITEMS As String
' <V21 - Private mvarDELETE_OTHER_USERS_ITEMS As Boolean
' <V22 - Private mvarDELETE_SENT_IN_N_DAYS As Integer
' <V23 - Private mvarWITH_SECURITY As Boolean
' <V24 - Private mvarREFRESH_IN_SECONDS As Long
' <V25 - Private mvarRELATE_L1_TO_S1 As Boolean
' <V26 - Private mvarSUPPLIER_BOX As String
' <V27 - Private mvarLAST_USED_PRINTER As String
' <V28 - Private mvarLAST_DV_PRINTER As String
' <V29 - Private mvarLOGID_DESCRIPTION As String
' <V30 - Private mvarFIXED_USER As Boolean
' <V31 - Private mvarLOGGED As Boolean
' <V32 - Private mvarSDICOUNT As Long
' <V33 - Private mvarSDECOUNT As Long
' <V34 - Private mvarSDTCOUNT As Long
' <V35 - Private mvarDECOUNT As Long
' <V36 - Private mvarSHOW_ONLY_DOCTYPE As Boolean
' <V37 - Private mvarSHOW_ONLY_VATNUM As Boolean
' <V38 - Private mvarSHOW_ONLY_CTRYCODE As Boolean
' <V39 - Private mvarSAVE_NEW_TARICCTRY As Long
' <V40 - Private mvarSHOW_ALL_TOBEPRINTED As Boolean
' <V41 - Private mvarSHOW_ALL_DRAFTS As Boolean
' <V42 - Private mvarSDI2COUNT As Long
' <V43 - Private mvarSDE2COUNT As Long
' <V44 - Private mvarSDT2COUNT As Long
' <V45 - Private mvarLANGUAGE_FIELD As Byte
'
' Objects: 0 (0/0)
' ************************************************************************************************
' ************************************************************************************************

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_CODE As String ' <V1
Private mvarFIELD_DTYPE As Byte ' <V2
Private mvarFIELD_DOCUMENT_NAME As String ' <V3
Private mvarFIELD_TREE_ID As String ' <V4
Private mvarFIELD_DATE_CREATED As Date ' <V5
Private mvarFIELD_DATE_LAST_MODIFIED As Date ' <V6
Private mvarFIELD_DATE_REQUESTED As Date ' <V7
Private mvarFIELD_DATE_SEND As Date ' <V8
Private mvarFIELD_LAST_MODIFIED_BY As String ' <V9
Private mvarFIELD_LOGID_DESCRIPTION As String ' <V10
Private mvarFIELD_REMARKS As String ' <V11
Private mvarFIELD_USER_NO As Long ' <V12
Private mvarFIELD_DOC_NUMBER As String ' <V13
Private mvarFIELD_DOC_TYPE As String ' <V14
Private mvarFIELD_Error_String As String ' <V15
Private mvarFIELD_Error_HD As String ' <V16
Private mvarFIELD_ORIGIN As String ' <V17
Private mvarFIELD_Type As String ' <V18
Private mvarFIELD_COMM As String ' <V19
Private mvarFIELD_LOGID As String ' <V20
Private mvarFIELD_PRINT As String ' <V21
Private mvarFIELD_VIEWED As Byte ' <V22
Private mvarFIELD_USERNAME As String ' <V23
Private mvarFIELD_HEADER As Integer ' <V24
Private mvarFIELD_Memo_Field As String ' <V25
Private mvarFIELD_DATE_PRINTED As Date ' <V26
Private mvarFIELD_PRINTED_BY As String ' <V27
Private mvarFIELD_MR As String ' <V28
Private mvarFIELD_BD As String ' <V29
Private mvarFIELD_AJ As String ' <V30
Private mvarFIELD_W8 As String ' <V31
Private mvarFIELD_W9 As String ' <V32
Private mvarFIELD_WA As String ' <V33
Private mvarFIELD_WB As String ' <V34
Private mvarFIELD_WD As String ' <V35
Private mvarFIELD_T7 As String ' <V36
Private mvarFIELD_WE As String ' <V37
Private mvarFIELD_SB As String ' <V38
Private mvarFIELD_DATE_LAST_RECEIVED As Date ' <V39
Private mvarFIELD_A9 As String ' <V40
Private mvarFIELD_AE As String ' <V41
Private mvarFIELD_AF As String ' <V42
Private mvarFIELD_B1 As String ' <V43
Private mvarFIELD_B7 As String ' <V44
Private mvarFIELD_L1 As String ' <V45
Private mvarFIELD_M1 As String ' <V46
Private mvarFIELD_M2 As String ' <V47
Private mvarFIELD_S1 As String ' <V48
Private mvarFIELD_S2 As String ' <V49
Private mvarFIELD_S3 As String ' <V50
Private mvarFIELD_S4 As String ' <V51
Private mvarFIELD_S5 As String ' <V52
Private mvarFIELD_S6 As String ' <V53
Private mvarFIELD_S7 As String ' <V54
Private mvarFIELD_S8 As String ' <V55
Private mvarFIELD_S9 As String ' <V56
Private mvarFIELD_SA As String ' <V57
Private mvarFIELD_V1 As String ' <V58
Private mvarFIELD_V2 As String ' <V59
Private mvarFIELD_V3 As String ' <V60
Private mvarFIELD_V4 As String ' <V61
Private mvarFIELD_V5 As String ' <V62
Private mvarFIELD_V6 As String ' <V63
Private mvarFIELD_V7 As String ' <V64
Private mvarFIELD_V8 As String ' <V65
Private mvarFIELD_Y1 As String ' <V66
Private mvarFIELD_Y2 As String ' <V67
Private mvarFIELD_Y3 As String ' <V68
Private mvarFIELD_Y4 As String ' <V69
Private mvarFIELD_Y5 As String ' <V70
Private mvarFIELD_CC As String ' <V71
Private mvarFIELD_CD As String ' <V72
Private mvarFIELD_CE As String ' <V73
Private mvarFIELD_CF As String ' <V74
Private mvarFIELD_CG As String ' <V75
Private mvarFIELD_CI As String ' <V76
Private mvarFIELD_UA As String ' <V77
Private mvarFIELD_UC As String ' <V78
Private mvarFIELD_UD As String ' <V79
Private mvarFIELD_UE As String ' <V80
Private mvarFIELD_UF As String ' <V81
Private mvarFIELD_CL As String ' <V82
Private mvarFIELD_CH As String ' <V83
Private mvarFIELD_REMOTE_ID As Long ' 84

' Code: >P1
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property
Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

' Code: >P2
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DTYPE(ByVal vData As Byte)
   mvarFIELD_DTYPE = vData
End Property
Public Property Get FIELD_DTYPE() As Byte
   FIELD_DTYPE = mvarFIELD_DTYPE
End Property

' Code: >P3
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DOCUMENT_NAME(ByVal vData As String)
   mvarFIELD_DOCUMENT_NAME = vData
End Property
Public Property Get FIELD_DOCUMENT_NAME() As String
   FIELD_DOCUMENT_NAME = mvarFIELD_DOCUMENT_NAME
End Property

' Code: >P4
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_TREE_ID(ByVal vData As String)
   mvarFIELD_TREE_ID = vData
End Property
Public Property Get FIELD_TREE_ID() As String
   FIELD_TREE_ID = mvarFIELD_TREE_ID
End Property

' Code: >P5
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_CREATED(ByVal vData As Date)
   mvarFIELD_DATE_CREATED = vData
End Property
Public Property Get FIELD_DATE_CREATED() As Date
   FIELD_DATE_CREATED = mvarFIELD_DATE_CREATED
End Property

' Code: >P6
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_LAST_MODIFIED(ByVal vData As Date)
   mvarFIELD_DATE_LAST_MODIFIED = vData
End Property
Public Property Get FIELD_DATE_LAST_MODIFIED() As Date
   FIELD_DATE_LAST_MODIFIED = mvarFIELD_DATE_LAST_MODIFIED
End Property

' Code: >P7
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_REQUESTED(ByVal vData As Date)
   mvarFIELD_DATE_REQUESTED = vData
End Property
Public Property Get FIELD_DATE_REQUESTED() As Date
   FIELD_DATE_REQUESTED = mvarFIELD_DATE_REQUESTED
End Property

' Code: >P8
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_SEND(ByVal vData As Date)
   mvarFIELD_DATE_SEND = vData
End Property
Public Property Get FIELD_DATE_SEND() As Date
   FIELD_DATE_SEND = mvarFIELD_DATE_SEND
End Property

' Code: >P9
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LAST_MODIFIED_BY(ByVal vData As String)
   mvarFIELD_LAST_MODIFIED_BY = vData
End Property
Public Property Get FIELD_LAST_MODIFIED_BY() As String
   FIELD_LAST_MODIFIED_BY = mvarFIELD_LAST_MODIFIED_BY
End Property

' Code: >P10
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property
Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

' Code: >P11
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_REMARKS(ByVal vData As String)
   mvarFIELD_REMARKS = vData
End Property
Public Property Get FIELD_REMARKS() As String
   FIELD_REMARKS = mvarFIELD_REMARKS
End Property

' Code: >P12
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_USER_NO(ByVal vData As Long)
   mvarFIELD_USER_NO = vData
End Property
Public Property Get FIELD_USER_NO() As Long
   FIELD_USER_NO = mvarFIELD_USER_NO
End Property

' Code: >P13
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DOC_NUMBER(ByVal vData As String)
   mvarFIELD_DOC_NUMBER = vData
End Property
Public Property Get FIELD_DOC_NUMBER() As String
   FIELD_DOC_NUMBER = mvarFIELD_DOC_NUMBER
End Property

' Code: >P14
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DOC_TYPE(ByVal vData As String)
   mvarFIELD_DOC_TYPE = vData
End Property
Public Property Get FIELD_DOC_TYPE() As String
   FIELD_DOC_TYPE = mvarFIELD_DOC_TYPE
End Property

' Code: >P15
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Error_String(ByVal vData As String)
   mvarFIELD_Error_String = vData
End Property
Public Property Get FIELD_Error_String() As String
   FIELD_Error_String = mvarFIELD_Error_String
End Property

' Code: >P16
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Error_HD(ByVal vData As String)
   mvarFIELD_Error_HD = vData
End Property
Public Property Get FIELD_Error_HD() As String
   FIELD_Error_HD = mvarFIELD_Error_HD
End Property

' Code: >P17
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_ORIGIN(ByVal vData As String)
   mvarFIELD_ORIGIN = vData
End Property
Public Property Get FIELD_ORIGIN() As String
   FIELD_ORIGIN = mvarFIELD_ORIGIN
End Property

' Code: >P18
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Type(ByVal vData As String)
   mvarFIELD_Type = vData
End Property
Public Property Get FIELD_Type() As String
   FIELD_Type = mvarFIELD_Type
End Property

' Code: >P19
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_COMM(ByVal vData As String)
   mvarFIELD_COMM = vData
End Property
Public Property Get FIELD_COMM() As String
   FIELD_COMM = mvarFIELD_COMM
End Property

' Code: >P20
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_LOGID(ByVal vData As String)
   mvarFIELD_LOGID = vData
End Property
Public Property Get FIELD_LOGID() As String
   FIELD_LOGID = mvarFIELD_LOGID
End Property

' Code: >P21
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_PRINT(ByVal vData As String)
   mvarFIELD_PRINT = vData
End Property
Public Property Get FIELD_PRINT() As String
   FIELD_PRINT = mvarFIELD_PRINT
End Property

' Code: >P22
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_VIEWED(ByVal vData As Byte)
   mvarFIELD_VIEWED = vData
End Property
Public Property Get FIELD_VIEWED() As Byte
   FIELD_VIEWED = mvarFIELD_VIEWED
End Property

' Code: >P23
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_USERNAME(ByVal vData As String)
   mvarFIELD_USERNAME = vData
End Property
Public Property Get FIELD_USERNAME() As String
   FIELD_USERNAME = mvarFIELD_USERNAME
End Property

' Code: >P24
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_HEADER(ByVal vData As Integer)
   mvarFIELD_HEADER = vData
End Property
Public Property Get FIELD_HEADER() As Integer
   FIELD_HEADER = mvarFIELD_HEADER
End Property

' Code: >P25
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Memo_Field(ByVal vData As String)
   mvarFIELD_Memo_Field = vData
End Property
Public Property Get FIELD_Memo_Field() As String
   FIELD_Memo_Field = mvarFIELD_Memo_Field
End Property

' Code: >P26
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
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
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
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
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_MR(ByVal vData As String)
   mvarFIELD_MR = vData
End Property
Public Property Get FIELD_MR() As String
   FIELD_MR = mvarFIELD_MR
End Property

' Code: >P29
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_BD(ByVal vData As String)
   mvarFIELD_BD = vData
End Property
Public Property Get FIELD_BD() As String
   FIELD_BD = mvarFIELD_BD
End Property

' Code: >P30
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_AJ(ByVal vData As String)
   mvarFIELD_AJ = vData
End Property
Public Property Get FIELD_AJ() As String
   FIELD_AJ = mvarFIELD_AJ
End Property

' Code: >P31
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_W8(ByVal vData As String)
   mvarFIELD_W8 = vData
End Property
Public Property Get FIELD_W8() As String
   FIELD_W8 = mvarFIELD_W8
End Property

' Code: >P32
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_W9(ByVal vData As String)
   mvarFIELD_W9 = vData
End Property
Public Property Get FIELD_W9() As String
   FIELD_W9 = mvarFIELD_W9
End Property

' Code: >P33
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_WA(ByVal vData As String)
   mvarFIELD_WA = vData
End Property
Public Property Get FIELD_WA() As String
   FIELD_WA = mvarFIELD_WA
End Property

' Code: >P34
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_WB(ByVal vData As String)
   mvarFIELD_WB = vData
End Property
Public Property Get FIELD_WB() As String
   FIELD_WB = mvarFIELD_WB
End Property

' Code: >P35
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_WD(ByVal vData As String)
   mvarFIELD_WD = vData
End Property
Public Property Get FIELD_WD() As String
   FIELD_WD = mvarFIELD_WD
End Property

' Code: >P36
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_T7(ByVal vData As String)
   mvarFIELD_T7 = vData
End Property
Public Property Get FIELD_T7() As String
   FIELD_T7 = mvarFIELD_T7
End Property

' Code: >P37
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_WE(ByVal vData As String)
   mvarFIELD_WE = vData
End Property
Public Property Get FIELD_WE() As String
   FIELD_WE = mvarFIELD_WE
End Property

' Code: >P38
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_SB(ByVal vData As String)
   mvarFIELD_SB = vData
End Property
Public Property Get FIELD_SB() As String
   FIELD_SB = mvarFIELD_SB
End Property

' Code: >P39
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_DATE_LAST_RECEIVED(ByVal vData As Date)
   mvarFIELD_DATE_LAST_RECEIVED = vData
End Property
Public Property Get FIELD_DATE_LAST_RECEIVED() As Date
   FIELD_DATE_LAST_RECEIVED = mvarFIELD_DATE_LAST_RECEIVED
End Property

Public Property Let FIELD_REMOTE_ID(ByVal vData As Long)
   mvarFIELD_REMOTE_ID = vData
End Property

Public Property Get FIELD_REMOTE_ID() As Long
   FIELD_REMOTE_ID = mvarFIELD_REMOTE_ID
End Property

' Code: >P40
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_A9(ByVal vData As String)
   mvarFIELD_A9 = vData
End Property
Public Property Get FIELD_A9() As String
   FIELD_A9 = mvarFIELD_A9
End Property

' Code: >P41
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_AE(ByVal vData As String)
   mvarFIELD_AE = vData
End Property
Public Property Get FIELD_AE() As String
   FIELD_AE = mvarFIELD_AE
End Property

' Code: >P42
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_AF(ByVal vData As String)
   mvarFIELD_AF = vData
End Property
Public Property Get FIELD_AF() As String
   FIELD_AF = mvarFIELD_AF
End Property

' Code: >P43
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_B1(ByVal vData As String)
   mvarFIELD_B1 = vData
End Property
Public Property Get FIELD_B1() As String
   FIELD_B1 = mvarFIELD_B1
End Property

' Code: >P44
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_B7(ByVal vData As String)
   mvarFIELD_B7 = vData
End Property
Public Property Get FIELD_B7() As String
   FIELD_B7 = mvarFIELD_B7
End Property

' Code: >P45
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_L1(ByVal vData As String)
   mvarFIELD_L1 = vData
End Property
Public Property Get FIELD_L1() As String
   FIELD_L1 = mvarFIELD_L1
End Property

' Code: >P46
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_M1(ByVal vData As String)
   mvarFIELD_M1 = vData
End Property
Public Property Get FIELD_M1() As String
   FIELD_M1 = mvarFIELD_M1
End Property

' Code: >P47
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_M2(ByVal vData As String)
   mvarFIELD_M2 = vData
End Property

Public Property Get FIELD_M2() As String
   FIELD_M2 = mvarFIELD_M2
End Property

' Code: >P48
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S1(ByVal vData As String)
   mvarFIELD_S1 = vData
End Property
Public Property Get FIELD_S1() As String
   FIELD_S1 = mvarFIELD_S1
End Property

' Code: >P49
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S2(ByVal vData As String)
   mvarFIELD_S2 = vData
End Property
Public Property Get FIELD_S2() As String
   FIELD_S2 = mvarFIELD_S2
End Property

' Code: >P50
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S3(ByVal vData As String)
   mvarFIELD_S3 = vData
End Property
Public Property Get FIELD_S3() As String
   FIELD_S3 = mvarFIELD_S3
End Property

' Code: >P51
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S4(ByVal vData As String)
   mvarFIELD_S4 = vData
End Property
Public Property Get FIELD_S4() As String
   FIELD_S4 = mvarFIELD_S4
End Property

' Code: >P52
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S5(ByVal vData As String)
   mvarFIELD_S5 = vData
End Property
Public Property Get FIELD_S5() As String
   FIELD_S5 = mvarFIELD_S5
End Property

' Code: >P53
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S6(ByVal vData As String)
   mvarFIELD_S6 = vData
End Property
Public Property Get FIELD_S6() As String
   FIELD_S6 = mvarFIELD_S6
End Property

' Code: >P54
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S7(ByVal vData As String)
   mvarFIELD_S7 = vData
End Property
Public Property Get FIELD_S7() As String
   FIELD_S7 = mvarFIELD_S7
End Property

' Code: >P55
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S8(ByVal vData As String)
   mvarFIELD_S8 = vData
End Property
Public Property Get FIELD_S8() As String
   FIELD_S8 = mvarFIELD_S8
End Property

' Code: >P56
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_S9(ByVal vData As String)
   mvarFIELD_S9 = vData
End Property
Public Property Get FIELD_S9() As String
   FIELD_S9 = mvarFIELD_S9
End Property

' Code: >P57
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_SA(ByVal vData As String)
   mvarFIELD_SA = vData
End Property
Public Property Get FIELD_SA() As String
   FIELD_SA = mvarFIELD_SA
End Property

' Code: >P58
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V1(ByVal vData As String)
   mvarFIELD_V1 = vData
End Property
Public Property Get FIELD_V1() As String
   FIELD_V1 = mvarFIELD_V1
End Property

' Code: >P59
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V2(ByVal vData As String)
   mvarFIELD_V2 = vData
End Property
Public Property Get FIELD_V2() As String
   FIELD_V2 = mvarFIELD_V2
End Property

' Code: >P60
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V3(ByVal vData As String)
   mvarFIELD_V3 = vData
End Property
Public Property Get FIELD_V3() As String
   FIELD_V3 = mvarFIELD_V3
End Property

' Code: >P61
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V4(ByVal vData As String)
   mvarFIELD_V4 = vData
End Property
Public Property Get FIELD_V4() As String
   FIELD_V4 = mvarFIELD_V4
End Property

' Code: >P62
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V5(ByVal vData As String)
   mvarFIELD_V5 = vData
End Property
Public Property Get FIELD_V5() As String
   FIELD_V5 = mvarFIELD_V5
End Property

' Code: >P63
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V6(ByVal vData As String)
   mvarFIELD_V6 = vData
End Property
Public Property Get FIELD_V6() As String
   FIELD_V6 = mvarFIELD_V6
End Property

' Code: >P64
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V7(ByVal vData As String)
   mvarFIELD_V7 = vData
End Property
Public Property Get FIELD_V7() As String
   FIELD_V7 = mvarFIELD_V7
End Property

' Code: >P65
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_V8(ByVal vData As String)
   mvarFIELD_V8 = vData
End Property
Public Property Get FIELD_V8() As String
   FIELD_V8 = mvarFIELD_V8
End Property

' Code: >P66
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Y1(ByVal vData As String)
   mvarFIELD_Y1 = vData
End Property
Public Property Get FIELD_Y1() As String
   FIELD_Y1 = mvarFIELD_Y1
End Property

' Code: >P67
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Y2(ByVal vData As String)
   mvarFIELD_Y2 = vData
End Property
Public Property Get FIELD_Y2() As String
   FIELD_Y2 = mvarFIELD_Y2
End Property

' Code: >P68
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Y3(ByVal vData As String)
   mvarFIELD_Y3 = vData
End Property
Public Property Get FIELD_Y3() As String
   FIELD_Y3 = mvarFIELD_Y3
End Property

' Code: >P69
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Y4(ByVal vData As String)
   mvarFIELD_Y4 = vData
End Property
Public Property Get FIELD_Y4() As String
   FIELD_Y4 = mvarFIELD_Y4
End Property

' Code: >P70
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_Y5(ByVal vData As String)
   mvarFIELD_Y5 = vData
End Property
Public Property Get FIELD_Y5() As String
   FIELD_Y5 = mvarFIELD_Y5
End Property

' Code: >P71
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CC(ByVal vData As String)
   mvarFIELD_CC = vData
End Property
Public Property Get FIELD_CC() As String
   FIELD_CC = mvarFIELD_CC
End Property

' Code: >P72
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CD(ByVal vData As String)
   mvarFIELD_CD = vData
End Property
Public Property Get FIELD_CD() As String
   FIELD_CD = mvarFIELD_CD
End Property

' Code: >P73
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CE(ByVal vData As String)
   mvarFIELD_CE = vData
End Property
Public Property Get FIELD_CE() As String
   FIELD_CE = mvarFIELD_CE
End Property

' Code: >P74
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CF(ByVal vData As String)
   mvarFIELD_CF = vData
End Property
Public Property Get FIELD_CF() As String
   FIELD_CF = mvarFIELD_CF
End Property

' Code: >P75
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CG(ByVal vData As String)
   mvarFIELD_CG = vData
End Property
Public Property Get FIELD_CG() As String
   FIELD_CG = mvarFIELD_CG
End Property

' Code: >P76
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CI(ByVal vData As String)
   mvarFIELD_CI = vData
End Property
Public Property Get FIELD_CI() As String
   FIELD_CI = mvarFIELD_CI
End Property

' Code: >P77
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_UA(ByVal vData As String)
   mvarFIELD_UA = vData
End Property
Public Property Get FIELD_UA() As String
   FIELD_UA = mvarFIELD_UA
End Property

' Code: >P78
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_UC(ByVal vData As String)
   mvarFIELD_UC = vData
End Property
Public Property Get FIELD_UC() As String
   FIELD_UC = mvarFIELD_UC
End Property

' Code: >P79
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_UD(ByVal vData As String)
   mvarFIELD_UD = vData
End Property
Public Property Get FIELD_UD() As String
   FIELD_UD = mvarFIELD_UD
End Property

' Code: >P80
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_UE(ByVal vData As String)
   mvarFIELD_UE = vData
End Property
Public Property Get FIELD_UE() As String
   FIELD_UE = mvarFIELD_UE
End Property

' Code: >P81
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_UF(ByVal vData As String)
   mvarFIELD_UF = vData
End Property
Public Property Get FIELD_UF() As String
   FIELD_UF = mvarFIELD_UF
End Property

' Code: >P82
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CL(ByVal vData As String)
   mvarFIELD_CL = vData
End Property
Public Property Get FIELD_CL() As String
   FIELD_CL = mvarFIELD_CL
End Property

' Code: >P83
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let FIELD_CH(ByVal vData As String)
   mvarFIELD_CH = vData
End Property
Public Property Get FIELD_CH() As String
   FIELD_CH = mvarFIELD_CH
End Property


