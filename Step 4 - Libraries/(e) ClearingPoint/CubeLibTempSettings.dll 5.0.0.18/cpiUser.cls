VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarUser_ID As Long
Private mvarUser_Name As String
Private mvarUser_Password As String
Private mvarUser_Rights As Long
Private mvarUser_StateSettings As Long
Private mvarUser_MainDimensionPosition As String
Private mvarUser_TreeWidth As Long
Private mvarUser_Level As Long
Private mvarUser_ButtonGroup As Long
Private mvarUser_ButtonBarWidth As Long

Private mvarADMINISTRATOR_RIGHTS As Boolean ' 1
Private mvarMAINTAIN_TABLES As Boolean ' 2
Private mvarALL_LOGICAL_IDS As Boolean ' 3
Private mvarSHOW_ALL_SENT As Boolean ' 4
Private mvarSHOW_ALL_WITH_ERRORS As Boolean ' 5
Private mvarSHOW_ALL_WAITING As Boolean ' 6
Private mvarSHOW_ALL_DELETED As Boolean ' 7
Private mvarCLEAN_UP_DELETED As Boolean ' 8
Private mvarEVERY_FIELD As Integer  ' 9
Private mvarDAYS_OR_ITEMS As String ' 10
Private mvarDELETE_OTHER_USERS_ITEMS As Boolean ' 11
Private mvarDELETE_SENT_IN_N_DAYS As Integer  ' 12
Private mvarWITH_SECURITY As Boolean ' 13
Private mvarREFRESH_IN_SECONDS As Long  ' 14
Private mvarRELATE_L1_TO_S1 As Boolean ' 15
Private mvarSUPPLIER_BOX As String ' 16
Private mvarLAST_USED_PRINTER As String ' 17
Private mvarLAST_DV_PRINTER As String ' 18
Private mvarLOGID_DESCRIPTION As String ' 19
Private mvarFIXED_USER As Boolean ' 20
Private mvarLOGGED As Boolean ' 21
Private mvarSDICOUNT As Long  ' 22
Private mvarSDECOUNT As Long ' 23
Private mvarSDTCOUNT As Long ' 24
Private mvarDECOUNT As Long ' 25
Private mvarSHOW_ONLY_DOCTYPE As Boolean ' 26
Private mvarSHOW_ONLY_VATNUM As Boolean ' 27
Private mvarSHOW_ONLY_CTRYCODE As Boolean ' 28
Private mvarSAVE_NEW_TARICCTRY As Long ' 29
Private mvarSHOW_ALL_TOBEPRINTED As Boolean ' 30
Private mvarSHOW_ALL_DRAFTS As Boolean ' 31
Private mvarSDI2COUNT As Long ' 32
Private mvarSDE2COUNT As Long ' 33
Private mvarSDT2COUNT As Long ' 34
Private mvarLANGUAGE_FIELD As Byte  ' 35

' ADMINISTRATOR_RIGHTS 1 B
' MAINTAIN_TABLES 2 B
' ALL_LOGICAL_IDS 3 B
' SHOW_ALL_SENT 4 B
' SHOW_ALL_WITH_ERRORS 5 B
' SHOW_ALL_WAITING 6 B
' SHOW_ALL_DELETED 7 B
' CLEAN_UP_DELETED 8 B
' EVERY_FIELD 9 INT
' DAYS_OR_ITEMS 10 TXT
' DELETE_OTHER_USERS_ITEMS 11 B
' DELETE_SENT_IN_N_DAYS 12 INT
' WITH_SECURITY 13 B
' REFRESH_IN_SECONDS 14 LONG
' RELATE_L1_TO_S1 15  B
' SUPPLIER_BOX 16 TXT
' LAST_USED_PRINTER 17 T
' LAST_DV_PRINTER 18 T
' LOGID_DESCRIPTION 19 T
' FIXED_USER 20 B
' LOGGED 21 B
' SDICOUNT 22 L
' SDECOUNT 23 L
' SDTCOUNT 24 L
' DECOUNT 25 L
' SHOW_ONLY_DOCTYPE 26 B
' SHOW_ONLY_VATNUM 27 B
' SHOW_ONLY_CTRYCODE 28 B
' SAVE_NEW_TARICCTRY 29 LNG
' SHOW_ALL_TOBEPRINTED 30 B
' SHOW_ALL_DRAFTS 31 B
' SDI2COUNT 32 L
' SDE2COUNT 33 L
' SDT2COUNT 34 L
' LANGUAGE_FIELD 35 BYTE

Public Property Let User_ID(ByVal vData As Long)
    
    mvarUser_ID = vData

End Property

Public Property Get User_ID() As Long
    
    User_ID = mvarUser_ID

End Property

Public Property Let User_Name(ByVal vData As String)
    
    mvarUser_Name = vData

End Property

Public Property Get User_Name() As String
    
    User_Name = mvarUser_Name

End Property

Public Property Let User_Password(ByVal vData As String)
    
    mvarUser_Password = vData

End Property

Public Property Get User_Password() As String
    
    User_Password = mvarUser_Password

End Property

Public Property Let User_Rights(ByVal vData As Long)
    
    mvarUser_Rights = vData

End Property

Public Property Get User_Rights() As Long
    
    User_Rights = mvarUser_Rights

End Property

Public Property Let User_StateSettings(ByVal vData As Long)
    
    mvarUser_StateSettings = vData

End Property

Public Property Get User_StateSettings() As Long
    
    User_StateSettings = mvarUser_StateSettings

End Property

Public Property Let User_MainDimensionPosition(ByVal vData As String)
    
    mvarUser_MainDimensionPosition = vData

End Property

Public Property Get User_MainDimensionPosition() As String
    
    User_MainDimensionPosition = mvarUser_MainDimensionPosition

End Property

Public Property Let User_TreeWidth(ByVal vData As Long)
    
    mvarUser_TreeWidth = vData

End Property

Public Property Get User_TreeWidth() As Long
    
    User_TreeWidth = mvarUser_TreeWidth

End Property

Public Property Let User_Level(ByVal vData As Long)
    
    mvarUser_Level = vData

End Property

Public Property Get User_Level() As Long
    
    User_Level = mvarUser_Level

End Property

Public Property Let User_ButtonGroup(ByVal vData As Long)
    
    mvarUser_ButtonGroup = vData

End Property

Public Property Get User_ButtonGroup() As Long
    
    User_ButtonGroup = mvarUser_ButtonGroup

End Property

Public Property Let User_ButtonBarWidth(ByVal vData As Long)
    
    mvarUser_ButtonBarWidth = vData

End Property

Public Property Get User_ButtonBarWidth() As Long
    
    User_ButtonBarWidth = mvarUser_ButtonBarWidth

End Property

' 1
Public Property Let ADMINISTRATOR_RIGHTS(ByVal vData As Boolean)
    
    mvarADMINISTRATOR_RIGHTS = vData

End Property

Public Property Get ADMINISTRATOR_RIGHTS() As Boolean
    
    ADMINISTRATOR_RIGHTS = mvarADMINISTRATOR_RIGHTS

End Property

' 2
Public Property Let MAINTAIN_TABLES(ByVal vData As Boolean)
    
    mvarMAINTAIN_TABLES = vData

End Property

Public Property Get MAINTAIN_TABLES() As Boolean
    
    MAINTAIN_TABLES = mvarMAINTAIN_TABLES

End Property

' 3
Public Property Let ALL_LOGICAL_IDS(ByVal vData As Boolean)
    
    mvarALL_LOGICAL_IDS = vData

End Property

Public Property Get ALL_LOGICAL_IDS() As Boolean
    
    ALL_LOGICAL_IDS = mvarALL_LOGICAL_IDS

End Property

' 4
Public Property Let SHOW_ALL_SENT(ByVal vData As Boolean)
    
    mvarSHOW_ALL_SENT = vData

End Property

Public Property Get SHOW_ALL_SENT() As Boolean
    
    SHOW_ALL_SENT = mvarSHOW_ALL_SENT

End Property

' 5
Public Property Let SHOW_ALL_WITH_ERRORS(ByVal vData As Boolean)
    
    mvarSHOW_ALL_WITH_ERRORS = vData

End Property

Public Property Get SHOW_ALL_WITH_ERRORS() As Boolean
    
    SHOW_ALL_WITH_ERRORS = mvarSHOW_ALL_WITH_ERRORS

End Property

' 6
Public Property Let SHOW_ALL_WAITING(ByVal vData As Boolean)
    
    mvarSHOW_ALL_WAITING = vData

End Property

Public Property Get SHOW_ALL_WAITING() As Boolean
    
    SHOW_ALL_WAITING = mvarSHOW_ALL_WAITING

End Property

' 7
Public Property Let SHOW_ALL_DELETED(ByVal vData As Boolean)
    
    mvarSHOW_ALL_DELETED = vData

End Property

Public Property Get SHOW_ALL_DELETED() As Boolean
    
    SHOW_ALL_DELETED = mvarSHOW_ALL_DELETED

End Property

' 8
Public Property Let CLEAN_UP_DELETED(ByVal vData As Boolean)
    
    mvarCLEAN_UP_DELETED = vData

End Property

Public Property Get CLEAN_UP_DELETED() As Boolean
    
    CLEAN_UP_DELETED = mvarCLEAN_UP_DELETED

End Property

' 9
Public Property Let EVERY_FIELD(ByVal vData As Integer)
    
    mvarEVERY_FIELD = vData

End Property

Public Property Get EVERY_FIELD() As Integer
    
    EVERY_FIELD = mvarEVERY_FIELD

End Property

' 10
Public Property Let DAYS_OR_ITEMS(ByVal vData As String)
    
    mvarDAYS_OR_ITEMS = vData

End Property

Public Property Get DAYS_OR_ITEMS() As String
    
    DAYS_OR_ITEMS = mvarDAYS_OR_ITEMS

End Property

' 11
Public Property Let DELETE_OTHER_USERS_ITEMS(ByVal vData As Boolean)
    
    mvarDELETE_OTHER_USERS_ITEMS = vData

End Property

Public Property Get DELETE_OTHER_USERS_ITEMS() As Boolean
    
    DELETE_OTHER_USERS_ITEMS = mvarDELETE_OTHER_USERS_ITEMS

End Property

' 12
Public Property Let DELETE_SENT_IN_N_DAYS(ByVal vData As Integer)
    
    mvarDELETE_SENT_IN_N_DAYS = vData

End Property

Public Property Get DELETE_SENT_IN_N_DAYS() As Integer
    
    DELETE_SENT_IN_N_DAYS = mvarDELETE_SENT_IN_N_DAYS

End Property

' 13
Public Property Let WITH_SECURITY(ByVal vData As Boolean)
    
    mvarWITH_SECURITY = vData

End Property

Public Property Get WITH_SECURITY() As Boolean
    
    WITH_SECURITY = mvarWITH_SECURITY

End Property

' 14
Public Property Let REFRESH_IN_SECONDS(ByVal vData As Long)
    
    mvarREFRESH_IN_SECONDS = vData

End Property

Public Property Get REFRESH_IN_SECONDS() As Long
    
    REFRESH_IN_SECONDS = mvarREFRESH_IN_SECONDS

End Property

' 15
Public Property Let RELATE_L1_TO_S1(ByVal vData As Boolean)
    
    mvarRELATE_L1_TO_S1 = vData

End Property

Public Property Get RELATE_L1_TO_S1() As Boolean
    
    RELATE_L1_TO_S1 = mvarRELATE_L1_TO_S1

End Property

' 16
Public Property Let SUPPLIER_BOX(ByVal vData As String)
    
    mvarSUPPLIER_BOX = vData

End Property

Public Property Get SUPPLIER_BOX() As String
    
    SUPPLIER_BOX = mvarSUPPLIER_BOX

End Property

' 17
Public Property Let LAST_USED_PRINTER(ByVal vData As String)
    
    mvarLAST_USED_PRINTER = vData

End Property

Public Property Get LAST_USED_PRINTER() As String
    
    LAST_USED_PRINTER = mvarLAST_USED_PRINTER

End Property

' 18
Public Property Let LAST_DV_PRINTER(ByVal vData As String)
    
    mvarLAST_DV_PRINTER = vData

End Property

Public Property Get LAST_DV_PRINTER() As String
    
    LAST_DV_PRINTER = mvarLAST_DV_PRINTER

End Property

' 19
Public Property Let LOGID_DESCRIPTION(ByVal vData As String)
    
    mvarLOGID_DESCRIPTION = vData

End Property

Public Property Get LOGID_DESCRIPTION() As String
    
    LOGID_DESCRIPTION = mvarLOGID_DESCRIPTION

End Property

' 20
Public Property Let FIXED_USER(ByVal vData As Boolean)
    
    mvarFIXED_USER = vData

End Property

Public Property Get FIXED_USER() As Boolean
    
    FIXED_USER = mvarFIXED_USER

End Property

' 21
Public Property Let LOGGED(ByVal vData As Boolean)
    
    mvarLOGGED = vData

End Property

Public Property Get LOGGED() As Boolean
    
    LOGGED = mvarLOGGED

End Property

' 22
Public Property Let SDICOUNT(ByVal vData As Long)
    
    mvarSDICOUNT = vData

End Property

Public Property Get SDICOUNT() As Long
    
    SDICOUNT = mvarSDICOUNT

End Property

' 23
Public Property Let SDECOUNT(ByVal vData As Long)
    
    mvarSDECOUNT = vData

End Property

Public Property Get SDECOUNT() As Long
    
    SDECOUNT = mvarSDECOUNT

End Property

' 24
Public Property Let SDTCOUNT(ByVal vData As Long)
    
    mvarSDTCOUNT = vData

End Property

Public Property Get SDTCOUNT() As Long
    
    SDTCOUNT = mvarSDTCOUNT

End Property

' 25
Public Property Let DECOUNT(ByVal vData As Long)
    
    mvarDECOUNT = vData

End Property

Public Property Get DECOUNT() As Long
    
    DECOUNT = mvarDECOUNT

End Property

' 26
Public Property Let SHOW_ONLY_DOCTYPE(ByVal vData As Boolean)
    
    mvarSHOW_ONLY_DOCTYPE = vData

End Property

Public Property Get SHOW_ONLY_DOCTYPE() As Boolean
    
    SHOW_ONLY_DOCTYPE = mvarSHOW_ONLY_DOCTYPE

End Property

' 27
Public Property Let SHOW_ONLY_VATNUM(ByVal vData As Boolean)
    
    mvarSHOW_ONLY_VATNUM = vData

End Property

Public Property Get SHOW_ONLY_VATNUM() As Boolean
    
    SHOW_ONLY_VATNUM = mvarSHOW_ONLY_VATNUM

End Property

' 28
Public Property Let SHOW_ONLY_CTRYCODE(ByVal vData As Boolean)
    
    mvarSHOW_ONLY_CTRYCODE = vData

End Property

Public Property Get SHOW_ONLY_CTRYCODE() As Boolean
    
    SHOW_ONLY_CTRYCODE = mvarSHOW_ONLY_CTRYCODE

End Property

' 29
Public Property Let SAVE_NEW_TARICCTRY(ByVal vData As Long)
    
    mvarSAVE_NEW_TARICCTRY = vData

End Property

Public Property Get SAVE_NEW_TARICCTRY() As Long
    
    SAVE_NEW_TARICCTRY = mvarSAVE_NEW_TARICCTRY

End Property

' 30
Public Property Let SHOW_ALL_TOBEPRINTED(ByVal vData As Boolean)
    
    mvarSHOW_ALL_TOBEPRINTED = vData

End Property

Public Property Get SHOW_ALL_TOBEPRINTED() As Boolean
    
    SHOW_ALL_TOBEPRINTED = mvarSHOW_ALL_TOBEPRINTED

End Property

' 31
Public Property Let SHOW_ALL_DRAFTS(ByVal vData As Boolean)
    
    mvarSHOW_ALL_DRAFTS = vData

End Property

Public Property Get SHOW_ALL_DRAFTS() As Boolean
    
    SHOW_ALL_DRAFTS = mvarSHOW_ALL_DRAFTS

End Property

' 32
Public Property Let SDI2COUNT(ByVal vData As Long)
    
    mvarSDI2COUNT = vData

End Property

Public Property Get SDI2COUNT() As Long
    
    SDI2COUNT = mvarSDI2COUNT

End Property

' 33
Public Property Let SDE2COUNT(ByVal vData As Long)
    
    mvarSDE2COUNT = vData

End Property

Public Property Get SDE2COUNT() As Long
    
    SDE2COUNT = mvarSDE2COUNT

End Property

' 34
Public Property Let SDT2COUNT(ByVal vData As Long)
    
    mvarSDT2COUNT = vData

End Property

Public Property Get SDT2COUNT() As Long
    
    SDT2COUNT = mvarSDT2COUNT

End Property

' 35
Public Property Let LANGUAGE_FIELD(ByVal vData As Byte)
    
    mvarLANGUAGE_FIELD = vData

End Property

Public Property Get LANGUAGE_FIELD() As Byte
    
    LANGUAGE_FIELD = mvarLANGUAGE_FIELD

End Property


