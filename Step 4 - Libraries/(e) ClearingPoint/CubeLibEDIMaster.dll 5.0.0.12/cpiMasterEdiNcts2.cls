VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiMasterEdiNcts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiMasterEdiNcts2
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Tuesday, September 30, 2003
' Description: IE07 Properties, mapped from MasterEdiNcts2 table
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

Private mvarCODE_FIELD As String  ' <V1
Private mvarDTYPE_FIELD As Byte   ' <V2
Private mvarDOCUMENT_NAME_FIELD As String ' <V3
Private mvarTREE_ID_FIELD  As String   ' <V4
Private mvarDATE_CREATED_FIELD  As Date   ' <V5
Private mvarDATE_LAST_MODIFIED_FIELD  As Date   ' <V6
Private mvarDATE_REQUESTED_FIELD As Date   ' <V7
Private mvarDATE_SEND_FIELD  As Date ' <V8
Private mvarLAST_MODIFIED_BY_FIELD  As String   ' <V9
Private mvarLOGID_DESCRIPTION_FIELD  As String   ' <V10
Private mvarREMARKS_FIELD  As String   ' <V11
Private mvarUSER_NO_FIELD  As Long   ' <V12
Private mvarDOC_NUMBER_FIELD  As String   ' <V13
Private mvarDOC_TYPE_FIELD  As String   ' <V14
Private mvarError_String_FIELD  As String   ' <V15
Private mvarError_HD_FIELD  As String   ' <V16
Private mvarORIGIN_FIELD  As String   ' <V17
Private mvarType_FIELD As String   ' <V18
Private mvarCOMM_FIELD  As String   ' <V19
Private mvarLOGID_FIELD  As String   ' <V20
Private mvarPrint_FIELD  As String   ' <V21
Private mvarVIEWED_FIELD  As Byte   ' <V22
Private mvarUSERNAME_FIELD As String    ' <V23
Private mvarHEADER_FIELD As Integer    ' <V24
Private mvarMemo_Field_FIELD As String    ' <V25
Private mvarDATE_PRINTED_FIELD As Date    ' <V26
Private mvarPRINTED_BY_FIELD As String    ' <V27
Private mvarMD_FIELD As String    ' <V28
Private mvarA6_FIELD As String  ' <V29
Private mvarME_FIELD As String    ' <V30
Private mvarMR_FIELD As String    ' <V31
Private mvarAH_FIELD As String    ' <V32
Private mvarAG_FIELD As String    ' <V33
Private mvarBD_FIELD As String    ' <V34
Private mvarBC_FIELD As String    ' <V35
Private mvarAI_FIELD As String    ' <V36
Private mvarMC_FIELD As String    ' <V37
Private mvarEP_FIELD As String    ' <V38
Private mvarAJ_FIELD As String    ' <V39
Private mvarBF_FIELD As String    ' <V40
Private mvarEQ_FIELD As String    ' <V41
Private mvarC7_FIELD As String    ' <V42
Private mvarER_FIELD As String    ' <V43
Private mvarC8_FIELD As String    ' <V44
Private mvarC9_FIELD As String    ' <V45
Private mvarCA_FIELD As String    ' <V46
Private mvarCB_FIELD As String    ' <V47
Private mvarES_FIELD As String   ' <V48
Private mvarAK_FIELD As String    ' <V49
Private mvarAL_FIELD As String    ' <V50
Private mvarAM_FIELD As String    ' <V51
Private mvarBG_FIELD As String    ' <V52
Private mvarBH_FIELD As String    ' <V53
Private mvarBI_FIELD As String    ' <V54
Private mvarBJ_FIELD As String    ' <V55
Private mvarBK_FIELD As String    ' <V56
Private mvarBL_FIELD As String    ' <V57
Private mvarSC_FIELD As String    ' <V58
Private mvarSD_FIELD As String    ' <V59
Private mvarSE_FIELD As String    ' <V60
Private mvarSF_FIELD As String    ' <V61
Private mvarSG_FIELD As String    ' <V62
Private mvarW8_FIELD As String    ' <V63
Private mvarW9_FIELD As String    ' <V64
Private mvarWA_FIELD As String    ' <V65
Private mvarWB_FIELD As String    ' <V66
Private mvarWC_FIELD As String    ' <V67
Private mvarWD_FIELD As String    ' <V68
Private mvarT7_FIELD As String    ' <V69
Private mvarWE_FIELD As String    ' <V70
Private mvarSB_FIELD As String    ' <V71
Private mvarDATE_LAST_RECEIVED_FIELD As Date ' <V72
Private mvarFIELD_REMOTE_ID As Long ' 72
Private mvarZA_FIELD As String    ' <V73
Private mvarZB_FIELD As String    ' <V74

' Code: >P1
' Name: CODE_FIELD (Let/Get)
' Description: set/return CODE_FIELD
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property
Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' Code: >P2
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DTYPE_FIELD(ByVal vData As Byte)
    mvarDTYPE_FIELD = vData
End Property
Public Property Get DTYPE_FIELD() As Byte
    DTYPE_FIELD = mvarDTYPE_FIELD
End Property

' Code: >P3
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DOCUMENT_NAME_FIELD(ByVal vData As String)
    mvarDOCUMENT_NAME_FIELD = vData
End Property
Public Property Get DOCUMENT_NAME_FIELD() As String
    DOCUMENT_NAME_FIELD = mvarDOCUMENT_NAME_FIELD
End Property

' Code: >P4
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let TREE_ID_FIELD(ByVal vData As String)
    mvarTREE_ID_FIELD = vData
End Property
Public Property Get TREE_ID_FIELD() As String
    TREE_ID_FIELD = mvarTREE_ID_FIELD
End Property

' Code: >P5
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_CREATED_FIELD(ByVal vData As Date)
    mvarDATE_CREATED_FIELD = vData
End Property
Public Property Get DATE_CREATED_FIELD() As Date
    DATE_CREATED_FIELD = mvarDATE_CREATED_FIELD
End Property

' Code: >P6
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_LAST_MODIFIED_FIELD(ByVal vData As Date)
    mvarDATE_LAST_MODIFIED_FIELD = vData
End Property
Public Property Get DATE_LAST_MODIFIED_FIELD() As Date
    DATE_LAST_MODIFIED_FIELD = mvarDATE_LAST_MODIFIED_FIELD
End Property

' Code: >P7
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_REQUESTED_FIELD(ByVal vData As Date)
    mvarDATE_REQUESTED_FIELD = vData
End Property
Public Property Get DATE_REQUESTED_FIELD() As Date
    DATE_REQUESTED_FIELD = mvarDATE_REQUESTED_FIELD
End Property

' Code: >P8
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_SEND_FIELD(ByVal vData As Date)
    mvarDATE_SEND_FIELD = vData
End Property
Public Property Get DATE_SEND_FIELD() As Date
    DATE_SEND_FIELD = mvarDATE_SEND_FIELD
End Property

' Code: >P9
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let LAST_MODIFIED_BY_FIELD(ByVal vData As String)
    mvarLAST_MODIFIED_BY_FIELD = vData
End Property
Public Property Get LAST_MODIFIED_BY_FIELD() As String
    LAST_MODIFIED_BY_FIELD = mvarLAST_MODIFIED_BY_FIELD
End Property

' Code: >P10
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let LOGID_DESCRIPTION_FIELD(ByVal vData As String)
    mvarLOGID_DESCRIPTION_FIELD = vData
End Property
Public Property Get LOGID_DESCRIPTION_FIELD() As String
    LOGID_DESCRIPTION_FIELD = mvarLOGID_DESCRIPTION_FIELD
End Property

' Code: >P11
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let REMARKS_FIELD(ByVal vData As String)
    mvarREMARKS_FIELD = vData
End Property
Public Property Get REMARKS_FIELD() As String
    REMARKS_FIELD = mvarREMARKS_FIELD
End Property

' Code: >P12
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let USER_NO_FIELD(ByVal vData As Long)
    mvarUSER_NO_FIELD = vData
End Property
Public Property Get USER_NO_FIELD() As Long
    USER_NO_FIELD = mvarUSER_NO_FIELD
End Property

' Code: >P13
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DOC_NUMBER_FIELD(ByVal vData As String)
    mvarDOC_NUMBER_FIELD = vData
End Property
Public Property Get DOC_NUMBER_FIELD() As String
    DOC_NUMBER_FIELD = mvarDOC_NUMBER_FIELD
End Property

' Code: >P14
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DOC_TYPE_FIELD(ByVal vData As String)
    mvarDOC_TYPE_FIELD = vData
End Property
Public Property Get DOC_TYPE_FIELD() As String
    DOC_TYPE_FIELD = mvarDOC_TYPE_FIELD
End Property

' Code: >P15
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let Error_String_FIELD(ByVal vData As String)
    mvarError_String_FIELD = vData
End Property
Public Property Get Error_String_FIELD() As String
    Error_String_FIELD = mvarError_String_FIELD
End Property

' Code: >P16
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let Error_HD_FIELD(ByVal vData As String)
    mvarError_HD_FIELD = vData
End Property
Public Property Get Error_HD_FIELD() As String
    Error_HD_FIELD = mvarError_HD_FIELD
End Property

' Code: >P17
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ORIGIN_FIELD(ByVal vData As String)
    mvarORIGIN_FIELD = vData
End Property
Public Property Get ORIGIN_FIELD() As String
    ORIGIN_FIELD = mvarORIGIN_FIELD
End Property

' Code: >P18
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let Type_FIELD(ByVal vData As String)
    mvarType_FIELD = vData
End Property
Public Property Get Type_FIELD() As String
    Type_FIELD = mvarType_FIELD
End Property

' Code: >P19
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let COMM_FIELD(ByVal vData As String)
    mvarCOMM_FIELD = vData
End Property
Public Property Get COMM_FIELD() As String
    COMM_FIELD = mvarCOMM_FIELD
End Property

' Code: >P20
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let LOGID_FIELD(ByVal vData As String)
    mvarLOGID_FIELD = vData
End Property
Public Property Get LOGID_FIELD() As String
    LOGID_FIELD = mvarLOGID_FIELD
End Property

' Code: >P21
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let Print_FIELD(ByVal vData As String)
    mvarPrint_FIELD = vData
End Property
Public Property Get Print_FIELD() As String
    Print_FIELD = mvarPrint_FIELD
End Property

' Code: >P22
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let VIEWED_FIELD(ByVal vData As Byte)
    mvarVIEWED_FIELD = vData
End Property
Public Property Get VIEWED_FIELD() As Byte
    VIEWED_FIELD = mvarVIEWED_FIELD
End Property

' Code: >P23
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let USERNAME_FIELD(ByVal vData As String)
    mvarUSERNAME_FIELD = vData
End Property
Public Property Get USERNAME_FIELD() As String
    USERNAME_FIELD = mvarUSERNAME_FIELD
End Property

' Code: >P24
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let HEADER_FIELD(ByVal vData As Integer)
    mvarHEADER_FIELD = vData
End Property
Public Property Get HEADER_FIELD() As Integer
    HEADER_FIELD = mvarHEADER_FIELD
End Property

' Code: >P25
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let Memo_Field_FIELD(ByVal vData As String)
    mvarMemo_Field_FIELD = vData
End Property
Public Property Get Memo_Field_FIELD() As String
    Memo_Field_FIELD = mvarMemo_Field_FIELD
End Property

' Code: >P26
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_PRINTED_FIELD(ByVal vData As Date)
    mvarDATE_PRINTED_FIELD = vData
End Property
Public Property Get DATE_PRINTED_FIELD() As Date
    DATE_PRINTED_FIELD = mvarDATE_PRINTED_FIELD
End Property

' Code: >P27
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let PRINTED_BY_FIELD(ByVal vData As String)
    mvarPRINTED_BY_FIELD = vData
End Property
Public Property Get PRINTED_BY_FIELD() As String
    PRINTED_BY_FIELD = mvarPRINTED_BY_FIELD
End Property

' Code: >P28
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let MD_FIELD(ByVal vData As String)
    mvarMD_FIELD = vData
End Property
Public Property Get MD_FIELD() As String
    MD_FIELD = mvarMD_FIELD
End Property

' Code: >P29
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ME_FIELD(ByVal vData As String)
    mvarME_FIELD = vData
End Property
Public Property Get ME_FIELD() As String
    ME_FIELD = mvarME_FIELD
End Property

' Code: >P30
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let A6_FIELD(ByVal vData As String)
    mvarA6_FIELD = vData
End Property
Public Property Get A6_FIELD() As String
    A6_FIELD = mvarA6_FIELD
End Property

' Code: >P73
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ZA_FIELD(ByVal vData As String)
    mvarZA_FIELD = vData
End Property
Public Property Get ZA_FIELD() As String
    ZA_FIELD = mvarZA_FIELD
End Property

' Code: >P74
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ZB_FIELD(ByVal vData As String)
    mvarZB_FIELD = vData
End Property
Public Property Get ZB_FIELD() As String
    ZB_FIELD = mvarZB_FIELD
End Property

' Code: >P31
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let MR_FIELD(ByVal vData As String)
    mvarMR_FIELD = vData
End Property
Public Property Get MR_FIELD() As String
    MR_FIELD = mvarMR_FIELD
End Property

' Code: >P32
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AH_FIELD(ByVal vData As String)
    mvarAH_FIELD = vData
End Property
Public Property Get AH_FIELD() As String
    AH_FIELD = mvarAH_FIELD
End Property

' Code: >P33
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AG_FIELD(ByVal vData As String)
    mvarAG_FIELD = vData
End Property
Public Property Get AG_FIELD() As String
    AG_FIELD = mvarAG_FIELD
End Property

' Code: >P34
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BD_FIELD(ByVal vData As String)
    mvarBD_FIELD = vData
End Property
Public Property Get BD_FIELD() As String
    BD_FIELD = mvarBD_FIELD
End Property

' Code: >P35
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BC_FIELD(ByVal vData As String)
    mvarBC_FIELD = vData
End Property
Public Property Get BC_FIELD() As String
    BC_FIELD = mvarBC_FIELD
End Property

' Code: >P36
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AI_FIELD(ByVal vData As String)
    mvarAI_FIELD = vData
End Property
Public Property Get AI_FIELD() As String
    AI_FIELD = mvarAI_FIELD
End Property

' Code: >P37
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let MC_FIELD(ByVal vData As String)
    mvarMC_FIELD = vData
End Property
Public Property Get MC_FIELD() As String
    MC_FIELD = mvarMC_FIELD
End Property

' Code: >P38
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let EP_FIELD(ByVal vData As String)
    mvarEP_FIELD = vData
End Property
Public Property Get EP_FIELD() As String
    EP_FIELD = mvarEP_FIELD
End Property

' Code: >P39
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AJ_FIELD(ByVal vData As String)
    mvarAJ_FIELD = vData
End Property
Public Property Get AJ_FIELD() As String
    AJ_FIELD = mvarAJ_FIELD
End Property

' Code: >P40
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BF_FIELD(ByVal vData As String)
    mvarBF_FIELD = vData
End Property
Public Property Get BF_FIELD() As String
    BF_FIELD = mvarBF_FIELD
End Property

' Code: >P41
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let EQ_FIELD(ByVal vData As String)
    mvarEQ_FIELD = vData
End Property
Public Property Get EQ_FIELD() As String
    EQ_FIELD = mvarEQ_FIELD
End Property

' Code: >P42
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let C7_FIELD(ByVal vData As String)
    mvarC7_FIELD = vData
End Property
Public Property Get C7_FIELD() As String
    C7_FIELD = mvarC7_FIELD
End Property

' Code: >P43
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ER_FIELD(ByVal vData As String)
    mvarER_FIELD = vData
End Property
Public Property Get ER_FIELD() As String
    ER_FIELD = mvarER_FIELD
End Property

' Code: >P44
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let C8_FIELD(ByVal vData As String)
    mvarC8_FIELD = vData
End Property
Public Property Get C8_FIELD() As String
    C8_FIELD = mvarC8_FIELD
End Property

' Code: >P45
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let C9_FIELD(ByVal vData As String)
    mvarC9_FIELD = vData
End Property
Public Property Get C9_FIELD() As String
    C9_FIELD = mvarC9_FIELD
End Property

' Code: >P46
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let CA_FIELD(ByVal vData As String)
    mvarCA_FIELD = vData
End Property
Public Property Get CA_FIELD() As String
    CA_FIELD = mvarCA_FIELD
End Property

' Code: >P47
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let CB_FIELD(ByVal vData As String)
    mvarCB_FIELD = vData
End Property
Public Property Get CB_FIELD() As String
    CB_FIELD = mvarCB_FIELD
End Property

' Code: >P48
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let ES_FIELD(ByVal vData As String)
    mvarES_FIELD = vData
End Property
Public Property Get ES_FIELD() As String
    ES_FIELD = mvarES_FIELD
End Property

' Code: >P49
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AK_FIELD(ByVal vData As String)
    mvarAK_FIELD = vData
End Property
Public Property Get AK_FIELD() As String
    AK_FIELD = mvarAK_FIELD
End Property

' Code: >P50
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AL_FIELD(ByVal vData As String)
    mvarAL_FIELD = vData
End Property
Public Property Get AL_FIELD() As String
    AL_FIELD = mvarAL_FIELD
End Property

' Code: >P51
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let AM_FIELD(ByVal vData As String)
    mvarAM_FIELD = vData
End Property
Public Property Get AM_FIELD() As String
    AM_FIELD = mvarAM_FIELD
End Property

' Code: >P52
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BG_FIELD(ByVal vData As String)
    mvarBG_FIELD = vData
End Property
Public Property Get BG_FIELD() As String
    BG_FIELD = mvarBG_FIELD
End Property

' Code: >P53
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BH_FIELD(ByVal vData As String)
    mvarBH_FIELD = vData
End Property
Public Property Get BH_FIELD() As String
    BH_FIELD = mvarBH_FIELD
End Property

' Code: >P54
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BI_FIELD(ByVal vData As String)
    mvarBI_FIELD = vData
End Property
Public Property Get BI_FIELD() As String
    BI_FIELD = mvarBI_FIELD
End Property

' Code: >P55
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BJ_FIELD(ByVal vData As String)
    mvarBJ_FIELD = vData
End Property
Public Property Get BJ_FIELD() As String
    BJ_FIELD = mvarBJ_FIELD
End Property

' Code: >P56
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BK_FIELD(ByVal vData As String)
    mvarBK_FIELD = vData
End Property
Public Property Get BK_FIELD() As String
    BK_FIELD = mvarBK_FIELD
End Property

' Code: >P57
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let BL_FIELD(ByVal vData As String)
    mvarBL_FIELD = vData
End Property
Public Property Get BL_FIELD() As String
    BL_FIELD = mvarBL_FIELD
End Property

' Code: >P58
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SC_FIELD(ByVal vData As String)
    mvarSC_FIELD = vData
End Property
Public Property Get SC_FIELD() As String
    SC_FIELD = mvarSC_FIELD
End Property

' Code: >P59
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SD_FIELD(ByVal vData As String)
    mvarSD_FIELD = vData
End Property
Public Property Get SD_FIELD() As String
    SD_FIELD = mvarSD_FIELD
End Property

' Code: >P60
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SE_FIELD(ByVal vData As String)
    mvarSE_FIELD = vData
End Property
Public Property Get SE_FIELD() As String
    SE_FIELD = mvarSE_FIELD
End Property

' Code: >P61
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SF_FIELD(ByVal vData As String)
    mvarSF_FIELD = vData
End Property
Public Property Get SF_FIELD() As String
    SF_FIELD = mvarSF_FIELD
End Property

' Code: >P62
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SG_FIELD(ByVal vData As String)
    mvarSG_FIELD = vData
End Property
Public Property Get SG_FIELD() As String
    SG_FIELD = mvarSG_FIELD
End Property

' Code: >P63
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let W8_FIELD(ByVal vData As String)
    mvarW8_FIELD = vData
End Property
Public Property Get W8_FIELD() As String
    W8_FIELD = mvarW8_FIELD
End Property

' Code: >P64
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let W9_FIELD(ByVal vData As String)
    mvarW9_FIELD = vData
End Property
Public Property Get W9_FIELD() As String
    W9_FIELD = mvarW9_FIELD
End Property

' Code: >P65
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let WA_FIELD(ByVal vData As String)
    mvarWA_FIELD = vData
End Property
Public Property Get WA_FIELD() As String
    WA_FIELD = mvarWA_FIELD
End Property

' Code: >P66
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let WB_FIELD(ByVal vData As String)
    mvarWB_FIELD = vData
End Property
Public Property Get WB_FIELD() As String
    WB_FIELD = mvarWB_FIELD
End Property

' Code: >P67
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let WC_FIELD(ByVal vData As String)
    mvarWC_FIELD = vData
End Property
Public Property Get WC_FIELD() As String
    WC_FIELD = mvarWC_FIELD
End Property

' Code: >P68
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let WD_FIELD(ByVal vData As String)
    mvarWD_FIELD = vData
End Property
Public Property Get WD_FIELD() As String
    WD_FIELD = mvarWD_FIELD
End Property

' Code: >P69
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let T7_FIELD(ByVal vData As String)
    mvarT7_FIELD = vData
End Property
Public Property Get T7_FIELD() As String
    T7_FIELD = mvarT7_FIELD
End Property

' Code: >P70
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let WE_FIELD(ByVal vData As String)
    mvarWE_FIELD = vData
End Property
Public Property Get WE_FIELD() As String
    WE_FIELD = mvarWE_FIELD
End Property

' Code: >P71
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let SB_FIELD(ByVal vData As String)
    mvarSB_FIELD = vData
End Property
Public Property Get SB_FIELD() As String
    SB_FIELD = mvarSB_FIELD
End Property

' Code: >P72
' Name: Lic_ID (Let/Get)
' Description: set/return Lic_ID
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Let DATE_LAST_RECEIVED_FIELD(ByVal vData As Date)
    mvarDATE_LAST_RECEIVED_FIELD = vData
End Property
Public Property Get DATE_LAST_RECEIVED_FIELD() As Date
    DATE_LAST_RECEIVED_FIELD = mvarDATE_LAST_RECEIVED_FIELD
End Property
Public Property Let FIELD_REMOTE_ID(ByVal vData As Long)
   mvarFIELD_REMOTE_ID = vData

End Property

Public Property Get FIELD_REMOTE_ID() As Long
   FIELD_REMOTE_ID = mvarFIELD_REMOTE_ID
End Property

