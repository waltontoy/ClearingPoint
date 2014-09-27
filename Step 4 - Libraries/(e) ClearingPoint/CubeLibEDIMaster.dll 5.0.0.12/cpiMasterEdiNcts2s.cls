VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiMasterEdiNcts2s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiMasterEdiNcts2s
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: Collection of cpiMasterEdiNcts2 class
' Methods: 9 (0/9)
' >M1 - Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveRecord As cpiUser) As Boolean
' >M2 - Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveRecord As cpiUser) As Boolean
' >M3 - Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveRecord As cpiUser) As Boolean
' >M4 - Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, _
                          ByRef ActiveRecord As cpiUser) As Boolean
' >M5 - Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, _
                         ByRef ActiveRecord As cpiUser) As Boolean
' >M6 - Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, _
                         ByRef ActiveRecord As cpiUser) As Boolean
' >M7 - Public Function GetTableRecord(ByRef ActiveRecord As cpiUser) As ADODB.Recordset
' >M8 - Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiUser
' >M9 - Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, _
                             ByVal SearchField, ByVal SearchValue) As Boolean
'
' Properties: 0 (0/0)
'
' Events: 0 (0/0)
'
' API Methods: 0 (0/0)
'
' Constants: 52 (52/0)
' <C1 - Private Const TABLE_NAME = "Users"
' <C2 - Private Const PK_FIELD = "User_ID"
' <C3 - Private Const FIELD_User_ID = "User_ID"
' <C4 - Private Const FIELD_User_Name = "User_Name"
' <C5 - Private Const FIELD_User_Password = "User_Password"
' <C6 - Private Const FIELD_User_Rights = "User_Rights"
' <C7 - Private Const FIELD_User_StateSettings = "User_StateSettings"
' <C8 - Private Const FIELD_User_MainDimensionPosition = "User_MainDimensionPosition"
' <C9 - Private Const FIELD_User_TreeWidth = "User_TreeWidth"
' <C10 - Private Const FIELD_User_Level = "User_Level"
' <C11 - Private Const FIELD_User_ButtonGroup = "User_ButtonGroup"
' <C12 - Private Const FIELD_User_ButtonBarWidth = "User_ButtonBarWidth"
' <C13 - Private Const FIELD_ADMINISTRATOR_RIGHTS = "ADMINISTRATOR RIGHTS"
' <C14 - Private Const FIELD_MAINTAIN_TABLES = "MAINTAIN TABLES"
' <C15 - Private Const FIELD_ALL_LOGICAL_IDS = "ALL LOGICAL IDS"
' <C16 - Private Const FIELD_SHOW_ALL_SENT = "SHOW ALL SENT"
' <C17 - Private Const FIELD_SHOW_ALL_WITH_ERRORS = "SHOW ALL WITH ERRORS"
' <C18 - Private Const FIELD_SHOW_ALL_WAITING = "SHOW ALL WAITING"
' <C19 - Private Const FIELD_SHOW_ALL_DELETED = "SHOW ALL DELETED"
' <C20 - Private Const FIELD_CLEAN_UP_DELETED = "CLEAN UP DELETED"
' <C21 - Private Const FIELD_EVERY_FIELD = "EVERY"
' <C22 - Private Const FIELD_DAYS_OR_ITEMS = "DAYS OR ITEMS"
' <C23 - Private Const FIELD_DELETE_OTHER_USERS_ITEMS = "DELETE OTHER USERS ITEMS"
' <C24 - Private Const FIELD_DELETE_SENT_IN_N_DAYS = "DELETE SENT IN N DAYS"
' <C25 - Private Const FIELD_WITH_SECURITY = "WITH SECURITY"
' <C26 - Private Const FIELD_REFRESH_IN_SECONDS = "REFRESH IN SECONDS"
' <C27 - Private Const FIELD_RELATE_L1_TO_S1 = "RELATE L1 TO S1"
' <C28 - Private Const FIELD_SUPPLIER_BOX = "SUPPLIER BOX"
' <C29 - Private Const FIELD_LAST_USED_PRINTER = "LAST USED PRINTER"
' <C30 - Private Const FIELD_LAST_DV_PRINTER = "LAST DV PRINTER"
' <C31 - Private Const FIELD_LOGID_DESCRIPTION = "LOGID DESCRIPTION"
' <C32 - Private Const FIELD_FIXED_USER = "FIXED USER"
' <C33 - Private Const FIELD_LOGGED = "LOGGED"
' <C34 - Private Const FIELD_SDICOUNT = "SDICOUNT"
' <C35 - Private Const FIELD_SDECOUNT = "SDECOUNT"
' <C36 - Private Const FIELD_SDTCOUNT = "SDTCOUNT"
' <C37 - Private Const FIELD_DECOUNT = "DECOUNT"
' <C38 - Private Const FIELD_SHOW_ONLY_DOCTYPE = "SHOW ONLY DOCTYPE"
' <C39 - Private Const FIELD_SHOW_ONLY_VATNUM = "SHOW ONLY VATNUM"
' <C40 - Private Const FIELD_SHOW_ONLY_CTRYCODE = "SHOW ONLY CTRYCODE"
' <C41 - Private Const FIELD_SAVE_NEW_TARICCTRY = "SAVE NEW TARICCTRY"
' <C42 - Private Const FIELD_SHOW_ALL_TOBEPRINTED = "SHOW ALL TOBEPRINTED"
' <C43 - Private Const FIELD_SHOW_ALL_DRAFTS = "SHOW ALL DRAFTS"
' <C44 - Private Const FIELD_SDI2COUNT = "SDI2COUNT"
' <C45 - Private Const FIELD_SDE2COUNT = "SDE2COUNT"
' <C46 - Private Const FIELD_SDT2COUNT = "SDT2COUNT"
' <C47 - Private Const FIELD_LANGUAGE_FIELD = "LANGUAGE"
' <C48 - "INSERT INTO [Users] ([User_ID],[User_Name],[User_Password],[User_Rights]," & _
              "[User_StateSettings],[User_MainDimensionPosition],[User_TreeWidth]," & _
              "[User_Level],[User_ButtonGroup],[User_ButtonBarWidth]," & _
              "[ADMINISTRATOR RIGHTS],[MAINTAIN TABLES],[ALL LOGICAL IDS],[SHOW ALL SENT]," & _
              "[SHOW ALL WITH ERRORS],[SHOW ALL WAITING],[SHOW ALL DELETED]," & _
              "[CLEAN UP DELETED],[EVERY],[DAYS OR ITEMS],[DELETE OTHER USERS ITEMS]," & _
              "[DELETE SENT IN N DAYS],[WITH SECURITY],[REFRESH IN SECONDS]," & _
              "[RELATE L1 TO S1],[SUPPLIER BOX],[LAST USED PRINTER],[LAST DV PRINTER]," & _
              "[LOGID DESCRIPTION],[FIXED USER],[LOGGED],[SDICOUNT],[SDECOUNT],[SDTCOUNT]," & _
              "[DECOUNT],[SHOW ONLY DOCTYPE],[SHOW ONLY VATNUM],[SHOW ONLY CTRYCODE]," & _
              "[SAVE NEW TARICCTRY],[SHOW ALL TOBEPRINTED],[SHOW ALL DRAFTS],[SDI2COUNT]," & _
              "[SDE2COUNT],[SDT2COUNT],[LANGUAGE]" & _
              ") VALUES "
' <C49 - Private Const SQL_DELETE_RECORD = "DELETE * FROM Users WHERE [User_ID]="
' <C50 - "SELECT [User_ID],[User_Name],[User_Password],[User_Rights]," & _
              "[User_StateSettings],[User_MainDimensionPosition],[User_TreeWidth]," & _
              "[User_Level],[User_ButtonGroup],[User_ButtonBarWidth]," & _
              "[ADMINISTRATOR RIGHTS],[MAINTAIN TABLES],[ALL LOGICAL IDS],[SHOW ALL SENT]," & _
              "[SHOW ALL WITH ERRORS],[SHOW ALL WAITING],[SHOW ALL DELETED]," & _
              "[CLEAN UP DELETED],[EVERY],[DAYS OR ITEMS],[DELETE OTHER USERS ITEMS]," & _
              "[DELETE SENT IN N DAYS],[WITH SECURITY],[REFRESH IN SECONDS]," & _
              "[RELATE L1 TO S1],[SUPPLIER BOX],[LAST USED PRINTER],[LAST DV PRINTER]," & _
              "[LOGID DESCRIPTION],[FIXED USER],[LOGGED],[SDICOUNT],[SDECOUNT],[SDTCOUNT]," & _
              "[DECOUNT],[SHOW ONLY DOCTYPE],[SHOW ONLY VATNUM],[SHOW ONLY CTRYCODE]," & _
              "[SAVE NEW TARICCTRY],[SHOW ALL TOBEPRINTED],[SHOW ALL DRAFTS],[SDI2COUNT]," & _
              "[SDE2COUNT],[SDT2COUNT],[LANGUAGE]" & _
              "FROM Users WHERE [User_ID]= "
' <C51 - Private Const SQL_GET_MAXID = "SELECT MAX(User_ID) AS [ID_MAX] FROM Users"
' <C52 - Private Const SQL_GET_MINID = "SELECT Min(User_ID) AS [ID_MIN] FROM Users"
'
' Enumerations: 0 (0/0)
' ************************************************************************************************
' ************************************************************************************************

Option Explicit
    
Private Const TABLE_NAME = "[MASTEREDINCTS2]" ' <C1
Private Const PK_FIELD = "[CODE]" ' <C2
    
Private Const FIELD_CODE_FIELD = "CODE" ' <C3
Private Const FIELD_DTYPE_FIELD = "DTYPE" ' <C4
Private Const FIELD_DOCUMENT_NAME_FIELD = "DOCUMENT NAME"  ' <C5
Private Const FIELD_TREE_ID_FIELD = "TREE ID"  ' <C6
Private Const FIELD_DATE_CREATED_FIELD = "DATE CREATED" ' <C7
Private Const FIELD_DATE_LAST_MODIFIED_FIELD = "DATE LAST MODIFIED" ' <C8
Private Const FIELD_DATE_REQUESTED_FIELD = "DATE REQUESTED" ' <C9
Private Const FIELD_DATE_SEND_FIELD = "DATE SEND" ' <C10
Private Const FIELD_LAST_MODIFIED_BY_FIELD = "LAST MODIFIED BY" ' <C11
Private Const FIELD_LOGID_DESCRIPTION_FIELD = "LOGID DESCRIPTION" ' <C12
Private Const FIELD_REMARKS_FIELD = "REMARKS" ' <C13
Private Const FIELD_USER_NO_FIELD = "USER NO" ' <C14
Private Const FIELD_DOC_NUMBER_FIELD = "DOC NUMBER" ' <C15
Private Const FIELD_DOC_TYPE_FIELD = "DOC TYPE" ' <C16
Private Const FIELD_Error_String_FIELD = "Error String" ' <C17
Private Const FIELD_Error_HD_FIELD = "Error HD" ' <C18
Private Const FIELD_ORIGIN_FIELD = "ORIGIN" ' <C19
Private Const FIELD_Type_FIELD = "Type" ' <C20
Private Const FIELD_COMM_FIELD = "COMM" ' <C21
Private Const FIELD_LOGID_FIELD = "LOGID" ' <C22
Private Const FIELD_PRINT_FIELD = "PRINT" ' <C23
Private Const FIELD_VIEWED_FIELD = "VIEWED" ' <C24
Private Const FIELD_USERNAME_FIELD = "USERNAME" ' <C25

Private Const FIELD_HEADER_FIELD = "HEADER" ' <C26
Private Const FIELD_Memo_Field_FIELD = "Memo Field" ' <C27
Private Const FIELD_DATE_PRINTED_FIELD = "DATE PRINTED" ' <C28
Private Const FIELD_PRINTED_BY_FIELD = "PRINTED BY" ' <C29
Private Const FIELD_MD_FIELD = "MD" ' <C30
Private Const FIELD_ME_FIELD = "ME" ' <C31
Private Const FIELD_MR_FIELD = "MR" ' <C32
Private Const FIELD_AH_FIELD = "AH" ' <C33
Private Const FIELD_AG_FIELD = "AG" ' <C34
Private Const FIELD_BD_FIELD = "BD" ' <C35
Private Const FIELD_BC_FIELD = "BC" ' <C36
Private Const FIELD_AI_FIELD = "AI" ' <C37
Private Const FIELD_MC_FIELD = "MC" ' <C38
Private Const FIELD_EP_FIELD = "EP" ' <C39
Private Const FIELD_AJ_FIELD = "AJ" ' <C40
Private Const FIELD_BF_FIELD = "BF" ' <C41
Private Const FIELD_EQ_FIELD = "EQ" ' <C42
Private Const FIELD_C7_FIELD = "C7" ' <C43
Private Const FIELD_ER_FIELD = "ER" ' <C44
Private Const FIELD_C8_FIELD = "C8" ' <C45
Private Const FIELD_C9_FIELD = "C9" ' <C46
Private Const FIELD_CA_FIELD = "CA" ' <C47
Private Const FIELD_CB_FIELD = "CB" ' <C48
Private Const FIELD_ES_FIELD = "ES" ' <C49
Private Const FIELD_AK_FIELD = "AK" ' <C50
Private Const FIELD_AL_FIELD = "AL" ' <C51
Private Const FIELD_AM_FIELD = "AM" ' <C52
Private Const FIELD_BG_FIELD = "BG" ' <C53
Private Const FIELD_BH_FIELD = "BH" ' <C54
Private Const FIELD_BI_FIELD = "BI" ' <C55
Private Const FIELD_BJ_FIELD = "BJ" ' <C56
Private Const FIELD_BK_FIELD = "BK" ' <C57
Private Const FIELD_BL_FIELD = "BL" ' <C58
Private Const FIELD_SC_FIELD = "SC" ' <C59
Private Const FIELD_SD_FIELD = "SD" ' <C60
Private Const FIELD_SE_FIELD = "SE" ' <C61
Private Const FIELD_SF_FIELD = "SF" ' <C62
Private Const FIELD_SG_FIELD = "SG" ' <C63
Private Const FIELD_W8_FIELD = "W8" ' <C64
Private Const FIELD_W9_FIELD = "W9" ' <C65
Private Const FIELD_WA_FIELD = "WA" ' <C66
Private Const FIELD_WB_FIELD = "WB" ' <C67
Private Const FIELD_WC_FIELD = "WC" ' <C68
Private Const FIELD_WD_FIELD = "WD" ' <C69
Private Const FIELD_T7_FIELD = "T7" ' <C70
Private Const FIELD_WE_FIELD = "WE" ' <C71
Private Const FIELD_SB_FIELD = "SB" ' <C72
Private Const FIELD_DATE_LAST_RECEIVED_FIELD = "DATE LAST RECEIVED" ' <C73
Private Const FIELD_REMOTE_ID = "REMOTE_ID" '72
Private Const FIELD_ZA_FIELD = "ZA" ' <C73
Private Const FIELD_ZB_FIELD = "ZB" ' <C74

Private Const SQL_ADD_RECORD0 = "INSERT INTO [MASTEREDINCTS2] (" ' <C74

Private Const SQL_ADD_RECORD1 = "[CODE] " & _
                                "," & "[DTYPE] " & _
                                "," & "[DOCUMENT NAME] " & _
                                "," & "[TREE ID] " & _
                                "," & "[DATE CREATED] " & _
                                "," & "[DATE LAST MODIFIED] " & _
                                "," & "[DATE REQUESTED] " & _
                                "," & "[DATE SEND] " & _
                                "," & "[LAST MODIFIED BY] " & _
                                "," & "[LOGID DESCRIPTION] " & _
                                "," & "[REMARKS] " & _
                                "," & "[USER NO] " & _
                                "," & "[DOC NUMBER] " & _
                                "," & "[DOC TYPE] " & _
                                "," & "[Error String] " & _
                                "," & "[Error HD] " & _
                                "," & "[ORIGIN] " & _
                                "," & "[Type] " & _
                                "," & "[COMM] " & _
                                "," & "[LOGID] " & _
                                "," & "[PRINT] " & _
                                "," & "[VIEWED] " & _
                                "," & "[USERNAME] "
                                            
Private Const SQL_ADD_RECORD2 = SQL_ADD_RECORD1 & _
                                "," & "[HEADER] " & _
                                "," & "[Memo Field] " & _
                                "," & "[DATE PRINTED] " & _
                                "," & "[PRINTED BY] " & _
                                "," & "[MD] " & _
                                "," & "[ME] " & _
                                "," & "[MR] " & _
                                "," & "[AH] " & _
                                "," & "[AG] " & _
                                "," & "[BD] " & _
                                "," & "[BC] " & _
                                "," & "[AI] " & _
                                "," & "[MC] " & _
                                "," & "[EP] " & _
                                "," & "[AJ] " & _
                                "," & "[BF] " & _
                                "," & "[EQ] " & _
                                "," & "[C7] " & _
                                "," & "[ER] " & _
                                "," & "[C8] " & _
                                "," & "[C9] " & _
                                "," & "[CA] " & _
                                "," & "[CB] "

Private Const SQL_ADD_RECORD4 = "," & "[ZA] " & _
                                "," & "[ZB] "

Private Const SQL_ADD_RECORD3 = SQL_ADD_RECORD2 & _
                                "," & "[ES] " & _
                                "," & "[AK] " & _
                                "," & "[AL] " & _
                                "," & "[AM] " & _
                                "," & "[BG] " & _
                                "," & "[BH] " & _
                                "," & "[BI] " & _
                                "," & "[BJ] " & _
                                "," & "[BK] " & _
                                "," & "[BL] " & _
                                "," & "[SC] " & _
                                "," & "[SD] " & _
                                "," & "[SE] " & _
                                "," & "[SF] " & _
                                "," & "[SG] " & _
                                "," & "[W8] " & _
                                "," & "[W9] " & _
                                "," & "[WA] " & _
                                "," & "[WB] " & _
                                "," & "[WC] " & _
                                "," & "[WD] " & _
                                "," & "[T7] " & _
                                "," & "[WE] " & "," & "[SB] " & _
                                "," & "[DATE LAST RECEIVED] ,[REMOTE_ID] "

Private Const SQL_ADD_RECORD = SQL_ADD_RECORD0 & SQL_ADD_RECORD3 & SQL_ADD_RECORD4 & ") VALUES "

Private Const SQL_DELETE_RECORD = "DELETE * FROM [MASTEREDINCTS2] WHERE [CODE]="

Private Const SQL_WHERE_KEY = " WHERE [CODE]= "

Private Const SQL_GET_ALL_RECORD = "SELECT " & SQL_ADD_RECORD3 & _
                                   " FROM [MASTEREDINCTS2] "

Private Const SQL_GET_RECORD = SQL_GET_ALL_RECORD & SQL_WHERE_KEY

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM [MASTEREDINCTS2]"
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM [MASTEREDINCTS2]"

Dim mvarRecordset As ADODB.Recordset

' Code: >M1
' Name: AddRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "(" & SetEnclosedChr(ActiveRecord.CODE_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DTYPE_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DOCUMENT_NAME_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.TREE_ID_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DATE_CREATED_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DATE_LAST_MODIFIED_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DATE_REQUESTED_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DATE_SEND_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.LAST_MODIFIED_BY_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.LOGID_DESCRIPTION_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.REMARKS_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.USER_NO_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DOC_NUMBER_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.DOC_TYPE_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.Error_String_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.Error_HD_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.ORIGIN_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.Type_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.COMM_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.LOGID_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.Print_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.VIEWED_FIELD) & _
                 "," & SetEnclosedChr(ActiveRecord.USERNAME_FIELD)

        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.HEADER_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.Memo_Field_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DATE_PRINTED_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.PRINTED_BY_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.MD_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ME_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.MR_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AH_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AG_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BD_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BC_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AI_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.MC_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.EP_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AJ_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BF_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.EQ_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.C7_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ER_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.C8_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.C9_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.CA_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.CB_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ES_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AK_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AL_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.AM_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BG_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BH_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BI_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BJ_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BK_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.BL_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SC_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SD_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SF_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SG_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.W8_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.W9_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WA_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WB_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WC_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WD_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.T7_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SB_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DATE_LAST_RECEIVED_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ZA_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ZB_FIELD)
        strSql = strSql & ")"
        
                        
    
        On Error GoTo ERROR_QUERY
        
        ExecuteNonQuery ActiveConnection, strSql
        'ActiveConnection.Execute strSql
    
        AddRecord = True
        Exit Function
        
    End If
    
    AddRecord = False
    
    Exit Function
    
ERROR_QUERY:
    AddRecord = False
End Function

' Code: >M2
' Name: DeleteRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)
    
        On Error GoTo ERROR_QUERY
        ExecuteNonQuery ActiveConnection, strSql
        'ActiveConnection.Execute strSql
    
        DeleteRecord = True
        Exit Function
        
    End If
    
    DeleteRecord = False
    
    Exit Function
    
ERROR_QUERY:
    DeleteRecord = False
End Function

' Code: >M3
' Name: ModifyRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_DTYPE_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DTYPE_FIELD) & _
                 ",[" & FIELD_DOCUMENT_NAME_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DOCUMENT_NAME_FIELD) & "" & _
                 ",[" & FIELD_TREE_ID_FIELD & "]=" & SetEnclosedChr(ActiveRecord.TREE_ID_FIELD) & _
                 ",[" & FIELD_DATE_CREATED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_CREATED_FIELD) & _
                 ",[" & FIELD_DATE_LAST_MODIFIED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_LAST_MODIFIED_FIELD) & _
                 ",[" & FIELD_DATE_REQUESTED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_REQUESTED_FIELD) & _
                 ",[" & FIELD_DATE_SEND_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_SEND_FIELD) & _
                 ",[" & FIELD_LAST_MODIFIED_BY_FIELD & "]=" & SetEnclosedChr(ActiveRecord.LAST_MODIFIED_BY_FIELD) & _
                 ",[" & FIELD_LOGID_DESCRIPTION_FIELD & "]=" & SetEnclosedChr(ActiveRecord.LOGID_DESCRIPTION_FIELD) & _
                 ",[" & FIELD_REMARKS_FIELD & "]=" & SetEnclosedChr(ActiveRecord.REMARKS_FIELD) & _
                 ",[" & FIELD_USER_NO_FIELD & "]=" & SetEnclosedChr(ActiveRecord.USER_NO_FIELD) & _
                 ",[" & FIELD_DOC_NUMBER_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DOC_NUMBER_FIELD) & _
                 ",[" & FIELD_DOC_TYPE_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DOC_TYPE_FIELD) & _
                 ",[" & FIELD_Error_String_FIELD & "]=" & SetEnclosedChr(ActiveRecord.Error_String_FIELD) & _
                 ",[" & FIELD_Error_HD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.Error_HD_FIELD) & _
                 ",[" & FIELD_ORIGIN_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ORIGIN_FIELD) & _
                 ",[" & FIELD_Type_FIELD & "]=" & SetEnclosedChr(ActiveRecord.Type_FIELD) & _
                 ",[" & FIELD_COMM_FIELD & "]=" & SetEnclosedChr(ActiveRecord.COMM_FIELD) & _
                 ",[" & FIELD_LOGID_FIELD & "]=" & SetEnclosedChr(ActiveRecord.LOGID_FIELD) & _
                 ",[" & FIELD_PRINT_FIELD & "]=" & SetEnclosedChr(ActiveRecord.Print_FIELD) & _
                 ",[" & FIELD_VIEWED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.VIEWED_FIELD) & _
                 ",[" & FIELD_USERNAME_FIELD & "]=" & SetEnclosedChr(ActiveRecord.USERNAME_FIELD)
        
        strSql = strSql & ",[" & FIELD_HEADER_FIELD & "]=" & SetEnclosedChr(ActiveRecord.HEADER_FIELD)
        strSql = strSql & ",[" & FIELD_Memo_Field_FIELD & "]=" & SetEnclosedChr(ActiveRecord.Memo_Field_FIELD)
        strSql = strSql & ",[" & FIELD_DATE_PRINTED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_PRINTED_FIELD)
        strSql = strSql & ",[" & FIELD_PRINTED_BY_FIELD & "]=" & SetEnclosedChr(ActiveRecord.PRINTED_BY_FIELD)
        strSql = strSql & ",[" & FIELD_MD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.MD_FIELD)
        strSql = strSql & ",[" & FIELD_ME_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ME_FIELD)
        strSql = strSql & ",[" & FIELD_MR_FIELD & "]=" & SetEnclosedChr(ActiveRecord.MR_FIELD)
        strSql = strSql & ",[" & FIELD_AH_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AH_FIELD)
        strSql = strSql & ",[" & FIELD_AG_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AG_FIELD)
        strSql = strSql & ",[" & FIELD_BD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BD_FIELD)
        strSql = strSql & ",[" & FIELD_BC_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BC_FIELD)
        strSql = strSql & ",[" & FIELD_AI_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AI_FIELD)
        strSql = strSql & ",[" & FIELD_MC_FIELD & "]=" & SetEnclosedChr(ActiveRecord.MC_FIELD)
        strSql = strSql & ",[" & FIELD_EP_FIELD & "]=" & SetEnclosedChr(ActiveRecord.EP_FIELD)
        strSql = strSql & ",[" & FIELD_AJ_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AJ_FIELD)
        strSql = strSql & ",[" & FIELD_BF_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BF_FIELD)
        strSql = strSql & ",[" & FIELD_EQ_FIELD & "]=" & SetEnclosedChr(ActiveRecord.EQ_FIELD)
        strSql = strSql & ",[" & FIELD_C7_FIELD & "]=" & SetEnclosedChr(ActiveRecord.C7_FIELD)
        strSql = strSql & ",[" & FIELD_ER_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ER_FIELD)
        strSql = strSql & ",[" & FIELD_C8_FIELD & "]=" & SetEnclosedChr(ActiveRecord.C8_FIELD)
        strSql = strSql & ",[" & FIELD_C9_FIELD & "]=" & SetEnclosedChr(ActiveRecord.C9_FIELD)
        strSql = strSql & ",[" & FIELD_CA_FIELD & "]=" & SetEnclosedChr(ActiveRecord.CA_FIELD)
        strSql = strSql & ",[" & FIELD_CB_FIELD & "]=" & SetEnclosedChr(ActiveRecord.CB_FIELD)
        strSql = strSql & ",[" & FIELD_ES_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ES_FIELD)
        strSql = strSql & ",[" & FIELD_AK_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AK_FIELD)
        strSql = strSql & ",[" & FIELD_AL_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AL_FIELD)
        strSql = strSql & ",[" & FIELD_AM_FIELD & "]=" & SetEnclosedChr(ActiveRecord.AM_FIELD)
        strSql = strSql & ",[" & FIELD_BG_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BG_FIELD)
        strSql = strSql & ",[" & FIELD_BH_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BH_FIELD)
        strSql = strSql & ",[" & FIELD_BI_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BI_FIELD)
        strSql = strSql & ",[" & FIELD_BJ_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BJ_FIELD)
        strSql = strSql & ",[" & FIELD_BK_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BK_FIELD)
        strSql = strSql & ",[" & FIELD_BL_FIELD & "]=" & SetEnclosedChr(ActiveRecord.BL_FIELD)
        strSql = strSql & ",[" & FIELD_SC_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SC_FIELD)
        strSql = strSql & ",[" & FIELD_SD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SD_FIELD)
        strSql = strSql & ",[" & FIELD_SE_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SE_FIELD)
        strSql = strSql & ",[" & FIELD_SF_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SF_FIELD)
        strSql = strSql & ",[" & FIELD_SG_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SG_FIELD)
        strSql = strSql & ",[" & FIELD_W8_FIELD & "]=" & SetEnclosedChr(ActiveRecord.W8_FIELD)
        strSql = strSql & ",[" & FIELD_W9_FIELD & "]=" & SetEnclosedChr(ActiveRecord.W9_FIELD)
        strSql = strSql & ",[" & FIELD_WA_FIELD & "]=" & SetEnclosedChr(ActiveRecord.WA_FIELD)
        strSql = strSql & ",[" & FIELD_WB_FIELD & "]=" & SetEnclosedChr(ActiveRecord.WB_FIELD)
        strSql = strSql & ",[" & FIELD_WC_FIELD & "]=" & SetEnclosedChr(ActiveRecord.WC_FIELD)
        strSql = strSql & ",[" & FIELD_WD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.WD_FIELD)
        strSql = strSql & ",[" & FIELD_T7_FIELD & "]=" & SetEnclosedChr(ActiveRecord.T7_FIELD)
        strSql = strSql & ",[" & FIELD_WE_FIELD & "]=" & SetEnclosedChr(ActiveRecord.WE_FIELD)
        strSql = strSql & ",[" & FIELD_SB_FIELD & "]=" & SetEnclosedChr(ActiveRecord.SB_FIELD)
        strSql = strSql & ",[" & FIELD_DATE_LAST_RECEIVED_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_LAST_RECEIVED_FIELD)
        strSql = strSql & ", [REMOTE_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID) & ""
        strSql = strSql & ",[" & FIELD_ZA_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ZA_FIELD) & ""
        strSql = strSql & ",[" & FIELD_ZB_FIELD & "]=" & SetEnclosedChr(ActiveRecord.ZA_FIELD) & ""
        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD)
    
        On Error GoTo ERROR_QUERY
        ExecuteNonQuery ActiveConnection, strSql
        'ActiveConnection.Execute strSql
    
        ModifyRecord = True
        Exit Function
        
    End If
    
    ModifyRecord = False
    
    Exit Function
    
ERROR_QUERY:
    ModifyRecord = False
End Function

' Code: >M4
' Name: GetRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset
    
    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        
        On Error GoTo ERROR_QUERY
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        'Set rstRecord = ActiveConnection.Execute(strSql)
        
        On Error GoTo ERROR_RECORDSET
        ActiveRecord.DTYPE_FIELD = FNullField(rstRecord.Fields(FIELD_DTYPE_FIELD))
        ActiveRecord.DOCUMENT_NAME_FIELD = FNullField(rstRecord.Fields(FIELD_DOCUMENT_NAME_FIELD))
        ActiveRecord.TREE_ID_FIELD = FNullField(rstRecord.Fields(FIELD_TREE_ID_FIELD))
        ActiveRecord.DATE_CREATED_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_CREATED_FIELD))
        ActiveRecord.DATE_LAST_MODIFIED_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_LAST_MODIFIED_FIELD))
        ActiveRecord.DATE_REQUESTED_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_REQUESTED_FIELD))
        ActiveRecord.DATE_SEND_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_SEND_FIELD))
        ActiveRecord.LAST_MODIFIED_BY_FIELD = FNullField(rstRecord.Fields(FIELD_LAST_MODIFIED_BY_FIELD))
        ActiveRecord.LOGID_DESCRIPTION_FIELD = FNullField(rstRecord.Fields(FIELD_LOGID_DESCRIPTION_FIELD))
        ActiveRecord.REMARKS_FIELD = FNullField(rstRecord.Fields(FIELD_REMARKS_FIELD))
        ActiveRecord.USER_NO_FIELD = FNullField(rstRecord.Fields(FIELD_USER_NO_FIELD))
        ActiveRecord.DOC_NUMBER_FIELD = FNullField(rstRecord.Fields(FIELD_DOC_NUMBER_FIELD))
        ActiveRecord.DOC_TYPE_FIELD = FNullField(rstRecord.Fields(FIELD_DOC_TYPE_FIELD))
        ActiveRecord.Error_String_FIELD = FNullField(rstRecord.Fields(FIELD_Error_String_FIELD))
        ActiveRecord.Error_HD_FIELD = FNullField(rstRecord.Fields(FIELD_Error_HD_FIELD))
        ActiveRecord.ORIGIN_FIELD = FNullField(rstRecord.Fields(FIELD_ORIGIN_FIELD))
        ActiveRecord.Type_FIELD = FNullField(rstRecord.Fields(FIELD_Type_FIELD))
        ActiveRecord.COMM_FIELD = FNullField(rstRecord.Fields(FIELD_COMM_FIELD))
        ActiveRecord.LOGID_FIELD = FNullField(rstRecord.Fields(FIELD_LOGID_FIELD))
        ActiveRecord.Print_FIELD = FNullField(rstRecord.Fields(FIELD_PRINT_FIELD))
        ActiveRecord.VIEWED_FIELD = FNullField(rstRecord.Fields(FIELD_VIEWED_FIELD))
        ActiveRecord.USERNAME_FIELD = FNullField(rstRecord.Fields(FIELD_USERNAME_FIELD))
        
        ActiveRecord.HEADER_FIELD = FNullField(rstRecord.Fields(FIELD_HEADER_FIELD))
        ActiveRecord.Memo_Field_FIELD = FNullField(rstRecord.Fields(FIELD_Memo_Field_FIELD))
        ActiveRecord.DATE_PRINTED_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_PRINTED_FIELD))
        ActiveRecord.PRINTED_BY_FIELD = FNullField(rstRecord.Fields(FIELD_PRINTED_BY_FIELD))
        ActiveRecord.MD_FIELD = FNullField(rstRecord.Fields(FIELD_MD_FIELD))
        ActiveRecord.ME_FIELD = FNullField(rstRecord.Fields(FIELD_ME_FIELD))
        ActiveRecord.MR_FIELD = FNullField(rstRecord.Fields(FIELD_MR_FIELD))
        ActiveRecord.AH_FIELD = FNullField(rstRecord.Fields(FIELD_AH_FIELD))
        ActiveRecord.AG_FIELD = FNullField(rstRecord.Fields(FIELD_AG_FIELD))
        ActiveRecord.BD_FIELD = FNullField(rstRecord.Fields(FIELD_BD_FIELD))
        ActiveRecord.BC_FIELD = FNullField(rstRecord.Fields(FIELD_BC_FIELD))
        ActiveRecord.AI_FIELD = FNullField(rstRecord.Fields(FIELD_AI_FIELD))
        ActiveRecord.MC_FIELD = FNullField(rstRecord.Fields(FIELD_MC_FIELD))
        ActiveRecord.EP_FIELD = FNullField(rstRecord.Fields(FIELD_EP_FIELD))
        ActiveRecord.AJ_FIELD = FNullField(rstRecord.Fields(FIELD_AJ_FIELD))
        ActiveRecord.BF_FIELD = FNullField(rstRecord.Fields(FIELD_BF_FIELD))
        ActiveRecord.EQ_FIELD = FNullField(rstRecord.Fields(FIELD_EQ_FIELD))
        ActiveRecord.C7_FIELD = FNullField(rstRecord.Fields(FIELD_C7_FIELD))
        ActiveRecord.ER_FIELD = FNullField(rstRecord.Fields(FIELD_ER_FIELD))
        ActiveRecord.C8_FIELD = FNullField(rstRecord.Fields(FIELD_C8_FIELD))
        ActiveRecord.C9_FIELD = FNullField(rstRecord.Fields(FIELD_C9_FIELD))
        ActiveRecord.CA_FIELD = FNullField(rstRecord.Fields(FIELD_CA_FIELD))
        ActiveRecord.CB_FIELD = FNullField(rstRecord.Fields(FIELD_CB_FIELD))
        ActiveRecord.ES_FIELD = FNullField(rstRecord.Fields(FIELD_ES_FIELD))
        ActiveRecord.AK_FIELD = FNullField(rstRecord.Fields(FIELD_AK_FIELD))
        ActiveRecord.AL_FIELD = FNullField(rstRecord.Fields(FIELD_AL_FIELD))
        ActiveRecord.AM_FIELD = FNullField(rstRecord.Fields(FIELD_AM_FIELD))
        ActiveRecord.BG_FIELD = FNullField(rstRecord.Fields(FIELD_BG_FIELD))
        ActiveRecord.BH_FIELD = FNullField(rstRecord.Fields(FIELD_BH_FIELD))
        ActiveRecord.BI_FIELD = FNullField(rstRecord.Fields(FIELD_BI_FIELD))
        ActiveRecord.BJ_FIELD = FNullField(rstRecord.Fields(FIELD_BJ_FIELD))
        ActiveRecord.BK_FIELD = FNullField(rstRecord.Fields(FIELD_BK_FIELD))
        ActiveRecord.BL_FIELD = FNullField(rstRecord.Fields(FIELD_BL_FIELD))
        ActiveRecord.SC_FIELD = FNullField(rstRecord.Fields(FIELD_SC_FIELD))
        ActiveRecord.SD_FIELD = FNullField(rstRecord.Fields(FIELD_SD_FIELD))
        ActiveRecord.SE_FIELD = FNullField(rstRecord.Fields(FIELD_SE_FIELD))
        ActiveRecord.SF_FIELD = FNullField(rstRecord.Fields(FIELD_SF_FIELD))
        ActiveRecord.SG_FIELD = FNullField(rstRecord.Fields(FIELD_SG_FIELD))
        ActiveRecord.W8_FIELD = FNullField(rstRecord.Fields(FIELD_W8_FIELD))
        ActiveRecord.W9_FIELD = FNullField(rstRecord.Fields(FIELD_W9_FIELD))
        ActiveRecord.WA_FIELD = FNullField(rstRecord.Fields(FIELD_WA_FIELD))
        ActiveRecord.WB_FIELD = FNullField(rstRecord.Fields(FIELD_WB_FIELD))
        ActiveRecord.WC_FIELD = FNullField(rstRecord.Fields(FIELD_WC_FIELD))
        ActiveRecord.WD_FIELD = FNullField(rstRecord.Fields(FIELD_WD_FIELD))
        ActiveRecord.T7_FIELD = FNullField(rstRecord.Fields(FIELD_T7_FIELD))
        ActiveRecord.WE_FIELD = FNullField(rstRecord.Fields(FIELD_WE_FIELD))
        ActiveRecord.SB_FIELD = FNullField(rstRecord.Fields(FIELD_SB_FIELD))
        ActiveRecord.DATE_LAST_RECEIVED_FIELD = FNullField(rstRecord.Fields(FIELD_DATE_LAST_RECEIVED_FIELD))
        ActiveRecord.FIELD_REMOTE_ID = FNullField(rstRecord.Fields("REMOTE_ID"))
        ActiveRecord.ZA_FIELD = FNullField(rstRecord.Fields(FIELD_ZA_FIELD))
        ActiveRecord.ZB_FIELD = FNullField(rstRecord.Fields(FIELD_ZB_FIELD))
        
        ADORecordsetClose rstRecord
        'Set rstRecord = Nothing
        
        GetRecord = True
    
    ElseIf ((ActiveRecord Is Nothing) = True) Then
        GetRecord = False
    End If
    
    Exit Function
    
ERROR_RECORDSET:
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetRecord = False
    Exit Function
ERROR_QUERY:
    GetRecord = False
End Function

' Code: >M5
' Name: GetMaxID
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiMasterEdiNcts2
    End If

    strSql = SQL_GET_MAXID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.CODE_FIELD = rstRecord.Fields("ID_MAX").Value
    
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetMaxID = True
    Exit Function
    
ERROR_RECORDSET:
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetMaxID = False
    Exit Function
ERROR_QUERY:
    GetMaxID = False
End Function

' Code: >M6
' Name: GetMinID
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMasterEdiNcts2) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiMasterEdiNcts2
    End If

    strSql = SQL_GET_MINID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.CODE_FIELD = rstRecord.Fields("ID_MIN").Value
    
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetMinID = True
    Exit Function
    
ERROR_RECORDSET:
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetMinID = False
    Exit Function
ERROR_QUERY:
    GetMinID = False
End Function

' Code: >M7
' Name: GetTableRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetTableRecord(ByRef ActiveRecord As cpiMasterEdiNcts2) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset
    
    Set rstRecord = New ADODB.Recordset
    
    rstRecord.Open
    
    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 30, 102 ' 1
    rstRecord.Fields.Append FIELD_DTYPE_FIELD, 17, 1, 118 ' 2
    rstRecord.Fields.Append FIELD_DOCUMENT_NAME_FIELD, 202, 40, 102 ' 3
    rstRecord.Fields.Append FIELD_TREE_ID_FIELD, 202, 10, 70 ' 4
    rstRecord.Fields.Append FIELD_DATE_CREATED_FIELD, 7, 8, 118 ' 5
    rstRecord.Fields.Append FIELD_DATE_LAST_MODIFIED_FIELD, 7, 8, 118 ' 6
    rstRecord.Fields.Append FIELD_DATE_REQUESTED_FIELD, 7, 8, 118 ' 7
    rstRecord.Fields.Append FIELD_DATE_SEND_FIELD, 7, 8, 118 ' 8
    rstRecord.Fields.Append FIELD_LAST_MODIFIED_BY_FIELD, 202, 25, 102 ' 9
    rstRecord.Fields.Append FIELD_LOGID_DESCRIPTION_FIELD, 202, 40, 102 ' 10
    rstRecord.Fields.Append FIELD_REMARKS_FIELD, 202, 50, 102 ' 11
    rstRecord.Fields.Append FIELD_USER_NO_FIELD, 3, 4, 118 ' 12
    rstRecord.Fields.Append FIELD_DOC_NUMBER_FIELD, 202, 7, 102 ' 13
    rstRecord.Fields.Append FIELD_DOC_TYPE_FIELD, 202, 3, 102 ' 14
    rstRecord.Fields.Append FIELD_Error_String_FIELD, 203, 536870910, 230 ' 15
    rstRecord.Fields.Append FIELD_Error_HD_FIELD, 202, 10, 102 ' 16
    rstRecord.Fields.Append FIELD_ORIGIN_FIELD, 202, 10, 102 ' 17
    rstRecord.Fields.Append FIELD_Type_FIELD, 202, 1, 102 ' 18
    rstRecord.Fields.Append FIELD_COMM_FIELD, 202, 1, 102 ' 19
    rstRecord.Fields.Append FIELD_LOGID_FIELD, 202, 4, 102 ' 20
    rstRecord.Fields.Append FIELD_PRINT_FIELD, 202, 1, 102 ' 21
    rstRecord.Fields.Append FIELD_VIEWED_FIELD, 17, 1, 118 ' 22
    rstRecord.Fields.Append FIELD_USERNAME_FIELD, 202, 25, 102 ' 23
    rstRecord.Fields.Append FIELD_HEADER_FIELD, 2, 2, 118 ' 24
    rstRecord.Fields.Append FIELD_Memo_Field_FIELD, 203, 536870910, 230 ' 25
    rstRecord.Fields.Append FIELD_DATE_PRINTED_FIELD, 7, 8, 118 ' 26
    rstRecord.Fields.Append FIELD_PRINTED_BY_FIELD, 202, 25, 102 ' 27
    rstRecord.Fields.Append FIELD_MD_FIELD, 202, 6, 102 ' 28
    rstRecord.Fields.Append FIELD_ME_FIELD, 202, 4, 102 ' 29
    rstRecord.Fields.Append FIELD_MR_FIELD, 202, 21, 102 ' 30
    rstRecord.Fields.Append FIELD_AH_FIELD, 202, 17, 102 ' 31
    rstRecord.Fields.Append FIELD_AG_FIELD, 202, 35, 102 ' 32
    rstRecord.Fields.Append FIELD_BD_FIELD, 202, 8, 102 ' 33
    rstRecord.Fields.Append FIELD_BC_FIELD, 202, 35, 102 ' 34
    rstRecord.Fields.Append FIELD_AI_FIELD, 202, 17, 102 ' 35
    rstRecord.Fields.Append FIELD_MC_FIELD, 202, 8, 102 ' 36
    rstRecord.Fields.Append FIELD_EP_FIELD, 202, 1, 102 ' 37
    rstRecord.Fields.Append FIELD_AJ_FIELD, 202, 2, 102 ' 38
    rstRecord.Fields.Append FIELD_BF_FIELD, 202, 35, 102 ' 39
    rstRecord.Fields.Append FIELD_EQ_FIELD, 202, 27, 102 ' 40
    rstRecord.Fields.Append FIELD_C7_FIELD, 202, 2, 102 ' 41
    rstRecord.Fields.Append FIELD_ER_FIELD, 202, 1, 102 ' 42
    rstRecord.Fields.Append FIELD_C8_FIELD, 202, 8, 102 ' 43
    rstRecord.Fields.Append FIELD_C9_FIELD, 202, 35, 102 ' 44
    rstRecord.Fields.Append FIELD_CA_FIELD, 202, 35, 102 ' 45
    rstRecord.Fields.Append FIELD_CB_FIELD, 202, 2, 102 ' 46
    rstRecord.Fields.Append FIELD_ES_FIELD, 203, 536870910, 230 ' 47
    rstRecord.Fields.Append FIELD_AK_FIELD, 202, 4, 102 ' 48
    rstRecord.Fields.Append FIELD_AL_FIELD, 202, 20, 102 ' 49
    rstRecord.Fields.Append FIELD_AM_FIELD, 202, 20, 102 ' 50
    rstRecord.Fields.Append FIELD_BG_FIELD, 202, 2, 102 ' 51
    rstRecord.Fields.Append FIELD_BH_FIELD, 202, 8, 102 ' 52
    rstRecord.Fields.Append FIELD_BI_FIELD, 202, 35, 102 ' 53
    rstRecord.Fields.Append FIELD_BJ_FIELD, 202, 35, 102 ' 54
    rstRecord.Fields.Append FIELD_BK_FIELD, 202, 2, 102 ' 55
    rstRecord.Fields.Append FIELD_BL_FIELD, 202, 27, 102 ' 56
    rstRecord.Fields.Append FIELD_SC_FIELD, 202, 17, 102 ' 57
    rstRecord.Fields.Append FIELD_SD_FIELD, 202, 17, 102 ' 58
    rstRecord.Fields.Append FIELD_SE_FIELD, 202, 17, 102 ' 59
    rstRecord.Fields.Append FIELD_SF_FIELD, 202, 17, 102 ' 60
    rstRecord.Fields.Append FIELD_SG_FIELD, 202, 17, 102 ' 61
    rstRecord.Fields.Append FIELD_W8_FIELD, 202, 17, 102 ' 62
    rstRecord.Fields.Append FIELD_W9_FIELD, 202, 35, 102 ' 63
    rstRecord.Fields.Append FIELD_WA_FIELD, 202, 35, 102 ' 64
    rstRecord.Fields.Append FIELD_WB_FIELD, 202, 35, 102 ' 65
    rstRecord.Fields.Append FIELD_WC_FIELD, 202, 2, 102 ' 66
    rstRecord.Fields.Append FIELD_WD_FIELD, 202, 9, 102 ' 67
    rstRecord.Fields.Append FIELD_T7_FIELD, 202, 1, 102 ' 68
    rstRecord.Fields.Append FIELD_WE_FIELD, 202, 2, 102 ' 69
    rstRecord.Fields.Append FIELD_SB_FIELD, 202, 1, 102 ' 70
    rstRecord.Fields.Append FIELD_DATE_LAST_RECEIVED_FIELD, 7, 8, 118 ' 71
    rstRecord.Fields.Append FIELD_REMOTE_ID, 3, 4, 118 ' 72
    rstRecord.Fields.Append FIELD_ZA_FIELD, 202, 2, 102 ' 73
    rstRecord.Fields.Append FIELD_ZB_FIELD, 202, 2, 102 ' 74

    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_DTYPE_FIELD) = ActiveRecord.DTYPE_FIELD
    rstRecord.Fields(FIELD_DOCUMENT_NAME_FIELD) = ActiveRecord.DOCUMENT_NAME_FIELD
    rstRecord.Fields(FIELD_TREE_ID_FIELD) = ActiveRecord.TREE_ID_FIELD
    rstRecord.Fields(FIELD_DATE_CREATED_FIELD) = ActiveRecord.DATE_CREATED_FIELD
    rstRecord.Fields(FIELD_DATE_LAST_MODIFIED_FIELD) = ActiveRecord.DATE_LAST_MODIFIED_FIELD
    rstRecord.Fields(FIELD_DATE_REQUESTED_FIELD) = ActiveRecord.DATE_REQUESTED_FIELD
    rstRecord.Fields(FIELD_DATE_SEND_FIELD) = ActiveRecord.DATE_SEND_FIELD
    rstRecord.Fields(FIELD_LAST_MODIFIED_BY_FIELD) = ActiveRecord.LAST_MODIFIED_BY_FIELD
    rstRecord.Fields(FIELD_LOGID_DESCRIPTION_FIELD) = ActiveRecord.LOGID_DESCRIPTION_FIELD
    rstRecord.Fields(FIELD_REMARKS_FIELD) = ActiveRecord.REMARKS_FIELD
    rstRecord.Fields(FIELD_USER_NO_FIELD) = ActiveRecord.USER_NO_FIELD
    rstRecord.Fields(FIELD_DOC_NUMBER_FIELD) = ActiveRecord.DOC_NUMBER_FIELD
    rstRecord.Fields(FIELD_DOC_TYPE_FIELD) = ActiveRecord.DOC_TYPE_FIELD
    rstRecord.Fields(FIELD_Error_String_FIELD) = ActiveRecord.Error_String_FIELD
    rstRecord.Fields(FIELD_Error_HD_FIELD) = ActiveRecord.Error_HD_FIELD
    rstRecord.Fields(FIELD_ORIGIN_FIELD) = ActiveRecord.ORIGIN_FIELD
    rstRecord.Fields(FIELD_Type_FIELD) = ActiveRecord.Type_FIELD
    rstRecord.Fields(FIELD_COMM_FIELD) = ActiveRecord.COMM_FIELD
    rstRecord.Fields(FIELD_LOGID_FIELD) = ActiveRecord.LOGID_FIELD
    rstRecord.Fields(FIELD_PRINT_FIELD) = ActiveRecord.Print_FIELD
    rstRecord.Fields(FIELD_VIEWED_FIELD) = ActiveRecord.VIEWED_FIELD
    rstRecord.Fields(FIELD_USERNAME_FIELD) = ActiveRecord.USERNAME_FIELD
    
    rstRecord.Fields(FIELD_HEADER_FIELD) = ActiveRecord.HEADER_FIELD
    rstRecord.Fields(FIELD_Memo_Field_FIELD) = ActiveRecord.Memo_Field_FIELD
    rstRecord.Fields(FIELD_DATE_PRINTED_FIELD) = ActiveRecord.DATE_PRINTED_FIELD
    rstRecord.Fields(FIELD_PRINTED_BY_FIELD) = ActiveRecord.PRINTED_BY_FIELD
    rstRecord.Fields(FIELD_MD_FIELD) = ActiveRecord.MD_FIELD
    rstRecord.Fields(FIELD_ME_FIELD) = ActiveRecord.ME_FIELD
    rstRecord.Fields(FIELD_MR_FIELD) = ActiveRecord.MR_FIELD
    rstRecord.Fields(FIELD_AH_FIELD) = ActiveRecord.AH_FIELD
    rstRecord.Fields(FIELD_AG_FIELD) = ActiveRecord.AG_FIELD
    rstRecord.Fields(FIELD_BD_FIELD) = ActiveRecord.BD_FIELD
    rstRecord.Fields(FIELD_BC_FIELD) = ActiveRecord.BC_FIELD
    rstRecord.Fields(FIELD_AI_FIELD) = ActiveRecord.AI_FIELD
    rstRecord.Fields(FIELD_MC_FIELD) = ActiveRecord.MC_FIELD
    rstRecord.Fields(FIELD_EP_FIELD) = ActiveRecord.EP_FIELD
    rstRecord.Fields(FIELD_AJ_FIELD) = ActiveRecord.AJ_FIELD
    rstRecord.Fields(FIELD_BF_FIELD) = ActiveRecord.BF_FIELD
    rstRecord.Fields(FIELD_EQ_FIELD) = ActiveRecord.EQ_FIELD
    rstRecord.Fields(FIELD_C7_FIELD) = ActiveRecord.C7_FIELD
    rstRecord.Fields(FIELD_ER_FIELD) = ActiveRecord.ER_FIELD
    rstRecord.Fields(FIELD_C8_FIELD) = ActiveRecord.C8_FIELD
    rstRecord.Fields(FIELD_C9_FIELD) = ActiveRecord.C9_FIELD
    rstRecord.Fields(FIELD_CA_FIELD) = ActiveRecord.CA_FIELD
    rstRecord.Fields(FIELD_CB_FIELD) = ActiveRecord.CB_FIELD
    rstRecord.Fields(FIELD_ES_FIELD) = ActiveRecord.ES_FIELD
    rstRecord.Fields(FIELD_AK_FIELD) = ActiveRecord.AK_FIELD
    rstRecord.Fields(FIELD_AL_FIELD) = ActiveRecord.AL_FIELD
    rstRecord.Fields(FIELD_AM_FIELD) = ActiveRecord.AM_FIELD
    rstRecord.Fields(FIELD_BG_FIELD) = ActiveRecord.BG_FIELD
    rstRecord.Fields(FIELD_BH_FIELD) = ActiveRecord.BH_FIELD
    rstRecord.Fields(FIELD_BI_FIELD) = ActiveRecord.BI_FIELD
    rstRecord.Fields(FIELD_BJ_FIELD) = ActiveRecord.BJ_FIELD
    rstRecord.Fields(FIELD_BK_FIELD) = ActiveRecord.BK_FIELD
    rstRecord.Fields(FIELD_BL_FIELD) = ActiveRecord.BL_FIELD
    rstRecord.Fields(FIELD_SC_FIELD) = ActiveRecord.SC_FIELD
    rstRecord.Fields(FIELD_SD_FIELD) = ActiveRecord.SD_FIELD
    rstRecord.Fields(FIELD_SE_FIELD) = ActiveRecord.SE_FIELD
    rstRecord.Fields(FIELD_SF_FIELD) = ActiveRecord.SF_FIELD
    rstRecord.Fields(FIELD_SG_FIELD) = ActiveRecord.SG_FIELD
    rstRecord.Fields(FIELD_W8_FIELD) = ActiveRecord.W8_FIELD
    rstRecord.Fields(FIELD_W9_FIELD) = ActiveRecord.W9_FIELD
    rstRecord.Fields(FIELD_WA_FIELD) = ActiveRecord.WA_FIELD
    rstRecord.Fields(FIELD_WB_FIELD) = ActiveRecord.WB_FIELD
    rstRecord.Fields(FIELD_WC_FIELD) = ActiveRecord.WC_FIELD
    rstRecord.Fields(FIELD_WD_FIELD) = ActiveRecord.WD_FIELD
    rstRecord.Fields(FIELD_T7_FIELD) = ActiveRecord.T7_FIELD
    rstRecord.Fields(FIELD_WE_FIELD) = ActiveRecord.WE_FIELD
    rstRecord.Fields(FIELD_SB_FIELD) = ActiveRecord.SB_FIELD
    rstRecord.Fields(FIELD_DATE_LAST_RECEIVED_FIELD) = ActiveRecord.DATE_LAST_RECEIVED_FIELD
    rstRecord.Fields("REMOTE_ID") = ActiveRecord.FIELD_REMOTE_ID
    rstRecord.Fields(FIELD_ZA_FIELD) = ActiveRecord.ZA_FIELD
    rstRecord.Fields(FIELD_ZB_FIELD) = ActiveRecord.ZB_FIELD
    
    rstRecord.Update
    
    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

' Code: >M8
' Name: GetClassRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiMasterEdiNcts2

    Dim clsRecord As cpiMasterEdiNcts2
    
    Set clsRecord = New cpiMasterEdiNcts2
    
    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.DTYPE_FIELD = FNullField(ActiveRecord.Fields(FIELD_DTYPE_FIELD))
    clsRecord.DOCUMENT_NAME_FIELD = FNullField(ActiveRecord.Fields(FIELD_DOCUMENT_NAME_FIELD))
    clsRecord.TREE_ID_FIELD = FNullField(ActiveRecord.Fields(FIELD_TREE_ID_FIELD))
    clsRecord.DATE_CREATED_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_CREATED_FIELD))
    clsRecord.DATE_LAST_MODIFIED_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_LAST_MODIFIED_FIELD))
    clsRecord.DATE_REQUESTED_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_REQUESTED_FIELD))
    clsRecord.DATE_SEND_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_SEND_FIELD))
    clsRecord.LAST_MODIFIED_BY_FIELD = FNullField(ActiveRecord.Fields(FIELD_LAST_MODIFIED_BY_FIELD))
    clsRecord.LOGID_DESCRIPTION_FIELD = FNullField(ActiveRecord.Fields(FIELD_LOGID_DESCRIPTION_FIELD))
    clsRecord.REMARKS_FIELD = FNullField(ActiveRecord.Fields(FIELD_REMARKS_FIELD))
    clsRecord.USER_NO_FIELD = FNullField(ActiveRecord.Fields(FIELD_USER_NO_FIELD))
    clsRecord.DOC_NUMBER_FIELD = FNullField(ActiveRecord.Fields(FIELD_DOC_NUMBER_FIELD))
    clsRecord.DOC_TYPE_FIELD = FNullField(ActiveRecord.Fields(FIELD_DOC_TYPE_FIELD))
    clsRecord.Error_String_FIELD = FNullField(ActiveRecord.Fields(FIELD_Error_String_FIELD))
    clsRecord.Error_HD_FIELD = FNullField(ActiveRecord.Fields(FIELD_Error_HD_FIELD))
    clsRecord.ORIGIN_FIELD = FNullField(ActiveRecord.Fields(FIELD_ORIGIN_FIELD))
    clsRecord.Type_FIELD = FNullField(ActiveRecord.Fields(FIELD_Type_FIELD))
    clsRecord.COMM_FIELD = FNullField(ActiveRecord.Fields(FIELD_COMM_FIELD))
    clsRecord.LOGID_FIELD = FNullField(ActiveRecord.Fields(FIELD_LOGID_FIELD))
    clsRecord.Print_FIELD = FNullField(ActiveRecord.Fields(FIELD_PRINT_FIELD))
    clsRecord.VIEWED_FIELD = FNullField(ActiveRecord.Fields(FIELD_VIEWED_FIELD))
    clsRecord.USERNAME_FIELD = FNullField(ActiveRecord.Fields(FIELD_USERNAME_FIELD))
    
    clsRecord.HEADER_FIELD = FNullField(ActiveRecord.Fields(FIELD_HEADER_FIELD))
    clsRecord.Memo_Field_FIELD = FNullField(ActiveRecord.Fields(FIELD_Memo_Field_FIELD))
    clsRecord.DATE_PRINTED_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_PRINTED_FIELD))
    clsRecord.PRINTED_BY_FIELD = FNullField(ActiveRecord.Fields(FIELD_PRINTED_BY_FIELD))
    clsRecord.MD_FIELD = FNullField(ActiveRecord.Fields(FIELD_MD_FIELD))
    clsRecord.ME_FIELD = FNullField(ActiveRecord.Fields(FIELD_ME_FIELD))
    clsRecord.MR_FIELD = FNullField(ActiveRecord.Fields(FIELD_MR_FIELD))
    clsRecord.AH_FIELD = FNullField(ActiveRecord.Fields(FIELD_AH_FIELD))
    clsRecord.AG_FIELD = FNullField(ActiveRecord.Fields(FIELD_AG_FIELD))
    clsRecord.BD_FIELD = FNullField(ActiveRecord.Fields(FIELD_BD_FIELD))
    clsRecord.BC_FIELD = FNullField(ActiveRecord.Fields(FIELD_BC_FIELD))
    clsRecord.AI_FIELD = FNullField(ActiveRecord.Fields(FIELD_AI_FIELD))
    clsRecord.MC_FIELD = FNullField(ActiveRecord.Fields(FIELD_MC_FIELD))
    clsRecord.EP_FIELD = FNullField(ActiveRecord.Fields(FIELD_EP_FIELD))
    clsRecord.AJ_FIELD = FNullField(ActiveRecord.Fields(FIELD_AJ_FIELD))
    clsRecord.BF_FIELD = FNullField(ActiveRecord.Fields(FIELD_BF_FIELD))
    clsRecord.EQ_FIELD = FNullField(ActiveRecord.Fields(FIELD_EQ_FIELD))
    clsRecord.C7_FIELD = FNullField(ActiveRecord.Fields(FIELD_C7_FIELD))
    clsRecord.ER_FIELD = FNullField(ActiveRecord.Fields(FIELD_ER_FIELD))
    clsRecord.C8_FIELD = FNullField(ActiveRecord.Fields(FIELD_C8_FIELD))
    clsRecord.C9_FIELD = FNullField(ActiveRecord.Fields(FIELD_C9_FIELD))
    clsRecord.CA_FIELD = FNullField(ActiveRecord.Fields(FIELD_CA_FIELD))
    clsRecord.CB_FIELD = FNullField(ActiveRecord.Fields(FIELD_CB_FIELD))
    clsRecord.ES_FIELD = FNullField(ActiveRecord.Fields(FIELD_ES_FIELD))
    clsRecord.AK_FIELD = FNullField(ActiveRecord.Fields(FIELD_AK_FIELD))
    clsRecord.AL_FIELD = FNullField(ActiveRecord.Fields(FIELD_AL_FIELD))
    clsRecord.AM_FIELD = FNullField(ActiveRecord.Fields(FIELD_AM_FIELD))
    clsRecord.BG_FIELD = FNullField(ActiveRecord.Fields(FIELD_BG_FIELD))
    clsRecord.BH_FIELD = FNullField(ActiveRecord.Fields(FIELD_BH_FIELD))
    clsRecord.BI_FIELD = FNullField(ActiveRecord.Fields(FIELD_BI_FIELD))
    clsRecord.BJ_FIELD = FNullField(ActiveRecord.Fields(FIELD_BJ_FIELD))
    clsRecord.BK_FIELD = FNullField(ActiveRecord.Fields(FIELD_BK_FIELD))
    clsRecord.BL_FIELD = FNullField(ActiveRecord.Fields(FIELD_BL_FIELD))
    clsRecord.SC_FIELD = FNullField(ActiveRecord.Fields(FIELD_SC_FIELD))
    clsRecord.SD_FIELD = FNullField(ActiveRecord.Fields(FIELD_SD_FIELD))
    clsRecord.SE_FIELD = FNullField(ActiveRecord.Fields(FIELD_SE_FIELD))
    clsRecord.SF_FIELD = FNullField(ActiveRecord.Fields(FIELD_SF_FIELD))
    clsRecord.SG_FIELD = FNullField(ActiveRecord.Fields(FIELD_SG_FIELD))
    clsRecord.W8_FIELD = FNullField(ActiveRecord.Fields(FIELD_W8_FIELD))
    clsRecord.W9_FIELD = FNullField(ActiveRecord.Fields(FIELD_W9_FIELD))
    clsRecord.WA_FIELD = FNullField(ActiveRecord.Fields(FIELD_WA_FIELD))
    clsRecord.WB_FIELD = FNullField(ActiveRecord.Fields(FIELD_WB_FIELD))
    clsRecord.WC_FIELD = FNullField(ActiveRecord.Fields(FIELD_WC_FIELD))
    clsRecord.WD_FIELD = FNullField(ActiveRecord.Fields(FIELD_WD_FIELD))
    clsRecord.T7_FIELD = FNullField(ActiveRecord.Fields(FIELD_T7_FIELD))
    clsRecord.WE_FIELD = FNullField(ActiveRecord.Fields(FIELD_WE_FIELD))
    clsRecord.SB_FIELD = FNullField(ActiveRecord.Fields(FIELD_SB_FIELD))
    clsRecord.DATE_LAST_RECEIVED_FIELD = FNullField(ActiveRecord.Fields(FIELD_DATE_LAST_RECEIVED_FIELD))
    clsRecord.FIELD_REMOTE_ID = FNullField(ActiveRecord.Fields(FIELD_REMOTE_ID))
    clsRecord.ZA_FIELD = FNullField(ActiveRecord.Fields(FIELD_ZA_FIELD))
    clsRecord.ZB_FIELD = FNullField(ActiveRecord.Fields(FIELD_ZB_FIELD))
    
    Set GetClassRecord = clsRecord
    Set clsRecord = Nothing

End Function

' Code: >M9
' Name: SearchRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, ByVal SearchField, ByVal SearchValue) As Boolean
    
    Dim strSql As String
    Dim lngRecordsAffected As Long
    Dim rstDB As ADODB.Recordset
    
    On Error GoTo ERROR_SEARCH
    SearchField = Trim$(SearchField)
    If ((SearchField <> "") And (SearchValue <> "")) Then
        If (Len(SearchField) > 2) Then
            If ((Left$(SearchField, 1) <> "[") And (Right$(SearchField, 1) <> "]")) Then
                SearchField = "[" & SearchField & "]"
            End If
        End If
    
        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TABLE_NAME & " WHERE " & SearchField & "=" & SetEnclosedChr((SearchValue))
        ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
        'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
        
        SearchRecord = Not (rstDB.EOF = True And rstDB.EOF = True)
        
        'If (rstDB.EOF = False) Then
        '    SearchRecord = True
        'ElseIf (rstDB.EOF = True) Then
        '    SearchRecord = False
        'End If
        
    End If
    
    Exit Function
    
ERROR_SEARCH:
    SearchValue = False
End Function

' Code: >M10
' Name: GetRecordset
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function GetRecordset(ByRef ActiveConnection As ADODB.Connection, Optional ByVal CommandText As String) As ADODB.Recordset

    Dim rstRecordset As ADODB.Recordset

    Set rstRecordset = New ADODB.Recordset

    On Error GoTo ERROR_RECORDSET
    
    If (CommandText = "") Then
        CommandText = SQL_GET_ALL_RECORD
    End If
    
    ADORecordsetOpen CommandText, rstRecordset, ActiveConnection, adOpenKeyset, adLockOptimistic
    'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

    Set GetRecordset = rstRecordset
    Set mvarRecordset = rstRecordset
    Set rstRecordset = Nothing

    Exit Function

ERROR_RECORDSET:
    Set rstRecordset = Nothing
End Function

' Code: >P1
' Name: Recordset (Get)
' Description: return Recordset
' Scope: Public
' Fan-In: <none>
' Data Type: ADODB.Recordset
Public Property Get Recordset() As ADODB.Recordset
Attribute Recordset.VB_UserMemId = 0
    Set Recordset = mvarRecordset
End Property

' Code: <M11
' Name: Class_Terminate
' Description: remove objects in memory
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Private Sub Class_Terminate()
    Set mvarRecordset = Nothing
End Sub
