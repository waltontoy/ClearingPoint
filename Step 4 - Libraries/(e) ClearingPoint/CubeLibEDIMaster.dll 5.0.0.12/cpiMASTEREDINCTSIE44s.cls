VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiMASTEREDINCTSIE44s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiMASTEREDINCTSIE44s
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: Collection of cpiMASTEREDINCTSIE44 class
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

Private mCol As Collection ' <V1
Private mvarRecordset As ADODB.Recordset ' <V2

Private Const TABLE_NAME = "MASTEREDINCTSIE44" ' <C1
Private Const PK_FIELD = "CODE" ' <C2

' field constants
Private Const FIELD_CODE = "CODE" ' <C3
Private Const FIELD_DTYPE = "DTYPE" ' <C4
Private Const FIELD_DOCUMENT_NAME = "DOCUMENT NAME" ' <C5
Private Const FIELD_TREE_ID = "TREE ID" ' <C6
Private Const FIELD_DATE_CREATED = "DATE CREATED" ' <C7
Private Const FIELD_DATE_LAST_MODIFIED = "DATE LAST MODIFIED" ' <C8
Private Const FIELD_DATE_REQUESTED = "DATE REQUESTED" ' <C9
Private Const FIELD_DATE_SEND = "DATE SEND" ' <C10
Private Const FIELD_LAST_MODIFIED_BY = "LAST MODIFIED BY" ' <C11
Private Const FIELD_LOGID_DESCRIPTION = "LOGID DESCRIPTION" ' <C12
Private Const FIELD_REMARKS = "REMARKS" ' <C13
Private Const FIELD_USER_NO = "USER NO" ' <C14
Private Const FIELD_DOC_NUMBER = "DOC NUMBER" ' <C15
Private Const FIELD_DOC_TYPE = "DOC TYPE" ' <C16
Private Const FIELD_Error_String = "Error String" ' <C17
Private Const FIELD_Error_HD = "Error HD" ' <C18
Private Const FIELD_ORIGIN = "ORIGIN" ' <C19
Private Const FIELD_Type = "Type" ' <C20
Private Const FIELD_COMM = "COMM" ' <C21
Private Const FIELD_LOGID = "LOGID" ' <C22
Private Const FIELD_PRINT = "PRINT" ' <C23
Private Const FIELD_VIEWED = "VIEWED" ' <C24
Private Const FIELD_USERNAME = "USERNAME" ' <C25
Private Const FIELD_HEADER = "HEADER" ' <C26
Private Const FIELD_Memo_Field = "Memo Field" ' <C27
Private Const FIELD_DATE_PRINTED = "DATE PRINTED" ' <C28
Private Const FIELD_PRINTED_BY = "PRINTED BY" ' <C29
Private Const FIELD_MR = "MR" ' <C30
Private Const FIELD_BD = "BD" ' <C31
Private Const FIELD_AJ = "AJ" ' <C32
Private Const FIELD_W8 = "W8" ' <C33
Private Const FIELD_W9 = "W9" ' <C34
Private Const FIELD_WA = "WA" ' <C35
Private Const FIELD_WB = "WB" ' <C36
Private Const FIELD_WD = "WD" ' <C37
Private Const FIELD_T7 = "T7" ' <C38
Private Const FIELD_WE = "WE" ' <C39
Private Const FIELD_SB = "SB" ' <C40
Private Const FIELD_DATE_LAST_RECEIVED = "DATE LAST RECEIVED" ' <C41
Private Const FIELD_A9 = "A9" ' <C42
Private Const FIELD_AE = "AE" ' <C43
Private Const FIELD_AF = "AF" ' <C44
Private Const FIELD_B1 = "B1" ' <C45
Private Const FIELD_B7 = "B7" ' <C46
Private Const FIELD_L1 = "L1" ' <C47
Private Const FIELD_M1 = "M1" ' <C48
Private Const FIELD_M2 = "M2" ' <C49
Private Const FIELD_S1 = "S1" ' <C50
Private Const FIELD_S2 = "S2" ' <C51
Private Const FIELD_S3 = "S3" ' <C52
Private Const FIELD_S4 = "S4" ' <C53
Private Const FIELD_S5 = "S5" ' <C54
Private Const FIELD_S6 = "S6" ' <C55
Private Const FIELD_S7 = "S7" ' <C56
Private Const FIELD_S8 = "S8" ' <C57
Private Const FIELD_S9 = "S9" ' <C58
Private Const FIELD_SA = "SA" ' <C59
Private Const FIELD_V1 = "V1" ' <C60
Private Const FIELD_V2 = "V2" ' <C61
Private Const FIELD_V3 = "V3" ' <C62
Private Const FIELD_V4 = "V4" ' <C63
Private Const FIELD_V5 = "V5" ' <C64
Private Const FIELD_V6 = "V6" ' <C65
Private Const FIELD_V7 = "V7" ' <C66
Private Const FIELD_V8 = "V8" ' <C67
Private Const FIELD_Y1 = "Y1" ' <C68
Private Const FIELD_Y2 = "Y2" ' <C69
Private Const FIELD_Y3 = "Y3" ' <C70
Private Const FIELD_Y4 = "Y4" ' <C71
Private Const FIELD_Y5 = "Y5" ' <C72
Private Const FIELD_CC = "CC" ' <C73
Private Const FIELD_CD = "CD" ' <C74
Private Const FIELD_CE = "CE" ' <C75
Private Const FIELD_CF = "CF" ' <C76
Private Const FIELD_CG = "CG" ' <C77
Private Const FIELD_CI = "CI" ' <C78
Private Const FIELD_UA = "UA" ' <C79
Private Const FIELD_UC = "UC" ' <C80
Private Const FIELD_UD = "UD" ' <C81
Private Const FIELD_UE = "UE" ' <C82
Private Const FIELD_UF = "UF" ' <C83
Private Const FIELD_CL = "CL" ' <C84
Private Const FIELD_CH = "CH" ' <C85
Private Const FIELD_REMOTE_ID = "REMOTE_ID" '84

Private Const SQL_ADD_RECORD = "INSERT INTO [" & TABLE_NAME & "] (" & _
              "[CODE],[DTYPE],[DOCUMENT NAME],[TREE ID],[DATE CREATED],[DATE LAST MODIFIED]" & _
              ",[DATE REQUESTED],[DATE SEND],[LAST MODIFIED BY],[LOGID DESCRIPTION],[REMARKS],[USER NO]" & _
              ",[DOC NUMBER],[DOC TYPE],[Error String],[Error HD],[ORIGIN],[Type]" & _
              ",[COMM],[LOGID],[PRINT],[VIEWED],[USERNAME],[HEADER]" & _
              ",[Memo Field],[DATE PRINTED],[PRINTED BY],[MR],[BD],[AJ]" & _
              ",[W8],[W9],[WA],[WB],[WD],[T7]" & _
              ",[WE],[SB],[DATE LAST RECEIVED],[REMOTE_ID],[A9],[AE],[AF]" & _
              ",[B1],[B7],[L1],[M1],[M2],[S1]" & _
              ",[S2],[S3],[S4],[S5],[S6],[S7]" & _
              ",[S8],[S9],[SA],[V1],[V2],[V3]" & _
              ",[V4],[V5],[V6],[V7],[V8],[Y1]" & _
              ",[Y2],[Y3],[Y4],[Y5],[CC],[CD]" & _
              ",[CE],[CF],[CG],[CI],[UA],[UC]" & _
              ",[UD],[UE],[UF],[CL],[CH]) VALUES " ' <C86

Private Const SQL_GET_RECORD = "SELECT [CODE],[DTYPE],[DOCUMENT NAME],[TREE ID],[DATE CREATED],[DATE LAST MODIFIED]" & _
              ",[DATE REQUESTED],[DATE SEND],[LAST MODIFIED BY],[LOGID DESCRIPTION],[REMARKS],[USER NO]" & _
              ",[DOC NUMBER],[DOC TYPE],[Error String],[Error HD],[ORIGIN],[Type]" & _
              ",[COMM],[LOGID],[PRINT],[VIEWED],[USERNAME],[HEADER]" & _
              ",[Memo Field],[DATE PRINTED],[PRINTED BY],[MR],[BD],[AJ]" & _
              ",[W8],[W9],[WA],[WB],[WD],[T7]" & _
              ",[WE],[SB],[DATE LAST RECEIVED],[REMOTE_ID],[A9],[AE],[AF]" & _
              ",[B1],[B7],[L1],[M1],[M2],[S1]" & _
              ",[S2],[S3],[S4],[S5],[S6],[S7]" & _
              ",[S8],[S9],[SA],[V1],[V2],[V3]" & _
              ",[V4],[V5],[V6],[V7],[V8],[Y1]" & _
              ",[Y2],[Y3],[Y4],[Y5],[CC],[CD]" & _
              ",[CE],[CF],[CG],[CI],[UA],[UC]" & _
              ",[UD],[UE],[UF],[CL],[CH] FROM [" & TABLE_NAME & "]" & _
              "WHERE [" & PK_FIELD & "] = " ' <C87

Private Const SQL_GET_ALL_RECORD = "SELECT [CODE],[DTYPE],[DOCUMENT NAME],[TREE ID],[DATE CREATED],[DATE LAST MODIFIED]" & _
              ",[DATE REQUESTED],[DATE SEND],[LAST MODIFIED BY],[LOGID DESCRIPTION],[REMARKS],[USER NO]" & _
              ",[DOC NUMBER],[DOC TYPE],[Error String],[Error HD],[ORIGIN],[Type]" & _
              ",[COMM],[LOGID],[PRINT],[VIEWED],[USERNAME],[HEADER]" & _
              ",[Memo Field],[DATE PRINTED],[PRINTED BY],[MR],[BD],[AJ]" & _
              ",[W8],[W9],[WA],[WB],[WD],[T7]" & _
              ",[WE],[SB],[DATE LAST RECEIVED], [REMOTE_ID], [A9],[AE],[AF]" & _
              ",[B1],[B7],[L1],[M1],[M2],[S1]" & _
              ",[S2],[S3],[S4],[S5],[S6],[S7]" & _
              ",[S8],[S9],[SA],[V1],[V2],[V3]" & _
              ",[V4],[V5],[V6],[V7],[V8],[Y1]" & _
              ",[Y2],[Y3],[Y4],[Y5],[CC],[CD]" & _
              ",[CE],[CF],[CG],[CI],[UA],[UC]" & _
              ",[UD],[UE],[UF],[CL],[CH] FROM [" & TABLE_NAME & "]" ' <C88

Private Const SQL_DELETE_RECORD = "DELETE * FROM [" & TABLE_NAME & "] " & _
              "WHERE [" & PK_FIELD & "] = " ' <C89

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] " & _
              "FROM [" & TABLE_NAME & "]" ' <C90

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] " & _
              "FROM [" & TABLE_NAME & "]" ' <91

' Code: >M1
' Name: AddRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

   Dim strSql As String

   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DTYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_NAME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TREE_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_CREATED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_MODIFIED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_REQUESTED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_SEND) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LAST_MODIFIED_BY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_REMARKS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USER_NO) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOC_NUMBER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOC_TYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Error_String) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Error_HD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_ORIGIN) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Type) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_COMM) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_VIEWED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USERNAME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_HEADER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Memo_Field) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_PRINTED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINTED_BY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_MR) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_AJ) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_W8) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_W9) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WA) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WB) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_T7) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SB) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_RECEIVED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_A9) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_AE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_AF) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_B1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_B7) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_L1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_M1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_M2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S3) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S4) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S5) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S6) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S7) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S8) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_S9) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SA) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V3) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V4) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V5) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V6) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V7) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_V8) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Y1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Y2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Y3) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Y4) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Y5) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CC) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CF) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CG) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CI) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_UA) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_UC) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_UD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_UE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_UF) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CL) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CH) & ")"
      

      On Error GoTo ERROR_QUERY
    
      ExecuteNonQuery ActiveConnection, strSql
      'ActiveConnection.Execute strSql


      AddRecord = True

      Exit Function

  End If


  AddRecord = False


  Exit Function


  AddRecord = False

ERROR_QUERY:

  AddRecord = False

End Function

' Code: >M2
' Name: DeleteRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_CODE)


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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

   Dim strSql As String

   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & TABLE_NAME & "] SET "
      strSql = strSql & " [DTYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_DTYPE) & ", "
      strSql = strSql & " [DOCUMENT NAME]=" & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_NAME) & ", "
      strSql = strSql & " [TREE ID]=" & SetEnclosedChr(ActiveRecord.FIELD_TREE_ID) & ", "
      strSql = strSql & " [DATE CREATED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_CREATED) & ", "
      strSql = strSql & " [DATE LAST MODIFIED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_MODIFIED) & ", "
      strSql = strSql & " [DATE REQUESTED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_REQUESTED) & ", "
      strSql = strSql & " [DATE SEND]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_SEND) & ", "
      strSql = strSql & " [LAST MODIFIED BY]=" & SetEnclosedChr(ActiveRecord.FIELD_LAST_MODIFIED_BY) & ", "
      strSql = strSql & " [LOGID DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION) & ", "
      strSql = strSql & " [REMARKS]=" & SetEnclosedChr(ActiveRecord.FIELD_REMARKS) & ", "
      strSql = strSql & " [USER NO]=" & SetEnclosedChr(ActiveRecord.FIELD_USER_NO) & ", "
      strSql = strSql & " [DOC NUMBER]=" & SetEnclosedChr(ActiveRecord.FIELD_DOC_NUMBER) & ", "
      strSql = strSql & " [DOC TYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_DOC_TYPE) & ", "
      strSql = strSql & " [Error String]=" & SetEnclosedChr(ActiveRecord.FIELD_Error_String) & ", "
      strSql = strSql & " [Error HD]=" & SetEnclosedChr(ActiveRecord.FIELD_Error_HD) & ", "
      strSql = strSql & " [ORIGIN]=" & SetEnclosedChr(ActiveRecord.FIELD_ORIGIN) & ", "
      strSql = strSql & " [Type]=" & SetEnclosedChr(ActiveRecord.FIELD_Type) & ", "
      strSql = strSql & " [COMM]=" & SetEnclosedChr(ActiveRecord.FIELD_COMM) & ", "
      strSql = strSql & " [LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_LOGID) & ", "
      strSql = strSql & " [PRINT]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT) & ", "
      strSql = strSql & " [VIEWED]=" & SetEnclosedChr(ActiveRecord.FIELD_VIEWED) & ", "
      strSql = strSql & " [USERNAME]=" & SetEnclosedChr(ActiveRecord.FIELD_USERNAME) & ", "
      strSql = strSql & " [HEADER]=" & SetEnclosedChr(ActiveRecord.FIELD_HEADER) & ", "
      strSql = strSql & " [Memo Field]=" & SetEnclosedChr(ActiveRecord.FIELD_Memo_Field) & ", "
      strSql = strSql & " [DATE PRINTED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_PRINTED) & ", "
      strSql = strSql & " [PRINTED BY]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINTED_BY) & ", "
      strSql = strSql & " [MR]=" & SetEnclosedChr(ActiveRecord.FIELD_MR) & ", "
      strSql = strSql & " [BD]=" & SetEnclosedChr(ActiveRecord.FIELD_BD) & ", "
      strSql = strSql & " [AJ]=" & SetEnclosedChr(ActiveRecord.FIELD_AJ) & ", "
      strSql = strSql & " [W8]=" & SetEnclosedChr(ActiveRecord.FIELD_W8) & ", "
      strSql = strSql & " [W9]=" & SetEnclosedChr(ActiveRecord.FIELD_W9) & ", "
      strSql = strSql & " [WA]=" & SetEnclosedChr(ActiveRecord.FIELD_WA) & ", "
      strSql = strSql & " [WB]=" & SetEnclosedChr(ActiveRecord.FIELD_WB) & ", "
      strSql = strSql & " [WD]=" & SetEnclosedChr(ActiveRecord.FIELD_WD) & ", "
      strSql = strSql & " [T7]=" & SetEnclosedChr(ActiveRecord.FIELD_T7) & ", "
      strSql = strSql & " [WE]=" & SetEnclosedChr(ActiveRecord.FIELD_WE) & ", "
      strSql = strSql & " [SB]=" & SetEnclosedChr(ActiveRecord.FIELD_SB) & ", "
      strSql = strSql & " [DATE LAST RECEIVED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_RECEIVED) & ", "
      strSql = strSql & " [REMOTE_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID) & ","
      strSql = strSql & " [A9]=" & SetEnclosedChr(ActiveRecord.FIELD_A9) & ", "
      strSql = strSql & " [AE]=" & SetEnclosedChr(ActiveRecord.FIELD_AE) & ", "
      strSql = strSql & " [AF]=" & SetEnclosedChr(ActiveRecord.FIELD_AF) & ", "
      strSql = strSql & " [B1]=" & SetEnclosedChr(ActiveRecord.FIELD_B1) & ", "
      strSql = strSql & " [B7]=" & SetEnclosedChr(ActiveRecord.FIELD_B7) & ", "
      strSql = strSql & " [L1]=" & SetEnclosedChr(ActiveRecord.FIELD_L1) & ", "
      strSql = strSql & " [M1]=" & SetEnclosedChr(ActiveRecord.FIELD_M1) & ", "
      strSql = strSql & " [M2]=" & SetEnclosedChr(ActiveRecord.FIELD_M2) & ", "
      strSql = strSql & " [S1]=" & SetEnclosedChr(ActiveRecord.FIELD_S1) & ", "
      strSql = strSql & " [S2]=" & SetEnclosedChr(ActiveRecord.FIELD_S2) & ", "
      strSql = strSql & " [S3]=" & SetEnclosedChr(ActiveRecord.FIELD_S3) & ", "
      strSql = strSql & " [S4]=" & SetEnclosedChr(ActiveRecord.FIELD_S4) & ", "
      strSql = strSql & " [S5]=" & SetEnclosedChr(ActiveRecord.FIELD_S5) & ", "
      strSql = strSql & " [S6]=" & SetEnclosedChr(ActiveRecord.FIELD_S6) & ", "
      strSql = strSql & " [S7]=" & SetEnclosedChr(ActiveRecord.FIELD_S7) & ", "
      strSql = strSql & " [S8]=" & SetEnclosedChr(ActiveRecord.FIELD_S8) & ", "
      strSql = strSql & " [S9]=" & SetEnclosedChr(ActiveRecord.FIELD_S9) & ", "
      strSql = strSql & " [SA]=" & SetEnclosedChr(ActiveRecord.FIELD_SA) & ", "
      strSql = strSql & " [V1]=" & SetEnclosedChr(ActiveRecord.FIELD_V1) & ", "
      strSql = strSql & " [V2]=" & SetEnclosedChr(ActiveRecord.FIELD_V2) & ", "
      strSql = strSql & " [V3]=" & SetEnclosedChr(ActiveRecord.FIELD_V3) & ", "
      strSql = strSql & " [V4]=" & SetEnclosedChr(ActiveRecord.FIELD_V4) & ", "
      strSql = strSql & " [V5]=" & SetEnclosedChr(ActiveRecord.FIELD_V5) & ", "
      strSql = strSql & " [V6]=" & SetEnclosedChr(ActiveRecord.FIELD_V6) & ", "
      strSql = strSql & " [V7]=" & SetEnclosedChr(ActiveRecord.FIELD_V7) & ", "
      strSql = strSql & " [V8]=" & SetEnclosedChr(ActiveRecord.FIELD_V8) & ", "
      strSql = strSql & " [Y1]=" & SetEnclosedChr(ActiveRecord.FIELD_Y1) & ", "
      strSql = strSql & " [Y2]=" & SetEnclosedChr(ActiveRecord.FIELD_Y2) & ", "
      strSql = strSql & " [Y3]=" & SetEnclosedChr(ActiveRecord.FIELD_Y3) & ", "
      strSql = strSql & " [Y4]=" & SetEnclosedChr(ActiveRecord.FIELD_Y4) & ", "
      strSql = strSql & " [Y5]=" & SetEnclosedChr(ActiveRecord.FIELD_Y5) & ", "
      strSql = strSql & " [CC]=" & SetEnclosedChr(ActiveRecord.FIELD_CC) & ", "
      strSql = strSql & " [CD]=" & SetEnclosedChr(ActiveRecord.FIELD_CD) & ", "
      strSql = strSql & " [CE]=" & SetEnclosedChr(ActiveRecord.FIELD_CE) & ", "
      strSql = strSql & " [CF]=" & SetEnclosedChr(ActiveRecord.FIELD_CF) & ", "
      strSql = strSql & " [CG]=" & SetEnclosedChr(ActiveRecord.FIELD_CG) & ", "
      strSql = strSql & " [CI]=" & SetEnclosedChr(ActiveRecord.FIELD_CI) & ", "
      strSql = strSql & " [UA]=" & SetEnclosedChr(ActiveRecord.FIELD_UA) & ", "
      strSql = strSql & " [UC]=" & SetEnclosedChr(ActiveRecord.FIELD_UC) & ", "
      strSql = strSql & " [UD]=" & SetEnclosedChr(ActiveRecord.FIELD_UD) & ", "
      strSql = strSql & " [UE]=" & SetEnclosedChr(ActiveRecord.FIELD_UE) & ", "
      strSql = strSql & " [UF]=" & SetEnclosedChr(ActiveRecord.FIELD_UF) & ", "
      strSql = strSql & " [CL]=" & SetEnclosedChr(ActiveRecord.FIELD_CL) & ", "
      strSql = strSql & " [CH]=" & SetEnclosedChr(ActiveRecord.FIELD_CH) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_CODE)

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
Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_CODE)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("CODE"))
      ActiveRecord.FIELD_DTYPE = FNullField(rstRecord.Fields("DTYPE"))
      ActiveRecord.FIELD_DOCUMENT_NAME = FNullField(rstRecord.Fields("DOCUMENT NAME"))
      ActiveRecord.FIELD_TREE_ID = FNullField(rstRecord.Fields("TREE ID"))
      ActiveRecord.FIELD_DATE_CREATED = FNullField(rstRecord.Fields("DATE CREATED"))
      ActiveRecord.FIELD_DATE_LAST_MODIFIED = FNullField(rstRecord.Fields("DATE LAST MODIFIED"))
      ActiveRecord.FIELD_DATE_REQUESTED = FNullField(rstRecord.Fields("DATE REQUESTED"))
      ActiveRecord.FIELD_DATE_SEND = FNullField(rstRecord.Fields("DATE SEND"))
      ActiveRecord.FIELD_LAST_MODIFIED_BY = FNullField(rstRecord.Fields("LAST MODIFIED BY"))
      ActiveRecord.FIELD_LOGID_DESCRIPTION = FNullField(rstRecord.Fields("LOGID DESCRIPTION"))
      ActiveRecord.FIELD_REMARKS = FNullField(rstRecord.Fields("REMARKS"))
      ActiveRecord.FIELD_USER_NO = FNullField(rstRecord.Fields("USER NO"))
      ActiveRecord.FIELD_DOC_NUMBER = FNullField(rstRecord.Fields("DOC NUMBER"))
      ActiveRecord.FIELD_DOC_TYPE = FNullField(rstRecord.Fields("DOC TYPE"))
      ActiveRecord.FIELD_Error_String = FNullField(rstRecord.Fields("Error String"))
      ActiveRecord.FIELD_Error_HD = FNullField(rstRecord.Fields("Error HD"))
      ActiveRecord.FIELD_ORIGIN = FNullField(rstRecord.Fields("ORIGIN"))
      ActiveRecord.FIELD_Type = FNullField(rstRecord.Fields("Type"))
      ActiveRecord.FIELD_COMM = FNullField(rstRecord.Fields("COMM"))
      ActiveRecord.FIELD_LOGID = FNullField(rstRecord.Fields("LOGID"))
      ActiveRecord.FIELD_PRINT = FNullField(rstRecord.Fields("PRINT"))
      ActiveRecord.FIELD_VIEWED = FNullField(rstRecord.Fields("VIEWED"))
      ActiveRecord.FIELD_USERNAME = FNullField(rstRecord.Fields("USERNAME"))
      ActiveRecord.FIELD_HEADER = FNullField(rstRecord.Fields("HEADER"))
      ActiveRecord.FIELD_Memo_Field = FNullField(rstRecord.Fields("Memo Field"))
      ActiveRecord.FIELD_DATE_PRINTED = FNullField(rstRecord.Fields("DATE PRINTED"))
      ActiveRecord.FIELD_PRINTED_BY = FNullField(rstRecord.Fields("PRINTED BY"))
      ActiveRecord.FIELD_MR = FNullField(rstRecord.Fields("MR"))
      ActiveRecord.FIELD_BD = FNullField(rstRecord.Fields("BD"))
      ActiveRecord.FIELD_AJ = FNullField(rstRecord.Fields("AJ"))
      ActiveRecord.FIELD_W8 = FNullField(rstRecord.Fields("W8"))
      ActiveRecord.FIELD_W9 = FNullField(rstRecord.Fields("W9"))
      ActiveRecord.FIELD_WA = FNullField(rstRecord.Fields("WA"))
      ActiveRecord.FIELD_WB = FNullField(rstRecord.Fields("WB"))
      ActiveRecord.FIELD_WD = FNullField(rstRecord.Fields("WD"))
      ActiveRecord.FIELD_T7 = FNullField(rstRecord.Fields("T7"))
      ActiveRecord.FIELD_WE = FNullField(rstRecord.Fields("WE"))
      ActiveRecord.FIELD_SB = FNullField(rstRecord.Fields("SB"))
      ActiveRecord.FIELD_DATE_LAST_RECEIVED = FNullField(rstRecord.Fields("DATE LAST RECEIVED"))
      ActiveRecord.FIELD_REMOTE_ID = FNullField(rstRecord.Fields("REMOTE_ID"))
      ActiveRecord.FIELD_A9 = FNullField(rstRecord.Fields("A9"))
      ActiveRecord.FIELD_AE = FNullField(rstRecord.Fields("AE"))
      ActiveRecord.FIELD_AF = FNullField(rstRecord.Fields("AF"))
      ActiveRecord.FIELD_B1 = FNullField(rstRecord.Fields("B1"))
      ActiveRecord.FIELD_B7 = FNullField(rstRecord.Fields("B7"))
      ActiveRecord.FIELD_L1 = FNullField(rstRecord.Fields("L1"))
      ActiveRecord.FIELD_M1 = FNullField(rstRecord.Fields("M1"))
      ActiveRecord.FIELD_M2 = FNullField(rstRecord.Fields("M2"))
      ActiveRecord.FIELD_S1 = FNullField(rstRecord.Fields("S1"))
      ActiveRecord.FIELD_S2 = FNullField(rstRecord.Fields("S2"))
      ActiveRecord.FIELD_S3 = FNullField(rstRecord.Fields("S3"))
      ActiveRecord.FIELD_S4 = FNullField(rstRecord.Fields("S4"))
      ActiveRecord.FIELD_S5 = FNullField(rstRecord.Fields("S5"))
      ActiveRecord.FIELD_S6 = FNullField(rstRecord.Fields("S6"))
      ActiveRecord.FIELD_S7 = FNullField(rstRecord.Fields("S7"))
      ActiveRecord.FIELD_S8 = FNullField(rstRecord.Fields("S8"))
      ActiveRecord.FIELD_S9 = FNullField(rstRecord.Fields("S9"))
      ActiveRecord.FIELD_SA = FNullField(rstRecord.Fields("SA"))
      ActiveRecord.FIELD_V1 = FNullField(rstRecord.Fields("V1"))
      ActiveRecord.FIELD_V2 = FNullField(rstRecord.Fields("V2"))
      ActiveRecord.FIELD_V3 = FNullField(rstRecord.Fields("V3"))
      ActiveRecord.FIELD_V4 = FNullField(rstRecord.Fields("V4"))
      ActiveRecord.FIELD_V5 = FNullField(rstRecord.Fields("V5"))
      ActiveRecord.FIELD_V6 = FNullField(rstRecord.Fields("V6"))
      ActiveRecord.FIELD_V7 = FNullField(rstRecord.Fields("V7"))
      ActiveRecord.FIELD_V8 = FNullField(rstRecord.Fields("V8"))
      ActiveRecord.FIELD_Y1 = FNullField(rstRecord.Fields("Y1"))
      ActiveRecord.FIELD_Y2 = FNullField(rstRecord.Fields("Y2"))
      ActiveRecord.FIELD_Y3 = FNullField(rstRecord.Fields("Y3"))
      ActiveRecord.FIELD_Y4 = FNullField(rstRecord.Fields("Y4"))
      ActiveRecord.FIELD_Y5 = FNullField(rstRecord.Fields("Y5"))
      ActiveRecord.FIELD_CC = FNullField(rstRecord.Fields("CC"))
      ActiveRecord.FIELD_CD = FNullField(rstRecord.Fields("CD"))
      ActiveRecord.FIELD_CE = FNullField(rstRecord.Fields("CE"))
      ActiveRecord.FIELD_CF = FNullField(rstRecord.Fields("CF"))
      ActiveRecord.FIELD_CG = FNullField(rstRecord.Fields("CG"))
      ActiveRecord.FIELD_CI = FNullField(rstRecord.Fields("CI"))
      ActiveRecord.FIELD_UA = FNullField(rstRecord.Fields("UA"))
      ActiveRecord.FIELD_UC = FNullField(rstRecord.Fields("UC"))
      ActiveRecord.FIELD_UD = FNullField(rstRecord.Fields("UD"))
      ActiveRecord.FIELD_UE = FNullField(rstRecord.Fields("UE"))
      ActiveRecord.FIELD_UF = FNullField(rstRecord.Fields("UF"))
      ActiveRecord.FIELD_CL = FNullField(rstRecord.Fields("CL"))
      ActiveRecord.FIELD_CH = FNullField(rstRecord.Fields("CH"))
      
      
      ADORecordsetClose rstRecord
      'Set rstRecord = Nothing

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


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
Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiMASTEREDINCTSIE44

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("ID_MAX"))
      
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
Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiMASTEREDINCTSIE44

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("ID_MIN"))
      
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
Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMASTEREDINCTSIE44) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   'rstRecord.Open strSql, ActiveConnection, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("CODE") = ActiveRecord.FIELD_CODE
   rstRecord.Fields("DTYPE") = ActiveRecord.FIELD_DTYPE
   rstRecord.Fields("DOCUMENT NAME") = ActiveRecord.FIELD_DOCUMENT_NAME
   rstRecord.Fields("TREE ID") = ActiveRecord.FIELD_TREE_ID
   rstRecord.Fields("DATE CREATED") = ActiveRecord.FIELD_DATE_CREATED
   rstRecord.Fields("DATE LAST MODIFIED") = ActiveRecord.FIELD_DATE_LAST_MODIFIED
   rstRecord.Fields("DATE REQUESTED") = ActiveRecord.FIELD_DATE_REQUESTED
   rstRecord.Fields("DATE SEND") = ActiveRecord.FIELD_DATE_SEND
   rstRecord.Fields("LAST MODIFIED BY") = ActiveRecord.FIELD_LAST_MODIFIED_BY
   rstRecord.Fields("LOGID DESCRIPTION") = ActiveRecord.FIELD_LOGID_DESCRIPTION
   rstRecord.Fields("REMARKS") = ActiveRecord.FIELD_REMARKS
   rstRecord.Fields("USER NO") = ActiveRecord.FIELD_USER_NO
   rstRecord.Fields("DOC NUMBER") = ActiveRecord.FIELD_DOC_NUMBER
   rstRecord.Fields("DOC TYPE") = ActiveRecord.FIELD_DOC_TYPE
   rstRecord.Fields("Error String") = ActiveRecord.FIELD_Error_String
   rstRecord.Fields("Error HD") = ActiveRecord.FIELD_Error_HD
   rstRecord.Fields("ORIGIN") = ActiveRecord.FIELD_ORIGIN
   rstRecord.Fields("Type") = ActiveRecord.FIELD_Type
   rstRecord.Fields("COMM") = ActiveRecord.FIELD_COMM
   rstRecord.Fields("LOGID") = ActiveRecord.FIELD_LOGID
   rstRecord.Fields("PRINT") = ActiveRecord.FIELD_PRINT
   rstRecord.Fields("VIEWED") = ActiveRecord.FIELD_VIEWED
   rstRecord.Fields("USERNAME") = ActiveRecord.FIELD_USERNAME
   rstRecord.Fields("HEADER") = ActiveRecord.FIELD_HEADER
   rstRecord.Fields("Memo Field") = ActiveRecord.FIELD_Memo_Field
   rstRecord.Fields("DATE PRINTED") = ActiveRecord.FIELD_DATE_PRINTED
   rstRecord.Fields("PRINTED BY") = ActiveRecord.FIELD_PRINTED_BY
   rstRecord.Fields("MR") = ActiveRecord.FIELD_MR
   rstRecord.Fields("BD") = ActiveRecord.FIELD_BD
   rstRecord.Fields("AJ") = ActiveRecord.FIELD_AJ
   rstRecord.Fields("W8") = ActiveRecord.FIELD_W8
   rstRecord.Fields("W9") = ActiveRecord.FIELD_W9
   rstRecord.Fields("WA") = ActiveRecord.FIELD_WA
   rstRecord.Fields("WB") = ActiveRecord.FIELD_WB
   rstRecord.Fields("WD") = ActiveRecord.FIELD_WD
   rstRecord.Fields("T7") = ActiveRecord.FIELD_T7
   rstRecord.Fields("WE") = ActiveRecord.FIELD_WE
   rstRecord.Fields("SB") = ActiveRecord.FIELD_SB
   rstRecord.Fields("DATE LAST RECEIVED") = ActiveRecord.FIELD_DATE_LAST_RECEIVED
   rstRecord.Fields("REMOTE_ID") = ActiveRecord.FIELD_REMOTE_ID
   rstRecord.Fields("A9") = ActiveRecord.FIELD_A9
   rstRecord.Fields("AE") = ActiveRecord.FIELD_AE
   rstRecord.Fields("AF") = ActiveRecord.FIELD_AF
   rstRecord.Fields("B1") = ActiveRecord.FIELD_B1
   rstRecord.Fields("B7") = ActiveRecord.FIELD_B7
   rstRecord.Fields("L1") = ActiveRecord.FIELD_L1
   rstRecord.Fields("M1") = ActiveRecord.FIELD_M1
   rstRecord.Fields("M2") = ActiveRecord.FIELD_M2
   rstRecord.Fields("S1") = ActiveRecord.FIELD_S1
   rstRecord.Fields("S2") = ActiveRecord.FIELD_S2
   rstRecord.Fields("S3") = ActiveRecord.FIELD_S3
   rstRecord.Fields("S4") = ActiveRecord.FIELD_S4
   rstRecord.Fields("S5") = ActiveRecord.FIELD_S5
   rstRecord.Fields("S6") = ActiveRecord.FIELD_S6
   rstRecord.Fields("S7") = ActiveRecord.FIELD_S7
   rstRecord.Fields("S8") = ActiveRecord.FIELD_S8
   rstRecord.Fields("S9") = ActiveRecord.FIELD_S9
   rstRecord.Fields("SA") = ActiveRecord.FIELD_SA
   rstRecord.Fields("V1") = ActiveRecord.FIELD_V1
   rstRecord.Fields("V2") = ActiveRecord.FIELD_V2
   rstRecord.Fields("V3") = ActiveRecord.FIELD_V3
   rstRecord.Fields("V4") = ActiveRecord.FIELD_V4
   rstRecord.Fields("V5") = ActiveRecord.FIELD_V5
   rstRecord.Fields("V6") = ActiveRecord.FIELD_V6
   rstRecord.Fields("V7") = ActiveRecord.FIELD_V7
   rstRecord.Fields("V8") = ActiveRecord.FIELD_V8
   rstRecord.Fields("Y1") = ActiveRecord.FIELD_Y1
   rstRecord.Fields("Y2") = ActiveRecord.FIELD_Y2
   rstRecord.Fields("Y3") = ActiveRecord.FIELD_Y3
   rstRecord.Fields("Y4") = ActiveRecord.FIELD_Y4
   rstRecord.Fields("Y5") = ActiveRecord.FIELD_Y5
   rstRecord.Fields("CC") = ActiveRecord.FIELD_CC
   rstRecord.Fields("CD") = ActiveRecord.FIELD_CD
   rstRecord.Fields("CE") = ActiveRecord.FIELD_CE
   rstRecord.Fields("CF") = ActiveRecord.FIELD_CF
   rstRecord.Fields("CG") = ActiveRecord.FIELD_CG
   rstRecord.Fields("CI") = ActiveRecord.FIELD_CI
   rstRecord.Fields("UA") = ActiveRecord.FIELD_UA
   rstRecord.Fields("UC") = ActiveRecord.FIELD_UC
   rstRecord.Fields("UD") = ActiveRecord.FIELD_UD
   rstRecord.Fields("UE") = ActiveRecord.FIELD_UE
   rstRecord.Fields("UF") = ActiveRecord.FIELD_UF
   rstRecord.Fields("CL") = ActiveRecord.FIELD_CL
   rstRecord.Fields("CH") = ActiveRecord.FIELD_CH
   

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
Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiMASTEREDINCTSIE44

   Dim clsRecord As cpiMASTEREDINCTSIE44
   Set clsRecord = New cpiMASTEREDINCTSIE44

   clsRecord.FIELD_CODE = FNullField(ActiveRecord.Fields(FIELD_CODE))
   clsRecord.FIELD_DTYPE = FNullField(ActiveRecord.Fields(FIELD_DTYPE))
   clsRecord.FIELD_DOCUMENT_NAME = FNullField(ActiveRecord.Fields(FIELD_DOCUMENT_NAME))
   clsRecord.FIELD_TREE_ID = FNullField(ActiveRecord.Fields(FIELD_TREE_ID))
   clsRecord.FIELD_DATE_CREATED = FNullField(ActiveRecord.Fields(FIELD_DATE_CREATED))
   clsRecord.FIELD_DATE_LAST_MODIFIED = FNullField(ActiveRecord.Fields(FIELD_DATE_LAST_MODIFIED))
   clsRecord.FIELD_DATE_REQUESTED = FNullField(ActiveRecord.Fields(FIELD_DATE_REQUESTED))
   clsRecord.FIELD_DATE_SEND = FNullField(ActiveRecord.Fields(FIELD_DATE_SEND))
   clsRecord.FIELD_LAST_MODIFIED_BY = FNullField(ActiveRecord.Fields(FIELD_LAST_MODIFIED_BY))
   clsRecord.FIELD_LOGID_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_LOGID_DESCRIPTION))
   clsRecord.FIELD_REMARKS = FNullField(ActiveRecord.Fields(FIELD_REMARKS))
   clsRecord.FIELD_USER_NO = FNullField(ActiveRecord.Fields(FIELD_USER_NO))
   clsRecord.FIELD_DOC_NUMBER = FNullField(ActiveRecord.Fields(FIELD_DOC_NUMBER))
   clsRecord.FIELD_DOC_TYPE = FNullField(ActiveRecord.Fields(FIELD_DOC_TYPE))
   clsRecord.FIELD_Error_String = FNullField(ActiveRecord.Fields(FIELD_Error_String))
   clsRecord.FIELD_Error_HD = FNullField(ActiveRecord.Fields(FIELD_Error_HD))
   clsRecord.FIELD_ORIGIN = FNullField(ActiveRecord.Fields(FIELD_ORIGIN))
   clsRecord.FIELD_Type = FNullField(ActiveRecord.Fields(FIELD_Type))
   clsRecord.FIELD_COMM = FNullField(ActiveRecord.Fields(FIELD_COMM))
   clsRecord.FIELD_LOGID = FNullField(ActiveRecord.Fields(FIELD_LOGID))
   clsRecord.FIELD_PRINT = FNullField(ActiveRecord.Fields(FIELD_PRINT))
   clsRecord.FIELD_VIEWED = FNullField(ActiveRecord.Fields(FIELD_VIEWED))
   clsRecord.FIELD_USERNAME = FNullField(ActiveRecord.Fields(FIELD_USERNAME))
   clsRecord.FIELD_HEADER = FNullField(ActiveRecord.Fields(FIELD_HEADER))
   clsRecord.FIELD_Memo_Field = FNullField(ActiveRecord.Fields(FIELD_Memo_Field))
   clsRecord.FIELD_DATE_PRINTED = FNullField(ActiveRecord.Fields(FIELD_DATE_PRINTED))
   clsRecord.FIELD_PRINTED_BY = FNullField(ActiveRecord.Fields(FIELD_PRINTED_BY))
   clsRecord.FIELD_MR = FNullField(ActiveRecord.Fields(FIELD_MR))
   clsRecord.FIELD_BD = FNullField(ActiveRecord.Fields(FIELD_BD))
   clsRecord.FIELD_AJ = FNullField(ActiveRecord.Fields(FIELD_AJ))
   clsRecord.FIELD_W8 = FNullField(ActiveRecord.Fields(FIELD_W8))
   clsRecord.FIELD_W9 = FNullField(ActiveRecord.Fields(FIELD_W9))
   clsRecord.FIELD_WA = FNullField(ActiveRecord.Fields(FIELD_WA))
   clsRecord.FIELD_WB = FNullField(ActiveRecord.Fields(FIELD_WB))
   clsRecord.FIELD_WD = FNullField(ActiveRecord.Fields(FIELD_WD))
   clsRecord.FIELD_T7 = FNullField(ActiveRecord.Fields(FIELD_T7))
   clsRecord.FIELD_WE = FNullField(ActiveRecord.Fields(FIELD_WE))
   clsRecord.FIELD_SB = FNullField(ActiveRecord.Fields(FIELD_SB))
   clsRecord.FIELD_DATE_LAST_RECEIVED = FNullField(ActiveRecord.Fields(FIELD_DATE_LAST_RECEIVED))
   clsRecord.FIELD_A9 = FNullField(ActiveRecord.Fields(FIELD_A9))
   clsRecord.FIELD_AE = FNullField(ActiveRecord.Fields(FIELD_AE))
   clsRecord.FIELD_AF = FNullField(ActiveRecord.Fields(FIELD_AF))
   clsRecord.FIELD_B1 = FNullField(ActiveRecord.Fields(FIELD_B1))
   clsRecord.FIELD_B7 = FNullField(ActiveRecord.Fields(FIELD_B7))
   clsRecord.FIELD_L1 = FNullField(ActiveRecord.Fields(FIELD_L1))
   clsRecord.FIELD_M1 = FNullField(ActiveRecord.Fields(FIELD_M1))
   clsRecord.FIELD_M2 = FNullField(ActiveRecord.Fields(FIELD_M2))
   clsRecord.FIELD_S1 = FNullField(ActiveRecord.Fields(FIELD_S1))
   clsRecord.FIELD_S2 = FNullField(ActiveRecord.Fields(FIELD_S2))
   clsRecord.FIELD_S3 = FNullField(ActiveRecord.Fields(FIELD_S3))
   clsRecord.FIELD_S4 = FNullField(ActiveRecord.Fields(FIELD_S4))
   clsRecord.FIELD_S5 = FNullField(ActiveRecord.Fields(FIELD_S5))
   clsRecord.FIELD_S6 = FNullField(ActiveRecord.Fields(FIELD_S6))
   clsRecord.FIELD_S7 = FNullField(ActiveRecord.Fields(FIELD_S7))
   clsRecord.FIELD_S8 = FNullField(ActiveRecord.Fields(FIELD_S8))
   clsRecord.FIELD_S9 = FNullField(ActiveRecord.Fields(FIELD_S9))
   clsRecord.FIELD_SA = FNullField(ActiveRecord.Fields(FIELD_SA))
   clsRecord.FIELD_V1 = FNullField(ActiveRecord.Fields(FIELD_V1))
   clsRecord.FIELD_V2 = FNullField(ActiveRecord.Fields(FIELD_V2))
   clsRecord.FIELD_V3 = FNullField(ActiveRecord.Fields(FIELD_V3))
   clsRecord.FIELD_V4 = FNullField(ActiveRecord.Fields(FIELD_V4))
   clsRecord.FIELD_V5 = FNullField(ActiveRecord.Fields(FIELD_V5))
   clsRecord.FIELD_V6 = FNullField(ActiveRecord.Fields(FIELD_V6))
   clsRecord.FIELD_V7 = FNullField(ActiveRecord.Fields(FIELD_V7))
   clsRecord.FIELD_V8 = FNullField(ActiveRecord.Fields(FIELD_V8))
   clsRecord.FIELD_Y1 = FNullField(ActiveRecord.Fields(FIELD_Y1))
   clsRecord.FIELD_Y2 = FNullField(ActiveRecord.Fields(FIELD_Y2))
   clsRecord.FIELD_Y3 = FNullField(ActiveRecord.Fields(FIELD_Y3))
   clsRecord.FIELD_Y4 = FNullField(ActiveRecord.Fields(FIELD_Y4))
   clsRecord.FIELD_Y5 = FNullField(ActiveRecord.Fields(FIELD_Y5))
   clsRecord.FIELD_CC = FNullField(ActiveRecord.Fields(FIELD_CC))
   clsRecord.FIELD_CD = FNullField(ActiveRecord.Fields(FIELD_CD))
   clsRecord.FIELD_CE = FNullField(ActiveRecord.Fields(FIELD_CE))
   clsRecord.FIELD_CF = FNullField(ActiveRecord.Fields(FIELD_CF))
   clsRecord.FIELD_CG = FNullField(ActiveRecord.Fields(FIELD_CG))
   clsRecord.FIELD_CI = FNullField(ActiveRecord.Fields(FIELD_CI))
   clsRecord.FIELD_UA = FNullField(ActiveRecord.Fields(FIELD_UA))
   clsRecord.FIELD_UC = FNullField(ActiveRecord.Fields(FIELD_UC))
   clsRecord.FIELD_UD = FNullField(ActiveRecord.Fields(FIELD_UD))
   clsRecord.FIELD_UE = FNullField(ActiveRecord.Fields(FIELD_UE))
   clsRecord.FIELD_UF = FNullField(ActiveRecord.Fields(FIELD_UF))
   clsRecord.FIELD_CL = FNullField(ActiveRecord.Fields(FIELD_CL))
   clsRecord.FIELD_CH = FNullField(ActiveRecord.Fields(FIELD_CH))
   
   
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
      strSql = "SELECT TOP 1 " & SearchField & " FROM [" & TABLE_NAME & "] WHERE " & SearchField & "=" & SetEnclosedChr((SearchValue))
      
      ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
      'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
    
      SearchRecord = Not (rstDB.EOF And rstDB.BOF)
      
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
    
   ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic
   'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function

' Code: >M11
' Name: Add
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [clsMASTEREDINCTSIE44] = Add([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiMASTEREDINCTSIE44

    'create a new object
    Dim objNewMember As cpiMASTEREDINCTSIE44
    Set objNewMember = New cpiMASTEREDINCTSIE44

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    'objNewMember.EDIPROP_ID = PKValue
    'objNewMember.ORDINAL = mCol.Count + 1

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

' Code: >P1
' Name: Item (Get)
' Description: return Item
' Scope: Public
' Fan-In: <none>
' Data Type: cpiMASTEREDINCTSIE44
Public Property Get Item(vntIndexKey As Variant) As cpiMASTEREDINCTSIE44
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

' Code: >P2
' Name: Count (Get)
' Description: return Count
' Scope: Public
' Fan-In: <none>
' Data Type: Long
Public Property Get Count() As Long
    Count = mCol.Count
End Property

' Code: >M12
' Name: Remove
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [clsMASTEREDINCTSIE44] = Add([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contEDIPROP_UserPasswordns either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub

' Code: >P3
' Name: NewEnum (Get)
' Description: return NewEnum
' Scope: Public
' Fan-In: <none>
' Data Type: IUnknown
Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

' Code: <M13
' Name: Class_Initialize
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [clsMASTEREDINCTSIE44] = Class_Initialize([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Set mvarRecordset = New ADODB.Recordset
End Sub

' Code: <M14
' Name: Class_Terminate
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [clsMASTEREDINCTSIE44] = Class_Initialize([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarRecordset = Nothing
End Sub

' Code: >P4
' Name: Recordset (Get/Set)
' Description: return/set Recordset
' Scope: Public
' Fan-In: <none>
' Data Type: ADODB.Recordset
Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mvarRecordset
End Property
Public Property Set Recordset(ByVal vNewValue As ADODB.Recordset)
    Set mvarRecordset = vNewValue
End Property

