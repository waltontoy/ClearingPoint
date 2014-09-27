VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDataNctsTables"
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
' Program Name: cpiDataNctsTables
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: Collection of cpiDataNctsTable class
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

Private Const TABLE_NAME = "DATA_NCTS" ' <C1
Private Const PK_FIELD = "DATA_NCTS_ID" ' <C2
' field constants
Private Const FIELD_DATA_NCTS_ID = "DATA_NCTS_ID" ' <C3
Private Const FIELD_DATA_NCTS_MSG_ID = "DATA_NCTS_MSG_ID" ' <C4
Private Const FIELD_CODE = "CODE" ' <C5
Private Const FIELD_LOGID_DESCRIPTION = "LOGID DESCRIPTION" ' <C6
Private Const FIELD_TYPE = "TYPE" ' <C7
Private Const FIELD_COMM = "COMM" ' <C8
Private Const FIELD_USER_NO = "USER NO" ' <C9
Private Const FIELD_LAST_MODIFIED_BY = "LAST MODIFIED BY" ' <C10
Private Const FIELD_PRINT = "PRINT" ' <C11
Private Const FIELD_DOCUMENT_NAME = "DOCUMENT NAME" ' <C12
Private Const FIELD_DATE_CREATED = "DATE CREATED" ' <C13
Private Const FIELD_DATE_LAST_MODIFIED = "DATE LAST MODIFIED" ' <C14
Private Const FIELD_DATE_REQUESTED = "DATE REQUESTED" ' <C15
Private Const FIELD_DATE_SEND = "DATE SEND" ' <C16
Private Const FIELD_HEADER = "HEADER" ' <C17
Private Const FIELD_DETAIL = "DETAIL" ' <C18
Private Const FIELD_DOCUMENT_COUNTER = "DOCUMENT COUNTER" ' <C19
Private Const FIELD_REMARKS = "REMARKS" ' <C20
Private Const FIELD_TREE_ID = "TREE ID" ' <C21
Private Const FIELD_SUPPLIER_NAME = "SUPPLIER NAME" ' <C22
Private Const FIELD_DTYPE = "DTYPE" ' <C23
Private Const FIELD_Error_String = "Error String" ' <C24
Private Const FIELD_Error_HD = "Error HD" ' <C25
Private Const FIELD_LOGID = "LOGID" ' <C26
Private Const FIELD_USERNAME = "USERNAME" ' <C27
Private Const FIELD_DATE_PRINTED = "DATE PRINTED" ' <C28
Private Const FIELD_PRINTED_BY = "PRINTED BY" ' <C29
Private Const FIELD_MRN = "MRN" ' <C30
Private Const FIELD_REMOTE_ID = "REMOTE_ID" ' <C30

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
              "[DATA_NCTS_ID],[DATA_NCTS_MSG_ID],[CODE]" & _
              ",[LOGID DESCRIPTION],[TYPE],[COMM]" & _
              ",[USER NO],[LAST MODIFIED BY],[PRINT]" & _
              ",[DOCUMENT NAME],[DATE CREATED],[DATE LAST MODIFIED]" & _
              ",[DATE REQUESTED],[DATE SEND],[HEADER]" & _
              ",[DETAIL],[DOCUMENT COUNTER],[REMARKS]" & _
              ",[TREE ID],[SUPPLIER NAME],[DTYPE]" & _
              ",[Error String],[Error HD],[LOGID]" & _
              ",[USERNAME],[DATE PRINTED],[PRINTED BY]" & _
              ",[MRN],[REMOTE_ID]) VALUES " ' <C31

Private Const SQL_GET_RECORD = "SELECT [DATA_NCTS_ID],[DATA_NCTS_MSG_ID],[CODE]" & _
              ",[LOGID DESCRIPTION],[TYPE],[COMM]" & _
              ",[USER NO],[LAST MODIFIED BY],[PRINT]" & _
              ",[DOCUMENT NAME],[DATE CREATED],[DATE LAST MODIFIED]" & _
              ",[DATE REQUESTED],[DATE SEND],[HEADER]" & _
              ",[DETAIL],[DOCUMENT COUNTER],[REMARKS]" & _
              ",[TREE ID],[SUPPLIER NAME],[DTYPE]" & _
              ",[Error String],[Error HD],[LOGID]" & _
              ",[USERNAME],[DATE PRINTED],[PRINTED BY]" & _
              ",[MRN], [REMOTE_ID] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = " ' <C32

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & _
              " WHERE [" & PK_FIELD & "] = " ' <C33

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] " & _
              "FROM [" & TABLE_NAME & "]" ' <C34

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] " & _
              "FROM [" & TABLE_NAME & "]" ' <C35

' Code: >M1
' Name: AddRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

   Dim strSql As String

   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_COMM) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USER_NO) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LAST_MODIFIED_BY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_NAME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_CREATED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_MODIFIED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_REQUESTED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_SEND) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_HEADER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DETAIL) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_COUNTER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_REMARKS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TREE_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SUPPLIER_NAME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DTYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Error_String) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Error_HD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USERNAME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATE_PRINTED) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINTED_BY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_MRN) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID) & ")"

      On Error GoTo ERROR_QUERY

      ExecuteNonQuery ActiveConnection, strSql


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
Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID)


      On Error GoTo ERROR_QUERY

      ExecuteNonQuery ActiveConnection, strSql


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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [DATA_NCTS_MSG_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID) & ", "
      strSql = strSql & " [CODE]=" & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ", "
      strSql = strSql & " [LOGID DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION) & ", "
      strSql = strSql & " [TYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_TYPE) & ", "
      strSql = strSql & " [COMM]=" & SetEnclosedChr(ActiveRecord.FIELD_COMM) & ", "
      strSql = strSql & " [USER NO]=" & SetEnclosedChr(ActiveRecord.FIELD_USER_NO) & ", "
      strSql = strSql & " [LAST MODIFIED BY]=" & SetEnclosedChr(ActiveRecord.FIELD_LAST_MODIFIED_BY) & ", "
      strSql = strSql & " [PRINT]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT) & ", "
      strSql = strSql & " [DOCUMENT NAME]=" & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_NAME) & ", "
      strSql = strSql & " [DATE CREATED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_CREATED) & ", "
      strSql = strSql & " [DATE LAST MODIFIED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_LAST_MODIFIED) & ", "
      strSql = strSql & " [DATE REQUESTED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_REQUESTED) & ", "
      strSql = strSql & " [DATE SEND]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_SEND) & ", "
      strSql = strSql & " [HEADER]=" & SetEnclosedChr(ActiveRecord.FIELD_HEADER) & ", "
      strSql = strSql & " [DETAIL]=" & SetEnclosedChr(ActiveRecord.FIELD_DETAIL) & ", "
      strSql = strSql & " [DOCUMENT COUNTER]=" & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT_COUNTER) & ", "
      strSql = strSql & " [REMARKS]=" & SetEnclosedChr(ActiveRecord.FIELD_REMARKS) & ", "
      strSql = strSql & " [TREE ID]=" & SetEnclosedChr(ActiveRecord.FIELD_TREE_ID) & ", "
      strSql = strSql & " [SUPPLIER NAME]=" & SetEnclosedChr(ActiveRecord.FIELD_SUPPLIER_NAME) & ", "
      strSql = strSql & " [DTYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_DTYPE) & ", "
      strSql = strSql & " [Error String]=" & SetEnclosedChr(ActiveRecord.FIELD_Error_String) & ", "
      strSql = strSql & " [Error HD]=" & SetEnclosedChr(ActiveRecord.FIELD_Error_HD) & ", "
      strSql = strSql & " [LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_LOGID) & ", "
      strSql = strSql & " [USERNAME]=" & SetEnclosedChr(ActiveRecord.FIELD_USERNAME) & ", "
      strSql = strSql & " [DATE PRINTED]=" & SetEnclosedChr(ActiveRecord.FIELD_DATE_PRINTED) & ", "
      strSql = strSql & " [PRINTED BY]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINTED_BY) & ", "
      strSql = strSql & " [MRN]=" & SetEnclosedChr(ActiveRecord.FIELD_MRN) & ", "
      strSql = strSql & " [REMOTE_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_REMOTE_ID) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID)

      On Error GoTo ERROR_QUERY

      ExecuteNonQuery ActiveConnection, strSql


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
Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID)

      On Error GoTo ERROR_QUERY
      
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_ID = FNullField(rstRecord.Fields("DATA_NCTS_ID"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(rstRecord.Fields("DATA_NCTS_MSG_ID"))
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("CODE"))
      ActiveRecord.FIELD_LOGID_DESCRIPTION = FNullField(rstRecord.Fields("LOGID DESCRIPTION"))
      ActiveRecord.FIELD_TYPE = FNullField(rstRecord.Fields("TYPE"))
      ActiveRecord.FIELD_COMM = FNullField(rstRecord.Fields("COMM"))
      ActiveRecord.FIELD_USER_NO = FNullField(rstRecord.Fields("USER NO"))
      ActiveRecord.FIELD_LAST_MODIFIED_BY = FNullField(rstRecord.Fields("LAST MODIFIED BY"))
      ActiveRecord.FIELD_PRINT = FNullField(rstRecord.Fields("PRINT"))
      ActiveRecord.FIELD_DOCUMENT_NAME = FNullField(rstRecord.Fields("DOCUMENT NAME"))
      ActiveRecord.FIELD_DATE_CREATED = FNullField(rstRecord.Fields("DATE CREATED"))
      ActiveRecord.FIELD_DATE_LAST_MODIFIED = FNullField(rstRecord.Fields("DATE LAST MODIFIED"))
      ActiveRecord.FIELD_DATE_REQUESTED = FNullField(rstRecord.Fields("DATE REQUESTED"))
      ActiveRecord.FIELD_DATE_SEND = FNullField(rstRecord.Fields("DATE SEND"))
      ActiveRecord.FIELD_HEADER = FNullField(rstRecord.Fields("HEADER"))
      ActiveRecord.FIELD_DETAIL = FNullField(rstRecord.Fields("DETAIL"))
      ActiveRecord.FIELD_DOCUMENT_COUNTER = FNullField(rstRecord.Fields("DOCUMENT COUNTER"))
      ActiveRecord.FIELD_REMARKS = FNullField(rstRecord.Fields("REMARKS"))
      ActiveRecord.FIELD_TREE_ID = FNullField(rstRecord.Fields("TREE ID"))
      ActiveRecord.FIELD_SUPPLIER_NAME = FNullField(rstRecord.Fields("SUPPLIER NAME"))
      ActiveRecord.FIELD_DTYPE = FNullField(rstRecord.Fields("DTYPE"))
      ActiveRecord.FIELD_Error_String = FNullField(rstRecord.Fields("Error String"))
      ActiveRecord.FIELD_Error_HD = FNullField(rstRecord.Fields("Error HD"))
      ActiveRecord.FIELD_LOGID = FNullField(rstRecord.Fields("LOGID"))
      ActiveRecord.FIELD_USERNAME = FNullField(rstRecord.Fields("USERNAME"))
      ActiveRecord.FIELD_DATE_PRINTED = FNullField(rstRecord.Fields("DATE PRINTED"))
      ActiveRecord.FIELD_PRINTED_BY = FNullField(rstRecord.Fields("PRINTED BY"))
      ActiveRecord.FIELD_MRN = FNullField(rstRecord.Fields("MRN"))
      ActiveRecord.FIELD_REMOTE_ID = FNullField(rstRecord.Fields("REMOTE_ID"))
      ADORecordsetClose rstRecord

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


   Exit Function


ERROR_RECORDSET:
       ADORecordsetClose rstRecord
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
Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDataNctsTable

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_ID = rstRecord.Fields("ID_MAX").Value
      ADORecordsetClose rstRecord

      GetMaxID = True

      Exit Function


ERROR_RECORDSET:

   ADORecordsetClose rstRecord

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
Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsTable) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDataNctsTable

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_ID = rstRecord.Fields("ID_MIN").Value
      ADORecordsetClose rstRecord

      GetMinID = True

      Exit Function


ERROR_RECORDSET:

   ADORecordsetClose rstRecord

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
Public Function GetTableRecord(ByRef ActiveRecord As cpiDataNctsTable) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_DATA_NCTS_ID, 3, 4, 90 ' 1
   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_ID, 3, 4, 118 ' 2
   rstRecord.Fields.Append FIELD_CODE, 202, 21, 102 ' 3
   rstRecord.Fields.Append FIELD_LOGID_DESCRIPTION, 202, 40, 102 ' 4
   rstRecord.Fields.Append FIELD_TYPE, 202, 1, 102 ' 5
   rstRecord.Fields.Append FIELD_COMM, 202, 1, 102 ' 6
   rstRecord.Fields.Append FIELD_USER_NO, 3, 4, 118 ' 7
   rstRecord.Fields.Append FIELD_LAST_MODIFIED_BY, 202, 25, 102 ' 8
   rstRecord.Fields.Append FIELD_PRINT, 202, 1, 102 ' 9
   rstRecord.Fields.Append FIELD_DOCUMENT_NAME, 202, 40, 102 ' 10
   rstRecord.Fields.Append FIELD_DATE_CREATED, 7, 8, 118 ' 11
   rstRecord.Fields.Append FIELD_DATE_LAST_MODIFIED, 7, 8, 118 ' 12
   rstRecord.Fields.Append FIELD_DATE_REQUESTED, 7, 8, 118 ' 13
   rstRecord.Fields.Append FIELD_DATE_SEND, 7, 8, 118 ' 14
   rstRecord.Fields.Append FIELD_HEADER, 2, 2, 118 ' 15
   rstRecord.Fields.Append FIELD_DETAIL, 2, 2, 118 ' 16
   rstRecord.Fields.Append FIELD_DOCUMENT_COUNTER, 2, 2, 118 ' 17
   rstRecord.Fields.Append FIELD_REMARKS, 202, 30, 102 ' 18
   rstRecord.Fields.Append FIELD_TREE_ID, 202, 10, 70 ' 19
   rstRecord.Fields.Append FIELD_SUPPLIER_NAME, 202, 40, 102 ' 20
   rstRecord.Fields.Append FIELD_DTYPE, 17, 1, 118 ' 21
   rstRecord.Fields.Append FIELD_Error_String, 203, 536870910, 230 ' 22
   rstRecord.Fields.Append FIELD_Error_HD, 202, 10, 102 ' 23
   rstRecord.Fields.Append FIELD_LOGID, 202, 4, 102 ' 24
   rstRecord.Fields.Append FIELD_USERNAME, 202, 25, 102 ' 25
   rstRecord.Fields.Append FIELD_DATE_PRINTED, 7, 8, 118 ' 26
   rstRecord.Fields.Append FIELD_PRINTED_BY, 202, 25, 102 ' 27
   rstRecord.Fields.Append FIELD_MRN, 202, 25, 102 ' 28
   rstRecord.Fields.Append FIELD_REMOTE_ID, 3, 4, 118 ' 28
   
   ' set values

   rstRecord.AddNew
   rstRecord.Fields("DATA_NCTS_ID") = ActiveRecord.FIELD_DATA_NCTS_ID
   rstRecord.Fields("DATA_NCTS_MSG_ID") = ActiveRecord.FIELD_DATA_NCTS_MSG_ID
   rstRecord.Fields("CODE") = ActiveRecord.FIELD_CODE
   rstRecord.Fields("LOGID DESCRIPTION") = ActiveRecord.FIELD_LOGID_DESCRIPTION
   rstRecord.Fields("TYPE") = ActiveRecord.FIELD_TYPE
   rstRecord.Fields("COMM") = ActiveRecord.FIELD_COMM
   rstRecord.Fields("USER NO") = ActiveRecord.FIELD_USER_NO
   rstRecord.Fields("LAST MODIFIED BY") = ActiveRecord.FIELD_LAST_MODIFIED_BY
   rstRecord.Fields("PRINT") = ActiveRecord.FIELD_PRINT
   rstRecord.Fields("DOCUMENT NAME") = ActiveRecord.FIELD_DOCUMENT_NAME
   rstRecord.Fields("DATE CREATED") = ActiveRecord.FIELD_DATE_CREATED
   rstRecord.Fields("DATE LAST MODIFIED") = ActiveRecord.FIELD_DATE_LAST_MODIFIED
   rstRecord.Fields("DATE REQUESTED") = ActiveRecord.FIELD_DATE_REQUESTED
   rstRecord.Fields("DATE SEND") = ActiveRecord.FIELD_DATE_SEND
   rstRecord.Fields("HEADER") = ActiveRecord.FIELD_HEADER
   rstRecord.Fields("DETAIL") = ActiveRecord.FIELD_DETAIL
   rstRecord.Fields("DOCUMENT COUNTER") = ActiveRecord.FIELD_DOCUMENT_COUNTER
   rstRecord.Fields("REMARKS") = ActiveRecord.FIELD_REMARKS
   rstRecord.Fields("TREE ID") = ActiveRecord.FIELD_TREE_ID
   rstRecord.Fields("SUPPLIER NAME") = ActiveRecord.FIELD_SUPPLIER_NAME
   rstRecord.Fields("DTYPE") = ActiveRecord.FIELD_DTYPE
   rstRecord.Fields("Error String") = ActiveRecord.FIELD_Error_String
   rstRecord.Fields("Error HD") = ActiveRecord.FIELD_Error_HD
   rstRecord.Fields("LOGID") = ActiveRecord.FIELD_LOGID
   rstRecord.Fields("USERNAME") = ActiveRecord.FIELD_USERNAME
   rstRecord.Fields("DATE PRINTED") = ActiveRecord.FIELD_DATE_PRINTED
   rstRecord.Fields("PRINTED BY") = ActiveRecord.FIELD_PRINTED_BY
   rstRecord.Fields("MRN") = ActiveRecord.FIELD_MRN
   rstRecord.Fields("REMOTE_ID") = ActiveRecord.FIELD_REMOTE_ID
   
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
Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiDataNctsTable

   Dim clsRecord As cpiDataNctsTable
   Set clsRecord = New cpiDataNctsTable

   clsRecord.FIELD_DATA_NCTS_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_ID))
   clsRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_ID))
   clsRecord.FIELD_CODE = FNullField(ActiveRecord.Fields(FIELD_CODE))
   clsRecord.FIELD_LOGID_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_LOGID_DESCRIPTION))
   clsRecord.FIELD_TYPE = FNullField(ActiveRecord.Fields(FIELD_TYPE))
   clsRecord.FIELD_COMM = FNullField(ActiveRecord.Fields(FIELD_COMM))
   clsRecord.FIELD_USER_NO = FNullField(ActiveRecord.Fields(FIELD_USER_NO))
   clsRecord.FIELD_LAST_MODIFIED_BY = FNullField(ActiveRecord.Fields(FIELD_LAST_MODIFIED_BY))
   clsRecord.FIELD_PRINT = FNullField(ActiveRecord.Fields(FIELD_PRINT))
   clsRecord.FIELD_DOCUMENT_NAME = FNullField(ActiveRecord.Fields(FIELD_DOCUMENT_NAME))
   clsRecord.FIELD_DATE_CREATED = FNullField(ActiveRecord.Fields(FIELD_DATE_CREATED))
   clsRecord.FIELD_DATE_LAST_MODIFIED = FNullField(ActiveRecord.Fields(FIELD_DATE_LAST_MODIFIED))
   clsRecord.FIELD_DATE_REQUESTED = FNullField(ActiveRecord.Fields(FIELD_DATE_REQUESTED))
   clsRecord.FIELD_DATE_SEND = FNullField(ActiveRecord.Fields(FIELD_DATE_SEND))
   clsRecord.FIELD_HEADER = FNullField(ActiveRecord.Fields(FIELD_HEADER))
   clsRecord.FIELD_DETAIL = FNullField(ActiveRecord.Fields(FIELD_DETAIL))
   clsRecord.FIELD_DOCUMENT_COUNTER = FNullField(ActiveRecord.Fields(FIELD_DOCUMENT_COUNTER))
   clsRecord.FIELD_REMARKS = FNullField(ActiveRecord.Fields(FIELD_REMARKS))
   clsRecord.FIELD_TREE_ID = FNullField(ActiveRecord.Fields(FIELD_TREE_ID))
   clsRecord.FIELD_SUPPLIER_NAME = FNullField(ActiveRecord.Fields(FIELD_SUPPLIER_NAME))
   clsRecord.FIELD_DTYPE = FNullField(ActiveRecord.Fields(FIELD_DTYPE))
   clsRecord.FIELD_Error_String = FNullField(ActiveRecord.Fields(FIELD_Error_String))
   clsRecord.FIELD_Error_HD = FNullField(ActiveRecord.Fields(FIELD_Error_HD))
   clsRecord.FIELD_LOGID = FNullField(ActiveRecord.Fields(FIELD_LOGID))
   clsRecord.FIELD_USERNAME = FNullField(ActiveRecord.Fields(FIELD_USERNAME))
   clsRecord.FIELD_DATE_PRINTED = FNullField(ActiveRecord.Fields(FIELD_DATE_PRINTED))
   clsRecord.FIELD_PRINTED_BY = FNullField(ActiveRecord.Fields(FIELD_PRINTED_BY))
   clsRecord.FIELD_MRN = FNullField(ActiveRecord.Fields(FIELD_MRN))
   clsRecord.FIELD_REMOTE_ID = FNullField(ActiveRecord.Fields(FIELD_REMOTE_ID))
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
      CommandText = SQL_GET_RECORD
   End If

   
   ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic

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
Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiDataNctsTable

    'create a new object
    Dim objNewMember As cpiDataNctsTable
    Set objNewMember = New cpiDataNctsTable

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
Public Property Get Item(vntIndexKey As Variant) As cpiDataNctsTable
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

