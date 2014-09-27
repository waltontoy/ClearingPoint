VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDataNctsMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************************************
' ************************************************************************************************
' Program Name: cpiDataNctsMessages
' Type: Class Module
' Instancing: 5 - MultiUse
' Author: Jason B. Sia
' Date Documented: Wednesday, October 01, 2003
' Description: Collection of cpiDataNctsMessage class
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

Private Const TABLE_NAME = "DATA_NCTS_MESSAGES" '<C1
Private Const PK_FIELD = "DATA_NCTS_MSG_ID" ' <C2

' field constants
Private Const FIELD_DATA_NCTS_MSG_ID = "DATA_NCTS_MSG_ID" ' <C3
Private Const FIELD_NCTS_IEM_ID = "NCTS_IEM_ID" ' <C4
Private Const FIELD_DATA_NCTS_ID = "DATA_NCTS_ID" ' <C5
Private Const FIELD_DATA_NCTS_MSG_Date = "DATA_NCTS_MSG_Date" ' <C6
Private Const FIELD_DATA_NCTS_MSG_StatusType = "DATA_NCTS_MSG_StatusType" ' <C7
Private Const FIELD_User_ID = "User_ID" ' <C8
Private Const FIELD_DATA_NCTS_MSG_Message = "DATA_NCTS_MSG_Message" ' <C9
Private Const FIELD_DATA_NCTS_MSG_Reference = "DATA_NCTS_MSG_Reference" ' <C10

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
              "[DATA_NCTS_MSG_ID],[NCTS_IEM_ID],[DATA_NCTS_ID]" & _
              ",[DATA_NCTS_MSG_Date],[DATA_NCTS_MSG_StatusType],[User_ID]" & _
              ",[DATA_NCTS_MSG_Message],[DATA_NCTS_MSG_Reference]) VALUES " ' <C11

Private Const SQL_GET_RECORD = "SELECT [DATA_NCTS_MSG_ID],[NCTS_IEM_ID],[DATA_NCTS_ID]" & _
              ",[DATA_NCTS_MSG_Date],[DATA_NCTS_MSG_StatusType],[User_ID]" & _
              ",[DATA_NCTS_MSG_Message],[DATA_NCTS_MSG_Reference] FROM " & TABLE_NAME & _
              " WHERE [" & PK_FIELD & "] = " ' <C12

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & _
              " WHERE [" & PK_FIELD & "] = " ' <C13

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] " & _
              "FROM [" & TABLE_NAME & "]" ' <C14

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] " & _
              "FROM [" & TABLE_NAME & "]" ' <C15

' Code: >M1
' Name: AddRecord
' Description: Returns created recordset based on query and connection passed.
' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
' Scope: Public
' Fan-In: <none>
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Date) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_StatusType) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_User_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Message) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Reference) & ")"

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
Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID)


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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [NCTS_IEM_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID) & ", "
      strSql = strSql & " [DATA_NCTS_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_ID) & ", "
      strSql = strSql & " [DATA_NCTS_MSG_Date]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Date) & ", "
      strSql = strSql & " [DATA_NCTS_MSG_StatusType]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_StatusType) & ", "
      strSql = strSql & " [User_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_User_ID) & ", "
      strSql = strSql & " [DATA_NCTS_MSG_Message]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Message) & ", "
      strSql = strSql & " [DATA_NCTS_MSG_Reference]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_Reference) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID)

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
Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(rstRecord.Fields("DATA_NCTS_MSG_ID"))
      ActiveRecord.FIELD_NCTS_IEM_ID = FNullField(rstRecord.Fields("NCTS_IEM_ID"))
      ActiveRecord.FIELD_DATA_NCTS_ID = FNullField(rstRecord.Fields("DATA_NCTS_ID"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_Date = FNullField(rstRecord.Fields("DATA_NCTS_MSG_Date"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_StatusType = FNullField(rstRecord.Fields("DATA_NCTS_MSG_StatusType"))
      ActiveRecord.FIELD_User_ID = FNullField(rstRecord.Fields("User_ID"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_Message = FNullField(rstRecord.Fields("DATA_NCTS_MSG_Message"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_Reference = FNullField(rstRecord.Fields("DATA_NCTS_MSG_Reference"))
      
      
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
Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDataNctsMessage

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_MSG_ID = rstRecord.Fields("ID_MAX").Value
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
Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDataNctsMessage) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDataNctsMessage

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_MSG_ID = rstRecord.Fields("ID_MIN").Value
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiDataNctsMessage) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_ID, 3, 4, 90 ' 1
   rstRecord.Fields.Append FIELD_NCTS_IEM_ID, 3, 4, 118 ' 2
   rstRecord.Fields.Append FIELD_DATA_NCTS_ID, 3, 4, 118 ' 3
   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_Date, 7, 8, 118 ' 4
   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_StatusType, 202, 15, 102 ' 5
   rstRecord.Fields.Append FIELD_User_ID, 3, 4, 118 ' 6
   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_Message, 203, 536870910, 230 ' 7
   rstRecord.Fields.Append FIELD_DATA_NCTS_MSG_Reference, 202, 22, 102 ' 8

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("DATA_NCTS_MSG_ID") = ActiveRecord.FIELD_DATA_NCTS_MSG_ID
   rstRecord.Fields("NCTS_IEM_ID") = ActiveRecord.FIELD_NCTS_IEM_ID
   rstRecord.Fields("DATA_NCTS_ID") = ActiveRecord.FIELD_DATA_NCTS_ID
   rstRecord.Fields("DATA_NCTS_MSG_Date") = ActiveRecord.FIELD_DATA_NCTS_MSG_Date
   rstRecord.Fields("DATA_NCTS_MSG_StatusType") = ActiveRecord.FIELD_DATA_NCTS_MSG_StatusType
   rstRecord.Fields("User_ID") = ActiveRecord.FIELD_User_ID
   rstRecord.Fields("DATA_NCTS_MSG_Message") = ActiveRecord.FIELD_DATA_NCTS_MSG_Message
   rstRecord.Fields("DATA_NCTS_MSG_Reference") = ActiveRecord.FIELD_DATA_NCTS_MSG_Reference

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
Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiDataNctsMessage

   Dim clsRecord As cpiDataNctsMessage
   Set clsRecord = New cpiDataNctsMessage

   clsRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_ID))
   clsRecord.FIELD_NCTS_IEM_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_ID))
   clsRecord.FIELD_DATA_NCTS_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_ID))
   clsRecord.FIELD_DATA_NCTS_MSG_Date = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_Date))
   clsRecord.FIELD_DATA_NCTS_MSG_StatusType = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_StatusType))
   clsRecord.FIELD_User_ID = FNullField(ActiveRecord.Fields(FIELD_User_ID))
   clsRecord.FIELD_DATA_NCTS_MSG_Message = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_Message))
   clsRecord.FIELD_DATA_NCTS_MSG_Reference = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_Reference))
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
Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiDataNctsMessage

    'create a new object
    Dim objNewMember As cpiDataNctsMessage
    Set objNewMember = New cpiDataNctsMessage

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
Public Property Get Item(vntIndexKey As Variant) As cpiDataNctsMessage
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

