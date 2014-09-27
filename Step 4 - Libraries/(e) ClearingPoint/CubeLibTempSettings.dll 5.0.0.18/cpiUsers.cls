VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Private Const TABLE_NAME = "Users"
Private Const PK_FIELD = "User_ID"
    
' field constants
Private Const FIELD_User_ID = "User_ID"
Private Const FIELD_User_Name = "User_Name"
Private Const FIELD_User_Password = "User_Password"
Private Const FIELD_User_Rights = "User_Rights"
Private Const FIELD_User_StateSettings = "User_StateSettings"
Private Const FIELD_User_MainDimensionPosition = "User_MainDimensionPosition"
Private Const FIELD_User_TreeWidth = "User_TreeWidth"
Private Const FIELD_User_Level = "User_Level"
Private Const FIELD_User_ButtonGroup = "User_ButtonGroup"
Private Const FIELD_User_ButtonBarWidth = "User_ButtonBarWidth"

Private Const FIELD_ADMINISTRATOR_RIGHTS = "ADMINISTRATOR RIGHTS"
Private Const FIELD_MAINTAIN_TABLES = "MAINTAIN TABLES"
Private Const FIELD_ALL_LOGICAL_IDS = "ALL LOGICAL IDS"
Private Const FIELD_SHOW_ALL_SENT = "SHOW ALL SENT"
Private Const FIELD_SHOW_ALL_WITH_ERRORS = "SHOW ALL WITH ERRORS"
Private Const FIELD_SHOW_ALL_WAITING = "SHOW ALL WAITING"
Private Const FIELD_SHOW_ALL_DELETED = "SHOW ALL DELETED"
Private Const FIELD_CLEAN_UP_DELETED = "CLEAN UP DELETED"
Private Const FIELD_EVERY_FIELD = "EVERY"
Private Const FIELD_DAYS_OR_ITEMS = "DAYS OR ITEMS"
Private Const FIELD_DELETE_OTHER_USERS_ITEMS = "DELETE OTHER USERS ITEMS"
Private Const FIELD_DELETE_SENT_IN_N_DAYS = "DELETE SENT IN N DAYS"
Private Const FIELD_WITH_SECURITY = "WITH SECURITY"
Private Const FIELD_REFRESH_IN_SECONDS = "REFRESH IN SECONDS"
Private Const FIELD_RELATE_L1_TO_S1 = "RELATE L1 TO S1"
Private Const FIELD_SUPPLIER_BOX = "SUPPLIER BOX"
Private Const FIELD_LAST_USED_PRINTER = "LAST USED PRINTER"
Private Const FIELD_LAST_DV_PRINTER = "LAST DV PRINTER"
Private Const FIELD_LOGID_DESCRIPTION = "LOGID DESCRIPTION"
Private Const FIELD_FIXED_USER = "FIXED USER"
Private Const FIELD_LOGGED = "LOGGED"
Private Const FIELD_SDICOUNT = "SDICOUNT"
Private Const FIELD_SDECOUNT = "SDECOUNT"
Private Const FIELD_SDTCOUNT = "SDTCOUNT"
Private Const FIELD_DECOUNT = "DECOUNT"
Private Const FIELD_SHOW_ONLY_DOCTYPE = "SHOW ONLY DOCTYPE"
Private Const FIELD_SHOW_ONLY_VATNUM = "SHOW ONLY VATNUM"
Private Const FIELD_SHOW_ONLY_CTRYCODE = "SHOW ONLY CTRYCODE"
Private Const FIELD_SAVE_NEW_TARICCTRY = "SAVE NEW TARICCTRY"
Private Const FIELD_SHOW_ALL_TOBEPRINTED = "SHOW ALL TOBEPRINTED"
Private Const FIELD_SHOW_ALL_DRAFTS = "SHOW ALL DRAFTS"
Private Const FIELD_SDI2COUNT = "SDI2COUNT"
Private Const FIELD_SDE2COUNT = "SDE2COUNT"
Private Const FIELD_SDT2COUNT = "SDT2COUNT"
Private Const FIELD_LANGUAGE_FIELD = "LANGUAGE"

Private Const SQL_ADD_RECORD = "INSERT INTO [Users] ([User_ID], [User_Name], [User_Password], [User_Rights], " & _
                            "[User_StateSettings], [User_MainDimensionPosition], [User_TreeWidth], [User_Level], [User_ButtonGroup], " & _
                            "[User_ButtonBarWidth] " & _
                            ",[ADMINISTRATOR RIGHTS],[MAINTAIN TABLES],[ALL LOGICAL IDS],[SHOW ALL SENT],[SHOW ALL WITH ERRORS] " & _
                            ",[SHOW ALL WAITING],[SHOW ALL DELETED],[CLEAN UP DELETED],[EVERY],[DAYS OR ITEMS] " & _
                            ",[DELETE OTHER USERS ITEMS],[DELETE SENT IN N DAYS],[WITH SECURITY],[REFRESH IN SECONDS],[RELATE L1 TO S1] " & _
                            ",[SUPPLIER BOX],[LAST USED PRINTER],[LAST DV PRINTER],[LOGID DESCRIPTION],[FIXED USER] " & _
                            ",[LOGGED],[SDICOUNT],[SDECOUNT],[SDTCOUNT],[DECOUNT] " & _
                            ",[SHOW ONLY DOCTYPE],[SHOW ONLY VATNUM],[SHOW ONLY CTRYCODE],[SAVE NEW TARICCTRY],[SHOW ALL TOBEPRINTED] " & _
                            ",[SHOW ALL DRAFTS],[SDI2COUNT],[SDE2COUNT],[SDT2COUNT],[LANGUAGE] " & _
                            ") VALUES "

Private Const SQL_DELETE_RECORD = "DELETE * FROM Users WHERE [User_ID]="

'Private Const SQL_MODIFY_RECORD '=

Private Const SQL_GET_RECORD = "SELECT [User_ID], [User_Name], [User_Password], [User_Rights], " & _
                            "[User_StateSettings], [User_MainDimensionPosition], [User_TreeWidth], [User_Level], [User_ButtonGroup], " & _
                            "[User_ButtonBarWidth] " & _
                            ",[ADMINISTRATOR RIGHTS],[MAINTAIN TABLES],[ALL LOGICAL IDS],[SHOW ALL SENT],[SHOW ALL WITH ERRORS] " & _
                            ",[SHOW ALL WAITING],[SHOW ALL DELETED],[CLEAN UP DELETED],[EVERY],[DAYS OR ITEMS] " & _
                            ",[DELETE OTHER USERS ITEMS],[DELETE SENT IN N DAYS],[WITH SECURITY],[REFRESH IN SECONDS],[RELATE L1 TO S1] " & _
                            ",[SUPPLIER BOX],[LAST USED PRINTER],[LAST DV PRINTER],[LOGID DESCRIPTION],[FIXED USER] " & _
                            ",[LOGGED],[SDICOUNT],[SDECOUNT],[SDTCOUNT],[DECOUNT] " & _
                            ",[SHOW ONLY DOCTYPE],[SHOW ONLY VATNUM],[SHOW ONLY CTRYCODE],[SAVE NEW TARICCTRY],[SHOW ALL TOBEPRINTED] " & _
                            ",[SHOW ALL DRAFTS],[SDI2COUNT],[SDE2COUNT],[SDT2COUNT],[LANGUAGE] " & _
                            "FROM Users WHERE [User_ID]= "

Private Const SQL_GET_MAXID = "SELECT MAX(User_ID) AS [ID_MAX] FROM Users"
Private Const SQL_GET_MINID = "SELECT Min(User_ID) AS [ID_MIN] FROM Users"

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & _
                        "(" & ActiveRecord.User_ID & ",'" & ActiveRecord.User_Name & "','" & ActiveRecord.User_Password & _
                        "'," & ActiveRecord.User_Rights & "," & ActiveRecord.User_StateSettings & _
                        ",'" & ActiveRecord.User_MainDimensionPosition & _
                        "'," & ActiveRecord.User_TreeWidth & "," & ActiveRecord.User_Level & "," & ActiveRecord.User_ButtonGroup & _
                        "," & ActiveRecord.User_ButtonBarWidth & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ADMINISTRATOR_RIGHTS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.MAINTAIN_TABLES) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ALL_LOGICAL_IDS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_SENT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_WITH_ERRORS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_WAITING) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_DELETED) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.CLEAN_UP_DELETED) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.EVERY_FIELD) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DAYS_OR_ITEMS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DELETE_OTHER_USERS_ITEMS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DELETE_SENT_IN_N_DAYS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.WITH_SECURITY) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.REFRESH_IN_SECONDS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.RELATE_L1_TO_S1) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SUPPLIER_BOX) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.LAST_USED_PRINTER) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.LAST_DV_PRINTER) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.LOGID_DESCRIPTION) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.FIXED_USER) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.LOGGED) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDICOUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDECOUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDTCOUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DECOUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ONLY_DOCTYPE) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ONLY_VATNUM) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ONLY_CTRYCODE) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SAVE_NEW_TARICCTRY) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_TOBEPRINTED) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SHOW_ALL_DRAFTS) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDI2COUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDE2COUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SDT2COUNT) & ""
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.LANGUAGE_FIELD) & ""
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

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & ActiveRecord.User_ID
    
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

' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET " & _
                            FIELD_User_Name & "='" & ActiveRecord.User_Name & "'," & _
                            FIELD_User_Password & "='" & ActiveRecord.User_Password & "'," & _
                            FIELD_User_Rights & "=" & ActiveRecord.User_Rights & "," & _
                            FIELD_User_StateSettings & "=" & ActiveRecord.User_StateSettings & "," & _
                            FIELD_User_MainDimensionPosition & "='" & ActiveRecord.User_MainDimensionPosition & "'," & _
                            FIELD_User_TreeWidth & "=" & ActiveRecord.User_TreeWidth & "," & _
                            FIELD_User_Level & "=" & ActiveRecord.User_Level & "," & _
                            FIELD_User_ButtonGroup & "=" & ActiveRecord.User_ButtonGroup & "," & _
                            FIELD_User_ButtonBarWidth & "=" & ActiveRecord.User_ButtonBarWidth
        
        strSql = strSql & "," & FIELD_ADMINISTRATOR_RIGHTS & "=" & SetEnclosedChr(ActiveRecord.ADMINISTRATOR_RIGHTS)
        strSql = strSql & "," & FIELD_MAINTAIN_TABLES & "=" & SetEnclosedChr(ActiveRecord.MAINTAIN_TABLES)
        strSql = strSql & "," & FIELD_ALL_LOGICAL_IDS & "=" & SetEnclosedChr(ActiveRecord.ALL_LOGICAL_IDS)
        strSql = strSql & "," & FIELD_SHOW_ALL_SENT & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_SENT)
        strSql = strSql & "," & FIELD_SHOW_ALL_WITH_ERRORS & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_WITH_ERRORS)
        strSql = strSql & "," & FIELD_SHOW_ALL_WAITING & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_WAITING)
        strSql = strSql & "," & FIELD_SHOW_ALL_DELETED & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_DELETED)
        strSql = strSql & "," & FIELD_CLEAN_UP_DELETED & "=" & SetEnclosedChr(ActiveRecord.CLEAN_UP_DELETED)
        strSql = strSql & "," & FIELD_EVERY_FIELD & "=" & SetEnclosedChr(ActiveRecord.EVERY_FIELD)
        strSql = strSql & "," & FIELD_DAYS_OR_ITEMS & "=" & SetEnclosedChr(ActiveRecord.DAYS_OR_ITEMS)
        strSql = strSql & "," & FIELD_DELETE_OTHER_USERS_ITEMS & "=" & SetEnclosedChr(ActiveRecord.DELETE_OTHER_USERS_ITEMS)
        strSql = strSql & "," & FIELD_DELETE_SENT_IN_N_DAYS & "=" & SetEnclosedChr(ActiveRecord.DELETE_SENT_IN_N_DAYS)
        strSql = strSql & "," & FIELD_WITH_SECURITY & "=" & SetEnclosedChr(ActiveRecord.WITH_SECURITY)
        strSql = strSql & "," & FIELD_REFRESH_IN_SECONDS & "=" & SetEnclosedChr(ActiveRecord.REFRESH_IN_SECONDS)
        strSql = strSql & "," & FIELD_RELATE_L1_TO_S1 & "=" & SetEnclosedChr(ActiveRecord.RELATE_L1_TO_S1)
        strSql = strSql & "," & FIELD_SUPPLIER_BOX & "=" & SetEnclosedChr(ActiveRecord.SUPPLIER_BOX)
        strSql = strSql & "," & FIELD_LAST_USED_PRINTER & "=" & SetEnclosedChr(ActiveRecord.LAST_USED_PRINTER)
        strSql = strSql & "," & FIELD_LAST_DV_PRINTER & "=" & SetEnclosedChr(ActiveRecord.LAST_DV_PRINTER)
        strSql = strSql & "," & FIELD_LOGID_DESCRIPTION & "=" & SetEnclosedChr(ActiveRecord.LOGID_DESCRIPTION)
        strSql = strSql & "," & FIELD_FIXED_USER & "=" & SetEnclosedChr(ActiveRecord.FIXED_USER)
        strSql = strSql & "," & FIELD_LOGGED & "=" & SetEnclosedChr(ActiveRecord.LOGGED)
        strSql = strSql & "," & FIELD_SDICOUNT & "=" & SetEnclosedChr(ActiveRecord.SDICOUNT)
        strSql = strSql & "," & FIELD_SDECOUNT & "=" & SetEnclosedChr(ActiveRecord.SDECOUNT)
        strSql = strSql & "," & FIELD_SDTCOUNT & "=" & SetEnclosedChr(ActiveRecord.SDTCOUNT)
        strSql = strSql & "," & FIELD_DECOUNT & "=" & SetEnclosedChr(ActiveRecord.DECOUNT)
        strSql = strSql & "," & FIELD_SHOW_ONLY_DOCTYPE & "=" & SetEnclosedChr(ActiveRecord.SHOW_ONLY_DOCTYPE)
        strSql = strSql & "," & FIELD_SHOW_ONLY_VATNUM & "=" & SetEnclosedChr(ActiveRecord.SHOW_ONLY_VATNUM)
        strSql = strSql & "," & FIELD_SHOW_ONLY_CTRYCODE & "=" & SetEnclosedChr(ActiveRecord.SHOW_ONLY_CTRYCODE)
        strSql = strSql & "," & FIELD_SAVE_NEW_TARICCTRY & "=" & SetEnclosedChr(ActiveRecord.SAVE_NEW_TARICCTRY)
        strSql = strSql & "," & FIELD_SHOW_ALL_TOBEPRINTED & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_TOBEPRINTED)
        strSql = strSql & "," & FIELD_SHOW_ALL_DRAFTS & "=" & SetEnclosedChr(ActiveRecord.SHOW_ALL_DRAFTS)
        strSql = strSql & "," & FIELD_SDI2COUNT & "=" & SetEnclosedChr(ActiveRecord.SDI2COUNT)
        strSql = strSql & "," & FIELD_SDE2COUNT & "=" & SetEnclosedChr(ActiveRecord.SDE2COUNT)
        strSql = strSql & "," & FIELD_SDT2COUNT & "=" & SetEnclosedChr(ActiveRecord.SDT2COUNT)
        strSql = strSql & "," & FIELD_LANGUAGE_FIELD & "=" & SetEnclosedChr(ActiveRecord.LANGUAGE_FIELD)
        
        strSql = strSql & " WHERE " & PK_FIELD & "=" & ActiveRecord.User_ID
    
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

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset
    
    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & ActiveRecord.User_ID
        
        On Error GoTo ERROR_QUERY
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        'Set rstRecord = ActiveConnection.Execute(strSql)
        
        On Error GoTo ERROR_RECORDSET
        ActiveRecord.User_Name = FNullField(rstRecord.Fields(FIELD_User_Name))
        ActiveRecord.User_Password = FNullField(rstRecord.Fields(FIELD_User_Password))
        ActiveRecord.User_Rights = FNullField(rstRecord.Fields(FIELD_User_Rights))
        ActiveRecord.User_StateSettings = FNullField(rstRecord.Fields(FIELD_User_StateSettings))
        ActiveRecord.User_MainDimensionPosition = FNullField(rstRecord.Fields(FIELD_User_MainDimensionPosition))
        ActiveRecord.User_TreeWidth = FNullField(rstRecord.Fields(FIELD_User_TreeWidth))
        ActiveRecord.User_Level = FNullField(rstRecord.Fields(FIELD_User_Level))
        ActiveRecord.User_ButtonGroup = FNullField(rstRecord.Fields(FIELD_User_ButtonGroup))
        ActiveRecord.User_ButtonBarWidth = FNullField(rstRecord.Fields(FIELD_User_ButtonBarWidth))
        
        ActiveRecord.ADMINISTRATOR_RIGHTS = FNullField(rstRecord.Fields(FIELD_ADMINISTRATOR_RIGHTS))
        ActiveRecord.MAINTAIN_TABLES = FNullField(rstRecord.Fields(FIELD_MAINTAIN_TABLES))
        ActiveRecord.ALL_LOGICAL_IDS = FNullField(rstRecord.Fields(FIELD_ALL_LOGICAL_IDS))
        ActiveRecord.SHOW_ALL_SENT = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_SENT))
        ActiveRecord.SHOW_ALL_WITH_ERRORS = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_WITH_ERRORS))
        ActiveRecord.SHOW_ALL_WAITING = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_WAITING))
        ActiveRecord.SHOW_ALL_DELETED = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_DELETED))
        ActiveRecord.CLEAN_UP_DELETED = FNullField(rstRecord.Fields(FIELD_CLEAN_UP_DELETED))
        ActiveRecord.EVERY_FIELD = FNullField(rstRecord.Fields(FIELD_EVERY_FIELD))
        ActiveRecord.DAYS_OR_ITEMS = FNullField(rstRecord.Fields(FIELD_DAYS_OR_ITEMS))
        ActiveRecord.DELETE_OTHER_USERS_ITEMS = FNullField(rstRecord.Fields(FIELD_DELETE_OTHER_USERS_ITEMS))
        ActiveRecord.DELETE_SENT_IN_N_DAYS = FNullField(rstRecord.Fields(FIELD_DELETE_SENT_IN_N_DAYS))
        ActiveRecord.WITH_SECURITY = FNullField(rstRecord.Fields(FIELD_WITH_SECURITY))
        ActiveRecord.REFRESH_IN_SECONDS = FNullField(rstRecord.Fields(FIELD_REFRESH_IN_SECONDS))
        ActiveRecord.RELATE_L1_TO_S1 = FNullField(rstRecord.Fields(FIELD_RELATE_L1_TO_S1))
        ActiveRecord.SUPPLIER_BOX = FNullField(rstRecord.Fields(FIELD_SUPPLIER_BOX))
        ActiveRecord.LAST_USED_PRINTER = FNullField(rstRecord.Fields(FIELD_LAST_USED_PRINTER))
        ActiveRecord.LAST_DV_PRINTER = FNullField(rstRecord.Fields(FIELD_LAST_DV_PRINTER))
        ActiveRecord.LOGID_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_LOGID_DESCRIPTION))
        ActiveRecord.FIXED_USER = FNullField(rstRecord.Fields(FIELD_FIXED_USER))
        ActiveRecord.LOGGED = FNullField(rstRecord.Fields(FIELD_LOGGED))
        ActiveRecord.SDICOUNT = FNullField(rstRecord.Fields(FIELD_SDICOUNT))
        ActiveRecord.SDECOUNT = FNullField(rstRecord.Fields(FIELD_SDECOUNT))
        ActiveRecord.SDTCOUNT = FNullField(rstRecord.Fields(FIELD_SDTCOUNT))
        ActiveRecord.DECOUNT = FNullField(rstRecord.Fields(FIELD_DECOUNT))
        ActiveRecord.SHOW_ONLY_DOCTYPE = FNullField(rstRecord.Fields(FIELD_SHOW_ONLY_DOCTYPE))
        ActiveRecord.SHOW_ONLY_VATNUM = FNullField(rstRecord.Fields(FIELD_SHOW_ONLY_VATNUM))
        ActiveRecord.SHOW_ONLY_CTRYCODE = FNullField(rstRecord.Fields(FIELD_SHOW_ONLY_CTRYCODE))
        ActiveRecord.SAVE_NEW_TARICCTRY = FNullField(rstRecord.Fields(FIELD_SAVE_NEW_TARICCTRY))
        ActiveRecord.SHOW_ALL_TOBEPRINTED = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_TOBEPRINTED))
        ActiveRecord.SHOW_ALL_DRAFTS = FNullField(rstRecord.Fields(FIELD_SHOW_ALL_DRAFTS))
        ActiveRecord.SDI2COUNT = FNullField(rstRecord.Fields(FIELD_SDI2COUNT))
        ActiveRecord.SDE2COUNT = FNullField(rstRecord.Fields(FIELD_SDE2COUNT))
        ActiveRecord.SDT2COUNT = FNullField(rstRecord.Fields(FIELD_SDT2COUNT))
        ActiveRecord.LANGUAGE_FIELD = FNullField(rstRecord.Fields(FIELD_LANGUAGE_FIELD))
        
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiUser
    End If

    strSql = SQL_GET_MAXID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.User_ID = rstRecord.Fields("ID_MAX").Value
    
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiUser) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiUser
    End If

    strSql = SQL_GET_MINID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.User_ID = rstRecord.Fields("ID_MIN").Value
    
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

' conversion
Public Function GetTableRecord(ByRef ActiveRecord As cpiUser) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset
    
    Set rstRecord = New ADODB.Recordset
    
    rstRecord.Open
    rstRecord.Fields.Append FIELD_User_ID, 3, 4, 90
    rstRecord.Fields.Append FIELD_User_Name, 202, 25, 70
    rstRecord.Fields.Append FIELD_User_Password, 202, 25, 70
    rstRecord.Fields.Append FIELD_User_Rights, 2, 2, 86
    rstRecord.Fields.Append FIELD_User_StateSettings, 3, 4, 86
    rstRecord.Fields.Append FIELD_User_MainDimensionPosition, 202, 75, 70
    rstRecord.Fields.Append FIELD_User_TreeWidth, 5, 8, 86
    rstRecord.Fields.Append FIELD_User_Level, 3, 4, 118
    rstRecord.Fields.Append FIELD_User_ButtonGroup, 3, 4, 118
    rstRecord.Fields.Append FIELD_User_ButtonBarWidth, 5, 8, 118
    
    rstRecord.Fields.Append FIELD_ADMINISTRATOR_RIGHTS, 11, 2, 86
    rstRecord.Fields.Append FIELD_MAINTAIN_TABLES, 11, 2, 86
    rstRecord.Fields.Append FIELD_ALL_LOGICAL_IDS, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ALL_SENT, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ALL_WITH_ERRORS, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ALL_WAITING, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ALL_DELETED, 11, 2, 86
    rstRecord.Fields.Append FIELD_CLEAN_UP_DELETED, 11, 2, 86
    rstRecord.Fields.Append FIELD_EVERY_FIELD, 2, 2, 118
    rstRecord.Fields.Append FIELD_DAYS_OR_ITEMS, 202, 1, 102
    rstRecord.Fields.Append FIELD_DELETE_OTHER_USERS_ITEMS, 11, 2, 86
    rstRecord.Fields.Append FIELD_DELETE_SENT_IN_N_DAYS, 2, 2, 118
    rstRecord.Fields.Append FIELD_WITH_SECURITY, 11, 2, 86
    rstRecord.Fields.Append FIELD_REFRESH_IN_SECONDS, 3, 4, 118
    rstRecord.Fields.Append FIELD_RELATE_L1_TO_S1, 11, 2, 86
    rstRecord.Fields.Append FIELD_SUPPLIER_BOX, 202, 2, 102
    rstRecord.Fields.Append FIELD_LAST_USED_PRINTER, 202, 150, 102
    rstRecord.Fields.Append FIELD_LAST_DV_PRINTER, 202, 50, 102
    rstRecord.Fields.Append FIELD_LOGID_DESCRIPTION, 202, 40, 102
    rstRecord.Fields.Append FIELD_FIXED_USER, 11, 2, 86
    rstRecord.Fields.Append FIELD_LOGGED, 11, 2, 86
    rstRecord.Fields.Append FIELD_SDICOUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_SDECOUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_SDTCOUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_DECOUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_SHOW_ONLY_DOCTYPE, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ONLY_VATNUM, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ONLY_CTRYCODE, 11, 2, 86
    rstRecord.Fields.Append FIELD_SAVE_NEW_TARICCTRY, 3, 4, 118
    rstRecord.Fields.Append FIELD_SHOW_ALL_TOBEPRINTED, 11, 2, 86
    rstRecord.Fields.Append FIELD_SHOW_ALL_DRAFTS, 11, 2, 86
    rstRecord.Fields.Append FIELD_SDI2COUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_SDE2COUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_SDT2COUNT, 3, 4, 118
    rstRecord.Fields.Append FIELD_LANGUAGE_FIELD, 17, 1, 118
    
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_User_ID) = ActiveRecord.User_ID
    rstRecord.Fields(FIELD_User_Name) = ActiveRecord.User_Name
    rstRecord.Fields(FIELD_User_Password) = ActiveRecord.User_Password
    rstRecord.Fields(FIELD_User_Rights) = ActiveRecord.User_Rights
    rstRecord.Fields(FIELD_User_StateSettings) = ActiveRecord.User_StateSettings
    rstRecord.Fields(FIELD_User_MainDimensionPosition) = ActiveRecord.User_MainDimensionPosition
    rstRecord.Fields(FIELD_User_TreeWidth) = ActiveRecord.User_TreeWidth
    rstRecord.Fields(FIELD_User_Level) = ActiveRecord.User_Level
    rstRecord.Fields(FIELD_User_ButtonGroup) = ActiveRecord.User_ButtonGroup
    rstRecord.Fields(FIELD_User_ButtonBarWidth) = ActiveRecord.User_ButtonBarWidth
    
    rstRecord.Fields(FIELD_ADMINISTRATOR_RIGHTS) = ActiveRecord.ADMINISTRATOR_RIGHTS
    rstRecord.Fields(FIELD_MAINTAIN_TABLES) = ActiveRecord.MAINTAIN_TABLES
    rstRecord.Fields(FIELD_ALL_LOGICAL_IDS) = ActiveRecord.ALL_LOGICAL_IDS
    rstRecord.Fields(FIELD_SHOW_ALL_SENT) = ActiveRecord.SHOW_ALL_SENT
    rstRecord.Fields(FIELD_SHOW_ALL_WITH_ERRORS) = ActiveRecord.SHOW_ALL_WITH_ERRORS
    rstRecord.Fields(FIELD_SHOW_ALL_WAITING) = ActiveRecord.SHOW_ALL_WAITING
    rstRecord.Fields(FIELD_SHOW_ALL_DELETED) = ActiveRecord.SHOW_ALL_DELETED
    rstRecord.Fields(FIELD_CLEAN_UP_DELETED) = ActiveRecord.CLEAN_UP_DELETED
    rstRecord.Fields(FIELD_EVERY_FIELD) = ActiveRecord.EVERY_FIELD
    rstRecord.Fields(FIELD_DAYS_OR_ITEMS) = ActiveRecord.DAYS_OR_ITEMS
    rstRecord.Fields(FIELD_DELETE_OTHER_USERS_ITEMS) = ActiveRecord.DELETE_OTHER_USERS_ITEMS
    rstRecord.Fields(FIELD_DELETE_SENT_IN_N_DAYS) = ActiveRecord.DELETE_SENT_IN_N_DAYS
    rstRecord.Fields(FIELD_WITH_SECURITY) = ActiveRecord.WITH_SECURITY
    rstRecord.Fields(FIELD_REFRESH_IN_SECONDS) = ActiveRecord.REFRESH_IN_SECONDS
    rstRecord.Fields(FIELD_RELATE_L1_TO_S1) = ActiveRecord.RELATE_L1_TO_S1
    rstRecord.Fields(FIELD_SUPPLIER_BOX) = ActiveRecord.SUPPLIER_BOX
    rstRecord.Fields(FIELD_LAST_USED_PRINTER) = ActiveRecord.LAST_USED_PRINTER
    rstRecord.Fields(FIELD_LAST_DV_PRINTER) = ActiveRecord.LAST_DV_PRINTER
    rstRecord.Fields(FIELD_LOGID_DESCRIPTION) = ActiveRecord.LOGID_DESCRIPTION
    rstRecord.Fields(FIELD_FIXED_USER) = ActiveRecord.FIXED_USER
    rstRecord.Fields(FIELD_LOGGED) = ActiveRecord.LOGGED
    rstRecord.Fields(FIELD_SDICOUNT) = ActiveRecord.SDICOUNT
    rstRecord.Fields(FIELD_SDECOUNT) = ActiveRecord.SDECOUNT
    rstRecord.Fields(FIELD_SDTCOUNT) = ActiveRecord.SDTCOUNT
    rstRecord.Fields(FIELD_DECOUNT) = ActiveRecord.DECOUNT
    rstRecord.Fields(FIELD_SHOW_ONLY_DOCTYPE) = ActiveRecord.SHOW_ONLY_DOCTYPE
    rstRecord.Fields(FIELD_SHOW_ONLY_VATNUM) = ActiveRecord.SHOW_ONLY_VATNUM
    rstRecord.Fields(FIELD_SHOW_ONLY_CTRYCODE) = ActiveRecord.SHOW_ONLY_CTRYCODE
    rstRecord.Fields(FIELD_SAVE_NEW_TARICCTRY) = ActiveRecord.SAVE_NEW_TARICCTRY
    rstRecord.Fields(FIELD_SHOW_ALL_TOBEPRINTED) = ActiveRecord.SHOW_ALL_TOBEPRINTED
    rstRecord.Fields(FIELD_SHOW_ALL_DRAFTS) = ActiveRecord.SHOW_ALL_DRAFTS
    rstRecord.Fields(FIELD_SDI2COUNT) = ActiveRecord.SDI2COUNT
    rstRecord.Fields(FIELD_SDE2COUNT) = ActiveRecord.SDE2COUNT
    rstRecord.Fields(FIELD_SDT2COUNT) = ActiveRecord.SDT2COUNT
    rstRecord.Fields(FIELD_LANGUAGE_FIELD) = ActiveRecord.LANGUAGE_FIELD
    
    rstRecord.Update
    
    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiUser

    Dim clsRecord As cpiUser
    
    Set clsRecord = New cpiUser
    
    clsRecord.User_ID = FNullField(ActiveRecord.Fields(FIELD_User_ID))
    clsRecord.User_Name = FNullField(ActiveRecord.Fields(FIELD_User_Name))
    clsRecord.User_Password = FNullField(ActiveRecord.Fields(FIELD_User_Password))
    clsRecord.User_Rights = FNullField(ActiveRecord.Fields(FIELD_User_Rights))
    clsRecord.User_StateSettings = FNullField(ActiveRecord.Fields(FIELD_User_StateSettings))
    clsRecord.User_MainDimensionPosition = FNullField(ActiveRecord.Fields(FIELD_User_MainDimensionPosition))
    clsRecord.User_TreeWidth = FNullField(ActiveRecord.Fields(FIELD_User_TreeWidth))
    clsRecord.User_Level = FNullField(ActiveRecord.Fields(FIELD_User_Level))
    clsRecord.User_ButtonGroup = FNullField(ActiveRecord.Fields(FIELD_User_ButtonGroup))
    clsRecord.User_ButtonBarWidth = FNullField(ActiveRecord.Fields(FIELD_User_ButtonBarWidth))
    
    clsRecord.ADMINISTRATOR_RIGHTS = FNullField(ActiveRecord.Fields(FIELD_ADMINISTRATOR_RIGHTS))
    clsRecord.MAINTAIN_TABLES = FNullField(ActiveRecord.Fields(FIELD_MAINTAIN_TABLES))
    clsRecord.ALL_LOGICAL_IDS = FNullField(ActiveRecord.Fields(FIELD_ALL_LOGICAL_IDS))
    clsRecord.SHOW_ALL_SENT = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_SENT))
    clsRecord.SHOW_ALL_WITH_ERRORS = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_WITH_ERRORS))
    clsRecord.SHOW_ALL_WAITING = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_WAITING))
    clsRecord.SHOW_ALL_DELETED = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_DELETED))
    clsRecord.CLEAN_UP_DELETED = FNullField(ActiveRecord.Fields(FIELD_CLEAN_UP_DELETED))
    clsRecord.EVERY_FIELD = FNullField(ActiveRecord.Fields(FIELD_EVERY_FIELD))
    clsRecord.DAYS_OR_ITEMS = FNullField(ActiveRecord.Fields(FIELD_DAYS_OR_ITEMS))
    clsRecord.DELETE_OTHER_USERS_ITEMS = FNullField(ActiveRecord.Fields(FIELD_DELETE_OTHER_USERS_ITEMS))
    clsRecord.DELETE_SENT_IN_N_DAYS = FNullField(ActiveRecord.Fields(FIELD_DELETE_SENT_IN_N_DAYS))
    clsRecord.WITH_SECURITY = FNullField(ActiveRecord.Fields(FIELD_WITH_SECURITY))
    clsRecord.REFRESH_IN_SECONDS = FNullField(ActiveRecord.Fields(FIELD_REFRESH_IN_SECONDS))
    clsRecord.RELATE_L1_TO_S1 = FNullField(ActiveRecord.Fields(FIELD_RELATE_L1_TO_S1))
    clsRecord.SUPPLIER_BOX = FNullField(ActiveRecord.Fields(FIELD_SUPPLIER_BOX))
    clsRecord.LAST_USED_PRINTER = FNullField(ActiveRecord.Fields(FIELD_LAST_USED_PRINTER))
    clsRecord.LAST_DV_PRINTER = FNullField(ActiveRecord.Fields(FIELD_LAST_DV_PRINTER))
    clsRecord.LOGID_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_LOGID_DESCRIPTION))
    clsRecord.FIXED_USER = FNullField(ActiveRecord.Fields(FIELD_FIXED_USER))
    clsRecord.LOGGED = FNullField(ActiveRecord.Fields(FIELD_LOGGED))
    clsRecord.SDICOUNT = FNullField(ActiveRecord.Fields(FIELD_SDICOUNT))
    clsRecord.SDECOUNT = FNullField(ActiveRecord.Fields(FIELD_SDECOUNT))
    clsRecord.SDTCOUNT = FNullField(ActiveRecord.Fields(FIELD_SDTCOUNT))
    clsRecord.DECOUNT = FNullField(ActiveRecord.Fields(FIELD_DECOUNT))
    clsRecord.SHOW_ONLY_DOCTYPE = FNullField(ActiveRecord.Fields(FIELD_SHOW_ONLY_DOCTYPE))
    clsRecord.SHOW_ONLY_VATNUM = FNullField(ActiveRecord.Fields(FIELD_SHOW_ONLY_VATNUM))
    clsRecord.SHOW_ONLY_CTRYCODE = FNullField(ActiveRecord.Fields(FIELD_SHOW_ONLY_CTRYCODE))
    clsRecord.SAVE_NEW_TARICCTRY = FNullField(ActiveRecord.Fields(FIELD_SAVE_NEW_TARICCTRY))
    clsRecord.SHOW_ALL_TOBEPRINTED = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_TOBEPRINTED))
    clsRecord.SHOW_ALL_DRAFTS = FNullField(ActiveRecord.Fields(FIELD_SHOW_ALL_DRAFTS))
    clsRecord.SDI2COUNT = FNullField(ActiveRecord.Fields(FIELD_SDI2COUNT))
    clsRecord.SDE2COUNT = FNullField(ActiveRecord.Fields(FIELD_SDE2COUNT))
    clsRecord.SDT2COUNT = FNullField(ActiveRecord.Fields(FIELD_SDT2COUNT))
    clsRecord.LANGUAGE_FIELD = FNullField(ActiveRecord.Fields(FIELD_LANGUAGE_FIELD))

    Set GetClassRecord = clsRecord
    Set clsRecord = Nothing

End Function

Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, ByVal SearchField, ByVal SearchValue) As Boolean
'
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
    
        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TABLE_NAME & " WHERE " & SearchField & "=" & Trim$(SearchValue)
        
         ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
        'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
        
        SearchRecord = Not (rstDB.EOF = True And rstDB.BOF)
        
        'If (rstDB.EOF = True And rstDB.BOF) Then
        '    SearchRecord = False
        'ElseIf (rstDB.EOF = False) Then
        '    SearchRecord = False
        'End If
        
    End If
    
    Exit Function
    
ERROR_SEARCH:
    SearchValue = False
End Function

' /* --------------- PRIVATE FUNCTIONS -------------------- */



