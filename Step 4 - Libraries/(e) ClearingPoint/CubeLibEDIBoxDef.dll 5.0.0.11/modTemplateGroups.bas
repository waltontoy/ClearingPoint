Attribute VB_Name = "modTemplateGroups"
'Option Explicit
'
'Private Const TABLE_NAME = "[BOX DEFAULT IE07 NCTS]"
'Private Const PK_FIELD = "[BOX CODE]"
'
'' field constants
'Private Const FIELD_BOX_CODE = TABLE_NAME & "." & "[BOX CODE]" ' 1
'Private Const FIELD_ENGLISH_DESCRIPTION = TABLE_NAME & "." & "[ENGLISH DESCRIPTION]" ' 2
'Private Const FIELD_DUTCH_DESCRIPTION = TABLE_NAME & "." & "[DUTCH DESCRIPTION]"  ' 3
'Private Const FIELD_FRENCH_DESCRIPTION = TABLE_NAME & "." & "[FRENCH DESCRIPTION]" ' 4
'Private Const FIELD_EMPTY_FIELD_VALUE = TABLE_NAME & "." & "[EMPTY FIELD VALUE]" ' 5
'Private Const FIELD_INSERT_FIELD = TABLE_NAME & "." & "[INSERT]" ' 6
'Private Const FIELD_JUSTIFY = TABLE_NAME & "." & "[FRENCH DESCRIPTION7]" ' 7
'Private Const FIELD_FRENCH_DESCRIPTION8 = TABLE_NAME & "." & "[FRENCH DESCRIPTION8]" ' 8
'Private Const FIELD_FRENCH_DESCRIPTION9 = TABLE_NAME & "." & "[FRENCH DESCRIPTION9]" ' 9
'Private Const FIELD_FRENCH_DESCRIPTION10 = TABLE_NAME & "." & "[FRENCH DESCRIPTION10]" ' 10
'Private Const FIELD_FRENCH_DESCRIPTION11 = TABLE_NAME & "." & "[FRENCH DESCRIPTION11]" ' 11
'Private Const FIELD_FRENCH_DESCRIPTION12 = TABLE_NAME & "." & "[FRENCH DESCRIPTION12]" ' 12
'Private Const FIELD_FRENCH_DESCRIPTION13 = TABLE_NAME & "." & "[FRENCH DESCRIPTION13]" ' 13
'Private Const FIELD_FRENCH_DESCRIPTION14 = TABLE_NAME & "." & "[FRENCH DESCRIPTION14]" ' 14
'Private Const FIELD_FRENCH_DESCRIPTION15 = TABLE_NAME & "." & "[FRENCH DESCRIPTION15]" ' 15
'Private Const FIELD_FRENCH_DESCRIPTION16 = TABLE_NAME & "." & "[FRENCH DESCRIPTION16]" ' 16
'Private Const FIELD_FRENCH_DESCRIPTION17 = TABLE_NAME & "." & "[FRENCH DESCRIPTION17]" ' 17
'Private Const FIELD_FRENCH_DESCRIPTION18 = TABLE_NAME & "." & "[FRENCH DESCRIPTION18]" ' 18
'Private Const FIELD_FRENCH_DESCRIPTION19 = TABLE_NAME & "." & "[FRENCH DESCRIPTION19]" ' 19
'Private Const FIELD_FRENCH_DESCRIPTION20 = TABLE_NAME & "." & "[FRENCH DESCRIPTION20]" ' 20
'Private Const FIELD_FRENCH_DESCRIPTION21 = TABLE_NAME & "." & "[FRENCH DESCRIPTION21]" ' 21
'Private Const FIELD_FRENCH_DESCRIPTION22 = TABLE_NAME & "." & "[FRENCH DESCRIPTION22]" ' 22
'
'Private Const SQL_ADD_RECORD = "INSERT INTO [BOX DEFAULT IE07 NCTS] (" & "[BOX CODE] " & _
'                                            "," & "[ENGLISH DESCRIPTION] " & _
'                                            "," & "[DUTCH DESCRIPTION] " & _
'                                            "," & "[FRENCH DESCRIPTION] " & _
'                                            "," & "[EMPTY FIELD VALUE] " & _
'                                            "," & "[INSERT] " & _
'                                            "," & "[FRENCH DESCRIPTION7] " & _
'                                            "," & "[FRENCH DESCRIPTION8] " & _
'                                            "," & "[FRENCH DESCRIPTION9] " & _
'                                            "," & "[FRENCH DESCRIPTION10] " & _
'                                            "," & "[FRENCH DESCRIPTION11] " & _
'                                            "," & "[FRENCH DESCRIPTION12] " & _
'                                            "," & "[FRENCH DESCRIPTION13] " & _
'                                            "," & "[FRENCH DESCRIPTION14] " & _
'                                            "," & "[FRENCH DESCRIPTION15] " & _
'                                            "," & "[FRENCH DESCRIPTION16] " & _
'                                            "," & "[FRENCH DESCRIPTION17] " & _
'                                            "," & "[FRENCH DESCRIPTION18] " & _
'                                            "," & "[FRENCH DESCRIPTION19] " & _
'                                            "," & "[FRENCH DESCRIPTION20] " & _
'                                            "," & "[FRENCH DESCRIPTION21] " & _
'                                            "," & "[FRENCH DESCRIPTION22] " & _
'                                            ") VALUES "
'
'Private Const SQL_DELETE_RECORD = "DELETE * FROM [BOX DEFAULT IE07 NCTS] WHERE [BOX CODE]="
'
''Private Const SQL_MODIFY_RECORD '=
'
'Private Const SQL_GET_RECORD = "SELECT [BOX CODE] " & _
'                                                    ",[ENGLISH DESCRIPTION] " & _
'                                                    ",[DUTCH DESCRIPTION] " & _
'                                                    ",[FRENCH DESCRIPTION] " & _
'                                                    ",[EMPTY FIELD VALUE] " & _
'                                                    ",[INSERT] " & _
'                                                    ",[FRENCH DESCRIPTION7] " & _
'                                                    ",[FRENCH DESCRIPTION8] " & _
'                                                    ",[FRENCH DESCRIPTION9] " & _
'                                                    ",[FRENCH DESCRIPTION10] " & _
'                                                    ",[FRENCH DESCRIPTION11] " & _
'                                                    ",[FRENCH DESCRIPTION12] " & _
'                                                    ",[FRENCH DESCRIPTION13] " & _
'                                                    ",[FRENCH DESCRIPTION14] " & _
'                                                    ",[FRENCH DESCRIPTION15] " & _
'                                                    ",[FRENCH DESCRIPTION16] " & _
'                                                    ",[FRENCH DESCRIPTION17] " & _
'                                                    ",[FRENCH DESCRIPTION18] " & _
'                                                    ",[FRENCH DESCRIPTION19] " & _
'                                                    ",[FRENCH DESCRIPTION20] " & _
'                                                    ",[FRENCH DESCRIPTION21] " & _
'                                                    ",[FRENCH DESCRIPTION22] " & _
'                                            "FROM [BOX DEFAULT IE07 NCTS] WHERE [BOX CODE]= "
'
'Private Const SQL_GET_MAXID = "SELECT MAX([BOX CODE]) AS [ID_MAX] FROM [BOX DEFAULT IE07 NCTS]"
'Private Const SQL_GET_MINID = "SELECT Min([BOX CODE]) AS [ID_MIN] FROM [BOX DEFAULT IE07 NCTS]"
'
'' /* --------------- BASIC FUNCTIONS -------------------- */
'Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'
'    If ((ActiveRecord Is Nothing) = False) Then
'
'        strSql = SQL_ADD_RECORD & "(" & "'" & ActiveRecord.BOX_CODE & "'" & _
'                        ",'" & ActiveRecord.ENGLISH_DESCRIPTION & "'" & _
'                        ",'" & ActiveRecord.DUTCH_DESCRIPTION & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION & "'" & _
'                        ",'" & ActiveRecord.EMPTY_FIELD_VALUE & "'" & _
'                        ",'" & ActiveRecord.INSERT_FIELD & "'" & _
'                        ",'" & ActiveRecord.JUSTIFY & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION8 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION9 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION10 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION11 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION12 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION13 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION14 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION15 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION16 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION17 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION18 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION19 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION20 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION21 & "'" & _
'                        ",'" & ActiveRecord.FRENCH_DESCRIPTION22 & "'" & _
'                        ")"
'
'        On Error GoTo ERROR_QUERY
'        ExecuteNonQuery ActiveConnection, strSql
'
'        AddRecord = True
'        Exit Function
'
'    End If
'
'    AddRecord = False
'
'    Exit Function
'
'ERROR_QUERY:
'    AddRecord = False
'End Function
'
'Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'
'    If ((ActiveRecord Is Nothing) = False) Then
'
'        strSql = SQL_DELETE_RECORD & ActiveRecord.BOX_CODE
'
'        On Error GoTo ERROR_QUERY
'        ExecuteNonQuery ActiveConnection, strSql
'
'        DeleteRecord = True
'        Exit Function
'
'    End If
'
'    DeleteRecord = False
'
'    Exit Function
'
'ERROR_QUERY:
'    DeleteRecord = False
'End Function
'
'' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
'Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'
'    If ((ActiveRecord Is Nothing) = False) Then
'
'        strSql = "UPDATE " & TABLE_NAME & " SET " & FIELD_ENGLISH_DESCRIPTION & "='" & ActiveRecord.ENGLISH_DESCRIPTION & "'" & _
'                        "," & FIELD_DUTCH_DESCRIPTION & "='" & ActiveRecord.DUTCH_DESCRIPTION & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION & "='" & ActiveRecord.FRENCH_DESCRIPTION & "'" & _
'                        "," & FIELD_EMPTY_FIELD_VALUE & "='" & ActiveRecord.EMPTY_FIELD_VALUE & "'" & _
'                        "," & FIELD_INSERT_FIELD & "='" & ActiveRecord.INSERT_FIELD & "'" & _
'                        "," & FIELD_JUSTIFY & "='" & ActiveRecord.JUSTIFY & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION8 & "='" & ActiveRecord.FRENCH_DESCRIPTION8 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION9 & "='" & ActiveRecord.FRENCH_DESCRIPTION9 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION10 & "='" & ActiveRecord.FRENCH_DESCRIPTION10 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION11 & "='" & ActiveRecord.FRENCH_DESCRIPTION11 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION12 & "='" & ActiveRecord.FRENCH_DESCRIPTION12 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION13 & "='" & ActiveRecord.FRENCH_DESCRIPTION13 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION14 & "='" & ActiveRecord.FRENCH_DESCRIPTION14 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION15 & "='" & ActiveRecord.FRENCH_DESCRIPTION15 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION16 & "='" & ActiveRecord.FRENCH_DESCRIPTION16 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION17 & "='" & ActiveRecord.FRENCH_DESCRIPTION17 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION18 & "='" & ActiveRecord.FRENCH_DESCRIPTION18 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION19 & "='" & ActiveRecord.FRENCH_DESCRIPTION19 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION20 & "='" & ActiveRecord.FRENCH_DESCRIPTION20 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION21 & "='" & ActiveRecord.FRENCH_DESCRIPTION21 & "'" & _
'                        "," & FIELD_FRENCH_DESCRIPTION22 & "='" & ActiveRecord.FRENCH_DESCRIPTION22 & "'" & _
'                        " WHERE " & PK_FIELD & "='" & ActiveRecord.BOX_CODE & "'"
'
'        On Error GoTo ERROR_QUERY
'        ExecuteNonQuery ActiveConnection, strSql
'
'        ModifyRecord = True
'        Exit Function
'
'    End If
'
'    ModifyRecord = False
'
'    Exit Function
'
'ERROR_QUERY:
'    ModifyRecord = False
'End Function
'
'Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'    Dim rstRecord As ADODB.Recordset
'
'    If ((ActiveRecord Is Nothing) = False) Then
'
'        strSql = SQL_GET_RECORD & ActiveRecord.BOX_CODE
'
'        On Error GoTo ERROR_QUERY
'        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
'
'        On Error GoTo ERROR_RECORDSET
'        ActiveRecord.ENGLISH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_ENGLISH_DESCRIPTION))
'        ActiveRecord.DUTCH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_DUTCH_DESCRIPTION))
'        ActiveRecord.FRENCH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION))
'        ActiveRecord.EMPTY_FIELD_VALUE = FNullField(rstRecord.Fields(FIELD_EMPTY_FIELD_VALUE))
'        ActiveRecord.INSERT_FIELD = FNullField(rstRecord.Fields(FIELD_INSERT_FIELD))
'        ActiveRecord.JUSTIFY = FNullField(rstRecord.Fields(FIELD_JUSTIFY))
'        ActiveRecord.FRENCH_DESCRIPTION8 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION8))
'        ActiveRecord.FRENCH_DESCRIPTION9 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION9))
'        ActiveRecord.FRENCH_DESCRIPTION10 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION10))
'        ActiveRecord.FRENCH_DESCRIPTION11 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION11))
'        ActiveRecord.FRENCH_DESCRIPTION12 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION12))
'        ActiveRecord.FRENCH_DESCRIPTION13 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION13))
'        ActiveRecord.FRENCH_DESCRIPTION14 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION14))
'        ActiveRecord.FRENCH_DESCRIPTION15 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION15))
'        ActiveRecord.FRENCH_DESCRIPTION16 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION16))
'        ActiveRecord.FRENCH_DESCRIPTION17 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION17))
'        ActiveRecord.FRENCH_DESCRIPTION18 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION18))
'        ActiveRecord.FRENCH_DESCRIPTION19 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION19))
'        ActiveRecord.FRENCH_DESCRIPTION20 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION20))
'        ActiveRecord.FRENCH_DESCRIPTION21 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION21))
'        ActiveRecord.FRENCH_DESCRIPTION22 = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION22))
'
'        ADORecordsetClose rstRecord
'        GetRecord = True
'
'    ElseIf ((ActiveRecord Is Nothing) = True) Then
'        GetRecord = False
'    End If
'
'    Exit Function
'
'ERROR_RECORDSET:
'    ADORecordsetClose rstRecord
'    GetRecord = False
'    Exit Function
'ERROR_QUERY:
'    GetRecord = False
'End Function
'
'Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'    Dim rstRecord As ADODB.Recordset
'
'    If ((ActiveRecord Is Nothing) = True) Then
'        Set ActiveRecord = New cpiBOX_DEFAULT_IE07_NCTS_I
'    End If
'
'    strSql = SQL_GET_MAXID
'
'    On Error GoTo ERROR_QUERY
'    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
'
'    On Error GoTo ERROR_RECORDSET
'    ActiveRecord.BOX_CODE = rstRecord.Fields("ID_MAX").Value
'
'    ADORecordsetClose rstRecord
'    GetMaxID = True
'    Exit Function
'
'ERROR_RECORDSET:
'    ADORecordsetClose rstRecord
'    GetMaxID = False
'    Exit Function
'ERROR_QUERY:
'    GetMaxID = False
'End Function
'
'Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As Boolean
'
'    Dim strSql As String
'    Dim rstRecord As ADODB.Recordset
'
'    If ((ActiveRecord Is Nothing) = True) Then
'        Set ActiveRecord = New cpiBOX_DEFAULT_IE07_NCTS_I
'    End If
'
'    strSql = SQL_GET_MINID
'
'    On Error GoTo ERROR_QUERY
'    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
'
'    On Error GoTo ERROR_RECORDSET
'    ActiveRecord.BOX_CODE = rstRecord.Fields("ID_MIN").Value
'
'    ADORecordsetClose rstRecord
'    GetMinID = True
'    Exit Function
'
'ERROR_RECORDSET:
'    ADORecordsetClose rstRecord
'    GetMinID = False
'    Exit Function
'ERROR_QUERY:
'    GetMinID = False
'End Function
'
'' conversion
'Public Function GetTableRecord(ByRef ActiveRecord As cpiBOX_DEFAULT_IE07_NCTS_I) As ADODB.Recordset
'    Dim rstRecord As ADODB.Recordset
'
'    Set rstRecord = New ADODB.Recordset
'
'    rstRecord.Open
'    rstRecord.Fields.Append FIELD_BOX_CODE, 3, 4, 90
'    rstRecord.Fields.Append FIELD_ENGLISH_DESCRIPTION, 3, 4, 118
'    rstRecord.Fields.Append FIELD_DUTCH_DESCRIPTION, 202, 50, 70
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION, 3, 4, 118
'    rstRecord.Fields.Append FIELD_EMPTY_FIELD_VALUE, 3, 4, 118
'    rstRecord.Fields.Append FIELD_INSERT_FIELD, 3, 4, 118
'    rstRecord.Fields.Append FIELD_JUSTIFY, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION8, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION9, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION10, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION11, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION12, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION13, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION14, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION15, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION16, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION17, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION18, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION19, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION20, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION21, 3, 4, 118
'    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION22, 3, 4, 118
'
'    ' set values
'    rstRecord.AddNew
'    rstRecord.Fields(FIELD_BOX_CODE) = ActiveRecord.BOX_CODE
'    rstRecord.Fields(FIELD_ENGLISH_DESCRIPTION) = ActiveRecord.ENGLISH_DESCRIPTION
'    rstRecord.Fields(FIELD_DUTCH_DESCRIPTION) = ActiveRecord.DUTCH_DESCRIPTION
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION) = ActiveRecord.FRENCH_DESCRIPTION
'    rstRecord.Fields(FIELD_EMPTY_FIELD_VALUE) = ActiveRecord.EMPTY_FIELD_VALUE
'    rstRecord.Fields(FIELD_INSERT_FIELD) = ActiveRecord.INSERT_FIELD
'    rstRecord.Fields(FIELD_JUSTIFY) = ActiveRecord.JUSTIFY
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION8) = ActiveRecord.FRENCH_DESCRIPTION8
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION9) = ActiveRecord.FRENCH_DESCRIPTION9
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION10) = ActiveRecord.FRENCH_DESCRIPTION10
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION11) = ActiveRecord.FRENCH_DESCRIPTION11
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION12) = ActiveRecord.FRENCH_DESCRIPTION12
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION13) = ActiveRecord.FRENCH_DESCRIPTION13
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION14) = ActiveRecord.FRENCH_DESCRIPTION14
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION15) = ActiveRecord.FRENCH_DESCRIPTION15
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION16) = ActiveRecord.FRENCH_DESCRIPTION16
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION17) = ActiveRecord.FRENCH_DESCRIPTION17
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION18) = ActiveRecord.FRENCH_DESCRIPTION18
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION19) = ActiveRecord.FRENCH_DESCRIPTION19
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION20) = ActiveRecord.FRENCH_DESCRIPTION20
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION21) = ActiveRecord.FRENCH_DESCRIPTION21
'    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION22) = ActiveRecord.FRENCH_DESCRIPTION22
'    rstRecord.Update
'
'    Set GetTableRecord = rstRecord
'    Set rstRecord = Nothing
'
'End Function
'
'Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiBOX_DEFAULT_IE07_NCTS_I
'
'    Dim clsRecord As cpiBOX_DEFAULT_IE07_NCTS_I
'
'    Set clsRecord = New cpiBOX_DEFAULT_IE07_NCTS_I
'
'    clsRecord.BOX_CODE = FNullField(ActiveRecord.Fields(FIELD_BOX_CODE))
'    clsRecord.ENGLISH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_ENGLISH_DESCRIPTION))
'    clsRecord.DUTCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_DUTCH_DESCRIPTION))
'    clsRecord.FRENCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION))
'    clsRecord.EMPTY_FIELD_VALUE = FNullField(ActiveRecord.Fields(FIELD_EMPTY_FIELD_VALUE))
'    clsRecord.INSERT_FIELD = FNullField(ActiveRecord.Fields(FIELD_INSERT_FIELD))
'    clsRecord.JUSTIFY = FNullField(ActiveRecord.Fields(FIELD_JUSTIFY))
'    clsRecord.FRENCH_DESCRIPTION8 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION8))
'    clsRecord.FRENCH_DESCRIPTION9 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION9))
'    clsRecord.FRENCH_DESCRIPTION10 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION10))
'    clsRecord.FRENCH_DESCRIPTION11 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION11))
'    clsRecord.FRENCH_DESCRIPTION12 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION12))
'    clsRecord.FRENCH_DESCRIPTION13 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION13))
'    clsRecord.FRENCH_DESCRIPTION14 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION14))
'    clsRecord.FRENCH_DESCRIPTION15 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION15))
'    clsRecord.FRENCH_DESCRIPTION16 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION16))
'    clsRecord.FRENCH_DESCRIPTION17 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION17))
'    clsRecord.FRENCH_DESCRIPTION18 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION18))
'    clsRecord.FRENCH_DESCRIPTION19 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION19))
'    clsRecord.FRENCH_DESCRIPTION20 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION20))
'    clsRecord.FRENCH_DESCRIPTION21 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION21))
'    clsRecord.FRENCH_DESCRIPTION22 = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION22))
'
'    Set GetClassRecord = clsRecord
'    Set clsRecord = Nothing
'
'End Function
'
'Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, ByVal SearchField, ByVal SearchValue) As Boolean
''
'    Dim strSql As String
'    Dim lngRecordsAffected As Long
'
'    On Error GoTo ERROR_SEARCH
'    SearchField = Trim$(SearchField)
'    If ((SearchField <> "") And (SearchValue <> "")) Then
'        If (Len(SearchField) > 2) Then
'            If ((Left$(SearchField, 1) <> "[") And (Right$(SearchField, 1) <> "]")) Then
'                SearchField = "[" & SearchField & "]"
'            End If
'        End If
'
'        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TABLE_NAME & " WHERE " & SearchField & "=" & Trim$(SearchValue)
'        ExecuteNonQuery ActiveConnection, strSql, lngRecordsAffected
'
'        If (lngRecordsAffected > 0) Then
'            SearchValue = True
'        ElseIf (lngRecordsAffected = 0) Then
'            SearchValue = False
'        End If
'
'    End If
'
'    Exit Function
'
'ERROR_SEARCH:
'    SearchValue = False
'End Function
'
''Public Function GetRecordset(ByRef ActiveConnection As ADODB.Connection, ByRef CommandText As String) As ADODB.Recordset
''
''    Dim rstRecordset As ADODB.Recordset
''
''    Set rstRecordset = New ADODB.Recordset
''
''    On Error GoTo ERROR_RECORDSET
''    ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic
''
''    Set GetRecordset = rstRecordset
''    Set rstRecordset = Nothing
''
''    Exit Function
''
''ERROR_RECORDSET:
''    Set rstRecordset = Nothing
''End Function
'
'
'' /* --------------- PRIVATE FUNCTIONS -------------------- */
'
'
'
'
'
'Global Const YYY = "0"
