VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBoxDefEdIIE07s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Private Const TABLE_NAME = "[BOX DEFAULT EDI NCTS2 ADMIN]"
Private Const PK_FIELD = "[BOX CODE]"
    
' field constants
''Private Const FIELD_BOX_CODE = TABLE_NAME & "." & "[BOX CODE]" ' 1
''Private Const FIELD_ENGLISH_DESCRIPTION = TABLE_NAME & "." & "[ENGLISH DESCRIPTION]" ' 2
''Private Const FIELD_DUTCH_DESCRIPTION = TABLE_NAME & "." & "[DUTCH DESCRIPTION]"  ' 3
''Private Const FIELD_FRENCH_DESCRIPTION = TABLE_NAME & "." & "[FRENCH DESCRIPTION]" ' 4
''Private Const FIELD_EMPTY_FIELD_VALUE = TABLE_NAME & "." & "[EMPTY FIELD VALUE]" ' 5
''Private Const FIELD_INSERT_FIELD = TABLE_NAME & "." & "[INSERT]" ' 6
''Private Const FIELD_JUSTIFY_FIELD = TABLE_NAME & "." & "[JUSTIFY]" ' 7
''Private Const FIELD_SEQUENTIAL_TABBING_DEFAULT = TABLE_NAME & "." & "[SEQUENTIAL TABBING DEFAULT]" ' 8
''Private Const FIELD_SEQUENTIAL_TABBING_ACTIVE = TABLE_NAME & "." & "[SEQUENTIAL TABBING ACTIVE]" ' 9
''Private Const FIELD_CHECK_VAT = TABLE_NAME & "." & "[CHECK VAT]" ' 10
''Private Const FIELD_CALCULATE_NET_WEIGHT = TABLE_NAME & "." & "[CALCULATE NET WEIGHT]" ' 11
''Private Const FIELD_CALCULATE_NO_OF_ITEMS = TABLE_NAME & "." & "[CALCULATE NO OF ITEMS]" ' 12
''Private Const FIELD_COPY_TO_NEXT = TABLE_NAME & "." & "[COPY TO NEXT]" ' 13
''Private Const FIELD_CHANGE_ONLY_IN_H1 = TABLE_NAME & "." & "[CHANGE ONLY IN H1]" ' 14
''Private Const FIELD_SEND_ONLY_IN_H1 = TABLE_NAME & "." & "[SEND ONLY IN H1]" ' 15
''Private Const FIELD_WIDTH_FIELD = TABLE_NAME & "." & "[WIDTH]" ' 16
''Private Const FIELD_DECIMAL_FIELD = TABLE_NAME & "." & "[DECIMAL]" ' 17
''Private Const FIELD_DATA_TYPE = TABLE_NAME & "." & "[DATA TYPE]" ' 18
''Private Const FIELD_AUTO_ADD = TABLE_NAME & "." & "[AUTO ADD]" ' 19
''Private Const FIELD_DEFAULT_VALUE = TABLE_NAME & "." & "[DEFAULT VALUE]" ' 20
''Private Const FIELD_CALCULATE_CUSTOMS_VALUE = TABLE_NAME & "." & "[CALCULATE CUSTOMS VALUE]" ' 21
''Private Const FIELD_RELATE_L1_TO_S1 = TABLE_NAME & "." & "[RELATE L1 TO S1]" ' 22
''Private Const FIELD_BOX_TYPE = TABLE_NAME & "." & "[BOX TYPE]" ' 23
''Private Const FIELD_BOX_INDEX = TABLE_NAME & "." & "[BOX INDEX]" ' 24
''Private Const FIELD_INTERNAL_CODE = TABLE_NAME & "." & "[INTERNAL CODE]" ' 25

Private Const FIELD_BOX_CODE = "BOX CODE" ' 1
Private Const FIELD_ENGLISH_DESCRIPTION = "ENGLISH DESCRIPTION" ' 2
Private Const FIELD_DUTCH_DESCRIPTION = "DUTCH DESCRIPTION"  ' 3
Private Const FIELD_FRENCH_DESCRIPTION = "FRENCH DESCRIPTION" ' 4
Private Const FIELD_EMPTY_FIELD_VALUE = "EMPTY FIELD VALUE" ' 5
Private Const FIELD_INSERT_FIELD = "INSERT" ' 6
Private Const FIELD_JUSTIFY_FIELD = "JUSTIFY" ' 7
Private Const FIELD_SEQUENTIAL_TABBING_DEFAULT = "SEQUENTIAL TABBING DEFAULT" ' 8
Private Const FIELD_SEQUENTIAL_TABBING_ACTIVE = "SEQUENTIAL TABBING ACTIVE" ' 9
Private Const FIELD_CHECK_VAT = "CHECK VAT" ' 10
Private Const FIELD_CALCULATE_NET_WEIGHT = "CALCULATE NET WEIGHT" ' 11
Private Const FIELD_CALCULATE_NO_OF_ITEMS = "CALCULATE NO OF ITEMS" ' 12
Private Const FIELD_COPY_TO_NEXT = "COPY TO NEXT" ' 13
Private Const FIELD_CHANGE_ONLY_IN_H1 = "CHANGE ONLY IN H1" ' 14
Private Const FIELD_SEND_ONLY_IN_H1 = "SEND ONLY IN H1" ' 15
Private Const FIELD_WIDTH_FIELD = "WIDTH" ' 16
Private Const FIELD_DECIMAL_FIELD = "DECIMAL" ' 17
Private Const FIELD_DATA_TYPE = "DATA TYPE" ' 18
Private Const FIELD_AUTO_ADD = "AUTO ADD" ' 19
Private Const FIELD_DEFAULT_VALUE = "DEFAULT VALUE" ' 20
Private Const FIELD_CALCULATE_CUSTOMS_VALUE = "CALCULATE CUSTOMS VALUE" ' 21
Private Const FIELD_RELATE_L1_TO_S1 = "RELATE L1 TO S1" ' 22
Private Const FIELD_BOX_TYPE = "BOX TYPE" ' 23
Private Const FIELD_BOX_INDEX = "BOX INDEX" ' 24
Private Const FIELD_INTERNAL_CODE = "INTERNAL CODE" ' 25
'INTERNAL_CODE

Private Const SQL_ADD_RECORD = "INSERT INTO [BOX DEFAULT EDI NCTS2 ADMIN] (" & "[BOX CODE] " & _
                                            "," & "[ENGLISH DESCRIPTION] " & _
                                            "," & "[DUTCH DESCRIPTION] " & _
                                            "," & "[FRENCH DESCRIPTION] " & _
                                            "," & "[EMPTY FIELD VALUE] " & _
                                            "," & "[INSERT] " & _
                                            "," & "[JUSTIFY] " & _
                                            "," & "[SEQUENTIAL TABBING DEFAULT] " & _
                                            "," & "[SEQUENTIAL TABBING ACTIVE] " & _
                                            "," & "[CHECK VAT] " & _
                                            "," & "[CALCULATE NET WEIGHT] " & _
                                            "," & "[CALCULATE NO OF ITEMS] " & _
                                            "," & "[COPY TO NEXT] " & _
                                            "," & "[CHANGE ONLY IN H1] " & _
                                            "," & "[SEND ONLY IN H1] " & _
                                            "," & "[WIDTH] " & _
                                            "," & "[DECIMAL] " & _
                                            "," & "[DATA TYPE] " & _
                                            "," & "[AUTO ADD] " & _
                                            "," & "[DEFAULT VALUE] " & _
                                            "," & "[CALCULATE CUSTOMS VALUE] " & _
                                            "," & "[RELATE L1 TO S1] " & _
                                            "," & "[BOX TYPE] " & "," & "[BOX INDEX] " & _
                                            "," & "[INTERNAL CODE] " & _
                                            ") VALUES "

Private Const SQL_DELETE_RECORD = "DELETE * FROM [BOX DEFAULT EDI NCTS2 ADMIN] WHERE [BOX CODE]="

'Private Const SQL_MODIFY_RECORD '=

Private Const SQL_WHERE_KEY = " WHERE [BOX CODE]= "

Private Const SQL_GET_ALL_RECORD = "SELECT [BOX CODE] " & _
                                                    ",[ENGLISH DESCRIPTION] " & _
                                                    ",[DUTCH DESCRIPTION] " & _
                                                    ",[FRENCH DESCRIPTION] " & _
                                                    ",[EMPTY FIELD VALUE] " & _
                                                    ",[INSERT] " & _
                                                    ",[JUSTIFY] " & _
                                                    ",[SEQUENTIAL TABBING DEFAULT] " & _
                                                    ",[SEQUENTIAL TABBING ACTIVE] " & _
                                                    ",[CHECK VAT] " & _
                                                    ",[CALCULATE NET WEIGHT] " & _
                                                    ",[CALCULATE NO OF ITEMS] " & _
                                                    ",[COPY TO NEXT] " & _
                                                    ",[CHANGE ONLY IN H1] " & _
                                                    ",[SEND ONLY IN H1] " & _
                                                    ",[WIDTH] " & _
                                                    ",[DECIMAL] " & _
                                                    ",[DATA TYPE] " & _
                                                    ",[AUTO ADD] " & _
                                                    ",[DEFAULT VALUE] " & _
                                                    ",[CALCULATE CUSTOMS VALUE] " & _
                                                    ",[RELATE L1 TO S1] " & _
                                                    ",[BOX TYPE] " & ",[BOX INDEX] " & _
                                                    ",[INTERNAL CODE] " & _
                                            "FROM [BOX DEFAULT EDI NCTS2 ADMIN] " 'WHERE [BOX CODE]= "

Private Const SQL_GET_RECORD = SQL_GET_ALL_RECORD & SQL_WHERE_KEY

Private Const SQL_GET_MAXID = "SELECT MAX([BOX CODE]) AS [ID_MAX] FROM [BOX DEFAULT EDI NCTS2 ADMIN]"
Private Const SQL_GET_MINID = "SELECT Min([BOX CODE]) AS [ID_MIN] FROM [BOX DEFAULT EDI NCTS2 ADMIN]"

Dim mvarRecordset As ADODB.Recordset

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "(" & SetEnclosedChr(ActiveRecord.BOX_CODE) & _
                        "," & SetEnclosedChr(ActiveRecord.ENGLISH_DESCRIPTION) & _
                        "," & SetEnclosedChr(ActiveRecord.DUTCH_DESCRIPTION) & _
                        "," & SetEnclosedChr(ActiveRecord.FRENCH_DESCRIPTION) & _
                        "," & SetEnclosedChr(ActiveRecord.EMPTY_FIELD_VALUE) & _
                        "," & SetEnclosedChr(ActiveRecord.INSERT_FIELD) & _
                        "," & SetEnclosedChr(ActiveRecord.JUSTIFY_FIELD) & _
                        "," & SetEnclosedChr(ActiveRecord.SEQUENTIAL_TABBING_DEFAULT) & _
                        "," & SetEnclosedChr(ActiveRecord.SEQUENTIAL_TABBING_ACTIVE) & _
                        "," & SetEnclosedChr(ActiveRecord.CHECK_VAT) & _
                        "," & SetEnclosedChr(ActiveRecord.CALCULATE_NET_WEIGHT) & _
                        "," & SetEnclosedChr(ActiveRecord.CALCULATE_NO_OF_ITEMS) & _
                        "," & SetEnclosedChr(ActiveRecord.COPY_TO_NEXT) & _
                        "," & SetEnclosedChr(ActiveRecord.CHANGE_ONLY_IN_H1) & _
                        "," & SetEnclosedChr(ActiveRecord.SEND_ONLY_IN_H1) & _
                        "," & SetEnclosedChr(ActiveRecord.WIDTH_FIELD) & _
                        "," & SetEnclosedChr(ActiveRecord.DECIMAL_FIELD) & _
                        "," & SetEnclosedChr(ActiveRecord.DATA_TYPE) & _
                        "," & SetEnclosedChr(ActiveRecord.AUTO_ADD) & _
                        "," & SetEnclosedChr(ActiveRecord.DEFAULT_VALUE) & _
                        "," & SetEnclosedChr(ActiveRecord.CALCULATE_CUSTOMS_VALUE) & _
                        "," & SetEnclosedChr(ActiveRecord.RELATE_L1_TO_S1) & _
                        "," & SetEnclosedChr(ActiveRecord.BOX_TYPE) & _
                        "," & SetEnclosedChr(ActiveRecord.BOX_INDEX)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.INTERNAL_CODE)
        strSql = strSql & ")"
    
        On Error GoTo ERROR_QUERY
        
        ExecuteNonQuery ActiveConnection, strSql
    
        AddRecord = True
        Exit Function
        
    End If
    
    AddRecord = False
    
    Exit Function
    
ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.BOX_CODE)
    
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

' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET " & FIELD_ENGLISH_DESCRIPTION & "=" & SetEnclosedChr(ActiveRecord.ENGLISH_DESCRIPTION) & _
                        "," & FIELD_DUTCH_DESCRIPTION & "=" & SetEnclosedChr(ActiveRecord.DUTCH_DESCRIPTION) & "" & _
                        "," & FIELD_FRENCH_DESCRIPTION & "=" & SetEnclosedChr(ActiveRecord.FRENCH_DESCRIPTION) & _
                        "," & FIELD_EMPTY_FIELD_VALUE & "=" & SetEnclosedChr(ActiveRecord.EMPTY_FIELD_VALUE) & _
                        "," & FIELD_INSERT_FIELD & "=" & SetEnclosedChr(ActiveRecord.INSERT_FIELD) & _
                        "," & FIELD_JUSTIFY_FIELD & "=" & SetEnclosedChr(ActiveRecord.JUSTIFY_FIELD) & _
                        "," & FIELD_SEQUENTIAL_TABBING_DEFAULT & "=" & SetEnclosedChr(ActiveRecord.SEQUENTIAL_TABBING_DEFAULT) & _
                        "," & FIELD_SEQUENTIAL_TABBING_ACTIVE & "=" & SetEnclosedChr(ActiveRecord.SEQUENTIAL_TABBING_ACTIVE) & _
                        "," & FIELD_CHECK_VAT & "=" & SetEnclosedChr(ActiveRecord.CHECK_VAT) & _
                        "," & FIELD_CALCULATE_NET_WEIGHT & "=" & SetEnclosedChr(ActiveRecord.CALCULATE_NET_WEIGHT) & _
                        "," & FIELD_CALCULATE_NO_OF_ITEMS & "=" & SetEnclosedChr(ActiveRecord.CALCULATE_NO_OF_ITEMS) & _
                        "," & FIELD_COPY_TO_NEXT & "=" & SetEnclosedChr(ActiveRecord.COPY_TO_NEXT) & _
                        "," & FIELD_CHANGE_ONLY_IN_H1 & "=" & SetEnclosedChr(ActiveRecord.CHANGE_ONLY_IN_H1) & _
                        "," & FIELD_SEND_ONLY_IN_H1 & "=" & SetEnclosedChr(ActiveRecord.SEND_ONLY_IN_H1) & _
                        "," & FIELD_WIDTH_FIELD & "=" & SetEnclosedChr(ActiveRecord.WIDTH_FIELD) & _
                        "," & FIELD_DECIMAL_FIELD & "=" & SetEnclosedChr(ActiveRecord.DECIMAL_FIELD) & _
                        "," & FIELD_DATA_TYPE & "=" & SetEnclosedChr(ActiveRecord.DATA_TYPE) & _
                        "," & FIELD_AUTO_ADD & "=" & SetEnclosedChr(ActiveRecord.AUTO_ADD) & _
                        "," & FIELD_DEFAULT_VALUE & "=" & SetEnclosedChr(ActiveRecord.DEFAULT_VALUE) & _
                        "," & FIELD_CALCULATE_CUSTOMS_VALUE & "=" & SetEnclosedChr(ActiveRecord.CALCULATE_CUSTOMS_VALUE) & _
                        "," & FIELD_RELATE_L1_TO_S1 & "=" & SetEnclosedChr(ActiveRecord.RELATE_L1_TO_S1) & _
                        "," & FIELD_BOX_TYPE & "=" & SetEnclosedChr(ActiveRecord.BOX_TYPE) & _
                        "," & FIELD_BOX_INDEX & "=" & SetEnclosedChr(ActiveRecord.BOX_INDEX)
        strSql = strSql & "," & FIELD_INTERNAL_CODE & "=" & SetEnclosedChr(ActiveRecord.INTERNAL_CODE)
        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.BOX_CODE)
    
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

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset
    
    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.BOX_CODE)
        
        On Error GoTo ERROR_QUERY
        
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        
        On Error GoTo ERROR_RECORDSET
        ActiveRecord.ENGLISH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_ENGLISH_DESCRIPTION))
        ActiveRecord.DUTCH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_DUTCH_DESCRIPTION))
        ActiveRecord.FRENCH_DESCRIPTION = FNullField(rstRecord.Fields(FIELD_FRENCH_DESCRIPTION))
        ActiveRecord.EMPTY_FIELD_VALUE = FNullField(rstRecord.Fields(FIELD_EMPTY_FIELD_VALUE))
        ActiveRecord.INSERT_FIELD = FNullField(rstRecord.Fields(FIELD_INSERT_FIELD))
        ActiveRecord.JUSTIFY_FIELD = FNullField(rstRecord.Fields(FIELD_JUSTIFY_FIELD))
        ActiveRecord.SEQUENTIAL_TABBING_DEFAULT = FNullField(rstRecord.Fields(FIELD_SEQUENTIAL_TABBING_DEFAULT))
        ActiveRecord.SEQUENTIAL_TABBING_ACTIVE = FNullField(rstRecord.Fields(FIELD_SEQUENTIAL_TABBING_ACTIVE))
        ActiveRecord.CHECK_VAT = FNullField(rstRecord.Fields(FIELD_CHECK_VAT))
        ActiveRecord.CALCULATE_NET_WEIGHT = FNullField(rstRecord.Fields(FIELD_CALCULATE_NET_WEIGHT))
        ActiveRecord.CALCULATE_NO_OF_ITEMS = FNullField(rstRecord.Fields(FIELD_CALCULATE_NO_OF_ITEMS))
        ActiveRecord.COPY_TO_NEXT = FNullField(rstRecord.Fields(FIELD_COPY_TO_NEXT))
        ActiveRecord.CHANGE_ONLY_IN_H1 = FNullField(rstRecord.Fields(FIELD_CHANGE_ONLY_IN_H1))
        ActiveRecord.SEND_ONLY_IN_H1 = FNullField(rstRecord.Fields(FIELD_SEND_ONLY_IN_H1))
        ActiveRecord.WIDTH_FIELD = FNullField(rstRecord.Fields(FIELD_WIDTH_FIELD))
        ActiveRecord.DECIMAL_FIELD = FNullField(rstRecord.Fields(FIELD_DECIMAL_FIELD))
        ActiveRecord.DATA_TYPE = FNullField(rstRecord.Fields(FIELD_DATA_TYPE))
        ActiveRecord.AUTO_ADD = FNullField(rstRecord.Fields(FIELD_AUTO_ADD))
        ActiveRecord.DEFAULT_VALUE = FNullField(rstRecord.Fields(FIELD_DEFAULT_VALUE))
        ActiveRecord.CALCULATE_CUSTOMS_VALUE = FNullField(rstRecord.Fields(FIELD_CALCULATE_CUSTOMS_VALUE))
        ActiveRecord.RELATE_L1_TO_S1 = FNullField(rstRecord.Fields(FIELD_RELATE_L1_TO_S1))
        ActiveRecord.BOX_TYPE = FNullField(rstRecord.Fields(FIELD_BOX_TYPE))
        ActiveRecord.BOX_INDEX = FNullField(rstRecord.Fields(FIELD_BOX_INDEX))
        ActiveRecord.INTERNAL_CODE = FNullField(rstRecord.Fields(FIELD_INTERNAL_CODE))
        
        ADORecordsetClose rstRecord
        GetRecord = True
    
    ElseIf ((ActiveRecord Is Nothing) = True) Then
        GetRecord = False
    End If
    
    Exit Function
    
ERROR_RECORDSET:
    ADORecordsetClose rstRecord
    GetRecord = False
    Exit Function
ERROR_QUERY:
    GetRecord = False
End Function

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiBoxDefEdIIE07
    End If

    strSql = SQL_GET_MAXID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.BOX_CODE = rstRecord.Fields("ID_MAX").Value
    
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBoxDefEdIIE07) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiBoxDefEdIIE07
    End If

    strSql = SQL_GET_MINID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.BOX_CODE = rstRecord.Fields("ID_MIN").Value
    
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

' conversion
Public Function GetTableRecord(ByRef ActiveRecord As cpiBoxDefEdIIE07) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset
    
    Set rstRecord = New ADODB.Recordset
    
    rstRecord.Open
    rstRecord.Fields.Append FIELD_BOX_CODE, 202, 3, 102
    rstRecord.Fields.Append FIELD_ENGLISH_DESCRIPTION, 202, 100, 102
    rstRecord.Fields.Append FIELD_DUTCH_DESCRIPTION, 202, 100, 102
    rstRecord.Fields.Append FIELD_FRENCH_DESCRIPTION, 202, 100, 102
    rstRecord.Fields.Append FIELD_EMPTY_FIELD_VALUE, 202, 50, 102
    rstRecord.Fields.Append FIELD_INSERT_FIELD, 2, 2, 118
    rstRecord.Fields.Append FIELD_JUSTIFY_FIELD, 202, 1, 102
    rstRecord.Fields.Append FIELD_SEQUENTIAL_TABBING_DEFAULT, 11, 2, 86
    rstRecord.Fields.Append FIELD_SEQUENTIAL_TABBING_ACTIVE, 11, 2, 86
    rstRecord.Fields.Append FIELD_CHECK_VAT, 11, 2, 86
    rstRecord.Fields.Append FIELD_CALCULATE_NET_WEIGHT, 11, 2, 86
    rstRecord.Fields.Append FIELD_CALCULATE_NO_OF_ITEMS, 11, 2, 86
    rstRecord.Fields.Append FIELD_COPY_TO_NEXT, 11, 2, 86
    rstRecord.Fields.Append FIELD_CHANGE_ONLY_IN_H1, 11, 2, 86
    rstRecord.Fields.Append FIELD_SEND_ONLY_IN_H1, 11, 2, 86
    rstRecord.Fields.Append FIELD_WIDTH_FIELD, 2, 2, 118
    rstRecord.Fields.Append FIELD_DECIMAL_FIELD, 2, 2, 118
    rstRecord.Fields.Append FIELD_DATA_TYPE, 202, 1, 102
    rstRecord.Fields.Append FIELD_AUTO_ADD, 11, 2, 86
    rstRecord.Fields.Append FIELD_DEFAULT_VALUE, 202, 255, 102
    rstRecord.Fields.Append FIELD_CALCULATE_CUSTOMS_VALUE, 11, 2, 86
    rstRecord.Fields.Append FIELD_RELATE_L1_TO_S1, 11, 2, 86
    rstRecord.Fields.Append FIELD_BOX_TYPE, 202, 15, 102
    rstRecord.Fields.Append FIELD_BOX_INDEX, 2, 2, 118
    rstRecord.Fields.Append FIELD_INTERNAL_CODE, 202, 50, 102
    
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_BOX_CODE) = ActiveRecord.BOX_CODE
    rstRecord.Fields(FIELD_ENGLISH_DESCRIPTION) = ActiveRecord.ENGLISH_DESCRIPTION
    rstRecord.Fields(FIELD_DUTCH_DESCRIPTION) = ActiveRecord.DUTCH_DESCRIPTION
    rstRecord.Fields(FIELD_FRENCH_DESCRIPTION) = ActiveRecord.FRENCH_DESCRIPTION
    rstRecord.Fields(FIELD_EMPTY_FIELD_VALUE) = ActiveRecord.EMPTY_FIELD_VALUE
    rstRecord.Fields(FIELD_INSERT_FIELD) = ActiveRecord.INSERT_FIELD
    rstRecord.Fields(FIELD_JUSTIFY_FIELD) = ActiveRecord.JUSTIFY_FIELD
    rstRecord.Fields(FIELD_SEQUENTIAL_TABBING_DEFAULT) = ActiveRecord.SEQUENTIAL_TABBING_DEFAULT
    rstRecord.Fields(FIELD_SEQUENTIAL_TABBING_ACTIVE) = ActiveRecord.SEQUENTIAL_TABBING_ACTIVE
    rstRecord.Fields(FIELD_CHECK_VAT) = ActiveRecord.CHECK_VAT
    rstRecord.Fields(FIELD_CALCULATE_NET_WEIGHT) = ActiveRecord.CALCULATE_NET_WEIGHT
    rstRecord.Fields(FIELD_CALCULATE_NO_OF_ITEMS) = ActiveRecord.CALCULATE_NO_OF_ITEMS
    rstRecord.Fields(FIELD_COPY_TO_NEXT) = ActiveRecord.COPY_TO_NEXT
    rstRecord.Fields(FIELD_CHANGE_ONLY_IN_H1) = ActiveRecord.CHANGE_ONLY_IN_H1
    rstRecord.Fields(FIELD_SEND_ONLY_IN_H1) = ActiveRecord.SEND_ONLY_IN_H1
    rstRecord.Fields(FIELD_WIDTH_FIELD) = ActiveRecord.WIDTH_FIELD
    rstRecord.Fields(FIELD_DECIMAL_FIELD) = ActiveRecord.DECIMAL_FIELD
    rstRecord.Fields(FIELD_DATA_TYPE) = ActiveRecord.DATA_TYPE
    rstRecord.Fields(FIELD_AUTO_ADD) = ActiveRecord.AUTO_ADD
    rstRecord.Fields(FIELD_DEFAULT_VALUE) = ActiveRecord.DEFAULT_VALUE
    rstRecord.Fields(FIELD_CALCULATE_CUSTOMS_VALUE) = ActiveRecord.CALCULATE_CUSTOMS_VALUE
    rstRecord.Fields(FIELD_RELATE_L1_TO_S1) = ActiveRecord.RELATE_L1_TO_S1
    rstRecord.Fields(FIELD_BOX_TYPE) = ActiveRecord.BOX_TYPE
    rstRecord.Fields(FIELD_BOX_INDEX) = ActiveRecord.BOX_INDEX
    rstRecord.Fields(FIELD_INTERNAL_CODE) = ActiveRecord.INTERNAL_CODE
    rstRecord.Update
    'INTERNAL_CODE
    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiBoxDefEdIIE07

    Dim clsRecord As cpiBoxDefEdIIE07
    
    Set clsRecord = New cpiBoxDefEdIIE07
    
    clsRecord.BOX_CODE = FNullField(ActiveRecord.Fields(FIELD_BOX_CODE))
    clsRecord.ENGLISH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_ENGLISH_DESCRIPTION))
    clsRecord.DUTCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_DUTCH_DESCRIPTION))
    clsRecord.FRENCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION))
    clsRecord.EMPTY_FIELD_VALUE = FNullField(ActiveRecord.Fields(FIELD_EMPTY_FIELD_VALUE))
    clsRecord.INSERT_FIELD = FNullField(ActiveRecord.Fields(FIELD_INSERT_FIELD))
    clsRecord.JUSTIFY_FIELD = FNullField(ActiveRecord.Fields(FIELD_JUSTIFY_FIELD))
    clsRecord.SEQUENTIAL_TABBING_DEFAULT = FNullField(ActiveRecord.Fields(FIELD_SEQUENTIAL_TABBING_DEFAULT))
    clsRecord.SEQUENTIAL_TABBING_ACTIVE = FNullField(ActiveRecord.Fields(FIELD_SEQUENTIAL_TABBING_ACTIVE))
    clsRecord.CHECK_VAT = FNullField(ActiveRecord.Fields(FIELD_CHECK_VAT))
    clsRecord.CALCULATE_NET_WEIGHT = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_NET_WEIGHT))
    clsRecord.CALCULATE_NO_OF_ITEMS = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_NO_OF_ITEMS))
    clsRecord.COPY_TO_NEXT = FNullField(ActiveRecord.Fields(FIELD_COPY_TO_NEXT))
    clsRecord.CHANGE_ONLY_IN_H1 = FNullField(ActiveRecord.Fields(FIELD_CHANGE_ONLY_IN_H1))
    clsRecord.SEND_ONLY_IN_H1 = FNullField(ActiveRecord.Fields(FIELD_SEND_ONLY_IN_H1))
    clsRecord.WIDTH_FIELD = FNullField(ActiveRecord.Fields(FIELD_WIDTH_FIELD))
    clsRecord.DECIMAL_FIELD = FNullField(ActiveRecord.Fields(FIELD_DECIMAL_FIELD))
    clsRecord.DATA_TYPE = FNullField(ActiveRecord.Fields(FIELD_DATA_TYPE))
    clsRecord.AUTO_ADD = FNullField(ActiveRecord.Fields(FIELD_AUTO_ADD))
    clsRecord.DEFAULT_VALUE = FNullField(ActiveRecord.Fields(FIELD_DEFAULT_VALUE))
    clsRecord.CALCULATE_CUSTOMS_VALUE = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_CUSTOMS_VALUE))
    clsRecord.RELATE_L1_TO_S1 = FNullField(ActiveRecord.Fields(FIELD_RELATE_L1_TO_S1))
    clsRecord.BOX_TYPE = FNullField(ActiveRecord.Fields(FIELD_BOX_TYPE))
    clsRecord.BOX_INDEX = FNullField(ActiveRecord.Fields(FIELD_BOX_INDEX))
    clsRecord.INTERNAL_CODE = FNullField(ActiveRecord.Fields(FIELD_INTERNAL_CODE))
'INTERNAL_CODE
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
        'ExecuteNonQuery ActiveConnection, strSql, lngRecordsAffected
        
        SearchValue = Not (rstDB.EOF And rstDB.BOF)
        
'        If (lngRecordsAffected > 0) Then
'            SearchValue = True
'        ElseIf (lngRecordsAffected = 0) Then
'            SearchValue = False
'        End If
        
    End If
    
    Exit Function
    
ERROR_SEARCH:
    SearchValue = False
End Function

Public Function GetRecordset(ByRef ActiveConnection As ADODB.Connection, Optional ByVal CommandText As String) As ADODB.Recordset

    Dim rstRecordset As ADODB.Recordset

    Set rstRecordset = New ADODB.Recordset

    On Error GoTo ERROR_RECORDSET
    
    If (CommandText = "") Then
        CommandText = SQL_GET_ALL_RECORD
    End If
    
    ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic

    Set GetRecordset = rstRecordset
    Set mvarRecordset = rstRecordset
    Set rstRecordset = Nothing

    Exit Function

ERROR_RECORDSET:
    Set rstRecordset = Nothing
End Function


' /* --------------- PRIVATE FUNCTIONS -------------------- */

Public Property Get Recordset() As ADODB.Recordset
Attribute Recordset.VB_UserMemId = 0
    Set Recordset = mvarRecordset
End Property

Private Sub Class_Terminate()
    Set mvarRecordset = Nothing
End Sub
