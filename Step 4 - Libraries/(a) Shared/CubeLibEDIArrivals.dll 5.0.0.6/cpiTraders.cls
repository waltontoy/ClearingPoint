VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiTraders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
Option Explicit

Private mCol As Collection

Private Const TABLE_NAME = "[DATA_NCTS_BERICHT_HANDELAAR]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_DESTINATION_TIN = "W8" ' 2
Private Const FIELD_DESTINATION_NAME = "W9" ' 3
Private Const FIELD_DESTINATION_STREET_AND_NUMBER = "WA" ' 3a
Private Const FIELD_DESTINATION_COUNTRY_CODE = "WE" ' 3b
Private Const FIELD_DESTINATION_POSTAL_CODE = "WD" ' 3c
Private Const FIELD_DESTINATION_CITY = "WB" ' 3d
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" ' 4
Private Const FIELD_ORDINAL = "ORDINAL" ' 5

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [W8] , [W9], [WA], [WE], [WD], [WB] , " & _
                                "[ORDINAL_PARENT],[ORDINAL]) VALUES "

Private Const SQL_GET_RECORD = "SELECT  [CODE] , [W8] , [W9], [WA], [WE], [WD], [WB] , " & _
                                "[ORDINAL_PARENT],[ORDINAL]" & _
                                " FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_TIN)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_NAME)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_STREET_AND_NUMBER)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_COUNTRY_CODE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_POSTAL_CODE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DESTINATION_CITY)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL)
        strSql = strSql & ")"

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
   
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End

        AddRecord = True
        Exit Function

    End If

    AddRecord = False

    Exit Function

ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End

        DeleteRecord = True
        Exit Function

    End If

    DeleteRecord = False

    Exit Function

ERROR_QUERY:
    DeleteRecord = False
End Function

' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_DESTINATION_TIN & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_TIN) & _
                        ",[" & FIELD_DESTINATION_NAME & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_NAME) & "" & _
                        ",[" & FIELD_DESTINATION_STREET_AND_NUMBER & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_STREET_AND_NUMBER) & "" & _
                        ",[" & FIELD_DESTINATION_COUNTRY_CODE & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_COUNTRY_CODE) & "" & _
                        ",[" & FIELD_DESTINATION_POSTAL_CODE & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_POSTAL_CODE) & "" & _
                        ",[" & FIELD_DESTINATION_CITY & "]=" & SetEnclosedChr(ActiveRecord.DESTINATION_CITY) & "" & _
                        ",[" & FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & "" & _
                        ",[" & FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & ""
        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End

        ModifyRecord = True
        Exit Function

    End If

    ModifyRecord = False

    Exit Function

ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        
        'Set rstRecord = ActiveConnection.Execute(strSql)
        ' MUCP-159 - End

        On Error GoTo ERROR_RECORDSET
        ActiveRecord.DESTINATION_TIN = FNullField(rstRecord.Fields(FIELD_DESTINATION_TIN))
        ActiveRecord.DESTINATION_NAME = FNullField(rstRecord.Fields(FIELD_DESTINATION_NAME))
        ActiveRecord.DESTINATION_STREET_AND_NUMBER = FNullField(rstRecord.Fields(FIELD_DESTINATION_STREET_AND_NUMBER))
        ActiveRecord.DESTINATION_COUNTRY_CODE = FNullField(rstRecord.Fields(FIELD_DESTINATION_COUNTRY_CODE))
        ActiveRecord.DESTINATION_POSTAL_CODE = FNullField(rstRecord.Fields(FIELD_DESTINATION_POSTAL_CODE))
        ActiveRecord.DESTINATION_CITY = FNullField(rstRecord.Fields(FIELD_DESTINATION_CITY))
        ActiveRecord.ORDINAL_PARENT = FNullField(rstRecord.Fields(FIELD_ORDINAL_PARENT))
        ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))
        
        'ActiveRecord.EnRouteEvents(9).
'        Private mvarHeaders As cpiTraders ' 6
'        Private mvarEnRouteEvents As cpiTraders  ' 7
'        Private mvarCustomOffices As cpiSealInfos  ' 8
'        Private mvarTraders As cpiTraders  ' 9
        
        Set rstRecord = Nothing
        GetRecord = True

    ElseIf ((ActiveRecord Is Nothing) = True) Then
        GetRecord = False
    End If

    Exit Function

ERROR_RECORDSET:
    Set rstRecord = Nothing
    GetRecord = False
    Exit Function
ERROR_QUERY:
    GetRecord = False
End Function

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiTrader
    End If

    strSql = SQL_GET_MAXID

    On Error GoTo ERROR_QUERY
    
    ' MUCP-159 - Start
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        
    'Set rstRecord = ActiveConnection.Execute(strSql)
    ' MUCP-159 - End

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.CODE_FIELD = rstRecord.Fields("ID_MAX").Value

    Set rstRecord = Nothing
    GetMaxID = True
    Exit Function

ERROR_RECORDSET:
    Set rstRecord = Nothing
    GetMaxID = False
    Exit Function
ERROR_QUERY:
    GetMaxID = False
End Function

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTrader) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiTrader
    End If

    strSql = SQL_GET_MINID

    On Error GoTo ERROR_QUERY
    
    ' MUCP-159 - Start
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    
    'Set rstRecord = ActiveConnection.Execute(strSql)
    ' MUCP-159 - End

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.CODE_FIELD = rstRecord.Fields("ID_MIN").Value

    Set rstRecord = Nothing
    GetMinID = True
    Exit Function

ERROR_RECORDSET:
    Set rstRecord = Nothing
    GetMinID = False
    Exit Function
ERROR_QUERY:
    GetMinID = False
End Function

' conversion
Public Function GetTableRecord(ByRef ActiveRecord As cpiTrader) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_DESTINATION_TIN, 202, 17, 102 ' 2
    rstRecord.Fields.Append FIELD_DESTINATION_NAME, 202, 35, 102 ' 3
    rstRecord.Fields.Append FIELD_DESTINATION_STREET_AND_NUMBER, 202, 35, 102 ' 4
    rstRecord.Fields.Append FIELD_DESTINATION_COUNTRY_CODE, 202, 2, 102 ' 5
    rstRecord.Fields.Append FIELD_DESTINATION_POSTAL_CODE, 202, 9, 102 ' 6
    rstRecord.Fields.Append FIELD_DESTINATION_CITY, 202, 35, 102 ' 7
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 8
    rstRecord.Fields.Append FIELD_ORDINAL_PARENT, 3, 4, 118 ' 9
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_DESTINATION_TIN) = ActiveRecord.DESTINATION_TIN
    rstRecord.Fields(FIELD_DESTINATION_NAME) = ActiveRecord.DESTINATION_NAME
    rstRecord.Fields(FIELD_DESTINATION_STREET_AND_NUMBER) = ActiveRecord.DESTINATION_STREET_AND_NUMBER
    rstRecord.Fields(FIELD_DESTINATION_COUNTRY_CODE) = ActiveRecord.DESTINATION_COUNTRY_CODE
    rstRecord.Fields(FIELD_DESTINATION_POSTAL_CODE) = ActiveRecord.DESTINATION_POSTAL_CODE
    rstRecord.Fields(FIELD_DESTINATION_CITY) = ActiveRecord.DESTINATION_CITY
    rstRecord.Fields(FIELD_ORDINAL_PARENT) = ActiveRecord.ORDINAL_PARENT
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiTrader

    Dim clsRecord As cpiTrader

    Set clsRecord = New cpiTrader

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.DESTINATION_TIN = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_TIN))
    clsRecord.DESTINATION_NAME = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_NAME))
    clsRecord.DESTINATION_STREET_AND_NUMBER = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_STREET_AND_NUMBER))
    clsRecord.DESTINATION_COUNTRY_CODE = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_COUNTRY_CODE))
    clsRecord.DESTINATION_POSTAL_CODE = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_POSTAL_CODE))
    clsRecord.DESTINATION_CITY = FNullField(ActiveRecord.Fields(FIELD_DESTINATION_CITY))
    clsRecord.ORDINAL_PARENT = FNullField(ActiveRecord.Fields(FIELD_ORDINAL_PARENT))
    clsRecord.ORDINAL = FNullField(ActiveRecord.Fields(FIELD_ORDINAL))
    
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

        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TABLE_NAME & " WHERE " & SearchField & "=" & SetEnclosedChr((SearchValue))
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
        
        SearchRecord = Not (rstDB.EOF And rstDB.BOF)
        
        'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
        '
        'If (rstDB.EOF = False) Then
        '    SearchRecord = True
        'ElseIf (rstDB.EOF = True) Then
        '    SearchRecord = False
        'End If
        ' MUCP-159 - Start
        
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
        CommandText = SQL_GET_RECORD
    End If

    ' MUCP-159 - Start
    ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic
        
    'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic
    ' MUCP-159 - End

    Set GetRecordset = rstRecordset
    'Set mvarRecordset = rstRecordset
    Set rstRecordset = Nothing

    Exit Function

ERROR_RECORDSET:
    Set rstRecordset = Nothing
End Function






Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiTrader
    
    'create a new object
    Dim objNewMember As cpiTrader
    Set objNewMember = New cpiTrader

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    objNewMember.CODE_FIELD = PKValue
    objNewMember.ORDINAL = mCol.Count + 1

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiTrader
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

