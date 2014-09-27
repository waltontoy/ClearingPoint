VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
Option Explicit
'
Private Const TBL_DATA_ROOT = "DATA" & "_" & "NCTS"
Private Const TBL_DATA_LEVEL1 = TBL_DATA_ROOT & "_" & "BERICHT"

Private Const TABLE_NAME = "[" & TBL_DATA_LEVEL1 & "]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_NCTS_IEM_ID = "NCTS_IEM_ID" ' 2
Private Const FIELD_MD_FIELD = "MD" ' 3
Private Const FIELD_ME_FIELD = "ME" ' 4
Private Const FIELD_ORDINAL = "ORDINAL" ' 5

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [NCTS_IEM_ID] , [MD] , " & _
                                "[ME], [ORDINAL]) VALUES "

Private Const SQL_GET_RECORD = "SELECT [CODE] , [NCTS_IEM_ID] , [MD] , " & _
                                "[ME], [ORDINAL] FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

Dim mvarRecordset As ADODB.Recordset

Private mCol As Collection

Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiMessage
    
    'create a new object
    Dim objNewMember As cpiMessage
    Set objNewMember = New cpiMessage

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    objNewMember.CODE_FIELD = PKValue
    objNewMember.ORDINAL = mCol.Count + 1
    
    Set objNewMember.Headers = New PCubeLibEDIArrivals.cpiHeaders
    objNewMember.Headers.Add sKey & "_1.1-" & CStr(objNewMember.Headers.Count + 1), PKValue
    
    Set objNewMember.EnRouteEvents = New cpiEnRouteEvents
    'objNewMember.EnRouteEvents.Add sKey & "_1.2-" & CStr(objNewMember.EnRouteEvents.Count + 1), PKValue
    
    Set objNewMember.CustomOffices = New cpiCustomOffices
    objNewMember.CustomOffices.Add sKey & "_1.3-" & CStr(objNewMember.CustomOffices.Count + 1), PKValue
    
    Set objNewMember.Traders = New cpiTraders
    objNewMember.Traders.Add sKey & "_1.4-" & CStr(objNewMember.Traders.Count + 1), PKValue
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiMessage
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
    Set mvarRecordset = Nothing

End Sub

'Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByVal strCode As String) As Boolean
''
'    'const DATA_NCTS_BERICHT
'
'    Dim blnFound As Boolean
'
'    ' load Ordinal from parent DATA_NCTS_BERICHT bericht
'    SearchRecord ActiveConnection, TBL_DATA_LEVEL1, "CODE", "'" & strCode & "'"
'
'
''
'End Function

' DB part -------------



' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.NCTS_IEM_ID)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DATE_OF_PREPARATION)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.TIME_OF_PREPARATION)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL)
        strSql = strSql & ")"

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
    
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End
        
        Dim intEventCtr As Integer
        
        For intEventCtr = 1 To ActiveRecord.Headers.Count
            ActiveRecord.Headers(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Headers.AddRecord ActiveConnection, ActiveRecord.Headers(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.CustomOffices.Count
            ActiveRecord.CustomOffices(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.CustomOffices.AddRecord ActiveConnection, ActiveRecord.CustomOffices(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.Traders.Count
            ActiveRecord.Traders(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Traders.AddRecord ActiveConnection, ActiveRecord.Traders(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.EnRouteEvents.Count
            ActiveRecord.EnRouteEvents(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.EnRouteEvents.AddRecord ActiveConnection, ActiveRecord.EnRouteEvents(intEventCtr)
        Next intEventCtr


        AddRecord = True
        Exit Function

    End If

    AddRecord = False

    Exit Function

ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_NCTS_IEM_ID & "]=" & SetEnclosedChr(ActiveRecord.NCTS_IEM_ID) & _
                        ",[" & FIELD_MD_FIELD & "]=" & SetEnclosedChr(ActiveRecord.DATE_OF_PREPARATION) & "" & _
                        ",[" & FIELD_ME_FIELD & "]=" & SetEnclosedChr(ActiveRecord.TIME_OF_PREPARATION) & "" & _
                        ",[" & FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & ""
        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End
        
        Dim intEventCtr As Integer
        
        For intEventCtr = 1 To ActiveRecord.Headers.Count
            ActiveRecord.Headers.ModifyRecord ActiveConnection, ActiveRecord.Headers(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.CustomOffices.Count
            ActiveRecord.CustomOffices.ModifyRecord ActiveConnection, ActiveRecord.CustomOffices(intEventCtr)
        Next intEventCtr
            
        For intEventCtr = 1 To ActiveRecord.Traders.Count
            ActiveRecord.Traders.ModifyRecord ActiveConnection, ActiveRecord.Traders(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.EnRouteEvents.Count
            ActiveRecord.EnRouteEvents.ModifyRecord ActiveConnection, ActiveRecord.EnRouteEvents(intEventCtr)
        Next intEventCtr

        ModifyRecord = True
        Exit Function

    End If

    ModifyRecord = False

    Exit Function

ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

    Dim strKey As String
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
        ActiveRecord.NCTS_IEM_ID = FNullField(rstRecord.Fields(FIELD_NCTS_IEM_ID))
        ActiveRecord.DATE_OF_PREPARATION = FNullField(rstRecord.Fields(FIELD_MD_FIELD))
        ActiveRecord.TIME_OF_PREPARATION = FNullField(rstRecord.Fields(FIELD_ME_FIELD))
        ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))
        
        'Dim clsHeader As cpiHeader
        'clsHeader.CODE_FIELD = ActiveRecord.Headers
        
        ' add new header here 1x 1.1
        Set ActiveRecord.Headers = New cpiHeaders
        strKey = ActiveRecord.Key & "_1.1-" & CStr(ActiveRecord.Headers.Count + 1)
        ActiveRecord.Headers.Add strKey, ActiveRecord.CODE_FIELD
        ActiveRecord.Headers(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
        ActiveRecord.Headers.GetRecord ActiveConnection, ActiveRecord.Headers(strKey)
        
        ' add new customs here 1x  1.3
        Set ActiveRecord.CustomOffices = New cpiCustomOffices
        strKey = ActiveRecord.Key & "_1.3-" & CStr(ActiveRecord.CustomOffices.Count + 1)
        ActiveRecord.CustomOffices.Add strKey, ActiveRecord.CODE_FIELD
        ActiveRecord.CustomOffices(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
        ActiveRecord.CustomOffices.GetRecord ActiveConnection, ActiveRecord.CustomOffices(strKey)
        
        ' add trader destination here 1x  1.4
        Set ActiveRecord.Traders = New cpiTraders
        strKey = ActiveRecord.Key & "_1.4-" & CStr(ActiveRecord.Traders.Count + 1)
        ActiveRecord.Traders.Add strKey, ActiveRecord.CODE_FIELD
        ActiveRecord.Traders(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
        ActiveRecord.Traders.GetRecord ActiveConnection, ActiveRecord.Traders(strKey)
        
        ' add en route event here 9x
        'Dim strSql As String
        
        strSql = "SELECT [CODE],[ORDINAL] FROM " & TBL_DATA_LEVEL1_2 & " WHERE [CODE]='" & ActiveRecord.CODE_FIELD & "'"
        
        Set ActiveRecord.EnRouteEvents = New cpiEnRouteEvents
        Set ActiveRecord.EnRouteEvents.ActiveRecordset = ActiveRecord.EnRouteEvents.GetRecordset(ActiveConnection, strSql)
        
        ' 1.2
        Do While (ActiveRecord.EnRouteEvents.ActiveRecordset.EOF = False)
            
            strKey = ActiveRecord.Key & "_1.2-" & CStr(ActiveRecord.EnRouteEvents.Count + 1)
            ActiveRecord.EnRouteEvents.Add strKey, ActiveRecord.EnRouteEvents.ActiveRecordset.Fields("CODE").Value
            
            ActiveRecord.EnRouteEvents(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.EnRouteEvents(strKey).ORDINAL = ActiveRecord.EnRouteEvents.ActiveRecordset.Fields("ORDINAL").Value
            
            ActiveRecord.EnRouteEvents.GetRecord ActiveConnection, ActiveRecord.EnRouteEvents(strKey)
            ActiveRecord.EnRouteEvents.ActiveRecordset.MoveNext
        Loop
        
        'ActiveRecord.EnRouteEvents(9).
'        Private mvarHeaders As cpiHeaders ' 6
'        Private mvarEnRouteEvents As cpiEnRouteEvents  ' 7
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiMessage
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiMessage) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiMessage
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiMessage) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_NCTS_IEM_ID, 3, 4, 118 ' 2
    rstRecord.Fields.Append FIELD_MD_FIELD, 202, 6, 102 ' 3
    rstRecord.Fields.Append FIELD_ME_FIELD, 202, 4, 102 ' 4
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 5
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_NCTS_IEM_ID) = ActiveRecord.NCTS_IEM_ID
    rstRecord.Fields(FIELD_MD_FIELD) = ActiveRecord.DATE_OF_PREPARATION
    rstRecord.Fields(FIELD_ME_FIELD) = ActiveRecord.TIME_OF_PREPARATION
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiMessage

    Dim clsRecord As cpiMessage

    Set clsRecord = New cpiMessage

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.NCTS_IEM_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_ID))
    clsRecord.DATE_OF_PREPARATION = FNullField(ActiveRecord.Fields(FIELD_MD_FIELD))
    clsRecord.TIME_OF_PREPARATION = FNullField(ActiveRecord.Fields(FIELD_ME_FIELD))
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
    Set mvarRecordset = rstRecordset
    Set rstRecordset = Nothing

    Exit Function

ERROR_RECORDSET:
    Set rstRecordset = Nothing
End Function


' /* --------------- PRIVATE FUNCTIONS -------------------- */

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mvarRecordset
End Property

'Private Sub Class_Terminate()
'    Set mvarRecordset = Nothing
'End Sub



' ****** ****** ****** ****** ****** ****** ****** ****** ****** ******

'Private Const TBL_DATA_ROOT = "DATA" & "_" & "NCTS"
'
'Private Const TBL_DATA_LEVEL1 = TBL_DATA_ROOT & "_" & "BERICHT"
'
'Private Const TBL_DATA_LEVEL1_1 = TBL_DATA_LEVEL1 & "_" & "HOOFDING"
'Private Const TBL_DATA_LEVEL1_2 = TBL_DATA_LEVEL1 & "_" & "VERVOER"
'Private Const TBL_DATA_LEVEL1_3 = TBL_DATA_LEVEL1 & "_" & "DOUANEKANTOOR"
'Private Const TBL_DATA_LEVEL1_4 = TBL_DATA_LEVEL1 & "_" & "HANDELAAR"
''
'Private Const TBL_DATA_LEVEL1_2_1 = TBL_DATA_LEVEL1_2 & "_" & "INCIDENT"
'Private Const TBL_DATA_LEVEL1_2_2 = TBL_DATA_LEVEL1_2 & "_" & "OVERLADING"
'Private Const TBL_DATA_LEVEL1_2_3 = TBL_DATA_LEVEL1_2 & "_" & "VERZEGELING_INFO"
'Private Const TBL_DATA_LEVEL1_2_4 = TBL_DATA_LEVEL1_2 & "_" & "CONTROLE"
''
'Private Const TBL_DATA_LEVEL1_2_2_1 = TBL_DATA_LEVEL1_2_2 & "_" & "CONTAINER"
'Private Const TBL_DATA_LEVEL1_2_3_1 = TBL_DATA_LEVEL1_2_2 & "_" & "ID"


