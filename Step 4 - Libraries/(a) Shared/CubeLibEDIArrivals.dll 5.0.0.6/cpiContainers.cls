VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiContainers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
Option Explicit

Dim mvarActiveRecordset As ADODB.Recordset

Dim mvarListIndex As Integer
Dim mvarGroupIndex As Integer
Dim mvarGroupCount As Integer
Dim mvarItemsPerGroup As Integer

Private mCol As Collection

Private Const TABLE_NAME = "[DATA_NCTS_BERICHT_VERVOER_OVERLADING_CONTAINER]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_NEW_CONTAINER_NUMBER = "SC" ' 2
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" ' 4
Private Const FIELD_ORDINAL = "ORDINAL" ' 5

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [SC] , " & _
                                "[ORDINAL_PARENT],[ORDINAL]) VALUES "

Private Const SQL_GET_RECORD = "SELECT  [CODE] , [SC] , " & _
                                "[ORDINAL_PARENT] " & _
                                " FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.NEW_CONTAINER_NUMBER)
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

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

    Dim strSql As String
    Dim rstDB As ADODB.Recordset
    Dim lngRecordsAffected As Long

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_NEW_CONTAINER_NUMBER & "]=" & SetEnclosedChr(ActiveRecord.NEW_CONTAINER_NUMBER)

'                        ",[" & FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & ""
                        
        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD) & " AND [" & _
                            FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & " AND [" & _
                            FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & ""

'                        ",[" & FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & ""

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        lngRecordsAffected = ExecuteNonQuery(ActiveConnection, strSql)
        
        'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
        ' MUCP-159 - End
        
        If (lngRecordsAffected > 0) Then
            
            ModifyRecord = True
        
        ElseIf (lngRecordsAffected = 0) Then
        
            ' add here
            'AddRecord ActiveConnection, ActiveRecord
            
        End If
        
        
        Exit Function

    End If

    ModifyRecord = False

    Exit Function

ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD) & " AND [ORDINAL]= " & _
                    SetEnclosedChr(ActiveRecord.ORDINAL) & " AND [ORDINAL_PARENT]= " & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        'Set rstRecord = ActiveConnection.Execute(strSql)
        ' MUCP-159 - End

        On Error GoTo ERROR_RECORDSET
        
        'If (rstRecord.EOF = False) Then
            ActiveRecord.NEW_CONTAINER_NUMBER = FNullField(rstRecord.Fields(FIELD_NEW_CONTAINER_NUMBER))
            'ActiveRecord.ORDINAL_PARENT = FNullField(rstRecord.Fields(FIELD_ORDINAL_PARENT))
        'End If
        'ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))

        
        'ActiveRecord.EnRouteEvents(9).
'        Private mvarHeaders As cpiContainers ' 6
'        Private mvarEnRouteEvents As cpiContainers  ' 7
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiContainer
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiContainer) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiContainer
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiContainer) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_NEW_CONTAINER_NUMBER, 202, 8, 102 ' 2
    rstRecord.Fields.Append FIELD_ORDINAL_PARENT, 3, 4, 118 ' 3
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 4
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_NEW_CONTAINER_NUMBER) = ActiveRecord.NEW_CONTAINER_NUMBER
    rstRecord.Fields(FIELD_ORDINAL_PARENT) = ActiveRecord.ORDINAL_PARENT
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiContainer

    Dim clsRecord As cpiContainer

    Set clsRecord = New cpiContainer

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.NEW_CONTAINER_NUMBER = FNullField(ActiveRecord.Fields(FIELD_NEW_CONTAINER_NUMBER))
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
        ' MUCP-159 - End
        
    End If

    Exit Function

ERROR_SEARCH:
    SearchValue = False
End Function

Public Function GetRecordset(ByRef ActiveConnection As ADODB.Connection, Optional ByVal CommandText As String) As ADODB.Recordset

    Dim rstRecordset As ADODB.Recordset

    ' MUCP-159 - Start
    'Set rstRecordset = New ADODB.Recordset
    ' MUCP-159 - End
    
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





Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiContainer
    
'    AddNewGroup sKey, PKValue
    
    'create a new object
    Dim objNewMember As cpiContainer
    Set objNewMember = New cpiContainer

    'set the properties passed into the method
    objNewMember.Key = sKey
    'objNewMember.NEW_CONTAINER_NUMBER = NewContainerNumber
    objNewMember.Index = mCol.Count + 1
    objNewMember.CODE_FIELD = PKValue

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function AddNewGroup(ByVal sParentKey As String, ByVal PKValue As String, ByVal ParentID As Integer)
    
    'create a new object
    Dim intGroupCtr As Integer
    Dim objNewMember As cpiContainer
    Dim sKey As String

    
    For intGroupCtr = 1 To mvarItemsPerGroup
        
        Set objNewMember = New cpiContainer
    
        sKey = sParentKey & "-" & CStr(mCol.Count + 1)
        'set the properties passed into the method
        objNewMember.Key = sKey '& "_" & CStr(mCol.Count + 1)
        'objNewMember.NEW_CONTAINER_NUMBER = NewContainerNumber
        objNewMember.Index = mCol.Count + 1
        objNewMember.CODE_FIELD = PKValue
        objNewMember.ORDINAL = mCol.Count + 1
        objNewMember.ORDINAL_PARENT = ParentID
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    
        'return the object created
        'Set Add = objNewMember

    Next intGroupCtr
    
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiContainer
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
    mvarItemsPerGroup = 5
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarActiveRecordset = Nothing

End Sub

Public Property Get ActiveRecordset() As ADODB.Recordset
    Set ActiveRecordset = mvarActiveRecordset
End Property

Public Property Set ActiveRecordset(ByRef vNewValue As ADODB.Recordset)
    Set mvarActiveRecordset = vNewValue
End Property


Public Property Get ListIndex() As Integer
    ListIndex = mvarListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Integer)
    mvarListIndex = vNewValue
End Property
'Dim mvarListIndex As Integer

Public Property Get ItemsPerGroup() As Integer
    ItemsPerGroup = mvarItemsPerGroup
End Property

Public Property Let ItemsPerGroup(ByVal vNewValue As Integer)
    mvarItemsPerGroup = vNewValue
End Property
'Dim mvarListIndex As Integer

Public Property Get GroupIndex() As Integer
    GroupIndex = mvarGroupIndex
End Property

Public Property Let GroupIndex(ByVal vNewValue As Integer)

    If (vNewValue <= GroupCount) And (vNewValue > 0) Then
        mvarGroupIndex = vNewValue
    End If
End Property

Public Property Get GroupCount() As Integer

    Dim intGroupTotal As Integer
    Dim intGroupCtr As Integer
    
    Do
        intGroupTotal = intGroupTotal + mvarItemsPerGroup
        intGroupCtr = intGroupCtr + 1
    Loop While (mCol.Count - intGroupTotal) > 0
    
    GroupCount = intGroupCtr
End Property

Public Property Let GroupCount(ByVal vNewValue As Integer)
    mvarGroupCount = vNewValue
End Property


