VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiTranshipments"
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

Private Const TABLE_NAME = "[DATA_NCTS_BERICHT_VERVOER_OVERLADING]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_NEW_TRANSPORT_MEANS_IDENTITY = "BL" ' 2
Private Const FIELD_NEW_TRANSPORT_MEANS_NATIONALITY = "BG" ' 3
Private Const FIELD_ENDORSEMENT_DATE = "BH" ' 3a
Private Const FIELD_ENDORSEMENT_AUTHORITY = "BI" ' 3b
Private Const FIELD_ENDORSEMENT_PLACE = "BJ" ' 3c
Private Const FIELD_ENDORSEMENT_COUNTRY = "BK" ' 3d
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" ' 4
Private Const FIELD_ORDINAL = "ORDINAL" ' 5

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [BL] , [BG], [BH], [BI], [BJ], [BK] , " & _
                                "[ORDINAL_PARENT],[ORDINAL]) VALUES "

Private Const SQL_GET_RECORD = "SELECT  [CODE] , [BL] , [BG], [BH], [BI], [BJ], [BK] , " & _
                                "[ORDINAL_PARENT] " & _
                                " FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.NEW_TRANSPORT_MEANS_IDENTITY)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.NEW_TRANSPORT_MEANS_NATIONALITY)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ENDORSEMENT_DATE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ENDORSEMENT_AUTHORITY)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ENDORSEMENT_PLACE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ENDORSEMENT_COUNTRY)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL)
        strSql = strSql & ")"

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
    
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End

        Dim intRecordCtr As Integer
        For intRecordCtr = 1 To ActiveRecord.Containers.Count
            ActiveRecord.Containers(intRecordCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Containers.AddRecord ActiveConnection, ActiveRecord.Containers(intRecordCtr)
        Next intRecordCtr
        
        AddRecord = True
        Exit Function

    End If

    AddRecord = False

    Exit Function

ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

'                        ",[" & FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & "" & _
'                        ",[" & FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & ""

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_NEW_TRANSPORT_MEANS_IDENTITY & "]=" & SetEnclosedChr(ActiveRecord.NEW_TRANSPORT_MEANS_IDENTITY) & _
                        ",[" & FIELD_NEW_TRANSPORT_MEANS_NATIONALITY & "]=" & SetEnclosedChr(ActiveRecord.NEW_TRANSPORT_MEANS_NATIONALITY) & "" & _
                        ",[" & FIELD_ENDORSEMENT_DATE & "]=" & SetEnclosedChr(ActiveRecord.ENDORSEMENT_DATE) & "" & _
                        ",[" & FIELD_ENDORSEMENT_AUTHORITY & "]=" & SetEnclosedChr(ActiveRecord.ENDORSEMENT_AUTHORITY) & "" & _
                        ",[" & FIELD_ENDORSEMENT_PLACE & "]=" & SetEnclosedChr(ActiveRecord.ENDORSEMENT_PLACE) & "" & _
                        ",[" & FIELD_ENDORSEMENT_COUNTRY & "]=" & SetEnclosedChr(ActiveRecord.ENDORSEMENT_COUNTRY) & ""

        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & " AND [ORDINAL_PARENT]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)


        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End
        
        Dim intRecordCtr As Integer
        For intRecordCtr = 1 To ActiveRecord.Containers.Count
            ActiveRecord.Containers.ModifyRecord ActiveConnection, ActiveRecord.Containers(intRecordCtr)
        Next intRecordCtr

        ModifyRecord = True
        Exit Function

    End If

    ModifyRecord = False

    Exit Function

ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & " AND [ORDINAL_PARENT]=" & ActiveRecord.ORDINAL_PARENT

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        
        'Set rstRecord = ActiveConnection.Execute(strSql)
        ' MUCP-159 - End

        On Error GoTo ERROR_RECORDSET
        ActiveRecord.NEW_TRANSPORT_MEANS_IDENTITY = FNullField(rstRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_IDENTITY))
        ActiveRecord.NEW_TRANSPORT_MEANS_NATIONALITY = FNullField(rstRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_NATIONALITY))
        ActiveRecord.ENDORSEMENT_DATE = FNullField(rstRecord.Fields(FIELD_ENDORSEMENT_DATE))
        ActiveRecord.ENDORSEMENT_AUTHORITY = FNullField(rstRecord.Fields(FIELD_ENDORSEMENT_AUTHORITY))
        ActiveRecord.ENDORSEMENT_PLACE = FNullField(rstRecord.Fields(FIELD_ENDORSEMENT_PLACE))
        ActiveRecord.ENDORSEMENT_COUNTRY = FNullField(rstRecord.Fields(FIELD_ENDORSEMENT_COUNTRY))
        'ActiveRecord.ORDINAL_PARENT = FNullField(rstRecord.Fields(FIELD_ORDINAL_PARENT))
        'ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))
        
        
        ' add containers here 99x
        Dim strKey As String
        Dim strParentKey As String
        
        ' :-)
        strSql = "SELECT [CODE],[ORDINAL],[ORDINAL_PARENT] FROM " & TBL_DATA_LEVEL1_2_2_1 & _
                    " WHERE [CODE]='" & ActiveRecord.CODE_FIELD & "'  AND [ORDINAL_PARENT]=" & CStr(ActiveRecord.ORDINAL_PARENT)
        Set ActiveRecord.Containers = New cpiContainers
        Set ActiveRecord.Containers.ActiveRecordset = ActiveRecord.Containers.GetRecordset(ActiveConnection, strSql)
        
        Dim intOrdinal As Integer
        intOrdinal = 0
        
        strParentKey = ActiveRecord.Key
        Do While (ActiveRecord.Containers.ActiveRecordset.EOF = False)
            'strKey = ActiveRecord.Seals.ActiveRecordset.Fields("CODE").Value & "_" & ActiveRecord.Seals.ActiveRecordset.Fields("ORDINAL").Value
            intOrdinal = intOrdinal + 1
            'strKey = strParentKey & "_1.2.2.1-" & ActiveRecord.Containers.ActiveRecordset.Fields("ORDINAL").Value  'ActiveRecord.Containers.ActiveRecordset.AbsolutePosition
            strKey = strParentKey & "_1.2.2.1-" & intOrdinal   'ActiveRecord.Containers.ActiveRecordset.AbsolutePosition
            
            ActiveRecord.Containers.Add strKey, ActiveRecord.Containers.ActiveRecordset.Fields("CODE").Value
            ActiveRecord.Containers(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Containers(strKey).ORDINAL_PARENT = ActiveRecord.Containers.ActiveRecordset.Fields("ORDINAL_PARENT").Value
            ActiveRecord.Containers(strKey).ORDINAL = ActiveRecord.Containers.ActiveRecordset.Fields("ORDINAL").Value
            
            ActiveRecord.Containers.GetRecord ActiveConnection, ActiveRecord.Containers(strKey)
            ActiveRecord.Containers(strKey).ORDINAL = intOrdinal
            
            ActiveRecord.Containers.ActiveRecordset.MoveNext
        Loop
        
        
        'ActiveRecord.EnRouteEvents(9).
'        Private mvarHeaders As cpiTranshipments ' 6
'        Private mvarEnRouteEvents As cpiTranshipments  ' 7
'        Private mvarCustomOffices As cpiSealInfos  ' 8
'        Private mvarTraders As cpiTranshipments  ' 9
        
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiTranshipment
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiTranshipment) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiTranshipment
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiTranshipment) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_NEW_TRANSPORT_MEANS_IDENTITY, 202, 17, 102 ' 2
    rstRecord.Fields.Append FIELD_NEW_TRANSPORT_MEANS_NATIONALITY, 202, 35, 102 ' 3
    rstRecord.Fields.Append FIELD_ENDORSEMENT_DATE, 202, 35, 102 ' 4
    rstRecord.Fields.Append FIELD_ENDORSEMENT_AUTHORITY, 202, 2, 102 ' 5
    rstRecord.Fields.Append FIELD_ENDORSEMENT_PLACE, 202, 9, 102 ' 6
    rstRecord.Fields.Append FIELD_ENDORSEMENT_COUNTRY, 202, 35, 102 ' 7
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 8
    rstRecord.Fields.Append FIELD_ORDINAL_PARENT, 3, 4, 118 ' 9
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_IDENTITY) = ActiveRecord.NEW_TRANSPORT_MEANS_IDENTITY
    rstRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_NATIONALITY) = ActiveRecord.NEW_TRANSPORT_MEANS_NATIONALITY
    rstRecord.Fields(FIELD_ENDORSEMENT_DATE) = ActiveRecord.ENDORSEMENT_DATE
    rstRecord.Fields(FIELD_ENDORSEMENT_AUTHORITY) = ActiveRecord.ENDORSEMENT_AUTHORITY
    rstRecord.Fields(FIELD_ENDORSEMENT_PLACE) = ActiveRecord.ENDORSEMENT_PLACE
    rstRecord.Fields(FIELD_ENDORSEMENT_COUNTRY) = ActiveRecord.ENDORSEMENT_COUNTRY
    rstRecord.Fields(FIELD_ORDINAL_PARENT) = ActiveRecord.ORDINAL_PARENT
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiTranshipment

    Dim clsRecord As cpiTranshipment

    Set clsRecord = New cpiTranshipment

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.NEW_TRANSPORT_MEANS_IDENTITY = FNullField(ActiveRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_IDENTITY))
    clsRecord.NEW_TRANSPORT_MEANS_NATIONALITY = FNullField(ActiveRecord.Fields(FIELD_NEW_TRANSPORT_MEANS_NATIONALITY))
    clsRecord.ENDORSEMENT_DATE = FNullField(ActiveRecord.Fields(FIELD_ENDORSEMENT_DATE))
    clsRecord.ENDORSEMENT_AUTHORITY = FNullField(ActiveRecord.Fields(FIELD_ENDORSEMENT_AUTHORITY))
    clsRecord.ENDORSEMENT_PLACE = FNullField(ActiveRecord.Fields(FIELD_ENDORSEMENT_PLACE))
    clsRecord.ENDORSEMENT_COUNTRY = FNullField(ActiveRecord.Fields(FIELD_ENDORSEMENT_COUNTRY))
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













Public Function Add(ByRef sKey As String, ByVal PKValue As String, ByVal ParentID As Integer) As cpiTranshipment
    
    'create a new object
    Dim objNewMember As cpiTranshipment
    Set objNewMember = New cpiTranshipment

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    objNewMember.CODE_FIELD = PKValue
    objNewMember.ORDINAL = mCol.Count + 1
    objNewMember.ORDINAL_PARENT = ParentID
    
    Set objNewMember.Containers = New cpiContainers
    'objNewMember.Containers.Add sKey & "_" & CStr(objNewMember.Containers.Count + 1), PKValue
    objNewMember.Containers.AddNewGroup sKey & "_1.2.2.1", PKValue, ParentID
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiTranshipment
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

