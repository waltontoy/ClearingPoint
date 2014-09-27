VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEnRouteEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
Option Explicit

Dim mvarListIndex As Integer
Private mCol As Collection
Private mvarActiveRecordset As ADODB.Recordset

Private Const TABLE_NAME = "[DATA_NCTS_BERICHT_VERVOER]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_PLACE_FIELD = "BF" ' 2
Private Const FIELD_COUNTRY_CODE = "C7" ' 3
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" ' 4
Private Const FIELD_ORDINAL = "ORDINAL" ' 5

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [BF] , [C7] , " & _
                                "[ORDINAL_PARENT],[ORDINAL]) VALUES "

Private Const SQL_GET_RECORD = "SELECT [CODE] , [BF] , [C7] , " & _
                                " [ORDINAL_PARENT] " & _
                                " FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.PLACE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.COUNTRY_CODE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL)
        strSql = strSql & ")"

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
   
        'ActiveConnection.Execute strSql
        ' MUCP-159 - End
        
        Dim intEventCtr As Integer
        
        For intEventCtr = 1 To ActiveRecord.Ctl_Controls.Count
            ActiveRecord.Ctl_Controls(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Ctl_Controls.AddRecord ActiveConnection, ActiveRecord.Ctl_Controls(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.Incidents.Count
            ActiveRecord.Incidents(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Incidents.AddRecord ActiveConnection, ActiveRecord.Incidents(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.NewSeals.Count
            ActiveRecord.NewSeals(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.NewSeals.AddRecord ActiveConnection, ActiveRecord.NewSeals(intEventCtr)
        Next intEventCtr
        
        For intEventCtr = 1 To ActiveRecord.Transhipments.Count
            ActiveRecord.Transhipments(intEventCtr).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.Transhipments.AddRecord ActiveConnection, ActiveRecord.Transhipments(intEventCtr)
        Next intEventCtr


        AddRecord = True
        Exit Function

    End If

    AddRecord = False

    Exit Function

ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        ActiveConnection.Execute strSql
        ' MUCP-159 - nd
        
        DeleteRecord = True
        Exit Function

    End If

    DeleteRecord = False

    Exit Function

ERROR_QUERY:
    DeleteRecord = False
End Function

' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strSql As String
    Dim lngRecordsAffected As Long
    Dim intEventCtr As Integer

    If ((ActiveRecord Is Nothing) = False) Then

'                        ",[" & FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & ""
        
        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_PLACE_FIELD & "]=" & SetEnclosedChr(ActiveRecord.PLACE) & _
                        ",[" & FIELD_COUNTRY_CODE & "]=" & SetEnclosedChr(ActiveRecord.COUNTRY_CODE) & ""

        strSql = strSql & " WHERE " & PK_FIELD & "=" & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & " AND [ORDINAL]= " & SetEnclosedChr(ActiveRecord.ORDINAL)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        lngRecordsAffected = ExecuteNonQuery(ActiveConnection, strSql)
        
        'ActiveConnection.Execute strSql, lngRecordsAffected
        ' MUCP-159 - End
        
        If (lngRecordsAffected > 0) Then
        
            For intEventCtr = 1 To ActiveRecord.Ctl_Controls.Count
                ActiveRecord.Ctl_Controls.ModifyRecord ActiveConnection, ActiveRecord.Ctl_Controls(intEventCtr)
            Next intEventCtr
            
            For intEventCtr = 1 To ActiveRecord.Incidents.Count
                ActiveRecord.Incidents.ModifyRecord ActiveConnection, ActiveRecord.Incidents(intEventCtr)
            Next intEventCtr
                
            For intEventCtr = 1 To ActiveRecord.NewSeals.Count
                ActiveRecord.NewSeals.ModifyRecord ActiveConnection, ActiveRecord.NewSeals(intEventCtr)
            Next intEventCtr
            
            For intEventCtr = 1 To ActiveRecord.Transhipments.Count
                ActiveRecord.Transhipments.ModifyRecord ActiveConnection, ActiveRecord.Transhipments(intEventCtr)
            Next intEventCtr

        ElseIf (lngRecordsAffected = 0) Then
        
            AddRecord ActiveConnection, ActiveRecord
            
            For intEventCtr = 1 To ActiveRecord.Ctl_Controls.Count
                ActiveRecord.Ctl_Controls.AddRecord ActiveConnection, ActiveRecord.Ctl_Controls(intEventCtr)
            Next intEventCtr
            
            For intEventCtr = 1 To ActiveRecord.Incidents.Count
                ActiveRecord.Incidents.AddRecord ActiveConnection, ActiveRecord.Incidents(intEventCtr)
            Next intEventCtr
                
            For intEventCtr = 1 To ActiveRecord.NewSeals.Count
                ActiveRecord.NewSeals.AddRecord ActiveConnection, ActiveRecord.NewSeals(intEventCtr)
            Next intEventCtr
            
            For intEventCtr = 1 To ActiveRecord.Transhipments.Count
                ActiveRecord.Transhipments.AddRecord ActiveConnection, ActiveRecord.Transhipments(intEventCtr)
            Next intEventCtr
            
        End If


        ModifyRecord = True
        Exit Function

    End If

    ModifyRecord = False

    Exit Function

ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strKey As String
    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & " AND [ORDINAL]=" & ActiveRecord.ORDINAL

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        
        'Set rstRecord = ActiveConnection.Execute(strSql)
        
        On Error GoTo ERROR_RECORDSET
        
        If Not (rstRecord.EOF And rstRecord.BOF) Then
        ' MUCP-159 - End
        
            ActiveRecord.PLACE = FNullField(rstRecord.Fields(FIELD_PLACE_FIELD))
            ActiveRecord.COUNTRY_CODE = FNullField(rstRecord.Fields(FIELD_COUNTRY_CODE))
            ActiveRecord.ORDINAL_PARENT = FNullField(rstRecord.Fields(FIELD_ORDINAL_PARENT))
            'ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))
            
            ' add incidents here 1x  1.2.1
            Set ActiveRecord.Incidents = New cpiIncidents
            strKey = ActiveRecord.Key & "_1.2.1-" & CStr(ActiveRecord.Incidents.Count + 1)
            ActiveRecord.Incidents.Add strKey, ActiveRecord.CODE_FIELD, ActiveRecord.ORDINAL
            ActiveRecord.Incidents.GetRecord ActiveConnection, ActiveRecord.Incidents(strKey)
            
            ' add control here here 1x 1.2.4
            Set ActiveRecord.Ctl_Controls = New cpiControls
            strKey = ActiveRecord.Key & "_1.2.4-" & CStr(ActiveRecord.Ctl_Controls.Count + 1)
            ActiveRecord.Ctl_Controls.Add strKey, ActiveRecord.CODE_FIELD, ActiveRecord.ORDINAL
            ActiveRecord.Ctl_Controls.GetRecord ActiveConnection, ActiveRecord.Ctl_Controls(strKey)
            
            ' 1.2.2
            ' add transhipment here 1x
            Set ActiveRecord.Transhipments = New cpiTranshipments
            strKey = ActiveRecord.Key & "_1.2.2-" & CStr(ActiveRecord.Transhipments.Count + 1)
            ActiveRecord.Transhipments.Add strKey, ActiveRecord.CODE_FIELD, ActiveRecord.ORDINAL
            ActiveRecord.Transhipments.GetRecord ActiveConnection, ActiveRecord.Transhipments(strKey)
            
            ' add seals here 1x
            Set ActiveRecord.NewSeals = New cpiSealInfos
            
            ' 1.2.3
            strKey = ActiveRecord.Key & "_1.2.3-" & CStr(ActiveRecord.NewSeals.Count + 1)
            ActiveRecord.NewSeals.Add strKey, ActiveRecord.CODE_FIELD, ActiveRecord.ORDINAL
            ActiveRecord.NewSeals(strKey).CODE_FIELD = ActiveRecord.CODE_FIELD
            ActiveRecord.NewSeals(strKey).ORDINAL = ActiveRecord.ORDINAL
            ActiveRecord.NewSeals.GetRecord ActiveConnection, ActiveRecord.NewSeals(strKey)
        
        End If
        'ActiveRecord.EnRouteEvents(9).
'        Private mvarHeaders As cpiEnRouteEvents ' 6
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiEnRouteEvent
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEnRouteEvent) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiEnRouteEvent
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiEnRouteEvent) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_PLACE_FIELD, 202, 35, 102 ' 2
    rstRecord.Fields.Append FIELD_COUNTRY_CODE, 202, 2, 102 ' 3
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 4
    rstRecord.Fields.Append FIELD_ORDINAL_PARENT, 3, 4, 118 ' 5
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_PLACE_FIELD) = ActiveRecord.PLACE
    rstRecord.Fields(FIELD_COUNTRY_CODE) = ActiveRecord.COUNTRY_CODE
    rstRecord.Fields(FIELD_ORDINAL_PARENT) = ActiveRecord.ORDINAL_PARENT
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiEnRouteEvent

    Dim clsRecord As cpiEnRouteEvent

    Set clsRecord = New cpiEnRouteEvent

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.PLACE = FNullField(ActiveRecord.Fields(FIELD_PLACE_FIELD))
    clsRecord.COUNTRY_CODE = FNullField(ActiveRecord.Fields(FIELD_COUNTRY_CODE))
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



Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiEnRouteEvent
    
    'create a new object
    Dim strKey As String
    Dim objNewMember As cpiEnRouteEvent
    Set objNewMember = New cpiEnRouteEvent

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    objNewMember.CODE_FIELD = PKValue
    objNewMember.ORDINAL = mCol.Count + 1
    
    ' 1.2.1
    Set objNewMember.Incidents = New cpiIncidents
    strKey = sKey & "_1.2.1-" & CStr(objNewMember.Incidents.Count + 1)
    objNewMember.Incidents.Add strKey, PKValue, objNewMember.ORDINAL
    
    ' 1.2.2
    Set objNewMember.Transhipments = New cpiTranshipments
    strKey = sKey & "_1.2.2-" & CStr(objNewMember.Transhipments.Count + 1)
    objNewMember.Transhipments.Add strKey, PKValue, objNewMember.ORDINAL
    'objNewMember.Transhipments(strKey).ORDINAL_PARENT = objNewMember.ORDINAL
    
    ' 1.2.3
    Set objNewMember.NewSeals = New cpiSealInfos
    strKey = sKey & "_1.2.3-" & CStr(objNewMember.NewSeals.Count + 1)
    objNewMember.NewSeals.Add strKey, PKValue, objNewMember.ORDINAL
    
    ' 1.2.4
    Set objNewMember.Ctl_Controls = New cpiControls
    strKey = sKey & "_1.2.4-" & CStr(objNewMember.Ctl_Controls.Count + 1)
    objNewMember.Ctl_Controls.Add strKey, PKValue, objNewMember.ORDINAL
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiEnRouteEvent
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
    'If (mCol.Count = 1) Then
    mCol.Remove vntIndexKey
    
    Dim lngCtr As Long
    Dim intColCtr As Integer
    Dim intChildCtr As Integer
    Dim intChildChildCtr As Integer
    
         'synchronize Index
    For intColCtr = 1 To mCol.Count
        mCol(intColCtr).Index = intColCtr
        mCol(intColCtr).ORDINAL = intColCtr
        mCol(intColCtr).Key = Mid(mCol(intColCtr).Key, 1, InStrRev(mCol(intColCtr).Key, "-")) & intColCtr
        
        ' update incidents parent
        For intChildCtr = 1 To mCol(intColCtr).Incidents.Count
            mCol(intColCtr).Incidents(intChildCtr).ORDINAL_PARENT = intColCtr
            mCol(intColCtr).Incidents(intChildCtr).Key = mCol(intColCtr).Key & "_1.2.1-" & mCol(intColCtr).Incidents.Count
        Next intChildCtr
        
        ' update Transhipments parent
        For intChildCtr = 1 To mCol(intColCtr).Transhipments.Count
            mCol(intColCtr).Transhipments(intChildCtr).ORDINAL_PARENT = intColCtr
            mCol(intColCtr).Transhipments(intChildCtr).Key = mCol(intColCtr).Key & "_1.2.2-" & mCol(intColCtr).Transhipments.Count
            
            For intChildChildCtr = 1 To mCol(intColCtr).Transhipments(intChildCtr).Containers.Count
                mCol(intColCtr).Transhipments(intChildCtr).Containers(intChildChildCtr).ORDINAL_PARENT = intColCtr
                mCol(intColCtr).Transhipments(intChildCtr).Containers(intChildChildCtr).Key = mCol(intColCtr).Transhipments(intChildCtr).Key & "_1.2.2.1-" & intChildChildCtr
            Next intChildChildCtr
        'objNewMember.Containers.AddNewGroup sKey & "_1.2.2.1", PKValue, ParentID
        Next intChildCtr
        
        ' update NewSeals parent
        For intChildCtr = 1 To mCol(intColCtr).NewSeals.Count
            mCol(intColCtr).NewSeals(intChildCtr).ORDINAL_PARENT = intColCtr
            mCol(intColCtr).NewSeals(intChildCtr).Key = mCol(intColCtr).Key & "_1.2.3-" & mCol(intColCtr).NewSeals.Count
            
            For intChildChildCtr = 1 To mCol(intColCtr).NewSeals(intChildCtr).Seals.Count
                mCol(intColCtr).NewSeals(intChildCtr).Seals(intChildChildCtr).ORDINAL_PARENT = intColCtr
                mCol(intColCtr).NewSeals(intChildCtr).Seals(intChildChildCtr).Key = mCol(intColCtr).NewSeals(intChildCtr).Key & "_1.2.3.1-" & intChildChildCtr
            Next intChildChildCtr
        'objNewMember.Containers.AddNewGroup sKey & "_1.2.2.1", PKValue, ParentID
        Next intChildCtr
        
        ' update CTlControls parent
        For intChildCtr = 1 To mCol(intColCtr).Ctl_Controls.Count
            mCol(intColCtr).Ctl_Controls(intChildCtr).ORDINAL_PARENT = intColCtr
            mCol(intColCtr).Ctl_Controls(intChildCtr).Key = mCol(intColCtr).Key & "_1.2.4-" & mCol(intColCtr).Ctl_Controls.Count
        Next intChildCtr
        ' update Containers parent
    Next intColCtr
    
    'Re-orders mCol so mCol(Index) corresponds with mCol(Index).Key
    For lngCtr = vntIndexKey To (mCol.Count)
        mCol.Add mCol(vntIndexKey), Mid(mCol(vntIndexKey).Key, 1, InStrRev(mCol(vntIndexKey).Key, "-")) & lngCtr
        mCol.Remove vntIndexKey
    Next
    
    'ElseIf ((mCol(vntIndexKey).ListIndex) <> mCol.Count) Then
        
    'End If

'    ' 1.2.1
'    Set objNewMember.Incidents = New cpiIncidents
'    strKey = sKey & "_1.2.1-" & CStr(objNewMember.Incidents.Count + 1)
'    objNewMember.Incidents.Add strKey, PKValue, objNewMember.ORDINAL
'
'    ' 1.2.2
'    Set objNewMember.Transhipments = New cpiTranshipments
'    strKey = sKey & "_1.2.2-" & CStr(objNewMember.Transhipments.Count + 1)
'    objNewMember.Transhipments.Add strKey, PKValue, objNewMember.ORDINAL
'    'objNewMember.Transhipments(strKey).ORDINAL_PARENT = objNewMember.ORDINAL
'
'    ' 1.2.3
'    Set objNewMember.NewSeals = New cpiSealInfos
'    strKey = sKey & "_1.2.3-" & CStr(objNewMember.NewSeals.Count + 1)
'    objNewMember.NewSeals.Add strKey, PKValue, objNewMember.ORDINAL
'
'    ' 1.2.4
'    Set objNewMember.Ctl_Controls = New cpiControls
'    strKey = sKey & "_1.2.4-" & CStr(objNewMember.Ctl_Controls.Count + 1)
'    objNewMember.Ctl_Controls.Add strKey, PKValue, objNewMember.ORDINAL

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

