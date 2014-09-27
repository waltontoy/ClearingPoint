VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiHeaders"
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

Private Const TABLE_NAME = "[DATA_NCTS_BERICHT_HOOFDING]"
Private Const PK_FIELD = "[CODE]"

' field constants
Private Const FIELD_CODE_FIELD = "CODE" ' 1
Private Const FIELD_MOVEMENT_REFERENCE_NUMBER = "MR" ' 2
Private Const FIELD_ARRIVAL_NOTIFICATION_DATE = "MC" ' 3
Private Const FIELD_ARRIVAL_NOTIFICATION_PLACE = "BC" ' 4
Private Const FIELD_ARRIVAL_AGREED_LOCATION_CODE = "AH" ' 5
Private Const FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS = "AG" ' 6
Private Const FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS = "AI" ' 7
Private Const FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = "AJ" ' 8
Private Const FIELD_SIMPLIFIED_PROCEDURE_FLAG = "EP" ' 9
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" ' 10
Private Const FIELD_ORDINAL = "ORDINAL" ' 11
Private Const FIELD_TIR_DISCHARGE = "ZA" ' 12
Private Const FIELD_TIR_VOLET_PAGE_NUMBER = "ZB" ' 12

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " ([CODE] , [MR] , [MC] , " & _
                                "BC, AH,AG,AI,[AJ],[EP],[ORDINAL_PARENT],[ORDINAL], [ZA], [ZB] ) VALUES "

Private Const SQL_GET_RECORD = "SELECT  [CODE] , [MR] , [MC] , " & _
                                "BC, AH,AG,AI,[AJ],[EP],[ORDINAL_PARENT],[ORDINAL], [ZA], [ZB] " & _
                                " FROM " & TABLE_NAME & " WHERE [CODE]= "


Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [CODE]= "

Private Const SQL_GET_MAXID = "SELECT MAX([CODE]) AS [ID_MAX] FROM " & TABLE_NAME
Private Const SQL_GET_MINID = "SELECT Min([CODE]) AS [ID_MIN] FROM " & TABLE_NAME

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & "("
        strSql = strSql & SetEnclosedChr(ActiveRecord.CODE_FIELD)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.MOVEMENT_REFERENCE_NUMBER)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ARRIVAL_NOTIFICATION_DATE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ARRIVAL_NOTIFICATION_PLACE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ARRIVAL_AGREED_LOCATION_CODE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ARRIVAL_AGREED_LOCATION_OF_GOODS)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ARRIVAL_AUTHORISED_LOCATION_OF_GOODS)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.SIMPLIFIED_PROCEDURE_FLAG)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.ORDINAL)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.TIR_DISCHARGE)
        strSql = strSql & "," & SetEnclosedChr(ActiveRecord.TIR_VOLET_PAGE_NUMBER)
        strSql = strSql & ")"



        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
   
        ActiveConnection.Execute strSql
        ' MUCP-159 - End
        
        AddRecord = True
        Exit Function

    End If

    AddRecord = False

    Exit Function

ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.CODE_FIELD)

        On Error GoTo ERROR_QUERY
        
        ' MUCP-159 - Start
        ExecuteNonQuery ActiveConnection, strSql
        
        ActiveConnection.Execute strSql
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
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET [" & FIELD_MOVEMENT_REFERENCE_NUMBER & "]=" & SetEnclosedChr(ActiveRecord.MOVEMENT_REFERENCE_NUMBER) & _
                        ",[" & FIELD_ARRIVAL_NOTIFICATION_DATE & "]=" & SetEnclosedChr(ActiveRecord.ARRIVAL_NOTIFICATION_DATE) & "" & _
                        ",[" & FIELD_ARRIVAL_NOTIFICATION_PLACE & "]=" & SetEnclosedChr(ActiveRecord.ARRIVAL_NOTIFICATION_PLACE) & "" & _
                        ",[" & FIELD_ARRIVAL_AGREED_LOCATION_CODE & "]=" & SetEnclosedChr(ActiveRecord.ARRIVAL_AGREED_LOCATION_CODE) & "" & _
                        ",[" & FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS & "]=" & SetEnclosedChr(ActiveRecord.ARRIVAL_AGREED_LOCATION_OF_GOODS) & "" & _
                        ",[" & FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS & "]=" & SetEnclosedChr(ActiveRecord.ARRIVAL_AUTHORISED_LOCATION_OF_GOODS) & "" & _
                        ",[" & FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION & "]=" & SetEnclosedChr(ActiveRecord.DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION) & "" & _
                        ",[" & FIELD_SIMPLIFIED_PROCEDURE_FLAG & "]=" & SetEnclosedChr(ActiveRecord.SIMPLIFIED_PROCEDURE_FLAG) & "" & _
                        ",[" & FIELD_ORDINAL_PARENT & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL_PARENT) & "" & _
                        ",[" & FIELD_ORDINAL & "]=" & SetEnclosedChr(ActiveRecord.ORDINAL) & "" & _
                        ",[" & FIELD_TIR_DISCHARGE & "]=" & SetEnclosedChr(ActiveRecord.TIR_DISCHARGE) & "" & _
                        ",[" & FIELD_TIR_VOLET_PAGE_NUMBER & "]=" & SetEnclosedChr(ActiveRecord.TIR_VOLET_PAGE_NUMBER) & ""
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

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

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
        ActiveRecord.MOVEMENT_REFERENCE_NUMBER = FNullField(rstRecord.Fields(FIELD_MOVEMENT_REFERENCE_NUMBER))
        ActiveRecord.ARRIVAL_NOTIFICATION_DATE = FNullField(rstRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_DATE))
        ActiveRecord.ARRIVAL_NOTIFICATION_PLACE = FNullField(rstRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_PLACE))
        ActiveRecord.ARRIVAL_AGREED_LOCATION_CODE = FNullField(rstRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_CODE))
        ActiveRecord.ARRIVAL_AGREED_LOCATION_OF_GOODS = FNullField(rstRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS))
        ActiveRecord.ARRIVAL_AUTHORISED_LOCATION_OF_GOODS = FNullField(rstRecord.Fields(FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS))
        ActiveRecord.DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = FNullField(rstRecord.Fields(FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION))
        ActiveRecord.SIMPLIFIED_PROCEDURE_FLAG = FNullField(rstRecord.Fields(FIELD_SIMPLIFIED_PROCEDURE_FLAG))
        ActiveRecord.ORDINAL_PARENT = FNullField(rstRecord.Fields(FIELD_ORDINAL_PARENT))
        ActiveRecord.ORDINAL = FNullField(rstRecord.Fields(FIELD_ORDINAL))
        ActiveRecord.TIR_DISCHARGE = FNullField(rstRecord.Fields(FIELD_TIR_DISCHARGE))
        ActiveRecord.TIR_VOLET_PAGE_NUMBER = FNullField(rstRecord.Fields(FIELD_TIR_VOLET_PAGE_NUMBER))
        
        
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

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiHeader
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiHeader) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiHeader
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
Public Function GetTableRecord(ByRef ActiveRecord As cpiHeader) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset

    Set rstRecord = New ADODB.Recordset

    rstRecord.Open

    rstRecord.Fields.Append FIELD_CODE_FIELD, 202, 21, 102 ' 1
    rstRecord.Fields.Append FIELD_MOVEMENT_REFERENCE_NUMBER, 202, 21, 102 ' 2
    rstRecord.Fields.Append FIELD_ARRIVAL_NOTIFICATION_DATE, 202, 8, 102 ' 3
    rstRecord.Fields.Append FIELD_ARRIVAL_NOTIFICATION_PLACE, 202, 35, 102 ' 4
    rstRecord.Fields.Append FIELD_ARRIVAL_AGREED_LOCATION_CODE, 202, 17, 102 ' 5
    rstRecord.Fields.Append FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS, 202, 35, 102 ' 6
    rstRecord.Fields.Append FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS, 202, 17, 102 ' 7
    rstRecord.Fields.Append FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION, 202, 2, 102 ' 8
    rstRecord.Fields.Append FIELD_SIMPLIFIED_PROCEDURE_FLAG, 1, 102 ' 9
    rstRecord.Fields.Append FIELD_ORDINAL_PARENT, 3, 4, 118 ' 10
    rstRecord.Fields.Append FIELD_ORDINAL, 3, 4, 118 ' 11
    rstRecord.Fields.Append FIELD_TIR_DISCHARGE, 202, 2, 102 ' 11
    rstRecord.Fields.Append FIELD_TIR_VOLET_PAGE_NUMBER, 202, 2, 102 ' 11
'
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_CODE_FIELD) = ActiveRecord.CODE_FIELD
    rstRecord.Fields(FIELD_MOVEMENT_REFERENCE_NUMBER) = ActiveRecord.MOVEMENT_REFERENCE_NUMBER
    rstRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_DATE) = ActiveRecord.ARRIVAL_NOTIFICATION_DATE
    rstRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_PLACE) = ActiveRecord.ARRIVAL_NOTIFICATION_PLACE
    rstRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_CODE) = ActiveRecord.ARRIVAL_AGREED_LOCATION_CODE
    rstRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS) = ActiveRecord.ARRIVAL_AGREED_LOCATION_OF_GOODS
    rstRecord.Fields(FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS) = ActiveRecord.ARRIVAL_AUTHORISED_LOCATION_OF_GOODS
    rstRecord.Fields(FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION) = ActiveRecord.DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION
    rstRecord.Fields(FIELD_SIMPLIFIED_PROCEDURE_FLAG) = ActiveRecord.SIMPLIFIED_PROCEDURE_FLAG
    rstRecord.Fields(FIELD_ORDINAL_PARENT) = ActiveRecord.ORDINAL_PARENT
    rstRecord.Fields(FIELD_ORDINAL) = ActiveRecord.ORDINAL
    rstRecord.Fields(FIELD_TIR_DISCHARGE) = ActiveRecord.TIR_DISCHARGE
    rstRecord.Fields(FIELD_TIR_VOLET_PAGE_NUMBER) = ActiveRecord.TIR_VOLET_PAGE_NUMBER
    '
    rstRecord.Update

    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiHeader

    Dim clsRecord As cpiHeader

    Set clsRecord = New cpiHeader

    clsRecord.CODE_FIELD = FNullField(ActiveRecord.Fields(FIELD_CODE_FIELD))
    clsRecord.MOVEMENT_REFERENCE_NUMBER = FNullField(ActiveRecord.Fields(FIELD_MOVEMENT_REFERENCE_NUMBER))
    clsRecord.ARRIVAL_NOTIFICATION_DATE = FNullField(ActiveRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_DATE))
    clsRecord.ARRIVAL_NOTIFICATION_PLACE = FNullField(ActiveRecord.Fields(FIELD_ARRIVAL_NOTIFICATION_PLACE))
    clsRecord.ARRIVAL_AGREED_LOCATION_CODE = FNullField(ActiveRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_CODE))
    clsRecord.ARRIVAL_AGREED_LOCATION_OF_GOODS = FNullField(ActiveRecord.Fields(FIELD_ARRIVAL_AGREED_LOCATION_OF_GOODS))
    clsRecord.ARRIVAL_AUTHORISED_LOCATION_OF_GOODS = FNullField(ActiveRecord.Fields(FIELD_ARRIVAL_AUTHORISED_LOCATION_OF_GOODS))
    clsRecord.DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = FNullField(ActiveRecord.Fields(FIELD_DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION))
    clsRecord.SIMPLIFIED_PROCEDURE_FLAG = FNullField(ActiveRecord.Fields(FIELD_SIMPLIFIED_PROCEDURE_FLAG))
    clsRecord.ORDINAL_PARENT = FNullField(ActiveRecord.Fields(FIELD_ORDINAL_PARENT))
    clsRecord.ORDINAL = FNullField(ActiveRecord.Fields(FIELD_ORDINAL))
    clsRecord.TIR_DISCHARGE = FNullField(ActiveRecord.Fields(FIELD_TIR_DISCHARGE))
    clsRecord.TIR_VOLET_PAGE_NUMBER = FNullField(ActiveRecord.Fields(FIELD_TIR_VOLET_PAGE_NUMBER))
    
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


' >>>
Public Function Add(ByRef sKey As String, ByVal PKValue As String) As PCubeLibEDIArrivals.cpiHeader
    
    'create a new object
    Dim objNewMember As cpiHeader
    Set objNewMember = New cpiHeader

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

Public Property Get Item(vntIndexKey As Variant) As cpiHeader
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
