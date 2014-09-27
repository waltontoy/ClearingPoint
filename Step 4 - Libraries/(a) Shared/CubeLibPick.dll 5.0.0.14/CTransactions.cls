VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cpiRecord"
Attribute VB_Ext_KEY = "Member0" ,"cpiRecord"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarActiveConnection As ADODB.Connection


Public Function Add(ADOFields As CADOFields _
                                    , RecordSource As ADODB.Recordset _
                                    , RecordSQL As String, Status As DataStatus _
                                    , Optional sKey As String _
                                    , Optional ByVal OldRecordSource As ADODB.Recordset) As CRecord
    'create a new object
    Dim objNewMember As CRecord
    Set objNewMember = New CRecord

    ' Set the properties passed into the method
    Set objNewMember.ADOFields = ADOFields
    Set objNewMember.RecordSource = RecordSource
    
    objNewMember.RecordSQL = RecordSQL
    objNewMember.Status = Status
    objNewMember.Key = sKey
    'objnewmember.TempDBID  =  ' REDESIGN
    Set objNewMember.OldRecordSource = OldRecordSource
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CRecord
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
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
      Set mvarActiveConnection = Nothing
End Sub

'ByRef DBConnection As ADODB.Connection
Public Property Set ActiveConnection(ByVal vData As ADODB.Connection)
    Set mvarActiveConnection = vData
End Property

Public Property Get ActiveConnection() As ADODB.Recordset
    Set ActiveConnection = mvarActiveConnection
End Property

Public Sub BeginTrans()
    ' conDBConnection.BeginTrans
    mvarActiveConnection.BeginTrans
End Sub

Public Sub CommitTrans(ByRef DBConnection As ADODB.Connection _
                             , ByVal SQLCommit As String _
                             , ByVal PKFieldBaseName As String)
                             
    Dim rstRecordsCommit As ADODB.Recordset
    Dim fldRecordsCommit As ADODB.Field
    
    Dim strPKFieldAlias As String
    Dim strSQLCommit As String
        
    Dim lngTransactionCnt As Long
    Dim strTableName As String
    Dim strTableCommand As String
    
    If mCol.Count > 0 Then
        strSQLCommit = RegenerateSQL(SQLCommit, True)
        
        strTableName = GetSQLFromTable(strSQLCommit)
        
        ADORecordsetOpen strSQLCommit, DBConnection, rstRecordsCommit, adOpenKeyset, adLockOptimistic
        
        For Each fldRecordsCommit In rstRecordsCommit.Fields
            If UCase(fldRecordsCommit.Properties(0).Value) = UCase(PKFieldBaseName) Then
                strPKFieldAlias = fldRecordsCommit.Name
                Exit For
            End If
        Next
                    
        For lngTransactionCnt = 1 To mCol.Count
            Select Case mCol(lngTransactionCnt).Status
                Case cpiStateDeleted
                    
                    rstRecordsCommit.MoveFirst
                    rstRecordsCommit.Find GetRecordCriteria(PKFieldBaseName, _
                                        mCol(lngTransactionCnt).RecordSource _
                                        .Fields(strPKFieldAlias).Value, _
                                        mCol(lngTransactionCnt).RecordSource _
                                        .Fields(strPKFieldAlias).Type), , adSearchForward
                    
                    If Not rstRecordsCommit.EOF Then
                        rstRecordsCommit.Delete
                        
                            strTableCommand = vbNullString
                            strTableCommand = strTableCommand & "DELETE "
                            strTableCommand = strTableCommand & "* "
                            strTableCommand = strTableCommand & "FROM "
                            strTableCommand = strTableCommand & "[" & strTableName & "] "
                            strTableCommand = strTableCommand & "WHERE "
                            strTableCommand = strTableCommand & GetRecordCriteria(PKFieldBaseName, _
                                                                mCol(lngTransactionCnt).RecordSource _
                                                                .Fields(strPKFieldAlias).Value, _
                                                                mCol(lngTransactionCnt).RecordSource _
                                                                .Fields(strPKFieldAlias).Type)
                        ExecuteNonQuery DBConnection, strTableCommand
                    End If
                    
                Case cpiStateModified
                    rstRecordsCommit.MoveFirst
                    rstRecordsCommit.Find GetRecordCriteria(PKFieldBaseName, _
                                        mCol(lngTransactionCnt).RecordSource _
                                        .Fields(strPKFieldAlias).Value, _
                                        mCol(lngTransactionCnt).RecordSource _
                                        .Fields(strPKFieldAlias).Type), , adSearchForward
                    
                    If Not rstRecordsCommit.EOF Then
                        On Error Resume Next 'ERROR POINT
                        For Each fldRecordsCommit In rstRecordsCommit.Fields
                            If Not mCol(lngTransactionCnt).RecordSource _
                                .Fields(fldRecordsCommit.Name) _
                                .Properties(5).Value Then
                                rstRecordsCommit.Fields(fldRecordsCommit.Name).Value = _
                                mCol(lngTransactionCnt).RecordSource _
                                .Fields(fldRecordsCommit.Name).Value
                            End If
                        Next
                        
                        Debug.Assert False
                        '#####
                        rstRecordsCommit.Update
                        
                        UpdateRecordset DBConnection, rstRecordsCommit, strTableName
                    End If
                    
                Case cpiStateNew
                    
                    Debug.Assert False
                    '#####
                    
                    rstRecordsCommit.AddNew
                    
                    For Each fldRecordsCommit In rstRecordsCommit.Fields
                        If Not mCol(lngTransactionCnt).RecordSource _
                            .Fields(fldRecordsCommit.Name).Properties(5).Value Then
                            rstRecordsCommit.Fields(fldRecordsCommit.Name).Value = _
                            mCol(lngTransactionCnt).RecordSource _
                            .Fields(fldRecordsCommit.Name).Value
                        End If   ' NULL, fieldtype=text, numeric return 0 or empty string
                    Next
                    rstRecordsCommit.Update
                    
                    InsertRecordset DBConnection, rstRecordsCommit, strTableName
            End Select
        Next lngTransactionCnt
        
        ADORecordsetClose rstRecordsCommit
    End If
End Sub
