VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEDIProps_Tbls"
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
Private mvarRecordset As ADODB.Recordset

Private Const TABLE_NAME = "EDIProperties"
Private Const PK_FIELD = "EDIPROP_ID"

' field constants
Private Const FIELD_EDIPROP_ID = "EDIPROP_ID" '1
Private Const FIELD_EDIPROP_Type = "EDIPROP_Type" '2
Private Const FIELD_EDIPROP_QueueName = "EDIPROP_QueueName" '3
Private Const FIELD_EDIPROP_Host = "EDIPROP_Host" '4
Private Const FIELD_EDIPROP_Port = "EDIPROP_Port" '5
Private Const FIELD_EDIPROP_UserName = "EDIPROP_UserName" '6
Private Const FIELD_EDIPROP_UserPassword = "EDIPROP_UserPassword" '7
Private Const FIELD_EDIPROP_SlowPollCycles = "EDIPROP_SlowPollCycles" '8
Private Const FIELD_EDIPROP_FastPollCycles = "EDIPROP_FastPollCycles" '9
Private Const FIELD_EDIPROP_TimeOut = "EDIPROP_TimeOut" '10

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[EDIPROP_ID],[EDIPROP_Type],[EDIPROP_QueueName]" & _
",[EDIPROP_Host],[EDIPROP_Port],[EDIPROP_UserName]" & _
",[EDIPROP_UserPassword],[EDIPROP_SlowPollCycles],[EDIPROP_FastPollCycles]" & _
",[EDIPROP_TimeOut]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [EDIPROP_ID],[EDIPROP_Type],[EDIPROP_QueueName]" & _
",[EDIPROP_Host],[EDIPROP_Port],[EDIPROP_UserName]" & _
",[EDIPROP_UserPassword],[EDIPROP_SlowPollCycles],[EDIPROP_FastPollCycles]" & _
",[EDIPROP_TimeOut] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Type) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_QueueName) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Host) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Port) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_UserName) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_UserPassword) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_SlowPollCycles) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_FastPollCycles) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_TimeOut) & ")"

      On Error GoTo ERROR_QUERY
        
      ExecuteNonQuery ActiveConnection, strSql
      'ActiveConnection.Execute strSql


      AddRecord = True

      Exit Function

  End If


  AddRecord = False


  Exit Function


  AddRecord = False

ERROR_QUERY:

  AddRecord = False

End Function


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_ID)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [EDIPROP_Type]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Type) & ", "
      strSql = strSql & " [EDIPROP_QueueName]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_QueueName) & ", "
      strSql = strSql & " [EDIPROP_Host]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Host) & ", "
      strSql = strSql & " [EDIPROP_Port]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_Port) & ", "
      strSql = strSql & " [EDIPROP_UserName]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_UserName) & ", "
      strSql = strSql & " [EDIPROP_UserPassword]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_UserPassword) & ", "
      strSql = strSql & " [EDIPROP_SlowPollCycles]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_SlowPollCycles) & ", "
      strSql = strSql & " [EDIPROP_FastPollCycles]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_FastPollCycles) & ", "
      strSql = strSql & " [EDIPROP_TimeOut]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_TimeOut) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_ID)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_EDIPROP_ID)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_EDIPROP_ID = FNullField(rstRecord.Fields("EDIPROP_ID"))
      ActiveRecord.FIELD_EDIPROP_Type = FNullField(rstRecord.Fields("EDIPROP_Type"))
      ActiveRecord.FIELD_EDIPROP_QueueName = FNullField(rstRecord.Fields("EDIPROP_QueueName"))
      ActiveRecord.FIELD_EDIPROP_Host = FNullField(rstRecord.Fields("EDIPROP_Host"))
      ActiveRecord.FIELD_EDIPROP_Port = FNullField(rstRecord.Fields("EDIPROP_Port"))
      ActiveRecord.FIELD_EDIPROP_UserName = FNullField(rstRecord.Fields("EDIPROP_UserName"))
      ActiveRecord.FIELD_EDIPROP_UserPassword = FNullField(rstRecord.Fields("EDIPROP_UserPassword"))
      ActiveRecord.FIELD_EDIPROP_SlowPollCycles = FNullField(rstRecord.Fields("EDIPROP_SlowPollCycles"))
      ActiveRecord.FIELD_EDIPROP_FastPollCycles = FNullField(rstRecord.Fields("EDIPROP_FastPollCycles"))
      ActiveRecord.FIELD_EDIPROP_TimeOut = FNullField(rstRecord.Fields("EDIPROP_TimeOut"))
      
      ADORecordsetClose rstRecord
      'Set rstRecord = Nothing

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


   Exit Function


ERROR_RECORDSET:
       ADORecordsetClose rstRecord
       'Set rstRecord = Nothing
       
       GetRecord = False
       Exit Function
ERROR_QUERY:

   GetRecord = False

End Function




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiEDIProps_Tbl

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_EDIPROP_ID = FNullField(rstRecord.Fields("ID_MAX"))
      
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiEDIProps_Tbl

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_EDIPROP_ID = FNullField(rstRecord.Fields("ID_MIN"))
      
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiEDIProps_Tbl) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   'rstRecord.Open strSql, ActiveConnection, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("EDIPROP_ID") = ActiveRecord.FIELD_EDIPROP_ID
   rstRecord.Fields("EDIPROP_Type") = ActiveRecord.FIELD_EDIPROP_Type
   rstRecord.Fields("EDIPROP_QueueName") = ActiveRecord.FIELD_EDIPROP_QueueName
   rstRecord.Fields("EDIPROP_Host") = ActiveRecord.FIELD_EDIPROP_Host
   rstRecord.Fields("EDIPROP_Port") = ActiveRecord.FIELD_EDIPROP_Port
   rstRecord.Fields("EDIPROP_UserName") = ActiveRecord.FIELD_EDIPROP_UserName
   rstRecord.Fields("EDIPROP_UserPassword") = ActiveRecord.FIELD_EDIPROP_UserPassword
   rstRecord.Fields("EDIPROP_SlowPollCycles") = ActiveRecord.FIELD_EDIPROP_SlowPollCycles
   rstRecord.Fields("EDIPROP_FastPollCycles") = ActiveRecord.FIELD_EDIPROP_FastPollCycles
   rstRecord.Fields("EDIPROP_TimeOut") = ActiveRecord.FIELD_EDIPROP_TimeOut

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiEDIProps_Tbl

   Dim clsRecord As cpiEDIProps_Tbl
   Set clsRecord = New cpiEDIProps_Tbl

   clsRecord.FIELD_EDIPROP_ID = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_ID))
   clsRecord.FIELD_EDIPROP_Type = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_Type))
   clsRecord.FIELD_EDIPROP_QueueName = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_QueueName))
   clsRecord.FIELD_EDIPROP_Host = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_Host))
   clsRecord.FIELD_EDIPROP_Port = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_Port))
   clsRecord.FIELD_EDIPROP_UserName = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_UserName))
   clsRecord.FIELD_EDIPROP_UserPassword = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_UserPassword))
   clsRecord.FIELD_EDIPROP_SlowPollCycles = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_SlowPollCycles))
   clsRecord.FIELD_EDIPROP_FastPollCycles = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_FastPollCycles))
   clsRecord.FIELD_EDIPROP_TimeOut = FNullField(ActiveRecord.Fields(FIELD_EDIPROP_TimeOut))
   Set GetClassRecord = clsRecord

   Set clsRecord = Nothing

End Function



Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, ByVal SearchField, ByVal SearchValue) As Boolean

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
      
      ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
      'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)

      SearchRecord = Not (rstDB.EOF And rstDB.BOF)
      
      'If (rstDB.EOF = False) Then
      '    SearchRecord = True
      'ElseIf (rstDB.EOF = True) Then
      '    SearchRecord = False
      'End If

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

   ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic
   'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiEDIProps_Tbl

    'create a new object
    Dim objNewMember As cpiEDIProps_Tbl
    Set objNewMember = New cpiEDIProps_Tbl

    'set the properties passed into the method
    objNewMember.Key = sKey
    objNewMember.Index = mCol.Count + 1
    'objNewMember.EDIPROP_ID = PKValue
    'objNewMember.ORDINAL = mCol.Count + 1

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cpiEDIProps_Tbl
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contEDIPROP_UserPasswordns either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Set mvarRecordset = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarRecordset = Nothing
End Sub

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mvarRecordset
End Property

Public Property Set Recordset(ByVal vNewValue As ADODB.Recordset)
    Set mvarRecordset = vNewValue
End Property




