VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiPICK_MAINTs"
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

'Private Const mvarTableName = "PICKLIST MAINTENANCE"
Private mvarTableName As String  ' = "PICKLIST MAINTENANCE"
Private mvarActiveLanguage As String  ' = "PICKLIST MAINTENANCE"

Private Const PK_FIELD = "INTERNAL CODE"

' field constants
Private Const FIELD_INTERNAL_CODE = "INTERNAL CODE" '1
Private Const FIELD_CODE = "CODE" '2

'Private Const strFIELD_DESCRIPTION = "DESCRIPTION" '3
Dim strFIELD_DESCRIPTION As String

Dim strSQL_ADD_RECORD As String
'Private Const strSQL_ADD_RECORD = "INSERT INTO [" & mvarTableName & "] (" & _
"[INTERNAL CODE],[CODE],[DESCRIPTION]                                     )  VALUES "

Dim strSQL_GET_RECORD As String
'Private Const strSQL_GET_RECORD = "SELECT [INTERNAL CODE],[CODE],[DESCRIPTION] FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

Dim strSQL_GET_ALL_RECORD As String
'Private Const strSQL_GET_ALL_RECORD = "SELECT [INTERNAL CODE],[CODE],[DESCRIPTION] FROM [" & mvarTableName & "]"

Dim strSQL_DELETE_RECORD As String
'Private Const strSQL_DELETE_RECORD = "DELETE * FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

Dim strSQL_GET_MAXID As String
'Private Const strSQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & mvarTableName & "]"

Dim strSQL_GET_MINID As String
'Private Const strSQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & mvarTableName & "]"


' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DESCRIPTION) & ")"

      On Error GoTo ERROR_QUERY

      ExecuteNonQuery ActiveConnection, strSql


      AddRecord = True

      Exit Function

  End If


  AddRecord = False


  Exit Function


  AddRecord = False

ERROR_QUERY:

  AddRecord = False

End Function


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & mvarTableName & "] SET "
      strSql = strSql & " [CODE]=" & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ", "
      strSql = strSql & " [DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_DESCRIPTION) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_INTERNAL_CODE = FNullField(rstRecord.Fields("INTERNAL CODE"))
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("CODE"))
      ActiveRecord.FIELD_DESCRIPTION = FNullField(rstRecord.Fields("DESCRIPTION"))
      ADORecordsetClose rstRecord

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


   Exit Function


ERROR_RECORDSET:
       ADORecordsetClose rstRecord
       GetRecord = False
       Exit Function
ERROR_QUERY:

   GetRecord = False

End Function




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiPICK_MAINT

   End If

      strSql = strSQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_INTERNAL_CODE = FNullField(rstRecord.Fields("ID_MAX"))
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiPICK_MAINT

   End If

      strSql = strSQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_INTERNAL_CODE = FNullField(rstRecord.Fields("ID_MIN"))
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_MAINT) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & mvarTableName & "]"
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("INTERNAL CODE") = ActiveRecord.FIELD_INTERNAL_CODE
   rstRecord.Fields("CODE") = ActiveRecord.FIELD_CODE
   rstRecord.Fields("DESCRIPTION") = ActiveRecord.FIELD_DESCRIPTION

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiPICK_MAINT

   Dim clsRecord As cpiPICK_MAINT
   Set clsRecord = New cpiPICK_MAINT

   clsRecord.FIELD_INTERNAL_CODE = FNullField(ActiveRecord.Fields(FIELD_INTERNAL_CODE))
   clsRecord.FIELD_CODE = FNullField(ActiveRecord.Fields(FIELD_CODE))
   clsRecord.FIELD_DESCRIPTION = FNullField(ActiveRecord.Fields(strFIELD_DESCRIPTION))
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
      strSql = "SELECT TOP 1 " & SearchField & " FROM [" & mvarTableName & "] WHERE " & SearchField & "=" & SetEnclosedChr((SearchValue))
      ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic

      SearchRecord = Not (rstDB.EOF And rstDB.BOF)
      
'      If (rstDB.EOF = False) Then
'          SearchRecord = True
'      ElseIf (rstDB.EOF = True) Then
'          SearchRecord = False
'      End If

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
      CommandText = strSQL_GET_ALL_RECORD
   End If

   rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiPICK_MAINT

    'create a new object
    Dim objNewMember As cpiPICK_MAINT
    Set objNewMember = New cpiPICK_MAINT

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

Public Property Get Item(vntIndexKey As Variant) As cpiPICK_MAINT
Attribute Item.VB_UserMemId = 0
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

Public Function SetSqlParameters(ByVal TableName As String, ByVal ActiveLanguage As String) As Boolean
'
    mvarTableName = TableName
    mvarActiveLanguage = UCase$(ActiveLanguage)
    
    strSQL_ADD_RECORD = "INSERT INTO [" & mvarTableName & "] (" & _
            "[INTERNAL CODE],[CODE],[DESCRIPTION]  )  VALUES "

    strSQL_GET_RECORD = "SELECT [INTERNAL CODE],[CODE],[DESCRIPTION] FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "
    
    strSQL_GET_ALL_RECORD = "SELECT [INTERNAL CODE],[CODE],[DESCRIPTION] FROM [" & mvarTableName & "]"
    
    strSQL_DELETE_RECORD = "DELETE * FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

    strSQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & mvarTableName & "]"

    strSQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & mvarTableName & "]"

    strFIELD_DESCRIPTION = "DESCRIPTION " & mvarActiveLanguage
'
End Function



