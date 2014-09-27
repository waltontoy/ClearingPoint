VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSETUPs"
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

Private Const TABLE_NAME = "SETUP"
Private Const PK_FIELD = "ENGLISH"

' field constants
Private Const FIELD_ENGLISH = "ENGLISH" '1
Private Const FIELD_FRENCH = "FRENCH" '2
Private Const FIELD_DUTCH = "DUTCH" '3
Private Const FIELD_CURRENT = "CURRENT" '4
Private Const FIELD_LAST_USER = "LAST USER" '5
Private Const FIELD_USER_NO_SERIES = "USER NO SERIES" '6
Private Const FIELD_WITH_SECURITY = "WITH SECURITY" '7
Private Const FIELD_EDIT_TIME = "EDIT TIME" '8
Private Const FIELD_REPERTORIUM_FOR_EXPORT = "REPERTORIUM FOR EXPORT" '9
Private Const FIELD_SENT_TIME = "SENT TIME" '10
Private Const FIELD_TREE_TIME = "TREE TIME" '11
Private Const FIELD_LAST_COMPACT_DATE = "LAST COMPACT DATE" '12
Private Const FIELD_NUMBER_OF_USERS = "NUMBER OF USERS" '13
Private Const FIELD_FIRSTRUN = "FIRSTRUN" '14
Private Const FIELD_Memo_Field = "Memo Field" '15
Private Const FIELD_DataTransferredToNCTS = "DataTransferredToNCTS" '16

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[ENGLISH],[FRENCH],[DUTCH]" & _
",[CURRENT],[LAST USER],[USER NO SERIES]" & _
",[WITH SECURITY],[EDIT TIME],[REPERTORIUM FOR EXPORT]" & _
",[SENT TIME],[TREE TIME],[LAST COMPACT DATE]" & _
",[NUMBER OF USERS],[FIRSTRUN],[Memo Field]" & _
",[DataTransferredToNCTS]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [ENGLISH],[FRENCH],[DUTCH]" & _
",[CURRENT],[LAST USER],[USER NO SERIES]" & _
",[WITH SECURITY],[EDIT TIME],[REPERTORIUM FOR EXPORT]" & _
",[SENT TIME],[TREE TIME],[LAST COMPACT DATE]" & _
",[NUMBER OF USERS],[FIRSTRUN],[Memo Field]" & _
",[DataTransferredToNCTS] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FRENCH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DUTCH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CURRENT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LAST_USER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USER_NO_SERIES) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WITH_SECURITY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDIT_TIME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_REPERTORIUM_FOR_EXPORT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SENT_TIME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TREE_TIME) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LAST_COMPACT_DATE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NUMBER_OF_USERS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FIRSTRUN) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Memo_Field) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DataTransferredToNCTS) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [FRENCH]=" & SetEnclosedChr(ActiveRecord.FIELD_FRENCH) & ", "
      strSql = strSql & " [DUTCH]=" & SetEnclosedChr(ActiveRecord.FIELD_DUTCH) & ", "
      strSql = strSql & " [CURRENT]=" & SetEnclosedChr(ActiveRecord.FIELD_CURRENT) & ", "
      strSql = strSql & " [LAST USER]=" & SetEnclosedChr(ActiveRecord.FIELD_LAST_USER) & ", "
      strSql = strSql & " [USER NO SERIES]=" & SetEnclosedChr(ActiveRecord.FIELD_USER_NO_SERIES) & ", "
      strSql = strSql & " [WITH SECURITY]=" & SetEnclosedChr(ActiveRecord.FIELD_WITH_SECURITY) & ", "
      strSql = strSql & " [EDIT TIME]=" & SetEnclosedChr(ActiveRecord.FIELD_EDIT_TIME) & ", "
      strSql = strSql & " [REPERTORIUM FOR EXPORT]=" & SetEnclosedChr(ActiveRecord.FIELD_REPERTORIUM_FOR_EXPORT) & ", "
      strSql = strSql & " [SENT TIME]=" & SetEnclosedChr(ActiveRecord.FIELD_SENT_TIME) & ", "
      strSql = strSql & " [TREE TIME]=" & SetEnclosedChr(ActiveRecord.FIELD_TREE_TIME) & ", "
      strSql = strSql & " [LAST COMPACT DATE]=" & SetEnclosedChr(ActiveRecord.FIELD_LAST_COMPACT_DATE) & ", "
      strSql = strSql & " [NUMBER OF USERS]=" & SetEnclosedChr(ActiveRecord.FIELD_NUMBER_OF_USERS) & ", "
      strSql = strSql & " [FIRSTRUN]=" & SetEnclosedChr(ActiveRecord.FIELD_FIRSTRUN) & ", "
      strSql = strSql & " [Memo Field]=" & SetEnclosedChr(ActiveRecord.FIELD_Memo_Field) & ", "
      strSql = strSql & " [DataTransferredToNCTS]=" & SetEnclosedChr(ActiveRecord.FIELD_DataTransferredToNCTS) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_ENGLISH = FNullField(rstRecord.Fields("ENGLISH"))
      ActiveRecord.FIELD_FRENCH = FNullField(rstRecord.Fields("FRENCH"))
      ActiveRecord.FIELD_DUTCH = FNullField(rstRecord.Fields("DUTCH"))
      ActiveRecord.FIELD_CURRENT = FNullField(rstRecord.Fields("CURRENT"))
      ActiveRecord.FIELD_LAST_USER = FNullField(rstRecord.Fields("LAST USER"))
      ActiveRecord.FIELD_USER_NO_SERIES = FNullField(rstRecord.Fields("USER NO SERIES"))
      ActiveRecord.FIELD_WITH_SECURITY = FNullField(rstRecord.Fields("WITH SECURITY"))
      ActiveRecord.FIELD_EDIT_TIME = FNullField(rstRecord.Fields("EDIT TIME"))
      ActiveRecord.FIELD_REPERTORIUM_FOR_EXPORT = FNullField(rstRecord.Fields("REPERTORIUM FOR EXPORT"))
      ActiveRecord.FIELD_SENT_TIME = FNullField(rstRecord.Fields("SENT TIME"))
      ActiveRecord.FIELD_TREE_TIME = FNullField(rstRecord.Fields("TREE TIME"))
      ActiveRecord.FIELD_LAST_COMPACT_DATE = FNullField(rstRecord.Fields("LAST COMPACT DATE"))
      ActiveRecord.FIELD_NUMBER_OF_USERS = FNullField(rstRecord.Fields("NUMBER OF USERS"))
      ActiveRecord.FIELD_FIRSTRUN = FNullField(rstRecord.Fields("FIRSTRUN"))
      ActiveRecord.FIELD_Memo_Field = FNullField(rstRecord.Fields("Memo Field"))
      ActiveRecord.FIELD_DataTransferredToNCTS = FNullField(rstRecord.Fields("DataTransferredToNCTS"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSETUP

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_ENGLISH = rstRecord.Fields("ID_MAX").Value
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSETUP) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSETUP

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_ENGLISH = rstRecord.Fields("ID_MIN").Value
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




Public Function GetTableRecord(ByRef ActiveRecord As cpiSETUP) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_ENGLISH, 11, 2, 86 ' 1
   rstRecord.Fields.Append FIELD_FRENCH, 11, 2, 86 ' 2
   rstRecord.Fields.Append FIELD_DUTCH, 11, 2, 86 ' 3
   rstRecord.Fields.Append FIELD_CURRENT, 2, 2, 118 ' 4
   rstRecord.Fields.Append FIELD_LAST_USER, 202, 25, 102 ' 5
   rstRecord.Fields.Append FIELD_USER_NO_SERIES, 3, 4, 118 ' 6
   rstRecord.Fields.Append FIELD_WITH_SECURITY, 11, 2, 86 ' 7
   rstRecord.Fields.Append FIELD_EDIT_TIME, 7, 8, 118 ' 8
   rstRecord.Fields.Append FIELD_REPERTORIUM_FOR_EXPORT, 11, 2, 86 ' 9
   rstRecord.Fields.Append FIELD_SENT_TIME, 7, 8, 118 ' 10
   rstRecord.Fields.Append FIELD_TREE_TIME, 7, 8, 118 ' 11
   rstRecord.Fields.Append FIELD_LAST_COMPACT_DATE, 7, 8, 118 ' 12
   rstRecord.Fields.Append FIELD_NUMBER_OF_USERS, 2, 2, 118 ' 13
   rstRecord.Fields.Append FIELD_FIRSTRUN, 11, 2, 86 ' 14
   rstRecord.Fields.Append FIELD_Memo_Field, 11, 2, 86 ' 15
   rstRecord.Fields.Append FIELD_DataTransferredToNCTS, 11, 2, 86 ' 16

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("ENGLISH") = ActiveRecord.FIELD_ENGLISH
   rstRecord.Fields("FRENCH") = ActiveRecord.FIELD_FRENCH
   rstRecord.Fields("DUTCH") = ActiveRecord.FIELD_DUTCH
   rstRecord.Fields("CURRENT") = ActiveRecord.FIELD_CURRENT
   rstRecord.Fields("LAST USER") = ActiveRecord.FIELD_LAST_USER
   rstRecord.Fields("USER NO SERIES") = ActiveRecord.FIELD_USER_NO_SERIES
   rstRecord.Fields("WITH SECURITY") = ActiveRecord.FIELD_WITH_SECURITY
   rstRecord.Fields("EDIT TIME") = ActiveRecord.FIELD_EDIT_TIME
   rstRecord.Fields("REPERTORIUM FOR EXPORT") = ActiveRecord.FIELD_REPERTORIUM_FOR_EXPORT
   rstRecord.Fields("SENT TIME") = ActiveRecord.FIELD_SENT_TIME
   rstRecord.Fields("TREE TIME") = ActiveRecord.FIELD_TREE_TIME
   rstRecord.Fields("LAST COMPACT DATE") = ActiveRecord.FIELD_LAST_COMPACT_DATE
   rstRecord.Fields("NUMBER OF USERS") = ActiveRecord.FIELD_NUMBER_OF_USERS
   rstRecord.Fields("FIRSTRUN") = ActiveRecord.FIELD_FIRSTRUN
   rstRecord.Fields("Memo Field") = ActiveRecord.FIELD_Memo_Field
   rstRecord.Fields("DataTransferredToNCTS") = ActiveRecord.FIELD_DataTransferredToNCTS

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiSETUP

   Dim clsRecord As cpiSETUP
   Set clsRecord = New cpiSETUP

   clsRecord.FIELD_ENGLISH = FNullField(ActiveRecord.Fields(FIELD_ENGLISH))
   clsRecord.FIELD_FRENCH = FNullField(ActiveRecord.Fields(FIELD_FRENCH))
   clsRecord.FIELD_DUTCH = FNullField(ActiveRecord.Fields(FIELD_DUTCH))
   clsRecord.FIELD_CURRENT = FNullField(ActiveRecord.Fields(FIELD_CURRENT))
   clsRecord.FIELD_LAST_USER = FNullField(ActiveRecord.Fields(FIELD_LAST_USER))
   clsRecord.FIELD_USER_NO_SERIES = FNullField(ActiveRecord.Fields(FIELD_USER_NO_SERIES))
   clsRecord.FIELD_WITH_SECURITY = FNullField(ActiveRecord.Fields(FIELD_WITH_SECURITY))
   clsRecord.FIELD_EDIT_TIME = FNullField(ActiveRecord.Fields(FIELD_EDIT_TIME))
   clsRecord.FIELD_REPERTORIUM_FOR_EXPORT = FNullField(ActiveRecord.Fields(FIELD_REPERTORIUM_FOR_EXPORT))
   clsRecord.FIELD_SENT_TIME = FNullField(ActiveRecord.Fields(FIELD_SENT_TIME))
   clsRecord.FIELD_TREE_TIME = FNullField(ActiveRecord.Fields(FIELD_TREE_TIME))
   clsRecord.FIELD_LAST_COMPACT_DATE = FNullField(ActiveRecord.Fields(FIELD_LAST_COMPACT_DATE))
   clsRecord.FIELD_NUMBER_OF_USERS = FNullField(ActiveRecord.Fields(FIELD_NUMBER_OF_USERS))
   clsRecord.FIELD_FIRSTRUN = FNullField(ActiveRecord.Fields(FIELD_FIRSTRUN))
   clsRecord.FIELD_Memo_Field = FNullField(ActiveRecord.Fields(FIELD_Memo_Field))
   clsRecord.FIELD_DataTransferredToNCTS = FNullField(ActiveRecord.Fields(FIELD_DataTransferredToNCTS))
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

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiSETUP

    'create a new object
    Dim objNewMember As cpiSETUP
    Set objNewMember = New cpiSETUP

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

Public Property Get Item(vntIndexKey As Variant) As cpiSETUP
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




