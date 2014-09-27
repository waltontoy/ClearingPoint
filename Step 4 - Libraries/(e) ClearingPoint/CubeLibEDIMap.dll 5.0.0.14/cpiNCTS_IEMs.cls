VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiNCTS_IEMs"
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

Private Const TABLE_NAME = "EDI_TMS_Core"
Private Const PK_FIELD = "NCTS_IEM_ID"

' field constants
Private Const FIELD_NCTS_IEM_ID = "NCTS_IEM_ID" '1
Private Const FIELD_Done = "Done" '2
Private Const FIELD_EDI_TMS_COR_ID = "EDI_TMS_COR_ID" '3
Private Const FIELD_NCTS_IEM_Name = "NCTS_IEM_Name" '4
Private Const FIELD_NCTS_IEM_Code = "NCTS_IEM_Code" '5
Private Const FIELD_NCTS_IEM_Description_E = "NCTS_IEM_Description_E" '6
Private Const FIELD_NCTS_IEM_Description_D = "NCTS_IEM_Description_D" '7
Private Const FIELD_NCTS_IEM_Description_F = "NCTS_IEM_Description_F" '8
Private Const FIELD_NCTS_IEM_Reference = "NCTS_IEM_Reference" '9

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[NCTS_IEM_ID],[Done],[EDI_TMS_COR_ID]" & _
",[NCTS_IEM_Name],[NCTS_IEM_Code],[NCTS_IEM_Description_E]" & _
",[NCTS_IEM_Description_D],[NCTS_IEM_Description_F],[NCTS_IEM_Reference]" & _
"                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT[NCTS_IEM_ID],[Done],[EDI_TMS_COR_ID]" & _
",[NCTS_IEM_Name],[NCTS_IEM_Code],[NCTS_IEM_Description_E]" & _
",[NCTS_IEM_Description_D],[NCTS_IEM_Description_F],[NCTS_IEM_Reference]" & _
" FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_Done) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_COR_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Name) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Code) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_E) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_D) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_F) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Reference) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [Done]=" & SetEnclosedChr(ActiveRecord.FIELD_Done) & ", "
      strSql = strSql & " [EDI_TMS_COR_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_COR_ID) & ", "
      strSql = strSql & " [NCTS_IEM_Name]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Name) & ", "
      strSql = strSql & " [NCTS_IEM_Code]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Code) & ", "
      strSql = strSql & " [NCTS_IEM_Description_E]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_E) & ", "
      strSql = strSql & " [NCTS_IEM_Description_D]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_D) & ", "
      strSql = strSql & " [NCTS_IEM_Description_F]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Description_F) & ", "
      strSql = strSql & " [NCTS_IEM_Reference]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_Reference) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID)

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

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID)

      On Error GoTo ERROR_QUERY
      
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_ID = FNullField(rstRecord.Fields("NCTS_IEM_ID"))
      ActiveRecord.FIELD_Done = FNullField(rstRecord.Fields("Done"))
      ActiveRecord.FIELD_EDI_TMS_COR_ID = FNullField(rstRecord.Fields("EDI_TMS_COR_ID"))
      ActiveRecord.FIELD_NCTS_IEM_Name = FNullField(rstRecord.Fields("NCTS_IEM_Name"))
      ActiveRecord.FIELD_NCTS_IEM_Code = FNullField(rstRecord.Fields("NCTS_IEM_Code"))
      ActiveRecord.FIELD_NCTS_IEM_Description_E = FNullField(rstRecord.Fields("NCTS_IEM_Description_E"))
      ActiveRecord.FIELD_NCTS_IEM_Description_D = FNullField(rstRecord.Fields("NCTS_IEM_Description_D"))
      ActiveRecord.FIELD_NCTS_IEM_Description_F = FNullField(rstRecord.Fields("NCTS_IEM_Description_F"))
      ActiveRecord.FIELD_NCTS_IEM_Reference = FNullField(rstRecord.Fields("NCTS_IEM_Reference"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_IEM

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_ID = rstRecord.Fields("ID_MAX").Value
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_IEM

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_ID = rstRecord.Fields("ID_MIN").Value
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




Public Function GetTableRecord(ByRef ActiveRecord As cpiNCTS_IEM) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_NCTS_IEM_ID, 3, 4, 90 ' 1
   rstRecord.Fields.Append FIELD_Done, 11, 2, 86 ' 2
   rstRecord.Fields.Append FIELD_EDI_TMS_COR_ID, 3, 4, 118 ' 3
   rstRecord.Fields.Append FIELD_NCTS_IEM_Name, 202, 10, 102 ' 4
   rstRecord.Fields.Append FIELD_NCTS_IEM_Code, 202, 8, 102 ' 5
   rstRecord.Fields.Append FIELD_NCTS_IEM_Description_E, 202, 100, 102 ' 6
   rstRecord.Fields.Append FIELD_NCTS_IEM_Description_D, 202, 100, 102 ' 7
   rstRecord.Fields.Append FIELD_NCTS_IEM_Description_F, 202, 100, 102 ' 8
   rstRecord.Fields.Append FIELD_NCTS_IEM_Reference, 202, 20, 102 ' 9

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("NCTS_IEM_ID") = ActiveRecord.FIELD_NCTS_IEM_ID
   rstRecord.Fields("Done") = ActiveRecord.FIELD_Done
   rstRecord.Fields("EDI_TMS_COR_ID") = ActiveRecord.FIELD_EDI_TMS_COR_ID
   rstRecord.Fields("NCTS_IEM_Name") = ActiveRecord.FIELD_NCTS_IEM_Name
   rstRecord.Fields("NCTS_IEM_Code") = ActiveRecord.FIELD_NCTS_IEM_Code
   rstRecord.Fields("NCTS_IEM_Description_E") = ActiveRecord.FIELD_NCTS_IEM_Description_E
   rstRecord.Fields("NCTS_IEM_Description_D") = ActiveRecord.FIELD_NCTS_IEM_Description_D
   rstRecord.Fields("NCTS_IEM_Description_F") = ActiveRecord.FIELD_NCTS_IEM_Description_F
   rstRecord.Fields("NCTS_IEM_Reference") = ActiveRecord.FIELD_NCTS_IEM_Reference

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiNCTS_IEM

   Dim clsRecord As cpiNCTS_IEM
   Set clsRecord = New cpiNCTS_IEM

   clsRecord.FIELD_NCTS_IEM_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_ID))
   clsRecord.FIELD_Done = FNullField(ActiveRecord.Fields(FIELD_Done))
   clsRecord.FIELD_EDI_TMS_COR_ID = FNullField(ActiveRecord.Fields(FIELD_EDI_TMS_COR_ID))
   clsRecord.FIELD_NCTS_IEM_Name = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Name))
   clsRecord.FIELD_NCTS_IEM_Code = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Code))
   clsRecord.FIELD_NCTS_IEM_Description_E = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Description_E))
   clsRecord.FIELD_NCTS_IEM_Description_D = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Description_D))
   clsRecord.FIELD_NCTS_IEM_Description_F = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Description_F))
   clsRecord.FIELD_NCTS_IEM_Reference = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_Reference))
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


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiNCTS_IEM

    'create a new object
    Dim objNewMember As cpiNCTS_IEM
    Set objNewMember = New cpiNCTS_IEM

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

Public Property Get Item(vntIndexKey As Variant) As cpiNCTS_IEM
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



