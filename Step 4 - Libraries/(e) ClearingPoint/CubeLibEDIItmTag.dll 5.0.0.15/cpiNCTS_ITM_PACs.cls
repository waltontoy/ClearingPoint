VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiNCTS_ITM_PACs"
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
Private Const PK_FIELD = "NCTS_ITM_PAC_ID"

' field constants
Private Const FIELD_NCTS_ITM_PAC_ID = "NCTS_ITM_PAC_ID" '1
Private Const FIELD_NCTS_IEM_TMS_ID = "NCTS_IEM_TMS_ID" '2
Private Const FIELD_EDI_TMS_ITM_ID = "EDI_TMS_ITM_ID" '3
Private Const FIELD_NCTS_ITM_PAC_Description = "NCTS_ITM_PAC_Description" '4
Private Const FIELD_NCTS_ITM_PAC_Value = "NCTS_ITM_PAC_Value" '5
Private Const FIELD_NCTS_ITM_PAC_Codelist = "NCTS_ITM_PAC_Codelist" '6
Private Const FIELD_NCTS_ITM_PAC_DataType = "NCTS_ITM_PAC_DataType" '7
Private Const FIELD_NCTS_ITM_PAC_Usage = "NCTS_ITM_PAC_Usage" '8

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[NCTS_ITM_PAC_ID],[NCTS_IEM_TMS_ID],[EDI_TMS_ITM_ID]" & _
",[NCTS_ITM_PAC_Description],[NCTS_ITM_PAC_Value],[NCTS_ITM_PAC_Codelist]" & _
",[NCTS_ITM_PAC_DataType],[NCTS_ITM_PAC_Usage]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT[NCTS_ITM_PAC_ID],[NCTS_IEM_TMS_ID],[EDI_TMS_ITM_ID]" & _
",[NCTS_ITM_PAC_Description],[NCTS_ITM_PAC_Value],[NCTS_ITM_PAC_Codelist]" & _
",[NCTS_ITM_PAC_DataType],[NCTS_ITM_PAC_Usage] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_ITM_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Description) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Value) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Codelist) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_DataType) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Usage) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_ID)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [NCTS_IEM_TMS_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID) & ", "
      strSql = strSql & " [EDI_TMS_ITM_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_ITM_ID) & ", "
      strSql = strSql & " [NCTS_ITM_PAC_Description]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Description) & ", "
      strSql = strSql & " [NCTS_ITM_PAC_Value]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Value) & ", "
      strSql = strSql & " [NCTS_ITM_PAC_Codelist]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Codelist) & ", "
      strSql = strSql & " [NCTS_ITM_PAC_DataType]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_DataType) & ", "
      strSql = strSql & " [NCTS_ITM_PAC_Usage]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_Usage) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_ID)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_ITM_PAC_ID)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_ITM_PAC_ID = FNullField(rstRecord.Fields("NCTS_ITM_PAC_ID"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(rstRecord.Fields("NCTS_IEM_TMS_ID"))
      ActiveRecord.FIELD_EDI_TMS_ITM_ID = FNullField(rstRecord.Fields("EDI_TMS_ITM_ID"))
      ActiveRecord.FIELD_NCTS_ITM_PAC_Description = FNullField(rstRecord.Fields("NCTS_ITM_PAC_Description"))
      ActiveRecord.FIELD_NCTS_ITM_PAC_Value = FNullField(rstRecord.Fields("NCTS_ITM_PAC_Value"))
      ActiveRecord.FIELD_NCTS_ITM_PAC_Codelist = FNullField(rstRecord.Fields("NCTS_ITM_PAC_Codelist"))
      ActiveRecord.FIELD_NCTS_ITM_PAC_DataType = FNullField(rstRecord.Fields("NCTS_ITM_PAC_DataType"))
      ActiveRecord.FIELD_NCTS_ITM_PAC_Usage = FNullField(rstRecord.Fields("NCTS_ITM_PAC_Usage"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_ITM_PAC

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_ITM_PAC_ID = rstRecord.Fields("ID_MAX").Value
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_ITM_PAC) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_ITM_PAC

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_ITM_PAC_ID = rstRecord.Fields("ID_MIN").Value
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




Public Function GetTableRecord(ByRef ActiveRecord As cpiNCTS_ITM_PAC) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_ID, 3, 4, 90 ' 1
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_ID, 3, 4, 118 ' 2
   rstRecord.Fields.Append FIELD_EDI_TMS_ITM_ID, 3, 4, 118 ' 3
   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_Description, 202, 50, 102 ' 4
   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_Value, 202, 50, 102 ' 5
   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_Codelist, 202, 5, 102 ' 6
   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_DataType, 202, 10, 102 ' 7
   rstRecord.Fields.Append FIELD_NCTS_ITM_PAC_Usage, 202, 1, 102 ' 8

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("NCTS_ITM_PAC_ID") = ActiveRecord.FIELD_NCTS_ITM_PAC_ID
   rstRecord.Fields("NCTS_IEM_TMS_ID") = ActiveRecord.FIELD_NCTS_IEM_TMS_ID
   rstRecord.Fields("EDI_TMS_ITM_ID") = ActiveRecord.FIELD_EDI_TMS_ITM_ID
   rstRecord.Fields("NCTS_ITM_PAC_Description") = ActiveRecord.FIELD_NCTS_ITM_PAC_Description
   rstRecord.Fields("NCTS_ITM_PAC_Value") = ActiveRecord.FIELD_NCTS_ITM_PAC_Value
   rstRecord.Fields("NCTS_ITM_PAC_Codelist") = ActiveRecord.FIELD_NCTS_ITM_PAC_Codelist
   rstRecord.Fields("NCTS_ITM_PAC_DataType") = ActiveRecord.FIELD_NCTS_ITM_PAC_DataType
   rstRecord.Fields("NCTS_ITM_PAC_Usage") = ActiveRecord.FIELD_NCTS_ITM_PAC_Usage

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiNCTS_ITM_PAC

   Dim clsRecord As cpiNCTS_ITM_PAC
   Set clsRecord = New cpiNCTS_ITM_PAC

   clsRecord.FIELD_NCTS_ITM_PAC_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_ID))
   clsRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_ID))
   clsRecord.FIELD_EDI_TMS_ITM_ID = FNullField(ActiveRecord.Fields(FIELD_EDI_TMS_ITM_ID))
   clsRecord.FIELD_NCTS_ITM_PAC_Description = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_Description))
   clsRecord.FIELD_NCTS_ITM_PAC_Value = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_Value))
   clsRecord.FIELD_NCTS_ITM_PAC_Codelist = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_Codelist))
   clsRecord.FIELD_NCTS_ITM_PAC_DataType = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_DataType))
   clsRecord.FIELD_NCTS_ITM_PAC_Usage = FNullField(ActiveRecord.Fields(FIELD_NCTS_ITM_PAC_Usage))
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


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiNCTS_ITM_PAC

    'create a new object
    Dim objNewMember As cpiNCTS_ITM_PAC
    Set objNewMember = New cpiNCTS_ITM_PAC

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

Public Property Get Item(vntIndexKey As Variant) As cpiNCTS_ITM_PAC
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




