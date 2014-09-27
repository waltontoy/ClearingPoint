VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiNCTS_IEM_TMS_Tables"
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
Private Const PK_FIELD = "NCTS_IEM_TMS_ID"

' field constants
Private Const FIELD_NCTS_IEM_TMS_ID = "NCTS_IEM_TMS_ID" '1
Private Const FIELD_NCTS_IEM_ID = "NCTS_IEM_ID" '2
Private Const FIELD_NCTS_IEM_TMS_RemarksIEMName = "NCTS_IEM_TMS_RemarksIEMName" '3
Private Const FIELD_EDI_TMS_ID = "EDI_TMS_ID" '4
Private Const FIELD_NCTS_IEM_TMS_RemarksSegmentTag = "NCTS_IEM_TMS_RemarksSegmentTag" '5
Private Const FIELD_NCTS_IEM_TMS_ParentID = "NCTS_IEM_TMS_ParentID" '6
Private Const FIELD_NCTS_IEM_TMS_RemarksTMSSequence = "NCTS_IEM_TMS_RemarksTMSSequence" '7
Private Const FIELD_NCTS_IEM_TMS_Ordinal = "NCTS_IEM_TMS_Ordinal" '8
Private Const FIELD_NCTS_IEM_TMS_Occurrence = "NCTS_IEM_TMS_Occurrence" '9
Private Const FIELD_NCTS_IEM_TMS_Usage = "NCTS_IEM_TMS_Usage" '10
Private Const FIELD_NCTS_IEM_TMS_RemarksQualifier = "NCTS_IEM_TMS_RemarksQualifier" '11

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[NCTS_IEM_TMS_ID],[NCTS_IEM_ID],[NCTS_IEM_TMS_RemarksIEMName]" & _
",[EDI_TMS_ID],[NCTS_IEM_TMS_RemarksSegmentTag],[NCTS_IEM_TMS_ParentID]" & _
",[NCTS_IEM_TMS_RemarksTMSSequence],[NCTS_IEM_TMS_Ordinal],[NCTS_IEM_TMS_Occurrence]" & _
",[NCTS_IEM_TMS_Usage],[NCTS_IEM_TMS_RemarksQualifier]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT[NCTS_IEM_TMS_ID],[NCTS_IEM_ID],[NCTS_IEM_TMS_RemarksIEMName]" & _
",[EDI_TMS_ID],[NCTS_IEM_TMS_RemarksSegmentTag],[NCTS_IEM_TMS_ParentID]" & _
",[NCTS_IEM_TMS_RemarksTMSSequence],[NCTS_IEM_TMS_Ordinal],[NCTS_IEM_TMS_Occurrence]" & _
",[NCTS_IEM_TMS_Usage],[NCTS_IEM_TMS_RemarksQualifier] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksIEMName) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksSegmentTag) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ParentID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksTMSSequence) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Ordinal) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Occurrence) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Usage) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksQualifier) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [NCTS_IEM_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_ID) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_RemarksIEMName]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksIEMName) & ", "
      strSql = strSql & " [EDI_TMS_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_EDI_TMS_ID) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_RemarksSegmentTag]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksSegmentTag) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_ParentID]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ParentID) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_RemarksTMSSequence]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksTMSSequence) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_Ordinal]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Ordinal) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_Occurrence]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Occurrence) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_Usage]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_Usage) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_RemarksQualifier]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksQualifier) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(rstRecord.Fields("NCTS_IEM_TMS_ID"))
      ActiveRecord.FIELD_NCTS_IEM_ID = FNullField(rstRecord.Fields("NCTS_IEM_ID"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksIEMName = FNullField(rstRecord.Fields("NCTS_IEM_TMS_RemarksIEMName"))
      ActiveRecord.FIELD_EDI_TMS_ID = FNullField(rstRecord.Fields("EDI_TMS_ID"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksSegmentTag = FNullField(rstRecord.Fields("NCTS_IEM_TMS_RemarksSegmentTag"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_ParentID = FNullField(rstRecord.Fields("NCTS_IEM_TMS_ParentID"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksTMSSequence = FNullField(rstRecord.Fields("NCTS_IEM_TMS_RemarksTMSSequence"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_Ordinal = FNullField(rstRecord.Fields("NCTS_IEM_TMS_Ordinal"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_Occurrence = FNullField(rstRecord.Fields("NCTS_IEM_TMS_Occurrence"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_Usage = FNullField(rstRecord.Fields("NCTS_IEM_TMS_Usage"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksQualifier = FNullField(rstRecord.Fields("NCTS_IEM_TMS_RemarksQualifier"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_IEM_TMS_Table

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_TMS_ID = rstRecord.Fields("ID_MAX").Value
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiNCTS_IEM_TMS_Table

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_NCTS_IEM_TMS_ID = rstRecord.Fields("ID_MIN").Value
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




Public Function GetTableRecord(ByRef ActiveRecord As cpiNCTS_IEM_TMS_Table) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_ID, 3, 4, 90 ' 1
   rstRecord.Fields.Append FIELD_NCTS_IEM_ID, 3, 4, 118 ' 2
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_RemarksIEMName, 202, 10, 102 ' 3
   rstRecord.Fields.Append FIELD_EDI_TMS_ID, 3, 4, 118 ' 4
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_RemarksSegmentTag, 202, 3, 102 ' 5
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_ParentID, 3, 4, 118 ' 6
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_RemarksTMSSequence, 3, 4, 118 ' 7
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_Ordinal, 3, 4, 118 ' 8
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_Occurrence, 3, 4, 118 ' 9
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_Usage, 202, 1, 102 ' 10
   rstRecord.Fields.Append FIELD_NCTS_IEM_TMS_RemarksQualifier, 202, 50, 102 ' 11

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("NCTS_IEM_TMS_ID") = ActiveRecord.FIELD_NCTS_IEM_TMS_ID
   rstRecord.Fields("NCTS_IEM_ID") = ActiveRecord.FIELD_NCTS_IEM_ID
   rstRecord.Fields("NCTS_IEM_TMS_RemarksIEMName") = ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksIEMName
   rstRecord.Fields("EDI_TMS_ID") = ActiveRecord.FIELD_EDI_TMS_ID
   rstRecord.Fields("NCTS_IEM_TMS_RemarksSegmentTag") = ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksSegmentTag
   rstRecord.Fields("NCTS_IEM_TMS_ParentID") = ActiveRecord.FIELD_NCTS_IEM_TMS_ParentID
   rstRecord.Fields("NCTS_IEM_TMS_RemarksTMSSequence") = ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksTMSSequence
   rstRecord.Fields("NCTS_IEM_TMS_Ordinal") = ActiveRecord.FIELD_NCTS_IEM_TMS_Ordinal
   rstRecord.Fields("NCTS_IEM_TMS_Occurrence") = ActiveRecord.FIELD_NCTS_IEM_TMS_Occurrence
   rstRecord.Fields("NCTS_IEM_TMS_Usage") = ActiveRecord.FIELD_NCTS_IEM_TMS_Usage
   rstRecord.Fields("NCTS_IEM_TMS_RemarksQualifier") = ActiveRecord.FIELD_NCTS_IEM_TMS_RemarksQualifier

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiNCTS_IEM_TMS_Table

   Dim clsRecord As cpiNCTS_IEM_TMS_Table
   Set clsRecord = New cpiNCTS_IEM_TMS_Table

   clsRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_ID))
   clsRecord.FIELD_NCTS_IEM_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_ID))
   clsRecord.FIELD_NCTS_IEM_TMS_RemarksIEMName = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_RemarksIEMName))
   clsRecord.FIELD_EDI_TMS_ID = FNullField(ActiveRecord.Fields(FIELD_EDI_TMS_ID))
   clsRecord.FIELD_NCTS_IEM_TMS_RemarksSegmentTag = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_RemarksSegmentTag))
   clsRecord.FIELD_NCTS_IEM_TMS_ParentID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_ParentID))
   clsRecord.FIELD_NCTS_IEM_TMS_RemarksTMSSequence = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_RemarksTMSSequence))
   clsRecord.FIELD_NCTS_IEM_TMS_Ordinal = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_Ordinal))
   clsRecord.FIELD_NCTS_IEM_TMS_Occurrence = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_Occurrence))
   clsRecord.FIELD_NCTS_IEM_TMS_Usage = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_Usage))
   clsRecord.FIELD_NCTS_IEM_TMS_RemarksQualifier = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_RemarksQualifier))
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


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiNCTS_IEM_TMS_Table

    'create a new object
    Dim objNewMember As cpiNCTS_IEM_TMS_Table
    Set objNewMember = New cpiNCTS_IEM_TMS_Table

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

Public Property Get Item(vntIndexKey As Variant) As cpiNCTS_IEM_TMS_Table
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




