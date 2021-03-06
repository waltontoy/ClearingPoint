VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDATA_NCTS_BGMs"
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

Private Const TABLE_NAME = "DATA_NCTS_BGM"
Private Const PK_FIELD = "DATA_NCTS_BGM_ID"

' field constants
Private Const FIELD_DATA_NCTS_BGM_ID = "DATA_NCTS_BGM_ID" '1
Private Const FIELD_DATA_NCTS_MSG_ID = "DATA_NCTS_MSG_ID" '2
Private Const FIELD_DATA_NCTS_BGM_ParentID = "DATA_NCTS_BGM_ParentID" '3
Private Const FIELD_NCTS_IEM_TMS_ID = "NCTS_IEM_TMS_ID" '4
Private Const FIELD_DATA_NCTS_BGM_Seq1 = "DATA_NCTS_BGM_Seq1" '5
Private Const FIELD_DATA_NCTS_BGM_Seq2 = "DATA_NCTS_BGM_Seq2" '6
Private Const FIELD_DATA_NCTS_BGM_Seq3 = "DATA_NCTS_BGM_Seq3" '7
Private Const FIELD_DATA_NCTS_BGM_Seq4 = "DATA_NCTS_BGM_Seq4" '8
Private Const FIELD_DATA_NCTS_BGM_Seq5 = "DATA_NCTS_BGM_Seq5" '9
Private Const FIELD_DATA_NCTS_BGM_Seq6 = "DATA_NCTS_BGM_Seq6" '10
Private Const FIELD_DATA_NCTS_BGM_Seq7 = "DATA_NCTS_BGM_Seq7" '11
Private Const FIELD_DATA_NCTS_BGM_Seq8 = "DATA_NCTS_BGM_Seq8" '12
Private Const FIELD_DATA_NCTS_BGM_Seq9 = "DATA_NCTS_BGM_Seq9" '13
Private Const FIELD_DATA_NCTS_BGM_Instance = "DATA_NCTS_BGM_Instance" '14

Private Const SQL_ADD_RECORD = "INSERT INTO " & TABLE_NAME & " (" & _
"[DATA_NCTS_BGM_ID],[DATA_NCTS_MSG_ID],[DATA_NCTS_BGM_ParentID]" & _
",[NCTS_IEM_TMS_ID],[DATA_NCTS_BGM_Seq1],[DATA_NCTS_BGM_Seq2]" & _
",[DATA_NCTS_BGM_Seq3],[DATA_NCTS_BGM_Seq4],[DATA_NCTS_BGM_Seq5]" & _
",[DATA_NCTS_BGM_Seq6],[DATA_NCTS_BGM_Seq7],[DATA_NCTS_BGM_Seq8]" & _
",[DATA_NCTS_BGM_Seq9],[DATA_NCTS_BGM_Instance]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [DATA_NCTS_BGM_ID],[DATA_NCTS_MSG_ID],[DATA_NCTS_BGM_ParentID]" & _
",[NCTS_IEM_TMS_ID],[DATA_NCTS_BGM_Seq1],[DATA_NCTS_BGM_Seq2]" & _
",[DATA_NCTS_BGM_Seq3],[DATA_NCTS_BGM_Seq4],[DATA_NCTS_BGM_Seq5]" & _
",[DATA_NCTS_BGM_Seq6],[DATA_NCTS_BGM_Seq7],[DATA_NCTS_BGM_Seq8]" & _
",[DATA_NCTS_BGM_Seq9],[DATA_NCTS_BGM_Instance] FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM " & TABLE_NAME & " WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ParentID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq3) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq4) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq5) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq6) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq7) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq8) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq9) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Instance) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ID)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE " & TABLE_NAME & " SET "
      strSql = strSql & " [DATA_NCTS_MSG_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_MSG_ID) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_ParentID]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ParentID) & ", "
      strSql = strSql & " [NCTS_IEM_TMS_ID]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_IEM_TMS_ID) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq1]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq1) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq2]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq2) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq3]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq3) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq4]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq4) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq5]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq5) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq6]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq6) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq7]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq7) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq8]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq8) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Seq9]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Seq9) & ", "
      strSql = strSql & " [DATA_NCTS_BGM_Instance]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_Instance) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ID)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_DATA_NCTS_BGM_ID)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_BGM_ID = FNullField(rstRecord.Fields("DATA_NCTS_BGM_ID"))
      ActiveRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(rstRecord.Fields("DATA_NCTS_MSG_ID"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_ParentID = FNullField(rstRecord.Fields("DATA_NCTS_BGM_ParentID"))
      ActiveRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(rstRecord.Fields("NCTS_IEM_TMS_ID"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq1 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq1"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq2 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq2"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq3 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq3"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq4 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq4"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq5 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq5"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq6 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq6"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq7 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq7"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq8 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq8"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Seq9 = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Seq9"))
      ActiveRecord.FIELD_DATA_NCTS_BGM_Instance = FNullField(rstRecord.Fields("DATA_NCTS_BGM_Instance"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDATA_NCTS_BGM

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_BGM_ID = rstRecord.Fields("ID_MAX").Value
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiDATA_NCTS_BGM

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_DATA_NCTS_BGM_ID = rstRecord.Fields("ID_MIN").Value
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiDATA_NCTS_BGM) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("DATA_NCTS_BGM_ID") = ActiveRecord.FIELD_DATA_NCTS_BGM_ID
   rstRecord.Fields("DATA_NCTS_MSG_ID") = ActiveRecord.FIELD_DATA_NCTS_MSG_ID
   rstRecord.Fields("DATA_NCTS_BGM_ParentID") = ActiveRecord.FIELD_DATA_NCTS_BGM_ParentID
   rstRecord.Fields("NCTS_IEM_TMS_ID") = ActiveRecord.FIELD_NCTS_IEM_TMS_ID
   rstRecord.Fields("DATA_NCTS_BGM_Seq1") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq1
   rstRecord.Fields("DATA_NCTS_BGM_Seq2") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq2
   rstRecord.Fields("DATA_NCTS_BGM_Seq3") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq3
   rstRecord.Fields("DATA_NCTS_BGM_Seq4") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq4
   rstRecord.Fields("DATA_NCTS_BGM_Seq5") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq5
   rstRecord.Fields("DATA_NCTS_BGM_Seq6") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq6
   rstRecord.Fields("DATA_NCTS_BGM_Seq7") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq7
   rstRecord.Fields("DATA_NCTS_BGM_Seq8") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq8
   rstRecord.Fields("DATA_NCTS_BGM_Seq9") = ActiveRecord.FIELD_DATA_NCTS_BGM_Seq9
   rstRecord.Fields("DATA_NCTS_BGM_Instance") = ActiveRecord.FIELD_DATA_NCTS_BGM_Instance

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiDATA_NCTS_BGM

    Dim clsRecord As cpiDATA_NCTS_BGM
    Set clsRecord = New cpiDATA_NCTS_BGM
    
    If Not ActiveRecord.EOF And Not ActiveRecord.BOF Then
        clsRecord.FIELD_DATA_NCTS_BGM_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_ID))
        clsRecord.FIELD_DATA_NCTS_MSG_ID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_MSG_ID))
        clsRecord.FIELD_DATA_NCTS_BGM_ParentID = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_ParentID))
        clsRecord.FIELD_NCTS_IEM_TMS_ID = FNullField(ActiveRecord.Fields(FIELD_NCTS_IEM_TMS_ID))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq1 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq1))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq2 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq2))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq3 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq3))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq4 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq4))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq5 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq5))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq6 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq6))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq7 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq7))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq8 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq8))
        clsRecord.FIELD_DATA_NCTS_BGM_Seq9 = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Seq9))
        clsRecord.FIELD_DATA_NCTS_BGM_Instance = FNullField(ActiveRecord.Fields(FIELD_DATA_NCTS_BGM_Instance))
    End If
    
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


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiDATA_NCTS_BGM

    'create a new object
    Dim objNewMember As cpiDATA_NCTS_BGM
    Set objNewMember = New cpiDATA_NCTS_BGM

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

Public Property Get Item(vntIndexKey As Variant) As cpiDATA_NCTS_BGM
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




