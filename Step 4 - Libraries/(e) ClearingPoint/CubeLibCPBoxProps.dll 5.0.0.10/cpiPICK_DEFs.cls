VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiPICK_DEFs"
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

Private Const TABLE_NAME = "PICKLIST DEFINITION"
Private Const PK_FIELD = "PICKLIST DESCRIPTION ENGLISH"

' field constants
Private Const FIELD_PICKLIST_DESCRIPTION_ENGLISH = "PICKLIST DESCRIPTION ENGLISH" '1
Private Const FIELD_PICKLIST_DESCRIPTION_DUTCH = "PICKLIST DESCRIPTION DUTCH" '2
Private Const FIELD_PICKLIST_DESCRIPTION_FRENCH = "PICKLIST DESCRIPTION FRENCH" '3
Private Const FIELD_DOCUMENT = "DOCUMENT" '4
Private Const FIELD_BOX_CODE = "BOX CODE" '5
Private Const FIELD_FROM = "FROM" '6
Private Const FIELD_TO = "TO" '7
Private Const FIELD_VALIDS = "VALIDS" '8
Private Const FIELD_WIDTH = "WIDTH" '9
Private Const FIELD_AUTO_ADD = "AUTO ADD" '10
Private Const FIELD_DATA_TYPE = "DATA TYPE" '11
Private Const FIELD_INTERNAL_CODE = "INTERNAL CODE" '12

Private Const SQL_ADD_RECORD = "INSERT INTO [" & TABLE_NAME & "] (" & _
"[PICKLIST DESCRIPTION ENGLISH],[PICKLIST DESCRIPTION DUTCH],[PICKLIST DESCRIPTION FRENCH],[DOCUMENT],[BOX CODE],[FROM]" & _
",[TO],[VALIDS],[WIDTH],[AUTO ADD],[DATA TYPE],[INTERNAL CODE]" & _
"                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [PICKLIST DESCRIPTION ENGLISH],[PICKLIST DESCRIPTION DUTCH],[PICKLIST DESCRIPTION FRENCH],[DOCUMENT],[BOX CODE],[FROM]" & _
",[TO],[VALIDS],[WIDTH],[AUTO ADD],[DATA TYPE],[INTERNAL CODE]" & _
" FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_ALL_RECORD = "SELECT [PICKLIST DESCRIPTION ENGLISH],[PICKLIST DESCRIPTION DUTCH],[PICKLIST DESCRIPTION FRENCH],[DOCUMENT],[BOX CODE],[FROM]" & _
",[TO],[VALIDS],[WIDTH],[AUTO ADD],[DATA TYPE],[INTERNAL CODE]" & _
" FROM [" & TABLE_NAME & "]"

Private Const SQL_DELETE_RECORD = "DELETE * FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_DUTCH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_FRENCH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FROM) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TO) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_VALIDS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WIDTH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_AUTO_ADD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_TYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & TABLE_NAME & "] SET "
      strSql = strSql & " [PICKLIST DESCRIPTION DUTCH]=" & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_DUTCH) & ", "
      strSql = strSql & " [PICKLIST DESCRIPTION FRENCH]=" & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_FRENCH) & ", "
      strSql = strSql & " [DOCUMENT]=" & SetEnclosedChr(ActiveRecord.FIELD_DOCUMENT) & ", "
      strSql = strSql & " [BOX CODE]=" & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE) & ", "
      strSql = strSql & " [FROM]=" & SetEnclosedChr(ActiveRecord.FIELD_FROM) & ", "
      strSql = strSql & " [TO]=" & SetEnclosedChr(ActiveRecord.FIELD_TO) & ", "
      strSql = strSql & " [VALIDS]=" & SetEnclosedChr(ActiveRecord.FIELD_VALIDS) & ", "
      strSql = strSql & " [WIDTH]=" & SetEnclosedChr(ActiveRecord.FIELD_WIDTH) & ", "
      strSql = strSql & " [AUTO ADD]=" & SetEnclosedChr(ActiveRecord.FIELD_AUTO_ADD) & ", "
      strSql = strSql & " [DATA TYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_TYPE) & ", "
      strSql = strSql & " [INTERNAL CODE]=" & SetEnclosedChr(ActiveRecord.FIELD_INTERNAL_CODE) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH = FNullField(rstRecord.Fields("PICKLIST DESCRIPTION ENGLISH"))
      ActiveRecord.FIELD_PICKLIST_DESCRIPTION_DUTCH = FNullField(rstRecord.Fields("PICKLIST DESCRIPTION DUTCH"))
      ActiveRecord.FIELD_PICKLIST_DESCRIPTION_FRENCH = FNullField(rstRecord.Fields("PICKLIST DESCRIPTION FRENCH"))
      ActiveRecord.FIELD_DOCUMENT = FNullField(rstRecord.Fields("DOCUMENT"))
      ActiveRecord.FIELD_BOX_CODE = FNullField(rstRecord.Fields("BOX CODE"))
      ActiveRecord.FIELD_FROM = FNullField(rstRecord.Fields("FROM"))
      ActiveRecord.FIELD_TO = FNullField(rstRecord.Fields("TO"))
      ActiveRecord.FIELD_VALIDS = FNullField(rstRecord.Fields("VALIDS"))
      ActiveRecord.FIELD_WIDTH = FNullField(rstRecord.Fields("WIDTH"))
      ActiveRecord.FIELD_AUTO_ADD = FNullField(rstRecord.Fields("AUTO ADD"))
      ActiveRecord.FIELD_DATA_TYPE = FNullField(rstRecord.Fields("DATA TYPE"))
      ActiveRecord.FIELD_INTERNAL_CODE = FNullField(rstRecord.Fields("INTERNAL CODE"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiPICK_DEF

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH = FNullField(rstRecord.Fields("ID_MAX"))
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiPICK_DEF

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH = FNullField(rstRecord.Fields("ID_MIN"))
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiPICK_DEF) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("PICKLIST DESCRIPTION ENGLISH") = ActiveRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH
   rstRecord.Fields("PICKLIST DESCRIPTION DUTCH") = ActiveRecord.FIELD_PICKLIST_DESCRIPTION_DUTCH
   rstRecord.Fields("PICKLIST DESCRIPTION FRENCH") = ActiveRecord.FIELD_PICKLIST_DESCRIPTION_FRENCH
   rstRecord.Fields("DOCUMENT") = ActiveRecord.FIELD_DOCUMENT
   rstRecord.Fields("BOX CODE") = ActiveRecord.FIELD_BOX_CODE
   rstRecord.Fields("FROM") = ActiveRecord.FIELD_FROM
   rstRecord.Fields("TO") = ActiveRecord.FIELD_TO
   rstRecord.Fields("VALIDS") = ActiveRecord.FIELD_VALIDS
   rstRecord.Fields("WIDTH") = ActiveRecord.FIELD_WIDTH
   rstRecord.Fields("AUTO ADD") = ActiveRecord.FIELD_AUTO_ADD
   rstRecord.Fields("DATA TYPE") = ActiveRecord.FIELD_DATA_TYPE
   rstRecord.Fields("INTERNAL CODE") = ActiveRecord.FIELD_INTERNAL_CODE

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiPICK_DEF

   Dim clsRecord As cpiPICK_DEF
   Set clsRecord = New cpiPICK_DEF

   clsRecord.FIELD_PICKLIST_DESCRIPTION_ENGLISH = FNullField(ActiveRecord.Fields(FIELD_PICKLIST_DESCRIPTION_ENGLISH))
   clsRecord.FIELD_PICKLIST_DESCRIPTION_DUTCH = FNullField(ActiveRecord.Fields(FIELD_PICKLIST_DESCRIPTION_DUTCH))
   clsRecord.FIELD_PICKLIST_DESCRIPTION_FRENCH = FNullField(ActiveRecord.Fields(FIELD_PICKLIST_DESCRIPTION_FRENCH))
   clsRecord.FIELD_DOCUMENT = FNullField(ActiveRecord.Fields(FIELD_DOCUMENT))
   clsRecord.FIELD_BOX_CODE = FNullField(ActiveRecord.Fields(FIELD_BOX_CODE))
   clsRecord.FIELD_FROM = FNullField(ActiveRecord.Fields(FIELD_FROM))
   clsRecord.FIELD_TO = FNullField(ActiveRecord.Fields(FIELD_TO))
   clsRecord.FIELD_VALIDS = FNullField(ActiveRecord.Fields(FIELD_VALIDS))
   clsRecord.FIELD_WIDTH = FNullField(ActiveRecord.Fields(FIELD_WIDTH))
   clsRecord.FIELD_AUTO_ADD = FNullField(ActiveRecord.Fields(FIELD_AUTO_ADD))
   clsRecord.FIELD_DATA_TYPE = FNullField(ActiveRecord.Fields(FIELD_DATA_TYPE))
   clsRecord.FIELD_INTERNAL_CODE = FNullField(ActiveRecord.Fields(FIELD_INTERNAL_CODE))
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
      strSql = "SELECT TOP 1 " & SearchField & " FROM [" & TABLE_NAME & "] WHERE " & SearchField & "=" & SetEnclosedChr((SearchValue))
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
      CommandText = SQL_GET_ALL_RECORD
   End If

   rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiPICK_DEF

    'create a new object
    Dim objNewMember As cpiPICK_DEF
    Set objNewMember = New cpiPICK_DEF

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

Public Property Get Item(vntIndexKey As Variant) As cpiPICK_DEF
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




