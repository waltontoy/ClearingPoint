VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSHEET_PROPS_Tbls"
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

Private Const TABLE_NAME = "SHEET PROPERTIES"
Private Const PK_FIELD = "user no"

' field constants
Private Const FIELD_user_no = "user no" '1
Private Const FIELD_font_name = "font name" '2
Private Const FIELD_bold = "bold" '3
Private Const FIELD_italic = "italic" '4
Private Const FIELD_BC_active_box = "BC active box" '5
Private Const FIELD_BC_inactive_box = "BC inactive box" '6
Private Const FIELD_BC_disabled_box = "BC disabled box" '7
Private Const FIELD_BC_werror_active = "BC werror active" '8
Private Const FIELD_BC_werror_inactive = "BC werror inactive" '9
Private Const FIELD_FC_active_box = "FC active box" '10
Private Const FIELD_FC_inactive_box = "FC inactive box" '11
Private Const FIELD_FC_disabled_box = "FC disabled box" '12
Private Const FIELD_FC_werror_active = "FC werror active" '13
Private Const FIELD_FC_werror_inactive = "FC werror inactive" '14
Private Const FIELD_size = "size" '15

Private Const SQL_ADD_RECORD = "INSERT INTO [" & TABLE_NAME & "] (" & _
"[user no],[font name],[bold],[italic],[BC active box],[BC inactive box]" & _
",[BC disabled box],[BC werror active],[BC werror inactive],[FC active box],[FC inactive box],[FC disabled box]" & _
",[FC werror active],[FC werror inactive],[size]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [user no],[font name],[bold],[italic],[BC active box],[BC inactive box]" & _
",[BC disabled box],[BC werror active],[BC werror inactive],[FC active box],[FC inactive box],[FC disabled box]" & _
",[FC werror active],[FC werror inactive],[size] FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_ALL_RECORD = "SELECT [user no],[font name],[bold],[italic],[BC active box],[BC inactive box]" & _
",[BC disabled box],[BC werror active],[BC werror inactive],[FC active box],[FC inactive box],[FC disabled box]" & _
",[FC werror active],[FC werror inactive],[size] FROM [" & TABLE_NAME & "]"

Private Const SQL_DELETE_RECORD = "DELETE * FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_user_no) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_font_name) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_bold) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_italic) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BC_active_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BC_inactive_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BC_disabled_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BC_werror_active) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BC_werror_inactive) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FC_active_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FC_inactive_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FC_disabled_box) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FC_werror_active) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FC_werror_inactive) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_size) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_user_no)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & TABLE_NAME & "] SET "
      strSql = strSql & " [font name]=" & SetEnclosedChr(ActiveRecord.FIELD_font_name) & ", "
      strSql = strSql & " [bold]=" & SetEnclosedChr(ActiveRecord.FIELD_bold) & ", "
      strSql = strSql & " [italic]=" & SetEnclosedChr(ActiveRecord.FIELD_italic) & ", "
      strSql = strSql & " [BC active box]=" & SetEnclosedChr(ActiveRecord.FIELD_BC_active_box) & ", "
      strSql = strSql & " [BC inactive box]=" & SetEnclosedChr(ActiveRecord.FIELD_BC_inactive_box) & ", "
      strSql = strSql & " [BC disabled box]=" & SetEnclosedChr(ActiveRecord.FIELD_BC_disabled_box) & ", "
      strSql = strSql & " [BC werror active]=" & SetEnclosedChr(ActiveRecord.FIELD_BC_werror_active) & ", "
      strSql = strSql & " [BC werror inactive]=" & SetEnclosedChr(ActiveRecord.FIELD_BC_werror_inactive) & ", "
      strSql = strSql & " [FC active box]=" & SetEnclosedChr(ActiveRecord.FIELD_FC_active_box) & ", "
      strSql = strSql & " [FC inactive box]=" & SetEnclosedChr(ActiveRecord.FIELD_FC_inactive_box) & ", "
      strSql = strSql & " [FC disabled box]=" & SetEnclosedChr(ActiveRecord.FIELD_FC_disabled_box) & ", "
      strSql = strSql & " [FC werror active]=" & SetEnclosedChr(ActiveRecord.FIELD_FC_werror_active) & ", "
      strSql = strSql & " [FC werror inactive]=" & SetEnclosedChr(ActiveRecord.FIELD_FC_werror_inactive) & ", "
      strSql = strSql & " [size]=" & SetEnclosedChr(ActiveRecord.FIELD_size) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_user_no)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_user_no)

      On Error GoTo ERROR_QUERY
      
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_user_no = FNullField(rstRecord.Fields("user no"))
      ActiveRecord.FIELD_font_name = FNullField(rstRecord.Fields("font name"))
      ActiveRecord.FIELD_bold = FNullField(rstRecord.Fields("bold"))
      ActiveRecord.FIELD_italic = FNullField(rstRecord.Fields("italic"))
      ActiveRecord.FIELD_BC_active_box = FNullField(rstRecord.Fields("BC active box"))
      ActiveRecord.FIELD_BC_inactive_box = FNullField(rstRecord.Fields("BC inactive box"))
      ActiveRecord.FIELD_BC_disabled_box = FNullField(rstRecord.Fields("BC disabled box"))
      ActiveRecord.FIELD_BC_werror_active = FNullField(rstRecord.Fields("BC werror active"))
      ActiveRecord.FIELD_BC_werror_inactive = FNullField(rstRecord.Fields("BC werror inactive"))
      ActiveRecord.FIELD_FC_active_box = FNullField(rstRecord.Fields("FC active box"))
      ActiveRecord.FIELD_FC_inactive_box = FNullField(rstRecord.Fields("FC inactive box"))
      ActiveRecord.FIELD_FC_disabled_box = FNullField(rstRecord.Fields("FC disabled box"))
      ActiveRecord.FIELD_FC_werror_active = FNullField(rstRecord.Fields("FC werror active"))
      ActiveRecord.FIELD_FC_werror_inactive = FNullField(rstRecord.Fields("FC werror inactive"))
      ActiveRecord.FIELD_size = FNullField(rstRecord.Fields("size"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSHEET_PROPS_Tbl

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_user_no = FNullField(rstRecord.Fields("ID_MAX"))
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSHEET_PROPS_Tbl

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_user_no = FNullField(rstRecord.Fields("ID_MIN"))
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSHEET_PROPS_Tbl) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("user no") = ActiveRecord.FIELD_user_no
   rstRecord.Fields("font name") = ActiveRecord.FIELD_font_name
   rstRecord.Fields("bold") = ActiveRecord.FIELD_bold
   rstRecord.Fields("italic") = ActiveRecord.FIELD_italic
   rstRecord.Fields("BC active box") = ActiveRecord.FIELD_BC_active_box
   rstRecord.Fields("BC inactive box") = ActiveRecord.FIELD_BC_inactive_box
   rstRecord.Fields("BC disabled box") = ActiveRecord.FIELD_BC_disabled_box
   rstRecord.Fields("BC werror active") = ActiveRecord.FIELD_BC_werror_active
   rstRecord.Fields("BC werror inactive") = ActiveRecord.FIELD_BC_werror_inactive
   rstRecord.Fields("FC active box") = ActiveRecord.FIELD_FC_active_box
   rstRecord.Fields("FC inactive box") = ActiveRecord.FIELD_FC_inactive_box
   rstRecord.Fields("FC disabled box") = ActiveRecord.FIELD_FC_disabled_box
   rstRecord.Fields("FC werror active") = ActiveRecord.FIELD_FC_werror_active
   rstRecord.Fields("FC werror inactive") = ActiveRecord.FIELD_FC_werror_inactive
   rstRecord.Fields("size") = ActiveRecord.FIELD_size

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiSHEET_PROPS_Tbl

   Dim clsRecord As cpiSHEET_PROPS_Tbl
   Set clsRecord = New cpiSHEET_PROPS_Tbl

   clsRecord.FIELD_user_no = FNullField(ActiveRecord.Fields(FIELD_user_no))
   clsRecord.FIELD_font_name = FNullField(ActiveRecord.Fields(FIELD_font_name))
   clsRecord.FIELD_bold = FNullField(ActiveRecord.Fields(FIELD_bold))
   clsRecord.FIELD_italic = FNullField(ActiveRecord.Fields(FIELD_italic))
   clsRecord.FIELD_BC_active_box = FNullField(ActiveRecord.Fields(FIELD_BC_active_box))
   clsRecord.FIELD_BC_inactive_box = FNullField(ActiveRecord.Fields(FIELD_BC_inactive_box))
   clsRecord.FIELD_BC_disabled_box = FNullField(ActiveRecord.Fields(FIELD_BC_disabled_box))
   clsRecord.FIELD_BC_werror_active = FNullField(ActiveRecord.Fields(FIELD_BC_werror_active))
   clsRecord.FIELD_BC_werror_inactive = FNullField(ActiveRecord.Fields(FIELD_BC_werror_inactive))
   clsRecord.FIELD_FC_active_box = FNullField(ActiveRecord.Fields(FIELD_FC_active_box))
   clsRecord.FIELD_FC_inactive_box = FNullField(ActiveRecord.Fields(FIELD_FC_inactive_box))
   clsRecord.FIELD_FC_disabled_box = FNullField(ActiveRecord.Fields(FIELD_FC_disabled_box))
   clsRecord.FIELD_FC_werror_active = FNullField(ActiveRecord.Fields(FIELD_FC_werror_active))
   clsRecord.FIELD_FC_werror_inactive = FNullField(ActiveRecord.Fields(FIELD_FC_werror_inactive))
   clsRecord.FIELD_size = FNullField(ActiveRecord.Fields(FIELD_size))
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


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiSHEET_PROPS_Tbl

    'create a new object
    Dim objNewMember As cpiSHEET_PROPS_Tbl
    Set objNewMember = New cpiSHEET_PROPS_Tbl

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

Public Property Get Item(vntIndexKey As Variant) As cpiSHEET_PROPS_Tbl
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




