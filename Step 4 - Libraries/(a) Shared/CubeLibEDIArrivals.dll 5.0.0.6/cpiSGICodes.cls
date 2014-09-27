VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSGICodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GridColumn"
Attribute VB_Ext_KEY = "Member0" ,"GridColumn"
Option Explicit

Dim mvarListIndex As Integer
Dim mvarGroupIndex As Integer
Dim mvarGroupCount As Integer
Dim mvarItemsPerGroup As Integer

Private mCol As Collection
Private mvarRecordset As ADODB.Recordset

Private Const TABLE_NAME = "SGICode"
Private Const PK_FIELD = "CODE"

' field constants
Private Const FIELD_CODE = "CODE" '1
Private Const FIELD_SENSITIVE_GOODS_CODE = "SENSITIVE GOODS CODE" '2
Private Const FIELD_SENSITIVE_QUANTITY = "SENSITIVE QUANTITY" '3
Private Const FIELD_ORDINAL_PARENT = "ORDINAL_PARENT" '4
Private Const FIELD_ORDINAL = "ORDINAL" '5

Private Const SQL_ADD_RECORD = "INSERT INTO [" & TABLE_NAME & "] (" & _
"[CODE],[SENSITIVE GOODS CODE],[SENSITIVE QUANTITY]" & _
",[ORDINAL_PARENT],[ORDINAL]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [CODE],[SENSITIVE GOODS CODE],[SENSITIVE QUANTITY]" & _
",[ORDINAL_PARENT],[ORDINAL] FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SENSITIVE_GOODS_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SENSITIVE_QUANTITY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_ORDINAL_PARENT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_ORDINAL) & ")"

      On Error GoTo ERROR_QUERY
    
      ' MUCP-159 - Start
      ExecuteNonQuery ActiveConnection, strSql
    
      'ActiveConnection.Execute strSql
      ' MUCP-159 - End

      AddRecord = True

      Exit Function

  End If


  AddRecord = False


  Exit Function


  AddRecord = False

ERROR_QUERY:

  AddRecord = False

End Function


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_CODE)


      On Error GoTo ERROR_QUERY
    
      ' MUCP-159 - Start
      ExecuteNonQuery ActiveConnection, strSql
      
      'ActiveConnection.Execute strSql
      ' MUCP-159 - End

      DeleteRecord = True

      Exit Function


   End If


   DeleteRecord = False


   Exit Function


ERROR_QUERY:

   DeleteRecord = False

End Function




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & TABLE_NAME & "] SET "
      strSql = strSql & " [SENSITIVE GOODS CODE]=" & SetEnclosedChr(ActiveRecord.FIELD_SENSITIVE_GOODS_CODE) & ", "
      strSql = strSql & " [SENSITIVE QUANTITY]=" & SetEnclosedChr(ActiveRecord.FIELD_SENSITIVE_QUANTITY) & ", "
      strSql = strSql & " [ORDINAL_PARENT]=" & SetEnclosedChr(ActiveRecord.FIELD_ORDINAL_PARENT) & ", "
      strSql = strSql & " [ORDINAL]=" & SetEnclosedChr(ActiveRecord.FIELD_ORDINAL) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_CODE)

      On Error GoTo ERROR_QUERY

      ' MUCP-159 - Start
      ExecuteNonQuery ActiveConnection, strSql
      
      'ActiveConnection.Execute strSql
      ' MUCP-159 - End

      ModifyRecord = True

      Exit Function


   End If


   ModifyRecord = False


   Exit Function


ERROR_QUERY:

   ModifyRecord = False

End Function




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_CODE)

      On Error GoTo ERROR_QUERY
      
      ' MUCP-159 - Start
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      
      'Set rstRecord = ActiveConnection.Execute(strSql)
      ' MUCP-159 - End
      
      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("CODE"))
      ActiveRecord.FIELD_SENSITIVE_GOODS_CODE = FNullField(rstRecord.Fields("SENSITIVE GOODS CODE"))
      ActiveRecord.FIELD_SENSITIVE_QUANTITY = FNullField(rstRecord.Fields("SENSITIVE QUANTITY"))
      ActiveRecord.FIELD_ORDINAL_PARENT = FNullField(rstRecord.Fields("ORDINAL_PARENT"))
      ActiveRecord.FIELD_ORDINAL = FNullField(rstRecord.Fields("ORDINAL"))
      Set rstRecord = Nothing

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


   Exit Function


ERROR_RECORDSET:
       Set rstRecord = Nothing
       GetRecord = False
       Exit Function
ERROR_QUERY:

   GetRecord = False

End Function




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSGICode

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      
      ' MUCP-159 - Start
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      
      'Set rstRecord = ActiveConnection.Execute(strSql)
      ' MUCP-159 - End
      
      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("ID_MAX"))
      Set rstRecord = Nothing

      GetMaxID = True

      Exit Function


ERROR_RECORDSET:

   Set rstRecord = Nothing

   GetMaxID = False

   Exit Function


ERROR_QUERY:

   GetMaxID = False

End Function




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiSGICode

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      
      ' MUCP-159 - Start
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      
      'Set rstRecord = ActiveConnection.Execute(strSql)
      ' MUCP-159 - End

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_CODE = FNullField(rstRecord.Fields("ID_MIN"))
      Set rstRecord = Nothing

      GetMinID = True

      Exit Function


ERROR_RECORDSET:

   Set rstRecord = Nothing

   GetMinID = False

   Exit Function


ERROR_QUERY:

   GetMinID = False

End Function




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiSGICode) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & TABLE_NAME & "]"
   
   ' MUCP-159 - Start
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      
   'rstRecord.Open strSql, ActiveConnection, adOpenKeyset, adLockOptimistic
   ' MUCP-159 - End
   
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("CODE") = ActiveRecord.FIELD_CODE
   rstRecord.Fields("SENSITIVE GOODS CODE") = ActiveRecord.FIELD_SENSITIVE_GOODS_CODE
   rstRecord.Fields("SENSITIVE QUANTITY") = ActiveRecord.FIELD_SENSITIVE_QUANTITY
   rstRecord.Fields("ORDINAL_PARENT") = ActiveRecord.FIELD_ORDINAL_PARENT
   rstRecord.Fields("ORDINAL") = ActiveRecord.FIELD_ORDINAL

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiSGICode

   Dim clsRecord As cpiSGICode
   Set clsRecord = New cpiSGICode

   clsRecord.FIELD_CODE = FNullField(ActiveRecord.Fields(FIELD_CODE))
   clsRecord.FIELD_SENSITIVE_GOODS_CODE = FNullField(ActiveRecord.Fields(FIELD_SENSITIVE_GOODS_CODE))
   clsRecord.FIELD_SENSITIVE_QUANTITY = FNullField(ActiveRecord.Fields(FIELD_SENSITIVE_QUANTITY))
   clsRecord.FIELD_ORDINAL_PARENT = FNullField(ActiveRecord.Fields(FIELD_ORDINAL_PARENT))
   clsRecord.FIELD_ORDINAL = FNullField(ActiveRecord.Fields(FIELD_ORDINAL))
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
      
      ' MUCP-159 - Start
      ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
   
      SearchRecord = Not (rstDB.EOF And rstDB.BOF)
      
      'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
      '
      'If (rstDB.EOF = False) Then
      '    SearchRecord = True
      'ElseIf (rstDB.EOF = True) Then
      '    SearchRecord = False
      'End If
      ' MUCP-159 - End
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

   ' MUCP-159 - Start
   ADORecordsetOpen CommandText, ActiveConnection, rstRecordset, adOpenKeyset, adLockOptimistic
      
   'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic
   ' MUCP-159 - End

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiSGICode

    'create a new object
    Dim objNewMember As cpiSGICode
    Set objNewMember = New cpiSGICode

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

Public Property Get Item(vntIndexKey As Variant) As cpiSGICode
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
    mvarItemsPerGroup = 4
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

Public Property Get ListIndex() As Integer
    ListIndex = mvarListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Integer)
    mvarListIndex = vNewValue
End Property
'Dim mvarListIndex As Integer

Public Property Get ItemsPerGroup() As Integer
    ItemsPerGroup = mvarItemsPerGroup
End Property

Public Property Let ItemsPerGroup(ByVal vNewValue As Integer)
    mvarItemsPerGroup = vNewValue
End Property
'Dim mvarListIndex As Integer

Public Property Get GroupIndex() As Integer
    GroupIndex = mvarGroupIndex
End Property

Public Property Let GroupIndex(ByVal vNewValue As Integer)

    If (vNewValue <= GroupCount) And (vNewValue > 0) Then
        mvarGroupIndex = vNewValue
    End If
End Property

Public Property Get GroupCount() As Integer

    Dim intGroupTotal As Integer
    Dim intGroupCtr As Integer
    
    Do
        intGroupTotal = intGroupTotal + mvarItemsPerGroup
        intGroupCtr = intGroupCtr + 1
    Loop While (mCol.Count - intGroupTotal) > 0
    
    GroupCount = intGroupCtr
End Property

Public Property Let GroupCount(ByVal vNewValue As Integer)
    mvarGroupCount = vNewValue
End Property

Public Function AddNewGroup(ByVal sParentKey As String, ByVal PKValue As String, ByVal ParentID As Integer)
    
    ' create a new object
    Dim intGroupCtr As Integer
    Dim objNewMember As cpiSGICode
    Dim sKey As String

    For intGroupCtr = 1 To mvarItemsPerGroup
        
        Set objNewMember = New cpiSGICode
    
        sKey = sParentKey & "-" & CStr(mCol.Count + 1)
        ' set the properties passed into the method
        objNewMember.Key = sKey '& "_" & CStr(mCol.Count + 1)
        objNewMember.Index = mCol.Count + 1
        objNewMember.FIELD_CODE = PKValue
        objNewMember.FIELD_ORDINAL = mCol.Count + 1
        objNewMember.FIELD_ORDINAL_PARENT = ParentID
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    
        ' return the object created
        ' Set Add = objNewMember

    Next intGroupCtr
    
    Set objNewMember = Nothing

End Function



