VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBOX_DEF_ADMINs"
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
Private mvarTableName  As String

'Private Const mvarTableName = "BOX DEFAULT ADMIN"
Private Const PK_FIELD = "BOX CODE"

' field constants
Private Const FIELD_BOX_CODE = "BOX CODE" '1
Private Const FIELD_ENGLISH_DESCRIPTION = "ENGLISH DESCRIPTION" '2
Private Const FIELD_DUTCH_DESCRIPTION = "DUTCH DESCRIPTION" '3
Private Const FIELD_FRENCH_DESCRIPTION = "FRENCH DESCRIPTION" '4
Private Const FIELD_EMPTY_FIELD_VALUE = "EMPTY FIELD VALUE" '5
Private Const FIELD_INSERT = "INSERT" '6
Private Const FIELD_JUSTIFY = "JUSTIFY" '7
Private Const FIELD_SEQUENTIAL_TABBING_DEFAULT = "SEQUENTIAL TABBING DEFAULT" '8
Private Const FIELD_SEQUENTIAL_TABBING_ACTIVE = "SEQUENTIAL TABBING ACTIVE" '9
Private Const FIELD_CHECK_VAT = "CHECK VAT" '10
Private Const FIELD_CALCULATE_NET_WEIGHT = "CALCULATE NET WEIGHT" '11
Private Const FIELD_CALCULATE_NO_OF_ITEMS = "CALCULATE NO OF ITEMS" '12
Private Const FIELD_COPY_TO_NEXT = "COPY TO NEXT" '13
Private Const FIELD_CHANGE_ONLY_IN_H1 = "CHANGE ONLY IN H1" '14
Private Const FIELD_SEND_ONLY_IN_H1 = "SEND ONLY IN H1" '15
Private Const FIELD_WIDTH = "WIDTH" '16
Private Const FIELD_DECIMAL = "DECIMAL" '17
Private Const FIELD_DATA_TYPE = "DATA TYPE" '18
Private Const FIELD_AUTO_ADD = "AUTO ADD" '19
Private Const FIELD_DEFAULT_VALUE = "DEFAULT VALUE" '20
Private Const FIELD_CALCULATE_CUSTOMS_VALUE = "CALCULATE CUSTOMS VALUE" '21
Private Const FIELD_RELATE_L1_TO_S1 = "RELATE L1 TO S1" '22
Private Const FIELD_VALIDATE_VALUE = "VALIDATE_VALUE" '23

Private strSqlAddRecord As String
'Private Const strSqlAddRecord = "INSERT INTO [" & mvarTableName & "] (" & _
"[BOX CODE],[ENGLISH DESCRIPTION],[DUTCH DESCRIPTION],[FRENCH DESCRIPTION],[EMPTY FIELD VALUE],[INSERT]" & _
",[JUSTIFY],[SEQUENTIAL TABBING DEFAULT],[SEQUENTIAL TABBING ACTIVE],[CHECK VAT],[CALCULATE NET WEIGHT],[CALCULATE NO OF ITEMS]" & _
",[COPY TO NEXT],[CHANGE ONLY IN H1],[SEND ONLY IN H1],[WIDTH],[DECIMAL],[DATA TYPE]" & _
",[AUTO ADD],[DEFAULT VALUE],[CALCULATE CUSTOMS VALUE],[RELATE L1 TO S1]                                     )  VALUES "

Private strSqlGetRecord As String
'Private Const strSqlGetRecord = "SELECT [BOX CODE],[ENGLISH DESCRIPTION],[DUTCH DESCRIPTION],[FRENCH DESCRIPTION],[EMPTY FIELD VALUE],[INSERT]" & _
",[JUSTIFY],[SEQUENTIAL TABBING DEFAULT],[SEQUENTIAL TABBING ACTIVE],[CHECK VAT],[CALCULATE NET WEIGHT],[CALCULATE NO OF ITEMS]" & _
",[COPY TO NEXT],[CHANGE ONLY IN H1],[SEND ONLY IN H1],[WIDTH],[DECIMAL],[DATA TYPE]" & _
",[AUTO ADD],[DEFAULT VALUE],[CALCULATE CUSTOMS VALUE],[RELATE L1 TO S1] FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

Private strSqlGetAllRecord As String
'Private Const strSqlGetAllRecord = "SELECT [BOX CODE],[ENGLISH DESCRIPTION],[DUTCH DESCRIPTION],[FRENCH DESCRIPTION],[EMPTY FIELD VALUE],[INSERT]" & _
",[JUSTIFY],[SEQUENTIAL TABBING DEFAULT],[SEQUENTIAL TABBING ACTIVE],[CHECK VAT],[CALCULATE NET WEIGHT],[CALCULATE NO OF ITEMS]" & _
",[COPY TO NEXT],[CHANGE ONLY IN H1],[SEND ONLY IN H1],[WIDTH],[DECIMAL],[DATA TYPE]" & _
",[AUTO ADD],[DEFAULT VALUE],[CALCULATE CUSTOMS VALUE],[RELATE L1 TO S1] FROM [" & mvarTableName & "]"

Private strSqlDeleteRecord As String
'Private Const strSqlDeleteRecord = "DELETE * FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

Private strSqlGetMaxID As String
'Private Const strSqlGetMaxID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & mvarTableName & "]"

Private strSqlGetMinID As String
'Private Const strSqlGetMinID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & mvarTableName & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSqlAddRecord & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DUTCH_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_FRENCH_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EMPTY_FIELD_VALUE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_INSERT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_JUSTIFY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEQUENTIAL_TABBING_DEFAULT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEQUENTIAL_TABBING_ACTIVE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CHECK_VAT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_NET_WEIGHT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_NO_OF_ITEMS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_COPY_TO_NEXT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CHANGE_ONLY_IN_H1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_ONLY_IN_H1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_WIDTH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DECIMAL) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DATA_TYPE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_AUTO_ADD) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_DEFAULT_VALUE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_CUSTOMS_VALUE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_RELATE_L1_TO_S1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_VALIDATE_VALUE) & ")"

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


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSqlDeleteRecord & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE)


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




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = "UPDATE [" & mvarTableName & "] SET "
      strSql = strSql & " [ENGLISH DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_ENGLISH_DESCRIPTION) & ", "
      strSql = strSql & " [DUTCH DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_DUTCH_DESCRIPTION) & ", "
      strSql = strSql & " [FRENCH DESCRIPTION]=" & SetEnclosedChr(ActiveRecord.FIELD_FRENCH_DESCRIPTION) & ", "
      strSql = strSql & " [EMPTY FIELD VALUE]=" & SetEnclosedChr(ActiveRecord.FIELD_EMPTY_FIELD_VALUE) & ", "
      strSql = strSql & " [INSERT]=" & SetEnclosedChr(ActiveRecord.FIELD_INSERT) & ", "
      strSql = strSql & " [JUSTIFY]=" & SetEnclosedChr(ActiveRecord.FIELD_JUSTIFY) & ", "
      strSql = strSql & " [SEQUENTIAL TABBING DEFAULT]=" & SetEnclosedChr(ActiveRecord.FIELD_SEQUENTIAL_TABBING_DEFAULT) & ", "
      strSql = strSql & " [SEQUENTIAL TABBING ACTIVE]=" & SetEnclosedChr(ActiveRecord.FIELD_SEQUENTIAL_TABBING_ACTIVE) & ", "
      strSql = strSql & " [CHECK VAT]=" & SetEnclosedChr(ActiveRecord.FIELD_CHECK_VAT) & ", "
      strSql = strSql & " [CALCULATE NET WEIGHT]=" & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_NET_WEIGHT) & ", "
      strSql = strSql & " [CALCULATE NO OF ITEMS]=" & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_NO_OF_ITEMS) & ", "
      strSql = strSql & " [COPY TO NEXT]=" & SetEnclosedChr(ActiveRecord.FIELD_COPY_TO_NEXT) & ", "
      strSql = strSql & " [CHANGE ONLY IN H1]=" & SetEnclosedChr(ActiveRecord.FIELD_CHANGE_ONLY_IN_H1) & ", "
      strSql = strSql & " [SEND ONLY IN H1]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_ONLY_IN_H1) & ", "
      strSql = strSql & " [WIDTH]=" & SetEnclosedChr(ActiveRecord.FIELD_WIDTH) & ", "
      strSql = strSql & " [DECIMAL]=" & SetEnclosedChr(ActiveRecord.FIELD_DECIMAL) & ", "
      strSql = strSql & " [DATA TYPE]=" & SetEnclosedChr(ActiveRecord.FIELD_DATA_TYPE) & ", "
      strSql = strSql & " [AUTO ADD]=" & SetEnclosedChr(ActiveRecord.FIELD_AUTO_ADD) & ", "
      strSql = strSql & " [DEFAULT VALUE]=" & SetEnclosedChr(ActiveRecord.FIELD_DEFAULT_VALUE) & ", "
      strSql = strSql & " [CALCULATE CUSTOMS VALUE]=" & SetEnclosedChr(ActiveRecord.FIELD_CALCULATE_CUSTOMS_VALUE) & ", "
      strSql = strSql & " [RELATE L1 TO S1]=" & SetEnclosedChr(ActiveRecord.FIELD_RELATE_L1_TO_S1) & ", "
      strSql = strSql & " [VALIDATE_VALUE]=" & SetEnclosedChr(ActiveRecord.FIELD_VALIDATE_VALUE) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE)

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




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSqlGetRecord & SetEnclosedChr(ActiveRecord.FIELD_BOX_CODE)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_BOX_CODE = FNullField(rstRecord.Fields("BOX CODE"))
      ActiveRecord.FIELD_ENGLISH_DESCRIPTION = FNullField(rstRecord.Fields("ENGLISH DESCRIPTION"))
      ActiveRecord.FIELD_DUTCH_DESCRIPTION = FNullField(rstRecord.Fields("DUTCH DESCRIPTION"))
      ActiveRecord.FIELD_FRENCH_DESCRIPTION = FNullField(rstRecord.Fields("FRENCH DESCRIPTION"))
      ActiveRecord.FIELD_EMPTY_FIELD_VALUE = FNullField(rstRecord.Fields("EMPTY FIELD VALUE"))
      ActiveRecord.FIELD_INSERT = FNullField(rstRecord.Fields("INSERT"))
      ActiveRecord.FIELD_JUSTIFY = FNullField(rstRecord.Fields("JUSTIFY"))
      ActiveRecord.FIELD_SEQUENTIAL_TABBING_DEFAULT = FNullField(rstRecord.Fields("SEQUENTIAL TABBING DEFAULT"))
      ActiveRecord.FIELD_SEQUENTIAL_TABBING_ACTIVE = FNullField(rstRecord.Fields("SEQUENTIAL TABBING ACTIVE"))
      ActiveRecord.FIELD_CHECK_VAT = FNullField(rstRecord.Fields("CHECK VAT"))
      ActiveRecord.FIELD_CALCULATE_NET_WEIGHT = FNullField(rstRecord.Fields("CALCULATE NET WEIGHT"))
      ActiveRecord.FIELD_CALCULATE_NO_OF_ITEMS = FNullField(rstRecord.Fields("CALCULATE NO OF ITEMS"))
      ActiveRecord.FIELD_COPY_TO_NEXT = FNullField(rstRecord.Fields("COPY TO NEXT"))
      ActiveRecord.FIELD_CHANGE_ONLY_IN_H1 = FNullField(rstRecord.Fields("CHANGE ONLY IN H1"))
      ActiveRecord.FIELD_SEND_ONLY_IN_H1 = FNullField(rstRecord.Fields("SEND ONLY IN H1"))
      ActiveRecord.FIELD_WIDTH = FNullField(rstRecord.Fields("WIDTH"))
      ActiveRecord.FIELD_DECIMAL = FNullField(rstRecord.Fields("DECIMAL"))
      ActiveRecord.FIELD_DATA_TYPE = FNullField(rstRecord.Fields("DATA TYPE"))
      ActiveRecord.FIELD_AUTO_ADD = FNullField(rstRecord.Fields("AUTO ADD"))
      ActiveRecord.FIELD_DEFAULT_VALUE = FNullField(rstRecord.Fields("DEFAULT VALUE"))
      ActiveRecord.FIELD_CALCULATE_CUSTOMS_VALUE = FNullField(rstRecord.Fields("CALCULATE CUSTOMS VALUE"))
      ActiveRecord.FIELD_RELATE_L1_TO_S1 = FNullField(rstRecord.Fields("RELATE L1 TO S1"))
      ActiveRecord.FIELD_VALIDATE_VALUE = FNullField(rstRecord.Fields("VALIDATE_VALUE"))
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




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiBOX_DEF_ADMIN

   End If

      strSql = strSqlGetMaxID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_BOX_CODE = FNullField(rstRecord.Fields("ID_MAX"))
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




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiBOX_DEF_ADMIN

   End If

      strSql = strSqlGetMinID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_BOX_CODE = FNullField(rstRecord.Fields("ID_MIN"))
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




Public Function GetTableRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiBOX_DEF_ADMIN) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset



   strSql = "SELECT * FROM [" & mvarTableName & "]"
   ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
   Set rstRecord = RstCopy(rstRecord, True, 0, 0, 1, True)

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("BOX CODE") = ActiveRecord.FIELD_BOX_CODE
   rstRecord.Fields("ENGLISH DESCRIPTION") = ActiveRecord.FIELD_ENGLISH_DESCRIPTION
   rstRecord.Fields("DUTCH DESCRIPTION") = ActiveRecord.FIELD_DUTCH_DESCRIPTION
   rstRecord.Fields("FRENCH DESCRIPTION") = ActiveRecord.FIELD_FRENCH_DESCRIPTION
   rstRecord.Fields("EMPTY FIELD VALUE") = ActiveRecord.FIELD_EMPTY_FIELD_VALUE
   rstRecord.Fields("INSERT") = ActiveRecord.FIELD_INSERT
   rstRecord.Fields("JUSTIFY") = ActiveRecord.FIELD_JUSTIFY
   rstRecord.Fields("SEQUENTIAL TABBING DEFAULT") = ActiveRecord.FIELD_SEQUENTIAL_TABBING_DEFAULT
   rstRecord.Fields("SEQUENTIAL TABBING ACTIVE") = ActiveRecord.FIELD_SEQUENTIAL_TABBING_ACTIVE
   rstRecord.Fields("CHECK VAT") = ActiveRecord.FIELD_CHECK_VAT
   rstRecord.Fields("CALCULATE NET WEIGHT") = ActiveRecord.FIELD_CALCULATE_NET_WEIGHT
   rstRecord.Fields("CALCULATE NO OF ITEMS") = ActiveRecord.FIELD_CALCULATE_NO_OF_ITEMS
   rstRecord.Fields("COPY TO NEXT") = ActiveRecord.FIELD_COPY_TO_NEXT
   rstRecord.Fields("CHANGE ONLY IN H1") = ActiveRecord.FIELD_CHANGE_ONLY_IN_H1
   rstRecord.Fields("SEND ONLY IN H1") = ActiveRecord.FIELD_SEND_ONLY_IN_H1
   rstRecord.Fields("WIDTH") = ActiveRecord.FIELD_WIDTH
   rstRecord.Fields("DECIMAL") = ActiveRecord.FIELD_DECIMAL
   rstRecord.Fields("DATA TYPE") = ActiveRecord.FIELD_DATA_TYPE
   rstRecord.Fields("AUTO ADD") = ActiveRecord.FIELD_AUTO_ADD
   rstRecord.Fields("DEFAULT VALUE") = ActiveRecord.FIELD_DEFAULT_VALUE
   rstRecord.Fields("CALCULATE CUSTOMS VALUE") = ActiveRecord.FIELD_CALCULATE_CUSTOMS_VALUE
   rstRecord.Fields("RELATE L1 TO S1") = ActiveRecord.FIELD_RELATE_L1_TO_S1
   rstRecord.Fields("VALIDATE_VALUE") = ActiveRecord.FIELD_VALIDATE_VALUE

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiBOX_DEF_ADMIN

   Dim clsRecord As cpiBOX_DEF_ADMIN
   Set clsRecord = New cpiBOX_DEF_ADMIN

   clsRecord.FIELD_BOX_CODE = FNullField(ActiveRecord.Fields(FIELD_BOX_CODE))
   clsRecord.FIELD_ENGLISH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_ENGLISH_DESCRIPTION))
   clsRecord.FIELD_DUTCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_DUTCH_DESCRIPTION))
   clsRecord.FIELD_FRENCH_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_FRENCH_DESCRIPTION))
   clsRecord.FIELD_EMPTY_FIELD_VALUE = FNullField(ActiveRecord.Fields(FIELD_EMPTY_FIELD_VALUE))
   clsRecord.FIELD_INSERT = FNullField(ActiveRecord.Fields(FIELD_INSERT))
   clsRecord.FIELD_JUSTIFY = FNullField(ActiveRecord.Fields(FIELD_JUSTIFY))
   clsRecord.FIELD_SEQUENTIAL_TABBING_DEFAULT = FNullField(ActiveRecord.Fields(FIELD_SEQUENTIAL_TABBING_DEFAULT))
   clsRecord.FIELD_SEQUENTIAL_TABBING_ACTIVE = FNullField(ActiveRecord.Fields(FIELD_SEQUENTIAL_TABBING_ACTIVE))
   clsRecord.FIELD_CHECK_VAT = FNullField(ActiveRecord.Fields(FIELD_CHECK_VAT))
   clsRecord.FIELD_CALCULATE_NET_WEIGHT = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_NET_WEIGHT))
   clsRecord.FIELD_CALCULATE_NO_OF_ITEMS = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_NO_OF_ITEMS))
   clsRecord.FIELD_COPY_TO_NEXT = FNullField(ActiveRecord.Fields(FIELD_COPY_TO_NEXT))
   clsRecord.FIELD_CHANGE_ONLY_IN_H1 = FNullField(ActiveRecord.Fields(FIELD_CHANGE_ONLY_IN_H1))
   clsRecord.FIELD_SEND_ONLY_IN_H1 = FNullField(ActiveRecord.Fields(FIELD_SEND_ONLY_IN_H1))
   clsRecord.FIELD_WIDTH = FNullField(ActiveRecord.Fields(FIELD_WIDTH))
   clsRecord.FIELD_DECIMAL = FNullField(ActiveRecord.Fields(FIELD_DECIMAL))
   clsRecord.FIELD_DATA_TYPE = FNullField(ActiveRecord.Fields(FIELD_DATA_TYPE))
   clsRecord.FIELD_AUTO_ADD = FNullField(ActiveRecord.Fields(FIELD_AUTO_ADD))
   clsRecord.FIELD_DEFAULT_VALUE = FNullField(ActiveRecord.Fields(FIELD_DEFAULT_VALUE))
   clsRecord.FIELD_CALCULATE_CUSTOMS_VALUE = FNullField(ActiveRecord.Fields(FIELD_CALCULATE_CUSTOMS_VALUE))
   clsRecord.FIELD_RELATE_L1_TO_S1 = FNullField(ActiveRecord.Fields(FIELD_RELATE_L1_TO_S1))
   clsRecord.FIELD_VALIDATE_VALUE = FNullField(ActiveRecord.Fields(FIELD_VALIDATE_VALUE))
   
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
      CommandText = strSqlGetAllRecord
   End If

   rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiBOX_DEF_ADMIN

    'create a new object
    Dim objNewMember As cpiBOX_DEF_ADMIN
    Set objNewMember = New cpiBOX_DEF_ADMIN

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

Public Property Get Item(vntIndexKey As Variant) As cpiBOX_DEF_ADMIN
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

'  2
Public Property Let TableName(ByVal vData As String)
   mvarTableName = vData
End Property

Public Property Get TableName() As String
   TableName = mvarTableName
End Property

Public Function SetSqlParameters(ByVal TableName As String) As Boolean
'
    mvarTableName = TableName
    
    strSqlAddRecord = "INSERT INTO [" & mvarTableName & "] (" & _
    "[BOX CODE],[ENGLISH DESCRIPTION],[DUTCH DESCRIPTION],[FRENCH DESCRIPTION],[EMPTY FIELD VALUE],[INSERT]" & _
    ",[JUSTIFY],[SEQUENTIAL TABBING DEFAULT],[SEQUENTIAL TABBING ACTIVE],[CHECK VAT],[CALCULATE NET WEIGHT],[CALCULATE NO OF ITEMS]" & _
    ",[COPY TO NEXT],[CHANGE ONLY IN H1],[SEND ONLY IN H1],[WIDTH],[DECIMAL],[DATA TYPE]" & _
    ",[AUTO ADD],[DEFAULT VALUE],[CALCULATE CUSTOMS VALUE],[RELATE L1 TO S1],[VALIDATE_VALUE]                                     )  VALUES "

    strSqlGetRecord = "SELECT [BOX CODE],[ENGLISH DESCRIPTION],[DUTCH DESCRIPTION],[FRENCH DESCRIPTION],[EMPTY FIELD VALUE],[INSERT]" & _
    ",[JUSTIFY],[SEQUENTIAL TABBING DEFAULT],[SEQUENTIAL TABBING ACTIVE],[CHECK VAT],[CALCULATE NET WEIGHT],[CALCULATE NO OF ITEMS]" & _
    ",[COPY TO NEXT],[CHANGE ONLY IN H1],[SEND ONLY IN H1],[WIDTH],[DECIMAL],[DATA TYPE]" & _
    ",[AUTO ADD],[DEFAULT VALUE],[CALCULATE CUSTOMS VALUE],[RELATE L1 TO S1],[VALIDATE_VALUE] FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

    strSqlDeleteRecord = "DELETE * FROM [" & mvarTableName & "] WHERE [" & PK_FIELD & "] = "

    strSqlGetMaxID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & mvarTableName & "]"

    strSqlGetMinID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & mvarTableName & "]"

End Function


' :-) :-) :-) :-) :-) :-) :-) :-) :-) :-) :-)
' :-)
' :-)

