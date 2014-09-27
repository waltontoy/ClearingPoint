VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiLOGICAL_IDs"
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

Private Const TABLE_NAME = "LOGICAL ID"
Private Const PK_FIELD = "LOGID DESCRIPTION"

' field constants
Private Const FIELD_LOGID_DESCRIPTION = "LOGID DESCRIPTION" '1
Private Const FIELD_SEND_OPERATIONAL_CORR = "SEND OPERATIONAL CORR" '2
Private Const FIELD_SEND_OPERATIONAL_LOGID = "SEND OPERATIONAL LOGID" '3
Private Const FIELD_SEND_OPERATIONAL_PASS = "SEND OPERATIONAL PASS" '4
Private Const FIELD_SEND_TEST_CORR = "SEND TEST CORR" '5
Private Const FIELD_SEND_TEST_LOGID = "SEND TEST LOGID" '6
Private Const FIELD_SEND_TEST_PASS = "SEND TEST PASS" '7
Private Const FIELD_PRINT_OPERATIONAL_CORR = "PRINT OPERATIONAL CORR" '8
Private Const FIELD_PRINT_OPERATIONAL_LOGID = "PRINT OPERATIONAL LOGID" '9
Private Const FIELD_PRINT_OPERATIONAL_PASS = "PRINT OPERATIONAL PASS" '10
Private Const FIELD_PRINT_TEST_CORR = "PRINT TEST CORR" '11
Private Const FIELD_PRINT_TEST_LOGID = "PRINT TEST LOGID" '12
Private Const FIELD_PRINT_TEST_PASS = "PRINT TEST PASS" '13
Private Const FIELD_A1 = "A1" '14
Private Const FIELD_A2 = "A2" '15
Private Const FIELD_HISTORY = "HISTORY" '16
Private Const FIELD_VAT = "VAT" '17
Private Const FIELD_TIN = "TIN" '18
Private Const FIELD_USAGE = "USAGE" '19
Private Const FIELD_LRN_USAGE = "LRN USAGE" '20
Private Const FIELD_BRANCH = "BRANCH" '21
Private Const FIELD_PROCEDURE = "PROCEDURE" '22
Private Const FIELD_LOCATION = "LOCATION" '23
Private Const FIELD_VOLGBRIEFJE_PRINTER = "VOLGBRIEFJE PRINTER" '24
Private Const FIELD_NCTS_DOC_PRINTER = "NCTS DOC PRINTER" '25
Private Const FIELD_PRINT_MODE = "PRINT MODE" '26
Private Const FIELD_SFTP_ADDRESS = "SFTP ADDRESS" '27
Private Const FIELD_EMAIL_ADDRESS = "EMAIL ADDRESS" '28
Private Const FIELD_SEND_EDI_RECIPIENT_OPERATIONAL = "SEND EDI RECIPIENT OPERATIONAL" '29
Private Const FIELD_SEND_EDI_RECIPIENT_TEST = "SEND EDI RECIPIENT TEST" '30
Private Const FIELD_LAST_EDI_REFERENCE = "LAST EDI REFERENCE" '31

Private Const SQL_ADD_RECORD = "INSERT INTO [" & TABLE_NAME & "] (" & _
"[LOGID DESCRIPTION],[SEND OPERATIONAL CORR],[SEND OPERATIONAL LOGID]" & _
",[SEND OPERATIONAL PASS],[SEND TEST CORR],[SEND TEST LOGID]" & _
",[SEND TEST PASS],[PRINT OPERATIONAL CORR],[PRINT OPERATIONAL LOGID]" & _
",[PRINT OPERATIONAL PASS],[PRINT TEST CORR],[PRINT TEST LOGID]" & _
",[PRINT TEST PASS],[A1],[A2]" & _
",[HISTORY],[VAT],[TIN]" & _
",[USAGE],[LRN USAGE],[BRANCH]" & _
",[PROCEDURE],[LOCATION],[VOLGBRIEFJE PRINTER]" & _
",[NCTS DOC PRINTER],[PRINT MODE],[SFTP ADDRESS]" & _
",[EMAIL ADDRESS],[SEND EDI RECIPIENT OPERATIONAL],[SEND EDI RECIPIENT TEST]" & _
",[LAST EDI REFERENCE]                                     )  VALUES "

Private Const SQL_GET_RECORD = "SELECT [LOGID DESCRIPTION],[SEND OPERATIONAL CORR],[SEND OPERATIONAL LOGID]" & _
",[SEND OPERATIONAL PASS],[SEND TEST CORR],[SEND TEST LOGID]" & _
",[SEND TEST PASS],[PRINT OPERATIONAL CORR],[PRINT OPERATIONAL LOGID]" & _
",[PRINT OPERATIONAL PASS],[PRINT TEST CORR],[PRINT TEST LOGID]" & _
",[PRINT TEST PASS],[A1],[A2]" & _
",[HISTORY],[VAT],[TIN]" & _
",[USAGE],[LRN USAGE],[BRANCH]" & _
",[PROCEDURE],[LOCATION],[VOLGBRIEFJE PRINTER]" & _
",[NCTS DOC PRINTER],[PRINT MODE],[SFTP ADDRESS]" & _
",[EMAIL ADDRESS],[SEND EDI RECIPIENT OPERATIONAL],[SEND EDI RECIPIENT TEST]" & _
",[LAST EDI REFERENCE] FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_DELETE_RECORD = "DELETE * FROM [" & TABLE_NAME & "] WHERE [" & PK_FIELD & "] = "

Private Const SQL_GET_MAXID = "SELECT MAX([" & PK_FIELD & "]) AS [ID_MAX] FROM [" & TABLE_NAME & "]"

Private Const SQL_GET_MINID = "SELECT Min([" & PK_FIELD & "]) AS [ID_MIN] FROM [" & TABLE_NAME & "]"



' /* --------------- BASIC FUNCTIONS -------------------- */

Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_ADD_RECORD & "("
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_CORR) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_PASS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_CORR) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_PASS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_CORR) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_PASS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_CORR) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_LOGID) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_PASS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_A1) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_A2) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_HISTORY) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_VAT) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_TIN) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_USAGE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LRN_USAGE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_BRANCH) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PROCEDURE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LOCATION) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_VOLGBRIEFJE_PRINTER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_NCTS_DOC_PRINTER) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_PRINT_MODE) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SFTP_ADDRESS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_EMAIL_ADDRESS) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_EDI_RECIPIENT_OPERATIONAL) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_SEND_EDI_RECIPIENT_TEST) & ","
      strSql = strSql & SetEnclosedChr(ActiveRecord.FIELD_LAST_EDI_REFERENCE) & ")"

      On Error GoTo ERROR_QUERY
    
      ExecuteNonQuery ActiveConnection, strSql
      'ActiveConnection.Execute strSql


      AddRecord = True

      Exit Function

  End If


  AddRecord = False


  Exit Function


  AddRecord = False

ERROR_QUERY:

  AddRecord = False

End Function


Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_DELETE_RECORD & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION)


      On Error GoTo ERROR_QUERY
      ExecuteNonQuery ActiveConnection, strSql
      'ActiveConnection.Execute strSql


      DeleteRecord = True

      Exit Function


   End If


   DeleteRecord = False


   Exit Function


ERROR_QUERY:

   DeleteRecord = False

End Function




Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

   Dim strSql As String


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = strSql & "UPDATE [" & TABLE_NAME & "] SET "
      strSql = strSql & " [SEND OPERATIONAL CORR]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_CORR) & ", "
      strSql = strSql & " [SEND OPERATIONAL LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_LOGID) & ", "
      strSql = strSql & " [SEND OPERATIONAL PASS]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_OPERATIONAL_PASS) & ", "
      strSql = strSql & " [SEND TEST CORR]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_CORR) & ", "
      strSql = strSql & " [SEND TEST LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_LOGID) & ", "
      strSql = strSql & " [SEND TEST PASS]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_TEST_PASS) & ", "
      strSql = strSql & " [PRINT OPERATIONAL CORR]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_CORR) & ", "
      strSql = strSql & " [PRINT OPERATIONAL LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_LOGID) & ", "
      strSql = strSql & " [PRINT OPERATIONAL PASS]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_OPERATIONAL_PASS) & ", "
      strSql = strSql & " [PRINT TEST CORR]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_CORR) & ", "
      strSql = strSql & " [PRINT TEST LOGID]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_LOGID) & ", "
      strSql = strSql & " [PRINT TEST PASS]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_TEST_PASS) & ", "
      strSql = strSql & " [A1]=" & SetEnclosedChr(ActiveRecord.FIELD_A1) & ", "
      strSql = strSql & " [A2]=" & SetEnclosedChr(ActiveRecord.FIELD_A2) & ", "
      strSql = strSql & " [HISTORY]=" & SetEnclosedChr(ActiveRecord.FIELD_HISTORY) & ", "
      strSql = strSql & " [VAT]=" & SetEnclosedChr(ActiveRecord.FIELD_VAT) & ", "
      strSql = strSql & " [TIN]=" & SetEnclosedChr(ActiveRecord.FIELD_TIN) & ", "
      strSql = strSql & " [USAGE]=" & SetEnclosedChr(ActiveRecord.FIELD_USAGE) & ", "
      strSql = strSql & " [LRN USAGE]=" & SetEnclosedChr(ActiveRecord.FIELD_LRN_USAGE) & ", "
      strSql = strSql & " [BRANCH]=" & SetEnclosedChr(ActiveRecord.FIELD_BRANCH) & ", "
      strSql = strSql & " [PROCEDURE]=" & SetEnclosedChr(ActiveRecord.FIELD_PROCEDURE) & ", "
      strSql = strSql & " [LOCATION]=" & SetEnclosedChr(ActiveRecord.FIELD_LOCATION) & ", "
      strSql = strSql & " [VOLGBRIEFJE PRINTER]=" & SetEnclosedChr(ActiveRecord.FIELD_VOLGBRIEFJE_PRINTER) & ", "
      strSql = strSql & " [NCTS DOC PRINTER]=" & SetEnclosedChr(ActiveRecord.FIELD_NCTS_DOC_PRINTER) & ", "
      strSql = strSql & " [PRINT MODE]=" & SetEnclosedChr(ActiveRecord.FIELD_PRINT_MODE) & ", "
      strSql = strSql & " [SFTP ADDRESS]=" & SetEnclosedChr(ActiveRecord.FIELD_SFTP_ADDRESS) & ", "
      strSql = strSql & " [EMAIL ADDRESS]=" & SetEnclosedChr(ActiveRecord.FIELD_EMAIL_ADDRESS) & ", "
      strSql = strSql & " [SEND EDI RECIPIENT OPERATIONAL]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_EDI_RECIPIENT_OPERATIONAL) & ", "
      strSql = strSql & " [SEND EDI RECIPIENT TEST]=" & SetEnclosedChr(ActiveRecord.FIELD_SEND_EDI_RECIPIENT_TEST) & ", "
      strSql = strSql & " [LAST EDI REFERENCE]=" & SetEnclosedChr(ActiveRecord.FIELD_LAST_EDI_REFERENCE) & ""
      strSql = strSql & " WHERE [" & PK_FIELD & "] = " & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION)

      On Error GoTo ERROR_QUERY
      ExecuteNonQuery ActiveConnection, strSql
      'ActiveConnection.Execute strSql


      ModifyRecord = True

      Exit Function


   End If


   ModifyRecord = False


   Exit Function


ERROR_QUERY:

   ModifyRecord = False

End Function




Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

   Dim strSql As String


   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = False) Then

      strSql = SQL_GET_RECORD & SetEnclosedChr(ActiveRecord.FIELD_LOGID_DESCRIPTION)

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_LOGID_DESCRIPTION = FNullField(rstRecord.Fields("LOGID DESCRIPTION"))
      ActiveRecord.FIELD_SEND_OPERATIONAL_CORR = FNullField(rstRecord.Fields("SEND OPERATIONAL CORR"))
      ActiveRecord.FIELD_SEND_OPERATIONAL_LOGID = FNullField(rstRecord.Fields("SEND OPERATIONAL LOGID"))
      ActiveRecord.FIELD_SEND_OPERATIONAL_PASS = FNullField(rstRecord.Fields("SEND OPERATIONAL PASS"))
      ActiveRecord.FIELD_SEND_TEST_CORR = FNullField(rstRecord.Fields("SEND TEST CORR"))
      ActiveRecord.FIELD_SEND_TEST_LOGID = FNullField(rstRecord.Fields("SEND TEST LOGID"))
      ActiveRecord.FIELD_SEND_TEST_PASS = FNullField(rstRecord.Fields("SEND TEST PASS"))
      ActiveRecord.FIELD_PRINT_OPERATIONAL_CORR = FNullField(rstRecord.Fields("PRINT OPERATIONAL CORR"))
      ActiveRecord.FIELD_PRINT_OPERATIONAL_LOGID = FNullField(rstRecord.Fields("PRINT OPERATIONAL LOGID"))
      ActiveRecord.FIELD_PRINT_OPERATIONAL_PASS = FNullField(rstRecord.Fields("PRINT OPERATIONAL PASS"))
      ActiveRecord.FIELD_PRINT_TEST_CORR = FNullField(rstRecord.Fields("PRINT TEST CORR"))
      ActiveRecord.FIELD_PRINT_TEST_LOGID = FNullField(rstRecord.Fields("PRINT TEST LOGID"))
      ActiveRecord.FIELD_PRINT_TEST_PASS = FNullField(rstRecord.Fields("PRINT TEST PASS"))
      ActiveRecord.FIELD_A1 = FNullField(rstRecord.Fields("A1"))
      ActiveRecord.FIELD_A2 = FNullField(rstRecord.Fields("A2"))
      ActiveRecord.FIELD_HISTORY = FNullField(rstRecord.Fields("HISTORY"))
      ActiveRecord.FIELD_VAT = FNullField(rstRecord.Fields("VAT"))
      ActiveRecord.FIELD_TIN = FNullField(rstRecord.Fields("TIN"))
      ActiveRecord.FIELD_USAGE = FNullField(rstRecord.Fields("USAGE"))
      ActiveRecord.FIELD_LRN_USAGE = FNullField(rstRecord.Fields("LRN USAGE"))
      ActiveRecord.FIELD_BRANCH = FNullField(rstRecord.Fields("BRANCH"))
      ActiveRecord.FIELD_PROCEDURE = FNullField(rstRecord.Fields("PROCEDURE"))
      ActiveRecord.FIELD_LOCATION = FNullField(rstRecord.Fields("LOCATION"))
      ActiveRecord.FIELD_VOLGBRIEFJE_PRINTER = FNullField(rstRecord.Fields("VOLGBRIEFJE PRINTER"))
      ActiveRecord.FIELD_NCTS_DOC_PRINTER = FNullField(rstRecord.Fields("NCTS DOC PRINTER"))
      ActiveRecord.FIELD_PRINT_MODE = FNullField(rstRecord.Fields("PRINT MODE"))
      ActiveRecord.FIELD_SFTP_ADDRESS = FNullField(rstRecord.Fields("SFTP ADDRESS"))
      ActiveRecord.FIELD_EMAIL_ADDRESS = FNullField(rstRecord.Fields("EMAIL ADDRESS"))
      ActiveRecord.FIELD_SEND_EDI_RECIPIENT_OPERATIONAL = FNullField(rstRecord.Fields("SEND EDI RECIPIENT OPERATIONAL"))
      ActiveRecord.FIELD_SEND_EDI_RECIPIENT_TEST = FNullField(rstRecord.Fields("SEND EDI RECIPIENT TEST"))
      ActiveRecord.FIELD_LAST_EDI_REFERENCE = FNullField(rstRecord.Fields("LAST EDI REFERENCE"))
      
      ADORecordsetClose rstRecord
      'Set rstRecord = Nothing

      GetRecord = True

      Exit Function


   End If


   GetRecord = False


   Exit Function


ERROR_RECORDSET:
       ADORecordsetClose rstRecord
       'Set rstRecord = Nothing
       
       GetRecord = False
       Exit Function
ERROR_QUERY:

   GetRecord = False

End Function




Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiLOGICAL_ID

   End If

      strSql = SQL_GET_MAXID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_LOGID_DESCRIPTION = rstRecord.Fields("ID_MAX").Value
      
      ADORecordsetClose rstRecord
      'Set rstRecord = Nothing

      GetMaxID = True

      Exit Function


ERROR_RECORDSET:
   ADORecordsetClose rstRecord
   'Set rstRecord = Nothing

   GetMaxID = False

   Exit Function


ERROR_QUERY:

   GetMaxID = False

End Function




Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLOGICAL_ID) As Boolean

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   If ((ActiveRecord Is Nothing) = True) Then

      Set ActiveRecord = New cpiLOGICAL_ID

   End If

      strSql = SQL_GET_MINID

      On Error GoTo ERROR_QUERY
      ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
      'Set rstRecord = ActiveConnection.Execute(strSql)

      On Error GoTo ERROR_RECORDSET
      ActiveRecord.FIELD_LOGID_DESCRIPTION = rstRecord.Fields("ID_MIN").Value
      
      ADORecordsetClose rstRecord
      'Set rstRecord = Nothing

      GetMinID = True

      Exit Function


ERROR_RECORDSET:
   ADORecordsetClose rstRecord
   'Set rstRecord = Nothing

   GetMinID = False

   Exit Function


ERROR_QUERY:

   GetMinID = False

End Function




Public Function GetTableRecord(ByRef ActiveRecord As cpiLOGICAL_ID) As ADODB.Recordset

   Dim strSql As String
   Dim rstRecord As ADODB.Recordset


   Set rstRecord = New ADODB.Recordset

   rstRecord.Open


   rstRecord.Fields.Append FIELD_LOGID_DESCRIPTION, 202, 40, 70 ' 1
   rstRecord.Fields.Append FIELD_SEND_OPERATIONAL_CORR, 202, 6, 102 ' 2
   rstRecord.Fields.Append FIELD_SEND_OPERATIONAL_LOGID, 202, 4, 102 ' 3
   rstRecord.Fields.Append FIELD_SEND_OPERATIONAL_PASS, 202, 5, 102 ' 4
   rstRecord.Fields.Append FIELD_SEND_TEST_CORR, 202, 6, 102 ' 5
   rstRecord.Fields.Append FIELD_SEND_TEST_LOGID, 202, 4, 102 ' 6
   rstRecord.Fields.Append FIELD_SEND_TEST_PASS, 202, 5, 102 ' 7
   rstRecord.Fields.Append FIELD_PRINT_OPERATIONAL_CORR, 202, 6, 102 ' 8
   rstRecord.Fields.Append FIELD_PRINT_OPERATIONAL_LOGID, 202, 4, 102 ' 9
   rstRecord.Fields.Append FIELD_PRINT_OPERATIONAL_PASS, 202, 5, 102 ' 10
   rstRecord.Fields.Append FIELD_PRINT_TEST_CORR, 202, 6, 102 ' 11
   rstRecord.Fields.Append FIELD_PRINT_TEST_LOGID, 202, 4, 102 ' 12
   rstRecord.Fields.Append FIELD_PRINT_TEST_PASS, 202, 5, 102 ' 13
   rstRecord.Fields.Append FIELD_A1, 202, 4, 102 ' 14
   rstRecord.Fields.Append FIELD_A2, 202, 5, 102 ' 15
   rstRecord.Fields.Append FIELD_HISTORY, 202, 4, 102 ' 16
   rstRecord.Fields.Append FIELD_VAT, 202, 9, 70 ' 17
   rstRecord.Fields.Append FIELD_TIN, 202, 17, 70 ' 18
   rstRecord.Fields.Append FIELD_USAGE, 17, 1, 86 ' 19
   rstRecord.Fields.Append FIELD_LRN_USAGE, 17, 1, 86 ' 20
   rstRecord.Fields.Append FIELD_BRANCH, 202, 3, 102 ' 21
   rstRecord.Fields.Append FIELD_PROCEDURE, 17, 1, 86 ' 22
   rstRecord.Fields.Append FIELD_LOCATION, 202, 17, 102 ' 23
   rstRecord.Fields.Append FIELD_VOLGBRIEFJE_PRINTER, 202, 50, 102 ' 24
   rstRecord.Fields.Append FIELD_NCTS_DOC_PRINTER, 202, 50, 102 ' 25
   rstRecord.Fields.Append FIELD_PRINT_MODE, 17, 1, 86 ' 26
   rstRecord.Fields.Append FIELD_SFTP_ADDRESS, 202, 50, 102 ' 27
   rstRecord.Fields.Append FIELD_EMAIL_ADDRESS, 202, 50, 102 ' 28
   rstRecord.Fields.Append FIELD_SEND_EDI_RECIPIENT_OPERATIONAL, 202, 35, 102 ' 29
   rstRecord.Fields.Append FIELD_SEND_EDI_RECIPIENT_TEST, 202, 35, 102 ' 30
   rstRecord.Fields.Append FIELD_LAST_EDI_REFERENCE, 202, 5, 102 ' 31

   ' set values

   rstRecord.AddNew
   rstRecord.Fields("LOGID DESCRIPTION") = ActiveRecord.FIELD_LOGID_DESCRIPTION
   rstRecord.Fields("SEND OPERATIONAL CORR") = ActiveRecord.FIELD_SEND_OPERATIONAL_CORR
   rstRecord.Fields("SEND OPERATIONAL LOGID") = ActiveRecord.FIELD_SEND_OPERATIONAL_LOGID
   rstRecord.Fields("SEND OPERATIONAL PASS") = ActiveRecord.FIELD_SEND_OPERATIONAL_PASS
   rstRecord.Fields("SEND TEST CORR") = ActiveRecord.FIELD_SEND_TEST_CORR
   rstRecord.Fields("SEND TEST LOGID") = ActiveRecord.FIELD_SEND_TEST_LOGID
   rstRecord.Fields("SEND TEST PASS") = ActiveRecord.FIELD_SEND_TEST_PASS
   rstRecord.Fields("PRINT OPERATIONAL CORR") = ActiveRecord.FIELD_PRINT_OPERATIONAL_CORR
   rstRecord.Fields("PRINT OPERATIONAL LOGID") = ActiveRecord.FIELD_PRINT_OPERATIONAL_LOGID
   rstRecord.Fields("PRINT OPERATIONAL PASS") = ActiveRecord.FIELD_PRINT_OPERATIONAL_PASS
   rstRecord.Fields("PRINT TEST CORR") = ActiveRecord.FIELD_PRINT_TEST_CORR
   rstRecord.Fields("PRINT TEST LOGID") = ActiveRecord.FIELD_PRINT_TEST_LOGID
   rstRecord.Fields("PRINT TEST PASS") = ActiveRecord.FIELD_PRINT_TEST_PASS
   rstRecord.Fields("A1") = ActiveRecord.FIELD_A1
   rstRecord.Fields("A2") = ActiveRecord.FIELD_A2
   rstRecord.Fields("HISTORY") = ActiveRecord.FIELD_HISTORY
   rstRecord.Fields("VAT") = ActiveRecord.FIELD_VAT
   rstRecord.Fields("TIN") = ActiveRecord.FIELD_TIN
   rstRecord.Fields("USAGE") = ActiveRecord.FIELD_USAGE
   rstRecord.Fields("LRN USAGE") = ActiveRecord.FIELD_LRN_USAGE
   rstRecord.Fields("BRANCH") = ActiveRecord.FIELD_BRANCH
   rstRecord.Fields("PROCEDURE") = ActiveRecord.FIELD_PROCEDURE
   rstRecord.Fields("LOCATION") = ActiveRecord.FIELD_LOCATION
   rstRecord.Fields("VOLGBRIEFJE PRINTER") = ActiveRecord.FIELD_VOLGBRIEFJE_PRINTER
   rstRecord.Fields("NCTS DOC PRINTER") = ActiveRecord.FIELD_NCTS_DOC_PRINTER
   rstRecord.Fields("PRINT MODE") = ActiveRecord.FIELD_PRINT_MODE
   rstRecord.Fields("SFTP ADDRESS") = ActiveRecord.FIELD_SFTP_ADDRESS
   rstRecord.Fields("EMAIL ADDRESS") = ActiveRecord.FIELD_EMAIL_ADDRESS
   rstRecord.Fields("SEND EDI RECIPIENT OPERATIONAL") = ActiveRecord.FIELD_SEND_EDI_RECIPIENT_OPERATIONAL
   rstRecord.Fields("SEND EDI RECIPIENT TEST") = ActiveRecord.FIELD_SEND_EDI_RECIPIENT_TEST
   rstRecord.Fields("LAST EDI REFERENCE") = ActiveRecord.FIELD_LAST_EDI_REFERENCE

   rstRecord.Update
   Set GetTableRecord = rstRecord
   Set rstRecord = Nothing

End Function


Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiLOGICAL_ID

   Dim clsRecord As cpiLOGICAL_ID
   Set clsRecord = New cpiLOGICAL_ID

   clsRecord.FIELD_LOGID_DESCRIPTION = FNullField(ActiveRecord.Fields(FIELD_LOGID_DESCRIPTION))
   clsRecord.FIELD_SEND_OPERATIONAL_CORR = FNullField(ActiveRecord.Fields(FIELD_SEND_OPERATIONAL_CORR))
   clsRecord.FIELD_SEND_OPERATIONAL_LOGID = FNullField(ActiveRecord.Fields(FIELD_SEND_OPERATIONAL_LOGID))
   clsRecord.FIELD_SEND_OPERATIONAL_PASS = FNullField(ActiveRecord.Fields(FIELD_SEND_OPERATIONAL_PASS))
   clsRecord.FIELD_SEND_TEST_CORR = FNullField(ActiveRecord.Fields(FIELD_SEND_TEST_CORR))
   clsRecord.FIELD_SEND_TEST_LOGID = FNullField(ActiveRecord.Fields(FIELD_SEND_TEST_LOGID))
   clsRecord.FIELD_SEND_TEST_PASS = FNullField(ActiveRecord.Fields(FIELD_SEND_TEST_PASS))
   clsRecord.FIELD_PRINT_OPERATIONAL_CORR = FNullField(ActiveRecord.Fields(FIELD_PRINT_OPERATIONAL_CORR))
   clsRecord.FIELD_PRINT_OPERATIONAL_LOGID = FNullField(ActiveRecord.Fields(FIELD_PRINT_OPERATIONAL_LOGID))
   clsRecord.FIELD_PRINT_OPERATIONAL_PASS = FNullField(ActiveRecord.Fields(FIELD_PRINT_OPERATIONAL_PASS))
   clsRecord.FIELD_PRINT_TEST_CORR = FNullField(ActiveRecord.Fields(FIELD_PRINT_TEST_CORR))
   clsRecord.FIELD_PRINT_TEST_LOGID = FNullField(ActiveRecord.Fields(FIELD_PRINT_TEST_LOGID))
   clsRecord.FIELD_PRINT_TEST_PASS = FNullField(ActiveRecord.Fields(FIELD_PRINT_TEST_PASS))
   clsRecord.FIELD_A1 = FNullField(ActiveRecord.Fields(FIELD_A1))
   clsRecord.FIELD_A2 = FNullField(ActiveRecord.Fields(FIELD_A2))
   clsRecord.FIELD_HISTORY = FNullField(ActiveRecord.Fields(FIELD_HISTORY))
   clsRecord.FIELD_VAT = FNullField(ActiveRecord.Fields(FIELD_VAT))
   clsRecord.FIELD_TIN = FNullField(ActiveRecord.Fields(FIELD_TIN))
   clsRecord.FIELD_USAGE = FNullField(ActiveRecord.Fields(FIELD_USAGE))
   clsRecord.FIELD_LRN_USAGE = FNullField(ActiveRecord.Fields(FIELD_LRN_USAGE))
   clsRecord.FIELD_BRANCH = FNullField(ActiveRecord.Fields(FIELD_BRANCH))
   clsRecord.FIELD_PROCEDURE = FNullField(ActiveRecord.Fields(FIELD_PROCEDURE))
   clsRecord.FIELD_LOCATION = FNullField(ActiveRecord.Fields(FIELD_LOCATION))
   clsRecord.FIELD_VOLGBRIEFJE_PRINTER = FNullField(ActiveRecord.Fields(FIELD_VOLGBRIEFJE_PRINTER))
   clsRecord.FIELD_NCTS_DOC_PRINTER = FNullField(ActiveRecord.Fields(FIELD_NCTS_DOC_PRINTER))
   clsRecord.FIELD_PRINT_MODE = FNullField(ActiveRecord.Fields(FIELD_PRINT_MODE))
   clsRecord.FIELD_SFTP_ADDRESS = FNullField(ActiveRecord.Fields(FIELD_SFTP_ADDRESS))
   clsRecord.FIELD_EMAIL_ADDRESS = FNullField(ActiveRecord.Fields(FIELD_EMAIL_ADDRESS))
   clsRecord.FIELD_SEND_EDI_RECIPIENT_OPERATIONAL = FNullField(ActiveRecord.Fields(FIELD_SEND_EDI_RECIPIENT_OPERATIONAL))
   clsRecord.FIELD_SEND_EDI_RECIPIENT_TEST = FNullField(ActiveRecord.Fields(FIELD_SEND_EDI_RECIPIENT_TEST))
   clsRecord.FIELD_LAST_EDI_REFERENCE = FNullField(ActiveRecord.Fields(FIELD_LAST_EDI_REFERENCE))
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
      'Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)

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
   'rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

   Set GetRecordset = rstRecordset
   Set mvarRecordset = rstRecordset
   Set rstRecordset = Nothing

   Exit Function

ERROR_RECORDSET:
   Set rstRecordset = Nothing
End Function


Public Function Add(ByRef sKey As String, ByVal PKValue As String) As cpiLOGICAL_ID

    'create a new object
    Dim objNewMember As cpiLOGICAL_ID
    Set objNewMember = New cpiLOGICAL_ID

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

Public Property Get Item(vntIndexKey As Variant) As cpiLOGICAL_ID
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




