VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiLicensees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Private Const TABLE_NAME = "Licensee"
Private Const PK_FIELD = "Lic_ID"
    
' field constants
Private Const FIELD_Lic_ID = "Lic_ID"
Private Const FIELD_Lic_Name = "Lic_Name"
Private Const FIELD_Lic_Address = "Lic_Address"
Private Const FIELD_Lic_City = "Lic_City"
Private Const FIELD_Lic_PostalCode = "Lic_PostalCode"
Private Const FIELD_Lic_Country = "Lic_Country"
Private Const FIELD_Lic_Phone = "Lic_Phone"
Private Const FIELD_Lic_Fax = "Lic_Fax"
Private Const FIELD_Lic_Email = "Lic_Email"
Private Const FIELD_Lic_LegalInfo = "Lic_LegalInfo"
Private Const FIELD_Lic_Currency = "Lic_Currency"
Private Const FIELD_Lic_Database = "Lic_Database"
Private Const FIELD_Lic_UseEXACTID = "Lic_UseEXACTID"
Private Const FIELD_Lic_Language = "Lic_Language"
Private Const FIELD_Lic_Logo = "Lic_Logo"
Private Const FIELD_Lic_Logosize = "Lic_Logosize"
Private Const FIELD_Lic_LogoProperties = "Lic_LogoProperties"
Private Const FIELD_Lic_Website = "Lic_Website"
Private Const FIELD_Lic_Key = "Lic_Key"
Private Const FIELD_Lic_User = "Lic_User"

Private Const SQL_ADD_RECORD = "INSERT INTO [Licensee] ([Lic_ID], [Lic_Name], [Lic_Address], [Lic_City], " & _
                            "[Lic_PostalCode], [Lic_Country], [Lic_Phone], [Lic_Fax], [Lic_Email], " & _
                            "[Lic_LegalInfo], [Lic_Currency], [Lic_Database], [Lic_UseEXACTID], [Lic_Language], " & _
                            "[Lic_Logosize], [Lic_LogoProperties], [Lic_Website], [Lic_Key], " & _
                            "[Lic_User]) VALUES "

Private Const SQL_DELETE_RECORD = "DELETE * FROM Licensee WHERE [Lic_ID]="

'Private Const SQL_MODIFY_RECORD '=

Private Const SQL_GET_RECORD = "SELECT [Lic_ID], [Lic_Name], [Lic_Address], [Lic_City], " & _
                            "[Lic_PostalCode], [Lic_Country], [Lic_Phone], [Lic_Fax], [Lic_Email], " & _
                            "[Lic_LegalInfo], [Lic_Currency], [Lic_Database], [Lic_UseEXACTID], [Lic_Language], " & _
                            "[Lic_Logo], [Lic_Logosize], [Lic_LogoProperties], [Lic_Website], [Lic_Key], " & _
                            "[Lic_User] FROM Licensee WHERE [Lic_ID]= "
Private Const SQL_GET_MAXID = "SELECT MAX(Lic_ID) AS [ID_MAX] FROM Licensee"
Private Const SQL_GET_MINID = "SELECT Min(Lic_ID) AS [ID_MIN] FROM Licensee"

' /* --------------- BASIC FUNCTIONS -------------------- */
Public Function AddRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_ADD_RECORD & _
                        "(" & ActiveRecord.Lic_ID & ",'" & ActiveRecord.Lic_Name & "','" & ActiveRecord.Lic_Address & _
                        "','" & ActiveRecord.Lic_City & "','" & ActiveRecord.Lic_PostalCode & _
                        "','" & ActiveRecord.Lic_Country & _
                        "','" & ActiveRecord.Lic_Phone & "','" & ActiveRecord.Lic_Fax & "','" & ActiveRecord.Lic_Email & _
                        "','" & ActiveRecord.Lic_LegalInfo & "','" & ActiveRecord.Lic_Currency & "','" & ActiveRecord.Lic_Database & _
                        "'," & ActiveRecord.Lic_UseEXACTID & ",'" & ActiveRecord.Lic_Language & "','" & ActiveRecord.Lic_Logosize & _
                        "','" & ActiveRecord.Lic_LogoProperties & "','" & ActiveRecord.Lic_Website & "','" & ActiveRecord.Lic_Key & _
                        "','" & ActiveRecord.Lic_User & "')"
    
        On Error GoTo ERROR_QUERY
        ExecuteNonQuery ActiveConnection, strSql
        'ActiveConnection.Execute strSql
    
        ' add procedures to add pictures here
        ' -> for Lic_Logo
    
        AddRecord = True
        Exit Function
        
    End If
    
    AddRecord = False
    
    Exit Function
    
ERROR_QUERY:
    AddRecord = False
End Function

Public Function DeleteRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean
    
    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_DELETE_RECORD & ActiveRecord.Lic_ID
    
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

' UPDATE [Customers]  SET [Customers].[CompanyName]='cOMPANY', [Customers].[ContactName]='jason' WHERE [Customers].[CustomerID]='BONAP'
Public Function ModifyRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean

    Dim strSql As String

    If ((ActiveRecord Is Nothing) = False) Then

        strSql = "UPDATE " & TABLE_NAME & " SET " & _
                            FIELD_Lic_Name & "='" & ActiveRecord.Lic_Name & "'," & _
                            FIELD_Lic_Address & "='" & ActiveRecord.Lic_Address & "'," & _
                            FIELD_Lic_City & "='" & ActiveRecord.Lic_City & "'," & _
                            FIELD_Lic_PostalCode & "='" & ActiveRecord.Lic_PostalCode & "'," & _
                            FIELD_Lic_Country & "='" & ActiveRecord.Lic_Country & "'," & _
                            FIELD_Lic_Phone & "='" & ActiveRecord.Lic_Phone & "'," & _
                            FIELD_Lic_Fax & "='" & ActiveRecord.Lic_Fax & "'," & _
                            FIELD_Lic_Email & "='" & ActiveRecord.Lic_Email & "'," & _
                            FIELD_Lic_LegalInfo & "='" & ActiveRecord.Lic_LegalInfo & "'," & _
                            FIELD_Lic_Currency & "='" & ActiveRecord.Lic_Currency & "'," & _
                            FIELD_Lic_Database & "='" & ActiveRecord.Lic_Database & "'," & _
                            FIELD_Lic_UseEXACTID & "=" & ActiveRecord.Lic_UseEXACTID & "," & _
                            FIELD_Lic_Language & "='" & ActiveRecord.Lic_Language & "'," & _
                            FIELD_Lic_Logosize & "='" & ActiveRecord.Lic_Logosize & "'," & _
                            FIELD_Lic_LogoProperties & "='" & ActiveRecord.Lic_LogoProperties & "'," & _
                            FIELD_Lic_Website & "='" & ActiveRecord.Lic_Website & "'," & _
                            FIELD_Lic_Key & "='" & ActiveRecord.Lic_Key & "'," & _
                            FIELD_Lic_User & "='" & ActiveRecord.Lic_User & _
                        "' WHERE " & PK_FIELD & "=" & ActiveRecord.Lic_ID
    
        On Error GoTo ERROR_QUERY
        ExecuteNonQuery ActiveConnection, strSql
        'ActiveConnection.Execute strSql
    
        ' add procedures here to modicy Lic_logo pics
    
        ModifyRecord = True
        Exit Function
        
    End If
    
    ModifyRecord = False
    
    Exit Function
    
ERROR_QUERY:
    ModifyRecord = False
End Function

Public Function GetRecord(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset
    
    If ((ActiveRecord Is Nothing) = False) Then

        strSql = SQL_GET_RECORD & ActiveRecord.Lic_ID
        
        On Error GoTo ERROR_QUERY
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        'Set rstRecord = ActiveConnection.Execute(strSql)
        
        On Error GoTo ERROR_RECORDSET
        ActiveRecord.Lic_Name = FNullField(rstRecord.Fields(FIELD_Lic_Name))
        ActiveRecord.Lic_Address = FNullField(rstRecord.Fields(FIELD_Lic_Address))
        ActiveRecord.Lic_City = FNullField(rstRecord.Fields(FIELD_Lic_City))
        ActiveRecord.Lic_PostalCode = FNullField(rstRecord.Fields(FIELD_Lic_PostalCode))
        ActiveRecord.Lic_Country = FNullField(rstRecord.Fields(FIELD_Lic_Country))
        ActiveRecord.Lic_Phone = FNullField(rstRecord.Fields(FIELD_Lic_Phone))
        ActiveRecord.Lic_Fax = FNullField(rstRecord.Fields(FIELD_Lic_Fax))
        ActiveRecord.Lic_Email = FNullField(rstRecord.Fields(FIELD_Lic_Email))
        ActiveRecord.Lic_LegalInfo = FNullField(rstRecord.Fields(FIELD_Lic_LegalInfo))
        ActiveRecord.Lic_Currency = FNullField(rstRecord.Fields(FIELD_Lic_Currency))
        ActiveRecord.Lic_Database = FNullField(rstRecord.Fields(FIELD_Lic_Database))
        ActiveRecord.Lic_UseEXACTID = FNullField(rstRecord.Fields(FIELD_Lic_UseEXACTID))
        ActiveRecord.Lic_Language = FNullField(rstRecord.Fields(FIELD_Lic_Language))
        ActiveRecord.Lic_Logosize = FNullField(rstRecord.Fields(FIELD_Lic_Logosize))
        ActiveRecord.Lic_LogoProperties = FNullField(rstRecord.Fields(FIELD_Lic_LogoProperties))
        ActiveRecord.Lic_Website = FNullField(rstRecord.Fields(FIELD_Lic_Website))
        ActiveRecord.Lic_Key = FNullField(rstRecord.Fields(FIELD_Lic_Key))
        ActiveRecord.Lic_User = FNullField(rstRecord.Fields(FIELD_Lic_User))
        ' creatr functon to get DB picture and save it to class
        ' ActiveRecord.Lic_Logo = GetDBPicture(rstRecord.Fields(FIELD_Lic_LegalInfo))
        
        ADORecordsetClose rstRecord
        'Set rstRecord = Nothing
        
        GetRecord = True
    
    ElseIf ((ActiveRecord Is Nothing) = True) Then
        GetRecord = False
    End If
    
    Exit Function
    
ERROR_RECORDSET:
    ADORecordsetClose rstRecord
    'Set rstRecord = Nothing
    
    GetRecord = False
    Exit Function
ERROR_QUERY:
    GetRecord = False
End Function

Public Function GetMaxID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiLicensee
    End If

    strSql = SQL_GET_MAXID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.Lic_ID = rstRecord.Fields("ID_MAX").Value
    
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

Public Function GetMinID(ByRef ActiveConnection As ADODB.Connection, ByRef ActiveRecord As cpiLicensee) As Boolean

    Dim strSql As String
    Dim rstRecord As ADODB.Recordset

    If ((ActiveRecord Is Nothing) = True) Then
        Set ActiveRecord = New cpiLicensee
    End If

    strSql = SQL_GET_MINID

    On Error GoTo ERROR_QUERY
    ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
    'Set rstRecord = ActiveConnection.Execute(strSql)

    On Error GoTo ERROR_RECORDSET
    ActiveRecord.Lic_ID = rstRecord.Fields("ID_MIN").Value
    
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

' conversion
Public Function GetTableRecord(ByRef ActiveRecord As cpiLicensee) As ADODB.Recordset
    Dim rstRecord As ADODB.Recordset
    
    Set rstRecord = New ADODB.Recordset
    
    rstRecord.Open
    rstRecord.Fields.Append FIELD_Lic_ID, 3, 4, 90
    rstRecord.Fields.Append FIELD_Lic_Name, 202, 50, 102
    rstRecord.Fields.Append FIELD_Lic_Address, 202, 150, 102
    rstRecord.Fields.Append FIELD_Lic_City, 202, 25, 102
    rstRecord.Fields.Append FIELD_Lic_PostalCode, 202, 25, 102
    rstRecord.Fields.Append FIELD_Lic_Country, 202, 25, 102
    rstRecord.Fields.Append FIELD_Lic_Phone, 202, 25, 102
    rstRecord.Fields.Append FIELD_Lic_Fax, 202, 25, 102
    rstRecord.Fields.Append FIELD_Lic_Email, 202, 75, 102
    rstRecord.Fields.Append FIELD_Lic_LegalInfo, 203, 536870910, 230
    rstRecord.Fields.Append FIELD_Lic_Currency, 202, 50, 102
    rstRecord.Fields.Append FIELD_Lic_Database, 202, 50, 102
    rstRecord.Fields.Append FIELD_Lic_UseEXACTID, 11, 2, 86
    rstRecord.Fields.Append FIELD_Lic_Language, 202, 15, 102
    rstRecord.Fields.Append FIELD_Lic_Logo, 205, 1073741823, 230
    rstRecord.Fields.Append FIELD_Lic_Logosize, 202, 100, 102
    rstRecord.Fields.Append FIELD_Lic_LogoProperties, 202, 255, 102
    rstRecord.Fields.Append FIELD_Lic_Website, 202, 255, 102
    rstRecord.Fields.Append FIELD_Lic_Key, 202, 50, 102
    rstRecord.Fields.Append FIELD_Lic_User, 202, 50, 102
    
    ' set values
    rstRecord.AddNew
    rstRecord.Fields(FIELD_Lic_ID) = ActiveRecord.Lic_ID
    rstRecord.Fields(FIELD_Lic_Name) = ActiveRecord.Lic_Name
    rstRecord.Fields(FIELD_Lic_Address) = ActiveRecord.Lic_Address
    rstRecord.Fields(FIELD_Lic_City) = ActiveRecord.Lic_City
    rstRecord.Fields(FIELD_Lic_PostalCode) = ActiveRecord.Lic_PostalCode
    rstRecord.Fields(FIELD_Lic_Country) = ActiveRecord.Lic_Country
    rstRecord.Fields(FIELD_Lic_Phone) = ActiveRecord.Lic_Phone
    rstRecord.Fields(FIELD_Lic_Fax) = ActiveRecord.Lic_Fax
    rstRecord.Fields(FIELD_Lic_Email) = ActiveRecord.Lic_Email
    rstRecord.Fields(FIELD_Lic_LegalInfo) = ActiveRecord.Lic_LegalInfo
    rstRecord.Fields(FIELD_Lic_Currency) = ActiveRecord.Lic_Currency
    rstRecord.Fields(FIELD_Lic_Database) = ActiveRecord.Lic_Database
    rstRecord.Fields(FIELD_Lic_UseEXACTID) = ActiveRecord.Lic_UseEXACTID
    rstRecord.Fields(FIELD_Lic_Language) = ActiveRecord.Lic_Language
    rstRecord.Fields(FIELD_Lic_Logosize) = ActiveRecord.Lic_Logosize
    rstRecord.Fields(FIELD_Lic_LogoProperties) = ActiveRecord.Lic_LogoProperties
    rstRecord.Fields(FIELD_Lic_Website) = ActiveRecord.Lic_Website
    rstRecord.Fields(FIELD_Lic_Key) = ActiveRecord.Lic_Key
    rstRecord.Fields(FIELD_Lic_User) = ActiveRecord.Lic_User
    
    ' create function here to save variant picture from rst
    'rstRecord.Fields(FIELD_Lic_LegalInfo) = ActiveRecord.Lic_Logo
    rstRecord.Update
    
    Set GetTableRecord = rstRecord
    Set rstRecord = Nothing

End Function

Public Function GetClassRecord(ByRef ActiveRecord As ADODB.Recordset) As cpiLicensee

    Dim clsRecord As cpiLicensee
    
    Set clsRecord = New cpiLicensee
    
    clsRecord.Lic_ID = FNullField(ActiveRecord.Fields(FIELD_Lic_ID))
    clsRecord.Lic_Name = FNullField(ActiveRecord.Fields(FIELD_Lic_Name))
    clsRecord.Lic_Address = FNullField(ActiveRecord.Fields(FIELD_Lic_Address))
    clsRecord.Lic_City = FNullField(ActiveRecord.Fields(FIELD_Lic_City))
    clsRecord.Lic_PostalCode = FNullField(ActiveRecord.Fields(FIELD_Lic_PostalCode))
    clsRecord.Lic_Country = FNullField(ActiveRecord.Fields(FIELD_Lic_Country))
    clsRecord.Lic_Phone = FNullField(ActiveRecord.Fields(FIELD_Lic_Phone))
    clsRecord.Lic_Fax = FNullField(ActiveRecord.Fields(FIELD_Lic_Fax))
    clsRecord.Lic_Email = FNullField(ActiveRecord.Fields(FIELD_Lic_Email))
    clsRecord.Lic_LegalInfo = FNullField(ActiveRecord.Fields(FIELD_Lic_LegalInfo))
    clsRecord.Lic_Currency = FNullField(ActiveRecord.Fields(FIELD_Lic_Currency))
    clsRecord.Lic_Database = FNullField(ActiveRecord.Fields(FIELD_Lic_Database))
    clsRecord.Lic_UseEXACTID = FNullField(ActiveRecord.Fields(FIELD_Lic_UseEXACTID))
    clsRecord.Lic_Language = FNullField(ActiveRecord.Fields(FIELD_Lic_Language))
    clsRecord.Lic_Logosize = FNullField(ActiveRecord.Fields(FIELD_Lic_Logosize))
    clsRecord.Lic_LogoProperties = FNullField(ActiveRecord.Fields(FIELD_Lic_LogoProperties))
    clsRecord.Lic_Website = FNullField(ActiveRecord.Fields(FIELD_Lic_Website))
    clsRecord.Lic_Key = FNullField(ActiveRecord.Fields(FIELD_Lic_Key))
    clsRecord.Lic_User = FNullField(ActiveRecord.Fields(FIELD_Lic_User))
    
    ' create function here to convert rst pics to class
    'clsRecord.Lic_Logo = FNullField(ActiveRecord.Fields(FIELD_Lic_Logo))

    Set GetClassRecord = clsRecord
    Set clsRecord = Nothing

End Function

Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, ByVal SearchField, ByVal SearchValue) As Boolean
'
    Dim strSql As String
    Dim lngRecordsAffected As Long
    Dim rstRecord As ADODB.Recordset
    
    On Error GoTo ERROR_SEARCH
    SearchField = Trim$(SearchField)
    If ((SearchField <> "") And (SearchValue <> "")) Then
        If (Len(SearchField) > 2) Then
            If ((Left$(SearchField, 1) <> "[") And (Right$(SearchField, 1) <> "]")) Then
                SearchField = "[" & SearchField & "]"
            End If
        End If
    
        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TABLE_NAME & " WHERE " & SearchField & "=" & Trim$(SearchValue)
        ADORecordsetOpen strSql, ActiveConnection, rstRecord, adOpenKeyset, adLockOptimistic
        'ActiveConnection.Execute strSql, lngRecordsAffected
        
        SearchValue = Not (rstRecord.EOF And rstRecord.BOF)
        
        'If (lngRecordsAffected > 0) Then
        '    SearchValue = True
        'ElseIf (lngRecordsAffected = 0) Then
        '    SearchValue = False
        'End If
        
    End If
    
    Exit Function
    
ERROR_SEARCH:
    SearchValue = False
End Function

' /* --------------- PRIVATE FUNCTIONS -------------------- */



