VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLinkedTable"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function WNetGetConnection Lib _
                "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName _
                As String, ByVal lpszRemoteName As String, _
                cbRemoteName As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                nSize As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex _
                As Long, ByVal lpValueName As String, lpcbValueName _
                As Long, ByVal lpReserved As Long, lpType As Long, _
                ByVal lpData As String, lpcbData As Long) As Long
   
Public Function RefreshLinkedTable(ByVal LinkedTableName As String)
    'Refresh Link...
    
    Dim adoCn As ADODB.Connection
    Dim adoCat As New ADOX.Catalog
    Dim adoTbl As New ADOX.Table
    
    Set adoCat = New ADOX.Catalog
    Set adoCat.ActiveConnection = adoCn
    
    Set adoTbl.ParentCatalog = adoCat
    
    For Each adoTbl In adoCat.Tables
        If (UCase$(Trim$(adoTbl.Type)) = UCase$(Trim$("PASS-THROUGH")) Or _
           UCase$(Trim$(adoTbl.Type)) = UCase$(Trim$("LINK"))) And _
           (UCase$(Trim$(adoTbl.Name)) = UCase$(Trim$(LinkedTableName))) Then
           
            '"Microsoft.JET.OLEDB.4.0;Jet OLEDB:Database Password=myPassword"
            adoTbl.Properties("Jet OLEDB:Link Provider String") = "MS Access;Pwd=myLinkPassword"
            adoTbl.Properties("Jet OLEDB:Link Datasource") = App.Path & "\myLinkDatabase.mdb"
        End If
    Next

End Function

Public Function DropLinkedTable(ByRef ADODatabase As ADODB.Connection, _
                                ByVal LinkedTableName As String) _
                                As Boolean
    
    Dim catCatalog As ADOX.Catalog
    Dim tblCatalog As ADOX.Table
    Dim blnIsLinkedTableExisting As Boolean
    
    Set catCatalog = New ADOX.Catalog
    Set catCatalog.ActiveConnection = ADODatabase
    
    For Each tblCatalog In catCatalog.Tables
        If (UCase$(Trim$(tblCatalog.Type)) = UCase$(Trim$("PASS-THROUGH")) Or _
           UCase$(Trim$(tblCatalog.Type)) = UCase$(Trim$("LINK"))) And _
           UCase$(Trim$(tblCatalog.Name)) = UCase$(Trim$(LinkedTableName)) Then
           
            catCatalog.Tables.Delete catCatalog.Tables.Item(LinkedTableName)
            
            Exit For
        End If
    Next
    
    DropLinkedTable = ADOXIsLinkedTableExisting_F(ADODatabase, LinkedTableName)
    blnIsLinkedTableExisting = ADOXIsLinkedTableExisting_F(ADODatabase, LinkedTableName)
    
    DropLinkedTable = (Not blnIsLinkedTableExisting)
    
    Set catCatalog = Nothing
End Function

Public Function CreateLinkedTable(ByRef DataSourceProperties As CDataSourceProperties, _
                                  ByVal DestinationCatalog As DBInstanceType, _
                                  ByVal DestinationTable As String, _
                                  ByVal SourceCatalog As DBInstanceType, _
                                  ByVal SourceTable As String, _
                         Optional ByVal DestinationCatalogYear As String = vbNullString, _
                         Optional ByVal SourceCatalogYear As String = vbNullString, _
                         Optional ByVal DestinationCatalogName As String = vbNullString, _
                         Optional ByVal SourceCatalogName As String = vbNullString, _
                         Optional ByVal AltDestinationInitialCatlogPathMSAccess As String = vbNullString, _
                         Optional ByVal AltSourceInitialCatlogPathMSAccess As String = vbNullString) As Boolean
    
    Dim adoCn As ADODB.Connection
    Dim adoCat As ADOX.Catalog
    Dim adoTbl As ADOX.Table

    Dim strSourceDataSource As String
    Dim strSourceCatalogYear As String
    Dim strLinkDataSource As String
    
    Dim strSourceFinalInitialCatalogName As String
    
    If LenB(Trim$(AltSourceInitialCatlogPathMSAccess)) > 0 Then
        strLinkDataSource = Trim$(AltSourceInitialCatlogPathMSAccess)
    Else
        strLinkDataSource = Trim$(DataSourceProperties.DataSource)
    End If
    
    If LenB(Trim$(DestinationCatalogName)) > 0 And _
        DestinationCatalog = DBInstanceType_DATABASE_OTHER Then
        
        ADOConnectDB_F adoCn, DataSourceProperties, DestinationCatalog, , DestinationCatalogName, , , AltDestinationInitialCatlogPathMSAccess
    Else
        ADOConnectDB_F adoCn, DataSourceProperties, DestinationCatalog, DestinationCatalogYear, , , , AltDestinationInitialCatlogPathMSAccess
    End If
    
    strSourceFinalInitialCatalogName = GetDBInstanceTypeDesc_F(SourceCatalog, SourceCatalogYear, SourceCatalogName)
    
    If Dir(AddBackSlashOnPath(strLinkDataSource) & strSourceFinalInitialCatalogName) = "" Then
        strSourceDataSource = GetUNCNameNT(AddBackSlashOnPath(strLinkDataSource) & strSourceFinalInitialCatalogName)
    Else
        strSourceDataSource = AddBackSlashOnPath(strLinkDataSource) & strSourceFinalInitialCatalogName
    End If

    ' Create Link...
    Set adoCat = New ADOX.Catalog
    Set adoCat.ActiveConnection = adoCn
    
    Dim strExistingLinkedTableDataSource As String
    Dim blnLinkEOF As Boolean
    
    blnLinkEOF = True
    
    For Each adoTbl In adoCat.Tables
        If (UCase$(Trim$(adoTbl.Type)) = UCase$(Trim$("PASS-THROUGH")) Or _
           UCase$(Trim$(adoTbl.Type)) = UCase$(Trim$("LINK"))) And _
           UCase$(Trim$(adoTbl.Name)) = UCase$(Trim$(DestinationTable)) Then

            strExistingLinkedTableDataSource = adoTbl.Properties("Jet OLEDB:Link Datasource")
            
            If Dir(strExistingLinkedTableDataSource) = "" Then
                strExistingLinkedTableDataSource = GetUNCNameNT(strExistingLinkedTableDataSource)
            End If
            
            If UCase$(strSourceDataSource) <> UCase$(strExistingLinkedTableDataSource) Or _
                adoTbl.Properties("Jet OLEDB:Remote Table Name") <> SourceTable Then
                
                If adoTbl.Properties("Jet OLEDB:Remote Table Name") <> SourceTable Then
                    adoCat.Tables.Delete adoTbl.Name
                Else
                    adoTbl.Properties("Jet OLEDB:Link Datasource") = strSourceDataSource
                    adoTbl.Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & DataSourceProperties.Password
                    
                    blnLinkEOF = False
                End If
            End If
            
            Exit For
        End If
    Next
    
    If blnLinkEOF Then
        If Not adoTbl Is Nothing Then
            Set adoTbl = Nothing
        End If
        Set adoTbl = New ADOX.Table
        
        Set adoTbl.ParentCatalog = adoCat
        adoTbl.Name = DestinationTable
        
        adoTbl.Properties("Jet OLEDB:Link Datasource") = strSourceDataSource
        adoTbl.Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & DataSourceProperties.Password
        adoTbl.Properties("Jet OLEDB:Remote Table Name") = SourceTable
        adoTbl.Properties("Jet OLEDB:Create Link") = True
    
        
        On Error Resume Next
        'Append the table to the tables collection
        adoCat.Tables.Append adoTbl
        
        Select Case Err.Number
        
            Case -2147217857    ' Object ' ' already exists.
            
                If UCase$(Left$(Err.Description, 8)) = UCase$("Object '") And _
                    UCase$(Right$(Err.Description, 17)) = UCase$("' already exists.") Then
                    
                    ' Do Nothing
                Else
                    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                End If
                
            Case 0
                ' Do Nothing
                
            Case Else
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        End Select
        
        On Error GoTo 0
    End If
    
    CreateLinkedTable = ADOXIsLinkedTableExisting_F(adoCn, DestinationTable)
    
    ADODisconnectDB_F adoCn
    
    Set adoCat = Nothing
    Set adoTbl = Nothing
    
End Function

Public Function GetUNCNameNT(pathName As String) As String

Dim hKey As Long
Dim hKey2 As Long
Dim exitFlag As Boolean
Dim i As Double
Dim ErrCode As Long
Dim rootKey As String
Dim Key As String
Dim computerName As String
Dim lComputerName As Long
Dim stPath As String
Dim firstLoop As Boolean
Dim Ret As Boolean

' first, verify whether the disk is connected to the network
If Mid(pathName, 2, 1) = ":" Then
   Dim UNCName As String
   Dim lenUNC As Long

   UNCName = String$(520, 0)
   lenUNC = 520
   ErrCode = WNetGetConnection(Left(pathName, 2), UNCName, lenUNC)

   If ErrCode = 0 Then
      UNCName = Trim(Left$(UNCName, InStr(UNCName, _
        vbNullChar) - 1))
      GetUNCNameNT = UNCName & Mid(pathName, 3)
      Exit Function
   End If
End If

' else, scan the registry looking for shared resources
'(NT version)
computerName = String$(255, 0)
lComputerName = Len(computerName)
ErrCode = GetComputerName(computerName, lComputerName)
If ErrCode <> 1 Then
   GetUNCNameNT = pathName
   Exit Function
End If

computerName = Trim(Left$(computerName, InStr(computerName, _
   vbNullChar) - 1))
rootKey = "SYSTEM\CurrentControlSet\Services\LanmanServer\Shares"
ErrCode = RegOpenKey(HKEY_LOCAL_MACHINE, rootKey, hKey)

If ErrCode <> 0 Then
   GetUNCNameNT = pathName
   Exit Function
End If

firstLoop = True

Do Until exitFlag
   Dim szValue As String
   Dim szValueName As String
   Dim cchValueName As Long
   Dim dwValueType As Long
   Dim dwValueSize As Long

   szValueName = String(1024, 0)
   cchValueName = Len(szValueName)
   szValue = String$(500, 0)
   dwValueSize = Len(szValue)

   ' loop on "i" to access all shared DLLs
   ' szValueName will receive the key that identifies an element
   ErrCode = RegEnumValue(hKey, i#, szValueName, _
       cchValueName, 0, dwValueType, szValue, dwValueSize)

   If ErrCode <> 0 Then
      If Not firstLoop Then
         exitFlag = True
      Else
         i = -1
         firstLoop = False
      End If
   Else
      stPath = GetPath(szValue)
      If firstLoop Then
         Ret = (UCase(stPath) = UCase(pathName))
         stPath = ""
      Else
         Ret = (UCase(stPath) = UCase(Left$(pathName, _
        Len(stPath))))
         stPath = Mid$(pathName, Len(stPath) + 1)
      End If
      If Ret Then
         exitFlag = True
         szValueName = Left$(szValueName, cchValueName)
         GetUNCNameNT = "\\" & computerName & "\" & _
            szValueName & stPath
      End If
   End If
   i = i + 1
Loop

RegCloseKey hKey
If GetUNCNameNT = "" Then GetUNCNameNT = pathName

End Function


Private Function GetPath(st As String) As String
   Dim pos1 As Long, pos2 As Long, pos3 As Long
   Dim stPath As String

   pos1 = InStr(st, "Path")
   If pos1 > 0 Then
      pos2 = InStr(pos1, st, vbNullChar)
      stPath = Mid$(st, pos1, pos2 - pos1)
      pos3 = InStr(stPath, "=")
      If pos3 > 0 Then
         stPath = Mid$(stPath, pos3 + 1)
         GetPath = stPath
      End If
   End If
End Function

