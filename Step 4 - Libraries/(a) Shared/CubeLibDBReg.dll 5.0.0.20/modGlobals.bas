Attribute VB_Name = "modGlobals"
Option Explicit


Global Const G_CONST_CP_APP_NAME = "ClearingPoint"
Global Const G_CONST_CPTS_APP_NAME = "ClearingPoint Scheduler"
Global Const G_CONST_CPMAIL_APP_NAME = "ClearingPoint Reports Auto-Email"
Global Const G_CONST_CPRP_APP_NAME = "RemotePrinter"
Global Const G_CONST_PERSISTENCE_FILE = "persistence.txt"

Global g_objDataSource As CDatasource
Global g_strPersistencePath As String
Global g_blnNewPersistencePath As Boolean
Global g_strDatabasePath As String
Global g_enuDatabaseType As CubeLibDataSource.DatabaseType
Global g_strDatabasePassword As String

Global g_blnIsSavingToTracefile As Boolean

' API function - use for loading resource string from target resource file
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" ( _
                         ByVal hInstance As Long, _
                         ByVal wID As Long, _
                         ByVal lpBuffer As String, _
                         ByVal nBufferMax As Long) As Long

' used for BrowseCallbackProcStr
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Public Declare Function PathFileExists _
    Lib "shlwapi.dll" Alias "PathFileExistsA" ( _
    ByVal pszPath As String _
    ) As Long

Public Declare Function SendMessage _
    Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
    ) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const KEY_ENCRYPT = ""

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal lParam As Long, _
                              ByVal lpData As Long) _
                              As Long
    
    If uMsg = 1 Then
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If

End Function

Public Function FunctionPointer(FunctionAddress As Long) As Long
    FunctionPointer = FunctionAddress
End Function

Public Sub ADORecordsetClose_F(rstToClose As ADODB.Recordset)
    
    On Error GoTo ERROR_HANDLER_BOOKMARK
    If Not rstToClose Is Nothing Then
        If rstToClose.State = adStateOpen Then
            rstToClose.Close
        End If
        Set rstToClose = Nothing
    End If
    On Error GoTo 0
    
    Exit Sub
    
ERROR_HANDLER_BOOKMARK:
    Select Case Err.Number
        Case -2147467259
            Resume
        Case Else
            Err.Raise Err.Number, , Err.Description
    End Select
    
End Sub

Public Function ExecuteNonQuery_F(ByRef ADOConnection As ADODB.Connection, _
                                  ByVal strSQL As String) As Long
    
    Dim success As Long
    
    Dim year As String
    Dim databaseName As String
    
    'Dim DataSource As CDatasource
    '
    'Set DataSource = New CDatasource
    'g_objDataSource.SetPersistencePath g_strPersistencePath
    
    If g_blnNewPersistencePath Then
        Set g_objDataSource = Nothing
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
            
        g_blnNewPersistencePath = False
    
    ElseIf g_objDataSource Is Nothing Then
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
        
    End If
    
    On Error GoTo ErrHandler
    
    databaseName = GetDatabaseName(ADOConnection)
        
    If InStr(databaseName, "mdb_history") > 0 Then
        year = Right(databaseName, 2)
        
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_HISTORY, year)
        
    ElseIf InStr(databaseName, "mdb_repertory") > 0 Then
        If IsNumeric(Right(databaseName, 4)) Then
            year = Right(databaseName, 4)
        End If
        
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_REPERTORY, year)
        
    ElseIf InStr(databaseName, "mdb_EDIhistory") > 0 Then
        If IsNumeric(Right(databaseName, 2)) Then
            year = Right(databaseName, 2)
        End If
        
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_EDI_HISTORY, year)
        
    ElseIf InStr(databaseName, "mdb_sadbel") > 0 Then
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_SADBEL)
    
    ElseIf InStr(databaseName, "mdb_data") > 0 Then
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_DATA)
        
    ElseIf InStr(databaseName, "edifact") > 0 Then
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_EDIFACT)
        
    ElseIf InStr(databaseName, "mdb_scheduler") > 0 Then
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_SCHEDULER)
        
    ElseIf InStr(databaseName, "TemplateCP") > 0 Then
        success = g_objDataSource.ExecuteNonQuery(strSQL, DBInstanceType_DATABASE_TEMPLATE)
        
    ElseIf InStr(databaseName, "mdb_taric") > 0 Then
        Err.Raise vbObjectError + 603, , "Error in ExecuteNonQuery() - Taric update is not supported yet."
        
    Else
        success = g_objDataSource.ExecuteNonQueryOtherDB(strSQL, databaseName)
        
        Err.Raise vbObjectError + 604, , "Error in ExecuteNonQuery() - Unrecognized database name."
        
    End If
    
    ExecuteNonQuery_F = success
    
ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            Err.Raise Err.Number, , Err.Description
    End Select
    
End Function

Public Function GetSQLCommandFromTableName_F(ByVal TableName As String) As String
    Dim strSQLCommandFromTableName As String
    
        strSQLCommandFromTableName = vbNullString
        strSQLCommandFromTableName = strSQLCommandFromTableName & "SELECT "
        strSQLCommandFromTableName = strSQLCommandFromTableName & "* "
        strSQLCommandFromTableName = strSQLCommandFromTableName & "FROM "
        strSQLCommandFromTableName = strSQLCommandFromTableName & "[" & TableName & "] "
    GetSQLCommandFromTableName_F = strSQLCommandFromTableName
End Function

Public Function GetDatabaseName(ByRef conToUse As ADODB.Connection) As String
    
    Dim connString As String
    Dim databaseName As String
    Dim matched As Integer
    
    connString = conToUse.ConnectionString
    matched = InStrRev(connString, "\")
    
    If matched > 0 Then
        connString = Mid(connString, matched + 1)
        matched = InStr(1, connString, ";")
        
        If matched > 0 Then
            connString = Left$(connString, matched - 1)
        End If
        
        If UCase$(Trim$(Right$(connString, 4))) = UCase$(Trim$(".MDB")) Then
            connString = Replace(connString, ".mdb", vbNullString)
        End If
        
        GetDatabaseName = Trim$(connString)
        
        Exit Function
    End If
    
    Err.Raise -9999, , "Error in getDatabaseName() - Invalid connection string."
    
End Function

Public Sub ADORecordsetOpen_F(ByVal Source As String, _
                   ByRef conToUse As ADODB.Connection, _
                   ByRef rstToOpen As ADODB.Recordset, _
                   ByVal CursorType As CursorTypeEnum, _
                   ByVal LockType As LockTypeEnum, _
          Optional ByVal lngCacheSize As Long = 1, _
          Optional ByVal UseDataShaping As Boolean = False)
    
    
    Dim year As String
    Dim databaseName As String
    
    On Error GoTo ErrHandler
    
    'Dim DataSource As CDatasource
    '
    'Set DataSource = New CDatasource
    'g_objDataSource.SetPersistencePath g_strPersistencePath
    
    If g_blnNewPersistencePath Then
        Set g_objDataSource = Nothing
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
            
        g_blnNewPersistencePath = False
    
    ElseIf g_objDataSource Is Nothing Then
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
        
    End If
    
    databaseName = GetDatabaseName(conToUse)
    
    ADORecordsetClose_F rstToOpen
    
    Set rstToOpen = New ADODB.Recordset
    
    If InStr(databaseName, "mdb_history") > 0 Then
        year = Right(databaseName, 2)
        
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_HISTORY, UseDataShaping, year)
        
    ElseIf InStr(databaseName, "mdb_repertory") > 0 Then
        If IsNumeric(Right(databaseName, 4)) Then
            year = Right(databaseName, 4)
        End If
        
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_REPERTORY, UseDataShaping, year)
        
    ElseIf InStr(databaseName, "mdb_EDIhistory") > 0 Then
        If IsNumeric(Right(databaseName, 2)) Then
            year = Right(databaseName, 2)
        End If
        
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_EDI_HISTORY, UseDataShaping, year)
        
    ElseIf InStr(databaseName, "mdb_sadbel") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_SADBEL, UseDataShaping)
    
    ElseIf InStr(databaseName, "mdb_data") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_DATA, UseDataShaping)
        
    ElseIf InStr(databaseName, "edifact") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_EDIFACT, UseDataShaping)
        
    ElseIf InStr(databaseName, "mdb_scheduler") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_SCHEDULER, UseDataShaping)
        
    ElseIf InStr(databaseName, "TemplateCP") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_TEMPLATE, UseDataShaping)
        
    ElseIf InStr(databaseName, "mdb_taric") > 0 Then
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_TARIC, UseDataShaping)
    
    Else
        
        Set rstToOpen = g_objDataSource.ExecuteQuery(Source, DBInstanceType_DATABASE_OTHER, UseDataShaping, , databaseName & ".mdb")
        
        'Err.Raise vbObjectError + 600, , "Error in cpiOpen() - Unrecognized database name."
    End If
    
ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            Err.Raise Err.Number, , Err.Description
            
    End Select

End Sub

Public Function GetADOConnectionString_F(ByRef DataSourceProperties As CDataSourceProperties, _
                                         ByVal InitialCatalog As DBInstanceType, _
                                Optional ByVal InitialCatalogYear As String = vbNullString, _
                                Optional ByVal InitialCatalogName As String = vbNullString, _
                                Optional ByVal OpenExclusive As Boolean = False, _
                                Optional ByVal AltInitialCatlogPathMSAccess As String = vbNullString) _
                                         As String
                                  
    Dim strConnectionString As String
    
    Dim strInitialCatalog As String
    
    strInitialCatalog = GetDBInstanceTypeDesc_F(InitialCatalog, InitialCatalogYear, InitialCatalogName)
    
    Select Case DataSourceProperties.DatabaseType
        Case DatabaseType.DatabaseType_ACCESS97, _
              DatabaseType.DatabaseType_ACCESS2003
              
            strInitialCatalog = strInitialCatalog & ".mdb"
        Case DatabaseType.DatabaseType_SQLSERVER
            ' Do Nothing
        Case Else
            Debug.Assert False
    End Select

    With DataSourceProperties
        Select Case .DatabaseType
            Case DatabaseType.DatabaseType_ACCESS97, _
                 DatabaseType.DatabaseType_ACCESS2003
                
                ' ---------------------------------
                ' Jet OLEDB:Database Locking Mode=1
                ' ---------------------------------
                '  Access does not support transaction isolation in the manner of locking
                '  a transaction block for each individual user. SQL Server fully supports '
                '  Transaction Isolation and returning a database to its normal state if a
                '  rollback occurs and the code fails. ADO does not compensate for the
                '  downfails of Access :). This line of code will set a lock for updating or
                '  review particular records. There are two potential values for this long parameter:
                '
                '      Page-level Locking 0
                '      Row-level Locking 1
                '
                '  An Access database can only be opened in one mode at a time. The first user
                '  to open the database determines the locking mode to be used while the database
                '  is open. So in a multi-user environment, if multiple users are hitting this
                '  database at the same time, this property can not be modified for various instances.
                '  There are some tedious work arounds, but overall this can become troublesome.
                '
                '  And if you really want to work with this, you need to set recordset locking
                '  levels. rs.Properties("Jet OLEDB:Locking Granularity") = 2. This will allow you
                '  to set various levels of locking per row.Unless the Record Locking Mode is 1,
                '  this property is ignored by default.

                strConnectionString = vbNullString
                'strConnectionString = strConnectionString & "Data Provider=Microsoft.Jet.OLEDB.4.0;"
                strConnectionString = strConnectionString & "Provider=Microsoft.Jet.OLEDB.4.0;"
                
                If LenB(Trim$(AltInitialCatlogPathMSAccess)) > 0 Then
                    strConnectionString = strConnectionString & "Data Source=" & NoBackSlash(AltInitialCatlogPathMSAccess) & "\" & strInitialCatalog & ";"
                Else
                    strConnectionString = strConnectionString & "Data Source=" & NoBackSlash(.DataSource) & "\" & strInitialCatalog & ";"
                End If
                strConnectionString = strConnectionString & "Persist Security Info=False;"
                If OpenExclusive Then
                    strConnectionString = strConnectionString & "Mode=" & adModeShareExclusive & ";"
                Else
                    strConnectionString = strConnectionString & "Mode=Share Deny None;"
                End If
                If .DatabaseType = DatabaseType.DatabaseType_ACCESS97 Then
                    strConnectionString = strConnectionString & "Jet OLEDB:Engine Type=5;"
                Else
                    strConnectionString = strConnectionString & "Jet OLEDB:Engine Type=4;"
                End If
                strConnectionString = strConnectionString & "Jet OLEDB:Database Password=" & .Password & ";"
                strConnectionString = strConnectionString & "Jet OLEDB:Database Locking Mode=1"
                
            Case DatabaseType.DatabaseType_SQLSERVER
                
                '   Integrated Security = SSPI : this is equivalant to true.
                '           false User ID and Password are specified in the
                '           connection string. true Windows account credentials
                '           are used for authentication. Recognized values are
                '           true , false , yes , no , and SSPI .
                
                '   Persist Security = true means that the Password used for SQL
                '           authentication is not removed from the ConnectionString
                '           property of the connection.
                '
                '           When Integrated Security = true is used then the Persist
                '           Security is completely irelevant since it only applies
                '           to SQL authentication, not to windows/Integrated/SSPI.
                '
                '   Provider=SQLOLEDB.1 or SQLOLEDB
                '           The only thing the .1 does is specify the precise version
                '           number. Normally you don't need this as (I'm pretty sure)
                '           any install of a new OLEDB version will either overwrite
                '           the old one, or will default the version-unspecified
                '           provider to point to the newest one.
                
                strConnectionString = vbNullString
                strConnectionString = strConnectionString & "Data Provider=SQLOLEDB;"
                strConnectionString = strConnectionString & "Persist Security Info=False;"
                If LenB(Trim$(.Username)) > 0 And LenB(Trim$(.Password)) > 0 Then
                    'strConnectionString = strConnectionString & "Integrated Security=false;"
                    strConnectionString = strConnectionString & "User ID=" & .Username & ";"
                    strConnectionString = strConnectionString & "Password=" & .Password & ";"
                Else
                    strConnectionString = strConnectionString & "Integrated Security=SSPI;"
                End If
                strConnectionString = strConnectionString & "Initial Catalog=" & strInitialCatalog & ";"
                strConnectionString = strConnectionString & "Data Source=" & .DataSource
        End Select
        
        
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' "Driver={SQL Server}; SERVER=MYSERVER\SQLEXPRESS; Database=MYDB; Uid=sa; Pwd=Pass!@
                                '''''                'Provider=SQLNCLI10;Server=10.1.100.1;Database=DataJualLama;Uid=sa;Pwd=sa;
                                '''''                'Server=MyServer;Database=northwind;Trusted_Connection=yes
                                '''''
                                '''''                ' SQL CLIENT
                                '''''                '   Windows Authentication
                                '''''                '       "Persist Security Info=False;Integrated Security=true;Initial Catalog=AdventureWorks;Server=MSSQL1"
                                '''''                '       "Persist Security Info=False;Integrated Security=SSPI;database=AdventureWorks;server=(local)"
                                '''''                '       "Persist Security Info=False;Trusted_Connection=True;database=AdventureWorks;server=(local)"
                                '''''                '   SQL Server Authentication
                                '''''                '       "Persist Security Info=False;User ID=*****;Password=*****;Initial Catalog=AdventureWorks;Server=MySqlServer"
                                '''''                ' DATASHAPE
                                '''''                '       "Provider=MSDataShape;Data Provider=SQLOLEDB;Data Source=(local);Initial Catalog=pubs;Integrated Security=SSPI;"
                                '''''
                                '''''
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' SQL Server Authentication
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "User ID=" & UserName & ";" & _
                                '''''                                   "Password=" & Password & ";" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' SQL Server Authentication
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "User ID=" & UserName & ";" & _
                                '''''                                   "Password=" & Password & ";" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' Windows Authentication
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "database=" & InitialCatalog & ";" & _
                                '''''                                   "server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' Windows Authentication
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "database=" & InitialCatalog & ";" & _
                                '''''                                   "server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' DATA SHAPE
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' DATA SHAPE
                                '''''                ADOConnection.Provider = "MSDataShape"
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''            Else
                                '''''
                                '''''
                                '''''                ' ORIGINAL
                                '''''                ADOConnection.Open "Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Data Source=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''
                                '''''                ' SQL Server Authentication
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "User ID=" & UserName & ";" & _
                                '''''                                   "Password=" & Password & ";" & _
                                '''''                                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                                   "Server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' Windows Authentication
                                '''''                'ADOConnection.Open "Persist Security Info=False;" & _
                                '''''                '                   "Integrated Security=true;" & _
                                '''''                '                   "Initial Catalog=" & InitialCatalog & ";" & _
                                '''''                '                   "Server=" & DataSource
                                '''''                'ADOConnection.Close
                                '''''
                                '''''                ' Windows Authentication
                                '''''                ADOConnection.Open "Data Provider=SQLOLEDB;" & _
                                '''''                                   "Persist Security Info=False;" & _
                                '''''                                   "Integrated Security=SSPI;" & _
                                '''''                                   "database=" & InitialCatalog & ";" & _
                                '''''                                   "server=" & DataSource
                                '''''                ADOConnection.Close
                                '''''
                                '''''                ' Windows Authentication
                                '''''                'ADOConnection.Open "Persist Security Info=False;" & _
                                '''''                '                   "Trusted_Connection=True;" & _
                                '''''                '                   "database=" & InitialCatalog & ";" & _
                                '''''                '                   "server=" & DataSource
                                '''''                'ADOConnection.Close
                                '''''
                                '''''
                                '''''            End If
    
    End With
    
    GetADOConnectionString_F = strConnectionString
End Function

Public Sub ADODisconnectDB_F(ByRef ConToClose As ADODB.Connection)
    
    If (ConToClose Is Nothing = False) Then
        If ConToClose.State = ADODB.ObjectStateEnum.adStateOpen Then
            ConToClose.Close
        End If
        
        Set ConToClose = Nothing
    End If
    
End Sub


Public Function ADOConnectDB_F(ByRef ADOConnection As ADODB.Connection, _
                               ByRef DataSourceProperties As CDataSourceProperties, _
                               ByVal InitialCatalog As DBInstanceType, _
                      Optional ByVal InitialCatalogYear As String = vbNullString, _
                      Optional ByVal InitialCatalogName As String = vbNullString, _
                      Optional ByVal UseDataShaping As Boolean = False, _
                      Optional ByVal OpenExclusive As Boolean = False, _
                      Optional ByVal AltInitialCatlogPathMSAccess As String = vbNullString) _
                               As CErrObject
          
          Dim strConnectionString As String
          Dim objErrObject As CErrObject
          
5         On Error GoTo Error_Handler
          
10        ADODisconnectDB_F ADOConnection

          Set objErrObject = New CErrObject
          
20        strConnectionString = GetADOConnectionString_F(DataSourceProperties, InitialCatalog, InitialCatalogYear, InitialCatalogName, OpenExclusive)
          
          If LenB(Trim$(strConnectionString)) <= 0 Then
              
              objErrObject.Number = -1
              objErrObject.Description = "Could not build connection string in GetADOConnectionString_F."
              
              Set ADOConnectDB_F = objErrObject
              
              Exit Function
          End If
          
90        Set ADOConnection = New ADODB.Connection

          '**************************************************************************
          'Open Connection depending on Provider Type
          '**************************************************************************
110       With DataSourceProperties
120           Select Case .DatabaseType
                  Case DatabaseType.DatabaseType_ACCESS97, _
                       DatabaseType.DatabaseType_ACCESS2003
                      
130                   If Len(Dir(NoBackSlash(.DataSource) & "\" & GetDBInstanceTypeDesc_F(InitialCatalog, InitialCatalogYear, InitialCatalogName) & ".mdb")) > 0 Then
                      
140                       If UseDataShaping = True Then
150                           ADOConnection.Provider = "MSDataShape"
160                           ADOConnection.Open strConnectionString
170                       Else
180                           ADOConnection.Open strConnectionString
190                       End If
                      
200                   Else
210                       Exit Function
220                   End If
                      
230               Case DatabaseType.DatabaseType_SQLSERVER

240                   If UseDataShaping = True Then
250                       ADOConnection.Provider = "MSDataShape"
260                   Else
270                       ADOConnection.Provider = "SQLOLEDB"
280                   End If

290                   ADOConnection.Open strConnectionString
          
                      
      '''''            Case ProviderType.[DBCFS Provider]
      '''''                'denDB.conDB.ConnectionString = "Extended Properties=" & Chr(34) & "FileDSN=" & DataSource & ";" & Chr(34)
      '''''                'denDB.conDB.Open
      '''''
      '''''                'Set ADOConnection = denDB.conDB
      '''''
      '''''            Case ProviderType.[Informix Provider]
      '''''                ADOConnection.Open "Provider=Ifxoledbc;" & _
      '''''                                   "Persist Security Info=False;" & _
      '''''                                   "Data Source=" & .DataSource & ";" & _
      '''''                                   "User ID=" & .Username & ";" & _
      '''''                                   "Password=" & .Password
300                   Case Else
310                       Debug.Assert False
320           End Select
330       End With
          '**************************************************************************
          
340       On Error GoTo 0
          
350       Set ADOConnectDB_F = objErrObject
360       Set objErrObject = Nothing
          
380       Exit Function

Error_Handler:

          objErrObject.CloneErrObject Err

410       Set ADOConnectDB_F = objErrObject
420       Set objErrObject = Nothing
End Function

Public Function IsCPDatabase_F(ByVal DBName As String) As Boolean
    Dim blnCPDatabase As Boolean
    Dim strDBName As String
    
    blnCPDatabase = False
    strDBName = UCase$(Trim$(DBName))
    
    If Right$(strDBName, 4) = UCase$(".mdb") Then
        strDBName = Left$(strDBName, Len(strDBName) - 4)
    End If

    blnCPDatabase = blnCPDatabase Or (strDBName = "EDIFACT")
    blnCPDatabase = blnCPDatabase Or (strDBName = "MDB_DATA")
    blnCPDatabase = blnCPDatabase Or (strDBName = "MDB_EDIHISTORY")
    blnCPDatabase = blnCPDatabase Or (strDBName = "MDB_SADBEL")
    blnCPDatabase = blnCPDatabase Or (strDBName = "MDB_SCHEDULER")
    blnCPDatabase = blnCPDatabase Or (strDBName = "MDB_TARIC")
    blnCPDatabase = blnCPDatabase Or (strDBName = "TEMPLATECP")
    
    blnCPDatabase = blnCPDatabase Or IsEDIHistoryDB_F(strDBName)
    blnCPDatabase = blnCPDatabase Or IsEDIHistoryDB_F(strDBName)
    blnCPDatabase = blnCPDatabase Or IsRepertoryDB_F(strDBName)
    
    IsCPDatabase_F = blnCPDatabase
End Function

Public Function IsEDIHistoryDB_F(ByVal DBName As String) As Boolean
    
    Dim strDBName As String
    Dim blnEDIHistoryDB As Boolean
    
    blnEDIHistoryDB = False
    strDBName = Trim$(UCase$(DBName))
    
    ' Remove Preceding backslash
    If Left$(strDBName, 1) = "\" Then
        strDBName = Mid(strDBName, 2)
    End If
    
    ' REMOVE FILE EXTENSION
    If Right$(strDBName, 4) = UCase$(".mdb") Then
        strDBName = Left$(strDBName, Len(strDBName) - 4)
    End If
    
    If Left$(strDBName, 14) = "MDB_EDIHISTORY" And _
        (Len(strDBName) = 16) Then
        
        If IsNumeric(Mid$(strDBName, 15, 2)) Then
            blnEDIHistoryDB = True
        End If
        
    End If
    
    IsEDIHistoryDB_F = blnEDIHistoryDB
End Function

Public Function IsHistoryDB_F(ByVal DBName As String) As Boolean
    Dim strDBName As String
    Dim blnHistoryDB As Boolean
    
    blnHistoryDB = False
    strDBName = Trim$(UCase$(DBName))
    
    ' Remove Preceding backslash
    If Left$(strDBName, 1) = "\" Then
        strDBName = Mid(strDBName, 2)
    End If
    
    ' REMOVE FILE EXTENSION
    If Right$(strDBName, 4) = UCase$(".mdb") Then
        strDBName = Left$(strDBName, Len(strDBName) - 4)
    End If
    
    If Left$(strDBName, 11) = "MDB_HISTORY" And _
        (Len(strDBName) = 13) Then

        If IsNumeric(Mid$(strDBName, 12, 2)) Then
            blnHistoryDB = True
        End If
    End If
    
    IsHistoryDB_F = blnHistoryDB
End Function

Public Function IsRepertoryDB_F(ByVal DBName As String) As Boolean
    
    Dim strDBName As String
    Dim blnRepertoryDB As Boolean
    
    blnRepertoryDB = False
    strDBName = Trim$(UCase$(DBName))
    
    ' Remove Preceding backslash
    If Left$(strDBName, 1) = "\" Then
        strDBName = Mid(strDBName, 2)
    End If
    
    ' REMOVE FILE EXTENSION
    If Right$(strDBName, 4) = UCase$(".mdb") Then
        strDBName = Left$(strDBName, Len(strDBName) - 4)
    End If
    
    If strDBName = UCase$("mdb_repertory") Then
        
        blnRepertoryDB = True
        
    ElseIf Left$(strDBName, 14) = "MDB_REPERTORY_" And _
            (Len(strDBName) = 18) Then
            
        If IsNumeric(Mid$(strDBName, 15, 4)) Then
            blnRepertoryDB = True
        End If
    End If
    
    IsRepertoryDB_F = blnRepertoryDB
End Function


Public Function GetRepertoryDBYear_F(ByVal RepertoryDBName As String) As String
    Dim strRepertoryDBName As String
    
    strRepertoryDBName = UCase$(Trim$(RepertoryDBName))
    
    ' Remove Preceding backslash like in strCurrentYear in CubeLibRepertorium
    If Left$(strRepertoryDBName, 1) = "\" Then
        strRepertoryDBName = Mid(strRepertoryDBName, 2)
    End If
    
    ' REMOVE .MDB FILE EXTENSION
    If Right$(strRepertoryDBName, 4) = UCase$(".mdb") Then
        strRepertoryDBName = Left$(strRepertoryDBName, Len(strRepertoryDBName) - 4)
    End If
    
    If IsRepertoryDB_F(strRepertoryDBName) Then
        If strRepertoryDBName = "MDB_REPERTORY" Then
            strRepertoryDBName = vbNullString
        Else
            strRepertoryDBName = Right$(strRepertoryDBName, 4)
        End If
    Else
        strRepertoryDBName = vbNullString
    End If
    
    GetRepertoryDBYear_F = strRepertoryDBName
End Function

Public Function GetHistoryDBYear_F(ByVal HistoryDBName As String) As String
    Dim strHistoryDBName As String
    
    strHistoryDBName = UCase$(Trim$(HistoryDBName))
    
    ' Remove Preceding backslash
    If Left$(strHistoryDBName, 1) = "\" Then
        strHistoryDBName = Mid(strHistoryDBName, 2)
    End If
    
    ' REMOVE .MDB FILE EXTENSION
    If Right$(strHistoryDBName, 4) = UCase$(".mdb") Then
        strHistoryDBName = Left$(strHistoryDBName, Len(strHistoryDBName) - 4)
    End If
    
    If IsHistoryDB_F(strHistoryDBName) Then
        strHistoryDBName = Right$(strHistoryDBName, 2)
    Else
        strHistoryDBName = vbNullString
    End If
    
    GetHistoryDBYear_F = strHistoryDBName
End Function

Public Function GetEDIHistoryDBYear_F(ByVal EDIHistoryDBName As String) As String
    Dim strEDIHistoryDBName As String
    
    strEDIHistoryDBName = UCase$(Trim$(EDIHistoryDBName))
    
    ' Remove Preceding backslash
    If Left$(strEDIHistoryDBName, 1) = "\" Then
        strEDIHistoryDBName = Mid(strEDIHistoryDBName, 2)
    End If
    
    ' REMOVE .MDB FILE EXTENSION
    If Right$(strEDIHistoryDBName, 4) = UCase$(".mdb") Then
        strEDIHistoryDBName = Left$(strEDIHistoryDBName, Len(strEDIHistoryDBName) - 4)
    End If
    
    If IsEDIHistoryDB_F(strEDIHistoryDBName) Then
        strEDIHistoryDBName = Right$(strEDIHistoryDBName, 2)
    Else
        strEDIHistoryDBName = vbNullString
    End If
    
    GetEDIHistoryDBYear_F = strEDIHistoryDBName
End Function


Public Function GetDBInstanceTypeDesc_F(ByVal InitialCatalog As DBInstanceType, _
                                        ByVal InitialCatalogYear As String, _
                                        ByVal InitialCatalogName As String) As String
    
    Dim strDBInstanceTypeDesc As String
    Dim strInitialCatalogYear As String
    Dim strInitialCatalog As String
    
    Dim blnValid As Boolean
    
    Select Case InitialCatalog
    
        Case DBInstanceType_DATABASE_OTHER
            
            strDBInstanceTypeDesc = Trim$(InitialCatalogName)
            
            If LenB(strDBInstanceTypeDesc) > 0 Then
                If Right$(UCase$(strDBInstanceTypeDesc), 4) = UCase$(Trim$(".mdb")) Then
                    strDBInstanceTypeDesc = Left$(strDBInstanceTypeDesc, Len(strDBInstanceTypeDesc) - 4)
                End If
            End If
            
        Case DBInstanceType.DBInstanceType_DATABASE_DATA
            strDBInstanceTypeDesc = "mdb_data"
        Case DBInstanceType.DBInstanceType_DATABASE_EDIFACT
            strDBInstanceTypeDesc = "edifact"
        Case DBInstanceType.DBInstanceType_DATABASE_SADBEL
            strDBInstanceTypeDesc = "mdb_sadbel"
        Case DBInstanceType.DBInstanceType_DATABASE_SCHEDULER
            strDBInstanceTypeDesc = "mdb_scheduler"
        Case DBInstanceType.DBInstanceType_DATABASE_TARIC
            strDBInstanceTypeDesc = "mdb_taric"
        Case DBInstanceType.DBInstanceType_DATABASE_TEMPLATE
            strDBInstanceTypeDesc = "TemplateCP"
            
        Case DBInstanceType_DATABASE_EDI_HISTORY, _
             DBInstanceType_DATABASE_HISTORY, _
             DBInstanceType_DATABASE_REPERTORY
                 
            strInitialCatalogYear = Trim$(InitialCatalogYear)
            
            If InStr(1, strInitialCatalogYear, "_") > 0 Then
            
                If LenB(Replace(strInitialCatalogYear, "_", "")) > 0 Then
                
                    Do While Left$(strInitialCatalogYear, "_")
                        strInitialCatalogYear = Mid(strInitialCatalogYear, 2)
                    Loop
                    
                Else
                    strDBInstanceTypeDesc = vbNullString
                    
                    GetDBInstanceTypeDesc_F = strDBInstanceTypeDesc
                    
                    Exit Function
                End If
            End If
            
            blnValid = True
            blnValid = blnValid And (LenB(strInitialCatalogYear) > 0)
            blnValid = blnValid And IsNumeric(strInitialCatalogYear)
            
            blnValid = blnValid And (LenB(strInitialCatalogYear) = 2 Or _
                         LenB(strInitialCatalogYear) = 4)
            
            If blnValid Then
            
                Select Case InitialCatalog
                
                    Case DBInstanceType_DATABASE_EDI_HISTORY
                    
                        strInitialCatalogYear = Format(InitialCatalogYear, "0000")
                        
                        strInitialCatalog = "mdb_EDIhistory" & Right$(strInitialCatalogYear, 2)
                        
                    Case DBInstanceType_DATABASE_HISTORY
    
                        strInitialCatalogYear = Format(InitialCatalogYear, "0000")
                        
                        strInitialCatalog = "mdb_history" & Right$(strInitialCatalogYear, 2)
                        
                    Case DBInstanceType_DATABASE_REPERTORY
                        
                        If LenB(strInitialCatalogYear) = 2 Then
                           
                           If Val(strInitialCatalogYear) > 97 Then
                                strInitialCatalogYear = "19" & strInitialCatalogYear
                            Else
                                strInitialCatalogYear = "20" & strInitialCatalogYear
                            End If
                        
                        End If
                        
                        strInitialCatalog = "mdb_repertory_" & strInitialCatalogYear
                End Select
                
                strDBInstanceTypeDesc = strInitialCatalog
                
            Else
                strDBInstanceTypeDesc = vbNullString
            End If
        
        
            
        Case Else
        
            strDBInstanceTypeDesc = vbNullString
            
            Debug.Assert False
            
    End Select
        
    GetDBInstanceTypeDesc_F = strDBInstanceTypeDesc
    
End Function

Public Function GetADOConnectionStringProperty_F(ByVal ConnectionString As String, _
                                               ByVal ADOConnectionStringProperty As ADOConnectionStringPropertyConstant)
    
    Dim strConnectionString As String
    Dim arrConnParams() As String
    Dim arrParamNameValue() As String
    Dim lngParamCtr As Long
    Dim blnNonJet As Boolean
    Dim strProperty As String
    
    strConnectionString = ConnectionString
    
    arrConnParams = Split(strConnectionString, ";")
    
    blnNonJet = False
    For lngParamCtr = LBound(arrConnParams) To UBound(arrConnParams)
        arrParamNameValue = Split(arrConnParams(lngParamCtr), "=")
        
        blnNonJet = blnNonJet Or (UCase$(Trim$(arrParamNameValue(0))) = UCase$(Trim$("Data Provider")))
        
        If blnNonJet Then
            Exit For
        End If
    Next
    
    For lngParamCtr = LBound(arrConnParams) To UBound(arrConnParams)
        arrParamNameValue = Split(arrConnParams(lngParamCtr), "=")
        
        Select Case ADOConnectionStringProperty
            Case ADOConnectionStringPropertyConstant.[Connection Initial Catalog]
                
                If blnNonJet Then
                    
                    If UCase$(Trim$(arrParamNameValue(0))) = UCase$(Trim$("INITIAL CATALOG")) Then
                    
                        strProperty = arrParamNameValue(1)
                        Exit For
                        
                    End If
                    
                ElseIf UCase$(Trim$(arrParamNameValue(0))) = UCase$(Trim$("DATA SOURCE")) Then
                    
                    strProperty = Mid(arrParamNameValue(1), InStrRev(arrParamNameValue(1), "\") + 1)
                    Exit For
                    
                End If
                
            Case ADOConnectionStringPropertyConstant.[Connection Initial Catalog Path]
                
                If blnNonJet Then
                    strProperty = vbNullString
                    Exit For
                    
                ElseIf UCase$(Trim$(arrParamNameValue(0))) = UCase$(Trim$("DATA SOURCE")) Then
                
                    strProperty = Left$(arrParamNameValue(1), InStrRev(arrParamNameValue(1), "\") - 1)
                    Exit For
                    
                End If
                
        End Select

    Next
    
    GetADOConnectionStringProperty_F = strProperty
    
End Function

