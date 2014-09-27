VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecordset"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum ExecuteRecordsetConstant
    [Insert]
    [Update]
End Enum

Public Enum ADOConnectionStringPropertyConstant
    [Connection Initial Catalog]
    [Connection Initial Catalog Path]
End Enum

Public Function ADOOptimizedFilter(ByRef ADORecordset As ADODB.Recordset, _
                                   ByVal Filter As String, _
                                   ByRef FilterFields() As String)
    
    Dim lngFieldCtr As Long
    
    If Not (ADORecordset.EOF And ADORecordset.BOF) Then
        
        If ADORecordset.CursorLocation = adUseClient Then
            For lngFieldCtr = LBound(FilterFields()) To UBound(FilterFields())
                ADORecordset.Fields(FilterFields(lngFieldCtr)).Properties("Optimize").Value = True  ' Create an Index
            Next
        End If
        
        ADORecordset.Filter = Filter
        
        If ADORecordset.CursorLocation = adUseClient Then
            For lngFieldCtr = LBound(FilterFields()) To UBound(FilterFields())
                ADORecordset.Fields(FilterFields(lngFieldCtr)).Properties("Optimize").Value = False 'Delete an index
            Next
        End If
    End If
End Function

Public Function ADOOptimizedFind(ByRef ADORecordset As ADODB.Recordset, _
                                 ByVal FindFieldName As String, _
                                 ByVal FindCriteria As String, _
                        Optional ByVal SkipRecords As Long, _
                        Optional ByVal SearchDirection As SearchDirectionEnum, _
                        Optional ByVal Start)
    Dim blnDone As Boolean
    
    blnDone = False
    
    If Not (ADORecordset.EOF And ADORecordset.BOF) Then
        If ADORecordset.CursorLocation = adUseClient Then
            ADORecordset.Fields(FindFieldName).Properties("Optimize").Value = True ' Create an Index
        End If
        
        If IsMissing(Start) And _
           IsMissing(SearchDirection) Then
                
            blnDone = True
            
            ADORecordset.Find FindCriteria, _
                              IIf(IsMissing(SkipRecords), vbEmpty, SkipRecords), _
                              adSearchForward, 1
        End If
        
        If Not blnDone Then
            ADORecordset.Find FindCriteria, _
                              IIf(IsMissing(SkipRecords), vbEmpty, SkipRecords), _
                              IIf(IsMissing(SearchDirection), vbEmpty, SearchDirection), _
                              IIf(IsMissing(Start), vbEmpty, Start)
        End If
        
        If ADORecordset.CursorLocation = adUseClient Then
            ADORecordset.Fields(FindFieldName).Properties("Optimize").Value = False 'Delete an index
        End If
        
    End If
End Function

Public Function ADOBeginTrans(ByVal InitialCatalog As DBInstanceType, _
                     Optional ByVal InitialCatalogYear As String = vbNullString, _
                     Optional ByVal InitialCatalogName As String = vbNullString, _
                     Optional ByVal UseDataShaping As Boolean = False) As Long
    
    Dim objRecord As CubeLibDataSource.DNetRecordset
    Dim success As Integer
    
    Dim year As String
    Dim databaseName As String
    
    'Dim DataSource As CDatasource
    '
    'Set DataSource = New CDatasource
    'DataSource.SetPersistencePath g_strPersistencePath
    
    If g_blnNewPersistencePath Then
        Set g_objDataSource = Nothing
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
            
        g_blnNewPersistencePath = False
        
    ElseIf g_objDataSource Is Nothing Then
        Set g_objDataSource = New CDatasource
        g_objDataSource.SetPersistencePath g_strPersistencePath
    End If
    
    g_objDataSource.BeginTransaction InitialCatalog, InitialCatalogYear, InitialCatalogName, UseDataShaping
End Function

Public Function GetSQLCommandFromTableName(ByVal TableName As String) As String

    GetSQLCommandFromTableName = GetSQLCommandFromTableName_F(TableName)
End Function

Public Function ExecuteNonQuery(ByRef ADOConnection As ADODB.Connection, _
                                ByVal strSQL As String) As Long
    
    ExecuteNonQuery = ExecuteNonQuery_F(ADOConnection, strSQL)
    
End Function

Public Sub DelayTime(ByVal NumOfSecs As Single)
    Dim sngEndTime As Single
    
    ' 5 seconds Delay loop - Edwin Jan16
    sngEndTime = Timer + 5
    Do
    Loop Until Timer >= sngEndTime
End Sub

Public Function ExecuteRecordset(ByVal ExecuteMode As ExecuteRecordsetConstant, _
                                 ByRef ADOConnection As ADODB.Connection, _
                                 ByRef rstADO As ADODB.Recordset, _
                                 ByVal tableToUpdate As String) _
                                 As Long

    Select Case ExecuteMode
    
        Case ExecuteRecordsetConstant.Insert
        
            ExecuteRecordset = InsertRecordset(ADOConnection, rstADO, tableToUpdate)
            
        Case ExecuteRecordsetConstant.Update
        
            ExecuteRecordset = UpdateRecordset(ADOConnection, rstADO, tableToUpdate)
    End Select
End Function

Public Function InsertRecordset(ByRef ADOConnection As ADODB.Connection, _
                                ByRef rstADO As ADODB.Recordset, _
                                ByVal tableToUpdate As String) As Long

    Dim objRecord As CubeLibDataSource.DNetRecordset
    Dim success As Integer
    
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
    
    Set objRecord = New CubeLibDataSource.DNetRecordset
    objRecord.InitializeClass rstADO, rstADO.Bookmark
    
    On Error GoTo ErrHandler
    
    databaseName = GetDatabaseName(ADOConnection)
        
    If InStr(databaseName, "mdb_history") > 0 Then
        year = Right(databaseName, 2)
        
        success = g_objDataSource.InsertSadbelHistory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_HISTORY), year)
        
    ElseIf InStr(databaseName, "mdb_repertory") > 0 Then
        If IsNumeric(Right(databaseName, 4)) Then
            year = Right(databaseName, 4)
        End If
        
        success = g_objDataSource.InsertRepertory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_REPERTORY), year)
        
    ElseIf InStr(databaseName, "mdb_EDIhistory") > 0 Then
        If IsNumeric(Right(databaseName, 2)) Then
            year = Right(databaseName, 2)
        End If
        
        success = g_objDataSource.InsertEdifactHistory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_EDI_HISTORY), year)
        
    ElseIf InStr(databaseName, "mdb_sadbel") > 0 Then
        success = g_objDataSource.InsertSadbel(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_SADBEL))
    
    ElseIf InStr(databaseName, "mdb_data") > 0 Then
        success = g_objDataSource.InsertData(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_DATA))
        
    ElseIf InStr(databaseName, "edifact") > 0 Then
        success = g_objDataSource.InsertEdifact(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_EDIFACT))
        
    ElseIf InStr(databaseName, "mdb_scheduler") > 0 Then
        success = g_objDataSource.InsertScheduler(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_SCHEDULER))
        
    ElseIf InStr(databaseName, "TemplateCP") > 0 Then
        success = g_objDataSource.InsertTemplateCP(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_TEMPLATE))
        
    ElseIf InStr(databaseName, "mdb_taric") > 0 Then
        Err.Raise vbObjectError + 603, , "Error in InsertRecordset() - Taric update is not supported yet."
        
    Else
        success = g_objDataSource.InsertOtherDB(objRecord, tableToUpdate, databaseName)
        
    End If
    
    InsertRecordset = success
    
ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            Err.Raise Err.Number, , Err.Description
    End Select
    
End Function

Public Function UpdateRecordset(ByRef ADOConnection As ADODB.Connection, _
                                ByRef rstADO As ADODB.Recordset, _
                                ByVal tableToUpdate As String) As Long
    
    Dim objRecord As CubeLibDataSource.DNetRecordset
    Dim success As Integer
    
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
    
    Set objRecord = New CubeLibDataSource.DNetRecordset
    objRecord.InitializeClass rstADO, rstADO.Bookmark
    
    On Error GoTo ErrHandler
    
    databaseName = GetDatabaseName(ADOConnection)
        
    If InStr(databaseName, "mdb_history") > 0 Then
        year = Right(databaseName, 2)
        
        success = g_objDataSource.UpdateSadbelHistory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_HISTORY), year)
        
    ElseIf InStr(databaseName, "mdb_repertory") > 0 Then
        If IsNumeric(Right(databaseName, 4)) Then
            year = Right(databaseName, 4)
        End If
        
        success = g_objDataSource.UpdateRepertory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_REPERTORY), year)
        
    ElseIf InStr(databaseName, "mdb_EDIhistory") > 0 Then
        If IsNumeric(Right(databaseName, 2)) Then
            year = Right(databaseName, 2)
        End If
        
        success = g_objDataSource.UpdateEdifactHistory(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_EDI_HISTORY), year)
        
    ElseIf InStr(databaseName, "mdb_sadbel") > 0 Then
        success = g_objDataSource.UpdateSadbel(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_SADBEL))
    
    ElseIf InStr(databaseName, "mdb_data") > 0 Then
        success = g_objDataSource.UpdateData(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_DATA))
        
    ElseIf InStr(databaseName, "edifact") > 0 Then
        success = g_objDataSource.UpdateEdifact(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_EDIFACT))
        
    ElseIf InStr(databaseName, "mdb_scheduler") > 0 Then
        success = g_objDataSource.UpdateScheduler(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_SCHEDULER))
        
    ElseIf InStr(databaseName, "TemplateCP") > 0 Then
        success = g_objDataSource.UpdateTemplateCP(objRecord, g_objDataSource.GetEnumFromTableName(tableToUpdate, DBInstanceType_DATABASE_TEMPLATE))
        
    ElseIf InStr(databaseName, "mdb_taric") > 0 Then
        Err.Raise vbObjectError + 601, , "Error in UpdateRecordset() - Taric update is not supported yet."
        
    Else
        success = g_objDataSource.UpdateOtherDB(objRecord, tableToUpdate, databaseName)
    End If
    
    UpdateRecordset = success
    
ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else

            Err.Raise Err.Number, , Err.Description
    End Select
    
End Function

'Hacked the original method to minimize code change
Public Sub ADORecordsetOpen(ByVal Source As String, _
                   ByRef conToUse As ADODB.Connection, _
                   ByRef rstToOpen As ADODB.Recordset, _
                   ByVal CursorType As CursorTypeEnum, _
                   ByVal LockType As LockTypeEnum, _
          Optional ByVal lngCacheSize As Long = 1, _
          Optional ByVal UseDataShaping As Boolean = False)
          
    ADORecordsetOpen_F Source, conToUse, rstToOpen, CursorType, LockType, lngCacheSize, UseDataShaping
End Sub

Public Sub ADORecordsetClose(rstToClose As ADODB.Recordset)
    ADORecordsetClose_F rstToClose
End Sub

Public Function ProcessQuotes(ByVal QuotedText As String, _
                                Optional ByVal ASCQuoteEnclosingData As Long = 39) As String
    Dim intFieldPosition As Integer
    Dim intLength As Integer
    Dim intLinkCtr As Integer
    Dim intLoopCtr As Integer
    
    Dim arrLink
    Dim blnDone As Boolean
    
    Dim strChain As String
    
    
    ReDim arrLink(0)
    blnDone = False
    intLinkCtr = 0
    strChain = ""
    
    Do While blnDone = False
        intLength = Len(QuotedText)
        
        If InStr(1, QuotedText, Chr(39)) > 0 Then    '-----> Apostrophe
            intFieldPosition = InStr(1, QuotedText, Chr(39))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            If Trim$(Chr(ASCQuoteEnclosingData)) = Chr(39) Then
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition) & Chr(39)
            Else
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition)
            End If
            
            If intFieldPosition <> Len(QuotedText) Then
                QuotedText = Mid(QuotedText, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        ElseIf InStr(1, QuotedText, Chr(34)) > 0 Then    '-----> Quote
            intFieldPosition = InStr(1, QuotedText, Chr(34))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            If Trim$(Chr(ASCQuoteEnclosingData)) = Chr(34) Then
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition) & Chr(34)
            Else
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition)
            End If

            
            If intFieldPosition <> Len(QuotedText) Then
                QuotedText = Mid(QuotedText, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        Else
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            arrLink(intLinkCtr - 1) = QuotedText
            blnDone = True
        End If
    Loop
                        
    For intLoopCtr = 0 To (UBound(arrLink) - 1)
        If Trim(strChain) = "" Then
            strChain = CStr(arrLink(intLoopCtr))
        Else
            strChain = strChain & CStr(arrLink(intLoopCtr))
        End If
    Next
    ProcessQuotes = Replace(strChain, " ", Chr(32))
    
End Function

Public Function ADOConnectDB(ByRef ADOConnection As ADODB.Connection, _
                             ByRef DataSourceProperties As CDataSourceProperties, _
                             ByVal InitialCatalog As DBInstanceType, _
                    Optional ByVal InitialCatalogYear As String = vbNullString, _
                    Optional ByVal InitialCatalogName As String = vbNullString, _
                    Optional ByVal UseDataShaping As Boolean = False, _
                    Optional ByVal OpenExclusive As Boolean = False, _
                    Optional ByVal AltInitialCatlogPathMSAccess As String = vbNullString) As CErrObject
            
    Set ADOConnectDB = ADOConnectDB_F(ADOConnection, _
                                      DataSourceProperties, _
                                      InitialCatalog, _
                                      InitialCatalogYear, _
                                      InitialCatalogName, _
                                      UseDataShaping, _
                                      OpenExclusive)
End Function

Public Sub ADOCloseOpenDB(ByRef ADOConnection As ADODB.Connection)
    If (ADOConnection Is Nothing = False) Then
        If ADOConnection.State = ADODB.ObjectStateEnum.adStateOpen Then
            ADOConnection.Close
        End If
        ADOConnection.Open
    End If
End Sub

Public Sub ADODisconnectDB(ByRef ConToClose As ADODB.Connection)
    ADODisconnectDB_F ConToClose
End Sub


Public Function FIsEmpty(ByVal Data As Variant) As Boolean
    Dim strDummy As String
    
    FIsEmpty = False
    
    If IsObject(Data) And Not TypeName(Data) = "Field" Then ' Check if Variable Passed is an Object Variable
        FIsEmpty = True
        
    ElseIf IsEmpty(Data) Then ' Check if Variable Passed is Not Initializaed
        FIsEmpty = True
        
    ElseIf IsNull(Data) Then ' Check if Variable Passed Contains Invalid Data
        FIsEmpty = True

    Else
        If IsArray(Data) Then
            If UBound(Data) = 0 And (Data(0) = "" Or IsEmpty(Data(0))) Then
                FIsEmpty = True
            End If
        Else
            strDummy = CStr(Data)
            
            If Trim(strDummy) = "" Then
                FIsEmpty = True
            End If
        End If
    End If
End Function

Public Function FNullField(ByRef Data As Variant) As Variant
    Dim strDataType As String
    
    If FIsEmpty(Data) Then
        strDataType = TypeName(Data)
        
        If InStr(1, strDataType, "Byte") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Integer") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Long") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Single") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "String") > 0 Then
            Data = ""
        ElseIf InStr(1, strDataType, "Double") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Currency") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Decimal") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Date") > 0 Then
            Data = #1/1/1900#
        ElseIf Trim(strDataType) = "Null" Then
            Data = ""
        ElseIf InStr(1, strDataType, "Field") > 0 Then
        
            Select Case Data.Type
                ' dbBoolean = 1
                ' adBoolean = 11
                Case adBoolean, 1                                       ' Boolean
                    Data = False
                
                                                                        
                ' adBSTR = 8                                            ' String
                ' dbText = 10
                ' dbMemo = 12
                ' adChar = 129
                ' adWChar = 130
                ' adVarChar = 200
                ' adLongVarChar = 201
                ' adVarWChar = 202
                ' adLongVarWChar= 203
                ' adVarBinary = 204
                Case adChar, adLongVarChar, adVarBinary, adVarChar, _
                        adVarWChar, adWChar, 10, 12, adBSTR
                        
                    Data = ""

                Case adDate, adDBDate                                   ' Date
                
                    Data = #1/1/1900#
                
                ' adDBFileTime = 137
                Case adDBTime, adFileTime, 137, adDBTimeStamp  ' Time
                
                    Data = #12:00:00 AM#
                                                                        'Numeric Field Types
                ' adSmallInt = 2
                ' adInteger = 3
                ' adSingle = 4
                ' adDouble = 5
                ' adCurrency = 6
                ' adIDispatch = 9
                ' adDecimal = 14
                ' adTinyInt= 16
                ' adUnsignedTinyInt = 17
                ' adUnsignedSmallInt = 18
                ' adUnsignedInt = 19
                ' adBigInt = 20
                ' adUnsignedBigInt = 21
                ' adBinary = 128
                ' adNumeric = 131
                ' adChapter = 136
                ' adPropVariant= 138
                ' adLongVarBinary = 205
                Case adUnsignedInt, adUnsignedSmallInt, adInteger, adUnsignedTinyInt, _
                        adSingle, adDouble, adCurrency, adNumeric, adDecimal, _
                        adLongVarBinary, adIDispatch, adPropVariant, adChapter, _
                        adBinary, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt
                                                                        
                    Data = 0
                    
                Case Else
                
                    Data = ""
            End Select
        End If
    End If
    FNullField = Data
End Function

Public Function GetADOConnectionStringPropertyFromADODB(ByRef DBConnection As ADODB.Connection, _
                                                        ByVal ADOConnectionStringProperty As ADOConnectionStringPropertyConstant)
    
    GetADOConnectionStringPropertyFromADODB = GetADOConnectionStringProperty(DBConnection.ConnectionString, ADOConnectionStringProperty)

End Function


Public Function GetADOConnectionStringProperty(ByVal ConnectionString As String, _
                                               ByVal ADOConnectionStringProperty As ADOConnectionStringPropertyConstant)

    GetADOConnectionStringProperty = GetADOConnectionStringProperty_F(ConnectionString, ADOConnectionStringProperty)
    
End Function

Public Sub SetMSAccessDBPassword(ByRef DataSourceProperties As CDataSourceProperties, _
                                 ByVal NewPassword As String, _
                                 ByVal InitialCatalog As DBInstanceType, _
                        Optional ByVal InitialCatalogYear As String = vbNullString, _
                        Optional ByVal InitialCatalogName As String = vbNullString)
                        
    Dim objConnection As ADODB.Connection
    
    ADOConnectDB_F objConnection, DataSourceProperties, InitialCatalog, InitialCatalogYear, InitialCatalogName, False, True
    
    If LenB(Trim$(NewPassword)) > 0 Then
        objConnection.Execute "ALTER Database Password " & NewPassword & "``"
    Else
        objConnection.Execute "ALTER DATABASE PASSWORD NULL [" & DataSourceProperties.Password & "]"
    End If
    
    ADODisconnectDB_F objConnection
End Sub

Public Function GetADOConnectionString(ByRef DataSourceProperties As CDataSourceProperties, _
                                       ByVal InitialCatalog As DBInstanceType, _
                              Optional ByVal InitialCatalogYear As String = vbNullString, _
                              Optional ByVal InitialCatalogName As String = vbNullString, _
                              Optional ByVal OpenExclusive As Boolean = False, _
                              Optional ByVal AltInitialCatlogPathMSAccess As String = vbNullString) _
                                       As String
    
    GetADOConnectionString = GetADOConnectionString_F(DataSourceProperties, _
                                                      InitialCatalog, _
                                                      InitialCatalogYear, _
                                                      InitialCatalogName, _
                                                      OpenExclusive, _
                                                      AltInitialCatlogPathMSAccess)
End Function
