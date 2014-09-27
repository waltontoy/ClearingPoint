VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPersistence"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum PersistencePropertyConstant
    [Persistence MdbPath]
    [Persistence Outputfilename]
End Enum

Public Function GetPersistenceFilePath(ByRef OwnerForm As Object, _
                                       ByRef AppObject As Object) As String

    Dim strPersistenceFilePath As String
    Dim strBrowseFilePath As String
    Dim strMsgBox As String

    Dim objDataSource As CDatasource
    Dim objBrowse As CBrowse

    strMsgBox = vbNullString
    strMsgBox = strMsgBox & "The program " & AppObject.Title & " needs to find the "
    strMsgBox = strMsgBox & "file " & G_CONST_PERSISTENCE_FILE & " to continue to start up. "
    strMsgBox = strMsgBox & "Would you like to retry finding the file?"

    strBrowseFilePath = GetSetting(G_CONST_CP_APP_NAME, "Settings", "PersistenceFilePath", AppObject.Path)

    If LenB(Trim$(NoBackSlash(strBrowseFilePath) & "\" & G_CONST_PERSISTENCE_FILE)) > 0 Then
        strPersistenceFilePath = strBrowseFilePath
    Else
        Set objBrowse = New CBrowse

        Do
            strBrowseFilePath = objBrowse.BrowseFile(OwnerForm, AppObject, strBrowseFilePath, G_CONST_PERSISTENCE_FILE)

            If LenB(Trim$(NoBackSlash(strBrowseFilePath) & "\" & G_CONST_PERSISTENCE_FILE)) <= 0 Then

                If MsgBox(strMsgBox, vbRetryCancel) = vbCancel Then
                    strBrowseFilePath = vbNullString

                    Exit Do
                End If

            End If


        Loop While (LenB(Trim$(NoBackSlash(strBrowseFilePath) & "\" & G_CONST_PERSISTENCE_FILE)) <= 0)

        Set objBrowse = Nothing
    End If

    ' Update Regsitry if a valid persistence file path
    If LenB(Trim$(strBrowseFilePath)) > 0 Then
        SaveSetting G_CONST_CP_APP_NAME, "settings", "PersistenceFilePath", strBrowseFilePath

        ' Set the persistence file path for CDataSource

        Set objDataSource = New CDatasource

        objDataSource.SetPersistencePath strBrowseFilePath

        objDataSource.Dispose
        Set objDataSource = Nothing
    End If

    GetPersistenceFilePath = strBrowseFilePath
End Function

Private Function GetPersistencePropertyKey(ByVal PersistenceProperty As PersistencePropertyConstant) As String

    Select Case PersistenceProperty
        Case PersistencePropertyConstant.[Persistence MdbPath]
        
            GetPersistencePropertyKey = "MdbPath"
            
        Case PersistencePropertyConstant.[Persistence Outputfilename]
        
            GetPersistencePropertyKey = "OutputFileName"
            
    End Select
    
End Function

Public Sub UpdatePersistenceFile(ByVal PersistenceFilePath As String, _
                                 ByVal PersistenceProperty As PersistencePropertyConstant, _
                                 ByVal NewValue As String)
    
    Dim strUpdatedMdbPath As String
        
    Dim lngFreeFile As Long
    Dim strFileString As String
    Dim strLineInput As String
    Dim arrLine() As String

    strFileString = vbNullString
    
    lngFreeFile = FreeFile()
    Open NoBackSlash(PersistenceFilePath) & "\" & G_CONST_PERSISTENCE_FILE For Input As #lngFreeFile

    Do While Not EOF(lngFreeFile)
        Line Input #lngFreeFile, strLineInput
        
        If Left$(strLineInput, 1) = "#" Or _
           LenB(Trim$(strLineInput)) <= 0 Then
           
           strFileString = strFileString & strLineInput
           
        Else
            If InStr(1, strLineInput, "=") > 0 Then
                arrLine = Split(strLineInput, "=")
                
                If UCase$(Trim$(arrLine(0))) = UCase$(Trim$(GetPersistencePropertyKey(PersistenceProperty))) Then
                
                    strFileString = strFileString & arrLine(0) & "=" & NewValue & vbCrLf
                Else
                    strFileString = strFileString & strLineInput
                End If
            Else
                strFileString = strFileString & strLineInput
            End If
        End If
    Loop
    Close #lngFreeFile
    
    lngFreeFile = FreeFile()
    
    Open NoBackSlash(PersistenceFilePath) & "\" & G_CONST_PERSISTENCE_FILE For Output As #lngFreeFile
    
    Print #lngFreeFile, strFileString

    Close #lngFreeFile
End Sub

Public Function GetDataSourceProperties(ByVal PersistencePath As String, _
                                        ByVal TracefileType As TracefileTypeConsant) As CDataSourceProperties

    Dim objDataSourceProperties As CDataSourceProperties
    
    Dim enuDatabaseType As CubeLibDataSource.DatabaseType
    Dim strDataSource As String
    Dim strUserName As String
    Dim strPassword As String
    Dim strOutputFilePath As String
    Dim strInitialCatalog As String
    Dim strInitialCatalogPath As String
    
    enuDatabaseType = DatabaseType.DatabaseType_ACCESS2003
    strDataSource = vbNullString
    strUserName = vbNullString
    strPassword = vbNullString
    strOutputFilePath = vbNullString
    
    g_blnNewPersistencePath = (g_strPersistencePath <> PersistencePath)
    g_strPersistencePath = PersistencePath
    
    Dim objDataSource As CubeLibDataSource.CDatasource
    Dim strPersistencePath As String
    
    Set objDataSource = New CubeLibDataSource.CDatasource
    
    If PathFileExists(NoBackSlash(g_strPersistencePath) & "\" & G_CONST_PERSISTENCE_FILE) Then
            
        Set objDataSourceProperties = New CDataSourceProperties
        
        objDataSource.SetPersistencePath NoBackSlash(g_strPersistencePath)
        
        enuDatabaseType = objDataSource.DatabaseType
        
        Select Case objDataSource.DatabaseType
            Case DatabaseType_ACCESS97, _
                DatabaseType_ACCESS2003
                
                strDataSource = objDataSource.DatabasePathFromPersistence
                
            Case DatabaseType_SQLSERVER
                strDataSource = objDataSource.ServerName
            Case Else
                Debug.Assert False
        End Select
        
        strUserName = objDataSource.Username
        strPassword = objDataSource.DatabasePassword
        strOutputFilePath = objDataSource.OutputFilePath
        
        With objDataSourceProperties
            .DatabaseType = enuDatabaseType
            .DataSource = strDataSource
            .OutputFilePath = strOutputFilePath
            .Username = strUserName
            .Password = strPassword
            .TracefileType = TracefileType
            
            ' TO DO FOR CP.NET
            .InitialCatalog = vbNullString
            
            Select Case objDataSource.DatabaseType
                Case DatabaseType_ACCESS97, _
                    DatabaseType_ACCESS2003
                    
                    .InitialCatalogPath = strDataSource
                
                Case DatabaseType_SQLSERVER
                    .InitialCatalogPath = vbNullString
            End Select
            
        End With
    Else
        g_strPersistencePath = vbNullString
    End If
    
    Set GetDataSourceProperties = objDataSourceProperties
    
    Set objDataSource = Nothing
    Set objDataSourceProperties = Nothing
End Function


