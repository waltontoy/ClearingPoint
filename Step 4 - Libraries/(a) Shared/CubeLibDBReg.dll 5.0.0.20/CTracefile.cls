VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTracefile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum TracefileTypeConsant
    [CPReports]
    [CPReportsNCTS]
    [RemotePrinter]
    [AutoMail]
    [Scheduler]
    [ClearingPoint]
    [Trace Off]
End Enum

Public Function AddToTrace(ByRef DataSourceProperties As CDataSourceProperties, _
                           ByVal strTraceString As String, _
                  Optional ByVal PrintToAppPath As Boolean = False, _
                  Optional ByRef FirstRun As Boolean = False) As String
          
          Dim intFreeFile As Integer
          Dim strTracefilePath As String
          Dim strTracefileBaseName As String
          Dim blnStartNewTraceFile As Boolean
          Dim strErrorInfo As String
          Dim enuTracefileType As TracefileTypeConsant
          Dim strComputerName As String
          Dim strUserName As String
          
10        On Error GoTo ErrHandler
          
          enuTracefileType = DataSourceProperties.TracefileType
          
20        strErrorInfo = vbNullString
          
30        g_blnIsSavingToTracefile = True
          
40        blnStartNewTraceFile = False
          
50        If PrintToAppPath Then
60            strTracefilePath = GetThreadEXEName
              
              ' Strip Path Name
70            strTracefilePath = Left$(strTracefilePath, InStrRev(strTracefilePath, "\") - 1)
80        Else
90            strTracefilePath = DataSourceProperties.OutputFilePath
100       End If
          
          ' Subfolder by Application Name
101       CreateDirectory NoBackSlash(strTracefilePath) & "\" & GetTracefileFolder(enuTracefileType)
          
102       strTracefilePath = NoBackSlash(strTracefilePath) & "\" & GetTracefileFolder(enuTracefileType)
          
          ' Subfolder by Computer Name
          strComputerName = Environ("COMPUTERNAME")
          
103       CreateDirectory NoBackSlash(strTracefilePath) & "\" & strComputerName
          
104       strTracefilePath = NoBackSlash(strTracefilePath) & "\" & strComputerName

110       strTracefileBaseName = GetTracefileBaseName(enuTracefileType)
          
120       intFreeFile = FreeFile()
          
130       If Len(Dir(NoBackSlash(strTracefilePath) & "\" & strTracefileBaseName & ".txt")) Then
140           If FileLen(NoBackSlash(strTracefilePath) & "\CPReportsTraceFile.txt") >= 360000 Then

150               Name NoBackSlash(strTracefilePath) & "\" & strTracefileBaseName & ".txt" As _
                       NoBackSlash(strTracefilePath) & "\" & GetRandomTraceFileName(enuTracefileType)
                  
160               Open NoBackSlash(strTracefilePath) & "\" & strTracefileBaseName & ".txt" For Output As #intFreeFile
                  
170               blnStartNewTraceFile = True
180           Else
190               Open NoBackSlash(strTracefilePath) & "\" & strTracefileBaseName & ".txt" For Append As #intFreeFile
200           End If
210       Else
220           Open NoBackSlash(strTracefilePath) & "\" & strTracefileBaseName & ".txt" For Output As #intFreeFile
              
230           blnStartNewTraceFile = True
240       End If
          
250       If IsMissing(FirstRun) Then
              
260           If blnStartNewTraceFile Then
270               Print #intFreeFile, AdditionalTracefileInfo(DataSourceProperties, enuTracefileType)
280           End If
              
290       ElseIf FirstRun Or blnStartNewTraceFile Then
300           FirstRun = False

310           Print #intFreeFile, AdditionalTracefileInfo(DataSourceProperties, enuTracefileType)
320       End If
          
321       If enuTracefileType = TracefileTypeConsant.ClearingPoint Then
322            Print #intFreeFile, Now & " " & strComputerName & "(" & strUserName & ") " & strTraceString
323       Else
324            Print #intFreeFile, strTraceString
325       End If
          
340       Close #intFreeFile

ErrHandler:
          
350       g_blnIsSavingToTracefile = False
          
360       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
370           Case Else
380               strErrorInfo = "Error in PCubeLibDBReg.CTracefile.AddToTrace: " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
390       End Select
          
400       AddToTrace = strErrorInfo
End Function

Public Function IsSavingToTracefile() As Boolean
    IsSavingToTracefile = g_blnIsSavingToTracefile
End Function

Private Function GetTracefileFolder(ByVal TracefileType As TracefileTypeConsant) As String

    Select Case TracefileType
        Case TracefileTypeConsant.CPReports
            GetTracefileFolder = "CP Reports"
        
        Case TracefileTypeConsant.CPReportsNCTS
            GetTracefileFolder = "CP Reports NCTS"
        
        Case TracefileTypeConsant.RemotePrinter
            GetTracefileFolder = "CP Remote Printer"
        
        Case TracefileTypeConsant.AutoMail
            GetTracefileFolder = "CP Reports AutoMail"
        
        Case TracefileTypeConsant.Scheduler
            GetTracefileFolder = "CP Task Scheduler"
        
        Case TracefileTypeConsant.Scheduler
            GetTracefileFolder = "ClearingPoint"
            
        Case Else
            Debug.Assert False
            GetTracefileFolder = ""
    End Select
End Function


Private Function GetTracefileBaseName(ByVal TracefileType As TracefileTypeConsant) As String

    Select Case TracefileType
        Case TracefileTypeConsant.CPReports
            GetTracefileBaseName = "CPReportsTraceFile"
        
        Case TracefileTypeConsant.CPReportsNCTS
            GetTracefileBaseName = "CPReportNCTSTraceFile"
        
        Case TracefileTypeConsant.RemotePrinter
            GetTracefileBaseName = "CPRPTraceFile"
        
        Case TracefileTypeConsant.AutoMail
            GetTracefileBaseName = "CPMailTraceFile"
        
        Case TracefileTypeConsant.Scheduler
            GetTracefileBaseName = "CPTSTraceFile"
        
        Case TracefileTypeConsant.Scheduler
            GetTracefileBaseName = "CPTraceFile"
            
        Case Else
            Debug.Assert False
            GetTracefileBaseName = "UnknownTracefileName"
    End Select
End Function

Private Function GetRandomTraceFileName(ByVal TracefileType As TracefileTypeConsant) As String
    
    Dim lngCtr As Long
    Dim strRandomNumber As String
    Dim strRandomFileName As String

    ' CSCLP-910 - Start
    strRandomFileName = GetTracefileBaseName(TracefileType) & Format(Now, "ddMMyyyyhhmm")
    ' CSCLP-910 - End
    
Repeat_Generator:
    Randomize
    
    For lngCtr = 1 To 4
        strRandomNumber = strRandomNumber & Int((9 * Rnd) + 0)
    Next lngCtr
    
    strRandomFileName = strRandomFileName & "_" & strRandomNumber & ".txt"
    
    If (Len(Dir(strRandomFileName)) > 0) Then
        GoTo Repeat_Generator
    End If
    
    GetRandomTraceFileName = strRandomFileName
    
End Function

Private Function AdditionalTracefileInfo(ByRef DataSourceProperties As CDataSourceProperties, _
                                         ByVal TracefileType As TracefileTypeConsant) As String
    
    Dim strInfo As String
    Dim strAppExeName As String
    Dim objFileInfo As CFileInfo
    Dim strSystemPath As String
    
    strSystemPath = GetSystemPath
    strAppExeName = GetThreadEXEName

    Set objFileInfo = GetFileInformation(strAppExeName)
    
    strInfo = vbNullString
    
    strInfo = strInfo & Now & ", Initialize Tracefile..." & vbCrLf
    strInfo = strInfo & Now & ", Get " & objFileInfo.ProductName & " version:" & vbCrLf
    strInfo = strInfo & Now & ", " & objFileInfo.ProductName & String(3, vbTab) & _
                                     objFileInfo.VerMajor & "." & objFileInfo.VerMinor & "." & _
                                     objFileInfo.VerNotUsedVB & "." & objFileInfo.VerRevision & _
                                     vbCrLf & vbCrLf
    
    Select Case TracefileType
        Case TracefileTypeConsant.RemotePrinter
            strInfo = strInfo & Now & ", Get CubeFTP.dll version:" & vbCrLf
            
            Set objFileInfo = GetFileInformation(NoBackSlash(strSystemPath) & "\CubeFTP.dll")
            strInfo = strInfo & Now & ", " & objFileInfo.ProductName & String(3, vbTab) & _
                                         objFileInfo.VerMajor & "." & objFileInfo.VerMinor & "." & _
                                         objFileInfo.VerNotUsedVB & "." & objFileInfo.VerRevision & _
                                         vbCrLf & vbCrLf
                                         
            strInfo = strInfo & Now & ", Get CubeFTPBrowse.dll version:" & vbCrLf
            Set objFileInfo = GetFileInformation(NoBackSlash(strSystemPath) & "\CubeFTPBrowse.dll")
            strInfo = strInfo & Now & ", " & objFileInfo.ProductName & String(3, vbTab) & _
                                         objFileInfo.VerMajor & "." & objFileInfo.VerMinor & "." & _
                                         objFileInfo.VerNotUsedVB & "." & objFileInfo.VerRevision & _
                                         vbCrLf & vbCrLf
                                         
        Case TracefileTypeConsant.Scheduler
        
            strInfo = strInfo & AdditionalTracefileInfoCPTS(DataSourceProperties)
    End Select
        
    Set objFileInfo = Nothing
    
    AdditionalTracefileInfo = strInfo
    
End Function

Public Function AdditionalTracefileInfoCPTS(ByRef DataSourceProperties As CDataSourceProperties) As String
    Dim strTraceString As String
    Dim lngFreeFile As Long
    Dim strBrowseFolder As String
    Dim blnBrowseFolder As Boolean
    Dim blnFilesAreComplete As Boolean
    Dim blnInitialLoop As Boolean
    
    Dim lngFileAttributes As Long
    Dim strMDBFilePath As String
    Dim strMDBFileName As String
    Dim strReturnConnection As String
    
    Dim strPathFolders() As String
    
    Dim blnPathExists As Boolean
    
    Dim rstReceivingCycles As ADODB.Recordset
    Dim conADOConnection As ADODB.Connection
    Dim clsRegistry As CRegistry
    
    On Error GoTo ErrHandler
    
    Set clsRegistry = New CRegistry
    
    strTraceString = vbNullString
    
    ' GET OS VERSION
    strTraceString = strTraceString & Now & ", Get Operating System Version: " & vbCrLf
    strTraceString = strTraceString & Now & ", " & GetOSVersion & vbCrLf & vbCrLf
    
    ' STEP 1: Check if with Admin Rights to Registry
    PROC_CheckRegistryRights DataSourceProperties, strTraceString
    
    ' STEP 2: Check Registry Values
    PROC_CheckDBPathRegistryValues DataSourceProperties, strTraceString
    
    ' STEP 3: Check Available Drives
    PROC_CheckAvailableLogicalDrives DataSourceProperties, strTraceString
    
    ' STEP 4: Check Network
    PROC_CheckNetwork DataSourceProperties, strTraceString
    
    ' STEP 5: Check Folder Path
    If Not PROC_CheckFolderPath(DataSourceProperties, strTraceString) Then
    
        ' Folder Path is read only
        blnBrowseFolder = True
  
        ' Just to make sure
        blnFilesAreComplete = False
    Else
        ' STEP 6:   Check if files exist
        blnFilesAreComplete = PROC_FilesExist(DataSourceProperties, strTraceString, _
                                              "TemplateCP.mdb", _
                                              "mdb_data.mdb", _
                                              "mdb_sadbel.mdb", _
                                              "mdb_taric.mdb", _
                                              "mdb_scheduler.mdb", _
                                              "edifact.mdb")
  
        If blnFilesAreComplete Then
                    
            ' Step 7: Test Database Connections
            ' Connect to TemplateCP.mdb
            IsADOConnectionOK DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE, strTraceString
            
            ' Connect to mdb_data.mdb
            IsADOConnectionOK DataSourceProperties, DBInstanceType_DATABASE_DATA, strTraceString
            
            ' Connect to mdb_sadbel.mdb
            IsADOConnectionOK DataSourceProperties, DBInstanceType_DATABASE_SADBEL, strTraceString
                    
            ' Connect to mdb_taric.mdb
            IsADOConnectionOK DataSourceProperties, DBInstanceType_DATABASE_TARIC, strTraceString
                    
            ' Connect to EDIFACT.mdb
            IsADOConnectionOK DataSourceProperties, DBInstanceType_DATABASE_EDIFACT, strTraceString
            
            ' Connect to mdb_scheduler.mdb
            If IsADOConnectionOK(DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER, strTraceString) Then
            
                    ' GET POLL CYCLES
                    ADOConnectDB_F conADOConnection, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
                    ADORecordsetOpen_F GetSQLCommandFromTableName_F("ReceivingCycles"), conADOConnection, rstReceivingCycles, adOpenKeyset, adLockOptimistic
                    
                    strTraceString = strTraceString & Now & ", Get Receiving Cycles Properties:" & vbCrLf
                    
                    strTraceString = strTraceString & "C.Sys" & vbTab & _
                                                      "S.PollCyc" & vbTab & _
                                                      "F.PollCyc" & vbTab & _
                                                      "F.PollIncr" & vbTab & _
                                                      "F.PollCycTyp" & vbTab & _
                                                      "F.PollInitVal" & vbCrLf
                                                          
                    If Not (rstReceivingCycles.EOF And rstReceivingCycles.BOF) Then
                        rstReceivingCycles.MoveFirst
                        
                        Do While rstReceivingCycles.EOF = False
                            strTraceString = strTraceString & rstReceivingCycles(1) & vbTab & _
                                                              rstReceivingCycles(2) & vbTab & vbTab & _
                                                              rstReceivingCycles(3) & vbTab & vbTab & _
                                                              rstReceivingCycles(4) & vbTab & vbTab & _
                                                              rstReceivingCycles(5) & vbTab & vbTab & vbTab & _
                                                              rstReceivingCycles(6) & vbCrLf
                            rstReceivingCycles.MoveNext
                        Loop
                        
                        strTraceString = strTraceString & vbCrLf & vbCrLf
                    Else
                        strTraceString = strTraceString & "Receiving Cycles Properties not found." & vbCrLf & vbCrLf
                    End If
                    ADORecordsetClose_F rstReceivingCycles
            End If
        End If
    End If
    
    
    
    'Get DLL Versions
    PROC_GetDllVersions DataSourceProperties, strTraceString
    
    ' For less user intervention, instead of displaying text on a form, and letting users copy and paste it
    ' to an e-mail, save it to a text file on the same folder as this application, and let users use this
    ' as an e-mail attachment.
    ' FInformation.ShowForm strTraceString
    'lngFreeFile = FreeFile()
    'Open App.Path & "\ConnTest.txt" For Output Lock Read Write As #lngFreeFile
    'Print #lngFreeFile, strTraceString
    'Close #lngFreeFile

ErrHandler:
        
    AdditionalTracefileInfoCPTS = strTraceString
    
    Set clsRegistry = Nothing
    
    Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
        Case 0                ' No error; normal exit.
            ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
        Case Else
            AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in AdditionalTracefileInfo, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
    End Select
End Function

Private Sub PROC_CheckRegistryRights(ByRef DataSourceProperties As CDataSourceProperties, _
                                     ByRef strTracefileString As String)
          Dim blnUserHasAdminRights As Boolean
          
10        On Error GoTo ErrHandler
          
          ' Check if with Admin Rights to Registry
20        strTracefileString = strTracefileString & "STEP 1 (Check Registry User Rights) : " & vbCrLf & vbCrLf
          
30        Call HasRegistryAdminRights_F(cpiLocalMachine, "Software\Microsoft\Windows", "CurrentVersion", "ProgramFilesDir", blnUserHasAdminRights)
          
40        If Not blnUserHasAdminRights Then
50            strTracefileString = strTracefileString & vbTab & "User has no admin rights to modify the Local Machine registry." & vbCrLf
60        Else
70            strTracefileString = strTracefileString & vbTab & "User has admin rights to modify the Local Machine registry." & vbCrLf
80        End If
          
90        Call HasRegistryAdminRights_F(cpiCurrentUser, "Software\Microsoft\Windows", "CurrentVersion", "ProgramFilesDir", blnUserHasAdminRights)
          
100       If Not blnUserHasAdminRights Then
110           strTracefileString = strTracefileString & vbTab & "User has no admin rights to modify the Current User registry." & vbCrLf & vbCrLf
120       Else
130           strTracefileString = strTracefileString & vbTab & "User has admin rights to modify the Current User registry." & vbCrLf & vbCrLf
140       End If
          
ErrHandler:
          
150       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
160           Case Else
170               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_CheckRegistryRights, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
180       End Select
End Sub

Private Sub PROC_CheckDBPathRegistryValues(ByRef DataSourceProperties As CDataSourceProperties, _
                                           ByRef strTracefileString As String)
          Dim strRegistryValue As String
          Dim clsRegistry As CRegistry

10        On Error GoTo ErrHandler
          
20        Set clsRegistry = New CRegistry
          
30        strTracefileString = strTracefileString & "STEP 2 (Check Registry Settings) : " & vbCrLf & vbCrLf
          
          ' Read Local Machine Registry - MdbPath
40        If clsRegistry.GetRegistry(cpiLocalMachine, G_CONST_CP_APP_NAME, "settings", "MdbPath") Then
50            strRegistryValue = clsRegistry.RegistryValue
              
60            If Trim(strRegistryValue) <> "" Then
70                strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Local Machine\Software\Cubepoint\ClearingPoint\Settings  [MdbPath, '" & strRegistryValue & "']" & vbCrLf
80            Else
90                strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Local Machine\Software\Cubepoint\ClearingPoint\Settings  [MdbPath, '" & strRegistryValue & "']" & vbCrLf
100           End If
110       Else
120           strTracefileString = strTracefileString & vbTab & "Registry Read Fail" & vbTab & "Local Machine\Software\Cubepoint\ClearingPoint\Settings  [MdbPath, Uhnknown]" & vbCrLf
130       End If
              
          ' Read Current User Registry - MdbPathDefault
140       If clsRegistry.GetRegistry(cpiCurrentUser, G_CONST_CP_APP_NAME, "Software\VB and VBA Program Settings\ClearingPoint\Settings", "MdbPathDefault", cpiUserDefined) Then
150           strRegistryValue = clsRegistry.RegistryValue
              
160           If Trim(strRegistryValue) <> "" Then
                  'strDBPathToUse = strRegistryValue
170               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathDefault, '" & strRegistryValue & "']" & vbCrLf
180           Else
190               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathDefault, '']" & vbCrLf
200           End If
210       Else
220           strTracefileString = strTracefileString & vbTab & "Registry Read Fail" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathDefault, Uhnknown]" & vbCrLf
230       End If
          
          ' Read Current User Registry - MdbPath
240       If clsRegistry.GetRegistry(cpiCurrentUser, G_CONST_CP_APP_NAME, "Software\VB and VBA Program Settings\ClearingPoint\Settings", "MdbPath", cpiUserDefined) Then
250           strRegistryValue = clsRegistry.RegistryValue
              
260           If strRegistryValue <> "" Then
270               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPath, '" & strRegistryValue & "']" & vbCrLf
280           Else
290               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPath, '']" & vbCrLf
300           End If
310       Else
320           strTracefileString = strTracefileString & vbTab & "Registry Read Fail" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPath, Uhnknown]" & vbCrLf
330       End If
          
          ' Read Current User Registry - MdbPathParameter
340       If clsRegistry.GetRegistry(cpiCurrentUser, G_CONST_CP_APP_NAME, "Software\VB and VBA Program Settings\ClearingPoint\Settings", "MdbPathParameter", cpiUserDefined) Then
350           strRegistryValue = clsRegistry.RegistryValue
              
360           If strRegistryValue <> "" Then
370               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathParameter, '" & strRegistryValue & "']" & vbCrLf & vbCrLf
380           Else
390               strTracefileString = strTracefileString & vbTab & "Registry Read OK" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathParameter, '']" & vbCrLf & vbCrLf
400           End If
410       Else
420           strTracefileString = strTracefileString & vbTab & "Registry Read Fail" & vbTab & "Current User\Software\VB and VBA Program Settings\ClearingPoint\Settings  [MdbPathParameter, Uhnknown]" & vbCrLf & vbCrLf
430       End If
          
          
ErrHandler:
440       Set clsRegistry = Nothing
          
450       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
460           Case Else
470               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_CheckDBPathRegistryValues, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
480       End Select
End Sub

Private Sub PROC_CheckAvailableLogicalDrives(ByRef DataSourceProperties As CDataSourceProperties, _
                                             ByRef strTracefileString As String)
          Dim arrDrives
          Dim lngDriveCtr As Long
          Dim strTemp As String
          
10        On Error GoTo ErrHandler
          
20        strTracefileString = strTracefileString & "STEP 3 (Check Available Logical Drives) : " & vbCrLf & vbCrLf
          
30        strTemp = GetSpaceSerializedLogicalDrives
          
40        If Trim(strTemp) <> "" Then
50            arrDrives = Split(strTemp, " ")
              
60            For lngDriveCtr = 0 To UBound(arrDrives)
70                strTracefileString = strTracefileString & vbTab & CStr(arrDrives(lngDriveCtr)) & ":\" & vbTab & GetDriveTypeDescription(CStr(arrDrives(lngDriveCtr)) & ":\") & vbCrLf
80            Next
90        Else
100           strTracefileString = strTracefileString & vbTab & "There is no available logical drive." & vbCrLf
110       End If
          
120       strTracefileString = strTracefileString & vbCrLf
          
ErrHandler:
          
130       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
140           Case Else
150               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_CheckAvailableLogicalDrives, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
160       End Select
End Sub

Private Sub PROC_CheckNetwork(ByRef DataSourceProperties As CDataSourceProperties, _
                              ByRef strTracefileString As String)
          Dim lngReturnNetworkConnection As Long
          Dim enuNetworkStatus As NetworkStatusConstant
          
10        On Error GoTo ErrHandler
          
20        strTracefileString = strTracefileString & "STEP 4 (Check Network) : " & vbCrLf & vbCrLf
          
30        enuNetworkStatus = GetNetworkStatus
          
40        Select Case enuNetworkStatus
              Case NetworkStatusConstant.[No Network]
50                strTracefileString = strTracefileString & vbTab & "The local system is not connected to a network." & vbCrLf & vbCrLf
60            Case NetworkStatusConstant.[AOL Network]
70                strTracefileString = strTracefileString & vbTab & "The local system is connected to an AOL network." & vbCrLf & vbCrLf
80            Case NetworkStatusConstant.[LAN Network]
90                strTracefileString = strTracefileString & vbTab & "The local system is connected to a LAN network." & vbCrLf & vbCrLf
100           Case NetworkStatusConstant.[WAN Network]
110               strTracefileString = strTracefileString & vbTab & "The local system is connected to an WAN network." & vbCrLf & vbCrLf
120           Case Else
130               strTracefileString = strTracefileString & vbTab & "The network connection type and status could not be determined." & vbCrLf & vbCrLf
140       End Select

ErrHandler:
          
150       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
160           Case Else
170               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_CheckNetwork, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
180       End Select
End Sub

Private Function PROC_CheckFolderPath(ByRef DataSourceProperties As CDataSourceProperties, _
                                      ByRef strTracefileString As String) As Boolean
          
          Dim blnNetworkPath As Boolean
          Dim blnUNCPath As Boolean
          Dim strLogicalDrive As String
          Dim lngFileAttributes As Long
          Dim strFolderPath As String

10        On Error GoTo ErrHandler
          
20        strFolderPath = NoBackSlash(DataSourceProperties.OutputFilePath)
          
30        strTracefileString = strTracefileString & "STEP 5 (Check Folder Path) : " & vbCrLf & vbTab & strFolderPath & vbCrLf & vbCrLf
          
40        If IsPathADirectory(NoBackSlash(DataSourceProperties.OutputFilePath)) Then
50            strTracefileString = strTracefileString & vbTab & "Folder Path exists." & vbCrLf
              
60            PROC_CheckFolderPath = True
              
              ' Check Directory Attributes
70            If IsFolderAttribute(strFolderPath, FileAttributeConstant.Hidden) Then
80                strTracefileString = strTracefileString & vbTab & "Hidden " & strFolderPath & vbCrLf
90            Else
100               strTracefileString = strTracefileString & vbTab & "Not Hidden " & strFolderPath & vbCrLf
110           End If
              
120           If IsFolderAttribute(strFolderPath, FileAttributeConstant.[Read Only]) Then
130               PROC_CheckFolderPath = False
140               strTracefileString = strTracefileString & vbTab & "Read Only " & strFolderPath & vbCrLf
150           Else
160               PROC_CheckFolderPath = True
170               strTracefileString = strTracefileString & vbTab & "Read/Write " & strFolderPath & vbCrLf
180           End If
190       Else
200           strTracefileString = strTracefileString & vbTab & "Folder Path is does not exist." & vbCrLf
210           PROC_CheckFolderPath = False
220       End If
          
230       If IsPathANetworkPath(strFolderPath) Then
240           blnNetworkPath = True
250           strTracefileString = strTracefileString & vbTab & "Folder Path is a network path." & vbCrLf
              
260           If IsPathAUNCPath(strFolderPath) Then
270               blnUNCPath = True
280               strTracefileString = strTracefileString & vbTab & "Folder Path is a UNC path." & vbCrLf
290           Else
300               blnUNCPath = False
310               strTracefileString = strTracefileString & vbTab & "Folder Path is not a UNC path." & vbCrLf
320           End If
330       Else
340           blnNetworkPath = False
350           strTracefileString = strTracefileString & vbTab & "Folder Path is a local path." & vbCrLf
360       End If
          
370       If Not blnUNCPath Or Not blnNetworkPath Then
380           strLogicalDrive = GetPathLogicalDrive(strFolderPath)
390           strTracefileString = strTracefileString & vbTab & strLogicalDrive & vbTab & GetDriveTypeDescription(strLogicalDrive) & vbCrLf & vbCrLf
400       Else
410           strTracefileString = strTracefileString & vbCrLf
420       End If
          
ErrHandler:
          
430       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
440           Case Else
450               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_CheckFolderPath, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
460       End Select
End Function

Private Function PROC_FilesExist(ByRef DataSourceProperties As CDataSourceProperties, _
                                 ByRef strTracefileString As String, _
                                 ParamArray Files()) As Boolean
          Dim lngFileCtr As Long
          Dim strFolderPath As String
          
10        On Error GoTo ErrHandler
          
20        PROC_FilesExist = True
          
30        strFolderPath = NoBackSlash(DataSourceProperties.OutputFilePath)
          
40        strTracefileString = strTracefileString & "STEP 6 (Check Database Files) : " & vbCrLf & vbTab & strFolderPath & vbCrLf & vbCrLf
          
50        For lngFileCtr = 0 To UBound(Files)
60            If IsFileExisting(strFolderPath & "\" & CStr(Files(lngFileCtr))) Then
70                strTracefileString = strTracefileString & vbTab & strFolderPath & "\" & CStr(Files(lngFileCtr)) & " exists." & vbCrLf
80            Else
90                PROC_FilesExist = False
100               strTracefileString = strTracefileString & vbTab & strFolderPath & "\" & CStr(Files(lngFileCtr)) & " does not exist." & vbCrLf
110           End If
120       Next
          
130       If Not PROC_FilesExist Then
140           strTracefileString = strTracefileString & vbTab & "Database files are incomplete." & vbCrLf
150       End If
          
160       strTracefileString = strTracefileString & vbCrLf
          
ErrHandler:
          
170       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
180           Case Else
190               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_FilesExist, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
200       End Select
End Function

Private Function IsADOConnectionOK(ByRef DataSourceProperties As CDataSourceProperties, _
                                   ByVal DBInstance As DBInstanceType, _
                                   ByRef strTracefileString As String) As Boolean
    Dim objErrObject As CErrObject
    Dim conADO As ADODB.Connection
    Dim rstDatabaseVersion As ADODB.Recordset
    Dim strReturnConnection As String
    Dim lngErrorLine As Long
    Dim objFileInfo As CFileInfo
    Dim strFullFilepath As String
    Dim strConnectionString As String
    
    Set objErrObject = ADOConnectDB_F(conADO, DataSourceProperties, DBInstance)
    
    Select Case objErrObject.Number
        Case 0
            IsADOConnectionOK = True
            
            strReturnConnection = vbNullString
            
        Case Else
        
            
            
            strConnectionString = GetADOConnectionString_F(DataSourceProperties, DBInstance)
          
            If LenB(Trim$(strConnectionString)) <= 0 Then
                strReturnConnection = objErrObject.Source & " (" & objErrObject.Number & ", " & objErrObject.ErrorLine & "): " & objErrObject.Description & vbCrLf & vbTab & "Connection String not built."
            Else
                strReturnConnection = objErrObject.Source & " (" & objErrObject.Number & ", " & objErrObject.ErrorLine & "): " & objErrObject.Description & vbCrLf & vbTab & strConnectionString
            End If
            
            IsADOConnectionOK = False
            
            Exit Function
    End Select
    
    If ADOXIsTableExisting_F(conADO, "DBProperties") Then
        ADORecordsetOpen_F "SELECT [DBProps_Version] FROM [DBProperties] ", conADO, rstDatabaseVersion, adOpenKeyset, adLockOptimistic

        If Not (rstDatabaseVersion.EOF And rstDatabaseVersion.BOF) Then
            rstDatabaseVersion.MoveFirst

            strTracefileString = strTracefileString & vbTab & GetDBInstanceTypeDesc_F(DBInstance, vbNullString, vbNullString) & ".mdb" & String(3, vbTab) & IIf(IsNull(rstDatabaseVersion.Fields("DBProps_Version").Value), "", rstDatabaseVersion.Fields("DBProps_Version").Value) & vbCrLf
        End If
    End If
    ADORecordsetClose_F rstDatabaseVersion
            
    ADODisconnectDB_F conADO
    
    If LenB(Trim$(strReturnConnection)) > 0 Then
        strTracefileString = strTracefileString & vbTab & strReturnConnection & vbCrLf
        'strDBPathToUse = BrowseForFolder(strDBPathToUse)
    Else
        Select Case DataSourceProperties.DatabaseType
            Case DatabaseType_ACCESS97, _
                 DatabaseType_ACCESS2003
                 
                strFullFilepath = NoBackSlash(DataSourceProperties.OutputFilePath) & "\" & GetDBInstanceTypeDesc_F(DBInstance, vbNullString, vbNullString) & ".mdb"
                strTracefileString = strTracefileString & vbTab & "Connected to " & strFullFilepath & vbCrLf
                
                Set objFileInfo = New CFileInfo
                Set objFileInfo = GetFileInformation(strFullFilepath)
                
                If objFileInfo.IsHidden Then
                    strTracefileString = strTracefileString & vbTab & "Hidden " & strFullFilepath & vbCrLf
                Else
                    strTracefileString = strTracefileString & vbTab & "Not Hidden " & strFullFilepath & vbCrLf
                End If
                      
                If objFileInfo.IsReadOnly Then
                    strTracefileString = strTracefileString & vbTab & "Read Only " & strFullFilepath & vbCrLf & vbCrLf
                Else
                    strTracefileString = strTracefileString & vbTab & "Read/Write " & strFullFilepath & vbCrLf & vbCrLf
                End If
          
                Set objFileInfo = Nothing
                
            Case Else
                strTracefileString = strTracefileString & vbTab & "Connected to " & GetDBInstanceTypeDesc_F(DBInstance, vbNullString, vbNullString) & vbCrLf
                
        End Select
    End If
End Function

Private Sub PROC_GetDllVersions(ByRef DataSourceProperties As CDataSourceProperties, _
                                ByRef strTracefileString As String)
          
          Dim strSystemPath As String
          
10        On Error GoTo ERR_HANDLER
          
20        strSystemPath = GetSystemPath
          
30        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib1.dll", strTracefileString
40        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib2.dll", strTracefileString
50        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib3.dll", strTracefileString
60        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib4.dll", strTracefileString
70        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib5.dll", strTracefileString
80        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib6.dll", strTracefileString
90        GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib7.dll", strTracefileString
100       GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib8.dll", strTracefileString
110       GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib9.dll", strTracefileString
120       GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib10.dll", strTracefileString
130       GetDLLVersionInfo DataSourceProperties, strSystemPath, "Cubelib12.dll", strTracefileString
          
140       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB1.dll", strTracefileString
150       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB6.dll", strTracefileString
160       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB7.dll", strTracefileString
170       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB16.dll", strTracefileString
180       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB17.dll", strTracefileString
190       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB21.dll", strTracefileString
200       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB22.dll", strTracefileString
210       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB23.dll", strTracefileString
220       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB24.dll", strTracefileString
230       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB25.dll", strTracefileString
240       GetDLLVersionInfo DataSourceProperties, strSystemPath, "CubeDB26.dll", strTracefileString
          
ERR_HANDLER:
          
          
250       Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
            ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
260           Case Else
270               AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in PROC_GetDllVersions, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
280       End Select
End Sub

Private Sub GetDLLVersionInfo(ByRef DataSourceProperties As CDataSourceProperties, _
                              ByVal DLLPath As String, _
                              ByVal DLLName As String, _
                              ByRef strTracefileString As String)
          
          Dim objFileInfo As CFileInfo
          
10        On Error GoTo ErrHandler
          
20        Set objFileInfo = GetFileInformation(NoBackSlash(DLLPath) & "\" & DLLName)
          
30        strTracefileString = strTracefileString & vbTab & DLLName & String(3, vbTab) & _
                                                            objFileInfo.VerMajor & "." & _
                                                            objFileInfo.VerMinor & "." & _
                                                            objFileInfo.VerNotUsedVB & "." & _
                                                            objFileInfo.VerRevision & vbCrLf
40        Set objFileInfo = Nothing
          
ErrHandler:
          
50        Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
              Case 0                ' No error; normal exit.
                  ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
60            Case Else
70                AddToTrace DataSourceProperties, TracefileTypeConsant.Scheduler, "Error in GetDLLVersionInfo, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
80        End Select
End Sub



