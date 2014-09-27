VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreCheck"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function CheckFiles(ByVal Application As Object, _
                           ByRef OwnerForm As Object, _
                           ByRef DataSourceProperties As CDataSourceProperties, _
                           ParamArray PathArgs()) As Boolean
    
'Public Function CheckFiles(ByVal Application As Object, ByRef OwnerForm As Object, _
                            ByVal MainDBPassword As String, ParamArray PathArgs()) As Boolean
                            
    Dim clsCheckConnection As ADODB.Connection
    Dim clsRegistry As CRegistry
    Dim clsBrowse As CBrowse
    Dim clsLocations As CLocations
    Dim enuResult As CheckResult
    Dim enuProvider As ProviderType
    Dim strYearSuffix As String
    
    Dim arrSubArguments
    Dim lngArgCtr As Long
    Dim strMissingPaths As String
    Dim strRegistryValue As String
    Dim strArguments As String
    
    Dim enuDBInstanceType As DBInstanceType
    
    Const IDX_DATABASE_NAME = 5
    Const IDX_DATABASE_PASSWORD = 6
    
    Set clsLocations = New CLocations
    
    Set clsCheckConnection = New ADODB.Connection
    Set clsRegistry = New CRegistry
    Set clsBrowse = New CBrowse
    
    enuResult = cpiInitialized
    strMissingPaths = ""
    CheckFiles = False
    
    enuResult = cpiInitialized
    strMissingPaths = ""
    CheckFiles = False
    
    
    
    If Application.PrevInstance Then
        If clsRegistry.GetRegistry(cpiCurrentUser, Application.ProductName, "Settings", "RestartInstance") Then
            If Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) = "No" Then
                MsgBox Application.ProductName & " is still running.", vbInformation, "Cubepoint Library"
                Set clsLocations = Nothing
                Set clsCheckConnection = Nothing
                Set clsRegistry = Nothing
                Set clsBrowse = Nothing
                Exit Function
            Else
                clsRegistry.SaveRegistry cpiCurrentUser, Application.ProductName, "Settings", "RestartInstance", Encrypt("No", KEY_ENCRYPT)
            End If
        End If
    End If

    
    ' Check for connectability with required files, folders, and databases
    Do While Not (enuResult = cpiCancel Or enuResult = cpiSuccess)
        strMissingPaths = ""
        strArguments = ""
        
        For lngArgCtr = 0 To UBound(PathArgs)
            arrSubArguments = Split(CStr(PathArgs(lngArgCtr)), "|")
            strArguments = strArguments & PathArgs(lngArgCtr) & ", "
            If UBound(arrSubArguments) >= 6 Then
                clsRegistry.GetRegistry arrSubArguments(6), Application.ProductName, CStr(arrSubArguments(1)), CStr(arrSubArguments(2))
            Else
                clsRegistry.GetRegistry cpiLocalMachine, Application.ProductName, CStr(arrSubArguments(1)), CStr(arrSubArguments(2))
            End If
            strRegistryValue = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
            
            Select Case UCase(CStr(arrSubArguments(0)))
                Case "P"
                    strMissingPaths = IIf(Trim(strRegistryValue) = "", strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)), strMissingPaths) & "| "
                Case "F"
                    If Trim(strRegistryValue) <> "" Then
                        If Mid(strRegistryValue, Len(strRegistryValue) - 3, 1) = "." Then
                            If Not CBool(PathFileExists(strRegistryValue)) Then
                                strMissingPaths = strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)) & "| "
                            End If
                        End If
                    Else
                        strMissingPaths = strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)) & "| "
                    End If
                Case "D"
                    If Trim(strRegistryValue) <> "" Then
                        Select Case UCase(CStr(arrSubArguments(4)))
                            Case "ACCESS"
                                enuProvider = ProviderType.[Access Provider]
                            Case "SQL"
                                enuProvider = ProviderType.[SQL Server Provider]
                        End Select
                                             
                        strYearSuffix = vbNullString
                        
                        Select Case UCase$(Trim$(CStr(arrSubArguments(IDX_DATABASE_NAME))))
                            Case UCase$("mdb_data.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_DATA
                            Case UCase$("mdb_scheduler.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_SCHEDULER
                            Case UCase$("mdb_sadbel.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_SADBEL
                            Case UCase$("mdb_scheduler.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_SCHEDULER
                            Case UCase$("edifact.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_EDIFACT
                            Case UCase$("mdb_repertory.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_REPERTORY
                            Case UCase$("mdb_taric.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_TARIC
                            Case UCase$("TemplateCP.mdb")
                                enuDBInstanceType = DBInstanceType_DATABASE_TEMPLATE
                            Case Else
                                
                                strYearSuffix = Left$(CStr(arrSubArguments(IDX_DATABASE_NAME)), Len(CStr(arrSubArguments(IDX_DATABASE_NAME))) - 4)
                                
                                If Left$(UCase$(Trim$(CStr(arrSubArguments(IDX_DATABASE_NAME)))), 11) = UCase$("mdb_history") Then
                                    enuDBInstanceType = DBInstanceType_DATABASE_HISTORY
                                    strYearSuffix = Right$(strYearSuffix, 2)
                                    
                                ElseIf Left$(UCase$(Trim$(CStr(arrSubArguments(IDX_DATABASE_NAME)))), 14) = UCase$("mdb_edihistory") Then
                                    enuDBInstanceType = DBInstanceType_DATABASE_EDI_HISTORY
                                    strYearSuffix = Right$(strYearSuffix, 2)
                                ElseIf Left$(UCase$(Trim$(CStr(arrSubArguments(IDX_DATABASE_NAME)))), 14) = UCase$("mdb_repertory") Then
                                    enuDBInstanceType = DBInstanceType_DATABASE_REPERTORY
                                    strYearSuffix = Right$(strYearSuffix, 4)
                                Else
                                    Debug.Assert False
                                End If
                        End Select
                        
                        If UBound(arrSubArguments) = IDX_DATABASE_PASSWORD Then
                            If Not ADOConnectDB(clsCheckConnection, DataSourceProperties, enuDBInstanceType, strYearSuffix) Then
                            
                                strMissingPaths = strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)) & "| "
                            End If
                        Else
                            If Not ADOConnectDB(clsCheckConnection, DataSourceProperties, enuDBInstanceType, strYearSuffix) Then
                                strMissingPaths = strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)) & "| "
                            End If
                        End If
                    Else
                        strMissingPaths = strMissingPaths & "|||||" & CStr(PathArgs(lngArgCtr)) & "| "
                    End If
            End Select
        Next lngArgCtr
        
        strMissingPaths = IIf(Len(strMissingPaths) <> 0, Mid(strMissingPaths, 6), strMissingPaths)
        
        If Trim(strMissingPaths) <> "" Then
            enuResult = FLocatePaths.OnLog(OwnerForm, Application, strMissingPaths)
            Set FLocatePaths = Nothing
            Select Case enuResult
                Case cpiSetPath
                    strArguments = Mid(strArguments, 1, Len(strArguments) - 2)
                    clsLocations.ViewLocations OwnerForm, Application, False, strArguments
                    enuResult = cpiRetry
            End Select
        Else
            enuResult = cpiSuccess
        End If
        
        If enuResult = cpiCancel Then
            Exit Do
        End If
    Loop
    
    If enuResult = cpiSuccess Then
        CheckFiles = True
    End If
    
    Set clsBrowse = Nothing
    Set clsRegistry = Nothing
    Set clsCheckConnection = Nothing

End Function
