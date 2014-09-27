Attribute VB_Name = "modTaricGlobals"
'for TARIC maintenance form taken from prjClearingPoint

Public Const G_CONST_NCTS1_TYPE = "Transit NCTS"        ' = cImport = "Import"
Public Const G_CONST_NCTS1_SHEET = "NCTS1Sheet"         ' = cCodisheet = "CodiSheet"
Public Const G_CONST_NCTS2_TYPE = "Combined NCTS"       ' = cImport = "Import"
Public Const G_CONST_NCTS2_SHEET = "NCTS2Sheet"         ' = cCodisheet = "CodiSheet"
Public Const G_CONST_EDINCTS1_TYPE = "EDI NCTS"        ' = cImport = "Import"
Public Const G_CONST_EDINCTS1_SHEET = "EDINCTS1Sheet"         ' = cCodisheet = "CodiSheet"
Public Const LOCK_FILENAME = "SADBELLock.sdb"
Public Const BIF_RETURNONLYFSDIRS = &H1

Global Const STRMDBNAME01 = "TemplateCP.mdb"
Global Const STRMDBNAME02 = "TemplateFMS.mdb"
Global Const STRMDBNAME03 = "mdb_sadbel.mdb"
Global Const STRMDBNAME04 = "mdb_data.mdb"
Global Const STRMDBNAME05 = "mdb_scheduler.mdb"

Public STRMDBPATH01 As String
Public STRMDBPATH02 As String

Public gstrTaricMainCallType As String
Public gstrTaricCNCallType As String

Public newform(10) As Form
Public newformE(10) As Form
Public newformT(10) As Form

Public frmNewFormNCTS1(10) As Form  ' = newformT(10)
Public frmNewFormNCTS2(10) As Form  ' = newformT(10)
Public frmNewFormEDINCTS1(10) As Form  ' = newformT(10)

Public csHead As String
Public csDetl As String
Public cAppPath As String

Public g_lngSelStart As Long
Public g_lngFrom As Long
Public g_strInternalCode As String
Public g_blnMultiplePick As Boolean

Public lMaintainTables As Boolean
Public lngUserNo As Long
Public cLanguage As String
Public gblnFormWasCanceled As Boolean
Public gstrMinLicValue As String          ' Added March 20, 2001
Public gstrMinValueCurr As String         ' Added March 20, 2001

Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
    (ByVal pIDList As Long, ByVal pszPath As String) As Long

Public Type BROWSEINFO
    hWndOwner As Long
    pIDListRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpFnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Enum StripCharPositionConstants
    sbpLeading = 0
    sbpTrailing
    sbpLeadingTrailing
End Enum

Public Enum ENCTSEvent
    eGot_Focus = 1
    eLost_Focus = 2
    eClick = 3
End Enum


'===========Used by TARIC maintenance form taken from prjClearingPoint===========
Public Function GetNewStr(ByVal nCode As Integer, Optional VarArray)
   Dim i As Integer, nPost As Integer, nCount As Integer, cStr1 As String
   Dim cstr2 As String, cNew As String, cString As String
   
   cString = Translate(nCode)
   nCount = CountChr(cString, "@@")
   cStr1 = cString
   cNew = ""
   For i = 1 To nCount
      nPost = InStr(1, cStr1, "@@")
      cstr2 = Left(cStr1, nPost - 1)
      cNew = cNew & cstr2 & VarArray(i - 1)
      cStr1 = Mid(cStr1, nPost + 2)
   Next
   GetNewStr = cNew & cStr1
End Function

Public Function CountChr(ByVal StringSearch As String, ByVal FindWhat As String) As Integer
    Dim nCtr As Integer
    Dim i As Integer
    Dim nFindwhat As Integer
    
    nCtr = 0
    nFindwhat = Len(Trim(FindWhat))
    
    For i = 1 To Len(Trim(StringSearch))
        If Mid(StringSearch, i, nFindwhat) = FindWhat Then
            nCtr = nCtr + 1
        End If
    Next
    
    CountChr = nCtr
End Function

Public Function StripChars(ByVal strStrippee As String, ByVal strCharToStrip As String, ByVal intPosition As StripCharPositionConstants) As String
    Dim intSearchCtr As Integer
    Dim intStrippeeLength As Integer
    Dim intCharToStripLength As Integer
    
    intStrippeeLength = Len(strStrippee)
    intCharToStripLength = Len(strCharToStrip)
    
    If intCharToStripLength Then
        Select Case intPosition
            Case sbpLeading
                For intSearchCtr = 1 To intStrippeeLength Step intCharToStripLength
                    If Mid(strStrippee, intSearchCtr, intCharToStripLength) <> strCharToStrip Then
                        Exit For
                    End If
                Next
                
                strStrippee = Mid(strStrippee, intSearchCtr)
            Case sbpTrailing
                For intSearchCtr = intStrippeeLength - (intCharToStripLength - 1) To 1 Step -intCharToStripLength
                    If Mid(strStrippee, intSearchCtr, intCharToStripLength) <> strCharToStrip Then
                        Exit For
                    End If
                Next
                
                strStrippee = Mid(strStrippee, 1, intSearchCtr + (intCharToStripLength - 1))
            Case sbpLeadingTrailing
                For intSearchCtr = 1 To intStrippeeLength Step intCharToStripLength
                    If Mid(strStrippee, intSearchCtr, intCharToStripLength) <> strCharToStrip Then
                        Exit For
                    End If
                Next
                
                strStrippee = Mid(strStrippee, intSearchCtr)
                intStrippeeLength = Len(strStrippee)
                
                For intSearchCtr = intStrippeeLength - (intCharToStripLength - 1) To 1 Step -intCharToStripLength
                    If Mid(strStrippee, intSearchCtr, intCharToStripLength) <> strCharToStrip Then
                        Exit For
                    End If
                Next
                
                strStrippee = Mid(strStrippee, 1, intSearchCtr + (intCharToStripLength - 1))
       End Select
    End If
    
    StripChars = strStrippee
End Function

Public Function GetVAT(ByVal strVatorClient As String) As String
    Dim intPos As Integer

    intPos = InStr(1, strVatorClient, "*")
    If intPos > 0 Then
        GetVAT = Left(strVatorClient, intPos - 1)
    Else
        GetVAT = strVatorClient
    End If
End Function

Public Sub CreateLogFile(ByVal FileName As String, ByVal Extension As String)
    Dim intFreeFile As Integer
    
    On Error Resume Next
    
    intFreeFile = FreeFile()
    
    If Len(Dir(cAppPath & "\" & Trim(FileName) & "." & Extension)) = 0 Then
        Open cAppPath & "\" & Trim(FileName) & "." & Extension For Output As #intFreeFile
        Print #intFreeFile, "user code " & Trim(FileName)
        Close #intFreeFile
    End If
End Sub

Public Sub UnloadControls(ByVal frmBeingUnloaded As Form, Optional ByVal blnTag As Boolean)
    Dim ctlToUnload As Control
    
    On Error Resume Next
    
    For Each ctlToUnload In frmBeingUnloaded.Controls
        Set ctlToUnload = Nothing
    Next
End Sub

Function CheckDate(ByVal DateVal, Optional ByVal lSw As Boolean, Optional ByRef frm As Form, Optional ByVal cTyp As String, Optional ByVal BoxCode As String, Optional ByVal cLog As String, Optional ByVal TextValue As String) As Boolean
'    Dim npos As Integer
'    Dim cDatStr As String
    
    Dim rsAdmin As ADODB.Recordset
    Dim rsBoxDef As ADODB.Recordset
    Dim rsDefUser As ADODB.Recordset
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim MyDate As String
    
    If lSw Then
        If NetUse("Default User " & cTyp, frm, DBInstanceType_DATABASE_SADBEL) Then
            Set rsDefUser = frm.rsTemp
        End If
        
        If NetUse("Box Default Value " & cTyp, frm, DBInstanceType_DATABASE_SADBEL) Then
            Set rsBoxDef = frm.rsTemp
        End If
        
        If NetUse("Box Default " & cTyp & " Admin", frm, DBInstanceType_DATABASE_SADBEL) Then
            Set rsAdmin = frm.rsTemp
        End If
        
        With rsDefUser
            .Filter = adFilterNone
            .Filter = "[USER NO] = " & lngUserNo & " AND [BOX CODE] = '" & BoxCode & "' AND [LOGID DESCRIPTION] = '" & cLog & "' "
            
            '.Index = "user_box_logid"
            '.Seek "=", lngUserNo, BoxCode, cLog
            
            'If Not .NoMatch Then
            If .RecordCount > 0 Then
                .MoveFirst
                
                If Trim(![Default Value]) = Trim(TextValue) Then
                    CheckDate = True
                    
                    GoTo ReleaseResources
                Else
                    GoTo A
                End If
            Else
A:              With rsBoxDef
                    .Filter = adFilterNone
                    .Filter = "[LOGID DESCRIPTION] = '" & cLog & "' AND [BOX CODE] = '" & BoxCode & "' "
                    '.Index = "LOGICAL_BOX"
                    '.Seek "=", cLog, BoxCode
                    
                    'If Not .NoMatch Then
                    If .RecordCount > 0 Then
                        .MoveFirst
                        
                        If Trim(![Default Value]) = Trim(TextValue) Then
                            CheckDate = True
                            
                            GoTo ReleaseResources
                        Else
                            GoTo B
                        End If
                    Else
B:                      With rsAdmin
                            .Filter = adFilterNone
                            .Filter = "[BOX CODE] = '" & BoxCode & "' "
                            '.Index = "box code"
                            '.Seek "=", BoxCode
                            
                            'If Not .NoMatch Then
                            If .RecordCount > 0 Then
                                .MoveFirst
                                If Trim(![EMPTY FIELD VALUE]) = Trim(TextValue) Then
                                    CheckDate = True
                                    
                                    GoTo ReleaseResources
                                Else
                                    CheckDate = False
                                End If
                            End If
                            .Filter = adFilterNone
                        End With
                    End If
                    .Filter = adFilterNone
                End With
            End If
            .Filter = adFilterNone
        End With
        
        If Len(Trim(DateVal)) <> 6 And (DateVal <> "" And Val(IIf(Len(Trim(DateVal)) > 0, DateVal, "0")) <> 0) Then
            MsgBox Translate(288), vbInformation    ' "The length of date must be equal to six."
            CheckDate = False
            
            GoTo ReleaseResources
        End If
        
        If DateVal = "" Or DateVal = " " Or Val(IIf(Len(Trim(DateVal)) > 0, DateVal, "0")) = 0 Then
            CheckDate = True
            
            GoTo ReleaseResources
        End If
        
        ADORecordsetClose rsAdmin
        ADORecordsetClose rsBoxDef
        ADORecordsetClose rsDefUser

    End If
    
    On Error GoTo ErrorDate
    
    strDay = Left(DateVal, 2)
    strMonth = Mid(DateVal, 3, 2)
    strYear = Mid(DateVal, 5, 2)
    
    Select Case Val(strYear)
        Case 0 To 29     ' 21st century
            strYear = "20" & strYear
        Case 30 To 99    ' 20th century
            strYear = "19" & strYear
    End Select
    
    MyDate = Format(CDate(strDay & "/" & strMonth & "/" & strYear), "dd/mm/yy")
    
    CheckDate = True
    
    GoTo ReleaseResources
    
ErrorDate:
    
    ' "You must specify a valid date. Check your entries to make sure " & Chr(13) & "they represent a valid date."
    MsgBox Translate(289), vbExclamation
    CheckDate = False
    
ReleaseResources:
    
    On Error Resume Next
    
    
    ADORecordsetClose rsAdmin
    ADORecordsetClose rsBoxDef
    ADORecordsetClose rsDefUser
End Function

Public Function OpenMDB(ByRef DataSourceProperties As CDataSourceProperties, _
                        ByRef CallingForm As Form, _
                        ByVal DBInstance As DBInstanceType, _
               Optional ByVal DBYear As String = vbNullString, _
               Optional ByVal ForNCTSUse As Boolean = False) As Boolean
               
'Public Function OpenMDB(ByVal strMDBpath As String, _
                        ByRef CallingForm As Form, _
                        ByVal DBInstance As DBInstanceType, _
               Optional ByRef wrkWorkspace As DAO.Workspace, _
               Optional ByVal ForNCTSUse As Boolean = False) As Boolean

    Dim conDatabase As ADODB.Connection     'Dim datDatabase As Database
    
    Dim strMsgBoxText As String
    Dim intNumOfRetries As Integer

    Const conDatabaseCorrupt As Integer = 3049

    On Error GoTo ErrHandler

'    DBEngine.SystemDB = "system.mdw"
'
'    ' added provision to open database in desired workspace.
'    If ((wrkWorkspace Is Nothing) = True) Then
'        ' use default workspace.
'        '<<< dandan 112306
'        '<<< Update with database password
'        'Set datDatabase = OpenDatabase(strMDBpath)
'        OpenDAODatabase datDatabase, strMDBpath
'
'    ElseIf ((wrkWorkspace Is Nothing) = False) Then
'        ' use desired workspace.
'        '<<< dandan 112306
'        '<<< Update with database password
'        'Set datDatabase = wrkWorkspace.OpenDatabase(strMDBpath)
'        OpenDAODatabase datDatabase, strMDBpath
'
'    End If
    
    
    Select Case DBInstance
        Case DBInstanceType_DATABASE_SADBEL
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
                'Set CallingForm.m_conSADBEL = datDatabase
                
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
                'Set CallingForm.conSadbel = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_DATA
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conData, DataSourceProperties, DBInstanceType_DATABASE_DATA
                'Set CallingForm.m_conData = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conData, DataSourceProperties, DBInstanceType_DATABASE_DATA
                'Set CallingForm.conData = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_HISTORY

            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conHistory, DataSourceProperties, DBInstanceType_DATABASE_HISTORY, DBYear
                'Set CallingForm.m_conHistory = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conHist, DataSourceProperties, DBInstanceType_DATABASE_HISTORY, DBYear
                Set CallingForm.conHist = datDatabase
            End If
            OpenMDB = True
            
        Case DBInstanceType_DATABASE_EDI_HISTORY

            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conHistory, DataSourceProperties, DBInstanceType_DATABASE_EDI_HISTORY, DBYear
                'Set CallingForm.m_conHistory = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conHist, DataSourceProperties, DBInstanceType_DATABASE_EDI_HISTORY, DBYear
                Set CallingForm.conHist = datDatabase
            End If
            OpenMDB = True
            
        Case DBInstanceType_DATABASE_TARIC
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conTaric, DataSourceProperties, DBInstanceType_DATABASE_TARIC
                'Set CallingForm.m_conTaric = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conTaric, DataSourceProperties, DBInstanceType_DATABASE_TARIC
                'Set CallingForm.conTaric = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_SCHEDULER
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conScheduler, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
                'Set CallingForm.m_conScheduler = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conScheduler, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
                'Set CallingForm.conScheduler = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_REPERTORY
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conRepertory, DataSourceProperties, DBInstanceType_DATABASE_REPERTORY, DBYear
                'Set CallingForm.m_conRepertory = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conRepertory, DataSourceProperties, DBInstanceType_DATABASE_REPERTORY, DBYear
                'Set CallingForm.conRepertory = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_TEMPLATE
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conTemplate, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
                'Set CallingForm.m_conTemplate = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conTemplate, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
                'Set CallingForm.conTemplate = datDatabase
            End If
            OpenMDB = True

        Case DBInstanceType_DATABASE_EDIFACT
            If (ForNCTSUse = True) Then
                ADOConnectDB CallingForm.m_conEDIFACT, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
                'Set CallingForm.m_conEDIFACT = datDatabase
            ElseIf (ForNCTSUse = False) Then
                ADOConnectDB CallingForm.conEdifact, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
                'Set CallingForm.conEdifact = datDatabase
            End If
            OpenMDB = True
    End Select

    Set datDatabase = Nothing
    
    Exit Function

ErrHandler:

    If (Err.Number = conDatabaseCorrupt) Then
        If (Err.Number <> 0) Then
            ' database is corrupted and cannot be opened. System will be terminated.
            MsgBox Translate(765) & vbCrLf & Err.Number & ": " & Err.Description
            OpenMDB = False

            CreateLogFile LOCK_FILENAME, "sdb"
            Open cAppPath & "\" & LOCK_FILENAME For Append Lock Read Write As #1

            Exit Function
        End If

    Else ' error

        Dim varAns As Variant

        ' do you want to try again?
        strMsgBoxText = "(" & Err.Number & ") " & Err.Description & Translate(766)

        varAns = MsgBox(strMsgBoxText, vbYesNo + vbInformation)

        If (varAns = vbYes) Then
            intNumOfRetries = intNumOfRetries + 1
            If ((intNumOfRetries >= 5) = True) Then
                ' too many retries. Please contact your administrator and report the error below:
                MsgBox Translate(764) & Chr(13) & Chr(13) & Err.Number & "  " & Err.Description, vbCritical
                OpenMDB = False

            ElseIf ((intNumOfRetries >= 5) = False) Then
                Resume
            End If

        ElseIf (varAns = vbNo) Then
            OpenMDB = False
        End If
    End If

    Set datDatabase = Nothing
End Function

Public Function NetUse(ByVal TableToOpen As String, _
                       ByRef CallingForm As Form, _
                       ByVal DBInstance As DBInstanceType, _
              Optional ByVal ForNCTSUse As Boolean = False) As Boolean

    Dim intNumOfRetries As Integer
    Dim strTableToOpen As String
    
    On Error GoTo ErrHandler
    
    strTableToOpen = GetSQLCommandFromTableName(TableToOpen)
    
    Select Case DBInstance

        Case DBInstanceType_DATABASE_SADBEL

            If (ForNCTSUse = True) Then
                
                 ADORecordsetOpen strTableToOpen, CallingForm.m_conSADBEL, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                 'Set CallingForm.m_rstDummy = CallingForm.m_conSADBEL.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then
                
                ADORecordsetOpen strTableToOpen, CallingForm.conSADBEL, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                Set CallingForm.m_rsTemp = CallingForm.conSADBEL.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_DATA

            If (ForNCTSUse = True) Then
        
                ADORecordsetOpen strTableToOpen, CallingForm.m_conData, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conData.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conData, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conData.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_HISTORY, _
             DBInstanceType_DATABASE_EDI_HISTORY

            If (ForNCTSUse = True) Then

                ADORecordsetOpen strTableToOpen, CallingForm.m_conHistory, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conHistory.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conHist, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conHist.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_TARIC

            If (ForNCTSUse = True) Then

                ADORecordsetOpen strTableToOpen, CallingForm.m_conTaric, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conTaric.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conTaric, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conTaric.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_SCHEDULER

            If (ForNCTSUse = True) Then

                ADORecordsetOpen strTableToOpen, CallingForm.m_conScheduler, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conScheduler.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conScheduler, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conScheduler.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_REPERTORY

            If (ForNCTSUse = True) Then
    
                ADORecordsetOpen strTableToOpen, CallingForm.m_conRepertory, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conRepertory.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then
    
                ADORecordsetOpen strTableToOpen, CallingForm.conRepertory, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conRepertory.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_TEMPLATE

            If (ForNCTSUse = True) Then

                ADORecordsetOpen strTableToOpen, CallingForm.m_conTemplate, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conTemplate.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conTemplate, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conTemplate.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

        Case DBInstanceType_DATABASE_EDIFACT

            If (ForNCTSUse = True) Then

                ADORecordsetOpen strTableToOpen, CallingForm.m_conEDIFACT, CallingForm.m_rstDummy, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rstDummy = CallingForm.m_conEDIFACT.OpenRecordset(strTableToOpen, dbOpenTable)

            ElseIf (ForNCTSUse = False) Then

                ADORecordsetOpen strTableToOpen, CallingForm.conEdifact, CallingForm.m_rsTemp, adOpenKeyset, adLockOptimistic
                'Set CallingForm.m_rsTemp = CallingForm.conEdifact.OpenRecordset(strTableToOpen, dbOpenTable)

            End If

            NetUse = True

    End Select

    Exit Function

ErrHandler:

    Dim varAns As Variant

    varAns = MsgBox(Err.Description & " " & Translate(766), vbInformation + vbYesNo)


    If (varAns = vbYes) Then

        intNumOfRetries = intNumOfRetries + 1

        If (intNumOfRetries >= 5) Then

            ' too many retries. Please contact your administrator and report the error below:
            MsgBox Translate(764) & Chr(13) & Chr(13) & Err.Number & " " & Err.Description, vbCritical
            NetUse = False

            Exit Function

        End If

        Resume

    ElseIf (varAns = vbNo) Then

        NetUse = False

        Exit Function

    End If

    NetUse = True

End Function
