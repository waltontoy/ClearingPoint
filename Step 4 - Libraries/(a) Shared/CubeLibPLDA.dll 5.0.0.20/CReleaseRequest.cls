VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReleaseRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private g_clsIHeader As CIHeader
Private g_clsIDetail As CIDetail
Private g_clsEHeader As CEHeader
Private g_clsEDetail As CEDetail
Private g_clsSummary As CISummary

Public Function CreateReleaseMessage(ByRef ConnectionPropeties As CDataSourceProperties, _
                                     ByVal Code As String, _
                            Optional ByVal QueueCode As String = "PLDA Belgium") As String
        
    Dim rstTemp As ADODB.Recordset
    Dim rstLogID As ADODB.Recordset
    
    Dim lngDType As Long
    
    Dim strSender As String     'Sender
    Dim strRecipient As String  'Recipient
    Dim strInterchangeReference As String   'Interchange Reference
    Dim strMRN As String    'MRN

    Dim strTempSQL As String
    

    ADOConnectDB g_conSADBEL, ConnectionPropeties, DBInstanceType_DATABASE_SADBEL
    
    'Get DType ------------------------------------------------
         strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT * "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "[PLDA MESSAGES] "
        strTempSQL = strTempSQL & "WHERE Code = '" & Code & "' "
        strTempSQL = strTempSQL & "AND Message_StatusType = 'Queued' "
        strTempSQL = strTempSQL & "AND Message_Request_Type = 3"
    
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        lngDType = rstTemp.Fields("DType").Value
    Else
        ADORecordsetClose rstTemp

        Exit Function
    End If
    ADORecordsetClose rstTemp
    '----------------------------------------------------------
    
    
    'Get Sender -----------------------------------------------
        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT "
        strTempSQL = strTempSQL & "QueueProp_QueueName AS PLDAPROP_QueueName "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "QueueProperties "
        strTempSQL = strTempSQL & "WHERE QueueProp_Type = 2 AND QueueProp_Code='" & QueueCode & "'"
        
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        strSender = FNullField(rstTemp.Fields("PLDAPROP_QueueName").Value)
    Else
        ADORecordsetClose rstTemp
        
        Exit Function
    End If
    ADORecordsetClose rstTemp
    '----------------------------------------------------------
    
    
    'Get Recipient --------------------------------------------
        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT "
        strTempSQL = strTempSQL & "[Send PLDA Recipient Operational] AS Operational, "
        strTempSQL = strTempSQL & "[Send PLDA Recipient Test] AS Test, "
        strTempSQL = strTempSQL & "Type "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "[Logical ID] "
        strTempSQL = strTempSQL & "INNER JOIN " & IIf(lngDType = 14, "[PLDA Import] ON [PLDA Import]", "[PLDA Combined] ON [PLDA Combined]") & ".[LOGID Description] = [Logical ID].[LOGID Description] "
        strTempSQL = strTempSQL & "WHERE Code = '" & Code & "'"
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        If FNullField(rstTemp.Fields("Type").Value) = "O" Then
            strRecipient = FNullField(rstTemp.Fields("Operational").Value)
        Else
            strRecipient = FNullField(rstTemp.Fields("Test").Value)
        End If
    Else
        ADORecordsetClose rstTemp
        Exit Function
    End If
    ADORecordsetClose rstTemp
    '----------------------------------------------------------
    
    'Get MRN --------------------------------------------------
        If lngDType = 14 Then
            strTempSQL = "SELECT MRN FROM [PLDA IMPORT HEADER] WHERE Code = '" & Code & "'"
        ElseIf lngDType = 18 Then
            strTempSQL = "SELECT MRN FROM [PLDA COMBINED HEADER] WHERE Code = '" & Code & "'"
        End If
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        strMRN = rstTemp.Fields("MRN").Value
    Else
        ADORecordsetClose rstTemp
        Exit Function
    End If
    ADORecordsetClose rstTemp
    
    '----------------------------------------------------------
    
    'Get Interchage Reference ---------------------------------
        If lngDType = 14 Then
            strTempSQL = "SELECT [LOGID Description] FROM [PLDA Import] WHERE Code = '" & Code & "'"
        ElseIf lngDType = 18 Then
            strTempSQL = "SELECT [LOGID Description] FROM [PLDA Combined] WHERE Code = '" & Code & "'"
        End If
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
            strTempSQL = vbNullString
            strTempSQL = strTempSQL & "SELECT "
            strTempSQL = strTempSQL & "[TIN], "
            strTempSQL = strTempSQL & "[Last EDI Reference] "
            strTempSQL = strTempSQL & "FROM "
            strTempSQL = strTempSQL & "[Logical ID] "
            strTempSQL = strTempSQL & "WHERE [LOGID Description] = '" & rstTemp.Fields("LogID Description").Value & "'"
        ADORecordsetOpen strTempSQL, g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
        'RstOpen strTempSQL, g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
        
        If rstLogID.RecordCount > 0 Then
             If Not IsNull(rstLogID.Fields("TIN").Value) And _
               Not IsNull(rstLogID.Fields("Last EDI Reference").Value) And _
               FNullField(rstLogID.Fields("TIN").Value) <> vbNullString And _
               FNullField(rstLogID.Fields("Last EDI Reference").Value) <> vbNullString Then
                
                strInterchangeReference = Left(FNullField(rstLogID.Fields("TIN").Value), 9) & Format(FNullField(rstLogID.Fields("Last EDI Reference").Value), "00000")
                
                If CLng(FNullField(rstLogID.Fields("Last EDI Reference").Value)) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(FNullField(rstLogID.Fields("Last EDI Reference").Value)) + 1
                End If

                rstLogID.Update
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, rstLogID, "Logical ID"
            End If
        End If
        ADORecordsetClose rstLogID
    End If
    ADORecordsetClose rstTemp
    '----------------------------------------------------------
    
    'Initialize module-level variables
    Call Initialize(lngDType, Code, g_conSADBEL)
    
    '*********************************************************************************************************
    '7. Create and Return the message
    '*********************************************************************************************************
    If lngDType = 14 Then
        CreateReleaseMessage = ImportReleaseMessage(strSender, _
                                                    strRecipient, _
                                                    strInterchangeReference, _
                                                    "IMPORT", _
                                                    strMRN)
    ElseIf lngDType = 18 Then
        CreateReleaseMessage = CombinedReleaseMessage(strSender, _
                                                      strRecipient, _
                                                      strInterchangeReference, _
                                                      "EXPORT", _
                                                      strMRN)
    End If
    '*********************************************************************************************************
    
End Function

Private Property Get ImportReleaseMessage(ByVal Sender As String, _
                                          ByVal Recipient As String, _
                                          ByVal InterchangeReference As String, _
                                          ByVal DType As String, _
                                          ByVal MRN As String) As String
    Dim strMessage As String
    Dim strTemp As String
    Dim astrTemp() As String
    Dim intCtr As Integer
    
    'UNB
    strMessage = "UNB+UNOA:2+" & ReplaceSpecialCharacters(Sender) & "+" & _
                                ReplaceSpecialCharacters(Recipient) & "+" & _
                                Format(Now, "yyMMdd") & ":" & _
                                Format(Now, "hhmm") & "+" & _
                                ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    'UNH
    strMessage = strMessage & "UNH+1+CUSDEC:D:04A:UN:" & _
                                ReplaceSpecialCharacters(DType) & "+" & _
                                ReplaceSpecialCharacters(MRN) & "'"
                                
    g_clsIHeader.m_lngHeaderSegments = g_clsIHeader.m_lngHeaderSegments + 1
    
    'BGM
    astrTemp = Split(g_clsIHeader.BGM, "+")
    astrTemp(2) = MRN & astrTemp(2)
    strMessage = strMessage & Join(astrTemp, "+")
    Erase astrTemp
                
    'LOC
    strMessage = strMessage & g_clsIHeader.LOC
    
    'DTM
    'strMessage = strMessage & g_clsIHeader.DTM
        
    'GEI
    strMessage = strMessage & g_clsIHeader.GEI
    
    'FII
    strMessage = strMessage & g_clsIHeader.FII
    
    'SEL
    strMessage = strMessage & g_clsIHeader.SEL
    
    'FTX
    strMessage = strMessage & g_clsIHeader.FTX
    
    'SG1
    strMessage = strMessage & g_clsIHeader.SG1
    
    'SG4
    strMessage = strMessage & g_clsIHeader.SG4
    
    'SG5
    strMessage = strMessage & g_clsIHeader.SG5
    
    'SG6
'    astrTemp = Split(g_clsIHeader.SG6, "'")
'    If UBound(astrTemp) <> -1 Then
'        For intCtr = 0 To UBound(astrTemp)
'            if astr
    strMessage = strMessage & g_clsIHeader.SG6
    
    'SG7
    strMessage = strMessage & g_clsIHeader.SG7
    
    'SG8
    strMessage = strMessage & g_clsIHeader.SG8
        
    g_clsIDetail.m_rstDetails.MoveFirst
    
    Do While Not g_clsIDetail.m_rstDetails.EOF
        
        g_clsIDetail.m_lngDetailNo = FNullField(g_clsIDetail.m_rstDetails.Fields("Detail").Value)

        'UNS
        strMessage = strMessage & IIf(g_clsIDetail.m_lngDetailNo > 1, "", g_clsIDetail.UNS)
        
        'CST
        strMessage = strMessage & g_clsIDetail.CST
        
        'FTX
        strMessage = strMessage & g_clsIDetail.FTX
        
        'LOC
        strMessage = strMessage & g_clsIDetail.LOC
        
        'SG31
        strMessage = strMessage & g_clsIDetail.SG31
        
        'SG33
        strMessage = strMessage & g_clsIDetail.SG33
        
        'SG35
        strMessage = strMessage & g_clsIDetail.SG35
        
        'SG37
        strMessage = strMessage & g_clsIDetail.SG37
        
        'SG38
        strMessage = strMessage & g_clsIDetail.SG38
        
        'SG40
        strMessage = strMessage & g_clsIDetail.SG40
        
        'SG41
        strMessage = strMessage & g_clsIDetail.SG41
        
        g_clsIDetail.m_rstDetails.MoveNext
    
    Loop
    
    'UNS
    strMessage = strMessage & g_clsSummary.UNS
    
    'CNT
    strMessage = strMessage & g_clsSummary.CNT
    
    'SG49
    strMessage = strMessage & g_clsSummary.SG49
    
    'SG50
    strMessage = strMessage & g_clsSummary.SG50
    
    'UNT
    g_clsSummary.m_lngSummarySegments = g_clsSummary.m_lngSummarySegments + 1
    strMessage = strMessage & "UNT+" & CStr(g_clsIHeader.m_lngHeaderSegments + g_clsIDetail.m_lngDetailSegments + g_clsSummary.m_lngSummarySegments) & "+1'"
    
    'UNZ
    strMessage = strMessage & "UNZ+1+" & ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    ImportReleaseMessage = strMessage
    
End Property

Private Property Get CombinedReleaseMessage(ByVal Sender As String, _
                                            ByVal Recipient As String, _
                                            ByVal InterchangeReference As String, _
                                            ByVal DType As String, _
                                            ByVal MRN As String) As String
    Dim strMessage As String
    Dim strTemp1 As String
    Dim strTemp2 As String
    Dim astrTemp() As String
    Dim strTemporaryWithoutReleaseCharacters As String
    
    Const CONST_DUMMY = "^&*("
    
    'UNB
    strMessage = "UNB+UNOA:2+" & ReplaceSpecialCharacters(Sender) & "+" & _
                                ReplaceSpecialCharacters(Recipient) & "+" & _
                                Format(Now, "yyMMdd") & ":" & _
                                Format(Now, "hhmm") & "+" & _
                                ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    'UNH
    strMessage = strMessage & "UNH+1+CUSDEC:D:04A:UN:" & _
                                ReplaceSpecialCharacters(DType) & "+" & _
                                ReplaceSpecialCharacters(MRN) & "'"
    'g_clsEHeader.m_lngHeaderSegments = g_clsEHeader.m_lngHeaderSegments + 1
    
    'BGM
    strTemp1 = g_clsEHeader.BGM
    astrTemp = Split(strTemp1, "+")
    astrTemp(2) = MRN & astrTemp(2)
    strMessage = strMessage & Join(astrTemp, "+")
    Erase astrTemp
                
    'LOC
    strTemporaryWithoutReleaseCharacters = g_clsEHeader.LOC
    
    ' Replace Special Characters with distinct characters (DUMMY)
    ' so that it will not be detected as a separator
    '   RELEASE CHARACTER
    strTemporaryWithoutReleaseCharacters = Replace(strTemporaryWithoutReleaseCharacters, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, CONST_DUMMY)
    
    ' Remove all LOC segments before LOC+14 including of LOC+14
    ' because this should not exist in Combined Release Message?
    If InStr(1, strTemporaryWithoutReleaseCharacters, "LOC+14") > 0 Then
        ' strTemp2 has no special characters because it is derived from a string without special characters
        strTemp2 = Left(strTemporaryWithoutReleaseCharacters, InStr(1, strTemporaryWithoutReleaseCharacters, "LOC+14") - 1)
    End If
    
    ' Return Released Segment Separtor Characters
    strTemporaryWithoutReleaseCharacters = Mid(strTemporaryWithoutReleaseCharacters, InStr(Len(strTemp2) + 1, strTemporaryWithoutReleaseCharacters, EDI_SEP_SEGMENT) + 1)
    
    strMessage = strMessage & Replace(strTemporaryWithoutReleaseCharacters, CONST_DUMMY, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
    
    'DTM
    strMessage = strMessage & g_clsEHeader.DTM
        
    'SEL
    strMessage = strMessage & g_clsEHeader.SEL
        
    'SG1
    strMessage = strMessage & g_clsEHeader.SG1
    
    'SG4
    strMessage = strMessage & g_clsEHeader.SG4
    
    'SG5
    strMessage = strMessage & g_clsEHeader.SG5
        
    'SG6
    strMessage = strMessage & g_clsEHeader.SG6
    
    'SG7
    strMessage = strMessage & g_clsEHeader.SG7
    
    'SG8
    strMessage = strMessage & g_clsEHeader.SG8
        
    g_clsEDetail.m_rstDetails.MoveFirst
    
    Do While Not g_clsEDetail.m_rstDetails.EOF
        
        g_clsEDetail.m_lngDetailNo = FNullField(g_clsEDetail.m_rstDetails.Fields("Detail").Value)
                
        'UNS
        strMessage = strMessage & IIf(g_clsEDetail.m_lngDetailNo > 1, "", g_clsEDetail.UNS)
        
        'CST
        strMessage = strMessage & g_clsEDetail.CST
        
        'FTX
        strTemp1 = g_clsEDetail.FTX

        strTemporaryWithoutReleaseCharacters = strTemp1
        ' Replace Special Characters with distinct characters (DUMMY)
        ' so that it will not be detected as a separator
        '   RELEASE CHARACTER
        strTemporaryWithoutReleaseCharacters = Replace(strTemporaryWithoutReleaseCharacters, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, CONST_DUMMY)
    
        astrTemp = Split(strTemporaryWithoutReleaseCharacters, EDI_SEP_SEGMENT)
        
        ' It is ok to use strTemp1 because CountSegments handle Released EDI Separator Characters
        If CountSegments(strTemp1) > 0 Then
            If InStr(1, astrTemp(0), "FTX+AAA") > 0 Then
                astrTemp(0) = Left(astrTemp(0), Len(astrTemp(0)) - 3)
            End If
        End If
        strTemporaryWithoutReleaseCharacters = Join(astrTemp, EDI_SEP_SEGMENT)
        strMessage = strMessage & Replace(strTemporaryWithoutReleaseCharacters, CONST_DUMMY, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
        
        'LOC
        strMessage = strMessage & g_clsEDetail.LOC
        
        'MEA
        strMessage = strMessage & g_clsEDetail.MEA
        
        'NAD
        strMessage = strMessage & g_clsEDetail.NAD
        
        'SG31
        strMessage = strMessage & g_clsEDetail.SG31
        
        'SG33
        strTemp1 = g_clsEDetail.SG33
        strMessage = strMessage & Mid(strTemp1, 1, Len(strTemp1) - InStr(1, strTemp1, "CUX+5"))
        'If CountSegments(strTemp1) > CountSegments(Mid(strTemp1, 1, Len(strTemp1) - InStr(1, strTemp1, "CUX+5"))) Then
        '    g_clsIDetail.m_lngDetailSegments = g_clsIDetail.m_lngDetailSegments - 1
        'End If
        
        'SG35
        strMessage = strMessage & g_clsEDetail.SG35
        
        'SG37
        strMessage = strMessage & g_clsEDetail.SG37
        
        'SG38
        strMessage = strMessage & g_clsEDetail.SG38
        
        g_clsEDetail.m_rstDetails.MoveNext
        
    Loop
    
    'UNS
    strMessage = strMessage & g_clsSummary.UNS
    
    'CNT
    strMessage = strMessage & g_clsSummary.CNT
    
    'SG49
    strMessage = strMessage & g_clsSummary.SG49
    
    'SG50
    strMessage = strMessage & g_clsSummary.SG50
    
    'UNT
    'g_clsSummary.m_lngSummarySegments = g_clsSummary.m_lngSummarySegments + 1
    'strMessage = strMessage & "UNT+" & CStr(g_clsEHeader.m_lngHeaderSegments + g_clsEDetail.m_lngDetailSegments + g_clsSummary.m_lngSummarySegments) & "+1'"
    strMessage = strMessage & "UNT+" & CStr(CountSegments(strMessage)) & "+1'"
    
    'UNZ
    strMessage = strMessage & "UNZ+1+" & ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    CombinedReleaseMessage = strMessage
    
End Property

Private Sub Initialize(ByVal DType As Long, _
                       ByVal Code As String, _
                       ByRef conConnection As ADODB.Connection)

    Dim rstHeader As ADODB.Recordset
    Dim rstHeaderZegels As ADODB.Recordset
    'Dim rstHeaderZekerheid As ADODB.Recordset  'Migs 10-14-2006
    Dim rstHeaderHandelaars As ADODB.Recordset  'Migs 10-14-2006
    Dim rstDetails As ADODB.Recordset
    Dim rstDetailsBijzondere As ADODB.Recordset
    Dim rstDetailsContainer As ADODB.Recordset
    Dim rstDetailsDocumenten As ADODB.Recordset
    Dim rstDetailsZelf As ADODB.Recordset
    Dim rstDetailsHandelaars As ADODB.Recordset 'Migs 10-14-2006
    
    Dim strSQLH As String
    Dim strSQLHZegels As String
    'Dim strSQLHZekerheid As String     'Migs 10-14-2006
    Dim strSQLHHandelaars As String     'Migs 10-14-2006
    Dim strSQLD As String
    Dim strSQLDBijzondere As String
    Dim strSQLDContainer As String
    Dim strSQLDDocumenten As String
    Dim strSQLDZelf As String
    Dim strSQLDBijgevoegde As String
    Dim strSQLDHandelaars As String     'Migs 10-14-2006
    
    Dim blnWithValue As Boolean
    
    Set rstHeader = New ADODB.Recordset
    Set rstHeaderZegels = New ADODB.Recordset
    'Set rstHeaderZekerheid = New ADODB.Recordset       'Migs 10-14-2006
    Set rstHeaderHandelaars = New ADODB.Recordset       'Migs 10-14-2006
    Set rstDetails = New ADODB.Recordset
    Set rstDetailsBijzondere = New ADODB.Recordset
    Set rstDetailsContainer = New ADODB.Recordset
    Set rstDetailsDocumenten = New ADODB.Recordset
    Set rstDetailsZelf = New ADODB.Recordset
    Set rstDetailsHandelaars = New ADODB.Recordset      'Migs 10-14-2006
        
    Select Case DType
    
        Case 14
        
            strSQLH = "SELECT * FROM [PLDA Import Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Import Header Zegels] WHERE Code = '" & Code & "'"
            strSQLHHandelaars = "SELECT * FROM [PLDA Import Header Handelaars] WHERE Code = '" & Code & "'"     'Migs 10-14-2006
            strSQLD = "SELECT * FROM [PLDA Import Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Import Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Import Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Import Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDZelf = "SELECT * FROM [PLDA Import Detail Zelf] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDHandelaars = "SELECT * FROM [PLDA Import Detail Handelaars] WHERE Code = '" & Code & "'"     'Migs 10-14-2006
                                                        
        Case 18
            
            strSQLH = "SELECT * FROM [PLDA Combined Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Combined Header Zegels] WHERE Code = '" & Code & "'"
            'strSQLHZekerheid = "SELECT * FROM [PLDA Combined Header Zekerheid] WHERE Code = '" & Code & "'"        'Migs 10-14-2006
            strSQLHHandelaars = "SELECT * FROM [PLDA Combined Header Handelaars] WHERE Code = '" & Code & "'"     'Migs 10-14-2006
            strSQLD = "SELECT * FROM [PLDA Combined Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Combined Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Combined Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Combined Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDHandelaars = "SELECT * FROM [PLDA Combined Detail Handelaars] WHERE Code = '" & Code & "'"     'Migs 10-14-2006
            
    End Select
    
    ADORecordsetOpen strSQLH, conConnection, rstHeader, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHZegels, conConnection, rstHeaderZegels, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHHandelaars, conConnection, rstHeaderHandelaars, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLD, conConnection, rstDetails, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDBijzondere, conConnection, rstDetailsBijzondere, adOpenKeyset, adLockOptimistic
    
    
    ADORecordsetOpen strSQLDContainer, conConnection, rstDetailsContainer, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDDocumenten, conConnection, rstDetailsDocumenten, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDHandelaars, conConnection, rstDetailsHandelaars, adOpenKeyset, adLockOptimistic
    
    'rstHeader.CursorLocation = adUseClient
    'rstHeader.Open strSQLH, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstHeader.ActiveConnection = Nothing
    
    'rstHeaderZegels.CursorLocation = adUseClient
    'rstHeaderZegels.Open strSQLHZegels, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstHeaderZegels.ActiveConnection = Nothing
    
    'Migs 10-14-2006
    'rstHeaderHandelaars.CursorLocation = adUseClient
    'rstHeaderHandelaars.Open strSQLHHandelaars, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstHeaderHandelaars.ActiveConnection = Nothing
    
    'rstDetails.CursorLocation = adUseClient
    'rstDetails.Open strSQLD, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstDetails.ActiveConnection = Nothing
    
    'rstDetailsBijzondere.CursorLocation = adUseClient
    'rstDetailsBijzondere.Open strSQLDBijzondere, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstDetailsBijzondere.ActiveConnection = Nothing
    
    'rstDetailsContainer.CursorLocation = adUseClient
    'rstDetailsContainer.Open strSQLDContainer, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstDetailsContainer.ActiveConnection = Nothing
    
    'rstDetailsDocumenten.CursorLocation = adUseClient
    'rstDetailsDocumenten.Open strSQLDDocumenten, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstDetailsDocumenten.ActiveConnection = Nothing
    
    'Migs 10-14-2006
    'rstDetailsHandelaars.CursorLocation = adUseClient
    'rstDetailsHandelaars.Open strSQLDHandelaars, conConnection, adOpenKeyset, adLockOptimistic
    'Set rstDetailsHandelaars.ActiveConnection = Nothing
    
    If DType = 14 Then
        
        ADORecordsetOpen strSQLDZelf, conConnection, rstDetailsZelf, adOpenKeyset, adLockOptimistic
        
        'rstDetailsZelf.CursorLocation = adUseClient
        'rstDetailsZelf.Open strSQLDZelf, conConnection, adOpenKeyset, adLockOptimistic
        'Set rstDetailsZelf.ActiveConnection = Nothing
        
    ElseIf DType = 18 Then
        'Migs 10-14-2006
        'rstHeaderZekerheid.CursorLocation = adUseClient
        'rstHeaderZekerheid.Open strSQLHZekerheid, conConnection, adOpenKeyset, adLockOptimistic
        'Set rstHeaderZekerheid.ActiveConnection = Nothing
    End If
    
    Set g_clsIHeader = New CIHeader
    Set g_clsIDetail = New CIDetail
    Set g_clsEHeader = New CEHeader
    Set g_clsEDetail = New CEDetail
    
    Set g_rstDetailsHandelaars = rstDetailsHandelaars   'Migs 10-14-2006
    
    If DType = 14 Then
            
        Set g_clsIHeader.m_rstHeader = rstHeader
        Set g_clsIHeader.m_rstHeaderZegels = rstHeaderZegels
        Set g_clsIHeader.m_rstHeaderHandelaars = rstHeaderHandelaars    'Migs 10-14-2006
        
        Set g_clsIDetail.m_rstDetails = rstDetails
        Set g_clsIDetail.m_rstDetailsBijzondere = rstDetailsBijzondere
        Set g_clsIDetail.m_rstDetailsContainer = rstDetailsContainer
        Set g_clsIDetail.m_rstDetailsDocumenten = rstDetailsDocumenten
        Set g_clsIDetail.m_rstDetailsZelf = rstDetailsZelf
        Set g_clsIDetail.m_rstHeader = rstHeader
        Set g_clsIDetail.m_rstDetailsHandelaars = rstDetailsHandelaars  'Migs 10-14-2006
        
        g_clsIHeader.m_lngHeaderSegments = 0
        g_clsIDetail.m_lngDetailSegments = 0
        g_clsIHeader.m_enuDeclarationMode = enuOriginal
        g_clsIDetail.m_enuDeclarationMode = enuOriginal
        
    ElseIf DType = 18 Then
        
        Set g_clsEHeader.m_rstHeader = rstHeader
        Set g_clsEHeader.m_rstHeaderZegels = rstHeaderZegels
        'Set g_clsEHeader.m_rstHeaderZekerheid = rstHeaderZekerheid     'Migs 10-14-2006
        Set g_clsEHeader.m_rstHeaderHandelaars = rstHeaderHandelaars    'Migs 10-14-2006
        
        Set g_clsEDetail.m_rstDetails = rstDetails
        Set g_clsEDetail.m_rstDetailsBijzondere = rstDetailsBijzondere
        Set g_clsEDetail.m_rstDetailsContainer = rstDetailsContainer
        Set g_clsEDetail.m_rstDetailsDocumenten = rstDetailsDocumenten
        Set g_clsEDetail.m_rstDetailsHandelaars = rstDetailsHandelaars  'Migs 10-14-2006
        
        g_clsEDetail.m_strLang = FNullField(rstHeader.Fields("A9").Value)
        'g_clsEDetail.m_strDeclarant = rstHeader!X1 'temporarily commented: Glenn
        g_clsEHeader.m_lngHeaderSegments = 0
        g_clsEDetail.m_lngDetailSegments = 0
        g_clsEHeader.m_enuDeclarationMode = enuOriginal
        g_clsEDetail.m_enuDeclarationMode = enuOriginal
        
    End If
    
    Set g_clsSummary = New CISummary
    Set g_clsSummary.m_rstSummary = rstHeader
    g_clsSummary.m_lngDType = DType
    g_clsSummary.m_lngNumberOfDetails = rstDetails.RecordCount
    g_clsSummary.m_lngSummarySegments = 0
    g_clsSummary.m_lngSealCount = 0
    g_clsSummary.m_enuDeclarationMode = enuOriginal
    
    If rstHeaderZegels.RecordCount <> 0 Then
        rstHeaderZegels.MoveFirst
        Do While Not rstHeaderZegels.EOF
            blnWithValue = LenB(Trim$(FNullField(rstHeaderZegels.Fields("E1").Value))) > 0 Or _
                           LenB(Trim$(FNullField(rstHeaderZegels.Fields("E2").Value))) > 0
            
            If blnWithValue Then
                g_clsSummary.m_lngSealCount = g_clsSummary.m_lngSealCount + 1
            End If
            
            If rstHeaderZegels.Fields("E3").Value = "E" Then Exit Do
            
            rstHeaderZegels.MoveNext
        Loop
    End If
    
    Set g_rstDetails = rstDetails
    
End Sub

Private Sub Class_Terminate()
      
    If g_clsIHeader Is Nothing Then Exit Sub
    If g_clsIDetail Is Nothing Then Exit Sub
    If g_clsEHeader Is Nothing Then Exit Sub
    If g_clsEDetail Is Nothing Then Exit Sub
    If g_clsSummary Is Nothing Then Exit Sub

    If Not g_clsIHeader.m_rstHeader Is Nothing Then
        If g_clsIHeader.m_rstHeader.State = adStateOpen Then g_clsIHeader.m_rstHeader.Close
        Set g_clsIHeader.m_rstHeader = Nothing
    End If
    
    If Not g_clsIHeader.m_rstHeaderZegels Is Nothing Then
        If g_clsIHeader.m_rstHeaderZegels.State = adStateOpen Then g_clsIHeader.m_rstHeaderZegels.Close
        Set g_clsIHeader.m_rstHeaderZegels = Nothing
    End If
    
    'Migs 10-14-2006
    If Not g_clsIHeader.m_rstHeaderHandelaars Is Nothing Then
        If g_clsIHeader.m_rstHeaderHandelaars.State = adStateOpen Then g_clsIHeader.m_rstHeaderHandelaars.Close
        Set g_clsIHeader.m_rstHeaderHandelaars = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetails Is Nothing Then
        If g_clsIDetail.m_rstDetails.State = adStateOpen Then g_clsIDetail.m_rstDetails.Close
        Set g_clsIDetail.m_rstDetails = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsBijzondere Is Nothing Then
        If g_clsIDetail.m_rstDetailsBijzondere.State = adStateOpen Then g_clsIDetail.m_rstDetailsBijzondere.Close
        Set g_clsIDetail.m_rstDetailsBijzondere = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsContainer Is Nothing Then
        If g_clsIDetail.m_rstDetailsContainer.State = adStateOpen Then g_clsIDetail.m_rstDetailsContainer.Close
        Set g_clsIDetail.m_rstDetailsContainer = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsDocumenten Is Nothing Then
        If g_clsIDetail.m_rstDetailsDocumenten.State = adStateOpen Then g_clsIDetail.m_rstDetailsDocumenten.Close
        Set g_clsIDetail.m_rstDetailsDocumenten = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsZelf Is Nothing Then
        If g_clsIDetail.m_rstDetailsZelf.State = adStateOpen Then g_clsIDetail.m_rstDetailsZelf.Close
        Set g_clsIDetail.m_rstDetailsZelf = Nothing
    End If
    
    'Migs 10-14-2006
    If Not g_clsIDetail.m_rstDetailsHandelaars Is Nothing Then
        If g_clsIDetail.m_rstDetailsHandelaars.State = adStateOpen Then g_clsIDetail.m_rstDetailsHandelaars.Close
        Set g_clsIDetail.m_rstDetailsHandelaars = Nothing
    End If
       
    If Not g_clsIDetail.m_rstHeader Is Nothing Then
        If g_clsIDetail.m_rstHeader.State = adStateOpen Then g_clsIDetail.m_rstHeader.Close
        Set g_clsIDetail.m_rstHeader = Nothing
    End If
    
    If Not g_clsEHeader.m_rstHeader Is Nothing Then
        If g_clsEHeader.m_rstHeader.State = adStateOpen Then g_clsEHeader.m_rstHeader.Close
        Set g_clsEHeader.m_rstHeader = Nothing
    End If
    
    If Not g_clsEHeader.m_rstHeaderZegels Is Nothing Then
        If g_clsEHeader.m_rstHeaderZegels.State = adStateOpen Then g_clsEHeader.m_rstHeaderZegels.Close
        Set g_clsEHeader.m_rstHeaderZegels = Nothing
    End If
    
    'Migs 10-14-2006
    'If Not g_clsEHeader.m_rstHeaderZekerheid Is Nothing Then
    '    If g_clsEHeader.m_rstHeaderZekerheid.State = adStateOpen Then g_clsEHeader.m_rstHeaderZekerheid.Close
    '    Set g_clsEHeader.m_rstHeaderZekerheid = Nothing
    'End If

    'Migs 10-14-2006
    If Not g_clsEHeader.m_rstHeaderHandelaars Is Nothing Then
        If g_clsEHeader.m_rstHeaderHandelaars.State = adStateOpen Then g_clsEHeader.m_rstHeaderHandelaars.Close
        Set g_clsEHeader.m_rstHeaderHandelaars = Nothing
    End If

    If Not g_clsEDetail.m_rstDetails Is Nothing Then
        If g_clsEDetail.m_rstDetails.State = adStateOpen Then g_clsEDetail.m_rstDetails.Close
        Set g_clsEDetail.m_rstDetails = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsBijzondere Is Nothing Then
        If g_clsEDetail.m_rstDetailsBijzondere.State = adStateOpen Then g_clsEDetail.m_rstDetailsBijzondere.Close
        Set g_clsEDetail.m_rstDetailsBijzondere = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsContainer Is Nothing Then
        If g_clsEDetail.m_rstDetailsContainer.State = adStateOpen Then g_clsEDetail.m_rstDetailsContainer.Close
        Set g_clsEDetail.m_rstDetailsContainer = Nothing
    End If
    
    If Not g_clsIDetail.m_rstDetailsDocumenten Is Nothing Then
        If g_clsEDetail.m_rstDetailsDocumenten.State = adStateOpen Then g_clsEDetail.m_rstDetailsDocumenten.Close
        Set g_clsEDetail.m_rstDetailsDocumenten = Nothing
    End If
        
    'Migs 10-14-2006
    If Not g_clsEDetail.m_rstDetailsHandelaars Is Nothing Then
        If g_clsEDetail.m_rstDetailsHandelaars.State = adStateOpen Then g_clsEDetail.m_rstDetailsHandelaars.Close
        Set g_clsEDetail.m_rstDetailsHandelaars = Nothing
    End If
        
    Set g_clsIHeader = Nothing
    Set g_clsIDetail = Nothing
    Set g_clsEHeader = Nothing
    Set g_clsEDetail = Nothing
    Set g_clsSummary = Nothing
    
    ADODisconnectDB g_conSADBEL
End Sub


