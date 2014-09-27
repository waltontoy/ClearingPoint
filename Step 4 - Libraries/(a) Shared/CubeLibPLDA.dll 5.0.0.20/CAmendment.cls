VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAmendment"
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

Public Function CreateAmendmentMessage(ByRef DataSourceProperties As CDataSourceProperties, _
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
    
    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
    
    'Get DType ------------------------------------------------
    Set rstTemp = New ADODB.Recordset


        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT * "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "[PLDA MESSAGES] "
        strTempSQL = strTempSQL & "WHERE "
        strTempSQL = strTempSQL & "Code = '" & Code & "' "
        strTempSQL = strTempSQL & "AND "
        strTempSQL = strTempSQL & "Message_StatusType = 'Queued' "
        strTempSQL = strTempSQL & "AND "
        strTempSQL = strTempSQL & "Message_Request_Type = 2"
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
    
    ' Get Sender -----------------------------------------------
    Set rstTemp = New ADODB.Recordset
    
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
    Set rstTemp = New ADODB.Recordset
    

    strTempSQL = vbNullString
    strTempSQL = strTempSQL & "SELECT "
    strTempSQL = strTempSQL & "[Send PLDA Recipient Operational] AS Operational, "
    strTempSQL = strTempSQL & "[Send PLDA Recipient Test] AS Test, "
    strTempSQL = strTempSQL & "Type "
    strTempSQL = strTempSQL & "FROM "
    strTempSQL = strTempSQL & "[Logical ID] "
    strTempSQL = strTempSQL & "INNER JOIN "
    strTempSQL = strTempSQL & IIf(lngDType = 14, "[PLDA Import] ON [PLDA Import]", "[PLDA Combined] ON [PLDA Combined]") & ".[LOGID Description] = [Logical ID].[LOGID Description] "
    strTempSQL = strTempSQL & "WHERE Code = '" & Code & "' "
    
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
    Set rstTemp = New ADODB.Recordset
    
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
    Set rstTemp = New ADODB.Recordset
    
    If lngDType = 14 Then
        strTempSQL = "SELECT [LOGID Description] FROM [PLDA Import] WHERE Code = '" & Code & "'"
    ElseIf lngDType = 18 Then
        strTempSQL = "SELECT [LOGID Description] FROM [PLDA Combined] WHERE Code = '" & Code & "'"
    End If
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        Set rstLogID = New ADODB.Recordset
        
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
               rstLogID.Fields("TIN").Value <> vbNullString And _
               rstLogID.Fields("Last EDI Reference").Value <> vbNullString Then
                
                strInterchangeReference = Left(rstLogID.Fields("TIN").Value, 9) & Format(rstLogID.Fields("Last EDI Reference").Value, "00000")
                
                If CLng(rstLogID.Fields("Last EDI Reference").Value) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(rstLogID.Fields("Last EDI Reference").Value) + 1
                End If
                
                rstLogID.Update
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, rstLogID, "Logical ID"
            End If
            
            ADORecordsetClose rstLogID
        End If
    End If
    ADORecordsetClose rstTemp
    '----------------------------------------------------------
    
    'Initialize module-level variables
    Call Initialize(lngDType, Code, g_conSADBEL)
    
    'Create and Return the message
    If lngDType = 14 Then
        CreateAmendmentMessage = ImportAmendmentMessage(strSender, _
                                                        strRecipient, _
                                                        strInterchangeReference, _
                                                        "IMPORT", _
                                                        strMRN)
    ElseIf lngDType = 18 Then
        CreateAmendmentMessage = CombinedAmendmentMessage(strSender, _
                                                          strRecipient, _
                                                          strInterchangeReference, _
                                                          "EXPORT", _
                                                          strMRN)
    End If
    
End Function

Private Property Get ImportAmendmentMessage(ByVal Sender As String, _
                                            ByVal Recipient As String, _
                                            ByVal InterchangeReference As String, _
                                            ByVal DType As String, _
                                            ByVal MRN As String) As String
    Dim strMessage As String
        
    'UNB
    strMessage = "UNB+UNOA:2+" & _
                ReplaceSpecialCharacters(Sender) & "+" & _
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
    strMessage = strMessage & g_clsIHeader.BGM
                
    'DTM
    strMessage = strMessage & g_clsIHeader.DTM
        
    'SEL
    strMessage = strMessage & g_clsIHeader.SEL
        
    'SG1
    strMessage = strMessage & g_clsIHeader.SG1
    
    'SG4
    strMessage = strMessage & g_clsIHeader.SG4
        
    'SG7
    strMessage = strMessage & g_clsIHeader.SG7
        
    g_clsIDetail.m_rstDetails.MoveFirst
    
    Do While Not g_clsIDetail.m_rstDetails.EOF
        
        g_clsIDetail.m_lngDetailNo = FNullField(g_clsIDetail.m_rstDetails.Fields("Detail").Value)

        'UNS
        strMessage = strMessage & IIf(g_clsIDetail.m_lngDetailNo > 1, "", g_clsIDetail.UNS)
        
        'CST
        strMessage = strMessage & g_clsIDetail.CST
        
        'FTX
        strMessage = strMessage & g_clsIDetail.FTX
        
        'SG31
        strMessage = strMessage & g_clsIDetail.SG31
        
        'SG35
        strMessage = strMessage & g_clsIDetail.SG35
        
        'SG37
        strMessage = strMessage & g_clsIDetail.SG37
        
        g_clsIDetail.m_rstDetails.MoveNext
    
    Loop
    
    'UNS
    strMessage = strMessage & g_clsSummary.UNS
    
    'CNT
    strMessage = strMessage & g_clsSummary.CNT
    
    'UNT
    g_clsSummary.m_lngSummarySegments = g_clsSummary.m_lngSummarySegments + 1
    strMessage = strMessage & "UNT+" & CStr(g_clsIHeader.m_lngHeaderSegments + g_clsIDetail.m_lngDetailSegments + g_clsSummary.m_lngSummarySegments) & "+1'"
    
    'UNZ
    strMessage = strMessage & "UNZ+1+" & ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    ImportAmendmentMessage = strMessage
    
End Property

Private Property Get CombinedAmendmentMessage(ByVal Sender As String, _
                                              ByVal Recipient As String, _
                                              ByVal InterchangeReference As String, _
                                              ByVal DType As String, _
                                              ByVal MRN As String) As String
    Dim strMessage As String
        
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
                                
    g_clsEHeader.m_lngHeaderSegments = g_clsEHeader.m_lngHeaderSegments + 1
    
    'BGM
    strMessage = strMessage & g_clsEHeader.BGM
                
    'DTM
    strMessage = strMessage & g_clsEHeader.DTM
        
    'SEL
    strMessage = strMessage & g_clsEHeader.SEL
        
    'SG1
    strMessage = strMessage & g_clsEHeader.SG1
    
    'SG4
    strMessage = strMessage & g_clsEHeader.SG4
        
    'SG7
    strMessage = strMessage & g_clsEHeader.SG7
        
    g_clsEDetail.m_rstDetails.MoveFirst
    
    Do While Not g_clsEDetail.m_rstDetails.EOF
        
        g_clsEDetail.m_lngDetailNo = FNullField(g_clsEDetail.m_rstDetails.Fields("Detail").Value)
                
        'UNS
        strMessage = strMessage & IIf(g_clsEDetail.m_lngDetailNo > 1, "", g_clsEDetail.UNS)
        
        'CST
        strMessage = strMessage & g_clsEDetail.CST
        
        'FTX
        strMessage = strMessage & g_clsEDetail.FTX
        
        'SG31
        strMessage = strMessage & g_clsEDetail.SG31
        
        'SG35
        strMessage = strMessage & g_clsEDetail.SG35
        
        'SG37
        strMessage = strMessage & g_clsEDetail.SG37
        
        g_clsEDetail.m_rstDetails.MoveNext
        
    Loop
    
    'UNS
    strMessage = strMessage & g_clsSummary.UNS
    
    'CNT
    strMessage = strMessage & g_clsSummary.CNT
    
    'UNT
    g_clsSummary.m_lngSummarySegments = g_clsSummary.m_lngSummarySegments + 1
    strMessage = strMessage & "UNT+" & CStr(g_clsEHeader.m_lngHeaderSegments + g_clsEDetail.m_lngDetailSegments + g_clsSummary.m_lngSummarySegments) & "+1'"
    
    'UNZ
    strMessage = strMessage & "UNZ+1+" & ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    CombinedAmendmentMessage = strMessage
    
End Property

Private Sub Initialize(ByVal DType As Long, _
                       ByVal Code As String, _
                       ByRef conConnection As ADODB.Connection)

    Dim rstHeader As ADODB.Recordset
    Dim rstHeaderZegels As ADODB.Recordset
    Dim rstHeaderTransitOffices As ADODB.Recordset
    Dim rstHeaderZekerheid As ADODB.Recordset
    
    Dim rstDetails As ADODB.Recordset
    Dim rstDetailsBijzondere As ADODB.Recordset
    Dim rstDetailsContainer As ADODB.Recordset
    Dim rstDetailsDocumenten As ADODB.Recordset
    Dim rstDetailsZelf As ADODB.Recordset
    Dim rstDetailsSensitiveGoods As ADODB.Recordset
    Dim rstDetailsBerekeningsEenheden As ADODB.Recordset
    
    Dim rstHeaderHandelaars As ADODB.Recordset
    Dim rstDetailsHandelaars As ADODB.Recordset
    
    Dim strSQLH As String
    Dim strSQLHZegels As String
    Dim strSQLHHandelaars As String
    Dim strSQLHZekerheid As String
    Dim strSQLD As String
    Dim strSQLDBijzondere As String
    Dim strSQLDContainer As String
    Dim strSQLDDocumenten As String
    Dim strSQLDZelf As String
    Dim strSQLDBijgevoegde As String
    Dim strSQLDHandelaars As String
    Dim strSQLDBerekeningsEenheden As String
    Dim strSQLHTransitOffices As String
    Dim strSQLDSensitiveGoods As String
    
    Dim blnWithValue As Boolean
    
    Select Case DType
    
        Case 14
            strSQLH = "SELECT * FROM [PLDA Import Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Import Header Zegels] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            strSQLHHandelaars = "SELECT * FROM [PLDA Import Header Handelaars] WHERE Code = '" & Code & "'"
            
            strSQLD = "SELECT * FROM [PLDA Import Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Import Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Import Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Import Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDZelf = "SELECT * FROM [PLDA Import Detail Zelf] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDHandelaars = "SELECT * FROM [PLDA Import Detail Handelaars] WHERE Code = '" & Code & "'"
            strSQLDBerekeningsEenheden = "SELECT * FROM [PLDA Import Detail Berekenings Eenheden] WHERE Code = '" & Code & "'"
                                                        
        Case 18
            strSQLH = "SELECT * FROM [PLDA Combined Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Combined Header Zegels] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            strSQLHZekerheid = "SELECT * FROM [PLDA Combined Header Zekerheid] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            strSQLHTransitOffices = "SELECT * FROM [PLDA Combined Header Transit Offices] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            strSQLHHandelaars = "SELECT * FROM [PLDA Combined Header Handelaars] WHERE Code = '" & Code & "'"
            
            strSQLD = "SELECT * FROM [PLDA Combined Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Combined Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Combined Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Combined Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDSensitiveGoods = "SELECT * FROM [PLDA Combined Detail Sensitive Goods] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDHandelaars = "SELECT * FROM [PLDA Combined Detail Handelaars] WHERE Code = '" & Code & "'"
            
    End Select
    
    ADORecordsetOpen strSQLH, conConnection, rstHeader, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHZegels, conConnection, rstHeaderZegels, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLD, conConnection, rstDetails, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDBijzondere, conConnection, rstDetailsBijzondere, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDContainer, conConnection, rstDetailsContainer, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDDocumenten, conConnection, rstDetailsDocumenten, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHHandelaars, conConnection, rstHeaderHandelaars, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDHandelaars, conConnection, rstDetailsHandelaars, adOpenKeyset, adLockOptimistic
    
    'RstOpen strSQLH, conConnection, rstHeader, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHZegels, conConnection, rstHeaderZegels, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLD, conConnection, rstDetails, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDBijzondere, conConnection, rstDetailsBijzondere, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDContainer, conConnection, rstDetailsContainer, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDDocumenten, conConnection, rstDetailsDocumenten, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHHandelaars, conConnection, rstHeaderHandelaars, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDHandelaars, conConnection, rstDetailsHandelaars, adOpenKeyset, adLockOptimistic, , True
    
    Set g_rstDetailsHandelaars = rstDetailsHandelaars   'This is for MProcedures.
    Set g_rstDetails = rstDetails                       'This is for MProcedures.
    
    If DType = 14 Then
        ADORecordsetOpen strSQLDBerekeningsEenheden, conConnection, rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDZelf, conConnection, rstDetailsZelf, adOpenKeyset, adLockOptimistic
        
        'RstOpen strSQLDBerekeningsEenheden, conConnection, rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDZelf, conConnection, rstDetailsZelf, adOpenKeyset, adLockOptimistic, , True
    ElseIf DType = 18 Then
        
        ADORecordsetOpen strSQLHZekerheid, conConnection, rstHeaderZekerheid, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLHTransitOffices, conConnection, rstHeaderTransitOffices, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDSensitiveGoods, conConnection, rstDetailsSensitiveGoods, adOpenKeyset, adLockOptimistic
        
        'RstOpen strSQLHZekerheid, conConnection, rstHeaderZekerheid, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLHTransitOffices, conConnection, rstHeaderTransitOffices, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDSensitiveGoods, conConnection, rstDetailsSensitiveGoods, adOpenKeyset, adLockOptimistic, , True
    End If
    
    Set g_clsIHeader = New CIHeader
    Set g_clsIDetail = New CIDetail
    Set g_clsEHeader = New CEHeader
    Set g_clsEDetail = New CEDetail
    
    Set g_rstDetailsHandelaars = rstDetailsHandelaars
    
    If DType = 14 Then
        Set g_clsIHeader.m_rstHeader = rstHeader
        Set g_clsIHeader.m_rstHeaderZegels = rstHeaderZegels
        Set g_clsIHeader.m_rstHeaderHandelaars = rstHeaderHandelaars
        
        Set g_clsIDetail.m_rstDetails = rstDetails
        Set g_clsIDetail.m_rstDetailsBijzondere = rstDetailsBijzondere
        Set g_clsIDetail.m_rstDetailsContainer = rstDetailsContainer
        Set g_clsIDetail.m_rstDetailsDocumenten = rstDetailsDocumenten
        Set g_clsIDetail.m_rstDetailsZelf = rstDetailsZelf
        Set g_clsIDetail.m_rstHeader = rstHeader
        Set g_clsIDetail.m_rstDetailsHandelaars = rstDetailsHandelaars
        Set g_clsIDetail.m_rstDetailsBerekeningsEenheden = rstDetailsBerekeningsEenheden
        
        g_clsIHeader.m_lngHeaderSegments = 0
        g_clsIDetail.m_lngDetailSegments = 0
        g_clsIHeader.m_enuDeclarationMode = enuAmendment
        g_clsIDetail.m_enuDeclarationMode = enuAmendment
        
    ElseIf DType = 18 Then
        
        Set g_clsEHeader.m_rstHeader = rstHeader
        Set g_clsEHeader.m_rstHeaderZegels = rstHeaderZegels
        Set g_clsEHeader.m_rstHeaderHandelaars = rstHeaderHandelaars
        Set g_clsEHeader.m_rstHeaderZekerheid = rstHeaderZekerheid
        Set g_clsEHeader.m_rstHeaderTransitOffices = rstHeaderTransitOffices
        
        Set g_clsEDetail.m_rstDetails = rstDetails
        Set g_clsEDetail.m_rstDetailsBijzondere = rstDetailsBijzondere
        Set g_clsEDetail.m_rstDetailsContainer = rstDetailsContainer
        Set g_clsEDetail.m_rstDetailsDocumenten = rstDetailsDocumenten
        Set g_clsEDetail.m_rstDetailsHandelaars = rstDetailsHandelaars
        Set g_clsEDetail.m_rstDetailsSensitiveGoods = rstDetailsSensitiveGoods
        
        g_clsEDetail.m_strLang = FNullField(rstHeader.Fields("A9").Value)
        'g_clsEDetail.m_strDeclarant = rstHeader!X1 'temporarily commented: Glenn
        g_clsEHeader.m_lngHeaderSegments = 0
        g_clsEDetail.m_lngDetailSegments = 0
        g_clsEHeader.m_enuDeclarationMode = enuAmendment
        g_clsEDetail.m_enuDeclarationMode = enuAmendment
        
    End If
    
    Set g_clsSummary = New CISummary
    Set g_clsSummary.m_rstSummary = rstHeader
    g_clsSummary.m_lngDType = DType
    g_clsSummary.m_lngNumberOfDetails = rstDetails.RecordCount
    g_clsSummary.m_lngSummarySegments = 0
    g_clsSummary.m_lngSealCount = 0
    g_clsSummary.m_enuDeclarationMode = enuAmendment
    
    If rstHeaderZegels.RecordCount <> 0 Then
        rstHeaderZegels.MoveFirst
        Do While Not rstHeaderZegels.EOF
            blnWithValue = FNullField(rstHeaderZegels.Fields("E1").Value) <> "" And FNullField(rstHeaderZegels.Fields("E2").Value) <> ""
            If blnWithValue Then g_clsSummary.m_lngSealCount = g_clsSummary.m_lngSealCount + 1
            If FNullField(rstHeaderZegels.Fields("E3").Value) = "E" Then Exit Do
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
    
    If Not g_clsIDetail.m_rstHeader Is Nothing Then
        If g_clsIDetail.m_rstHeader.State = adStateOpen Then g_clsIDetail.m_rstHeader.Close
        Set g_clsIDetail.m_rstHeader = Nothing
    End If
    
    'Migs 10-14-2006
    If Not g_clsIDetail.m_rstDetailsHandelaars Is Nothing Then
        If g_clsIDetail.m_rstDetailsHandelaars.State = adStateOpen Then g_clsIDetail.m_rstDetailsHandelaars.Close
        Set g_clsIDetail.m_rstDetailsHandelaars = Nothing
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
    If Not g_clsEHeader.m_rstHeaderHandelaars Is Nothing Then
        If g_clsEHeader.m_rstHeaderHandelaars.State = adStateOpen Then g_clsEHeader.m_rstHeaderHandelaars.Close
        Set g_clsEHeader.m_rstHeaderHandelaars = Nothing
    End If
    
    'Migs 10-14-2006
    'If Not g_clsEHeader.m_rstHeaderZekerheid Is Nothing Then
    '    If g_clsEHeader.m_rstHeaderZekerheid.State = adStateOpen Then g_clsEHeader.m_rstHeaderZekerheid.Close
    '    Set g_clsEHeader.m_rstHeaderZekerheid = Nothing
    'End If
    
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
    If Not g_clsIDetail.m_rstDetailsHandelaars Is Nothing Then
        If g_clsEDetail.m_rstDetailsHandelaars.State = adStateOpen Then g_clsEDetail.m_rstDetailsHandelaars.Close
        Set g_clsEDetail.m_rstDetailsHandelaars = Nothing
    End If
    
    Set g_clsIHeader = Nothing
    Set g_clsIDetail = Nothing
    Set g_clsEHeader = Nothing
    Set g_clsEDetail = Nothing
    Set g_clsSummary = Nothing
    
End Sub


