VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDIMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CreateEDIMessage(ByRef DataSourceProperties As CDataSourceProperties, _
                                 ByVal Code As String, _
                                 ByVal DType As Long, _
                                 ByRef CallingForm As Object, _
                        Optional ByVal QueueCode As String = "PLDA Belgium") As String
    
    Dim rstDetailsCount As ADODB.Recordset
    
    Dim rstTemp As ADODB.Recordset
    Dim rstLogID As ADODB.Recordset
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
    
    Dim rstHeaderHandelaars As ADODB.Recordset
    Dim rstDetailsHandelaars As ADODB.Recordset
    Dim rstDetailsBerekeningsEenheden As ADODB.Recordset
    
    Dim rstPLDAMessages As ADODB.Recordset
    
    Dim CImpHeader As CIHeader
    Dim CExpHeader As CEHeader
    Dim CSummary As CISummary
    Dim CPLDADetails As CDetails
    
    Dim strCommand As String
    Dim strSQLH As String
    Dim strSQLHZegels As String
    Dim strSQLHZekerheid As String
    Dim strSQLHTransitOffices As String
    
    Dim strSQLD As String
    Dim strSQLDBijzondere As String
    Dim strSQLDContainer As String
    Dim strSQLDDocumenten As String
    Dim strSQLDZelf As String
    Dim strSQLDBijgevoegde As String
    Dim strSQLDSensitiveGoods As String
    
    Dim strSQLHHandelaars As String
    Dim strSQLDHandelaars As String
    Dim strSQLDBerekeningsEenheden As String
    
    Dim strSQLPLDAMessages As String
    
    Dim strEDIUserName As String
    Dim strSendRecipient
    Dim strAcct49 As String
    Dim strLogTIN As String
    Dim strLastEDIReference As String
    Dim strMessage As String
    Dim lngSegments As Long
    Dim lngSealsCount As Long
    Dim lngNumberOfDetails As Long
    Dim lngCtr As Long
    Dim blnWithValue As Boolean
    Dim dblNetWt As Double
    
    Dim clsDigisign As CDigiSignFields
    
    Dim conScheduler As ADODB.Connection
    
    Dim lngDetailCountMainTable As Long 'csclp-473
    
    ADOConnectDB conScheduler, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL

    'Set Global DType for Digital Signature
    g_lngDType = DType
    
    Set rstTemp = New ADODB.Recordset
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "QueueProp_QueueName AS PLDAPROP_QueueName "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "QueueProperties "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "QueueProp_Type = 2 AND QueueProp_Code ='" & QueueCode & "' "
    ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If (rstTemp.RecordCount > 0) Then
        strEDIUserName = FNullField(rstTemp.Fields("PLDAPROP_QueueName").Value)
    Else
        ADORecordsetClose rstTemp

        Exit Function
    End If
    ADORecordsetClose rstTemp
    
    'Get Recipient
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[Send PLDA Recipient Operational] AS Operational, "
        strCommand = strCommand & "[Send PLDA Recipient Test] AS Test, "
        strCommand = strCommand & "Type "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[Logical ID] INNER JOIN " & IIf(DType = 14, "[PLDA Import] ON [PLDA Import]", "[PLDA Combined] ON [PLDA Combined]") & ".[LOGID Description] = [Logical ID].[LOGID Description] "
        strCommand = strCommand & "WHERE Code = '" & Code & "'"
    ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    If (rstTemp.RecordCount > 0) Then
        If (rstTemp![Type] = "O") Then
            strSendRecipient = FNullField(rstTemp.Fields("Operational").Value)
        Else
            strSendRecipient = FNullField(rstTemp.Fields("Test").Value)
        End If
    Else
        ADORecordsetClose rstTemp
        
        Exit Function
    End If
    ADORecordsetClose rstTemp
    
    
    Set rstLogID = New ADODB.Recordset
    Set rstHeader = New ADODB.Recordset
    Set rstHeaderZegels = New ADODB.Recordset
    Set rstHeaderZekerheid = New ADODB.Recordset
    Set rstHeaderTransitOffices = New ADODB.Recordset
    
    Set rstDetails = New ADODB.Recordset
    Set rstDetailsBijzondere = New ADODB.Recordset
    Set rstDetailsContainer = New ADODB.Recordset
    Set rstDetailsDocumenten = New ADODB.Recordset
    Set rstDetailsZelf = New ADODB.Recordset
    Set rstDetailsSensitiveGoods = New ADODB.Recordset
    
    Set rstHeaderHandelaars = New ADODB.Recordset
    Set rstDetailsHandelaars = New ADODB.Recordset
    Set rstDetailsBerekeningsEenheden = New ADODB.Recordset
    
    Set rstPLDAMessages = New ADODB.Recordset
    
    If (DType = 14) Then
        Set CImpHeader = New CIHeader
    ElseIf DType = 18 Then
        Set CExpHeader = New CEHeader
    End If
    Set CSummary = New CISummary
    Set CPLDADetails = New CDetails
        
    'Reset Indicator for Digital Signature
    g_blnDigiSignActivated = False
        
    '*****************************************************************************************************************
    ' 1. Prepare SQL commands for recordset to use in Creation of EDI Message
    ' 2. Get Last EDI Reference used
    ' 3. Get Digital Signature Option
    ' 4. Update the last EDI Reference
    '*****************************************************************************************************************
    Select Case DType
        Case 14
            strSQLH = "SELECT * FROM [PLDA Import Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Import Header Zegels] WHERE Code = '" & Code & "'"
            strSQLD = "SELECT * FROM [PLDA Import Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Import Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Import Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Import Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDZelf = "SELECT * FROM [PLDA Import Detail Zelf] WHERE Code = '" & Code & "' ORDER BY Detail"
            
            strSQLHHandelaars = "SELECT * FROM [PLDA Import Header Handelaars] WHERE Code = '" & Code & "'"
            strSQLDHandelaars = "SELECT * FROM [PLDA Import Detail Handelaars] WHERE Code = '" & Code & "'"
            
            strSQLDBerekeningsEenheden = "SELECT * FROM [PLDA Import Detail Berekenings Eenheden] WHERE Code = '" & Code & "'"
            
            
            ADORecordsetOpen "SELECT [LOGID Description], [Detail] FROM [PLDA Import] WHERE Code = '" & Code & "'", g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
            If rstTemp.RecordCount > 0 Then
                
                ADORecordsetOpen "SELECT A2, TIN, [Last EDI Reference], [DIGITAL SIGNATURE OPTION], [DIGITAL SIGNATURE SELECTED] FROM [Logical ID] WHERE [LOGID Description] = '" & FNullField(rstTemp.Fields("LogID Description").Value) & "'", g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
                
                If rstLogID.RecordCount > 0 Then
                    strAcct49 = rstLogID.Fields("A2").Value
                    strLogTIN = rstLogID.Fields("TIN").Value
                    
                    'Get the Last EDI Reference
                    strLastEDIReference = rstLogID.Fields("Last EDI Reference").Value
                    
                    g_intDigitalSignatureType = rstLogID.Fields("DIGITAL SIGNATURE OPTION").Value
                    
                    If g_intDigitalSignatureType = DigitalSignatureType.Fixed Then
                        g_strCertificateToUse = FNullField(rstLogID.Fields("DIGITAL SIGNATURE SELECTED").Value)
                        g_blnDigiSignActivated = True
                    ElseIf g_intDigitalSignatureType = DigitalSignatureType.[User Defined] Then
                        'Get the user-defined certificate
                        GetDigitalCertificateToUse DataSourceProperties, Code, "[PLDA Import]", g_strCertificateToUse, g_intDigitalSignatureType
                        g_blnDigiSignActivated = True
                    End If
                    
                End If
                
                'Update the Last EDI Reference
                If CLng(rstLogID.Fields("Last EDI Reference").Value) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(rstLogID.Fields("Last EDI Reference").Value) + 1
                End If
                rstLogID.Update
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, rstLogID, "Logical ID"
                
                ADORecordsetClose rstLogID
            End If
            
            ' CSCLP-513 + CSCLP-514
            'csclp-473
            'lngDetailCountMainTable = rstTemp.Fields("Detail").Value
            
            ADORecordsetClose rstTemp
                                            
        Case 18
            strSQLH = "SELECT * FROM [PLDA Combined Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Combined Header Zegels] WHERE Code = '" & Code & "'"
            strSQLHZekerheid = "SELECT * FROM [PLDA Combined Header Zekerheid] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            strSQLHTransitOffices = "SELECT * FROM [PLDA Combined Header Transit Offices] WHERE Code = '" & Code & "' ORDER BY Ordinal"
            
            strSQLD = "SELECT * FROM [PLDA Combined Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Combined Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Combined Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Combined Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDSensitiveGoods = "SELECT * FROM [PLDA Combined Detail Sensitive Goods] WHERE Code = '" & Code & "' ORDER BY Detail"
        
            strSQLPLDAMessages = "SELECT * FROM [PLDA Messages] WHERE Code ='" & Code & "' AND Message_Request_Type = 1 AND DType =" & DType
            
            strSQLHHandelaars = "SELECT * FROM [PLDA Combined Header Handelaars] WHERE Code = '" & Code & "'"
            strSQLDHandelaars = "SELECT * FROM [PLDA Combined Detail Handelaars] WHERE Code = '" & Code & "'"
            
            ADORecordsetOpen "SELECT [LOGID Description],[Detail] FROM [PLDA Combined] WHERE Code = '" & Code & "'", g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
            If rstTemp.RecordCount > 0 Then
            
                ADORecordsetOpen "SELECT A2, TIN, [Last EDI Reference], [DIGITAL SIGNATURE OPTION], [DIGITAL SIGNATURE SELECTED] FROM [Logical ID] WHERE [LOGID Description] = '" & FNullField(rstTemp.Fields("LogID Description").Value) & "'", g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
                
                If rstLogID.RecordCount > 0 Then
                    strAcct49 = rstLogID.Fields("A2").Value
                    strLogTIN = rstLogID.Fields("TIN").Value
                    
                    'Get the Last EDI Reference
                    strLastEDIReference = rstLogID.Fields("Last EDI Reference").Value
                    
                    g_intDigitalSignatureType = rstLogID.Fields("DIGITAL SIGNATURE OPTION").Value
                    
                    'Get Digital Certificate to use
                    If g_intDigitalSignatureType = DigitalSignatureType.Fixed Then
                        g_strCertificateToUse = FNullField(rstLogID.Fields("DIGITAL SIGNATURE SELECTED").Value)
                        g_blnDigiSignActivated = True
                    ElseIf g_intDigitalSignatureType = DigitalSignatureType.[User Defined] Then
                        'Get the user-defined certificate
                        GetDigitalCertificateToUse DataSourceProperties, Code, "[PLDA Combined]", g_strCertificateToUse, g_intDigitalSignatureType
                        g_blnDigiSignActivated = True
                    End If
                    
                End If
                
                'Update the Last EDI Reference
                If CLng(rstLogID.Fields("Last EDI Reference").Value) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(rstLogID.Fields("Last EDI Reference").Value) + 1
                End If
                rstLogID.Update
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, rstLogID, "Logical ID"
                
                ADORecordsetClose rstLogID
            End If
            
            ' CSCLP-513 + CSCLP-514
            'csclp-473
            'lngDetailCountMainTable = rstTemp.Fields("Detail").Value
            
            ADORecordsetClose rstTemp
    End Select
            
    '*****************************************************************************************************************
    ' 5. Open Offline recordsets for Creation of EDI messages
    '*****************************************************************************************************************
    ADORecordsetOpen strSQLH, g_conSADBEL, rstHeader, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHZegels, g_conSADBEL, rstHeaderZegels, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLD, g_conSADBEL, rstDetails, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDBijzondere, g_conSADBEL, rstDetailsBijzondere, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDContainer, g_conSADBEL, rstDetailsContainer, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDDocumenten, g_conSADBEL, rstDetailsDocumenten, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHHandelaars, g_conSADBEL, rstHeaderHandelaars, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDHandelaars, g_conSADBEL, rstDetailsHandelaars, adOpenKeyset, adLockOptimistic
    
    'RstOpen strSQLH, g_conSADBEL, rstHeader, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHZegels, g_conSADBEL, rstHeaderZegels, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLD, g_conSADBEL, rstDetails, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDBijzondere, g_conSADBEL, rstDetailsBijzondere, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDContainer, g_conSADBEL, rstDetailsContainer, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDDocumenten, g_conSADBEL, rstDetailsDocumenten, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHHandelaars, g_conSADBEL, rstHeaderHandelaars, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDHandelaars, g_conSADBEL, rstDetailsHandelaars, adOpenKeyset, adLockOptimistic, , True
    
    Set g_rstDetailsHandelaars = rstDetailsHandelaars   'This is for MProcedures.
    Set g_rstDetails = rstDetails                       'This is for MProcedures.
    
    If DType = 14 Then
        ADORecordsetOpen strSQLDBerekeningsEenheden, g_conSADBEL, rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDZelf, g_conSADBEL, rstDetailsZelf, adOpenKeyset, adLockOptimistic
        
        'RstOpen strSQLDBerekeningsEenheden, g_conSADBEL, rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDZelf, g_conSADBEL, rstDetailsZelf, adOpenKeyset, adLockOptimistic, , True
        
    ElseIf DType = 18 Then
        ADORecordsetOpen strSQLHZekerheid, g_conSADBEL, rstHeaderZekerheid, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLHTransitOffices, g_conSADBEL, rstHeaderTransitOffices, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDSensitiveGoods, g_conSADBEL, rstDetailsSensitiveGoods, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLPLDAMessages, g_conSADBEL, rstPLDAMessages, adOpenKeyset, adLockOptimistic
        
        'RstOpen strSQLHZekerheid, g_conSADBEL, rstHeaderZekerheid, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLHTransitOffices, g_conSADBEL, rstHeaderTransitOffices, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDSensitiveGoods, g_conSADBEL, rstDetailsSensitiveGoods, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLPLDAMessages, g_conSADBEL, rstPLDAMessages, adOpenKeyset, adLockOptimistic, , True
    
    End If
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' CSCLP-513 + CSCLP-514
    '*****************************************************************************************************************
    Set rstDetailsCount = rstDetails.Clone
   
    lngDetailCountMainTable = 0
    
    If rstDetailsCount.RecordCount > 0 Then
        rstDetailsCount.MoveFirst
        Do Until rstDetailsCount.EOF
            lngDetailCountMainTable = lngDetailCountMainTable + 1
            
            If UCase(Trim(rstDetailsCount.Fields("T7").Value)) = "F" Then
                Exit Do
            End If
            rstDetailsCount.MoveNext
        Loop
    End If
    '*****************************************************************************************************************
        
    '*****************************************************************************************************************
    'Initiliaze DigiSign Field Verifier / Offline Recordset for DigiSign
    '*****************************************************************************************************************
    If g_blnDigiSignActivated = True Then
        Call BuildDigisignFieldVerifier(g_conSADBEL, DType, CallingForm)
        Call InitializeRecordsToDigiSign(CallingForm)
    End If
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 6. Create EDI Header
    '*****************************************************************************************************************
    If DType = 14 Then
        strMessage = CImpHeader.CreateHeader(lngSegments, _
                                             rstHeader, _
                                             rstHeaderZegels, _
                                             rstHeaderHandelaars, _
                                             enuOriginal, _
                                             strEDIUserName, _
                                             strLogTIN, _
                                             strSendRecipient, _
                                             strLastEDIReference)
        
    ElseIf DType = 18 Then
        strMessage = CExpHeader.CreateHeader(lngSegments, _
                                             rstHeader, _
                                             rstHeaderZegels, _
                                             rstHeaderHandelaars, _
                                             rstHeaderZekerheid, _
                                             rstHeaderTransitOffices, _
                                             rstPLDAMessages, _
                                             rstDetails, _
                                             enuOriginal, _
                                             strEDIUserName, _
                                             strLogTIN, _
                                             strSendRecipient, _
                                             strLastEDIReference, Code)
    End If
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 7. Create EDI Detail
    '*****************************************************************************************************************
    lngNumberOfDetails = 0
    strMessage = strMessage & CPLDADetails.CreateDetails(lngSegments, _
                                                         rstHeader, _
                                                         rstDetails, _
                                                         rstDetailsBijzondere, _
                                                         rstDetailsContainer, _
                                                         rstDetailsDocumenten, _
                                                         rstDetailsZelf, _
                                                         rstDetailsBerekeningsEenheden, _
                                                         rstDetailsSensitiveGoods, _
                                                         DType, _
                                                         lngNumberOfDetails, Code)
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 8. Add string to be digitally signed
    '*****************************************************************************************************************
    If g_blnDigiSignActivated = True Then
        Set g_clsSignData = New CDataToSign
        Set clsDigisign = New CDigiSignFields
        
        Call BuildWholeDigitalSignature(DType)
        
        g_clsSignData.StringToSign = clsDigisign.StringToSign(g_rstDigiSignData)
    End If
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 9. Create EDI Summary
    '*****************************************************************************************************************
    lngSealsCount = 0
    If rstHeaderZegels.RecordCount <> 0 Then
        rstHeaderZegels.MoveFirst
        Do While Not rstHeaderZegels.EOF
            If LenB(Trim$(FNullField(rstHeaderZegels.Fields("E1").Value))) > 0 Or _
               LenB(Trim$(FNullField(rstHeaderZegels.Fields("E2").Value))) > 0 Then
                lngSealsCount = lngSealsCount + 1
            End If
            
            If rstHeaderZegels.Fields("E3").Value = "E" Then Exit Do

            rstHeaderZegels.MoveNext
        Loop
    End If
    
    
    
    strMessage = strMessage & CSummary.CreateSummary(lngSegments, _
                                                     rstHeader, _
                                                     lngSealsCount, _
                                                     dblNetWt, _
                                                     strLogTIN, _
                                                     DType, _
                                                     strLastEDIReference, _
                                                     lngDetailCountMainTable, _
                                                     enuOriginal)
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 10. Return Created EDI Message
    '*****************************************************************************************************************
    CreateEDIMessage = strMessage
    '*****************************************************************************************************************
    
    '*****************************************************************************************************************
    ' 11. Close and Destroy Recordsets to Free Memory
    '*****************************************************************************************************************
    ADORecordsetClose rstTemp
    ADORecordsetClose rstLogID
    
    ADORecordsetClose rstHeader
    ADORecordsetClose rstHeaderZegels
    ADORecordsetClose rstHeaderZekerheid
    ADORecordsetClose rstHeaderTransitOffices
    
    ADORecordsetClose rstDetails
    ADORecordsetClose rstDetailsBijzondere
    ADORecordsetClose rstDetailsContainer
    ADORecordsetClose rstDetailsDocumenten
    ADORecordsetClose rstDetailsZelf
    ADORecordsetClose rstDetailsSensitiveGoods
    ADORecordsetClose rstDetailsBerekeningsEenheden
    
    ADORecordsetClose rstHeaderHandelaars
    ADORecordsetClose rstDetailsHandelaars
    
    ADORecordsetClose rstPLDAMessages
    '*****************************************************************************************************************
    
End Function


Private Function GetDigitalCertificateToUse(ByRef DataSourceProperties As CDataSourceProperties, _
                                            ByVal UniqueCode As String, _
                                            ByVal TableName As String, _
                                            ByRef strCertificateToUse As String, _
                                            ByRef intDigitalSignatureType As Integer) As String
    
    Dim conTemplate As ADODB.Connection
    Dim rstUser As ADODB.Recordset
    
    Dim strCommand As String
    Dim lngUserNo As String
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[USER NO] "
        strCommand = strCommand & "FROM " & TableName & " "
        strCommand = strCommand & "WHERE Code = '" & Replace(UniqueCode, "'", "''") & "'"
    ADORecordsetOpen strCommand, g_conSADBEL, rstUser, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstUser, adOpenKeyset, adLockOptimistic, , True
    
    If (rstUser.RecordCount > 0) Then
        lngUserNo = FNullField(rstUser.Fields("USER NO").Value)
    End If
    
    ADORecordsetClose rstUser
    
    
    If (lngUserNo > 0) Then
        'Search USER ID and get the selected digital signature type
        
        ADOConnectDB conTemplate, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
        
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "[DIGITAL SIGNATURE OPTION], "
            strCommand = strCommand & "[DIGITAL SIGNATURE SELECTED] "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "USERS "
            strCommand = strCommand & "WHERE [USER_ID] = " & lngUserNo
        
        ADORecordsetOpen strCommand, conTemplate, rstUser, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, conTemplate, rstUser, adOpenKeyset, adLockOptimistic, , True
        
        If (rstUser.RecordCount > 0) Then
            intDigitalSignatureType = IIf(IsNull(rstUser.Fields("DIGITAL SIGNATURE OPTION").Value), 0, rstUser.Fields("DIGITAL SIGNATURE OPTION").Value)
            strCertificateToUse = IIf(IsNull(rstUser.Fields("DIGITAL SIGNATURE SELECTED").Value), "", Trim(rstUser.Fields("DIGITAL SIGNATURE SELECTED").Value))
        End If
        
        ADORecordsetClose rstUser
        
        ADODisconnectDB conTemplate
        
    Else
        
        intDigitalSignatureType = 0
        strCertificateToUse = ""
        
    End If
    
End Function
