VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum CustomsDocumentTypeConstant
    [SADBEL Import Document] = 1
    [SADBEL Export Document] = 2
    [SADBEL Transit Document] = 3
    [SADBEL Import Document Duplicata] = 4
    [SADBEL Export Document Duplicata] = 5
    [SADBEL Transit Document Duplicata] = 6
    [Transit NCTS Document] = 7
    [Transit NCTS Document Duplicata] = 8
    [Combined NCTS Document] = 9
    [Combined NCTS Document Duplicata] = 10
    [NCTS Departure Document] = 11
    [NCTS Arrival Document] = 12
    [PLDA Import Document] = 14
    [PLDA Import Document Duplicata] = 15
    [PLDA Export Document] = 16
    [PLDA Export Document Duplicata] = 17
    [PLDA Combined Document] = 18
    [PLDA Combined Document Duplicata] = 19
    [PLDA Volgbriefje] = 20 'Edwin Sept 28
    [PLDA EAD] = 21 'Edwin Sept 28
    
    [PLDA DV1] = 22 'CSCLP-638
    [PLDA EUR1] = 23
    [PLDA ATR] = 24
    [PLDA Import Codisheet] = 25
    [PLDA Combined Codisheet] = 26
    [PLDA Volgbriefje Import] = 27
End Enum


Private Enum MessagingTypeConstant
    [EDI Messaging] = 1
    [XML Messaging] = 2
    [EDI Follow-up Request Messaging] = 3 'CSCLP-578
    [Unknown] = 99
End Enum

Public Enum PLDABeXmlMsgTypeConstant
    [PLDA Import AC4] = 1
    [PLDA Import AC4 Cancellation] = 2
End Enum

Private Function GetMessagingType(ByRef ADODBSadbel As ADODB.Connection, _
                                  ByVal LogID_Desc As String, _
                                  ByVal DType As CustomsDocumentTypeConstant, _
                         Optional ByRef SendMode As Long) As MessagingTypeConstant 'NCTSCustomsLocalSystemConstant
        
    Dim strCommand As String
    Dim rstTemp As ADODB.Recordset
    Dim CustomsDocumentType As CustomsDocumentTypeConstant
    
    If Len(Trim$(LogID_Desc)) = 0 Then
        GetMessagingType = -99
        Exit Function
    End If
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[LOGICAL ID] "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "[LogID Description] = '" & LogID_Desc & "' "
    
    ADORecordsetOpen strCommand, ADODBSadbel, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, ADODBSadbel, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    If rstTemp.RecordCount > 0 Then
        SendMode = IIf(FNullField(rstTemp.Fields("SEND MODE").Value) = "", 2, FNullField(rstTemp.Fields("SEND MODE").Value))
        
        Select Case DType
            Case 11 To 12 'NCTS
                If IsNull(rstTemp.Fields("Messaging Type").Value) Then
                    rstTemp.Fields("Messaging Type").Value = MessagingTypeConstant.[EDI Messaging]
                End If
                
                If rstTemp.Fields("Messaging Type").Value < 2 Then
                    GetMessagingType = MessagingTypeConstant.[EDI Messaging]
                Else
                    GetMessagingType = MessagingTypeConstant.[XML Messaging]
                End If
            Case 14 To 21 'PLDA
                If IsNull(rstTemp.Fields("PLDA_MessagingType").Value) Then
                    rstTemp.Fields("PLDA_MessagingType").Value = MessagingTypeConstant.[EDI Messaging]
                End If
                
                If rstTemp.Fields("PLDA_MessagingType").Value < 2 Then
                    GetMessagingType = MessagingTypeConstant.[EDI Messaging]
                Else
                    GetMessagingType = MessagingTypeConstant.[XML Messaging]
                End If
        End Select
    
    Else
        GetMessagingType = MessagingTypeConstant.Unknown
    End If
           
End Function

Private Function GetPLDAQueuePropertiesToUse(ByRef ADODBSadbel As ADODB.Connection, _
                                             ByVal LogID As String, _
                                             ByVal DType As Long, _
                                             ByVal InfoType As String) As String
    
    Dim lngSendMode As Long
    
    Select Case GetMessagingType(ADODBSadbel, LogID, DType, lngSendMode)
        Case MessagingTypeConstant.[EDI Messaging]
        
            Select Case lngSendMode
                Case 0, 1
                    If InfoType = "T" Then
                        GetPLDAQueuePropertiesToUse = "PLDA Belgium Test"
                    Else
                        GetPLDAQueuePropertiesToUse = "PLDA Belgium"
                    End If
                    
                Case 2
                    GetPLDAQueuePropertiesToUse = "PLDA Belgium Test"
                    
            End Select
            
        Case MessagingTypeConstant.[XML Messaging]
        
            GetPLDAQueuePropertiesToUse = ""
            
    End Select
    
End Function

Public Function CreateXMLMessage(ByRef DataSourceProperties As ADODB.Connection, _
                                 ByVal UniqueCode As String, _
                                 ByVal InfoType As String, _
                                 ByVal LogID As String, _
                                 ByVal DType As Long, _
                                 ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant) As String

    Dim strCommand As String
    
    Dim strEDIUserName As String
    Dim strSendRecipient As String
    Dim strMessage As String
    Dim strReason As String
    
    Dim rstTemp As ADODB.Recordset
    Dim rstUseSFTP As ADODB.Recordset
    Dim strQuePropCode As String
    Dim blnUseSFTP As Boolean

    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
    
    strQuePropCode = GetPLDAQueuePropertiesToUse(g_conSADBEL, LogID, DType, InfoType)
    
    ' ---------------------------------------------------------------------------------- '
    ' Get Cancellation Reason if for Cancellation
    ' ---------------------------------------------------------------------------------- '
    If PLDABeXmlMsgType = [PLDA Import AC4 Cancellation] Then
            strCommand = vbNullString
            strCommand = strCommand & "SELECT * FROM [PLDA MESSAGES] "
            strCommand = strCommand & "WHERE Code = '" & UniqueCode & "' "
            strCommand = strCommand & "AND Message_StatusType = 'Queued' "
            strCommand = strCommand & "AND Message_Request_Type = 1"
        
        ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
        With rstTemp
            If .RecordCount > 0 Then
                strReason = rstTemp.Fields("Message_Reason").Value
            Else
                ADORecordsetClose rstTemp
                Exit Function
            End If
            
            ADORecordsetClose rstTemp
        End With
    End If
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    ' ---------------------------------------------------------------------------------- '
    ' GET EDI USERNAME
    ' ---------------------------------------------------------------------------------- '
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "SFTP_ENABLED AS SFTP_ENABLED "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "QueueProperties "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "QueueProp_Type = 1 "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "QueueProp_Code ='" & strQuePropCode & "' "
    
    ADORecordsetOpen strCommand, g_conSADBEL, rstUseSFTP, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstUseSFTP, adOpenKeyset, adLockOptimistic

    If (rstUseSFTP.RecordCount > 0) Then
        
        blnUseSFTP = FNullField(rstUseSFTP.Fields("SFTP_ENABLED").Value)
    Else
        ADORecordsetClose rstUseSFTP

        Exit Function
    End If
    ADORecordsetClose rstUseSFTP
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "QueueProp_QueueName AS QueueProp_QueueName, "
        strCommand = strCommand & "SFTP_QueueProp_QueueName AS SFTP_QueueProp_QueueName "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "QueueProperties "
        strCommand = strCommand & "WHERE "
        If blnUseSFTP Then
            strCommand = strCommand & "QueueProp_Type = 2 "
        Else
            strCommand = strCommand & "QueueProp_Type = 1 "
        End If
        strCommand = strCommand & "AND "
        strCommand = strCommand & "QueueProp_Code ='" & strQuePropCode & "' "
    ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic

    If (rstTemp.RecordCount > 0) Then
    
        If blnUseSFTP Then
            g_strXmlSender = FNullField(rstTemp.Fields("SFTP_QueueProp_QueueName").Value)
        Else
            g_strXmlSender = FNullField(rstTemp.Fields("QueueProp_QueueName").Value)
        End If
        
    Else
        ADORecordsetClose rstTemp

        Exit Function
    End If
    ADORecordsetClose rstTemp
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    
    ' ---------------------------------------------------------------------------------- '
    ' GET RECIPIENT
    ' ---------------------------------------------------------------------------------- '
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[Send PLDA Recipient Operational] AS Operational, "
        strCommand = strCommand & "[Send PLDA Recipient Test] AS Test, "
        strCommand = strCommand & "Type "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[Logical ID] "
        strCommand = strCommand & "INNER JOIN "
        strCommand = strCommand & IIf(DType = 14, "[PLDA Import] ", "[PLDA Combined] ")
        strCommand = strCommand & "ON "
        strCommand = strCommand & IIf(DType = 14, "[PLDA Import]", "[PLDA Combined]")
        strCommand = strCommand & ".[LOGID Description] = [Logical ID].[LOGID Description] "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "Code = '" & UniqueCode & "'"
    ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    If (rstTemp.RecordCount > 0) Then
        If (rstTemp.Fields("Type").Value = "O") Then
            g_strXmlRecipient = FNullField(rstTemp.Fields("Operational").Value)
        Else
            g_strXmlRecipient = FNullField(rstTemp.Fields("Test").Value)
        End If
    Else
        ADORecordsetClose rstTemp
        
        Exit Function
    End If
    ADORecordsetClose rstTemp
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    ' ---------------------------------------------------------------------------------- '
    ' 1. Prepare SQL commands for recordset to use in Creation of EDI Message
    ' ---------------------------------------------------------------------------------- '
    Select Case DType
        Case 14
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Header] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlHeader, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlHeader, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Header Zegels] "
                strCommand = strCommand & "WHERE Code = '" & UniqueCode & "' "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlHeaderZegels, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlHeaderZegels, adOpenKeyset, adLockOptimistic, , True
                
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "Detail "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetails, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetails, adOpenKeyset, adLockOptimistic, , True
            
            If Not (g_rstXmlDetails.EOF And g_rstXmlDetails.BOF) Then
                
                g_lngTotalItemsF = 0
                
                g_rstXmlDetails.MoveFirst
                Do Until g_rstXmlDetails.EOF
                    
                    g_lngTotalItemsF = g_lngTotalItemsF + 1
                    
                    If UCase$(Trim$(FNullField(g_rstXmlDetails.Fields("T7").Value))) = "F" Then
                        Exit Do
                    End If
                    
                    g_rstXmlDetails.MoveNext
                Loop
                
                g_rstXmlDetails.MoveFirst
            End If
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Bijzondere] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "Detail "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsBijzondere, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsBijzondere, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Container] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "Detail "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsContainer, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsContainer, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Documenten] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "Detail "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsDocumenten, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsDocumenten, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Zelf] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "Detail "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsZelf, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsZelf, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Header Handelaars] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlHeaderHandelaars, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlHeaderHandelaars, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Handelaars] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsHandelaars, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsHandelaars, adOpenKeyset, adLockOptimistic, , True
            
                strCommand = ""
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[PLDA Import Detail Berekenings Eenheden] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "Code = '" & UniqueCode & "' "
            ADORecordsetOpen strCommand, g_conSADBEL, g_rstXmlDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, g_conSADBEL, g_rstXmlDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic, , True
    End Select
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    ' ---------------------------------------------------------------------------------- '
    ' 6. Create XML Message
    ' ---------------------------------------------------------------------------------- '
    Select Case PLDABeXmlMsgType
    
        Case PLDABeXmlMsgTypeConstant.[PLDA Import AC4]
            strMessage = CreateXML(PLDABeXmlMsgTypeConstant.[PLDA Import AC4])
        
        Case PLDABeXmlMsgTypeConstant.[PLDA Import AC4 Cancellation]
            strMessage = CreateXML(PLDABeXmlMsgTypeConstant.[PLDA Import AC4 Cancellation], strReason)
            
        Case Else
            Debug.Assert False
    End Select
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    ' ---------------------------------------------------------------------------------- '
    ' 10. Return Created XML Message
    ' ---------------------------------------------------------------------------------- '
    CreateXMLMessage = strMessage
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
    
    ' ---------------------------------------------------------------------------------- '
    ' 11. Close and Destroy Recordsets to Free Memory
    ' ---------------------------------------------------------------------------------- '
    ADORecordsetClose rstTemp
    
    ADORecordsetClose g_rstXmlHeader
    ADORecordsetClose g_rstXmlHeaderZegels
    
    ADORecordsetClose g_rstXmlDetails
    ADORecordsetClose g_rstXmlDetailsBijzondere
    ADORecordsetClose g_rstXmlDetailsContainer
    ADORecordsetClose g_rstXmlDetailsDocumenten
    ADORecordsetClose g_rstXmlDetailsZelf
    ADORecordsetClose g_rstXmlDetailsBerekeningsEenheden
    
    ADORecordsetClose g_rstXmlHeaderHandelaars
    ADORecordsetClose g_rstXmlDetailsHandelaars
    ' ---------------------------------------------------------------------------------- '
    ' ---------------------------------------------------------------------------------- '
End Function

