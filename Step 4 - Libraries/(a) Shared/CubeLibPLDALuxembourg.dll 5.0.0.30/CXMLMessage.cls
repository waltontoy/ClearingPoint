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

Public Enum m_enumXMLMessage
    enumExportAmendment = 513
    enumExportDeclaration = 515
    enumImportAmendment = 613
    enumImportDeclaration = 615
    enumCancellationRequest = 914
End Enum


'03242008 - Additional param
Public Function CreateXMLMessage(ByVal Code As String, _
                                 ByVal DType As Long, _
                                 ByRef conConnection As ADODB.Connection, _
                                 ByVal strMDBPath As String, _
                                 ByVal XMLMessageType As m_enumXMLMessage, _
                        Optional ByVal LogIDDescription As String, _
                        Optional ByVal QueueCode As String = "PLDA Luxembourg") As String
    
    Dim rstTemp As ADODB.Recordset
    Dim rstLogID As ADODB.Recordset
        
    Dim strSQLMain As String
    Dim strSQLH As String
    Dim strSQLHHandelaars As String
    Dim strSQLHZegels As String
    Dim strSQLD As String
    Dim strSQLDBerekeningsEenheden As String
    Dim strSQLDBijzondere As String
    Dim strSQLDContainer As String
    Dim strSQLDDocumenten As String
    Dim strSQLDHandelaars As String
    Dim strSQLDZelf As String

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
            
    Dim strCommand As String
            
    '***********************************************************************************************
    'Check if for cancellation
    '***********************************************************************************************
    If XMLMessageType = enumCancellationRequest Then
        
        Set rstTemp = New ADODB.Recordset
        
        strCommand = vbNullString
        strCommand = "SELECT * "
        strCommand = strCommand & "FROM [PLDA MESSAGES] "
        strCommand = strCommand & "WHERE Code = '" & Code & "' "
        strCommand = strCommand & "AND Message_StatusType = 'Queued' "
        strCommand = strCommand & "AND Message_Request_Type = 1"
        
        ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
        With rstTemp
            '.open strCommand, conConnection, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                G_strCancelReason = !Message_Reason
            Else
                .Close
                Set rstTemp = Nothing
                Exit Function
            End If
            
            .Close
            Set rstTemp = Nothing
        End With
        
    End If
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Check if for amendment
    '***********************************************************************************************
    If XMLMessageType = enumImportAmendment Or XMLMessageType = enumExportAmendment Then
    
        Set rstTemp = New ADODB.Recordset
        
        strCommand = vbNullString
        strCommand = "SELECT * "
        strCommand = strCommand & "FROM [PLDA MESSAGES] "
        strCommand = strCommand & "WHERE Code = '" & Code & "' "
        strCommand = strCommand & "AND Message_StatusType = 'Queued' "
        strCommand = strCommand & "AND Message_Request_Type = 2"
        
        ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
        With rstTemp
            '.open strCommand, conConnection, adOpenKeyset, adLockOptimistic
        
            If .RecordCount > 0 Then
                'lngDType = rstTemp!DType
            Else
                .Close
                Set rstTemp = Nothing
                Exit Function
            End If
            
            .Close
            Set rstTemp = Nothing
        End With
        
    End If
    '***********************************************************************************************
    
    G_strIEFunctionCode = "IE" & XMLMessageType
        
    '***********************************************************************************************
    'Get XML Username
    '***********************************************************************************************
'    strCommand = vbNullString
'    strCommand = strCommand & "SELECT "
'    strCommand = strCommand & "PLDAPROP_QueueName "
'    strCommand = strCommand & "FROM "
'    strCommand = strCommand & "MDB_PLDAProperties "
'    strCommand = strCommand & "WHERE "
'    strCommand = strCommand & "PLDAPROP_Type = 2"

    strCommand = vbNullString
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "QueueProp_QueueName AS PLDAPROP_QueueName "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "QueueProperties "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "QueueProp_Type =2 "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "QueueProp_Code ='" & QueueCode & "'"
    
    ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    If (rstTemp.RecordCount > 0) Then
        G_strXMLUserName = Trim$(FNullField(rstTemp.Fields("PLDAPROP_QueueName").Value))
    Else
        ADORecordsetClose rstTemp
        Exit Function
    End If
    
    ADORecordsetClose rstTemp
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Set Message Sender and Message Recipient
    'G_lngTestOnly is always set to zero according to Frank
    '***********************************************************************************************
    G_lngTestOnly = 0
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "[Send PLDA Sender Operational] AS MessageSenderOperational, "
    strCommand = strCommand & "[Send PLDA Sender Test] AS MessageSenderTest, "
    strCommand = strCommand & "[Send PLDA Recipient Operational] AS MessageRecipientOperational, "
    strCommand = strCommand & "[Send PLDA Recipient Test] AS MessageRecipientTest, "
    strCommand = strCommand & "Type "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[Logical ID] "
    strCommand = strCommand & "INNER JOIN " & IIf(DType = 14, "[PLDA Import] ON [PLDA Import]", "[PLDA Combined] ON [PLDA Combined]") & ".[LOGID Description] = [Logical ID].[LOGID Description] "
    strCommand = strCommand & "WHERE Code = '" & Code & "'"
    
    ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    If rstTemp.RecordCount > 0 Then
        If rstTemp.Fields("Type").Value = "O" Then
            G_strMessageSender = Trim$(FNullField(rstTemp.Fields("MessageSenderOperational").Value))
            G_strMessageRecipient = Trim$(FNullField(rstTemp.Fields("MessageRecipientOperational").Value))
        Else
            G_strMessageSender = Trim$(FNullField(rstTemp.Fields("MessageSenderTest").Value))
            G_strMessageRecipient = Trim$(FNullField(rstTemp.Fields("MessageRecipientTest").Value))
        End If
    Else
        ADORecordsetClose rstTemp
        Exit Function
    End If
    
    ADORecordsetClose rstTemp
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Set Recordset for Authorisation Fields - '03242008
    '***********************************************************************************************
    Set G_rstLogIDFields = GetDataFromLogIDTable(conConnection, LogIDDescription)
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Set Recordset Queries and incrementing of EDI Reference
    '***********************************************************************************************
    Select Case DType
        Case 14
            
            strSQLMain = "SELECT * FROM [PLDA Import] WHERE Code = '" & Code & "'"
            
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
            
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "[LOGID Description] "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[PLDA Import] "
            strCommand = strCommand & "WHERE Code = '" & Code & "'"
            
            ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "A2, "
                strCommand = strCommand & "TIN, "
                strCommand = strCommand & "[Last EDI Reference] "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[Logical ID] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "[LOGID Description] = '" & rstTemp.Fields("LogID Description").Value & "'"
                
                ADORecordsetOpen strCommand, conConnection, rstLogID, adOpenKeyset, adLockOptimistic
                'RstOpen strCommand, conConnection, rstLogID, adOpenKeyset, adLockOptimistic, , True
                
                If rstLogID.RecordCount > 0 Then
                    strAcct49 = Trim$(FNullField(rstLogID.Fields("A2").Value))
                    strLogTIN = Trim$(FNullField(rstLogID.Fields("TIN").Value))
                    strLastEDIReference = Trim$(FNullField(rstLogID.Fields("Last EDI Reference").Value))
                End If
                
                If CLng(rstLogID.Fields("Last EDI Reference").Value) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(rstLogID.Fields("Last EDI Reference").Value) + 1
                End If
                
                rstLogID.Update
                
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, conConnection, rstLogID, "Logical ID"
                
                ADORecordsetClose rstLogID
            End If
            
            ADORecordsetClose rstTemp
            
        Case 18
            
            strSQLH = "SELECT * FROM [PLDA Combined Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Combined Header Zegels] WHERE Code = '" & Code & "'"
            strSQLD = "SELECT * FROM [PLDA Combined Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Combined Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Combined Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Combined Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
        
            strSQLHHandelaars = "SELECT * FROM [PLDA Combined Header Handelaars] WHERE Code = '" & Code & "'"
            
            strSQLDHandelaars = "SELECT * FROM [PLDA Combined Detail Handelaars] WHERE Code = '" & Code & "'"
            
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "[LOGID Description] "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[PLDA Combined] "
            strCommand = strCommand & "WHERE Code = '" & Code & "'"
            
            ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "A2, "
                strCommand = strCommand & "TIN, "
                strCommand = strCommand & "[Last EDI Reference] "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[Logical ID] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "[LOGID Description] = '" & rstTemp.Fields("LogID Description").Value & "'"
                   
                ADORecordsetOpen strCommand, conConnection, rstLogID, adOpenKeyset, adLockOptimistic
                'RstOpen strCommand, conConnection, rstLogID, adOpenKeyset, adLockOptimistic, , True
                
                If rstLogID.RecordCount > 0 Then
                    strAcct49 = Trim$(FNullField(rstLogID.Fields("A2").Value))
                    strLogTIN = Trim$(FNullField(rstLogID.Fields("TIN").Value))
                    strLastEDIReference = Trim$(FNullField(rstLogID.Fields("Last EDI Reference").Value))
                End If
                
                If CLng(rstLogID.Fields("Last EDI Reference").Value) >= 99999 Then
                    rstLogID.Fields("Last EDI Reference").Value = "1"
                Else
                    rstLogID.Fields("Last EDI Reference").Value = CLng(rstLogID.Fields("Last EDI Reference").Value) + 1
                End If
                
                rstLogID.Update
                
                ExecuteRecordset ExecuteRecordsetConstant.Update, conConnection, rstLogID, "Logical ID"
                
                ADORecordsetClose rstLogID
            End If
            
            ADORecordsetClose rstTemp
        
        Case Else
            Debug.Assert False
    
    End Select
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Open Offline Recordsets
    '***********************************************************************************************
    ADORecordsetOpen strSQLH, conConnection, G_rstHeader, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHZegels, conConnection, G_rstHeaderZegels, adOpenKeyset, adLockOptimistic
    
    ADORecordsetOpen strSQLD, conConnection, G_rstDetails, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDBijzondere, conConnection, G_rstDetailsBijzondere, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDContainer, conConnection, G_rstDetailsContainer, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDDocumenten, conConnection, G_rstDetailsDocumenten, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHHandelaars, conConnection, G_rstHeaderHandelaars, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDHandelaars, conConnection, G_rstDetailsHandelaars, adOpenKeyset, adLockOptimistic
    
    
    'RstOpen strSQLH, conConnection, G_rstHeader, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHZegels, conConnection, G_rstHeaderZegels, adOpenKeyset, adLockOptimistic, , True
    
    'RstOpen strSQLD, conConnection, G_rstDetails, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDBijzondere, conConnection, G_rstDetailsBijzondere, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDContainer, conConnection, G_rstDetailsContainer, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDDocumenten, conConnection, G_rstDetailsDocumenten, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLHHandelaars, conConnection, G_rstHeaderHandelaars, adOpenKeyset, adLockOptimistic, , True
    'RstOpen strSQLDHandelaars, conConnection, G_rstDetailsHandelaars, adOpenKeyset, adLockOptimistic, , True
    
    If DType = 14 Then
        ADORecordsetOpen strSQLMain, conConnection, G_rstMain, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDBerekeningsEenheden, conConnection, G_rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic
        ADORecordsetOpen strSQLDZelf, conConnection, G_rstDetailsZelf, adOpenKeyset, adLockOptimistic
        
        'RstOpen strSQLMain, conConnection, G_rstMain, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDBerekeningsEenheden, conConnection, G_rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic, , True
        'RstOpen strSQLDZelf, conConnection, G_rstDetailsZelf, adOpenKeyset, adLockOptimistic, , True
    End If
    '***********************************************************************************************
        
    '***********************************************************************************************
    'Start creating the XML data here...
    '***********************************************************************************************
    CreateXMLMessage = CreateXML(DType, XMLMessageType)
    '***********************************************************************************************
    
    '***********************************************************************************************
    'Close Recordsets
    '***********************************************************************************************
    ADORecordsetClose G_rstMain
    ADORecordsetClose G_rstHeader
    ADORecordsetClose G_rstHeaderZegels
    ADORecordsetClose G_rstDetails
    ADORecordsetClose G_rstDetailsBijzondere
    ADORecordsetClose G_rstDetailsContainer
    ADORecordsetClose G_rstDetailsDocumenten
    ADORecordsetClose G_rstHeaderHandelaars
    ADORecordsetClose G_rstDetailsHandelaars
    
    If DType = 14 Then
        ADORecordsetClose G_rstDetailsBerekeningsEenheden
        ADORecordsetClose G_rstDetailsZelf
    End If
    '***********************************************************************************************
    
End Function

