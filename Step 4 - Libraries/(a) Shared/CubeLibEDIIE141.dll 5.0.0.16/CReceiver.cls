VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_dtmReceived As Date

Public Sub ProcessFollowRequestMessage(ByRef DataSourceProperties As CDataSourceProperties, _
                                       ByVal Message As String, _
                                       ByVal DBPath As String, _
                                       ByRef CallingForm As Object)
    
    Dim lngDummyMessageStart As Long
    Dim lngDataNCTSID As Long
    
    Dim strErrorReply As String
    Dim strMessageReceived As String
    Dim strCode As String
    Dim strMessageReference As String
    
    On Error GoTo ErrHandler
    
    Set G_CallingForm = CallingForm
    
    m_dtmReceived = Now()
    
    Call OpenConnections(DataSourceProperties, DBPath)
    
    G_CallingForm.AddtoTraceFollowUpReq "MessageReceived, Message to prepare:" & vbCrLf & Message & vbCrLf, False
    
    lngDummyMessageStart = InStr(1, Message, "Error reply:", vbTextCompare)
     
    'Extract Message
    If lngDummyMessageStart > 0 Then
        strMessageReceived = Mid(Message, lngDummyMessageStart)
          
        lngDummyMessageStart = InStr(1, strMessageReceived, "Original message:", vbTextCompare)
          
        If lngDummyMessageStart > 0 Then
            strErrorReply = Mid(strMessageReceived, 1, lngDummyMessageStart - 1)
        End If
    Else
        
        lngDummyMessageStart = InStr(1, Message, "UNB+UNOC:3", vbTextCompare)

        'strMessageReceived = Mid(Message, lngDummyMessageStart)
        'strMessageReceived = Trim(strMessageReceived)
        
        ' MUCP-87 - Start
        If InStr(1, Message, "CUSRES:D:96B:UN:CD906A") > 0 Then
            strErrorReply = Trim$(Mid(Message, lngDummyMessageStart))
            
            lngDummyMessageStart = InStr(1, strErrorReply, vbNullChar)
      
            'Strip terminating NULL character
            If lngDummyMessageStart > 0 Then
                strErrorReply = Left(strErrorReply, lngDummyMessageStart - 1)
            End If
        Else
            strMessageReceived = Mid(Message, lngDummyMessageStart)
            strMessageReceived = Trim(strMessageReceived)
        End If
        ' MUCP-87 - End
    End If
      
    lngDummyMessageStart = InStr(1, strMessageReceived, vbNullChar)
      
    'Strip terminating NULL character
    If lngDummyMessageStart > 0 Then
        strMessageReceived = Left(strMessageReceived, lngDummyMessageStart - 1)
    End If
    
    'Strip terminating carriage return and line feed
    If Right(strMessageReceived, 2) = vbCrLf Then
        strMessageReceived = Left(strMessageReceived, Len(strMessageReceived) - 2)
    End If
    
    If LenB(Trim$(strMessageReceived)) > 0 Then
        strMessageReference = GetMessageReference(strMessageReceived)
    Else
        strMessageReference = GetMessageReference(strErrorReply)
    End If
    
    'Update DATA_NCTS_MESSAGES and Get Code and DATA_NCTS_ID field value
    strCode = GetCodeAndUpdateNCTSMessages(strMessageReference, IIf(Len(strErrorReply) > 0, strErrorReply, strMessageReceived), lngDataNCTSID)
    
    'Update DATA_NCTS and MasterEDINCTSIE141
    UpdateDataNCTS lngDataNCTSID, strErrorReply
    UpdateMasterEDINCTSIE strCode, strErrorReply
    
    Call CloseConnections(DBPath)
    
ErrHandler:
    Select Case Err.Number
  Case 0
      'Do Nothing
      
  Case Else
      G_CallingForm.AddtoTraceFollowUpReq "Function: ProcessFollowRequestMessage" & vbCrLf, False
      G_CallingForm.AddtoTraceFollowUpReq "Error: " & Err.Number & " - " & Err.Description & vbCrLf, False
      
    End Select

End Sub


Private Sub OpenConnections(ByRef DataSourceProperties As CDataSourceProperties, _
                            ByVal DBPath As String)

    ADOConnectDB G_conEdifact, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    ADOConnectDB G_conData, DataSourceProperties, DBInstanceType_DATABASE_DATA
    
    'ConnectDB G_conEdifact, DBPath, "edifact.mdb"
    'ConnectDB G_conData, DBPath, "mdb_data.mdb"
    
End Sub


Private Sub CloseConnections(ByVal DBPath As String)
    
    ADODisconnectDB G_conEdifact
    ADODisconnectDB G_conData
    
End Sub

Private Function GetCodeAndUpdateNCTSMessages(ByVal MessageReference As String, ByVal Message As String, ByRef DataNCTSID As Long) As String
    
    Dim rstNCTSMessages As ADODB.Recordset
    Dim rstNCTSData As ADODB.Recordset
    
    Dim strCommand As String
    
    Dim lngData_NCTS_ID As Long
            
    strCommand = vbNullString
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "[NCTS_IEM_ID], "
    strCommand = strCommand & "[DATA_NCTS_MSG_Message], "
    strCommand = strCommand & "[DATA_NCTS_MSG_Reference], "
    strCommand = strCommand & "[DATA_NCTS_ID], "
    strCommand = strCommand & "[DATA_NCTS_MSG_Date], "
    strCommand = strCommand & "[DATA_NCTS_MSG_StatusType], "
    strCommand = strCommand & "[User_ID] "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[DATA_NCTS_MESSAGES] "
    strCommand = strCommand & "WHERE DATA_NCTS_MSG_Reference = '" & MessageReference & "' "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "DATA_NCTS_MSG_StatusType = 'Sent' "

    
    strCommand = strCommand & "ORDER BY DATA_NCTS_MSG_Date DESC "
    
    ADORecordsetOpen strCommand, G_conEdifact, rstNCTSMessages, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conEdifact, rstNCTSMessages, adOpenKeyset, adLockOptimistic
    
    If rstNCTSMessages.RecordCount > 0 Then
    
        lngData_NCTS_ID = FNullField(rstNCTSMessages.Fields("DATA_NCTS_ID").Value)
        
        ' MUCP-87 - Start
'        If FNullField(rstNCTSMessages.Fields("NCTS_IEM_ID").Value) = 34 Then    ' Follow-up Request
'
'            strCommand = vbNullString
'            strCommand = strCommand & "SELECT * "
'            strCommand = strCommand & "FROM "
'            strCommand = strCommand & "[DATA_NCTS] "
'            strCommand = strCommand & "WHERE "
'            strCommand = strCommand & "DATA_NCTS_ID = '" & lngData_NCTS_ID & "' "
'            strCommand = strCommand & "AND "
'            strCommand = strCommand & "NCTS_IEM_ID = " & 5 & " "
'            strCommand = strCommand & "AND "
'            strCommand = strCommand & "DATA_NCTS_MSG_StatusType = 'Sent' "
'
'            RstOpen strCommand, G_conEdifact, rstNCTSData, adOpenKeyset, adLockOptimistic
'
'        End If
        ' MUCP-87 - End
            
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[CODE] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[DATA_NCTS] "
        strCommand = strCommand & "WHERE DATA_NCTS_ID =" & lngData_NCTS_ID
            
        ADORecordsetOpen strCommand, G_conEdifact, rstNCTSData, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, G_conEdifact, rstNCTSData, adOpenKeyset, adLockOptimistic
        
        If rstNCTSData.RecordCount > 0 Then
            rstNCTSMessages.AddNew
            
            ' MUCP-87 - Start
            'rstNCTSMessages.Fields("NCTS_IEM_ID").Value = 29
            If InStr(1, Message, "CUSRES:D:96B:UN:CD906A") > 0 Then
                rstNCTSMessages.Fields("NCTS_IEM_ID").Value = 25
            Else
                rstNCTSMessages.Fields("NCTS_IEM_ID").Value = 29
            End If
            ' MUCP-87 - End
            
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Message").Value = Message
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Reference").Value = MessageReference
            rstNCTSMessages.Fields("DATA_NCTS_ID").Value = lngData_NCTS_ID
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Date").Value = m_dtmReceived
            rstNCTSMessages.Fields("DATA_NCTS_MSG_StatusType").Value = "Received"
            rstNCTSMessages.Fields("User_ID").Value = 0
            rstNCTSMessages.Update
                
            InsertRecordset G_conEdifact, rstNCTSMessages, "DATA_NCTS_MESSAGES"
            
            DataNCTSID = lngData_NCTS_ID
            GetCodeAndUpdateNCTSMessages = Trim$(FNullField(rstNCTSData.Fields("CODE").Value))
        End If
    Else
        
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[DATA_NCTS_ID] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[DATA_NCTS] "
        strCommand = strCommand & "WHERE MRN = '" & MessageReference & "' "
        
        ADORecordsetOpen strCommand, G_conEdifact, rstNCTSData, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, G_conEdifact, rstNCTSData, adOpenKeyset, adLockOptimistic
        
        If rstNCTSData.RecordCount > 0 Then
            
            lngData_NCTS_ID = FNullField(rstNCTSData.Fields("DATA_NCTS_ID").Value)
            
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "[NCTS_IEM_ID], "
            strCommand = strCommand & "[DATA_NCTS_MSG_Message], "
            strCommand = strCommand & "[DATA_NCTS_MSG_Reference], "
            strCommand = strCommand & "[DATA_NCTS_ID], "
            strCommand = strCommand & "[DATA_NCTS_MSG_Date], "
            strCommand = strCommand & "[DATA_NCTS_MSG_StatusType], "
            strCommand = strCommand & "[User_ID] "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[DATA_NCTS_MESSAGES] "
                
            ADORecordsetOpen strCommand, G_conEdifact, rstNCTSMessages, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, G_conEdifact, rstNCTSMessages, adOpenKeyset, adLockOptimistic
            
            rstNCTSMessages.AddNew
            rstNCTSMessages.Fields("NCTS_IEM_ID").Value = 33
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Message").Value = Message
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Reference").Value = MessageReference
            rstNCTSMessages.Fields("DATA_NCTS_ID").Value = lngData_NCTS_ID
            rstNCTSMessages.Fields("DATA_NCTS_MSG_Date").Value = m_dtmReceived
            rstNCTSMessages.Fields("DATA_NCTS_MSG_StatusType").Value = "Received"
            rstNCTSMessages.Fields("User_ID").Value = 0
            rstNCTSMessages.Update
            
            InsertRecordset G_conEdifact, rstNCTSMessages, "DATA_NCTS_MESSAGES"
            
            DataNCTSID = lngData_NCTS_ID
            GetCodeAndUpdateNCTSMessages = Trim$(FNullField(rstNCTSData.Fields("CODE").Value))
        End If
    End If
    
    ADORecordsetClose rstNCTSData
    ADORecordsetClose rstNCTSMessages
    
    'RstClose rstNCTSData
    'RstClose rstNCTSMessages
    
End Function


Private Function GetMessageReference(ByVal Message As String) As String
       
    Dim lngBGM As Long
    Dim strTemp As String
    
    lngBGM = InStr(1, Message, "BGM", vbTextCompare)
    
    If lngBGM > 0 Then
        strTemp = Mid(Message, lngBGM)
        
        lngBGM = InStr(1, strTemp, "'", vbTextCompare)
        
        If lngBGM > 0 Then
            strTemp = Mid(strTemp, 1, lngBGM - 1)
            strTemp = Replace(strTemp, "BGM++", "")
            
            ' MUCP-87 - Start
            If InStr(1, strTemp, "+") > 0 Then
                strTemp = Left$(strTemp, InStr(1, strTemp, "+") - 1)
            End If
            ' MUCP-87 - End
            GetMessageReference = strTemp
        End If
    End If
    
End Function


Private Sub UpdateDataNCTS(ByVal DataNCTSID As Long, ByVal ErrorMessage As String)
    
    Dim strCommand As String
    Dim rstTemp As ADODB.Recordset
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "[Tree ID], "
    strCommand = strCommand & "[Error String], "
    strCommand = strCommand & "[Date Last Received] "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[DATA_NCTS] "
    strCommand = strCommand & "WHERE DATA_NCTS_ID = " & DataNCTSID
    
    ADORecordsetOpen strCommand, G_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
        
        If Len(ErrorMessage) > 0 Then
            rstTemp.Fields("Tree ID").Value = "58ED"
            rstTemp.Fields("Error String").Value = ErrorMessage
        Else
            rstTemp.Fields("Tree ID").Value = "56ED"
        End If
        
        rstTemp.Fields("Date Last Received").Value = m_dtmReceived
        rstTemp.Update
        
        UpdateRecordset G_conEdifact, rstTemp, "DATA_NCTS"
    End If
    
    ADORecordsetClose rstTemp
End Sub


Private Sub UpdateMasterEDINCTSIE(ByVal Code As String, ByVal ErrorMessage As String)
    
    Dim strCommand As String
    Dim rstTemp As ADODB.Recordset
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "[Tree ID],"
    strCommand = strCommand & "[Date Last Received], "
    strCommand = strCommand & "[Viewed] "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[MASTEREDINCTS] "
    strCommand = strCommand & "WHERE CODE = '" & Code & "' "
    
    ADORecordsetOpen strCommand, G_conData, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conData, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
    
        If Len(ErrorMessage) > 0 Then
            rstTemp.Fields("Tree ID").Value = "58ED"
        Else
            rstTemp.Fields("Tree ID").Value = "56ED"
        End If
        
        rstTemp.Fields("Date Last Received").Value = m_dtmReceived
        rstTemp.Fields("Viewed").Value = 0
        rstTemp.Update
        
        UpdateRecordset G_conData, rstTemp, "MASTEREDINCTS"
    End If
    
    ADORecordsetClose rstTemp
End Sub

Private Sub Class_Terminate()
    
    Set G_CallingForm = Nothing
    
End Sub
