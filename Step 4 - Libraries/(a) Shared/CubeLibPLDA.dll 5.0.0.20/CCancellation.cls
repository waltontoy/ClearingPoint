VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DOCUMENT_TYPE_IMPORT  As Long = 14
Private Const DOCUMENT_TYPE_COMBINED  As Long = 18

Public Function CreateCancellationMessage(ByRef DataSourceProperties As CDataSourceProperties, _
                                          ByVal Code As String, _
                                 Optional ByVal QueueCode As String = "PLDA Belgium") As String
    
    
    Dim rstTemp As ADODB.Recordset
    Dim rstLogID As ADODB.Recordset
    
    Dim lngDType As Long
    Dim strDocType As String
    
    Dim strSender As String     'Sender
    Dim strRecipient As String  'Recipient
    Dim strInterchangeReference As String   'Interchange Reference
    
    Dim strMRN As String    'MRN
    Dim strLRN As String    'LRN
    Dim strDeclarationType As String   'Declaration Type
    Dim strLanguage As String   'Language of the Declaration
    Dim strReason As String     'Reason for Cancellation
    Dim strTempSQL As String
    
    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
    
    'Get DType, Reason ----------------------------------------
        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT * FROM [PLDA MESSAGES] "
        strTempSQL = strTempSQL & "WHERE Code = '" & Code & "' "
        strTempSQL = strTempSQL & "AND Message_StatusType = 'Queued' "
        strTempSQL = strTempSQL & "AND Message_Request_Type = 1"
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    With rstTemp
        If .RecordCount > 0 Then
            lngDType = rstTemp.Fields("DType").Value
            strReason = rstTemp.Fields("Message_Reason").Value
        Else
            ADORecordsetClose rstTemp
            Exit Function
        End If
        
        ADORecordsetClose rstTemp
    End With
    '----------------------------------------------------------
    
    
    'Get Sender -----------------------------------------------
        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT "
        strTempSQL = strTempSQL & "QueueProp_QueueName AS PLDAPROP_QueueName "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "QueueProperties "
        strTempSQL = strTempSQL & "WHERE QueueProp_Type = 2 AND QueueProp_Code = '" & QueueCode & "'"
    
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    With rstTemp
        'If .RecordCount > 0 Then
        If Not (.EOF And .BOF) Then
            If Not IsNull(rstTemp.Fields("PLDAPROP_QueueName").Value) Then
                strSender = rstTemp.Fields("PLDAPROP_QueueName").Value
            Else
                ADORecordsetClose rstTemp
    
                Exit Function
            End If
        Else
            ADORecordsetClose rstTemp
            Exit Function
        End If
        
        ADORecordsetClose rstTemp
    End With
    '----------------------------------------------------------
    

    
    'Get Recipient --------------------------------------------
        strTempSQL = vbNullString
        strTempSQL = strTempSQL & "SELECT "
        strTempSQL = strTempSQL & "[Send PLDA Recipient Operational] AS Operational, "
        strTempSQL = strTempSQL & "[Send PLDA Recipient Test] AS Test, "
        strTempSQL = strTempSQL & "Type "
        strTempSQL = strTempSQL & "FROM "
        strTempSQL = strTempSQL & "[Logical ID] "
        strTempSQL = strTempSQL & "INNER JOIN "
        strTempSQL = strTempSQL & IIf(lngDType = DOCUMENT_TYPE_IMPORT, "[PLDA Import] ON [PLDA Import]", "[PLDA Combined] ON [PLDA Combined]") & ".[LOGID Description] = [Logical ID].[LOGID Description] WHERE Code = '" & Code & "'"
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    With rstTemp
        If .RecordCount > 0 Then
            If rstTemp.Fields("Type").Value = "O" And Not IsNull(rstTemp.Fields("Operational").Value) Then
                strRecipient = rstTemp.Fields("Operational").Value
                
            ElseIf rstTemp.Fields("Type").Value = "T" And Not IsNull(rstTemp.Fields("Test").Value) Then
                strRecipient = rstTemp.Fields("Test").Value
                
            Else
                ADORecordsetClose rstTemp
                Exit Function
            End If
        Else
            ADORecordsetClose rstTemp
            Exit Function
        End If
        
        ADORecordsetClose rstTemp
    End With
    '----------------------------------------------------------
    
    'Get MRN, LRN ---------------------------------------------
        If lngDType = DOCUMENT_TYPE_IMPORT Then
            strTempSQL = "SELECT MRN, A1, A2, A3, A9 FROM [PLDA IMPORT HEADER] WHERE Code = '" & Code & "'"
        ElseIf lngDType = DOCUMENT_TYPE_COMBINED Then
            strTempSQL = "SELECT MRN, A1, A2, A3, A9 FROM [PLDA COMBINED HEADER] WHERE Code = '" & Code & "'"
        End If
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    With rstTemp
        If .RecordCount > 0 Then
            strMRN = FNullField(rstTemp.Fields("MRN").Value)
            strLRN = FNullField(rstTemp.Fields("A3").Value)
            strDeclarationType = FNullField(rstTemp.Fields("A1").Value) & FNullField(rstTemp.Fields("A2").Value)
            strLanguage = FNullField(rstTemp.Fields("A9").Value)
        Else
            ADORecordsetClose rstTemp
            Exit Function
        End If
        
        ADORecordsetClose rstTemp
    End With
    '----------------------------------------------------------
    
    'Get Interchage Reference ---------------------------------
        If lngDType = DOCUMENT_TYPE_IMPORT Then
            strTempSQL = "SELECT [LOGID Description] FROM [PLDA Import] WHERE Code = '" & Code & "'"
        ElseIf lngDType = DOCUMENT_TYPE_COMBINED Then
            strTempSQL = "SELECT [LOGID Description] FROM [PLDA Combined] WHERE Code = '" & Code & "'"
        End If
    ADORecordsetOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strTempSQL, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If rstTemp.RecordCount > 0 Then
            strTempSQL = vbNullString
            strTempSQL = strTempSQL & "SELECT "
            strTempSQL = strTempSQL & "[TIN], "
            strTempSQL = strTempSQL & "[Last EDI Reference] "
            strTempSQL = strTempSQL & "FROM [Logical ID] "
            strTempSQL = strTempSQL & "WHERE [LOGID Description] = '" & rstTemp.Fields("LogID Description").Value & "'"
        
        ADORecordsetOpen strTempSQL, g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
        'RstOpen strTempSQL, g_conSADBEL, rstLogID, adOpenKeyset, adLockOptimistic
        
        With rstLogID
            If .RecordCount > 0 Then
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
                End If
                
                ADORecordsetClose rstLogID
            End If
        End With
    End If
    ADORecordsetClose rstLogID
    '----------------------------------------------------------
    
    If lngDType = DOCUMENT_TYPE_IMPORT Then
        strDocType = "IMPORT"
    ElseIf lngDType = DOCUMENT_TYPE_COMBINED Then
        strDocType = "EXPORT"
    End If
    
    'Create and Return the message
    CreateCancellationMessage = CancelMessage(strSender, _
                                              strRecipient, _
                                              strInterchangeReference, _
                                              strDocType, _
                                              strMRN, _
                                              strReason, _
                                              strLRN, _
                                              strDeclarationType, _
                                              strLanguage)
End Function

Private Property Get CancelMessage(ByVal Sender As String, _
                                   ByVal Recipient As String, _
                                   ByVal InterchangeReference As String, _
                                   ByVal DType As String, _
                                   ByVal MRN As String, _
                                   ByVal Reason As String, _
                                   ByVal LRN As String, _
                                   ByVal DecType As String, _
                                   ByVal Language As String) As String
    Dim strMessage As String
    
    'UNB
    strMessage = "UNB+UNOA:2+" & ReplaceSpecialCharacters(Sender) & "+" & _
                                ReplaceSpecialCharacters(Recipient) & "+" & _
                                Format(Now, "yyMMdd") & ":" & _
                                Format(Now, "hhmm") & "+" & _
                                ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    'UNH
    strMessage = strMessage & "UNH+1+CUSDEC:D:04A:UN:" & ReplaceSpecialCharacters(DType) & "+" & _
                                                        ReplaceSpecialCharacters(MRN) & "'"
    
    'BGM
    strMessage = strMessage & "BGM+:::" & ReplaceSpecialCharacters(DecType) & "+" & _
                                            ReplaceSpecialCharacters(MRN) & ":" & _
                                            ReplaceSpecialCharacters(Language) & "+" & _
                                            ReplaceSpecialCharacters(enuCancellation) & "'"
    
    'DTM
    strMessage = strMessage & "DTM+137:" & Format(Now, "yyyyMMdd") & ":102'"
    
    'FTX - Modified for JIRA CSCLP-476
    'strMessage = strMessage & "FTX+ACD+++" & ReplaceSpecialCharacters(Reason) & "'"
    strMessage = strMessage & "FTX+ACD++NIHIL+" & ReplaceSpecialCharacters(Reason) & "'"
    
    'RFF
    strMessage = strMessage & "RFF+ABE:" & ReplaceSpecialCharacters(LRN) & "'"
    
    'UNS
    strMessage = strMessage & "UNS+D'"
    
    strMessage = strMessage & "UNS+S'"
    
    'UNT
    strMessage = strMessage & "UNT+8+1'"
    
    'UNZ
    strMessage = strMessage & "UNZ+1+" & ReplaceSpecialCharacters(InterchangeReference) & "'"
    
    CancelMessage = strMessage
    
End Property


