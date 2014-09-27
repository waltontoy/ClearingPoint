Attribute VB_Name = "MProcedures"
Option Explicit


Public Function GetMapFunctionValue(ByVal MapFunction As String) As String
    
    Dim strReturnValue As String
    Dim strCommand As String
    Dim strTIN As String
    
    Dim lngLastEDIReference As Long
    
    Dim rstTemp As ADODB.Recordset
    
    Select Case MapFunction
        Case "F<RECEIVE QUEUE>"
            strCommand = vbNullString
            strCommand = strCommand & "SELECT QueueProperties.QueueProp_QueueName "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[QueueProperties] "
            strCommand = strCommand & "INNER JOIN "
            strCommand = strCommand & "[LOGICAL ID] "
            strCommand = strCommand & "ON "
            strCommand = strCommand & "[QueueProperties].QueueProp_Code = [LOGICAL ID].[NCTS_QueuePropCode] "
            strCommand = strCommand & "WHERE [QueueProperties].[QueueProp_Type] = 2 "
            strCommand = strCommand & "AND [LOGICAL ID].[LOGID DESCRIPTION]= " & Chr(39) & ProcessQuotes(Trim$(FNullField(G_rstNCTSData.Fields("LogIDDescription").Value))) & Chr(39)
            
            ADORecordsetOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                strReturnValue = Trim$(FNullField(rstTemp.Fields("QueueProp_QueueName").Value))
            End If
            
            ADORecordsetClose rstTemp
            'RstClose rstTemp
            
        Case "F<RECIPIENT>"
            strCommand = vbNullString
            strCommand = strCommand & "SELECT * "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[LOGICAL ID] "
            strCommand = strCommand & "WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(Trim$(FNullField(G_rstNCTSData.Fields("LogIDDescription").Value))) & Chr(39)
            
            ADORecordsetOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                If Trim$(FNullField(G_rstNCTSData.Fields("SendMode").Value)) = "O" Then
                    strReturnValue = Trim$(FNullField(rstTemp.Fields("SEND EDI RECIPIENT OPERATIONAL").Value))
                Else
                    strReturnValue = Trim$(FNullField(rstTemp.Fields("SEND EDI RECIPIENT TEST").Value))
                End If
            End If
            
            ADORecordsetClose rstTemp
            'RstClose rstTemp
            
        Case "F<DATE, YYMMDD>"
            strReturnValue = Format(Date, "YYMMDD")
            
        Case "F<TIME, HHMM>"
            strReturnValue = Format(Time, "HHMM")
        
        Case "F<1 TIN REF>"
            strCommand = vbNullString
            strCommand = strCommand & "SELECT * "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[LOGICAL ID] "
            strCommand = strCommand & "WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(Trim$(FNullField(G_rstNCTSData.Fields("LogIDDescription").Value))) & Chr(39)
            
            ADORecordsetOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strCommand, G_conSadbel, rstTemp, adOpenKeyset, adLockOptimistic
            
            If rstTemp.RecordCount > 0 Then
                strTIN = Trim$(FNullField(rstTemp.Fields("TIN").Value))
                lngLastEDIReference = Val(Trim$(FNullField(rstTemp.Fields("LAST EDI REFERENCE").Value)))
            Else
                strTIN = ""
                lngLastEDIReference = 0
            End If
            
            lngLastEDIReference = lngLastEDIReference + 1
            
            If lngLastEDIReference > 99999 Then
                lngLastEDIReference = 0
            End If
            
            rstTemp.Fields("LAST EDI REFERENCE").Value = CStr(lngLastEDIReference)
            rstTemp.Update
            
            InsertRecordset G_conSadbel, rstTemp, "LOGICAL ID"
            
            ADORecordsetClose rstTemp
            'RstClose rstTemp
            
            strReturnValue = Format(Left(strTIN, 9) & Format(lngLastEDIReference, "00000"), Replace(Space(14), " ", "0"))
            
        Case "F<MESSAGE REFERENCE>"
            strReturnValue = "1"
            
    End Select
    
    GetMapFunctionValue = strReturnValue
    
End Function

Public Function GetValueFromNCTSFollowUpRequest(ByVal FieldName As String)
    
    Select Case FieldName
        Case "AN", "AR"
            If Len(Trim$(FNullField(G_rstFollowUpRequest.Fields(FieldName).Value))) > 0 Then
                GetValueFromNCTSFollowUpRequest = Trim$(FNullField(G_rstFollowUpRequest.Fields(FieldName).Value))
            Else
                GetValueFromNCTSFollowUpRequest = 0
            End If
            
        Case Else
            If Len(Trim$(FNullField(G_rstFollowUpRequest.Fields(FieldName).Value))) > 0 Then
                If Trim$(FNullField(G_rstFollowUpRequest.Fields(FieldName).Value)) <> "0" Then
                    GetValueFromNCTSFollowUpRequest = Trim$(FNullField(G_rstFollowUpRequest.Fields(FieldName).Value))
                End If
            End If
    
    End Select
    
End Function


Public Function ReplaceSpecialCharacters(ByVal SourceString As String) As String
    Dim strReturnValue As String
    strReturnValue = Trim(SourceString)
    If strReturnValue <> vbNullString Then
        '----->  RELEASE CHARACTER
        strReturnValue = Replace(strReturnValue, EDI_SEP_RELEASE_CHARACTER, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER)
        '----->  SEGMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_SEGMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
        '----->  COMPOSITE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_COMPOSITE_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT)
        '----->  SIMPLE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT)
    End If
    ReplaceSpecialCharacters = strReturnValue
End Function
