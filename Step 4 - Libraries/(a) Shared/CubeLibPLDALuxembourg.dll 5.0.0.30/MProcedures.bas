Attribute VB_Name = "MProcedures"
Option Explicit

Public Function CreateXML(ByVal DType As Long, ByVal XMLMessageType As m_enumXMLMessage) As String
    
    Dim objDOM As DOMDocument
    Dim objParentNode As IXMLDOMNode
    Dim objChildNode As IXMLDOMNode
    
    Set objDOM = New DOMDocument
    
    objDOM.async = False
    objDOM.validateOnParse = False
    objDOM.resolveExternals = False
    objDOM.preserveWhiteSpace = False
    
    ' Create a processing instruction targeted for xml
    Set objParentNode = objDOM.createProcessingInstruction("xml", "version='1.0'")
    
    objDOM.appendChild objParentNode
    Set objParentNode = Nothing
        
    ' Create the root node
    If XMLMessageType = enumCancellationRequest Then
        Set objParentNode = objDOM.appendChild(objDOM.createElement("PLDAResponse"))
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        Call CreateCancellationXML(DType, XMLMessageType, objDOM, objParentNode, objChildNode)
        
    Else
    
        If DType = 14 Then
            Set objParentNode = objDOM.appendChild(objDOM.createElement("PLDAImportDV1"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateImportXML(XMLMessageType, objDOM, objParentNode, objChildNode)
        ElseIf DType = 18 Then
            Set objParentNode = objDOM.appendChild(objDOM.createElement("PLDAExport"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateExportXML(XMLMessageType, objDOM, objParentNode, objChildNode)
        End If
                        
    End If
        
    CreateXML = objDOM.xml
    
    ''*********************************************************************************
    ''FOR TESTING ONLY
    ''*********************************************************************************
    'Dim strTemp As String
    'strTemp = objDOM.xml
    '
    'Open App.Path & "/test_" & DType & ".xml" For Output As #1
    'Print #1, strTemp
    'Close #1
    ''*********************************************************************************
    
    Set objChildNode = Nothing
    Set objParentNode = Nothing
    Set objDOM = Nothing
                    
End Function


Public Function GenerateData(ByRef CurrentTable As ADODB.Recordset, _
                             ByVal strField As String, _
                    Optional ByVal lngPos As Long = 1) As String
            
    Dim rstCurrentTable As ADODB.Recordset
    Dim lngDetail As Long
    Dim strFilter As String
    
    Set rstCurrentTable = CurrentTable.Clone
    
    'Check first if Field is existing on the Current Recordset
    If Not IsFieldExisting(rstCurrentTable, strField) Then
        Debug.Print "Selected field is not present on table: " & strField
        Exit Function
    End If
    
    '********************************************************************************************
    'Add filter on Detail Number and/or Ordinal depending on the situation
    '   1. Detail only for non-grouped fields on Detail tab
    '   2. Ordinal only for grouped fields on header tab
    '   3. Detail and Ordinal for grouped fields on detail tab
    '********************************************************************************************
    rstCurrentTable.Filter = adFilterNone
    
    If IsFieldExisting(CurrentTable, "Detail") Then
        lngDetail = CurrentTable.Fields("Detail").Value
        strFilter = "Detail = " & lngDetail & " "
        
        If IsFieldExisting(CurrentTable, "Ordinal") Then
            strFilter = strFilter & "AND Ordinal = " & lngPos & " "
        End If
    Else
        If IsFieldExisting(CurrentTable, "Ordinal") Then
            strFilter = "Ordinal = " & lngPos & " "
        End If
    End If
    
    If strFilter <> vbNullString Then
        rstCurrentTable.Filter = strFilter
    End If
    '********************************************************************************************
    
    'If a record is found, return value
    If rstCurrentTable.RecordCount > 0 Then
        GenerateData = Trim$(FNullField(rstCurrentTable.Fields(strField).Value))
    End If
    
End Function


Public Function IsFieldExisting(ByVal rstCurrentTable As ADODB.Recordset, _
                                ByVal FieldName As String) As Boolean
    
    Dim strFieldName As String
    
    On Error Resume Next
    strFieldName = rstCurrentTable.Fields(FieldName).Name
    IsFieldExisting = (Err.Number = 0)
    On Error GoTo 0
   
End Function

Public Function GetDataFromLogIDTable(ByRef conConnection As ADODB.Connection, _
                                      ByVal LogIDDescription As String) As ADODB.Recordset

    Dim strCommand As String
    Dim rstTemp As New ADODB.Recordset
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[LOGICAL ID] "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "[LOGID DESCRIPTION] = '" & LogIDDescription & "' "
    
    ADORecordsetOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, conConnection, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    Set GetDataFromLogIDTable = rstTemp
    
End Function
