Attribute VB_Name = "MCancellationRequest"
Option Explicit

'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <objChildElement>
'               <objChildElement2>
'                   <objChildElement3>

Public Sub CreateXMLMessageIE14(ByRef DataSourceProperties As CDataSourceProperties, _
                                ByRef objDOM As DOMDocument, _
                                ByRef objParentNode As IXMLDOMNode, _
                                ByRef objChildNode As IXMLDOMNode)
    
    Dim objChildElement As IXMLDOMNode
    Dim objChildElement2 As IXMLDOMNode
    
    'Interchange
    CreateMessageInterchangeIE14 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Header
    CreateMessageHeaderIE14 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Principal
    CreateMessagePrincipalIE14 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Departure Customs Office
    CreateMessageDepartureOfficeIE14 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    
    Set objChildElement = Nothing
    Set objChildElement2 = Nothing
     
End Sub


Private Sub CreateMessageInterchangeIE14(ByRef objDOM As DOMDocument, _
                                         ByRef objParentNode As IXMLDOMNode, _
                                         ByRef objChildNode As IXMLDOMNode, _
                                         ByRef objChildElement As IXMLDOMNode, _
                                         ByRef objChildElement2 As IXMLDOMNode)

    'Syntax Identifier
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("SynIdeMES1"))
    objChildNode.Text = "UNOC"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Syntax Version Number
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("SynVerNumMES2"))
    objChildNode.Text = "3"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Message Sender
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesSenMES3"))
    objChildNode.Text = GetMapFunctionValue("F<RECEIVE QUEUE>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Sender Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("SenIdeCodQuaMES4"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Message Recipient
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesRecMES6"))
    objChildNode.Text = GetMapFunctionValue("F<RECIPIENT>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Recipient Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecIdeCodQuaMES7"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Date of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("DatOfPreMES9"))
    objChildNode.Text = GetMapFunctionValue("F<DATE, YYMMDD>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Time of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("TimOfPreMES10"))
    objChildNode.Text = GetMapFunctionValue("F<TIME, HHMM>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Interchange Control Reference
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("IntConRefMES11"))
    objChildNode.Text = GetMapFunctionValue("F<1 TIN REF>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Recipient's Reference/Password
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecRefMES12"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Recipient's Reference/Password Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecRefQuaMES13"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Application Reference
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("AppRefMES14"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Priority
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("PriMES15"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Acknowledgement Request
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("AckReqMES16"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Communications Agreement ID
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("ComAgrIdMES17"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Test Indicator
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("TesIndMES18"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Message Identification
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesIdeMES19"))
    objChildNode.Text = "1"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Message Type
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesTypMES20"))
    objChildNode.Text = "CC014A"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Common Access Reference
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("ComAccRefMES21"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Sequence Number - NOT ON EDIFACT XML MAPPING FOR CUSDEC
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesSeqNumMES22"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'First and Last Transaction - NOT ON EDIFACT XML MAPPING FOR CUSDEC
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("FirAndLasTraMES23"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub


Private Sub CreateMessageHeaderIE14(ByRef objDOM As DOMDocument, _
                               ByRef objParentNode As IXMLDOMNode, _
                               ByRef objChildNode As IXMLDOMNode, _
                               ByRef objChildElement As IXMLDOMNode, _
                               ByRef objChildElement2 As IXMLDOMNode)

    
    Dim strTemp As String
    
    'Header
    Set objChildNode = objDOM.createElement("HEAHEA")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Document Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DocNumHEA5"))
        objChildElement.Text = RetrieveRecordFromEDI("DocNumHEA5")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Date of Cancellation request
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DatOfCanReqHEA147"))
        objChildElement.Text = RetrieveRecordFromEDI("DatOfCanReqHEA147")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Cancellation Reason
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CanReaHEA250"))
        objChildElement.Text = RetrieveRecordFromEDI("CanReaHEA250")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Cancellation Reason language
        If Len(Trim$(GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0))) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CanReaHEA250LNG"))
            objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub


Private Sub CreateMessagePrincipalIE14(ByRef objDOM As DOMDocument, _
                                  ByRef objParentNode As IXMLDOMNode, _
                                  ByRef objChildNode As IXMLDOMNode, _
                                  ByRef objChildElement As IXMLDOMNode, _
                                  ByRef objChildElement2 As IXMLDOMNode)
    
    
    'Principal
    Set objChildNode = objDOM.createElement("TRAPRIPC1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamPC17"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X1", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Street and Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumPC122"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X2", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Postal Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodPC123"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X6", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'City
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitPC124"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X3", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouPC125"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X5", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'NAD Language
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGPC"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'TIN
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINPC159"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "X4", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Public Sub CreateMessageDepartureOfficeIE14(ByRef objDOM As DOMDocument, _
                                        ByRef objParentNode As IXMLDOMNode, _
                                        ByRef objChildNode As IXMLDOMNode, _
                                        ByRef objChildElement As IXMLDOMNode, _
                                        ByRef objChildElement2 As IXMLDOMNode)
    
    'Customs Offices of Departure
    Set objChildNode = objDOM.createElement("CUSOFFDEPEPT")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Reference Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumEPT1"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A4", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

End Sub

Private Function RetrieveRecordFromEDI(ByVal XMLNode As String) As String
    
    Dim arrEDISegments() As String
    Dim arrSubSegments() As String
    Dim arrDataSegments() As String
    Dim lngctr As Long
    Dim strSegmentToFind As String
    
    'Check what segment to find
    Select Case XMLNode
        Case "DocNumHEA5"
            strSegmentToFind = "BGM"
            
        Case "DatOfCanReqHEA147"
            strSegmentToFind = "DTM+318:"
            
        Case "CanReaHEA250"
            strSegmentToFind = "FTX+ACD"
            
        Case Else
            Exit Function
            
    End Select
    
    If Len(G_strEDICancellation) = 0 Then Exit Function
    
    arrEDISegments = Split(G_strEDICancellation, "'")
    
    For lngctr = 0 To UBound(arrEDISegments)
        If InStr(1, arrEDISegments(lngctr), strSegmentToFind) > 0 Then
            arrSubSegments = Split(arrEDISegments(lngctr), "+")
            
            Select Case XMLNode
                Case "DocNumHEA5"
                    If UBound(arrSubSegments) >= 2 Then
                        arrDataSegments = Split(arrSubSegments(2), ":")
                        
                        If UBound(arrDataSegments) >= 0 Then
                            RetrieveRecordFromEDI = arrDataSegments(0)
                        End If
                    End If
                    
                Case "DatOfCanReqHEA147"
                    If UBound(arrSubSegments) >= 1 Then
                        arrDataSegments = Split(arrSubSegments(1), ":")
                        
                        If UBound(arrDataSegments) >= 1 Then
                            RetrieveRecordFromEDI = arrDataSegments(1)
                        End If
                    End If
                    
                    
                Case "CanReaHEA250"
                    If UBound(arrSubSegments) >= 4 Then
                        arrDataSegments = Split(arrSubSegments(4), ":")
                        
                        If UBound(arrDataSegments) >= 0 Then
                            RetrieveRecordFromEDI = arrDataSegments(0)
                        End If
                    End If
                    
            End Select
            
        End If
    Next
    
End Function
