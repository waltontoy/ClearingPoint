Attribute VB_Name = "MCancellation"
Option Explicit

Public Sub CreateCancellationXML(ByVal DType As Long, _
                        ByVal XMLMessageType As m_enumXMLMessage, _
                        ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode)
    
    Dim objChildNode2 As IXMLDOMNode
    Dim objChildElement As IXMLDOMElement
    Dim objChildElement2 As IXMLDOMElement
    
    Call CreateCancellationInterchangeHeader(XMLMessageType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    Call CreateCancellationResponsePLDAHeader(DType, XMLMessageType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    Call CreateCancellationResponsePLDAItem(objDOM, objParentNode, objChildNode, objChildElement)
    Call CreateCancellationResponseCalculationResults(objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    
    Set objChildNode2 = Nothing
    Set objChildElement2 = Nothing
    Set objChildElement = Nothing
    
End Sub

Private Sub CreateCancellationInterchangeHeader(ByVal XMLMessageType As m_enumXMLMessage, ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode, _
                        ByRef objChildElement As IXMLDOMElement, _
                        ByRef objChildElement2 As IXMLDOMElement)

    '*****************
    'Interchange Header
    '*****************
    Set objChildNode = objDOM.createElement("InterchangeHeader")
    objDOM.documentElement.appendChild objChildNode
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageSender"))
        objChildElement.Text = G_strMessageSender   ' Client AS2 Identifier
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageRecipient"))
        objChildElement.Text = G_strMessageRecipient    ' PLDA AS2 Identifier
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageVersion"))
        objChildElement.Text = "V0.1"
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("testIndicator"))
        objChildElement.Text = G_lngTestOnly
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("functionCode"))
        objChildElement.Text = "IE" & XMLMessageType     ' Cancellation Request
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DateTimeOfPreparation"))
        
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("dateOfPreparation"))
            objChildElement2.Text = Format(Now, "YYYYMMDD")
            
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("timeOfPreparation"))
            objChildElement2.Text = Format(Now, "HHMMSS")
            
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub

Private Sub CreateCancellationResponsePLDAHeader(ByVal DType As Long, _
                        ByVal XMLMessageType As m_enumXMLMessage, _
                        ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode, _
                        ByRef objChildElement As IXMLDOMElement, _
                        ByRef objChildElement2 As IXMLDOMElement)
    
    '*****************
    'Response PLDA Header
    '*****************
    
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement
    Dim objChildElement5 As IXMLDOMElement
    
    Dim lngPos As Long
    
    'Row Position of Declarant on Header Handelaars
    If DType = 14 Then
        lngPos = 2
    ElseIf DType = 18 Then
        lngPos = 1
    End If
    
    Set objChildNode = objDOM.createElement("ResponsePLDAHeader")
    objDOM.documentElement.appendChild objChildNode
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("localReferenceNumber"))
        objChildElement.Text = GenerateData(G_rstHeader, "A3")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("customsMainReferenceNumber"))
        objChildElement.Text = GenerateData(G_rstHeader, "MRN")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("cancellationRequestDate"))
        objChildElement.Text = Format(Now, "YYYYMMDD")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("cancellationRequestTime"))
        objChildElement.Text = Format(Now, "HHMMSS")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("cancellationDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("cancellationTime"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("investigationStartDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("investigationStartTime"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentAcceptanceDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("rejectionDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CancellationReason"))
        objChildElement.Text = G_strCancelReason
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("rejectionElementPath"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("rejectionCode"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("registeredDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("registeredTime"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("presentationFlag"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("presentationRequiredDate"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
'
'        Set objChildElement = objChildNode.appendChild(objDOM.createElement("presentationRequiredTime"))
'        objChildElement.Text = ""
'        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("exportOffice"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("destinationOffice"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Declarant
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Declarant"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DeclarantOperator"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                'Operator Identity +
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorIdentity"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                                            
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("OperatorIdentity"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X1", lngPos)
                                
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Operator +
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Operator"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                    
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("operatorName"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X2", lngPos)
                                        
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("OperatorAddress"))
                    objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("postalCode"))
                        objChildElement5.Text = GenerateData(G_rstHeaderHandelaars, "X5", lngPos)

                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("streetAndNumber1"))
                        objChildElement5.Text = GenerateData(G_rstHeaderHandelaars, "X3", lngPos)

                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("streetAndNumber2"))
                        objChildElement5.Text = GenerateData(G_rstHeaderHandelaars, "X4", lngPos)

                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("city"))
                        objChildElement5.Text = GenerateData(G_rstHeaderHandelaars, "X6", lngPos)

                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("country"))
                        objChildElement5.Text = GenerateData(G_rstHeaderHandelaars, "X8", lngPos)

                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    Set objChildElement3 = Nothing
    Set objChildElement4 = Nothing
    Set objChildElement5 = Nothing
    
End Sub

Private Sub CreateCancellationResponsePLDAItem(ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode, _
                        ByRef objChildElement As IXMLDOMElement)
                        
    Set objChildNode = objDOM.createElement("ResponsePLDAItem")
    objDOM.documentElement.appendChild objChildNode
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("dataElementPath"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("dataElementSequence"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("errorCode"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("errorContent"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("errorDescription"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub

Private Sub CreateCancellationResponseCalculationResults(ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode, _
                        ByRef objChildElement As IXMLDOMElement, _
                        ByRef objChildElement2 As IXMLDOMElement)

    Set objChildNode = objDOM.createElement("CalculationResults")
    objDOM.documentElement.appendChild objChildNode
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("sequence"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Duties"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taxType"))
            objChildElement2.Text = ""
        
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taxBase"))
            objChildElement2.Text = ""
        
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taxAmount"))
            objChildElement2.Text = ""
        
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("paymentMethodTaxes"))
            objChildElement2.Text = ""
                
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine)

End Sub


