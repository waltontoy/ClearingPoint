Attribute VB_Name = "MExport"
Option Explicit

'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'            <ChildElement>
'                <ChildElement2>
'                    <ChildElement3>

Public Sub CreateExportXML(ByVal XMLMessageType As m_enumXMLMessage, _
                           ByRef objDOM As DOMDocument, _
                           ByRef objParentNode As IXMLDOMNode, _
                           ByRef objChildNode As IXMLDOMNode)

    Dim objChildElement As IXMLDOMElement
    Dim objChildElement2 As IXMLDOMElement
    
    Call CreateExportInterchangeHeader(XMLMessageType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    Call CreateExportGoodsDeclaration(objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    Call CreateExportProcedureType(objDOM, objParentNode, objChildNode)
    Call CreateExportGoodsItem(objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    
    Set objChildElement2 = Nothing
    Set objChildElement = Nothing
    
End Sub

Private Sub CreateExportInterchangeHeader(ByVal XMLMessageType As m_enumXMLMessage, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement)

    '***********************************************************************************************
    'Interchange Header
    '***********************************************************************************************
    '   - Procedure to generate InterchangeHeader Node for PLDA Lux Export
    '***********************************************************************************************
    'Date Modified: August 29, 2007
    'Modifications: 1. Added In-line comment.
    '***********************************************************************************************
    Set objChildNode = objDOM.createElement("InterchangeHeader")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'PLDA Sender
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageSender"))
        objChildElement.Text = G_strMessageSender   ' Client AS2 Identifier
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'PLDA Recipient
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageRecipient"))
        objChildElement.Text = G_strMessageRecipient    ' PLDA AS2 Identifier
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Message Version
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageVersion"))
        objChildElement.Text = "V0.1"
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Test Indicator
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("testIndicator"))
        objChildElement.Text = G_lngTestOnly
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Function Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("functionCode"))
        objChildElement.Text = G_strIEFunctionCode      ' Export Declaration Data
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DateTimeOfPreparation"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Preparation Date
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("dateOfPreparation"))
            objChildElement2.Text = Format(Now, "YYYYMMDD")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Preparation Time
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("timeOfPreparation"))
            objChildElement2.Text = Format(Now, "HHMMSS")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    '***********************************************************************************************
    
End Sub

Private Sub CreateExportGoodsDeclaration(ByRef objDOM As DOMDocument, _
                                         ByRef objParentNode As IXMLDOMNode, _
                                         ByRef objChildNode As IXMLDOMNode, _
                                         ByRef objChildElement As IXMLDOMElement, _
                                         ByRef objChildElement2 As IXMLDOMElement)
    
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement
        
    Dim lngCountSeal As Long
    Dim lngSealTotal As Long
    
    Dim strAuthCodes As String
    Dim lngDocumenten As Long
    Dim lngDetailCtr As Long
    
    '***********************************************************************************************
    'Goods Declaration
    '***********************************************************************************************
    '   - Procedure to generate GoodsDeclaration Node for PLDA Lux Export
    '***********************************************************************************************
    'Date Modified: August 29, 2007
    'Modifications: 1. Nodes are not created when value is empty string or null (Optional nodes).
    '               2. Nodes are not created when Not Used based on G_strIEFunctionCode.
    '               3. Formatting of some nodes were fixed.
    '               4. Seal Affixed are not counted when value is empty or null.
    '***********************************************************************************************
    Set objChildNode = objDOM.createElement("GoodsDeclaration")
    objDOM.documentElement.appendChild objChildNode
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Loading List
        If (GenerateData(G_rstHeader, "A8") <> "" And GenerateData(G_rstHeader, "A8") <> 0) Or _
           G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
           
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("loadingList"))
            objChildElement.Text = GenerateData(G_rstHeader, "A8")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Local Reference Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("localReferenceNumber"))
        objChildElement.Text = GenerateData(G_rstHeader, "A3")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Commercial Reference
        If GenerateData(G_rstHeader, "AC") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("commercialReference"))
            objChildElement.Text = GenerateData(G_rstHeader, "AC")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Totals +
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Totals"))
        
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Totals Items
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("items"))
            objChildElement2.Text = G_rstDetails.RecordCount
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Totals TotalGrossMass
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalGrossmass"))
            objChildElement2.Text = GenerateData(G_rstHeader, "D1")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Totals Packages
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
            objChildElement2.Text = GenerateData(G_rstHeader, "D3")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Transaction Nature +
        If GenerateData(G_rstHeader, "C7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TransactionNature"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Transaction Nature 1
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature1"))
                objChildElement2.Text = Left$(GenerateData(G_rstHeader, "C7"), 1) ' First character of C7
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Transaction Nature 2
                If Len(Trim$(GenerateData(G_rstHeader, "C7"))) = 2 Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature2"))
                    If Len(Trim$(GenerateData(G_rstHeader, "C7"))) = 2 Then
                        objChildElement2.Text = Trim$(Right$(GenerateData(G_rstHeader, "C7"), 1)) ' Second character of C7
                    Else
                        objChildElement2.Text = vbNullString
                    End If
                                
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Registration Number
        If G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("registrationNumber"))
            objChildElement.Text = GenerateData(G_rstHeader, "AB")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Issue Place
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("issuePlace"))
        objChildElement.Text = GenerateData(G_rstHeader, "A5")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Signature +
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("signature"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Signature Operator Contact Name
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorContactName"))
            objChildElement2.Text = GenerateData(G_rstHeader, "AA")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Signature Capacity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("capacity"))
            objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XG", 2)
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Type Part One
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartOne"))
        objChildElement.Text = GenerateData(G_rstHeader, "A1")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Type Part Two
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartTwo"))
        objChildElement.Text = GenerateData(G_rstHeader, "A2")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Type Part Three
        If G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartThree"))
            ' NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
            Debug.Assert False
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Consignor +
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Consignor"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Consignor Operator Identity +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
                'Consignor Operator Identity
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1", 5)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Consignor Operator +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Consignor Operator Name
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 5)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                'Consignor Operator Address +
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Consignor Operator Address Postal Code
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 5)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Consignor Operator Address Street and Number 1
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 5)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Consignor Operator Address Street and Number 2
                    If GenerateData(G_rstHeaderHandelaars, "X4", 5) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                        objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 5)
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'Consignor Operator Address City
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 5)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Consignor Operator Address Country - Edwin Nov28
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 5)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Consignee +
        If GenerateData(G_rstHeaderHandelaars, "X5", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X3", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X4", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X6", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X8", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X2", 4) <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X1", 4) <> "" Or _
           G_BlnDoNotIncludeIfEmpty = False Then
               
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("Consignee"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Consignee Operator Identity +
                If GenerateData(G_rstHeaderHandelaars, "X1", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                        'Consignee Operator Identity
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
                        objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1", 4)
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Consignee Operator +
                If GenerateData(G_rstHeaderHandelaars, "X5", 4) <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X3", 4) <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X4", 4) <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X6", 4) <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X8", 4) <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X2", 4) <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                   
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                        'Consignee Operator Name
                        If GenerateData(G_rstHeaderHandelaars, "X2", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 4)
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                                        
                        'Consignee Operator Address +
                        If GenerateData(G_rstHeaderHandelaars, "X5", 4) <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X3", 4) <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X4", 4) <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X6", 4) <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X8", 4) <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then
                           
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                            objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
                                'Consignee Operator Address Postal Code
                                If GenerateData(G_rstHeaderHandelaars, "X5", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 4)
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Consignee Operator Address Street and Number 1
                                If GenerateData(G_rstHeaderHandelaars, "X3", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 4)
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Consignee Operator Address Street and Number 2
                                If GenerateData(G_rstHeaderHandelaars, "X4", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 4)
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Consignee Operator Address City
                                If GenerateData(G_rstHeaderHandelaars, "X6", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 4)
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Consignee Operator Address Country
                                If GenerateData(G_rstHeaderHandelaars, "X8", 4) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 4)
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
            
        'Transport Means +
        If GenerateData(G_rstHeader, "C2") <> "" Or _
           GenerateData(G_rstHeader, "C3") <> "" Or _
           GenerateData(G_rstHeader, "D8") <> "" Or _
           GenerateData(G_rstHeader, "D7") <> "" Or _
           GenerateData(G_rstHeader, "D6") <> "" Or _
           GenerateData(G_rstHeader, "D5") <> "" Or _
           GenerateData(G_rstHeader, "D9") <> "" Or _
           G_rstDetailsContainer.RecordCount > 0 Or _
           G_BlnDoNotIncludeIfEmpty = False Then
            
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TransportMeans"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                'Transport Means Delivery Terms +
                If GenerateData(G_rstHeader, "C2") <> "" Or _
                   GenerateData(G_rstHeader, "C3") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DeliveryTerms"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                        'Transport Means Delivery Terms
                        If GenerateData(G_rstHeader, "C2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("deliveryTerms"))
                            objChildElement3.Text = GenerateData(G_rstHeader, "C2")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        
                        'Transport Means Delivery Terms Place
                        If GenerateData(G_rstHeader, "C3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("deliveryTermsPlace"))
                            objChildElement3.Text = GenerateData(G_rstHeader, "C3")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Border Mode
                If GenerateData(G_rstHeader, "D8") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("borderMode"))
                    objChildElement2.Text = GenerateData(G_rstHeader, "D8")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Border Nationality
                If GenerateData(G_rstHeader, "D7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("borderNationality"))
                    objChildElement2.Text = GenerateData(G_rstHeader, "D7")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Border Identity
                If GenerateData(G_rstHeader, "D6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("borderIdentity"))
                    objChildElement2.Text = GenerateData(G_rstHeader, "D6")
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Departure Identity
                If GenerateData(G_rstHeader, "D5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("departureIdentity"))
                    objChildElement2.Text = GenerateData(G_rstHeader, "D5")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Inland Mode
                If GenerateData(G_rstHeader, "D9") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("inlandMode"))
                    objChildElement2.Text = GenerateData(G_rstHeader, "D9")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Transport Means Containerized Indicator
                If G_rstDetailsContainer.RecordCount > 0 Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("container"))
                    
                    'CSCLP-354
                    Dim blnContainersHaveValue As Boolean
                    Dim lngCurrentDetail As Long
                    Dim blnGoNextDetail As Boolean
                    
                    G_rstDetailsContainer.Sort = "DETAIL ASC, ORDINAL ASC"
                    G_rstDetailsContainer.MoveFirst
                    
                    'Memo current detail tab
                    lngCurrentDetail = FNullField(G_rstDetailsContainer.fields("Detail").Value)
                    
                    Do Until G_rstDetailsContainer.EOF
                        If blnGoNextDetail = False Then
                            If (Len((FNullField(G_rstDetailsContainer.fields("S4").Value))) > 0 _
                                And FNullField(G_rstDetailsContainer.fields("S4").Value <> 0)) _
                                Or _
                               (Len((FNullField(G_rstDetailsContainer.fields("S5").Value))) > 0 _
                                And FNullField(G_rstDetailsContainer.fields("S5").Value <> 0)) Then
                                'Container value found, give up
                                blnContainersHaveValue = True
                                Exit Do
                            Else
                                If UCase((FNullField(G_rstDetailsContainer.fields("S6").Value))) = "E" Then
                                    'Enable detail search
                                    blnGoNextDetail = True
                                Else
                                    'Move to next ordinal
                                    G_rstDetailsContainer.MoveNext
                                End If
                            End If
                        Else
                            'Keep searching next detail until last record
                            Do Until G_rstDetailsContainer.EOF
                                If lngCurrentDetail = FNullField(G_rstDetailsContainer.fields("Detail").Value) Then
                                    G_rstDetailsContainer.MoveNext
                                Else
                                    'Memo new detail tab
                                    lngCurrentDetail = FNullField(G_rstDetailsContainer.fields("Detail").Value)
                                    'Disable detail search
                                    blnGoNextDetail = False
                                    Exit Do
                                End If
                            Loop
                        End If
                    Loop
                    objChildElement2.Text = -(blnContainersHaveValue)
                    G_rstDetailsContainer.Sort = vbNullString
                    
                    'Original
                    'objChildElement2.Text = IIf(G_rstDetailsContainer.RecordCount > 0, "1", "")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                                    
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
            
        'Customs +
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Customs"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Customs Exit Office
            If GenerateData(G_rstHeader, "A7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("exitOffice"))
                objChildElement2.Text = GenerateData(G_rstHeader, "A7")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Customs Goods Location +
            If GenerateData(G_rstHeader, "D4") <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("GoodsLocation"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                    'Customs Goods Location Precise
                    If GenerateData(G_rstHeader, "D4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("precise"))
                        objChildElement3.Text = GenerateData(G_rstHeader, "D4")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'Customs Goods Location Postal Code ( NOT OPTIONAL according to the XSD file )
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("postalcode"))
                    objChildElement3.Text = GenerateData(G_rstHeader, "DG") 'Edwin Oct 22
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Customs Export Office
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("exportCustomsOffice"))
            objChildElement2.Text = GenerateData(G_rstHeader, "A6")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
            'Seal +
            If G_rstHeaderZegels.RecordCount > 0 Then
                G_rstHeaderZegels.Sort = "ORDINAL ASC"
                G_rstHeaderZegels.MoveFirst
                
                Do While Not G_rstHeaderZegels.EOF
                    lngCountSeal = lngCountSeal + 1
                    
                    If GenerateData(G_rstHeaderZegels, "E1", lngCountSeal) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Seal"))
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                            
                            'Seal SealAffixed
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("sealAffixed"))
                            objChildElement3.Text = GenerateData(G_rstHeaderZegels, "E1", lngCountSeal)
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Count the non-empty Seals
                            lngSealTotal = lngSealTotal + 1
                            
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    If Trim$(FNullField(G_rstHeaderZegels.fields("E3").Value)) = "E" Then Exit Do
                    
                    G_rstHeaderZegels.MoveNext
                Loop
            End If
                                    
            'Total Number of Seals
            If lngSealTotal > 0 And G_BlnDoNotIncludeIfEmpty = True Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("seals"))
                objChildElement2.Text = lngSealTotal
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Representative +
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Representative"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Representative Operator Identity +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
                'Representative Operator Identity
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1", 2)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Representative Operator +
            If GenerateData(G_rstHeaderHandelaars, "X2", 2) <> "" Or _
               GenerateData(G_rstHeaderHandelaars, "X5", 2) <> "" Or _
               GenerateData(G_rstHeaderHandelaars, "X3", 2) <> "" Or _
               GenerateData(G_rstHeaderHandelaars, "X4", 2) <> "" Or _
               GenerateData(G_rstHeaderHandelaars, "X6", 2) <> "" Or _
               GenerateData(G_rstHeaderHandelaars, "X8", 2) <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
            
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                    'Representative Operator Name
                    If GenerateData(G_rstHeaderHandelaars, "X2", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                        objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 2)
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'Representative Operator Address +
                    If GenerateData(G_rstHeaderHandelaars, "X5", 2) <> "" Or _
                       GenerateData(G_rstHeaderHandelaars, "X3", 2) <> "" Or _
                       GenerateData(G_rstHeaderHandelaars, "X4", 2) <> "" Or _
                       GenerateData(G_rstHeaderHandelaars, "X6", 2) <> "" Or _
                       GenerateData(G_rstHeaderHandelaars, "X8", 2) <> "" Or _
                       G_BlnDoNotIncludeIfEmpty = False Then
                        
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                            'Representative Operator Address Postal Code
                            If GenerateData(G_rstHeaderHandelaars, "X5", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 2)
                                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Representative Operator Address Street and Number 1
                            If GenerateData(G_rstHeaderHandelaars, "X3", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 2)
                                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Representative Operator Address Street and Number 2
                            If GenerateData(G_rstHeaderHandelaars, "X4", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 2)
                                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Representative Operator Address City
                            If GenerateData(G_rstHeaderHandelaars, "X6", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 2)
                                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Representative Operator Address Country
                            If GenerateData(G_rstHeaderHandelaars, "X8", 2) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 2)
                                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
            
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Representative Capacity - Edwin Nov28
            If GenerateData(G_rstHeaderHandelaars, "XG", 2) <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("capacity"))
                objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XG", 2)
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Representative Declarant Status
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("declarantStatus"))
            objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XF", 2)
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Representative Contact Person Name
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("contactPersonName"))
            objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "X9", 2)
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Representative Authorised Identity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisedIdentity"))
            objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XH", 2)
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Principal +
        If GenerateData(G_rstHeaderHandelaars, "X2") <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X5") <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X3") <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X4") <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X6") <> "" Or _
           GenerateData(G_rstHeaderHandelaars, "X8") <> "" Or _
           G_BlnDoNotIncludeIfEmpty = False Then
               
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("Principal"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Principal Operator Identity +
                If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then ' XXXXX Not in Used for CONFIRMATION XXXXX
                    ' Start - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
                    Debug.Assert False
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
                        objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    ' End - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
                End If
                
                'Principal Operator +
                If GenerateData(G_rstHeaderHandelaars, "X2") <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X5") <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X3") <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X4") <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X6") <> "" Or _
                   GenerateData(G_rstHeaderHandelaars, "X8") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                   
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                        'Principal Operator Name
                        If GenerateData(G_rstHeaderHandelaars, "X2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        
                        'Principal Operator Address +
                        If GenerateData(G_rstHeaderHandelaars, "X5") <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X3") <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X4") <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X6") <> "" Or _
                           GenerateData(G_rstHeaderHandelaars, "X8") <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then
                           
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                            objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
                                'Principal Operator Address Postal Code
                                If GenerateData(G_rstHeaderHandelaars, "X5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5")
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Principal Operator Address  Street and Number 1
                                If GenerateData(G_rstHeaderHandelaars, "X3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3")
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Principal Operator Address Street and Number 2
                                If GenerateData(G_rstHeaderHandelaars, "X4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4")
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Principal Operator Address City
                                If GenerateData(G_rstHeaderHandelaars, "X6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6")
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Principal Operator Address Country
                                If GenerateData(G_rstHeaderHandelaars, "X8") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8")
                                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                End If
                                                
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                    
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Invoice +
        If GenerateData(G_rstHeader, "C6") <> "" Or _
           GenerateData(G_rstHeader, "C5") <> "" Or _
           GenerateData(G_rstHeader, "C4") <> "" Or _
           G_BlnDoNotIncludeIfEmpty = False Then
               
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("Invoice"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Exchange Rate +
                If GenerateData(G_rstHeader, "C6") <> "" Or _
                   GenerateData(G_rstHeader, "C5") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                   
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ExchangeRate"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                        'Exchange Rate
                        If GenerateData(G_rstHeader, "C6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("exchangeRate"))
                            objChildElement3.Text = GenerateData(G_rstHeader, "C6")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        
                        'Exchange Rate Currency
                        If GenerateData(G_rstHeader, "C5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("currency"))
                            objChildElement3.Text = GenerateData(G_rstHeader, "C5")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        
                        'Exchange Rate Date
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("exchangeDate"))
                        objChildElement3.Text = GenerateData(G_rstHeader, "A4")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Invoice Amount
                If GenerateData(G_rstHeader, "C4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("invoiceAmount"))
                    'objChildElement2.Text = Format(GenerateData(G_rstHeader, "C4"), "################.00") ' Format 16.2
                    objChildElement2.Text = GenerateData(G_rstHeader, "C4")
                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Deferred Payment
        If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
            ' Start - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
            Debug.Assert False
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("deferredPayment"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Destination Country
        If GenerateData(G_rstDetails, "N7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("destinationCountry"))
            objChildElement.Text = GenerateData(G_rstDetails, "N7")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Dispatch Country
        If GenerateData(G_rstHeader, "DB") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("dispatchCountry"))
            objChildElement.Text = GenerateData(G_rstHeader, "DB")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Customs Main Reference Number
        If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
            ' Start - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
            Debug.Assert False
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("customsMainReferenceNumber"))
            objChildElement.Text = GenerateData(G_rstHeader, "MRN")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Declaration Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationDate"))
        objChildElement.Text = GenerateData(G_rstHeader, "A4")
        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Declaration Period From Date
        If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationPeriodFromDate"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Declaration Period To Date
        If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationPeriodToDate"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                        
        'Incomplete Declaration Reason
        If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
            ' Must inquire frank and luc where to get this info from the codisheet and if a Box needs to be added
            Debug.Assert False
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("incompleteDeclarationReason"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                                   
        '**********************************************************************************************
        'Authorisation +
        '**********************************************************************************************
        If (GenerateData(G_rstLogIDFields, "Authorisation Code") <> "" Or _
           GenerateData(G_rstLogIDFields, "Authorisation Reference") <> "") Or G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("Authorisation"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                If GenerateData(G_rstLogIDFields, "Authorisation Code") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationCode"))
                    objChildElement2.Text = GenerateData(G_rstLogIDFields, "Authorisation Code")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                                
                'Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationDate"))
                'objChildElement2.Text = ""
                'objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                If GenerateData(G_rstLogIDFields, "Authorisation Reference") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationReference"))
                    objChildElement2.Text = GenerateData(G_rstLogIDFields, "Authorisation Reference")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("controllingOffice"))
                'objChildElement2.Text = ""
                'objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        '**********************************************************************************************
        
        '**********************************************************************************************
        'CSCLP-232: Check for Authorisation Documents from Tariff Lines
        '**********************************************************************************************
        If G_rstDetails.RecordCount > 0 Then
            
            For lngDetailCtr = 1 To G_rstDetails.RecordCount
            
                G_rstDetailsDocumenten.Filter = adFilterNone
                G_rstDetailsDocumenten.Filter = "Detail = " & lngDetailCtr
                
                G_rstDetailsDocumenten.Sort = "Ordinal ASC"
                
                'G_rstDetailsDocumenten.MoveFirst
                
                If G_rstDetailsDocumenten.RecordCount > 0 Then
                    G_rstDetailsDocumenten.MoveFirst
                    
                    lngDocumenten = 0
                    
                    Do While Not G_rstDetailsDocumenten.EOF
                        lngDocumenten = lngDocumenten + 1
                        
                        
                        If GenerateData(G_rstDetailsDocumenten, "Q5", lngDocumenten) = "99" And _
                           GenerateData(G_rstDetailsDocumenten, "Q1", lngDocumenten) <> "" And _
                           GenerateData(G_rstDetailsDocumenten, "Q2", lngDocumenten) <> "" Then
                                
                                If InStr(1, strAuthCodes, GenerateData(G_rstDetailsDocumenten, "Q1", lngDocumenten)) = 0 Then
                                    
                                    'Save on strAuthCodes all the Distinct Authorisation Codes
                                    strAuthCodes = strAuthCodes & "@@@@@" & GenerateData(G_rstDetailsDocumenten, "Q1", lngDocumenten)
                                    
                                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Authorisation"))
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                                    
                                        'Authorisation Code
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationCode"))
                                        objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q1", lngDocumenten)
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                                                                            
                                        'Authorisation Date
                                        If GenerateData(G_rstDetailsDocumenten, "Q3", lngDocumenten) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationDate"))
                                            objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q3", lngDocumenten)
                                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                        End If
                                        
                                        'Authorisation Reference
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisationReference"))
                                        objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q2", lngDocumenten)
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                        
                                        'Controlling Office
                                        If GenerateData(G_rstDetailsDocumenten, "Q4", lngDocumenten) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("controllingOffice"))
                                            objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q4", lngDocumenten)
                                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                        End If
                                        
                                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                End If
                        End If
                        
                        If Trim$(FNullField(G_rstDetailsDocumenten.fields("QA").Value)) = "E" Then Exit Do
                        
                        G_rstDetailsDocumenten.MoveNext
                    Loop
                End If
                
            Next
        
        End If
        '**********************************************************************************************
                                            
        If G_strIEFunctionCode = "IE" & m_enumXMLMessage.enumExportAmendment Then
            'Amendment Date
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentDate"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                
            'Amendment Place
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentPlace"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                                    
        'Acceptance Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("acceptanceDate"))
        objChildElement.Text = GenerateData(G_rstHeader, "A4")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        
        If G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumExportDeclaration Then
            'Release Date
            ' Start - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
            Debug.Assert False
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("releaseDate"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                        
            'Release Time
            ' Start - NU (Not Used) in PLDA Lux Specifications; Must find out if this can be here with an empty string value
            Debug.Assert False
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("releaseTime"))
            objChildElement.Text = ""
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        End If
    
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    '***********************************************************************************************
    
    Set objChildElement3 = Nothing
    Set objChildElement4 = Nothing
    
End Sub

Private Sub CreateExportProcedureType(ByRef objDOM As DOMDocument, _
                                      ByRef objParentNode As IXMLDOMNode, _
                                      ByRef objChildNode As IXMLDOMNode)
    
    '***********************************************************************************************
    'Procedure Type
    '***********************************************************************************************
    '   - Procedure to generate ProcedureType Node for PLDA Lux Export
    '***********************************************************************************************
    'Date Modified: August 29, 2007
    'Modifications: 1. Added In-line comment.
    '***********************************************************************************************
    Set objChildNode = objDOM.createElement("ProcedureType")
    objChildNode.Text = GenerateData(G_rstDetails, "N4")
    objDOM.documentElement.appendChild objChildNode
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    '***********************************************************************************************

End Sub

Private Sub CreateExportGoodsItem(ByRef objDOM As DOMDocument, _
                                  ByRef objParentNode As IXMLDOMNode, _
                                  ByRef objChildNode As IXMLDOMNode, _
                                  ByRef objChildElement As IXMLDOMElement, _
                                  ByRef objChildElement2 As IXMLDOMElement)
 
    Dim objChildElement3 As IXMLDOMElement
    
    Dim lngDetailCounter As Long
    Dim strDummy As String
    
    Dim lngBijzondere As Long
    Dim lngContainer As Long
    Dim lngDocumenten As Long
    
    Dim strCheckMass As String
        
    '***********************************************************************************************
    'Goods Item
    '***********************************************************************************************
    '   - Procedure to generate GoodsItem Nodes for PLDA Lux Export
    '***********************************************************************************************
    'Date Modified: August 29, 2007
    'Modifications: 1. Nodes are not created when value is empty string or null (Optional nodes).
    '               2. Nodes are not created when Not Used based on G_strIEFunctionCode.
    '               3. Formatting of some nodes were fixed.
    '***********************************************************************************************
    If G_rstDetails.RecordCount > 0 Then
        G_rstDetails.Sort = "DETAIL ASC"
        G_rstDetails.MoveFirst
    
        For lngDetailCounter = 1 To G_rstDetails.RecordCount
            
            If G_rstDetailsBijzondere.RecordCount > 0 Then G_rstDetailsBijzondere.MoveFirst
            If G_rstDetailsContainer.RecordCount > 0 Then G_rstDetailsContainer.MoveFirst
            If G_rstDetailsDocumenten.RecordCount > 0 Then G_rstDetailsDocumenten.MoveFirst
            
            lngBijzondere = 0
            lngContainer = 0
            lngDocumenten = 0
            
            Set objChildNode = objDOM.createElement("GoodsItem")
            objDOM.documentElement.appendChild objChildNode
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
                'Item Number
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("ItemNumber"))
                objChildElement.Text = G_rstDetails.fields("Detail").Value '""
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                '***********************************************************************************************
                'Commodity Code +
                '***********************************************************************************************
                If GenerateData(G_rstDetails, "L1") <> "" Or _
                   GenerateData(G_rstDetails, "L2") <> "" Or _
                   GenerateData(G_rstDetails, "L3") <> "" Or _
                   GenerateData(G_rstDetails, "L4") <> "" Or _
                   GenerateData(G_rstDetails, "L5") <> "" Or _
                   GenerateData(G_rstDetails, "L6") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("commodityCode"))
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                    
                        'Commodity Code
                        If GenerateData(G_rstDetails, "L1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("commodityCode"))
                            objChildElement2.Text = Left$(GenerateData(G_rstDetails, "L1"), 8)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code Taric
                        If GenerateData(G_rstDetails, "L1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taric"))
                            objChildElement2.Text = Right$(GenerateData(G_rstDetails, "L1"), 2)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code firstAdditionalCommodity
                        If GenerateData(G_rstDetails, "L2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("firstAdditionalCommodity"))
                            objChildElement2.Text = IIf(GenerateData(G_rstDetails, "L2") = "", "0000", GenerateData(G_rstDetails, "L2")) ' Zero filled when empty
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code secondAdditionalCommodity
                        If GenerateData(G_rstDetails, "L3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("secondAdditionalCommodity"))
                            objChildElement2.Text = IIf(GenerateData(G_rstDetails, "L3") = "", "0000", GenerateData(G_rstDetails, "L3")) ' Zero filled when empty
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code nationalAdditionalCommodity1
                        If GenerateData(G_rstDetails, "L4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity1"))
                            objChildElement2.Text = IIf(GenerateData(G_rstDetails, "L4") = "", "0000", GenerateData(G_rstDetails, "L4")) ' Zero filled when empty
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code nationalAdditionalCommodity2
                        If GenerateData(G_rstDetails, "L5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity2"))
                            objChildElement2.Text = GenerateData(G_rstDetails, "L5")
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'Commodity Code nationalAdditionalCommodity3
                        If GenerateData(G_rstDetails, "L6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity3"))
                            objChildElement2.Text = GenerateData(G_rstDetails, "L6")
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                '***********************************************************************************************
                
                'Net Mass
                If GenerateData(G_rstDetails, "LA") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("netMass"))
                    strCheckMass = GenerateData(G_rstDetails, "LA")
                    If Val(strCheckMass) = 0 Then
                        strCheckMass = ""
                    End If
                    objChildElement.Text = strCheckMass

                    'FINALIZE
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                'Gross Mass
                If (GenerateData(G_rstDetails, "L9") <> "" And GenerateData(G_rstDetails, "L9") <> "0") Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("grossMass"))
                    
                    'CSCLP-209
'                    strCheckMass = GenerateData(G_rstDetails, "L9")
'                    If Val(strCheckMass) = 0 Then
'                        strCheckMass = ""
'                    End If
                    
                    'FINALIZE
                    objChildElement.Text = GenerateData(G_rstDetails, "L9") 'strCheckMass
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                'Goods Description
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("goodsDescription"))
                objChildElement.Text = GenerateData(G_rstDetails, "L8")
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                '***********************************************************************************************
                'Packaging +
                '***********************************************************************************************
                If GenerateData(G_rstDetails, "S3") <> "" Or _
                   GenerateData(G_rstDetails, "S2") <> "" Or _
                   GenerateData(G_rstDetails, "S1") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                    
                    '***********************************************************************************************
                    'Marks and Numbers is not filled up while Packages and Package Type are filled up
                    '***********************************************************************************************
                    If Len(Trim$(GenerateData(G_rstDetails, "S3"))) = 0 Then
                        If GenerateData(G_rstDetails, "S2") <> "" Or _
                           GenerateData(G_rstDetails, "S1") <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then
                             
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("Packaging"))
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                
                                'Packages
                                If GenerateData(G_rstDetails, "S2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
                                    objChildElement2.Text = GenerateData(G_rstDetails, "S2")
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                                'Package Type
                                If GenerateData(G_rstDetails, "S1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packageType"))
                                    objChildElement2.Text = GenerateData(G_rstDetails, "S1")
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                                'Type Marks
                                If G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
                                    ' Must ask frank and luc where to get this value
                                    ' enumeration MAWB
                                    ' enumeration HAWB
                                    ' enumeration TABAC
                                    ' enumeration RESTI
                                    Debug.Assert False
                                    If Len(Trim(GenerateData(G_rstDetails, "SF"))) > 0 Then
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TypeMarks"))
                                        objChildElement2.Text = GenerateData(G_rstDetails, "SF")
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                    End If
                                End If
                                
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                    End If
                    '***********************************************************************************************
                    
                    '***********************************************************************************************
                    'First Part
                    '***********************************************************************************************
                    If Len(Trim$(GenerateData(G_rstDetails, "S3"))) > 0 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Packaging"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                        
                            'Marks and Number
                            If GenerateData(G_rstDetails, "S3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("marksNumber"))
                                objChildElement2.Text = Mid(GenerateData(G_rstDetails, "S3"), 1, 35)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Packages
                            If GenerateData(G_rstDetails, "S2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S2")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Package Type
                            If GenerateData(G_rstDetails, "S1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packageType"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S1")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Type Marks
                                If G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
                                    ' Must ask frank and luc where to get this value
                                    ' enumeration MAWB
                                    ' enumeration HAWB
                                    ' enumeration TABAC
                                    ' enumeration RESTI
                                    Debug.Assert False
                                    If Len(Trim(GenerateData(G_rstDetails, "SF"))) > 0 Then
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TypeMarks"))
                                        objChildElement2.Text = GenerateData(G_rstDetails, "SF")
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                    End If
                                End If
                        
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    '***********************************************************************************************
                    
                    '***********************************************************************************************
                    'Second Part
                    '***********************************************************************************************
                    If Len(Trim$(GenerateData(G_rstDetails, "S3"))) > 35 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Packaging"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                        
                            'Marks and Number
                            If GenerateData(G_rstDetails, "S3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("marksNumber"))
                                objChildElement2.Text = Mid(GenerateData(G_rstDetails, "S3"), 36, 35)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Packages
                            If GenerateData(G_rstDetails, "S2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S2")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Package Type
                            If GenerateData(G_rstDetails, "S1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packageType"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S1")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Type Marks
                                If G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
                                    ' Must ask frank and luc where to get this value
                                    ' enumeration MAWB
                                    ' enumeration HAWB
                                    ' enumeration TABAC
                                    ' enumeration RESTI
                                    Debug.Assert False
                                    If Len(Trim(GenerateData(G_rstDetails, "SF"))) > 0 Then
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TypeMarks"))
                                        objChildElement2.Text = GenerateData(G_rstDetails, "SF")
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                    End If
                                End If
                        
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    '***********************************************************************************************
                    
                    '***********************************************************************************************
                    'Third Part
                    '***********************************************************************************************
                    If Len(Trim$(GenerateData(G_rstDetails, "S3"))) > 70 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Packaging"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                        
                            'Marks and Number
                            If GenerateData(G_rstDetails, "S3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("marksNumber"))
                                objChildElement2.Text = Mid(GenerateData(G_rstDetails, "S3"), 71, 35)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Packages
                            If GenerateData(G_rstDetails, "S2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S2")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Package Type
                            If GenerateData(G_rstDetails, "S1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packageType"))
                                objChildElement2.Text = GenerateData(G_rstDetails, "S1")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Type Marks
                                If G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
                                    ' Must ask frank and luc where to get this value
                                    ' enumeration MAWB
                                    ' enumeration HAWB
                                    ' enumeration TABAC
                                    ' enumeration RESTI
                                    Debug.Assert False
                                    If Len(Trim(GenerateData(G_rstDetails, "SF"))) > 0 Then
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TypeMarks"))
                                        objChildElement2.Text = GenerateData(G_rstDetails, "SF")
                                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                    End If
                                End If
                        
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    '***********************************************************************************************
                    
                End If
                '***********************************************************************************************
                
                '***********************************************************************************************
                'Container Identifier +
                '***********************************************************************************************
                G_rstDetailsContainer.Filter = adFilterNone
                G_rstDetailsContainer.Filter = "Detail = " & lngDetailCounter
                
                If G_rstDetailsContainer.RecordCount > 0 Then
                    G_rstDetailsContainer.Sort = "ORDINAL ASC"
                    G_rstDetailsContainer.MoveFirst
                    
                    Do Until G_rstDetailsContainer.EOF
                        lngContainer = lngContainer + 1
                        
                        strDummy = Trim$(GenerateData(G_rstDetailsContainer, "S4", lngContainer))
                        If LenB(strDummy) > 0 Then
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("containerIdentifier"))
                            objChildElement.Text = strDummy
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                        
                        strDummy = Trim$(GenerateData(G_rstDetailsContainer, "S5", lngContainer))
                        If LenB(strDummy) > 0 Then
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("containerIdentifier"))
                            objChildElement.Text = strDummy
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                        
                        If Trim$(FNullField(G_rstDetailsContainer.fields("S6").Value)) = "E" Then Exit Do
                        
                        G_rstDetailsContainer.MoveNext
                    Loop
                End If
                '***********************************************************************************************
                
                '***********************************************************************************************
                'Produced Document +
                '***********************************************************************************************
                G_rstDetailsDocumenten.Filter = adFilterNone
                G_rstDetailsDocumenten.Filter = "Detail = " & lngDetailCounter
                
                If G_rstDetailsDocumenten.RecordCount > 0 Then
                    G_rstDetailsDocumenten.Sort = "ORDINAL ASC"
                    G_rstDetailsDocumenten.MoveFirst
                    
                    Do Until G_rstDetailsDocumenten.EOF
                        lngDocumenten = lngDocumenten + 1
                        
                        '****************************************************************************************
                        'CSCLP-232 IF Q5 = 99 then it is an authorisation document and must be skipped
                        '****************************************************************************************
                        If GenerateData(G_rstDetailsDocumenten, "Q5", lngDocumenten) <> "99" Then
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ProducedDocument"))
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                            
                                'Document +
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Document"))
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                                    'Document Reference
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentReference"))
                                    objChildElement3.Text = GenerateData(G_rstDetailsDocumenten, "Q2", lngDocumenten)
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                                    'Document Type
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentType"))
                                    objChildElement3.Text = GenerateData(G_rstDetailsDocumenten, "Q1", lngDocumenten)
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                                    'Document Date
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentDate"))
                                    objChildElement3.Text = GenerateData(G_rstDetailsDocumenten, "Q3", lngDocumenten)
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                    
                                If GenerateData(G_rstDetailsDocumenten, "Q9") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("complementaryInformation"))
                                    objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q9", lngDocumenten)
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                End If
                                    
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                        '****************************************************************************************
                        
                        If Trim$(FNullField(G_rstDetailsDocumenten.fields("QA").Value)) = "E" Then Exit Do
                        
                        G_rstDetailsDocumenten.MoveNext
                    Loop
                End If
                '***********************************************************************************************
                
                '***********************************************************************************************
                ' Supplementary Units +
                '***********************************************************************************************
                If GenerateData(G_rstDetails, "M2") <> "" Or _
                   GenerateData(G_rstDetails, "M1") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                   
                   Set objChildElement = objChildNode.appendChild(objDOM.createElement("SupplementaryUnits"))
                       objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                   
                       ' Supplementary Units
                       If GenerateData(G_rstDetails, "M2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                           Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("supplementaryUnits"))
                           objChildElement2.Text = GenerateData(G_rstDetails, "M2")
                           objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                       End If
    
                       ' Supplementary Units Code
                       If GenerateData(G_rstDetails, "M1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                           Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("supplementaryUnitsCode"))
                           objChildElement2.Text = GenerateData(G_rstDetails, "M1")
                           objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                       End If
                   
                   objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                '***********************************************************************************************
                
                'Origin Country
                If GenerateData(G_rstDetails, "NB") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("originCountry"))
                    objChildElement.Text = GenerateData(G_rstDetails, "NB")
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
            
                'Statistical Value
                If GenerateData(G_rstDetails, "O2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("statisticalValue"))
                    objChildElement.Text = GenerateData(G_rstDetails, "O2")
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                '***********************************************************************************************
                'Customs Treatment +
                '***********************************************************************************************
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("CustomsTreatment"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Warehouse +
                    If GenerateData(G_rstDetails, "M3") <> "" Or _
                       GenerateData(G_rstDetails, "M4") <> "" Or _
                       GenerateData(G_rstDetails, "M5") <> "" Or _
                       G_BlnDoNotIncludeIfEmpty = False Then
                    
                       Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Warehouse"))
                       objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                           
                           'Warehouse Type
                           If GenerateData(G_rstDetails, "M3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                               Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("warehouseType"))
                               objChildElement3.Text = GenerateData(G_rstDetails, "M3")
                               objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                           End If
                           
                           'Warehouse Identity
                           If GenerateData(G_rstDetails, "M4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                               Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("warehouseIdentity"))
                               objChildElement3.Text = GenerateData(G_rstDetails, "M4")
                               objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                           End If
                           
                           'Warehouse Country
                           If GenerateData(G_rstDetails, "M5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                               Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("warehouseCountry"))
                               objChildElement3.Text = GenerateData(G_rstDetails, "M5")
                               objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                           End If
    
                       objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    '***********************************************************************************************
 
                    '***********************************************************************************************
                    'TProcedure +
                    '***********************************************************************************************
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TProcedure"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                        'procedurePart1
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedurePart1"))
                        objChildElement3.Text = GenerateData(G_rstDetails, "N1")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                        'procedurePart2
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedurePart2"))
                        objChildElement3.Text = GenerateData(G_rstDetails, "N2")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                        'procedureNat
                        If GenerateData(G_rstDetails, "N3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedureNat"))
                            objChildElement3.Text = GenerateData(G_rstDetails, "N3")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'procedureNat
                        If GenerateData(G_rstDetails, "ND") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedureNat"))
                            objChildElement3.Text = GenerateData(G_rstDetails, "ND")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                        
                        'procedureNat
                        If GenerateData(G_rstDetails, "NE") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedureNat"))
                            objChildElement3.Text = GenerateData(G_rstDetails, "NE")
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                '***********************************************************************************************
 
                '***********************************************************************************************
                'Additional Information +
                '***********************************************************************************************
                G_rstDetailsBijzondere.Filter = adFilterNone
                G_rstDetailsBijzondere.Filter = "Detail = " & lngDetailCounter
                
                If G_rstDetailsBijzondere.RecordCount > 0 Then
                    G_rstDetailsBijzondere.Sort = "ORDINAL ASC"
                    G_rstDetailsBijzondere.MoveFirst
                    
                    Do Until G_rstDetailsBijzondere.EOF
                        lngBijzondere = lngBijzondere + 1
                        
                        If GenerateData(G_rstDetailsBijzondere, "P2", lngBijzondere) <> "" Or _
                           GenerateData(G_rstDetailsBijzondere, "P1", lngBijzondere) <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then
                        
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("AdditionalInformation"))
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                            
                                'Additional Information Content
                                If GenerateData(G_rstDetailsBijzondere, "P2", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContent"))
                                    objChildElement2.Text = GenerateData(G_rstDetailsBijzondere, "P2", lngBijzondere)
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                                'Additional Information Type
                                If GenerateData(G_rstDetailsBijzondere, "P1", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationType"))
                                    objChildElement2.Text = GenerateData(G_rstDetailsBijzondere, "P1", lngBijzondere)
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                                End If
                                                                
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

                        End If
                        
                        If Trim$(FNullField(G_rstDetailsBijzondere.fields("P5").Value)) = "E" Then Exit Do
                        
                        G_rstDetailsBijzondere.MoveNext
                    Loop
                
                End If
                '***********************************************************************************************
                
                '***********************************************************************************************
                'Previous Document +
                '***********************************************************************************************
                If GenerateData(G_rstDetails, "R5") <> "" Or _
                   GenerateData(G_rstDetails, "R2") <> "" Or _
                   GenerateData(G_rstDetails, "R1") <> "" Or _
                   G_BlnDoNotIncludeIfEmpty = False Then
                   
                    ' Must ask Frank if we need to redesign the Codisheet to accommodate more than
                    ' one (1) previous document
                    Debug.Assert False
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("PreviousDocument"))
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                    
                        'Document +
                        If GenerateData(G_rstDetails, "R5") <> "" Or _
                           GenerateData(G_rstDetails, "R2") <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then
                        
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Document"))
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                                'Previous Document Reference
                                If GenerateData(G_rstDetails, "R5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentReference"))
                                    
                                    Select Case GenerateData(G_rstDetails, "R2")
                                        Case "705"
                                            objChildElement3.Text = Left$(GenerateData(G_rstDetails, "R5"), 6)
                                        Case "740"
                                            objChildElement3.Text = Left$(GenerateData(G_rstDetails, "R5"), 5)
                                        Case Else
                                            objChildElement3.Text = GenerateData(G_rstDetails, "R5")
                                    End Select
                                    
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                End If
                            
                                'Previous Document Type
                                If GenerateData(G_rstDetails, "R2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentType"))
                                    objChildElement3.Text = GenerateData(G_rstDetails, "R2")
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                End If
                                
                                'Previous Document Date
                                If GenerateData(G_rstDetails, "R3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentDate"))
                                    objChildElement3.Text = GenerateData(G_rstDetails, "R3")
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        End If
                                            
                        'Previous Document Category
                        If GenerateData(G_rstDetails, "R1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("previousDocumentCategory"))
                            objChildElement2.Text = GenerateData(G_rstDetails, "R1")
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                            
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                '***********************************************************************************************
                
                'Function Code
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("functionCode"))
                objChildElement.Text = G_strIEFunctionCode
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            objParentNode.appendChild objDOM.createTextNode(vbNewLine)
        
            Set objChildElement3 = Nothing
        
            G_rstDetails.MoveNext
        
        Next
        
    End If
    '***********************************************************************************************
    
End Sub

