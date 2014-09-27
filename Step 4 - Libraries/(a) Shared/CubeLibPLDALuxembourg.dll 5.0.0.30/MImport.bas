Attribute VB_Name = "MImport"
Option Explicit

'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <ChildNode2>
'                <ChildElement>
'                   <ChildElement2>
'                       <ChildElement3>

Public Sub CreateImportXML(ByVal XMLMessageType As m_enumXMLMessage, _
                           ByRef objDOM As DOMDocument, _
                           ByRef objParentNode As IXMLDOMNode, _
                           ByRef objChildNode As IXMLDOMNode)
    
    Dim objChildNode2 As IXMLDOMNode
    Dim objChildElement As IXMLDOMElement
    Dim objChildElement2 As IXMLDOMElement
    
    Dim lngDetailCtr As Long
    Dim blnDV1Found As Boolean
    
    Call CreateImportInterchangeHeader(XMLMessageType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2)
    
    'PLDA Import
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("PLDAImport"))
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Goods Declaration
        Set objChildNode2 = objChildNode.appendChild(objDOM.createElement("GoodsDeclaration"))
        objDOM.documentElement.appendChild objChildNode
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
        Call CreateImportGoodsDeclaration(objDOM, objParentNode, objChildNode2, objChildElement, objChildElement2)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Procedure Type
        Set objChildNode2 = objChildNode.appendChild(objDOM.createElement("ProcedureType"))
        objChildNode2.Text = GenerateData(G_rstDetails, "N4")
        objDOM.documentElement.appendChild objChildNode
        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Goods Item
        Call CreateImportGoodsItem(objDOM, objChildNode, objChildNode2, objChildElement, objChildElement2)
                
        objParentNode.appendChild objDOM.createTextNode(vbNewLine)
               
        'PLDA DV1 - included when Q1 = "C602" or "N934"
        If G_rstDetails.RecordCount > 0 Then
            
            For lngDetailCtr = 1 To G_rstDetails.RecordCount
            
                G_rstDetailsDocumenten.Filter = adFilterNone
                G_rstDetailsDocumenten.Filter = "Detail = " & lngDetailCtr
                
                G_rstDetailsDocumenten.Sort = "Ordinal ASC"
                
                G_rstDetailsDocumenten.MoveFirst
                
                If G_rstDetailsDocumenten.RecordCount > 0 Then
                    Do While Not G_rstDetailsDocumenten.EOF
                        If IsNull(G_rstDetailsDocumenten.Fields("Q1").Value) = False Then
                            If UCase(Trim$(FNullField(G_rstDetailsDocumenten.Fields("Q1").Value))) = "C602" Or _
                               UCase(Trim$(FNullField(G_rstDetailsDocumenten.Fields("Q1").Value))) = "N934" Then
                                blnDV1Found = True
                                Exit Do
                            End If
                        End If
                        
                        If Trim$(FNullField(G_rstDetailsDocumenten.Fields("QA").Value)) = "E" Then Exit Do
                        
                        G_rstDetailsDocumenten.MoveNext
                    Loop
                End If
                
                If blnDV1Found = True Then Exit For
            Next
        
        End If
    
        If blnDV1Found = True Then
            Set objChildNode = objParentNode.appendChild(objDOM.createElement("PLDADV1"))
            objDOM.documentElement.appendChild objChildNode
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
            Call CreateImportDV1Header(objDOM, objParentNode, objChildNode, objChildNode2, objChildElement, objChildElement2)
        
            Call CreateImportDV1GoodsItem(objDOM, objParentNode, objChildNode, objChildNode2, objChildElement, objChildElement2)
        End If
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine)
    
    Set objChildNode2 = Nothing
    Set objChildElement2 = Nothing
    Set objChildElement = Nothing

End Sub
 
Private Sub CreateImportInterchangeHeader(ByVal XMLMessageType As m_enumXMLMessage, _
                                          ByRef objDOM As DOMDocument, _
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
    
        'PLDA Sender
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageSender"))
        objChildElement.Text = G_strMessageSender
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'PLDA Recipient
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("messageRecipient"))
        objChildElement.Text = G_strMessageRecipient
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
        objChildElement.Text = G_strIEFunctionCode
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
    
End Sub

Private Sub CreateImportGoodsDeclaration(ByRef objDOM As DOMDocument, _
                                         ByRef objParentNode As IXMLDOMNode, _
                                         ByRef objChildNode As IXMLDOMNode, _
                                         ByRef objChildElement As IXMLDOMElement, _
                                         ByRef objChildElement2 As IXMLDOMElement)
    
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement
    Dim objChildElement5 As IXMLDOMElement
    
    Dim strAuthCodes As String
    Dim lngDocumenten As Long
    Dim lngDetailCtr As Long
    
    '>>>>>>>>>>>>>>>>>>>>>

    '>>>>>>>>>>>>>>>>>>>>>>
    
    'Loading List
    If ((GenerateData(G_rstHeader, "A8") <> "" And GenerateData(G_rstHeader, "A8") <> 0) Or _
       G_BlnDoNotIncludeIfEmpty = False) Or (GenerateData(G_rstHeader, "A1") = "AC" And GenerateData(G_rstHeader, "A2") = "4") Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("loadingList"))
        objChildElement.Text = GenerateData(G_rstHeader, "A8")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Local Reference Number
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("localReferenceNumber"))
    objChildElement.Text = GenerateData(G_rstHeader, "A3")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    'Commercial Reference Number
    If GenerateData(G_rstHeader, "AC") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("commercialReference"))
        objChildElement.Text = GenerateData(G_rstHeader, "AC")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Invoice Number
    If G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("invoiceNumber")) 'Edwin Jan14
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    '**********************************************************************************************
    'Totals +
    '**********************************************************************************************
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Totals"))
    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("items"))
        objChildElement2.Text = G_rstDetails.RecordCount 'Edwin Jan14
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalGrossmass"))
        objChildElement2.Text = GenerateData(G_rstHeader, "D1")
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
        objChildElement2.Text = GenerateData(G_rstHeader, "D3")
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    '**********************************************************************************************
    
    '**********************************************************************************************
    'Transaction Nature +
    '**********************************************************************************************
    If GenerateData(G_rstHeader, "C7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TransactionNature"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature1"))
            objChildElement2.Text = Left$(GenerateData(G_rstHeader, "C7"), 1) 'Edwin Jan14
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            If Len(Trim$(GenerateData(G_rstHeader, "C7"))) = 2 Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature2"))
                
                If Len(Trim$(GenerateData(G_rstHeader, "C7"))) = 2 Then
                    objChildElement2.Text = Trim$(Right$(GenerateData(G_rstHeader, "C7"), 1))
                Else
                    objChildElement2.Text = vbNullString
                End If
                
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    '**********************************************************************************************
    
    'Registration Number
    If G_strIEFunctionCode = "IE" & m_enumXMLMessage.enumImportDeclaration Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("registrationNumber"))
        objChildElement.Text = GenerateData(G_rstHeader, "AB")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Issue Place
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("issuePlace"))
    objChildElement.Text = GenerateData(G_rstHeader, "A5")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    'Signature +
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("signature"))
    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Signature Operator Contact Name
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorContactName"))
        objChildElement2.Text = GenerateData(G_rstHeader, "AA")
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Signature Capacity
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("capacity"))
        objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XG", 2)
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
    'Type Part One
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartOne"))
    objChildElement.Text = GenerateData(G_rstHeader, "A1")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    'Type Part Two
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartTwo"))
    objChildElement.Text = GenerateData(G_rstHeader, "A2")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    '**********************************************************************************************
    'Consignee +
    '**********************************************************************************************
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Consignee"))
    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Consignee Operator Identity +
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
        
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1")
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Consignee Operator +
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
        
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2")
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
            objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5")
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3")
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4")
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6")
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8")
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    '**********************************************************************************************
    
    '**********************************************************************************************
    'Transport Means +
    '**********************************************************************************************
    If GenerateData(G_rstHeader, "C2") <> "" Or _
       GenerateData(G_rstHeader, "C3") <> "" Or _
       GenerateData(G_rstHeader, "D8") <> "" Or _
       GenerateData(G_rstHeader, "D7") <> "" Or _
       GenerateData(G_rstHeader, "D5") <> "" Or _
       GenerateData(G_rstHeader, "D9") <> "" Or _
       G_rstDetailsContainer.RecordCount > 0 Or _
       G_BlnDoNotIncludeIfEmpty = False Then
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TransportMeans"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
            'Transport Means Delivery Terms +
            If GenerateData(G_rstHeader, "C2") <> "" Or _
               GenerateData(G_rstHeader, "C3") <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
                
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DeliveryTerms"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Transport Means Delivery Terms
                    If GenerateData(G_rstHeader, "C2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("deliveryTerms"))
                        objChildElement3.Text = GenerateData(G_rstHeader, "C2")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    End If
                                    
                    'Transport Means Delivery Terms Place
                    If GenerateData(G_rstHeader, "C3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("deliveryTermsPlace"))
                        objChildElement3.Text = GenerateData(G_rstHeader, "C3")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
            If GenerateData(G_rstHeader, "D8") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("borderMode"))
                objChildElement2.Text = GenerateData(G_rstHeader, "D8")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
            If GenerateData(G_rstHeader, "D7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("borderNationality"))
                objChildElement2.Text = GenerateData(G_rstHeader, "D7")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
            If GenerateData(G_rstHeader, "D5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("departureIdentity"))
                objChildElement2.Text = GenerateData(G_rstHeader, "D5")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
                        
            If GenerateData(G_rstHeader, "D9") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("inlandMode"))
                objChildElement2.Text = GenerateData(G_rstHeader, "D9")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
            If G_rstDetailsContainer.RecordCount > 0 Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("container"))
                
                'CSCLP-354
                Dim blnContainersHaveValue As Boolean
                Dim lngCurrentDetail As Long
                Dim blnGoNextDetail As Boolean
                
                G_rstDetailsContainer.Sort = "DETAIL ASC, ORDINAL ASC"
                G_rstDetailsContainer.MoveFirst
                
                'Memo current detail tab
                lngCurrentDetail = G_rstDetailsContainer.Fields("Detail").Value
                
                Do Until G_rstDetailsContainer.EOF
                    If blnGoNextDetail = False Then
                        If (Len((FNullField(G_rstDetailsContainer.Fields("S4").Value))) > 0 _
                            And FNullField(G_rstDetailsContainer.Fields("S4").Value <> 0)) _
                            Or _
                           (Len((FNullField(G_rstDetailsContainer.Fields("S5").Value))) > 0 _
                            And FNullField(G_rstDetailsContainer.Fields("S5").Value <> 0)) Then
                            'Container value found, give up
                            blnContainersHaveValue = True
                            Exit Do
                        Else
                            If UCase((FNullField(G_rstDetailsContainer.Fields("S6").Value))) = "E" Then
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
                            If lngCurrentDetail = FNullField(G_rstDetailsContainer.Fields("Detail").Value) Then
                                G_rstDetailsContainer.MoveNext
                            Else
                                'Memo new detail tab
                                lngCurrentDetail = FNullField(G_rstDetailsContainer.Fields("Detail").Value)
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
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    '**********************************************************************************************
                            
    '**********************************************************************************************
    'Customs
    '**********************************************************************************************
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Customs"))
    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
           
        'Customs Goods Location +
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("GoodsLocation"))
        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            
            'Customs Goods Location Precise
            If GenerateData(G_rstHeader, "D4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("precise"))
                objChildElement3.Text = GenerateData(G_rstHeader, "D4")
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            End If
            
            'Customs Goods Location Postal Code
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("postalcode"))
            objChildElement3.Text = GenerateData(G_rstHeader, "DG")
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
        'Customs Office Destination
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("customsOfficeDestination"))
        objChildElement2.Text = GenerateData(G_rstHeader, "A6")
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    '**********************************************************************************************
    
    '**********************************************************************************************
    'Invoice Issuer +
    '**********************************************************************************************
    If G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("invoiceIssuer"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Invoice Issuer +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("invoiceIssuer"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            
                'Operator Identity +
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorIdentity"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                    'Invoice Issuer Operator Identity
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("operatorIdentity"))
                    objChildElement4.Text = ""
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
                'Invoice Issuer Operator +
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Operator"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Invoice Issuer Operator Name
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("operatorName"))
                    objChildElement4.Text = ""
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                                    
                    'Invoice Issuer Operator Address +
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("OperatorAddress"))
                    objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Invoice Issuer Postal Code
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("postalCode"))
                        objChildElement5.Text = ""
                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Invoice Issuer Street and Number 1
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("streetAndNumber1"))
                        objChildElement5.Text = ""
                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Invoice Issuer Street and Number 2
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("streetAndNumber2"))
                        objChildElement5.Text = ""
                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Invoice Issuer City
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("city"))
                        objChildElement5.Text = ""
                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Invoice Issuer Country
                        Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("country"))
                        objChildElement5.Text = ""
                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
            'Contact Person +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("contactPerson"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Contact Person Name
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonName"))
                objChildElement3.Text = ""
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Contact Person Telephone Number
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonTelNumber"))
                objChildElement3.Text = ""
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Contact Person Fax Number
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonFaxNumber"))
                objChildElement3.Text = ""
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Contact Person Email
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonEmail"))
                objChildElement3.Text = ""
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    '**********************************************************************************************
            
    '**********************************************************************************************
    'Representative +
    '**********************************************************************************************
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Representative"))
    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Operator Identity +
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
        
            'Representative Operator Identity
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1", 4)
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Operator +
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
        
            'Representative Name
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
            objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 4)
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                        
            'Representative Address +
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
            objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Representative Postal Code
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 4)
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Representative Street and Number 1
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 4)
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Representative Street and Number 2
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 4)
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Representative City
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 4)
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Representative Country
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 4)
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
        
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
    
        'Registration Number
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("registrationNumber"))
        objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XD", 4)
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        'Declarant Status
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("declarantStatus"))
        objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XF", 4)
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
        'Should not be sent according to version 2.2 of PLDA Lux MIG - June 10, 2010
        'Representative Contact Person Name
        'Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("contactPersonName"))
        'objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "X9", 4)
        'objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
        'Authorised Identity
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisedIdentity"))
        objChildElement2.Text = GenerateData(G_rstHeaderHandelaars, "XH", 4)
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    '**********************************************************************************************
        
    '**********************************************************************************************
    'Taxes +
    '**********************************************************************************************
    If GenerateData(G_rstDetailsZelf, "U1") <> "" Or _
       GenerateData(G_rstDetailsZelf, "U2") <> "" Or _
       G_BlnDoNotIncludeIfEmpty = False Then
        
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("taxes"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Tax Calculation +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taxCalculation"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Tax Type
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxType"))
                objChildElement3.Text = GenerateData(G_rstDetailsZelf, "U1")
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Tax Base
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxBase"))
                objChildElement3.Text = ""
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                'Tax Amount
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxAmount"))
                objChildElement3.Text = GenerateData(G_rstDetailsZelf, "U2")
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Total Amount +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalAmount"))
            objChildElement2.Text = ""
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Total Currency +
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalCurrency"))
            objChildElement2.Text = ""
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    '**********************************************************************************************
            
    '**********************************************************************************************
    'Finance +
    '**********************************************************************************************
    If (GenerateData(G_rstHeader, "C6") <> "" And GenerateData(G_rstHeader, "C5") <> "") Or _
       GenerateData(G_rstHeader, "C4") <> "" Or _
       G_BlnDoNotIncludeIfEmpty = False Then
       
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Finance"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Invoice Amount
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("invoiceAmount"))
            objChildElement2.Text = GenerateData(G_rstHeader, "C4")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Invoice Amount Estimated Flag
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("invoiceAmountEstimatedFlag"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Exchange Rate +
            If (GenerateData(G_rstHeader, "C6") <> "" And GenerateData(G_rstHeader, "C5") <> "") Or _
               G_BlnDoNotIncludeIfEmpty = False Then
            
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("exchangeRate"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("exchangeRate"))
                    objChildElement3.Text = GenerateData(G_rstHeader, "C6")
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("currency"))
                    objChildElement3.Text = GenerateData(G_rstHeader, "C5")
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
            
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("exchangeDate"))
                    objChildElement3.Text = GenerateData(G_rstHeader, "A4")
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
            
            If G_BlnDoNotIncludeIfEmpty = False Then
                'Transport EU Border +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TransportEUBorder"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Transport EU Border Charges
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Charges"))
                    objChildElement3.Text = ""
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Transport EU Border Estimated Flag
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement3.Text = "0"
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Transport EU Inland +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TransportEUInland"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Transport EU Inland Charges
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Charges"))
                    objChildElement3.Text = ""
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Transport EU Inland Estimated Flag
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement3.Text = "0"
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Insurance +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Insurance"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Insurance Charges
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Charges"))
                    objChildElement3.Text = ""
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Insurance Estimated Flag
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement3.Text = "0"
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Other EU Border +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OtherEUBorder"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Other EU Border Charges
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Charges"))
                    objChildElement3.Text = ""
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Other EU Border Estimated Flag
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement3.Text = "0"
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                            
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Other EU Inland +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OtherEUInland"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Other EU Inland Charges
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Charges"))
                    objChildElement3.Text = ""
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Other EU Inland Estimated Flag
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement3.Text = "0"
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                                    
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            End If
                        
            'OtherCharges CSCLP-671
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OtherCharges"))
            objChildElement2.Text = GenerateData(G_rstHeader, "C9")
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    '**********************************************************************************************
    
    'Method of payment
    If GenerateData(G_rstHeader, "B1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("methodOfPayment"))
        objChildElement.Text = GenerateData(G_rstHeader, "B1")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Deferred Payment - CSCLP-631
    If GenerateData(G_rstHeader, "B5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("deferredPayment"))
        objChildElement.Text = GenerateData(G_rstHeader, "B5")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Destination Country
    If GenerateData(G_rstDetails, "N7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("destinationCountry"))
        objChildElement.Text = GenerateData(G_rstDetails, "N7")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Dispatch Country
    If GenerateData(G_rstHeader, "DA") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("dispatchCountry"))
        objChildElement.Text = GenerateData(G_rstHeader, "DA")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Customs Main Reference Number
    If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumImportDeclaration) Or _
       G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("customsMainReferenceNumber"))
        objChildElement.Text = GenerateData(G_rstHeader, "MRN")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
                        
    'Invoice Date
    If G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("invoiceDate"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
            
    'Invoice Place
    If G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("invoicePlace"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Declaration Date
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationDate"))
    objChildElement.Text = GenerateData(G_rstHeader, "A4")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

    'Declaration Period From Date
    If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumImportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationPeriodFromDate"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Declaration Period To Date
    If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumImportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("declarationPeriodToDate"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Incomplete Declaration Reason
    If (G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumImportDeclaration) Or G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("incompleteDeclarationReason"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
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
                    
                    If Trim$(FNullField(G_rstDetailsDocumenten.Fields("QA").Value)) = "E" Then Exit Do
                    
                    G_rstDetailsDocumenten.MoveNext
                Loop
            End If
            
        Next
    
    End If
    '**********************************************************************************************
                                            
    ' Value Details Flag
    If G_BlnDoNotIncludeIfEmpty = False Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("valueDetailsFlag"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
                                            
    If G_strIEFunctionCode = "IE" & m_enumXMLMessage.enumImportAmendment Then
        'Amendment Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentDate"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
        'Amendment Place
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("amendmentPlace"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    End If
    
    'Acceptance Date
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("acceptanceDate"))
    objChildElement.Text = GenerateData(G_rstHeader, "A4")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
    
    If G_strIEFunctionCode <> "IE" & m_enumXMLMessage.enumImportDeclaration Then
        'Release Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("releaseDate"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
        'Release Time
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("releaseTime"))
        objChildElement.Text = ""
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    End If
    
    Set objChildElement3 = Nothing
    Set objChildElement4 = Nothing
    Set objChildElement5 = Nothing
    
End Sub

Private Sub CreateImportProcedureType(ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode)
    
    '*****************
    'Procedure Type
    '*****************
    Set objChildNode = objDOM.createElement("ProcedureType")
    objChildNode.Text = GenerateData(G_rstDetails, "N4")
    
    objDOM.documentElement.appendChild objChildNode
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

End Sub

Private Sub CreateImportGoodsItem(ByRef objDOM As DOMDocument, _
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
    Dim lngBerekenings As Long
    'Dim lngZelf As Long
    
    Dim strCheckMass As String
    
    'CSCLP-673 TYPE MISMATCH MOD DOUBLE TO STRING
    Dim dblInvoiceAmt As Double
    Dim dblAdaptCharge As Double
    
    If (G_rstDetails.RecordCount > 0) Then
        G_rstDetails.MoveFirst
    End If
    
    For lngDetailCounter = 1 To G_rstDetails.RecordCount
        
        If G_rstDetailsBijzondere.RecordCount > 0 Then G_rstDetailsBijzondere.MoveFirst
        If G_rstDetailsContainer.RecordCount > 0 Then G_rstDetailsContainer.MoveFirst
        If G_rstDetailsDocumenten.RecordCount > 0 Then G_rstDetailsDocumenten.MoveFirst
        If G_rstDetailsBerekeningsEenheden.RecordCount > 0 Then G_rstDetailsBerekeningsEenheden.MoveFirst
        If G_rstDetailsZelf.RecordCount > 0 Then G_rstDetailsZelf.MoveFirst
        
        lngBijzondere = 0
        lngContainer = 0
        lngDocumenten = 0
        lngBerekenings = 0
        'lngZelf = 0
        
        'Goods Item
        Set objChildNode = objParentNode.appendChild(objDOM.createElement("GoodsItem"))
        objDOM.documentElement.appendChild objParentNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
            'Item Number
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ItemNumber"))
            objChildElement.Text = lngDetailCounter
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Commodity Code +
            If GenerateData(G_rstDetails, "L1") <> "" Or _
               GenerateData(G_rstDetails, "L2") <> "" Or _
               GenerateData(G_rstDetails, "L3") <> "" Or _
               GenerateData(G_rstDetails, "L4") <> "" Or _
               GenerateData(G_rstDetails, "L5") <> "" Or _
               GenerateData(G_rstDetails, "L6") <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("commodityCode"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    'Commodity Code
                    If GenerateData(G_rstDetails, "L1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("commodityCode"))
                        objChildElement2.Text = Left$(GenerateData(G_rstDetails, "L1"), 8)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'Commodity Code Taric
                    If GenerateData(G_rstDetails, "L1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("taric"))
                        objChildElement2.Text = Right$(GenerateData(G_rstDetails, "L1"), 2)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'First Additional Commodity
                    If GenerateData(G_rstDetails, "L2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("firstAdditionalCommodity"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "L2")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                                        
                    'Second Additional Commodity
                    If GenerateData(G_rstDetails, "L3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("secondAdditionalCommodity"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "L3")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'National Additional Commodity 1
                    If GenerateData(G_rstDetails, "L4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity1"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "L4")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'National Additional Commodity 2
                    If GenerateData(G_rstDetails, "L5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity2"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "L5")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    'National Additional Commodity 3
                    If GenerateData(G_rstDetails, "L6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("nationalAdditionalCommodity3"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "L6")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Net Mass
            If GenerateData(G_rstDetails, "LA") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("netMass"))
                
                strCheckMass = GenerateData(G_rstDetails, "LA")
                If Val(strCheckMass) = 0 Then
                    strCheckMass = ""
                End If
                
                'FINALIZE
                objChildElement.Text = strCheckMass
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Gross Mass
            If (GenerateData(G_rstDetails, "L9") <> "" And GenerateData(G_rstDetails, "L9") <> "0") Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("grossMass"))
                
                'CSCLP-209
'                strCheckMass = GenerateData(G_rstDetails, "L9")
'                If Val(strCheckMass) = 0 Then
'                    strCheckMass = ""
'                End If
                
                'FINALIZE
                objChildElement.Text = GenerateData(G_rstDetails, "L9") 'strCheckMass
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Goods Description
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("goodsDescription"))
            objChildElement.Text = GenerateData(G_rstDetails, "L8")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
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
                            
            '**********************************************************************************************
            'Container Identifier
            '**********************************************************************************************
            G_rstDetailsContainer.Filter = adFilterNone
            G_rstDetailsContainer.Filter = "Detail = " & lngDetailCounter
            
            If G_rstDetailsContainer.RecordCount > 0 Then
                
                G_rstDetailsContainer.Sort = "ORDINAL ASC"
                G_rstDetailsContainer.MoveFirst
                
                Do Until G_rstDetailsContainer.EOF
                    lngContainer = lngContainer + 1
                    
                    'S4
                    strDummy = Trim$(GenerateData(G_rstDetailsContainer, "S4", lngContainer))
                    If LenB(strDummy) > 0 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("containerIdentifier"))
                        objChildElement.Text = strDummy
                        
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'S5
                    strDummy = Trim$(GenerateData(G_rstDetailsContainer, "S5", lngContainer))
                    If LenB(strDummy) > 0 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("containerIdentifier"))
                        objChildElement.Text = strDummy
                        
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    If Trim$(FNullField(G_rstDetailsContainer.Fields("S6").Value)) = "E" Then Exit Do
                    
                    G_rstDetailsContainer.MoveNext
                Loop
                
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Produced Document +
            '**********************************************************************************************
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
                                
                            'If GenerateData(G_rstDetailsDocumenten, "Q9", lngDocumenten) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then CSCLP-632
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("complementaryInformation"))
                                objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q9", lngDocumenten)
                                
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                            'End If
                                
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    '****************************************************************************************
                    
                    If Trim$(FNullField(G_rstDetailsDocumenten.Fields("QA").Value)) = "E" Then Exit Do
                    
                    G_rstDetailsDocumenten.MoveNext
                Loop
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Supplementary Units +
            '**********************************************************************************************
            If GenerateData(G_rstDetails, "M2") <> "" Or _
               GenerateData(G_rstDetails, "M1") <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("SupplementaryUnits"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    If GenerateData(G_rstDetails, "M2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("supplementaryUnits"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "M2")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    End If
                    
                    If GenerateData(G_rstDetails, "M1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("supplementaryUnitsCode"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "M1")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
            End If
            '**********************************************************************************************
            
            'Origin Country
            If GenerateData(G_rstDetails, "NB") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("originCountry"))
                objChildElement.Text = GenerateData(G_rstDetails, "NB")
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            'Statistical Value
            If GenerateData(G_rstDetails, "O2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("statisticalValue"))
                objChildElement.Text = GenerateData(G_rstDetails, "O2")
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            '******************************************************************************************************
            'Customs Treatment + ( Procedure and Warehouse was interchanged based on version 2.2 of PLDA Lux MIG )
            '******************************************************************************************************
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CustomsTreatment"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                'Procedure +
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Procedure"))
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
                    
                'quota CSCLP-663
                If GenerateData(G_rstDetails, "L7") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("quota"))
                    objChildElement2.Text = GenerateData(G_rstDetails, "L7")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                    
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Calculation Units +
            '**********************************************************************************************
            G_rstDetailsBerekeningsEenheden.Filter = adFilterNone
            G_rstDetailsBerekeningsEenheden.Filter = "Detail = " & lngDetailCounter
            
            If G_rstDetailsBerekeningsEenheden.RecordCount > 0 Then
                G_rstDetailsBerekeningsEenheden.Sort = "ORDINAL ASC"
                G_rstDetailsBerekeningsEenheden.MoveFirst
                
                Do Until G_rstDetailsBerekeningsEenheden.EOF
                    lngBerekenings = lngBerekenings + 1
                    
                    If GenerateData(G_rstDetailsBerekeningsEenheden, "T8", lngBerekenings) <> "" Or _
                       GenerateData(G_rstDetailsBerekeningsEenheden, "TZ", lngBerekenings) <> "" Or _
                       G_BlnDoNotIncludeIfEmpty = False Then
                        
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CalculationUnits"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                        
                            'Claculation Units
                            If GenerateData(G_rstDetailsBerekeningsEenheden, "T8", lngBerekenings) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("calculationUnits"))
                            objChildElement2.Text = GenerateData(G_rstDetailsBerekeningsEenheden, "T8", lngBerekenings)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Calculation Units Code
                            If GenerateData(G_rstDetailsBerekeningsEenheden, "TZ", lngBerekenings) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("calculationUnitsCode"))
                                objChildElement2.Text = GenerateData(G_rstDetailsBerekeningsEenheden, "TZ", lngBerekenings)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                                                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                    End If
                    
                    If Trim$(FNullField(G_rstDetailsBerekeningsEenheden.Fields("T9").Value)) = "E" Then Exit Do
                    
                    G_rstDetailsBerekeningsEenheden.MoveNext
                Loop
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Additional Information +
            '**********************************************************************************************
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
                    
                    If Trim$(FNullField(G_rstDetailsBijzondere.Fields("P5").Value)) = "E" Then Exit Do
                    
                    G_rstDetailsBijzondere.MoveNext
                Loop
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Previous Document +
            '**********************************************************************************************
            If GenerateData(G_rstDetails, "R5") <> "" Or _
               GenerateData(G_rstDetails, "R2") <> "" Or _
               GenerateData(G_rstDetails, "R1") <> "" Or _
               G_BlnDoNotIncludeIfEmpty = False Then
                   
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("PreviousDocument"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    'Document +
                    If GenerateData(G_rstDetails, "R5") <> "" Or _
                       GenerateData(G_rstDetails, "R2") <> "" Or _
                       G_BlnDoNotIncludeIfEmpty = False Then
                        
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Document"))
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
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
                                
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Previous Document Type
                            If GenerateData(G_rstDetails, "R2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentType"))
                                objChildElement3.Text = GenerateData(G_rstDetails, "R2")
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                            
                            'Previous Document Date
                            If GenerateData(G_rstDetails, "R3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentDate"))
                                objChildElement3.Text = GenerateData(G_rstDetails, "R3")
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            End If
                        
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                        
                    End If
                                        
                    'Previous Document Category
                    If GenerateData(G_rstDetails, "R1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("previousDocumentCategory"))
                        objChildElement2.Text = GenerateData(G_rstDetails, "R1")
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Price +
            '**********************************************************************************************
            'If (GenerateData(G_rstDetails, "O5") <> "" And _
               GenerateData(G_rstDetails, "O1") <> "") Or _
               G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("price"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("price"))
                    'objChildElement2.Text = GenerateData(G_rstDetails, "O5")
                    dblInvoiceAmt = Val(IIf(GenerateData(G_rstDetails, "O5") = "", 0, GenerateData(G_rstDetails, "O5")))
                    dblAdaptCharge = Val(IIf(GenerateData(G_rstDetails, "O7") = "", 0, GenerateData(G_rstDetails, "O7")))
                    objChildElement2.Text = Replace(CStr(dblInvoiceAmt + dblAdaptCharge), ",", ".")

                    
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("valuationMethod"))
                    objChildElement2.Text = GenerateData(G_rstDetails, "O1")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("estimatedFlag"))
                    objChildElement2.Text = "0"
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Taxes +
            '**********************************************************************************************
            If G_BlnDoNotIncludeIfEmpty = False Then
               
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("taxes"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                    'Tax Calculation
                    Do While Not G_rstDetailsZelf.EOF
                        If G_rstDetailsZelf![DETAIL] = lngDetailCounter Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TaxCalculation"))
                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxType"))
                                objChildElement3.Text = GenerateData(G_rstDetailsZelf, "U1")
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                                
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxBase"))
                                objChildElement3.Text = ""
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxAmount"))
                                objChildElement3.Text = GenerateData(G_rstDetailsZelf, "U2") '""
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("paymentMethodTaxes"))
                                objChildElement3.Text = ""
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("taxRate"))
                                objChildElement3.Text = ""
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                            
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        End If
                        G_rstDetailsZelf.MoveNext
                    Loop
                                        
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalAmount"))
                    objChildElement2.Text = ""
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
            End If
            '**********************************************************************************************
            
            '**********************************************************************************************
            'Preference +
            '**********************************************************************************************
            If Len(Trim$(GenerateData(G_rstDetails, "N5"))) = 3 Then
            
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("preference"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("preference1"))
                    objChildElement2.Text = Left(Trim$(GenerateData(G_rstDetails, "N5")), 1)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("preference2"))
                    objChildElement2.Text = Right(Trim$(GenerateData(G_rstDetails, "N5")), 2)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            End If
            '**********************************************************************************************
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
        Set objChildElement3 = Nothing
    
        G_rstDetails.MoveNext
        
    Next lngDetailCounter
    
End Sub

Private Sub CreateImportDV1Header(ByRef objDOM As DOMDocument, _
                        ByRef objParentNode As IXMLDOMNode, _
                        ByRef objChildNode As IXMLDOMNode, _
                        ByRef objChildNode2 As IXMLDOMNode, _
                        ByRef objChildElement As IXMLDOMElement, _
                        ByRef objChildElement2 As IXMLDOMElement)
    
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement
    
    Dim lngInvoiceCtr As Long
    Dim blnQ2Found As Boolean
    
    'DV1 Header
    Set objChildNode2 = objChildNode.appendChild(objDOM.createElement("DV1Header"))
    objDOM.documentElement.appendChild objChildNode
    objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
        'Forms
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Forms"))
        objChildElement.Text = "1"
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
'        'Declaration Date
'        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("declarationDate"))
'        objChildElement.Text = GenerateData(G_rstHeader, "A4")
'        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
'        'Issue Place
'        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("issuePlace"))
'        objChildElement.Text = GenerateData(G_rstHeader, "A5")
'        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
'        'Customs Reference
'        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("CustomsReference"))
'        objChildElement.Text = ""
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Local Reference Number
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("localReferenceNumber"))
'            objChildElement2.Text = GenerateData(G_rstHeader, "A3")
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Type Part One
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("typePartOne"))
'            objChildElement2.Text = GenerateData(G_rstHeader, "A1")
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Type Part Two
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("typePartTwo"))
'            objChildElement2.Text = GenerateData(G_rstHeader, "A2")
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Customs Main Reference Number
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("customsMainReferenceNumber"))
'            objChildElement2.Text = GenerateData(G_rstHeader, "MRN")
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Acceptance Date
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AcceptanceDate"))
'            objChildElement2.Text = ""
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Customs Office Destination
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("customsOfficeDestination"))
'            objChildElement2.Text = ""
'
'        'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        
        '***********************************************************************************************
        'Trader Seller + CSCLP-637 As per Benny's Instruction these tags are mapped to intracom
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("TraderSeller"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

'            'Operator Identity
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
'            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
'                objChildElement3.Text = GenerateData(G_rstMain, "Supplier Name")
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'Operator
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 3)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 3)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 3)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 3)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 3)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 3)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************
        
        '***********************************************************************************************
        'Trader Buyer +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("TraderBuyer"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'Operator Identity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X1", 1)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'Operator
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
                objChildElement3.Text = GenerateData(G_rstHeaderHandelaars, "X2", 1)
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X5", 1)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X3", 1)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X4", 1)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X6", 1)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)

                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                    objChildElement4.Text = GenerateData(G_rstHeaderHandelaars, "X8", 1)
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************
        
'        '***********************************************************************************************
'        'Representative +
'        '***********************************************************************************************
'        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Representative"))
'        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Operator Identity
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))
'            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
'                objChildElement3.Text = ""
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Operator
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))
'            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
'                objChildElement3.Text = ""
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
'                objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
'                    objChildElement4.Text = ""
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
'                    objChildElement4.Text = ""
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
'                    objChildElement4.Text = ""
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
'                    objChildElement4.Text = ""
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
'                    objChildElement4.Text = ""
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Authorised Identity
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisedIdentity"))
'            objChildElement2.Text = ""
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
'        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'        '***********************************************************************************************
            
        '***********************************************************************************************
        'Delivery Terms +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("DeliveryTerms"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Delivery Terms
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("deliveryTerms"))
            objChildElement2.Text = GenerateData(G_rstHeader, "C2")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'Delivery Terms Place
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("deliveryTermsPlace"))
            objChildElement2.Text = GenerateData(G_rstHeader, "C3")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************
        
        '***********************************************************************************************
        'Invoice +
        '***********************************************************************************************
        G_rstDetails.Filter = adFilterNone
        G_rstDetails.MoveFirst
        
        If G_rstDetails.RecordCount > 0 Then
            
            For lngInvoiceCtr = 1 To G_rstDetails.RecordCount
                
                G_rstDetailsDocumenten.Filter = adFilterNone
                G_rstDetailsDocumenten.Filter = "Detail = " & lngInvoiceCtr
                
                
                If G_rstDetailsDocumenten.RecordCount > 0 Then
                    G_rstDetailsDocumenten.MoveFirst
                    
                    Do While Not G_rstDetailsDocumenten.EOF
                        
                        If LenB(Trim$(GenerateData(G_rstDetailsDocumenten, "Q2"))) > 0 Then
                            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Invoice"))
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                                'Document Reference
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentReference"))
                                objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q2")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                
                                'Document Date
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentDate"))
                                objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q3")
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                
                            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            blnQ2Found = True
                            Exit Do
                        End If
                        
                        G_rstDetailsDocumenten.MoveNext
                    Loop
                    
                End If
                
                If blnQ2Found = True Then Exit For
            Next
        End If
        '***********************************************************************************************
        
        '***********************************************************************************************
        'Contract +
        '***********************************************************************************************
'        G_rstDetails.Filter = adFilterNone
'        G_rstDetails.MoveFirst
'
'        If G_rstDetails.RecordCount > 0 Then
'
'            For lngInvoiceCtr = 1 To G_rstDetails.RecordCount
'
'                G_rstDetailsDocumenten.Filter = adFilterNone
'                G_rstDetailsDocumenten.Filter = "Detail = " & lngInvoiceCtr
'
'
'                If G_rstDetailsDocumenten.RecordCount > 0 Then
'                    G_rstDetailsDocumenten.MoveFirst
'
'                    Do While Not G_rstDetailsDocumenten.EOF
'
'                        If LenB(Trim$(GenerateData(G_rstDetailsDocumenten, "Q2"))) > 0 Then
'                            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Contract"))
'                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                                'Document Reference
'                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentReference"))
'                                objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q2")
'                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                                'Document Date
'                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentDate"))
'                                objChildElement2.Text = GenerateData(G_rstDetailsDocumenten, "Q3")
'                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
'                            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
'                            blnQ2Found = True
'                            Exit Do
'                        End If
'
'                        G_rstDetailsDocumenten.MoveNext
'                    Loop
'
'                End If
'
'                If blnQ2Found = True Then Exit For
'            Next
'        End If
        '***********************************************************************************************
        
        '***********************************************************************************************
        'Previous Customs Decision +
        '***********************************************************************************************
        If G_BlnDoNotIncludeIfEmpty = False Then
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("PreviousCustomsDecision"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Document Reference
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentReference"))
                objChildElement2.Text = ""
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Document Date
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocumentDate"))
                objChildElement2.Text = ""
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        End If
        '***********************************************************************************************
        
        '***********************************************************************************************
        'Relationship +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("RelationShip"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'YesNo Flag1
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag1"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

'            'YesNo Flag2
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag2"))
'            objChildElement2.Text = "0"
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'YesNo Flag3
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag3"))
'            objChildElement2.Text = "0"
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'            'Additional Information Content
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContent"))
'            objChildElement2.Text = ""
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************

        '***********************************************************************************************
        'Restrictions +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Restrictions"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'YesNo Flag1
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag1"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'YesNo Flag2
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag2"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

'            'Additional Information Content
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContent"))
'            objChildElement2.Text = ""
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************

        '***********************************************************************************************
        'Royalties License Fees +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("RoyaltiesLicenseFees"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'YesNo Flag1
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag1"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

            'YesNo Flag2
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("YesNoFlag2"))
            objChildElement2.Text = "0"
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

'            'Additional Information Content
'            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContent"))
'            objChildElement2.Text = ""
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************
                                                                
        '***********************************************************************************************
        'signature +
        '***********************************************************************************************
        Set objChildElement = objChildNode2.appendChild(objDOM.createElement("signature"))
        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
            'operatorContactName
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorContactName"))
            objChildElement2.Text = GenerateData(G_rstHeader, "AA")
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
        
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
        '***********************************************************************************************
    
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
    Set objChildElement3 = Nothing
    Set objChildElement4 = Nothing
    
End Sub

Private Sub CreateImportDV1GoodsItem(ByRef objDOM As DOMDocument, _
                                     ByRef objParentNode As IXMLDOMNode, _
                                     ByRef objChildNode As IXMLDOMNode, _
                                     ByRef objChildNode2 As IXMLDOMNode, _
                                     ByRef objChildElement As IXMLDOMElement, _
                                     ByRef objChildElement2 As IXMLDOMElement)
            
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement
    Dim objChildElement5 As IXMLDOMElement
    
    Dim lngDetailCounter As Long
    'revert
    Dim dblInvoiceAmt As Double
    Dim dblAdaptCharge As Double
    
    If (G_rstDetails.RecordCount > 0) Then
        G_rstDetails.MoveFirst
    End If
    
    For lngDetailCounter = 1 To G_rstDetails.RecordCount
        
        Set objChildNode2 = objChildNode.appendChild(objDOM.createElement("DV1GoodsItem"))
        objDOM.documentElement.appendChild objChildNode
        objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Item Number
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("ItemNumber"))
            objChildElement.Text = lngDetailCounter
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'SAD Reference
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("SADReference"))
            objChildElement.Text = lngDetailCounter
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Amount 'CSCLP-669
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Amount"))
            objChildElement.Text = GenerateData(G_rstDetails, "O5")
            dblInvoiceAmt = Val(IIf(GenerateData(G_rstDetails, "O5") = "", 0, GenerateData(G_rstDetails, "O5"))) 'revert
            dblAdaptCharge = Val(IIf(GenerateData(G_rstDetails, "O7") = "", 0, GenerateData(G_rstDetails, "O7")))
            objChildElement.Text = Replace(CStr(dblInvoiceAmt + dblAdaptCharge), ",", ".")
            
            
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            '***********************************************************************************************
            'Basis Of Calculation +
            '***********************************************************************************************
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("BasisOfCalculation"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
            
                'Item Price
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ItemPrice"))
                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    'Amount Group
                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                                            
                        'Amount 'CSCLP-669
                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
                        'objChildElement4.Text = GenerateData(G_rstDetails, "O5")
                        dblInvoiceAmt = Val(IIf(GenerateData(G_rstDetails, "O5") = "", 0, GenerateData(G_rstDetails, "O5"))) 'revert
                        dblAdaptCharge = Val(IIf(GenerateData(G_rstDetails, "O7") = "", 0, GenerateData(G_rstDetails, "O7")))
                        objChildElement4.Text = Replace(CStr(dblInvoiceAmt + dblAdaptCharge), ",", ".")
                        
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                        
                        'Currency
                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
                        objChildElement4.Text = GenerateData(G_rstDetails, "O6")
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                        'Exchange Rate
                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
                        objChildElement4.Text = GenerateData(G_rstDetails, "OB")
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                            
                        'Estimated Flag
                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
                        objChildElement4.Text = "1"
                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
                    
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
'                'Indirect Payments
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("IndirectPayments"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Amount Group
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        'Amount
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        'Currency
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        'Exchange Rate
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        'Estimated Flag
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                
                'Total Amount 'CSCLP-669
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Amount"))
                'objChildElement2.Text = GenerateData(G_rstDetails, "O5") 'objChildElement2.Text = CStr(dblInvoiceAmt + dblAdaptCharge) 'revert
                dblInvoiceAmt = Val(IIf(GenerateData(G_rstDetails, "O5") = "", 0, GenerateData(G_rstDetails, "O5"))) 'revert
                dblAdaptCharge = Val(IIf(GenerateData(G_rstDetails, "O7") = "", 0, GenerateData(G_rstDetails, "O7")))
                objChildElement2.Text = Replace(CStr(dblInvoiceAmt + dblAdaptCharge), ",", ".")
                
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            '***********************************************************************************************
            
            '***********************************************************************************************
            'Additions +
            '***********************************************************************************************
            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Additions"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

'                'Costs +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Costs"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Commissions +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Commissions"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Brokerage +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Brokerage"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Containers Packing +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ContainersPacking"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'FreeOfCharge +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("FreeOfCharge"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Materials Components Parts +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("MaterialsComponentsParts"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Tools Dies Moulds +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ToolsDiesMoulds"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Consumed In Production +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ConsumedInProduction"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    'Engineering Design +
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("EngineeringDesign"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("AmountGroup"))
'                        objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Amount"))
'                            objChildElement5.Text = "0.00"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("Currency"))
'                            objChildElement5.Text = "EUR"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("ExchangeRate"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                            Set objChildElement5 = objChildElement4.appendChild(objDOM.createElement("EstimatedFlag"))
'                            objChildElement5.Text = "1"
'                            objChildElement4.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Royalties Licenses +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("RoyaltiesLicenses"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Proceeds +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Proceeds"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
                'Transport Handling Insurance +
                'If G_BlnDoNotIncludeIfEmpty = False Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TransportHandlingInsurance"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Place"))
                        objChildElement3.Text = GenerateData(G_rstHeader, "C3")
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Transport"))
                        objChildElement3.Text = "0.00"
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Handling"))
                        objChildElement3.Text = "0.00"
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Insurance"))
                        objChildElement3.Text = "0.00"
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)

                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("EstimatedFlag"))
                        objChildElement3.Text = "0"
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)

                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                'End If
'
                'Total Amount
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Amount"))
                objChildElement2.Text = GenerateData(G_rstDetails, "O2")
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'            '***********************************************************************************************
            
            '***********************************************************************************************
            'Deductions +
            '***********************************************************************************************
'            Set objChildElement = objChildNode2.appendChild(objDOM.createElement("Deductions"))
'            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'EUInland +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EUInland"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Construction Charges +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConstructionCharges"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Other Charges +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OtherCharges"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("Type"))
'                    objChildElement3.Text = ""
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Duties Taxes +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DutiesTaxes"))
'                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("AmountGroup"))
'                    objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Amount"))
'                        objChildElement4.Text = "0.00"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("Currency"))
'                        objChildElement4.Text = "EUR"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("ExchangeRate"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                        Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("EstimatedFlag"))
'                        objChildElement4.Text = "1"
'                        objChildElement3.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab)
'
'                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
'
'                'Amount +
'                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Amount"))
'                objChildElement2.Text = "0.00"
'                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
'
'            objChildNode2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
        Set objChildElement3 = Nothing
        Set objChildElement4 = Nothing
        Set objChildElement5 = Nothing
        
        G_rstDetails.MoveNext
    Next
    
End Sub


