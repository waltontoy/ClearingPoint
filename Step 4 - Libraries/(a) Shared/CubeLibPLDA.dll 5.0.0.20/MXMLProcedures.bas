Attribute VB_Name = "MXMLProcedures"
Option Explicit

Public g_rstXmlHeader As ADODB.Recordset
Public g_rstXmlHeaderHandelaars As ADODB.Recordset
Public g_rstXmlHeaderZegels As ADODB.Recordset
'''''    Public g_rstXmlHeaderTransitOffices As ADODB.Recordset
'''''    Public g_rstXmlHeaderZekerheid As ADODB.Recordset
'''''
Public g_rstXmlDetails As ADODB.Recordset
Public g_rstXmlDetailsHandelaars As ADODB.Recordset
Public g_rstXmlDetailsBijzondere As ADODB.Recordset
Public g_rstXmlDetailsContainer As ADODB.Recordset
Public g_rstXmlDetailsDocumenten As ADODB.Recordset
Public g_rstXmlDetailsZelf As ADODB.Recordset
'''''    Public g_rstXmlDetailsSensitiveGoods As ADODB.Recordset
'''''
Public g_rstXmlDetailsBerekeningsEenheden As ADODB.Recordset

Public g_lngTotalItemsF As Long
Public g_strXmlSender As String
Public g_strXmlRecipient As String

Public Const G_BlnDoNotIncludeIfEmpty = True
    
Public Function CreateXML(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, Optional ByVal CancellationReason As String) As String

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

    Select Case PLDABeXmlMsgType
        Case PLDABeXmlMsgTypeConstant.[PLDA Import AC4]

            Set objParentNode = objDOM.appendChild(objDOM.createElement("SADAC4"))
            Call CreateSadAC4XML(PLDABeXmlMsgType, objDOM, objParentNode, objChildNode)
        
        Case PLDABeXmlMsgTypeConstant.[PLDA Import AC4 Cancellation]

            Set objParentNode = objDOM.appendChild(objDOM.createElement("SADCancellation"))
            Call CreateSadAC4CancellationXML(PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, CancellationReason)
            
    End Select

    CreateXML = objDOM.xml

    ''*********************************************************************************
    ''FOR TESTING ONLY
    ''*********************************************************************************
    Dim strTemp As String
    Dim lngFreeFile As Long

    strTemp = objDOM.xml

    lngFreeFile = FreeFile()

    If Len(Dir(App.Path & "/test_AC4.xml")) <> 0 Then
        Kill App.Path & "/test_AC4.xml"
    End If

    Open App.Path & "/test_AC4.xml" For Output As #lngFreeFile
    Print #lngFreeFile, strTemp
    Close #lngFreeFile
    ''*********************************************************************************

    Set objChildNode = Nothing
    Set objParentNode = Nothing
    Set objDOM = Nothing
End Function

'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <ChildNode2>
'                <ChildElement>
'                   <ChildElement2>
'                       <ChildElement3>

Public Sub CreateSadAC4XML(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                           ByRef objDOM As DOMDocument, _
                           ByRef objParentNode As IXMLDOMNode, _
                           ByRef objChildNode As IXMLDOMNode)

    Dim objChildNode2 As IXMLDOMNode
    Dim objChildElement As IXMLDOMElement
    Dim objChildElement2 As IXMLDOMElement
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement


    ' Message Sender
    
    If LenB(Trim$(g_strXmlSender)) > 0 Or G_BlnDoNotIncludeIfEmpty = False Then
        CreateSadAC4MessageSender PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
    End If

    ' Function Code
    Set objChildNode = objDOM.createElement("functionCode")
    objChildNode.Text = 9
    objDOM.documentElement.appendChild objChildNode

    ' Language Code
    Set objChildNode = objDOM.createElement("languageCode")
    objChildNode.Text = GenerateData(g_rstXmlHeader, "A9")
    objDOM.documentElement.appendChild objChildNode

    ' Goods Declaration
    CreateSadAC4GoodsDeclaration PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
    
    ' Goods Item
    CreateSadAC4GoodsItem PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
End Sub

Private Function FormatDateToAC4Date(ByVal CPPLDABoxDate As String) As String
    Dim strCPPLDABoxDate As String
    
    strCPPLDABoxDate = CPPLDABoxDate
    
    strCPPLDABoxDate = Left$(strCPPLDABoxDate, 4) & "-" & Mid(strCPPLDABoxDate, 5, 2) & "-" & Right$(strCPPLDABoxDate, 2)
    
    FormatDateToAC4Date = strCPPLDABoxDate
End Function

Private Sub CreateSadAC4GoodsDeclaration(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement)
                                                  
    
    Dim strTotalGrossMass As String
    
    ' Goods Declaration
    Set objChildNode = objDOM.createElement("GoodsDeclaration")
    objDOM.documentElement.appendChild objChildNode
    
        ' Loading List
        If (GenerateData(g_rstXmlHeader, "A8") <> "" And GenerateData(g_rstXmlHeader, "A8") <> "0") Or _
           G_BlnDoNotIncludeIfEmpty = False Or (GenerateData(g_rstXmlHeader, "A1") = "AC" And GenerateData(g_rstXmlHeader, "A2") = "4") Then
           
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("loadingList"))
            objChildElement.Text = GenerateData(g_rstXmlHeader, "A8")
        End If
        
        ' Acceptance Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("acceptanceDate"))
        If GenerateData(g_rstXmlHeader, "A4") <> "" Then
            objChildElement.Text = FormatDateToAC4Date(GenerateData(g_rstXmlHeader, "A4"))
        Else
            objChildElement.Text = ""
        End If
        
        ' Local Reference Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("localReferenceNumber"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A3")
        
        ' Commercial Reference
        If GenerateData(g_rstXmlHeader, "AC") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
        
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("commercialReference"))
            objChildElement.Text = GenerateData(g_rstXmlHeader, "AC")
        End If
        
        ' Totals
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Totals"))

            'Items
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("items"))
            objChildElement2.Text = g_lngTotalItemsF
            
            'Total Gross Mass
            strTotalGrossMass = GenerateData(g_rstXmlHeader, "D1")
            strTotalGrossMass = Trim$(strTotalGrossMass)
            If IsNumeric(strTotalGrossMass) Then
                If Val(strTotalGrossMass) = 0 Then
                    strTotalGrossMass = ""
                End If
            End If
            If LenB(Trim$(strTotalGrossMass)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalGrossmass"))
                objChildElement2.Text = strTotalGrossMass
            End If
            
            'Total Net Mass
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("totalNetmass"))
            objChildElement2.Text = GenerateData(g_rstXmlHeader, "D2")
            
            'Packages
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
            objChildElement2.Text = GenerateData(g_rstXmlHeader, "D3")
        
        ' Transaction Nature
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TransactionNature"))

            'Transaction Nature 1
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature1"))
            objChildElement2.Text = Left$(GenerateData(g_rstXmlHeader, "C7"), 1) ' First character of C7
            
            'Transaction Nature 2
            If Len(Trim$(GenerateData(g_rstXmlHeader, "C7"))) = 2 Or G_BlnDoNotIncludeIfEmpty = False Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("transactionNature2"))
                
                If Len(Trim$(GenerateData(g_rstXmlHeader, "C7"))) = 2 Then
                    objChildElement2.Text = Trim$(Right$(GenerateData(g_rstXmlHeader, "C7"), 1)) ' Second character of C7
                Else
                    objChildElement2.Text = vbNullString
                End If
            End If
                
        ' Declarant
        CreateSadAC4Declarant PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
        
        ' Registration Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("registrationNumber"))
        objChildElement.Text = GenerateData(g_rstXmlHeaderHandelaars, "XD", 2)
        
        ' Issue Place
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("issuePlace"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A5")
        
        ' Signature
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("signature"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "AA")
        
        ' Type Part One
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartOne"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A1")
        
        ' Type Part Two
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartTwo"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A2")
        
        ' Consignee
        CreateSadAC4Consignee PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
        
        ' Payment Taxes
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PaymentTaxes"))

            'Payment Method Taxes
            If GenerateData(g_rstXmlHeader, "B1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("paymentMethodTaxes"))
                objChildElement2.Text = GenerateData(g_rstXmlHeader, "B1")
            End If
            
            'Deferred Payment
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("deferredPayment"))
            objChildElement2.Text = GenerateData(g_rstXmlHeader, "B5")
            
            'Deferred Payment Account Holder
            If GenerateData(g_rstXmlHeader, "B4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
            
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("deferredPaymentAccountHolder"))
                objChildElement2.Text = GenerateData(g_rstXmlHeader, "B4")
            End If
    
        ' Payment Vat
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PaymentVat"))

            'Payment Vat
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("paymentVat"))
            Select Case UCase$(GenerateData(g_rstXmlHeader, "B6"))
                Case "TRUE", "FALSE"
                    objChildElement2.Text = GenerateData(g_rstXmlHeader, "B6")
                Case Else
                    If GenerateData(g_rstXmlHeader, "B6") <> "" Then
                        objChildElement2.Text = "true"
                    Else
                        objChildElement2.Text = "false"
                    End If
            End Select
            'objChildElement2.Text = GenerateData(g_rstXmlHeader, "B6")
            
            'Payment Method Vat
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("paymentMethodVat"))
            objChildElement2.Text = GenerateData(g_rstXmlHeader, "B2")
                
        ' Customs
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Customs"))

            'Validation Office
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("validationOffice"))
            objChildElement2.Text = GenerateData(g_rstXmlHeader, "A6")
        
        ' Period
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Period"))
            
            'Begin Date Period
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("beginDatePeriod"))
            
            If GenerateData(g_rstXmlHeader, "H1") <> "" Then
                objChildElement2.Text = FormatDateToAC4Date(GenerateData(g_rstXmlHeader, "H1"))
            Else
                objChildElement2.Text = ""
            End If
            
            'End Date Period
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("endDatePeriod"))
            
            If GenerateData(g_rstXmlHeader, "H2") <> "" Then
                objChildElement2.Text = FormatDateToAC4Date(GenerateData(g_rstXmlHeader, "H2"))
            Else
                objChildElement2.Text = ""
            End If
            
            'Authorization Month Declaration
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorizationMonthDeclaration"))
            Select Case UCase$(GenerateData(g_rstXmlHeader, "H3"))
                Case "TRUE", "FALSE"
                    objChildElement2.Text = GenerateData(g_rstXmlHeader, "H3")
                Case Else
                    If GenerateData(g_rstXmlHeader, "H3") <> "" Then
                        objChildElement2.Text = "true"
                    Else
                        objChildElement2.Text = "false"
                    End If
            End Select
            
            'objChildElement2.Text = GenerateData(g_rstXmlHeader, "H3")
                
End Sub

Private Sub CreateSadAC4Declarant(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement)
                                          
    
    ' MUCP-48 - Start
    Dim strOperatorIdentity As String
    Dim strOperatorIdentifier As String
    ' MUCP-48 - End
    
    ' Declarant
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Declarant"))

        'Declarant Status
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("declarantstatus"))
        objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "XF", 2)
        
        'Authorised Identity
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("authorisedIdentity"))
        objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "XH", 2)
        
        ' Operator Identity
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))

            'Country
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("country"))
            objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
            
            'Identifier
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("identifier"))
            
            ' MUCP-48 - Start
            strOperatorIdentifier = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentifier = Trim$(strOperatorIdentifier)
            strOperatorIdentifier = Left$(strOperatorIdentifier, 5)
            If Len(strOperatorIdentifier) > 2 Then
                strOperatorIdentifier = Mid(strOperatorIdentifier, 3)
            End If
            objChildElement3.Text = strOperatorIdentifier
            ' MUCP-48 - End
            
            'Operator Identity
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
            
            ' MUCP-48 - Start
            strOperatorIdentity = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentity = Trim$(strOperatorIdentity)
            If Len(strOperatorIdentity) > 5 Then
                strOperatorIdentity = Mid(strOperatorIdentity, 6)
            End If
            objChildElement3.Text = strOperatorIdentity
            ' MUCP-48 - End
            
        ' Operator
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))

            ' Operator Name
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
            objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X2", 2)

            ' Operator Address
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
            
                ' Postal Code
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X5", 2)
                
                ' Street And Address 1
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X3", 2)
                
                ' Street And Address 2
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X4", 2)
                
                ' City
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X6", 2)
                
                ' Country Sub-Entity
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("countrySubEntity"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X7", 2)
                
                ' Country
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
        
            ' Contact Person
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ContactPerson"))
            
                ' Contact Person Name
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonName"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X9", 2)
                
                ' Contact Person Communication Number
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonCommunicationNumber"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XA", 2)
                
                ' Contact Person Email
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonEmail"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XC", 2)
                
                ' Contact Person Fax Number
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonFaxNumber"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XB", 2)
End Sub

Private Sub CreateSadAC4Consignee(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement)
                                          
    ' MUCP-48 - Start
    Dim strOperatorIdentity As String
    Dim strOperatorIdentifier As String
    ' MUCP-48 - End
            
    ' Consignee
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("Consignee"))

        'License Status
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("consigneeLicenseStatute"))
        objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "XF", 1)
        
        'License Number
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("consigneeLicenseNumber"))
        objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "XH", 1)
        
        ' Operator Identity
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorIdentity"))

            'Country
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("country"))
            objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 1)
            
            'Identifier
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("identifier"))
            
            ' MUCP-48 - Start
            strOperatorIdentifier = GenerateData(g_rstXmlHeaderHandelaars, "X1", 1)
            strOperatorIdentifier = Trim$(strOperatorIdentifier)
            strOperatorIdentifier = Left$(strOperatorIdentifier, 5)
            If Len(strOperatorIdentifier) > 2 Then
                strOperatorIdentifier = Mid(strOperatorIdentifier, 3)
            End If
            objChildElement3.Text = strOperatorIdentifier
            ' MUCP-48 - End
            
            'Operator Identity
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorIdentity"))
            
            ' MUCP-48 - Start
            strOperatorIdentity = GenerateData(g_rstXmlHeaderHandelaars, "X1", 1)
            strOperatorIdentity = Trim$(strOperatorIdentity)
            If Len(strOperatorIdentity) > 5 Then
                strOperatorIdentity = Mid(strOperatorIdentity, 6)
            End If
            objChildElement3.Text = strOperatorIdentity
            ' MUCP-48 - End
    
        ' Operator
        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Operator"))

            ' Operator Name
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("operatorName"))
            objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X2", 1)

            ' Operator Address
            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("OperatorAddress"))
            
                ' Postal Code
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("postalCode"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X5", 1)
                
                ' Street And Address 1
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber1"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X3", 1)
                
                ' Street And Address 2
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("streetAndNumber2"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X4", 1)
                
                ' City
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("city"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X6", 1)
                
                ' Country Sub-Entity
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("countrySubEntity"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X7", 1)
                
                ' Country
                Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("country"))
                objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 1)
        
            If GenerateData(g_rstXmlHeaderHandelaars, "X9", 1) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XA", 1) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XC", 1) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XB", 1) <> "" Or _
                G_BlnDoNotIncludeIfEmpty = False Then
            
                ' Contact Person
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ContactPerson"))
                
                    ' Contact Person Name
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonName"))
                    objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "X9", 1)
                    
                    ' Contact Person Communication Number
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonCommunicationNumber"))
                    objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XA", 1)
                    
                    ' Contact Person Email
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonEmail"))
                    objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XC", 1)
                    
                    ' Contact Person Fax Number
                    Set objChildElement4 = objChildElement3.appendChild(objDOM.createElement("contactPersonFaxNumber"))
                    objChildElement4.Text = GenerateData(g_rstXmlHeaderHandelaars, "XB", 1)
            End If
End Sub

Private Sub CreateSadAC4GoodsItem(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement)

    Dim lngDetailCounter As Long

    Dim lngBijzondere As Long
    Dim lngContainer As Long
    Dim lngDocumenten As Long
    Dim lngBerekenings As Long
    Dim lngZelf As Long

    Dim dblInvoiceAmt As Double
    Dim dblAdaptCharge As Double

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
    If g_rstXmlDetails.RecordCount > 0 Then
        g_rstXmlDetails.Sort = "DETAIL ASC"
        g_rstXmlDetails.MoveFirst

        For lngDetailCounter = 1 To g_rstXmlDetails.RecordCount

            If g_rstXmlDetailsBijzondere.RecordCount > 0 Then g_rstXmlDetailsBijzondere.MoveFirst
            If g_rstXmlDetailsContainer.RecordCount > 0 Then g_rstXmlDetailsContainer.MoveFirst
            If g_rstXmlDetailsDocumenten.RecordCount > 0 Then g_rstXmlDetailsDocumenten.MoveFirst

            lngBijzondere = 0
            lngContainer = 0
            lngDocumenten = 0

            'Goods Item
            Set objChildNode = objDOM.createElement("GoodsItem")
            objDOM.documentElement.appendChild objChildNode

                    'Sequence
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("sequence"))
                    objChildElement.Text = g_rstXmlDetails.Fields("Detail").Value

                    'National Additional Commodity 1
                    If GenerateData(g_rstXmlDetails, "L4") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("nationalAdditionalCommodity1"))
                        objChildElement.Text = GenerateData(g_rstXmlDetails, "L4")
                    End If

                    'National Additional Commodity 2
                    If GenerateData(g_rstXmlDetails, "L5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("nationalAdditionalCommodity2"))
                        objChildElement.Text = GenerateData(g_rstXmlDetails, "L5")
                    End If

                    'National Additional Commodity 3
                    If GenerateData(g_rstXmlDetails, "L6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("nationalAdditionalCommodity3"))
                        objChildElement.Text = GenerateData(g_rstXmlDetails, "L6")
                    End If

                    'Net Mass
                    If GenerateData(g_rstXmlDetails, "LA") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("netMass"))

                        strCheckMass = GenerateData(g_rstXmlDetails, "LA")
                        If Val(strCheckMass) = 0 Then
                            strCheckMass = ""
                        End If
                        objChildElement.Text = strCheckMass

                    End If

                    'Goods Description
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("goodsDescription"))
                    objChildElement.Text = GenerateData(g_rstXmlDetails, "L8")

                    ' Packaging
                    If Len(Trim$(GenerateData(g_rstXmlDetails, "S3"))) > 0 Then

                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Packaging"))

                            'Marks Number
                            If GenerateData(g_rstXmlDetails, "S3") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("marksNumber"))
                                objChildElement2.Text = Mid(GenerateData(g_rstXmlDetails, "S3"), 1, 35)
                            End If

                            'Packages
                            If GenerateData(g_rstXmlDetails, "S2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packages"))
                                objChildElement2.Text = GenerateData(g_rstXmlDetails, "S2")
                            End If

                            'Package Type
                            If GenerateData(g_rstXmlDetails, "S1") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("packageType"))
                                objChildElement2.Text = GenerateData(g_rstXmlDetails, "S1")
                            End If
                    End If

                    'Price

                    If (Not (Trim$(GenerateData(g_rstXmlDetails, "O5")) = "" Or Trim$(GenerateData(g_rstXmlDetails, "O5")) = "0")) Or _
                        (Not (Trim$(GenerateData(g_rstXmlDetails, "O7")) = "" Or Trim$(GenerateData(g_rstXmlDetails, "O7")) = "0")) Or _
                        G_BlnDoNotIncludeIfEmpty = False Then

                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Price"))

                            ' Price
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("price"))

                                dblInvoiceAmt = Val(IIf(GenerateData(g_rstXmlDetails, "O5") = "", 0, GenerateData(g_rstXmlDetails, "O5")))
                                dblAdaptCharge = Val(IIf(GenerateData(g_rstXmlDetails, "O7") = "", 0, GenerateData(g_rstXmlDetails, "O7")))

                                objChildElement2.Text = Replace(CStr(dblInvoiceAmt + dblAdaptCharge), ",", ".")

                            ' Exchange Rate
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ExchangeRate"))

                                'Exchange Rate
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("exchangeRate"))
                                objChildElement3.Text = GenerateData(g_rstXmlDetails, "OB")

                                'Currency
                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("currency"))
                                objChildElement3.Text = GenerateData(g_rstXmlDetails, "O6")
                    End If

                    ' Customs Treatment
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("CustomsTreatment"))

                        'Procedure
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Procedure"))

                            'procedurePart1
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedurePart1"))
                            objChildElement3.Text = GenerateData(g_rstXmlDetails, "N1")

                            'procedurePart2
                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("procedurePart2"))
                            objChildElement3.Text = GenerateData(g_rstXmlDetails, "N2")


                        '---------------------------------------------------------------------------------------------------------------------------
                        ' CALCULATION UNITS
                        '---------------------------------------------------------------------------------------------------------------------------
                        g_rstXmlDetailsBerekeningsEenheden.Filter = adFilterNone
                        g_rstXmlDetailsBerekeningsEenheden.Filter = "Detail = " & lngDetailCounter

                        If g_rstXmlDetailsBerekeningsEenheden.RecordCount > 0 Then
                            g_rstXmlDetailsBerekeningsEenheden.Sort = "ORDINAL ASC"
                            g_rstXmlDetailsBerekeningsEenheden.MoveFirst

                            lngBerekenings = 0

                            Do Until g_rstXmlDetailsBerekeningsEenheden.EOF
                                lngBerekenings = lngBerekenings + 1

                                If GenerateData(g_rstXmlDetailsBerekeningsEenheden, "T8", lngBerekenings) <> "" Or _
                                   GenerateData(g_rstXmlDetailsBerekeningsEenheden, "TZ", lngBerekenings) <> "" Or _
                                   G_BlnDoNotIncludeIfEmpty = False Then

                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CalculationUnits"))

                                        'Calculation Units Code
                                        If GenerateData(g_rstXmlDetailsBerekeningsEenheden, "TZ", lngBerekenings) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("calculationCode"))
                                            objChildElement3.Text = GenerateData(g_rstXmlDetailsBerekeningsEenheden, "TZ", lngBerekenings)
                                        End If

                                        'Claculation Units
                                        If GenerateData(g_rstXmlDetailsBerekeningsEenheden, "T8", lngBerekenings) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("calculationUnits"))
                                            objChildElement3.Text = GenerateData(g_rstXmlDetailsBerekeningsEenheden, "T8", lngBerekenings)

                                        End If
                                End If

                                If Trim$(FNullField(g_rstXmlDetailsBerekeningsEenheden.Fields("T9").Value)) = "E" Then Exit Do

                                g_rstXmlDetailsBerekeningsEenheden.MoveNext
                            Loop
                        End If
                        '---------------------------------------------------------------------------------------------------------------------------
                        '---------------------------------------------------------------------------------------------------------------------------

                    '---------------------------------------------------------------------------------------------------------------------------
                    'SELF ADDED DUTY
                    '---------------------------------------------------------------------------------------------------------------------------
                    If G_BlnDoNotIncludeIfEmpty = False Then

                        g_rstXmlDetailsZelf.Filter = adFilterNone
                        g_rstXmlDetailsZelf.Filter = "Detail = " & lngDetailCounter

                        If g_rstXmlDetailsZelf.RecordCount > 0 Then
                            g_rstXmlDetailsZelf.Sort = "ORDINAL ASC"
                            g_rstXmlDetailsZelf.MoveFirst

                            lngZelf = 0

                            'Tax Calculation
                            Do Until g_rstXmlDetailsZelf.EOF
                                lngZelf = lngZelf + 1

                                If GenerateData(g_rstXmlDetailsZelf, "U1", lngZelf) <> "" Or _
                                   GenerateData(g_rstXmlDetailsZelf, "U2", lngZelf) <> "" Or _
                                   GenerateData(g_rstXmlDetailsZelf, "U4", lngZelf) <> "" Or _
                                   G_BlnDoNotIncludeIfEmpty = False Then


                                    ' Self Added Duty
                                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("SelfAddedDuty"))

                                        ' Self Added Duty Type
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("selfAddedDutyType"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsZelf, "U1", lngZelf)

                                        ' Self Added Duty Code
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("selfAddedDutyCode"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsZelf, "U4", lngZelf)

                                        ' Self Added Duty Amount
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("selfAddedDutyAmount"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsZelf, "U2", lngZelf)

                                End If

                                If Trim$(FNullField(g_rstXmlDetailsBerekeningsEenheden.Fields("U3").Value)) = "E" Then Exit Do

                                g_rstXmlDetailsZelf.MoveNext
                            Loop

                        End If

                    End If
                    '---------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------


                    '---------------------------------------------------------------------------------------------------------------------------
                    'Additional Information +
                    '---------------------------------------------------------------------------------------------------------------------------
                    g_rstXmlDetailsBijzondere.Filter = adFilterNone
                    g_rstXmlDetailsBijzondere.Filter = "Detail = " & lngDetailCounter

                    If g_rstXmlDetailsBijzondere.RecordCount > 0 Then
                        g_rstXmlDetailsBijzondere.Sort = "ORDINAL ASC"
                        g_rstXmlDetailsBijzondere.MoveFirst

                        Do Until g_rstXmlDetailsBijzondere.EOF
                            lngBijzondere = lngBijzondere + 1

                            If GenerateData(g_rstXmlDetailsBijzondere, "P2", lngBijzondere) <> "" Or _
                               GenerateData(g_rstXmlDetailsBijzondere, "P1", lngBijzondere) <> "" Or _
                               GenerateData(g_rstXmlDetailsBijzondere, "P3", lngBijzondere) <> "" Or _
                               GenerateData(g_rstXmlDetailsBijzondere, "P4", lngBijzondere) <> "" Or _
                               G_BlnDoNotIncludeIfEmpty = False Then

                                Set objChildElement = objChildNode.appendChild(objDOM.createElement("AdditionalInformation"))

                                    'Additional Information Content
                                    If GenerateData(g_rstXmlDetailsBijzondere, "P2", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContent"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsBijzondere, "P2", lngBijzondere)
                                    End If


                                    'Additional Information Content Energy
                                    If GenerateData(g_rstXmlDetailsBijzondere, "P3", lngBijzondere) <> "" Or _
                                        GenerateData(g_rstXmlDetailsBijzondere, "P4", lngBijzondere) <> "" Or _
                                        G_BlnDoNotIncludeIfEmpty = False Then
                                    
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationContentEnergy"))
                                             
                                            'Additional Information License Number
                                            If GenerateData(g_rstXmlDetailsBijzondere, "P3", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("additionalInformationLicenseNumber"))
                                                objChildElement3.Text = GenerateData(g_rstXmlDetailsBijzondere, "P3", lngBijzondere)
                                            End If
                                    
                                            'Additional Information Product Code
                                            If GenerateData(g_rstXmlDetailsBijzondere, "P4", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("additionalInformationProductCode"))
                                                objChildElement3.Text = GenerateData(g_rstXmlDetailsBijzondere, "P4", lngBijzondere)
                                            End If
                                    End If

                                    'Additional Information Type
                                    If GenerateData(g_rstXmlDetailsBijzondere, "P1", lngBijzondere) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("additionalInformationType"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsBijzondere, "P1", lngBijzondere)
                                    End If

                            End If

                            If Trim$(FNullField(g_rstXmlDetailsBijzondere.Fields("P5").Value)) = "E" Then Exit Do

                            g_rstXmlDetailsBijzondere.MoveNext
                        Loop
                    End If
                    '---------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------

                    '---------------------------------------------------------------------------------------------------------------------------
                    'Previous Document +
                    '---------------------------------------------------------------------------------------------------------------------------
                    If GenerateData(g_rstXmlDetails, "R5") <> "" Or _
                       GenerateData(g_rstXmlDetails, "R2") <> "" Or _
                       GenerateData(g_rstXmlDetails, "R1") <> "" Or _
                       G_BlnDoNotIncludeIfEmpty = False Then

                        ' Previous Document
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PreviousDocument"))

                        'Document +
                        If GenerateData(g_rstXmlDetails, "R5") <> "" Or _
                           GenerateData(g_rstXmlDetails, "R2") <> "" Or _
                           G_BlnDoNotIncludeIfEmpty = False Then


                            ' Document
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Document"))

                                'Document Reference
                                If GenerateData(g_rstXmlDetails, "R5") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentReference"))

                                    'Select Case GenerateData(g_rstXmlDetails, "R2")
                                    '    Case "705"
                                    '        objChildElement3.Text = Left$(GenerateData(g_rstXmlDetails, "R5"), 6)
                                    '    Case "740"
                                    '        objChildElement3.Text = Left$(GenerateData(g_rstXmlDetails, "R5"), 5)
                                    '    Case Else
                                            objChildElement3.Text = GenerateData(g_rstXmlDetails, "R5")
                                    'End Select
                                End If

                                'Previous Document Type
                                If GenerateData(g_rstXmlDetails, "R2") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then

                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentType"))
                                    objChildElement3.Text = GenerateData(g_rstXmlDetails, "R2")
                                End If

                            ' Previous Document Date
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("previousDocumentDate"))
                            If GenerateData(g_rstXmlDetails, "R3") <> "" Then
                                objChildElement2.Text = FormatDateToAC4Date(GenerateData(g_rstXmlDetails, "R3"))
                            Else
                                objChildElement2.Text = ""
                            End If
                            'objChildElement2.Text = GenerateData(g_rstXmlDetails, "R3")

                            ' Previous Document Article
                            If GenerateData(g_rstXmlDetails, "R6") <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("previousDocumentArt"))
                                objChildElement2.Text = GenerateData(g_rstXmlDetails, "R6")
                            End If

                        End If
                    End If
                    '---------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------

                    '---------------------------------------------------------------------------------------------------------------------------
                    'Produced Document +
                    '---------------------------------------------------------------------------------------------------------------------------
                    g_rstXmlDetailsDocumenten.Filter = adFilterNone
                    g_rstXmlDetailsDocumenten.Filter = "Detail = " & lngDetailCounter

                    If g_rstXmlDetailsDocumenten.RecordCount > 0 Then
                        g_rstXmlDetailsDocumenten.Sort = "ORDINAL ASC"
                        g_rstXmlDetailsDocumenten.MoveFirst

                        Do Until g_rstXmlDetailsDocumenten.EOF
                            lngDocumenten = lngDocumenten + 1

                            '****************************************************************************************
                            'CSCLP-232 IF Q5 = 99 then it is an authorisation document and must be skipped
                            '****************************************************************************************
                            If GenerateData(g_rstXmlDetailsDocumenten, "Q5", lngDocumenten) <> "99" Then

                                Set objChildElement = objChildNode.appendChild(objDOM.createElement("ProducedDocument"))

                                    'Document +
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("Document"))

                                        'Document Reference
                                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentReference"))
                                        objChildElement3.Text = GenerateData(g_rstXmlDetailsDocumenten, "Q2", lngDocumenten)

                                        'Document Type
                                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("documentType"))
                                        objChildElement3.Text = GenerateData(g_rstXmlDetailsDocumenten, "Q1", lngDocumenten)

                                    'Document Date
                                    If GenerateData(g_rstXmlDetailsDocumenten, "Q3", lngDocumenten) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                    
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("producedDocumentsInformationDate"))
                                        If GenerateData(g_rstXmlDetailsDocumenten, "Q3", lngDocumenten) <> "" Then
                                            objChildElement2.Text = FormatDateToAC4Date(GenerateData(g_rstXmlDetailsDocumenten, "Q3", lngDocumenten))
                                        Else
                                            objChildElement2.Text = ""
                                        End If
                                        'objChildElement3.Text = GenerateData(g_rstXmlDetailsDocumenten, "Q3", lngDocumenten)
                                    End If

                                    'Complementary Information
                                    If GenerateData(g_rstXmlDetailsDocumenten, "Q9", lngDocumenten) <> "" Or G_BlnDoNotIncludeIfEmpty = False Then
                                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("complementaryInformation"))
                                        objChildElement2.Text = GenerateData(g_rstXmlDetailsDocumenten, "Q9", lngDocumenten)
                                    End If
                            End If
                            '****************************************************************************************

                            If Trim$(FNullField(g_rstXmlDetailsDocumenten.Fields("QA").Value)) = "E" Then Exit Do

                            g_rstXmlDetailsDocumenten.MoveNext
                        Loop
                    End If
                    '---------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------

            If Trim$(FNullField(g_rstXmlDetails.Fields("T7").Value)) = "F" Then
                Exit For
            End If

            g_rstXmlDetails.MoveNext

        Next
    End If
End Sub

Private Sub CreateSadAC4MessageSender(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                          ByRef objDOM As DOMDocument, _
                                          ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement)

    ' MUCP-48 - Start
    Dim strOperatorIdentity As String
    Dim strOperatorIdentifier As String
    Dim strCountry As String
    
    ' MUCP-48 - End
            
    '*****************
    ' Message Sender
    '*****************
    Set objChildNode = objDOM.createElement("MessageSender")
    objDOM.documentElement.appendChild objChildNode

        ' Operator Identity
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("OperatorIdentity"))

            'Country
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("country"))
            strCountry = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
            If Len(Trim$(strCountry)) <= 0 Then
                strCountry = "BE"
            End If
            objChildElement2.Text = strCountry
            
            'Identifier
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("identifier"))
            
            ' MUCP-48 - Start
            strOperatorIdentifier = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentifier = Trim$(strOperatorIdentifier)
            strOperatorIdentifier = Left$(strOperatorIdentifier, 5)
            If Len(strOperatorIdentifier) > 2 Then
                strOperatorIdentifier = Mid(strOperatorIdentifier, 3)
            End If
            
            ' MUCP-47 - Start
            objChildElement2.Text = strOperatorIdentifier
            ' MUCP-47 - End
            
            'objChildElement3.Text = strOperatorIdentifier
            ' MUCP-48 - End
            
            'Operator Identity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorIdentity"))
            
            ' MUCP-48 - Start
            strOperatorIdentity = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentity = Trim$(strOperatorIdentity)
            If Len(strOperatorIdentity) > 5 Then
                strOperatorIdentity = Mid(strOperatorIdentity, 6)
            End If
            objChildElement2.Text = strOperatorIdentity
            'objChildElement2.Text = g_strXmlSender
            ' MUCP-48 - End

        ' Operator
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Operator"))

            ' Operator Name
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorName"))
            ' MUCP-48 - Start
            objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "X2", 2)
            ' MUCP-48 - End

            ' Operator Address
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorAddress"))
            
                ' Postal Code
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("postalCode"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X5", 2)
                ' MUCP-48 - End
                
                ' Street And Number 1
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("streetAndNumber1"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X3", 2)
                ' MUCP-48 - End
                
                ' Street And Number 2
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("streetAndNumber2"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X4", 2)
                ' MUCP-48 - End
                
                ' City
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("city"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X6", 2)
                ' MUCP-48 - End
                
                ' Country Sub-Entity
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("countrySubEntity"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X7", 2)
                ' MUCP-48 - End
                
                ' Country
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("country"))
                ' MUCP-48 - Start
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
                'objChildElement3.Text = "BE"
                ' MUCP-48 - End
        
        ' Contact Person
        ' MUCP-48 - Start
        If GenerateData(g_rstXmlHeaderHandelaars, "X9", 2) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XA", 2) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XC", 2) <> "" Or _
                GenerateData(g_rstXmlHeaderHandelaars, "XB", 2) <> "" Or _
                G_BlnDoNotIncludeIfEmpty = False Then
                
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ContactPerson"))
            
                ' Contact Person Name
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonName"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X9", 2)
                'objChildElement3.Text = g_strXmlSender
                
                ' Contact Person Communication Number
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonCommunicationNumber"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "XA", 2)
                
                ' Contact Person Email
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonEmail"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "XC", 2)
                
                ' Contact Person Fax Number
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("contactPersonFaxNumber"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "XB", 2)
        End If
        ' MUCP-48 - End
        
End Sub


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


Public Sub CreateSadAC4CancellationXML(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                       ByRef objDOM As DOMDocument, _
                                       ByRef objParentNode As IXMLDOMNode, _
                                       ByRef objChildNode As IXMLDOMNode, _
                                       ByVal CancellationReason As String)

    Dim objChildNode2 As IXMLDOMNode
    Dim objChildElement As IXMLDOMElement
    Dim objChildElement2 As IXMLDOMElement
    Dim objChildElement3 As IXMLDOMElement
    Dim objChildElement4 As IXMLDOMElement

    ' Message Sender
    If LenB(Trim$(g_strXmlSender)) > 0 Or G_BlnDoNotIncludeIfEmpty = False Then
        CreateSadAC4MessageSenderForCancellation PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4
    End If

    ' Function Code
    Set objChildNode = objDOM.createElement("functionCode")
    objChildNode.Text = 3
    objDOM.documentElement.appendChild objChildNode

    ' Goods Declaration
    CreateSadAC4GoodsDeclarationForCancel PLDABeXmlMsgType, objDOM, objParentNode, objChildNode, objChildElement, objChildElement2, objChildElement3, objChildElement4, CancellationReason
    
End Sub

Private Sub CreateSadAC4GoodsDeclarationForCancel(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                                  ByRef objDOM As DOMDocument, _
                                                    ByRef objParentNode As IXMLDOMNode, _
                                          ByRef objChildNode As IXMLDOMNode, _
                                          ByRef objChildElement As IXMLDOMElement, _
                                          ByRef objChildElement2 As IXMLDOMElement, _
                                          ByRef objChildElement3 As IXMLDOMElement, _
                                          ByRef objChildElement4 As IXMLDOMElement, _
                                          ByVal CancellationReason As String)
                                                  
    
    ' Goods Declaration
    Set objChildNode = objDOM.createElement("GoodsDeclaration")
    objDOM.documentElement.appendChild objChildNode
    
        ' Type Part One
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartOne"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A1")
        
        ' Type Part Two
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartTwo"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A2")
        
        ' Type Part Three???
        'Set objChildElement = objChildNode.appendChild(objDOM.createElement("typePartTwo"))
        'objChildElement.Text = GenerateData(g_rstXmlHeader, "A2")
        
        ' MRN
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("customsMainReferenceNumber"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "MRN")
        
        ' LRN
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("localReferenceNumber"))
        objChildElement.Text = GenerateData(g_rstXmlHeader, "A3")
        
        ' Cancellation Reason
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("reason"))
        objChildElement.Text = CancellationReason
        
                
End Sub


Private Sub CreateSadAC4MessageSenderForCancellation(ByVal PLDABeXmlMsgType As PLDABeXmlMsgTypeConstant, _
                                                     ByRef objDOM As DOMDocument, _
                                                     ByRef objParentNode As IXMLDOMNode, _
                                                     ByRef objChildNode As IXMLDOMNode, _
                                                     ByRef objChildElement As IXMLDOMElement, _
                                                     ByRef objChildElement2 As IXMLDOMElement, _
                                                     ByRef objChildElement3 As IXMLDOMElement, _
                                                     ByRef objChildElement4 As IXMLDOMElement)

    Dim strOperatorIdentity As String
    Dim strOperatorIdentifier As String
    Dim strCountry As String
            
    '*****************
    ' Message Sender
    '*****************
    Set objChildNode = objDOM.createElement("MessageSender")
    objDOM.documentElement.appendChild objChildNode

        ' Operator Identity
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("OperatorIdentity"))

            'Country
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("country"))
            strCountry = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
            If Len(Trim$(strCountry)) <= 0 Then
                strCountry = "BE"
            End If
            objChildElement2.Text = strCountry
            
            'Identifier
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("identifier"))
            
            strOperatorIdentifier = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentifier = Trim$(strOperatorIdentifier)
            strOperatorIdentifier = Left$(strOperatorIdentifier, 5)
            If Len(strOperatorIdentifier) > 2 Then
                strOperatorIdentifier = Mid(strOperatorIdentifier, 3)
            End If
            
            objChildElement2.Text = strOperatorIdentifier
            
            'Operator Identity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorIdentity"))
            
            strOperatorIdentity = GenerateData(g_rstXmlHeaderHandelaars, "X1", 2)
            strOperatorIdentity = Trim$(strOperatorIdentity)
            If Len(strOperatorIdentity) > 5 Then
                strOperatorIdentity = Mid(strOperatorIdentity, 6)
            End If
            objChildElement2.Text = strOperatorIdentity
            
        ' Operator
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("Operator"))

            ' Operator Name
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("operatorName"))
            objChildElement2.Text = GenerateData(g_rstXmlHeaderHandelaars, "X2", 2)
            
            ' Operator Address
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OperatorAddress"))
            
                ' Postal Code
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("postalCode"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X5", 2)
                
                ' Street And Number 1
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("streetAndNumber1"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X3", 2)
                
                ' Street And Number 2
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("streetAndNumber2"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X4", 2)
                
                ' City
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("city"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X6", 2)
                
                ' Country Sub-Entity
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("countrySubEntity"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X7", 2)
                
                ' Country
                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("country"))
                objChildElement3.Text = GenerateData(g_rstXmlHeaderHandelaars, "X8", 2)
        
End Sub

