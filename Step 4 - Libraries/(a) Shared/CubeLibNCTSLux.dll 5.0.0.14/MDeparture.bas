Attribute VB_Name = "MDeparture"
Option Explicit

Private lngStoragePackage As Long
Private lngStorageContainer As Long
Private lngStoragSpecialMention As Long
Private lngStorageProducedDocs As Long
Private lngStoragePreviuosDocs As Long
Private lngSensitiveGoods As Long

Private m_blnW6IsTheSame As Boolean
Private m_blnU6IsTheSame As Boolean
Private m_strA5 As String


'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <objChildElement>
'               <objChildElement2>
'                   <objChildElement3>

Public Sub CreateXMLMessageIE15(ByRef DataSourceProperties As CDataSourceProperties, _
                                ByRef objDOM As DOMDocument, _
                                ByRef objParentNode As IXMLDOMNode, _
                                ByRef objChildNode As IXMLDOMNode)
    
    Dim objChildElement As IXMLDOMNode
    Dim objChildElement2 As IXMLDOMNode
    
    InitializeRecorsetsForXML DataSourceProperties
    
    'Check Venture Number U6
    m_blnU6IsTheSame = VentureNumberAreTheSame("U6")
    
    'Check Venture Number W6
    m_blnW6IsTheSame = VentureNumberAreTheSame("W6")
    
    'Get Language
    m_strA5 = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
    
    'Interchange
    CreateMessageInterchangeIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Header
    CreateMessageHeaderIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Principal
    CreateMessagePrincipalIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Consignor
    If m_blnU6IsTheSame = True Then
        CreateMessageConsignorHeaderIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    End If
    
    'Consignee
    If m_blnW6IsTheSame = True Then
        CreateMessageConsigneeHeaderIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    End If
    'Next
    'Authorised
    CreateMessageAuthorisedIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Departure Customs Office
    CreateMessageDepartureOfficeIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Transit Customs Office
    CreateMessageTransitOfficesIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Destination Customs Office
    CreateMessageDestinationOfficeIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Control Result
    CreateMessageControlResultIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Representative
    CreateMessageRepresentativeIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Seals Info
    CreateMessageSealsInfoIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2 'allan ncts no entry
    
    'Guarantee
    CreateMessageGuaranteeIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Goods Item
    lngStoragePackage = 0
    lngStorageContainer = 0
    lngStoragSpecialMention = 0
    lngStorageProducedDocs = 0
    lngStoragePreviuosDocs = 0
    CreateMessageGoodsItemIE15 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    Set objChildElement = Nothing
    Set objChildElement2 = Nothing
     
End Sub

Private Sub CreateMessageInterchangeIE15(ByRef objDOM As DOMDocument, _
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
    'objChildNode.Text = GetMapFunctionValue("F<RECEIVE QUEUE>")
    objChildNode.Text = GetValueIfNotNull(rstUNB.Fields("DATA_NCTS_UNB_Seq3"))  'allan ncts
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Sender Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("SenIdeCodQuaMES4"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Message Recipient
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesRecMES6"))
    'objChildNode.Text = GetMapFunctionValue("F<RECIPIENT>")
    objChildNode.Text = GetValueIfNotNull(rstUNB.Fields("DATA_NCTS_UNB_Seq6")) 'allan ncts
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Recipient Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecIdeCodQuaMES7"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Date of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("DatOfPreMES9"))
    'objChildNode.Text = GetMapFunctionValue("F<DATE, YYMMDD>")
    objChildNode.Text = GetValueIfNotNull(rstUNB.Fields("DATA_NCTS_UNB_Seq9"))
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Time of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("TimOfPreMES10"))
    'objChildNode.Text = GetMapFunctionValue("F<TIME, HHMM>")
    objChildNode.Text = GetValueIfNotNull(rstUNB.Fields("DATA_NCTS_UNB_Seq10"))
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    'Interchange Control Reference
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("IntConRefMES11"))
    'objChildNode.Text = GetMapFunctionValue("F<1 TIN REF>")
    objChildNode.Text = GetValueIfNotNull(rstUNB.Fields("DATA_NCTS_UNB_Seq11"))
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
    objChildNode.Text = "CC015A"
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


Private Sub CreateMessageHeaderIE15(ByRef objDOM As DOMDocument, _
                                    ByRef objParentNode As IXMLDOMNode, _
                                    ByRef objChildNode As IXMLDOMNode, _
                                    ByRef objChildElement As IXMLDOMNode, _
                                    ByRef objChildElement2 As IXMLDOMNode)

    
    Dim strTemp As String
    
    'Header
    Set objChildNode = objDOM.createElement("HEAHEA")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Reference number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumHEA4"))
        objChildElement.Text = GetValueForSegment(rstRFF, "RFF", "17", 2) 'GetValueIfNotNull(rstRFF.Fields("DATA_NCTS_RFF_Seq2")) 'allan ncts
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Type of Declaration
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TypOfDecHEA24"))
        objChildElement.Text = GetValueIfNotNull(rstBGM.Fields("DATA_NCTS_BGM_Seq4")) 'allan ncts
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country of Destination Code
        strTemp = GetValueForSegment(rstLOC, "LOC", "4", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouOfDesCodHEA30"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Agreed Location of Goods Code
        'Set objChildElement = objChildNode.appendChild(objDOM.createElement("AgrLocOfGooCodHEA38"))
        'objChildElement.Text = vbNullString
        'objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Agreed Location of Goods
        'strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "AB", 1)(0)
        strTemp = GetValueForSegment(rstLOC, "LOC", "5", "2")
        'CSCLP-336 commented
        'If Len(Trim$(strTemp)) > 0 Then
        '    Set objChildElement = objChildNode.appendChild(objDOM.createElement("AgrLocOfGooHEA39"))
        '    objChildElement.Text = strTemp
        '    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        'End If
        
        'CSCLP-336 commented
        'Agreed Location of Goods Language
        'If GetSegmentOptionForLanguage("AgrLocOfGooHEA39LNG") = False Or Len(Trim$(strTemp)) > 0 Then
        '    Set objChildElement = objChildNode.appendChild(objDOM.createElement("AgrLocOfGooHEA39LNG"))
        '    objChildElement.Text = m_strA5
        '    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        'End If
        
        'CSCLP-336 uncommented and modified objChildElement.Text = vbnullstring
        'Authorised Location of Goods Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("AutLocOfGooCodHEA41"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Place of Loading Code
        strTemp = GetValueForSegment(rstLOC, "LOC", "6", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PlaOfLoaCodHEA46"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Country of Dispatch Code
        strTemp = GetValueForSegment(rstLOC, "LOC", "7", "2")
        If Len(Trim$(strTemp)) >= 2 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouOfDisCodHEA55"))
            'objChildElement.Text = Mid(strTemp, 1, 2)
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Customs Sub Place - 'CSCLP-168
        'Set objChildElement = objChildNode.appendChild(objDOM.createElement("CusSubPlaHEA66"))
        'objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A4", 1)(0)
        'objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Inland Transport Mode
        strTemp = GetValueForSegment(rstTDT, "TDT", "22", "3")
        If Len(Trim$(strTemp)) >= 2 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("InlTraModHEA75"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Transport Mode at Border
        strTemp = GetValueForSegment(rstTDT, "TDT", "23", "3")
        If Len(Trim$(strTemp)) >= 2 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TraModAtBorHEA76"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Identity of Means of Transport at Departure
        'strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B1", 1)(0)
        strTemp = GetValueForSegment(rstTDT, "TDT", "25", "18")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraAtDHEA78"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'seafood
        'Identity of Means of Transport at Departure Language
        If (GetSegmentOptionForLanguage("IdeOfMeaOfTraAtDHEA78LNG") = True) And (Len(Trim$(strTemp)) > 0) Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraAtDHEA78LNG"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Nationality of Means of Transport at Departure
        strTemp = GetValueForSegment(rstTDT, "TDT", "23", "19")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NatOfMeaOfTraAtDHEA80"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Identity of Means of Transport at Crossing Border
        strTemp = GetValueForSegment(rstTDT, "TDT", "23", "18")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraCroHEA85"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                        
        'seafood
        'Identity of Means of Transport at Crossing Border Language
        If (GetSegmentOptionForLanguage("IdeOfMeaOfTraCroHEA85LNG") = True) And (Len(Trim$(strTemp)) > 0) Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraCroHEA85LNG"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Nationality of Means of Transport at Crossing Border
        strTemp = GetValueForSegment(rstTDT, "TDT", "23", "19")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NatOfMeaOfTraCroHEA87"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Type of Means of Transport Crossing Border
        strTemp = GetValueForSegment(rstTDT, "TDT", "23", "5")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TypOfMeaOfTraCroHEA88"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Containerized Indicator
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("ConIndHEA96"))
        strTemp = IIf(Len(GetValueForSegment(rstGIS, "GIS", "10", "1")) > 0, GetValueForSegment(rstGIS, "GIS", "10", "1"), "0")
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Dialog Language Indicator at Departure
        If Len(Trim$(m_strA5)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("DiaLanIndAtDepHEA254"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'NCTS Accompanying Document Language Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NCTSAccDocHEA601LNG"))
        objChildElement.Text = m_strA5
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'seafood
        'Number of Loading Lists
        strTemp = IIf(Len(GetValueForSegment(rstCNT, "CNT", "51", "2")) > 0, GetValueForSegment(rstCNT, "CNT", "51", "2"), "0")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NumOfLoaLisHEA304"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Total Number of Items
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotNumOfIteHEA305"))
        strTemp = GetValueForSegment(rstCNT, "CNT", "53", "2")
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Total Number of Packages
        strTemp = GetValueForSegment(rstCNT, "CNT", "50", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotNumOfPacHEA306"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Total Gross Mass
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotGroMasHEA307"))
        strTemp = GetValueForSegment(rstMEA, "MEA", "11", "7")
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Declaration Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DecDatHEA383"))
        'objChildElement.Text = GetMapFunctionValue("F<DATE, YYYYMMDD>")
        objChildElement.Text = GetValueForSegment(rstDTM, "DTM", "8", "2")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Declaration Place
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DecPlaHEA394"))
        strTemp = GetValueForSegment(rstLOC, "LOC", "57", "5")
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Declaration Place Language
        If GetSegmentOptionForLanguage("DecPlaHEA394LNG") = True And Len(Trim$(m_strA5)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("DecPlaHEA394LNG"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub


Private Sub CreateMessagePrincipalIE15(ByRef objDOM As DOMDocument, _
                                       ByRef objParentNode As IXMLDOMNode, _
                                       ByRef objChildNode As IXMLDOMNode, _
                                       ByRef objChildElement As IXMLDOMNode, _
                                       ByRef objChildElement2 As IXMLDOMNode)
    
    
    Dim strTemp As String
    
    'Principal
    Set objChildNode = objDOM.createElement("TRAPRIPC1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "10")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamPC17"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Street and Number
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "16")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumPC122"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Postal Code
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "22")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodPC123"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'City
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "20")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitPC124"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Country Code
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "23")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouPC125"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'NAD Language
        If Len(Trim$(m_strA5)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGPC"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'TIN
        strTemp = GetValueForSegment(rstNAD, "NAD", "27", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINPC159"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Private Sub CreateMessageConsignorHeaderIE15(ByRef objDOM As DOMDocument, _
                                             ByRef objParentNode As IXMLDOMNode, _
                                             ByRef objChildNode As IXMLDOMNode, _
                                             ByRef objChildElement As IXMLDOMNode, _
                                             ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    'Consignor
    Set objChildNode = objDOM.createElement("TRACONCO1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamCO17"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "30", "10")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Street and Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumCO122"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "30", "16")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Postal Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodCO123"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "30", "22")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'City
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitCO124"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "30", "20")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouCO125"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "30", "23")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'NAD Language
        If Len(Trim$(m_strA5)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGCO"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'TIN
        strTemp = GetValueForSegment(rstNAD, "NAD", "30", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINCO159"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Private Sub CreateMessageConsigneeHeaderIE15(ByRef objDOM As DOMDocument, _
                                             ByRef objParentNode As IXMLDOMNode, _
                                             ByRef objChildNode As IXMLDOMNode, _
                                             ByRef objChildElement As IXMLDOMNode, _
                                             ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    'Consignee
    Set objChildNode = objDOM.createElement("TRACONCE1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamCE17"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "29", "10")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Street and Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumCE122"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "29", "16")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Postal Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodCE123"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "29", "22")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'City
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitCE124"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "29", "20")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouCE125"))
        objChildElement.Text = GetValueForSegment(rstNAD, "NAD", "29", "23")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'NAD Language
        If Len(Trim$(m_strA5)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGCE"))
            objChildElement.Text = m_strA5
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'TIN
        strTemp = GetValueForSegment(rstNAD, "NAD", "29", "2")
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINCE159"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Private Sub CreateMessageAuthorisedIE15(ByRef objDOM As DOMDocument, _
                                        ByRef objParentNode As IXMLDOMNode, _
                                        ByRef objChildNode As IXMLDOMNode, _
                                        ByRef objChildElement As IXMLDOMNode, _
                                        ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    Dim lngDetailCount As Long
    Dim lngctr As Long
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    
    '************************************************************************************************************
    'Rule for Authorised Consignee
    '************************************************************************************************************
    '       If W7 = Empty String or W7 <> "Y" then do not include "Authorised Consignee" segment
    '       If W6 = Empty String then do not include "Authorised Consignee" segment
    '************************************************************************************************************
    lngDetailCount = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "F<DETAIL COUNT>", 0)(0)
    
    If m_blnW6IsTheSame = True Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "DATA_NCTS_DETAIL "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "Detail = 1"
        ADORecordsetOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
        'RstOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
        
        If rstTemp.RecordCount > 0 Then
            strTemp = IIf(IsNull(rstTemp.Fields("W7").Value), "", Trim$(rstTemp.Fields("W7").Value))
        Else
            strTemp = vbNullString
        End If
        
        If Not IsNull(strTemp) Then
            If UCase$(Trim$(strTemp)) = "Y" Then
                
                strTemp = GetValueForSegment(rstNAD, "NAD", "39", "2", True, 1)
                If Len(Trim$(strTemp)) > 0 Then
                    'Authorised
                    Set objChildNode = objDOM.createElement("TRAAUTCONTRA")
                    objDOM.documentElement.appendChild objChildNode
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                        'TIN
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRA59"))
                        objChildElement.Text = strTemp
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
                End If
                
            End If
        End If
    
    Else
        For lngctr = 1 To lngDetailCount
            
                strSQL = vbNullString
                strSQL = strSQL & "SELECT * "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "DATA_NCTS_DETAIL "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Detail = " & lngctr
            ADORecordsetOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                strTemp = IIf(IsNull(rstTemp.Fields("W7").Value), "", Trim$(rstTemp.Fields("W7").Value))
            Else
                strTemp = vbNullString
            End If
            
            If Not IsNull(strTemp) Then
                If UCase$(Trim$(strTemp)) = "Y" Then
                    
                    strTemp = GetValueForSegment(rstNAD, "NAD", "39", "2", True, lngctr)
                    If Len(Trim$(strTemp)) > 0 Then
                        'Authorised
                        Set objChildNode = objDOM.createElement("TRAAUTCONTRA")
                        objDOM.documentElement.appendChild objChildNode
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                            'TIN
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRA59"))
                            objChildElement.Text = strTemp
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
                        
                        Exit For
                    End If
                    
                End If
            End If
            
            strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "T7", lngctr)(0)
            
            If UCase$(Trim$(strTemp)) = "F" Then Exit For
            
        Next
        
        
    End If
    
    ADORecordsetClose rstTemp
    'RstClose rstTemp
    
End Sub


Private Sub CreateMessageTransitOfficesIE15(ByRef objDOM As DOMDocument, _
                                            ByRef objParentNode As IXMLDOMNode, _
                                            ByRef objChildNode As IXMLDOMNode, _
                                            ByRef objChildElement As IXMLDOMNode, _
                                            ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    'Header behavior changed to repeat for each reference number entered
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 1)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 2)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 3)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
       
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "EC", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 4)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
       
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 5)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
       
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
    'Reference Number
    strTemp = GetValueForSegment(rstLOC, "LOC", "59", "2", True, 6)
    If (Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0") Then
        'Transit Customs Offices
        Set objChildNode = objDOM.createElement("CUSOFFTRARNS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
       
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRNS1"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If

End Sub



Private Sub CreateMessageDepartureOfficeIE15(ByRef objDOM As DOMDocument, _
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
        objChildElement.Text = GetValueForSegment(rstLOC, "LOC", "58", "2")
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

End Sub

Private Sub CreateMessageDestinationOfficeIE15(ByRef objDOM As DOMDocument, _
                                               ByRef objParentNode As IXMLDOMNode, _
                                               ByRef objChildNode As IXMLDOMNode, _
                                               ByRef objChildElement As IXMLDOMNode, _
                                               ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    strTemp = GetValueForSegment(rstLOC, "LOC", "60", "2")
    'If Len(Trim$(strTemp)) >= 10 Then 'allan ncts
    
        'Customs Offices of Departure
        Set objChildNode = objDOM.createElement("CUSOFFDESEST")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
            'Reference Number
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumEST1"))
            objChildElement.Text = strTemp 'allan ncts
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
    'End If 'allan ncts

End Sub


Private Sub CreateMessageControlResultIE15(ByRef objDOM As DOMDocument, _
                                           ByRef objParentNode As IXMLDOMNode, _
                                           ByRef objChildNode As IXMLDOMNode, _
                                           ByRef objChildElement As IXMLDOMNode, _
                                           ByRef objChildElement2 As IXMLDOMNode)
    
    If Len(Trim$(GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "AC", 1)(0))) > 0 And _
       Len(Trim$(GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "AD", 1)(0))) > 0 Then
    
        'Control Result
        Set objChildNode = objDOM.createElement("CONRESERS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
            'Control Result Code
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ConResCodERS16"))
            objChildElement.Text = GetValueForSegment(rstFTX, "FTX", "15", "3")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Date Limit
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("DatLimERS69"))
            objChildElement.Text = GetValueForSegment(rstDTM, "DTM", "9", "2")
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    End If

End Sub


Private Sub CreateMessageRepresentativeIE15(ByRef objDOM As DOMDocument, _
                                            ByRef objParentNode As IXMLDOMNode, _
                                            ByRef objChildNode As IXMLDOMNode, _
                                            ByRef objChildElement As IXMLDOMNode, _
                                            ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    strTemp = GetValueForSegment(rstNAD, "NAD", "28", "10")
    If Len(Trim$(strTemp)) > 0 Then
    
        'Representative
        Set objChildNode = objDOM.createElement("REPREP")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
            'Name
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamREP5"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Representative Capacity
            
            strTemp = GetValueForSegment(rstFTX, "FTX", "14", "6")
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("RepCapREP18"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
                        
            'Representative Capacity Language
            If GetSegmentOptionForLanguage("RepCapREP18LNG") = False Or _
               Len(Trim$(strTemp)) > 0 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("RepCapREP18LNG"))
                objChildElement.Text = m_strA5
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
    End If

End Sub


Private Sub CreateMessageSealsInfoIE15(ByRef objDOM As DOMDocument, _
                                       ByRef objParentNode As IXMLDOMNode, _
                                       ByRef objChildNode As IXMLDOMNode, _
                                       ByRef objChildElement As IXMLDOMNode, _
                                       ByRef objChildElement2 As IXMLDOMNode)
    
    
    Dim varSealsIdentity As Variant
    Dim lngSealCount As Long
    
    Dim lngctr As Long
    
    Call GetSealsNumberAndIdentity(lngSealCount, varSealsIdentity)
    
    If lngSealCount > 0 Then
        'Seals Information
        Set objChildNode = objDOM.createElement("SEAINFSLI")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Seals Number
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("SeaNumSLI2"))
            objChildElement.Text = lngSealCount
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
            If IsArray(varSealsIdentity) = True Then
                For lngctr = LBound(varSealsIdentity) To UBound(varSealsIdentity)
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("SEAIDSID"))
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        'Seals Identity
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1"))
                        objChildElement2.Text = varSealsIdentity(lngctr)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        'Seals Identity Language
                        If Len(Trim$(m_strA5)) > 0 Then
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1LNG"))
                            objChildElement2.Text = m_strA5
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        End If
                    
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                Next
            Else
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("SEAIDSID"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Seals Identity
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1"))
                    objChildElement2.Text = varSealsIdentity
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Seals Identity Language
                    If Len(Trim$(m_strA5)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1L"))
                        objChildElement2.Text = m_strA5
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
                
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    End If
    
End Sub


Private Sub CreateMessageGuaranteeIE15(ByRef objDOM As DOMDocument, _
                                       ByRef objParentNode As IXMLDOMNode, _
                                       ByRef objChildNode As IXMLDOMNode, _
                                       ByRef objChildElement As IXMLDOMNode, _
                                       ByRef objChildElement2 As IXMLDOMNode)
    
    
    Dim objChildElement3 As IXMLDOMElement
    Dim lngctr As Long
    
    Dim varE1 As Variant
    Dim varE3 As Variant
    Dim varE4 As Variant
    Dim varE5 As Variant
    Dim varE6 As Variant
    Dim varE7 As Variant
    Dim varEK As Variant
    Dim varEM As Variant
    Dim varEN As Variant
    Dim lngE1Count As Long
    Dim rstTemp As ADODB.Recordset
    Dim lngIncrementZekerheid As Long
    'Retrieve Box EO and EJ
    GetGroupRecordsFromDataNCTSTables "DATA_NCTS_HEADER_ZEKERHEID", rstTemp
 
    rstRFF.Filter = adFilterNone
    rstRFF.Filter = "[NCTS_IEM_TMS_ID] = 18"
    lngE1Count = rstRFF.RecordCount
    
    If (lngE1Count) > rstTemp.RecordCount Then Exit Sub 'allan ncts
    rstTemp.MoveFirst
    
    For lngctr = 0 To rstRFF.RecordCount - 1
        
        If lngctr = 0 Then
            lngIncrementZekerheid = 0
        Else
            lngIncrementZekerheid = lngIncrementZekerheid + 6
        End If
        
        'Guarantee
        Set objChildNode = objDOM.createElement("GUAGUA")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Guarantee Type
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GuaTypGUA1"))
            varE1 = GetValueForSegment(rstRFF, "RFF", "18", "2", True, lngctr + 1)
            objChildElement.Text = varE1
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GUAREFREF"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                If rstTemp.Fields("EJ").Value = "G" Then
                    'Guarantee Reference Number
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("GuaRefNumGRNREF1"))
                    varE3 = GetValueForSegment(rstPAC, "PAC", "19", "10", True, lngctr + 1)
                    objChildElement2.Text = varE3
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                Else
                    'Other Guarantee Reference Number
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("OthGuaRefREF4"))
                    varE3 = GetValueForSegment(rstPAC, "PAC", "19", "10", True, lngctr + 1)
                    objChildElement2.Text = varE3
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                    
                'Access Code
                varEK = GetValueForSegment(rstPAC, "PAC", "19", "9", True, lngctr + 1)
                If Len(Trim$(varEK)) >= 4 Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AccCodREF6"))
                    objChildElement2.Text = varEK
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                    
                'Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMECVLE"))
                'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        
                    'Not valid for EC
                '    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForECVL"))
                '    objChildElement3.Text = vbNullString
                '    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    
                'objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                'Not valid for other contracting parties
                varE4 = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (1 + lngIncrementZekerheid))
                If Len(Trim$(varE4)) >= 2 Then 'allan ncts
                   Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varE4
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Not valid for other contracting parties
                varE5 = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (2 + lngIncrementZekerheid))
                If Len(Trim$(varE5)) >= 2 Then 'allan ncts
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varE5
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Not valid for other contracting parties
                varE6 = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (3 + lngIncrementZekerheid))
                If Len(Trim$(varE6)) >= 2 Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varE6
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Not valid for other contracting parties
                varE7 = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (4 + lngIncrementZekerheid))
                If Len(Trim$(varE7)) >= 2 Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varE7
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Not valid for other contracting parties
                varEM = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (5 + lngIncrementZekerheid))
                If Len(Trim$(varEM)) >= 2 Then
                   Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varEM
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
                'Not valid for other contracting parties
                varEN = GetValueForSegment(rstPCI, "PCI", "21", "2", True, (6 + lngIncrementZekerheid))
                If Len(Trim$(varEN)) >= 2 Then
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("VALLIMNONECLIM"))
                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                        Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("NotValForOthConPLIM2"))
                        objChildElement3.Text = varEN
                        objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                End If
                
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
        If UCase$(Trim$(rstTemp.Fields("EO").Value)) = "E" Then Exit For
        
        rstTemp.MoveNext
    Next

End Sub



Private Sub CreateMessageGoodsItemIE15(ByRef objDOM As DOMDocument, _
                                       ByRef objParentNode As IXMLDOMNode, _
                                       ByRef objChildNode As IXMLDOMNode, _
                                       ByRef objChildElement As IXMLDOMNode, _
                                       ByRef objChildElement2 As IXMLDOMNode)
                                  
    Dim strTemp As String
    
    Dim lngDetailCount As Long
    Dim lngctr As Long
    
    Dim lngPrevDocCtr As Long
    Dim varDocType As Variant
    
    Dim lngAddInfo As Long
    Dim varAddInfoZ1 As Variant
    Dim varAddInfoZ2 As Variant
    Dim varAddInfoZ3 As Variant
    
    Dim lngContainerCtr As Long
    Dim varContNumS6 As Variant
    Dim varContNumS7 As Variant
    Dim varContNumS8 As Variant
    Dim varContNumS9 As Variant
    Dim varContNumSA As Variant
    
    Dim lngPackagesCtr As Long
    Dim varMarksAndNum As Variant
    Dim varKindPackage As Variant
    Dim varNumPackage As Variant
    
    Dim strV2 As String
    Dim strV4 As String
    Dim strV6 As String
    Dim strV8 As String
    
    Dim strV1 As String
    Dim strV3 As String
    Dim strV5 As String
    Dim strV7 As String

    Dim lngIncrementV As Long
    Dim lngIncrement As Long
    Dim strMarksAndNum As String
    Dim strKindPackage As String
    Dim strNumPackage As String
    Dim rstTemp As ADODB.Recordset
    
    lngDetailCount = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "F<DETAIL COUNT>", 0)(0)
    
    For lngctr = 1 To lngDetailCount
        'Goods Item
        Set objChildNode = objDOM.createElement("GOOITEGDS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Item Number
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("IteNumGDS7"))
            objChildElement.Text = lngctr
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Commodity Code
            strTemp = GetValueForSegment(rstDetailCST, "CST", "33", "2", True, lngctr)
            If Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0" Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("ComCodTarCodGDS10"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Declaration Type
            strTemp = GetValueForSegment(rstDetailCST, "CST", "33", "14", True, lngctr)
            If Len(Trim$(strTemp)) > 0 And Trim$(strTemp) <> "0" Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("DecTypGDS15"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Goods Description
            strTemp = GetValueForSegment(rstDetailFTX, "FTX", "34", "6", True, lngctr) & GetValueForSegment(rstDetailFTX, "FTX", "34", "7", True, lngctr)
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GooDesGDS23"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
            'Goods Description Language
            If GetSegmentOptionForLanguage("GooDesGDS23LNG") = False Or _
               Len(Trim$(strTemp)) > 0 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("GooDesGDS23LNG"))
                objChildElement.Text = m_strA5
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Gross Mass
            strTemp = GetValueForSegment(rstDetailMEA, "MEA", "37", "7", True, lngctr)
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("GroMasGDS46"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Net Mass
            strTemp = GetValueForSegment(rstDetailMEA, "MEA", "38", "7", True, lngctr)
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("NetMasGDS48"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Country of Dispatch
            strTemp = GetValueForSegment(rstDetailLOC, "LOC", "35", "2", True, lngctr)
            If Len(Trim$(strTemp)) >= 2 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouOfDisGDS58"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            'Country of Destination
            strTemp = GetValueForSegment(rstDetailLOC, "LOC", "36", "2", True, lngctr)
            If Len(Trim$(strTemp)) >= 2 Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouOfDesGDS59"))
                objChildElement.Text = strTemp
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            
            '*******************************************************************************************************
            'Previous Document
            '*******************************************************************************************************
            'Retrieve Box Y1 and Y5
            GetGroupRecordsFromDataNCTSTables "DATA_NCTS_DETAIL_DOCUMENTEN", rstTemp, lngctr
            
            ReDim varDocType(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            
            If (UBound(varDocType) + 1) = rstTemp.RecordCount Then
                
                rstTemp.MoveFirst
                
                For lngPrevDocCtr = 0 To UBound(varDocType)
                    If UCase$(Trim$(rstTemp.Fields("Y1").Value)) = "V" Then
                        
                        lngStoragePreviuosDocs = lngStoragePreviuosDocs + 1
                        
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PREADMREFAR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Previous Document Type
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "4", True, lngPrevDocCtr + 1)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("PreDocTypAR21"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Previous Document Reference
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "5", True, lngPrevDocCtr + 1)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("PreDocRefAR26"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Previous Document Reference Language
                            If GetSegmentOptionForLanguage("PreDocRefLNG") = False Or _
                               Len(Trim$(GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "Y3", lngctr)(lngPrevDocCtr))) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("PreDocRefLNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Complement of Information
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "7", True, lngPrevDocCtr + 1)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfAR29"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Complement of Information Language
                            If GetSegmentOptionForLanguage("ComOfInfAR29LNG") = False Or _
                               Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfAR29LNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                    End If
                    
                    If UCase$(Trim$(rstTemp.Fields("Y5").Value)) = "E" Then Exit For
                    
                    rstTemp.MoveNext
                Next
                
            End If
            '*******************************************************************************************************
                
            '*******************************************************************************************************
            'Produced Document
            '*******************************************************************************************************
            'Retrieve Box Y1 and Y5
            ''commented allan ncts redundant
            'GetGroupRecordsFromDataNCTSTables "DATA_NCTS_DETAIL_DOCUMENTEN", rstTemp, lngCtr
            'varDocType = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "Y2", lngCtr)
            
            If (UBound(varDocType) + 1) = rstTemp.RecordCount Then
                
                rstTemp.MoveFirst
                
                For lngPrevDocCtr = 0 To UBound(varDocType)
                   
                    lngStorageProducedDocs = lngStorageProducedDocs + 1
                    
                    If UCase$(Trim$(rstTemp.Fields("Y1").Value)) = "P" Then
                        
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PRODOCDC2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Produced Document Type
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "4", True, lngStorageProducedDocs)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocTypDC21"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Produced Document Reference
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "5", True, lngStorageProducedDocs)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocRefDC23"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Produced Document Reference Language
                            If GetSegmentOptionForLanguage("DocRefDCLNG") = False Or _
                               Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocRefDCLNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Complement of Information
                            strTemp = GetValueForSegment(rstDetailDOC, "DOC", "44", "7", True, lngStorageProducedDocs)
                            If Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfDC25"))
                                objChildElement2.Text = strTemp
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Complement of Information Language
                            If GetSegmentOptionForLanguage("ComOfInfDC25LNG") = False Or _
                               Len(Trim$(strTemp)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfDC25LNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                    End If
                    
                    If UCase$(Trim$(rstTemp.Fields("Y5").Value)) = "E" Then Exit For
                    
                    rstTemp.MoveNext
                Next
                
            End If
            '*******************************************************************************************************
                
            '*******************************************************************************************************
            'Special Mentions
            '*******************************************************************************************************
            'Retrieve box Z4
            GetGroupRecordsFromDataNCTSTables "DATA_NCTS_DETAIL_BIJZONDERE", rstTemp, lngctr
            
            ReDim varAddInfoZ1(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varAddInfoZ2(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varAddInfoZ3(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            
            rstDetailFTX.Filter = adFilterNone
            rstDetailFTX.Filter = "[NCTS_IEM_TMS_ID] = 47 and [DATA_NCTS_FTX_Instance]= " & lngctr & " "
            'strTemp = GetValueIfNotNull(rstDetailDOC.Fields("DATA_NCTS_DOC_Seq7").Value) 'allan ncts not sure
            
            If (UBound(varAddInfoZ1) + 1) = rstTemp.RecordCount Then
            
                rstTemp.MoveFirst
                
                For lngAddInfo = 0 To UBound(varAddInfoZ1)
                    
                    If lngStoragSpecialMention = 0 And lngctr = 1 And lngAddInfo = 0 Then
                        lngStoragSpecialMention = 0
                    Else
                        lngStoragSpecialMention = lngStoragSpecialMention + 1
                    End If
                    
                    'lngStoragSpecialMention = lngStoragSpecialMention + 1
                    Dim strTemp2 As String
                    strTemp = GetValueForSegment(rstDetailFTX, "FTX", "47", "3", True, lngStoragSpecialMention + 1)
                    strTemp2 = GetValueForSegment(rstDetailFTX, "FTX", "47", "6", True, lngStoragSpecialMention + 1)
                    
                    If Len(Trim$(strTemp)) = 0 Or _
                    Len(Trim$(strTemp2)) = 0 Then
                        
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("SPEMENMT2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                    End If
                    
                    If Len(Trim$(strTemp)) Or _
                    Len(Trim$(strTemp2)) Then
                    
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("SPEMENMT2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Additional Information
                            strTemp = GetValueForSegment(rstDetailFTX, "FTX", "47", "6", True, lngStoragSpecialMention + 1)
                            If Len(Trim$(strTemp)) > 0 Then 'allan ncts
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AddInfMT21"))
                                objChildElement2.Text = GetValueForSegment(rstDetailFTX, "FTX", "47", "6", True, lngStoragSpecialMention + 1)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Additional Information Language
                            If GetSegmentOptionForLanguage("AddInfMT21LNG") = False Or _
                               Len(Trim$(GetValueIfNotNull(rstDetailFTX.Fields("DATA_NCTS_FTX_Seq6").Value))) > 0 Then 'allan ncts
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AddInfMT21LNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Additional Information Code
                            If Len(Trim$(GetValueIfNotNull(rstDetailFTX.Fields("DATA_NCTS_FTX_Seq3").Value))) > 0 Then 'allan ncts
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AddInfCodMT23"))
                                objChildElement2.Text = GetValueForSegment(rstDetailFTX, "FTX", "47", "3", True, lngStoragSpecialMention + 1)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            If Trim$(GetValueIfNotNull(rstDetailFTX.Fields("DATA_NCTS_FTX_Seq3").Value)) = "DG0" Or _
                               Trim$(GetValueIfNotNull(rstDetailFTX.Fields("DATA_NCTS_FTX_Seq3").Value)) = "DG1" Then

                                strTemp = GetValueForSegment(rstDetailTOD, "TOD", "46", "3", True, lngStoragSpecialMention + 1)
                                'Export from EC - FOR CONFIRMATION
                                If Len(Trim$(strTemp)) = 2 Then
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ExpFroECMT24"))
                                    objChildElement2.Text = strTemp
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                            ElseIf Trim$(GetValueIfNotNull(rstDetailFTX.Fields("DATA_NCTS_FTX_Seq3").Value)) = "DG2" Then 'allan ncts
                                'Export from Country - FOR CONFIRMATION
                                If Len(Trim$(strTemp)) = 2 Then 'allan ncts
                                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ExpFroCouMT25"))
                                    objChildElement2.Text = strTemp
                                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                                End If
                                
                            End If
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                    End If
                    
                    If UCase$(Trim$(rstTemp.Fields("Z4").Value)) = "E" Then Exit For
                    
                    rstTemp.MoveNext
                Next
                
            End If
            
            '*******************************************************************************************************
            
            '*******************************************************************************************************
            'Consignor
            '*******************************************************************************************************
            If m_blnU6IsTheSame = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("TRACONCO2"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Name
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NamCO27"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "30", "10")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'Street and Number
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("StrAndNumCO222"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "30", "16")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'Postal Code
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("PosCodCO223"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "30", "22")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'City
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CitCO224"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "30", "20")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Country Code
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CouCO225"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "30", "23")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'NAD Language
                    If Len(Trim$(m_strA5)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NADLNGGTCO"))
                        objChildElement2.Text = m_strA5
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'TIN
                    strTemp = GetValueForSegment(rstNAD, "NAD", "30", "2")
                    If Len(Trim$(strTemp)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TINCO259"))
                        objChildElement2.Text = strTemp
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                            
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            '*******************************************************************************************************
            
            '*******************************************************************************************************
            'Consignee
            '*******************************************************************************************************
            If m_blnW6IsTheSame = False Then
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("TRACONCE2"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Name
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NamCE27"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "29", "10")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'Street and Number
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("StrAndNumCE222"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "29", "16")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'Postal Code
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("PosCodCE223"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "29", "22")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'City
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CitCE224"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "29", "20")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Country Code
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CouCE225"))
                    objChildElement2.Text = GetValueForSegment(rstNAD, "NAD", "29", "23")
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'NAD Language
                    If Len(Trim$(m_strA5)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NADLNGGICE"))
                        objChildElement2.Text = m_strA5
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'TIN
                    strTemp = GetValueForSegment(rstNAD, "NAD", "29", "2")
                    If Len(Trim$(strTemp)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("TINCE259"))
                        objChildElement2.Text = strTemp
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                            
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            '*******************************************************************************************************
            
            '*******************************************************************************************************
            'Container
            '*******************************************************************************************************
            'Retrieve box SB
            GetGroupRecordsFromDataNCTSTables "DATA_NCTS_DETAIL_CONTAINER", rstTemp, lngctr
          
            ReDim varContNumS6(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varContNumS7(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varContNumS8(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varContNumS9(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varContNumSA(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            
            If (UBound(varContNumS6) + 1) = rstTemp.RecordCount Then
            
                rstTemp.MoveFirst
            
                For lngContainerCtr = 0 To UBound(varContNumS6)
                    
                    If lngContainerCtr = 0 And lngctr = 1 Then
                        lngStorageContainer = 0
                    Else
                        lngStorageContainer = lngStorageContainer + 5
                    End If
                    strTemp = GetValueForSegment(rstDetailRFF, "RFF", "43", "2", True, (lngStorageContainer + 1))
                    If strTemp <> vbNullString Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
                            objChildElement2.Text = Trim$(strTemp)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    strTemp = GetValueForSegment(rstDetailRFF, "RFF", "43", "2", True, (lngStorageContainer + 2))
                    If strTemp <> vbNullString Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
                            objChildElement2.Text = Trim$(strTemp)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    strTemp = GetValueForSegment(rstDetailRFF, "RFF", "43", "2", True, (lngStorageContainer + 3))
                    If strTemp <> vbNullString Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
                            objChildElement2.Text = Trim$(strTemp)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                     
                    strTemp = GetValueForSegment(rstDetailRFF, "RFF", "43", "2", True, (lngStorageContainer + 4))
                    If strTemp <> vbNullString Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
                            objChildElement2.Text = Trim$(strTemp)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    strTemp = GetValueForSegment(rstDetailRFF, "RFF", "43", "2", True, (lngStorageContainer + 5))
                    If strTemp <> vbNullString Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
                            objChildElement2.Text = Trim$(strTemp)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    If UCase$(Trim$(rstTemp.Fields("SB").Value)) = "E" Then Exit For
                    
                    rstTemp.MoveNext
                Next
                
            End If
            '*******************************************************************************************************
            
            '*******************************************************************************************************
            'Packages
            '*******************************************************************************************************
            'Retrieve box SB
            GetGroupRecordsFromDataNCTSTables "DATA_NCTS_DETAIL_COLLI", rstTemp, lngctr
   
            ReDim varMarksAndNum(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varKindPackage(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            ReDim varNumPackage(IIf(rstTemp.RecordCount = 1, 0, (rstTemp.RecordCount - 1)))
            
            If (UBound(varMarksAndNum) + 1) = rstTemp.RecordCount Then
            
                rstTemp.MoveFirst
                
                For lngPackagesCtr = 0 To UBound(varMarksAndNum)
                    
                    lngStoragePackage = lngStoragePackage + 1
                    
                    strKindPackage = GetValueForSegment(rstDetailPAC, "PAC", "41", "9", True, lngStoragePackage)
                    If Len(Trim$(strKindPackage)) > 0 Then
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PACGS2"))
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Marks and Numbers
                            strMarksAndNum = GetValueForSegment(rstDetailPCI, "PCI", "42", "2", True, lngStoragePackage)
                            If Len(Trim$(strMarksAndNum)) > 0 Then 'varMarksAndNum(lngPackagesCtr))) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("MarNumOfPacGS21"))
                                objChildElement2.Text = Trim$(strMarksAndNum) 'varMarksAndNum(lngPackagesCtr))
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Marks and Numbers Language
                            If GetSegmentOptionForLanguage("MarNumOfPacGS21LNG") = False Or _
                               Len(Trim$(strMarksAndNum)) > 0 Then 'varMarksAndNum(lngPackagesCtr))) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("MarNumOfPacGS21LNG"))
                                objChildElement2.Text = m_strA5
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Package Type
                            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("KinOfPacGS23"))
                            objChildElement2.Text = strKindPackage 'varKindPackage(lngPackagesCtr)
                            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            
                            'Number of Packages
                            strNumPackage = GetValueForSegment(rstDetailPAC, "PAC", "41", "10", True, lngStoragePackage)
                            If Len(Trim$(strNumPackage)) > 0 Then
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NumOfPacGS24"))
                                objChildElement2.Text = strNumPackage 'varNumPackage(lngPackagesCtr)
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                            'Number of Pieces 'CSCLP-168
                            'If Len(Trim$(GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "S3", lngctr)(lngPackagesCtr))) > 0 Then
                            '    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NumOfPieGS25"))
                            '    objChildElement2.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "S3", lngctr)(lngPackagesCtr)
                            '    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            'End If
                            
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    If UCase$(Trim$(rstTemp.Fields("S5").Value)) = "1" Or _
                       UCase$(Trim$(rstTemp.Fields("S5").Value)) = "0" Then Exit For
                    
                    rstTemp.MoveNext
                Next
                
            End If
            '*******************************************************************************************************
            
            '*******************************************************************************************************
            'Sensitive Goods
            '*******************************************************************************************************
            
            lngSensitiveGoods = lngSensitiveGoods + 1
            strV2 = GetValueForSegment(rstDetailGIR, "GIR", "48", "2", True, lngSensitiveGoods, "V2")
            strV1 = GetValueForSegment(rstDetailGIR, "GIR", "48", "5", True, lngSensitiveGoods, "V2")
            lngSensitiveGoods = lngSensitiveGoods + 1
            strV4 = GetValueForSegment(rstDetailGIR, "GIR", "48", "2", True, lngSensitiveGoods, "V4")
            strV3 = GetValueForSegment(rstDetailGIR, "GIR", "48", "5", True, lngSensitiveGoods, "V2")
            lngSensitiveGoods = lngSensitiveGoods + 1
            strV6 = GetValueForSegment(rstDetailGIR, "GIR", "48", "2", True, lngSensitiveGoods, "V6")
            strV5 = GetValueForSegment(rstDetailGIR, "GIR", "48", "5", True, lngSensitiveGoods, "V2")
            lngSensitiveGoods = lngSensitiveGoods + 1
            strV8 = GetValueForSegment(rstDetailGIR, "GIR", "48", "2", True, lngSensitiveGoods, "V8")
            strV7 = GetValueForSegment(rstDetailGIR, "GIR", "48", "5", True, lngSensitiveGoods, "V2")
            
            If strV2 <> "0" Or strV4 <> "0" Or strV6 <> "0" Or strV8 <> "0" Then
               
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("SGICODSD2"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Sensitive Goods Code / Sensitive Goods Quantity
                    If strV2 <> "0" Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenGooCodSD22"))
                        objChildElement2.Text = strV1
                        'objChildElement2.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "V1", lngCtr)(0)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenQuaSD23"))
                        objChildElement2.Text = strV2
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Sensitive Goods Code / Sensitive Goods Quantity
                    If strV4 <> "0" Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenGooCodSD22"))
                        objChildElement2.Text = strV3 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "V3", lngCtr)(0)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenQuaSD23"))
                        objChildElement2.Text = strV4
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Sensitive Goods Code / Sensitive Goods Quantity
                    If strV6 <> "0" Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenGooCodSD22"))
                        objChildElement2.Text = strV5 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "V5", lngCtr)(0)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenQuaSD23"))
                        objChildElement2.Text = strV6
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Sensitive Goods Code / Sensitive Goods Quantity
                    If strV8 <> "0" Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenGooCodSD22"))
                        objChildElement2.Text = strV7 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "V7", lngCtr)(0)
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                        
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenQuaSD23"))
                        objChildElement2.Text = strV8
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
            '*******************************************************************************************************
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    Next
    
End Sub
    



