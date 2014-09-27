Attribute VB_Name = "MArrivals"
Option Explicit

'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <objChildElement>
'               <objChildElement2>
'                   <objChildElement3>

Public Sub CreateXMLMessageIE07(ByRef DataSourceProperties As CDataSourceProperties, _
                                ByRef objDOM As DOMDocument, _
                                ByRef objParentNode As IXMLDOMNode, _
                                ByRef objChildNode As IXMLDOMNode)
    
    Dim objChildElement As IXMLDOMNode
    Dim objChildElement2 As IXMLDOMNode
    
    'Interchange
    CreateMessageInterchangeIE07 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Header
    CreateMessageHeaderIE07 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Trader
    CreateMessageTraderIE07 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Customs Office
    CreateMessageCustomsOfficeIE07 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'En Route Events
    CreateMessageEnRouteEventsIE07 objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    Set objChildElement = Nothing
    Set objChildElement2 = Nothing
     
End Sub


Private Sub CreateMessageInterchangeIE07(ByRef objDOM As DOMDocument, _
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
    objChildNode.Text = "CC007A"
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



Private Sub CreateMessageHeaderIE07(ByRef objDOM As DOMDocument, _
                                    ByRef objParentNode As IXMLDOMNode, _
                                    ByRef objChildNode As IXMLDOMNode, _
                                    ByRef objChildElement As IXMLDOMNode, _
                                    ByRef objChildElement2 As IXMLDOMNode)
    
    'Header
    Set objChildNode = objDOM.createElement("HEAHEA")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Document Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("DocNumHEA5"))
        objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).MOVEMENT_REFERENCE_NUMBER
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Customs Sub Place - FOR CONFIRMATION
        'Set objChildElement = objChildNode.appendChild(objDOM.createElement("CusSubPlaHEA66"))
        'objChildElement.Text = vbNullString
        'objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Notification Place
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrNotPlaHEA60"))
        objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_NOTIFICATION_PLACE
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Notification Place Language
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrNotPlaHEA60LNG"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_NOTIFICATION_PLACE_LNG
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Agreed Location Code
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_CODE)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrAgrLocCodHEA62"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_CODE
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Agreed Location of Goods
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_OF_GOODS)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrAgrLocOfGooHEA63"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_OF_GOODS
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Agreed Location of Goods Language
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrAgrLocOfGooHEA63LNG"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Authorised Location of Goods
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AUTHORISED_LOCATION_OF_GOODS)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrAutLocOfGooHEA65"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_AUTHORISED_LOCATION_OF_GOODS
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Simplified Procedure Flag
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("SimProFlaHEA132"))
        objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).SIMPLIFIED_PROCEDURE_FLAG
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Notification Date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("ArrNotDatHEA141"))
        objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).ARRIVAL_NOTIFICATION_DATE
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Dialog Language Indicator at Destination
        If Len(Trim$(G_clsEDIArrival.Headers(G_strHeaderKey).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("DiaLanIndAtDesHEA255"))
            objChildElement.Text = G_clsEDIArrival.Headers(G_strHeaderKey).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub


Private Sub CreateMessageTraderIE07(ByRef objDOM As DOMDocument, _
                                    ByRef objParentNode As IXMLDOMNode, _
                                    ByRef objChildNode As IXMLDOMNode, _
                                    ByRef objChildElement As IXMLDOMNode, _
                                    ByRef objChildElement2 As IXMLDOMNode)
    
    'Trader
    Set objChildNode = objDOM.createElement("TRADESTRD")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Name
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_NAME)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamTRD7"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_NAME
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Street and Number
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_STREET_AND_NUMBER)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumTRD22"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_STREET_AND_NUMBER
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Postal Code
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_POSTAL_CODE)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodTRD23"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_POSTAL_CODE
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'City
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_CITY)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitTRD24"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_CITY
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Country Code
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_COUNTRY_CODE)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouTRD25"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_COUNTRY_CODE
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'NAD Language
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_NAD_LNG)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGRD"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_NAD_LNG
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'TIN Number
        If Len(Trim$(G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_TIN)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRD59"))
            objChildElement.Text = G_clsEDIArrival.Traders(G_strTraderKey).DESTINATION_TIN
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub


Private Sub CreateMessageCustomsOfficeIE07(ByRef objDOM As DOMDocument, _
                                           ByRef objParentNode As IXMLDOMNode, _
                                           ByRef objChildNode As IXMLDOMNode, _
                                           ByRef objChildElement As IXMLDOMNode, _
                                           ByRef objChildElement2 As IXMLDOMNode)
    
    
    If Len(Trim$(G_clsEDIArrival.CustomOffices(G_strCustomOfcKey).REFERENCE_NUMBER)) > 0 Then
        
        Set objChildNode = objDOM.createElement("CUSOFFPREOFFRES")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
            'Reference Number
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRES1"))
            objChildElement.Text = G_clsEDIArrival.CustomOffices(G_strCustomOfcKey).REFERENCE_NUMBER
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
    End If
        
End Sub



Private Sub CreateMessageEnRouteEventsIE07(ByRef objDOM As DOMDocument, _
                                           ByRef objParentNode As IXMLDOMNode, _
                                           ByRef objChildNode As IXMLDOMNode, _
                                           ByRef objChildElement As IXMLDOMNode, _
                                           ByRef objChildElement2 As IXMLDOMNode)
    
    
    Dim objChildElement3 As IXMLDOMNode
    
    Dim lngctr As Long
    Dim lngSealsCtr As Long
    Dim lngContCtr As Long
    
    If G_clsEDIArrival.EnRouteEvents.Count > 0 Then
        
        For lngctr = 1 To G_clsEDIArrival.EnRouteEvents.Count
            
            'En Route Events
            Set objChildNode = objDOM.createElement("ENROUEVETEV")
            objDOM.documentElement.appendChild objChildNode
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
                'Place
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("PlaTEV10"))
                objChildElement.Text = G_clsEDIArrival.EnRouteEvents(lngctr).PLACE
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                'Place Language
                If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).PLACE_LNG)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("PlaTEV10LNG"))
                    objChildElement.Text = G_clsEDIArrival.EnRouteEvents(lngctr).PLACE_LNG
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                'Country Code
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouTEV13"))
                objChildElement.Text = G_clsEDIArrival.EnRouteEvents(lngctr).COUNTRY_CODE
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                '********************************************************************************
                'Control +
                '********************************************************************************
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("CTLCTL"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
                    'Already in NCTS Flag
                    Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("AlrInNCTCTL29"))
                    objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Ctl_Controls(1).ALREADY_IN_NCTS
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                '********************************************************************************
                
                '********************************************************************************
                'Incident +
                '********************************************************************************
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("INCINC"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'Incident Flag
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_FLAG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("IncFlaINC3"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_FLAG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Incident Information
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_INFORMATION)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("IncInfINC4"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_INFORMATION
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Incident Information Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_INFORMATION_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("IncInfINC4LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).INCIDENT_INFORMATION_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Date
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_DATE)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndDatINC6"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_DATE
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Authority
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_AUTHORITY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndAutINC7"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_AUTHORITY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Authority Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_AUTHORITY_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndAutINC7LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_AUTHORITY_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Place
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_PLACE)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndPlaINC10"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_PLACE
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Place Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_PLACE_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndPlaINC10LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_PLACE_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                    'Endorsement Country
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_COUNTRY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndCouINC12"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Incidents(1).ENDORSEMENT_COUNTRY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                '********************************************************************************
                
                '********************************************************************************
                'Seals +
                '********************************************************************************
                If G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).NEW_SEALS_NO > 0 Then
                    
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("SEAINFSF1"))
                    objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                        'Seals Count
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaNumSF12"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).NEW_SEALS_NO
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                        For lngSealsCtr = 1 To G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).NEW_SEALS_NO
                            
                            If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START)) > 0 Then
                                
                                'Seals ID
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SEAIDSI1"))
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                                    If IsNumeric(G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START) = True And _
                                       IsNumeric(G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_END) = True Then
                                        
                                        If G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START < G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_END Then
                                            'Seals Identity
                                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            objChildElement3.Text = CLng(G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START) + CLng(lngSealsCtr - 1)
                                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                            
                                            'Seals Identity Language
                                            'Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            'objChildElement3.Text = ""
                                            'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                        
                                        Else
                                            
                                            If lngSealsCtr = 1 Then
                                                'Seals Identity
                                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                                objChildElement3.Text = G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START
                                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                                
                                                'Seals Identity Language
                                                'Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                                'objChildElement3.Text = ""
                                                'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                                
                                            ElseIf lngSealsCtr = 2 Then
                                                'Seals Identity
                                                Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                                objChildElement3.Text = G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_END
                                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                                
                                                'Seals Identity Language
                                                'Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                                'objChildElement3.Text = ""
                                                'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                                
                                            End If
                                            
                                        End If
                                    
                                    Else
                                        
                                        If lngSealsCtr = 1 Then
                                            'Seals Identity
                                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            objChildElement3.Text = G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_START
                                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                            
                                            'Seals Identity Language
                                            'Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            'objChildElement3.Text = ""
                                            'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                            
                                        ElseIf lngSealsCtr = 2 Then
                                            'Seals Identity
                                            Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            objChildElement3.Text = G_clsEDIArrival.EnRouteEvents(lngctr).NewSeals(1).SEAL_END
                                            objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                            
                                            'Seals Identity Language
                                            'Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("SeaIdeSI11"))
                                            'objChildElement3.Text = ""
                                            'objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                            
                                        End If
                                    
                                    End If
                                    
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                        Next
                        
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                End If
                '********************************************************************************
                
                '********************************************************************************
                'Transhipment +
                '********************************************************************************
                Set objChildElement = objChildNode.appendChild(objDOM.createElement("TRASHP"))
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
                    'New Transport Means Identity
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_IDENTITY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NewTraMeaIdeSHP26"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_IDENTITY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'New Transport Means Identity Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_IDENTITY_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NewTraMeaIdeSHP26LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_IDENTITY_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'New Transport Means Identity Nationality
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_NATIONALITY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NewTraMeaNatSHP54"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).NEW_TRANSPORT_MEANS_NATIONALITY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Date
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_DATE)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndDatSHP60"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_DATE
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Authority
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_AUTHORITY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndAutSHP61"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_AUTHORITY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Authority Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_AUTHORITY_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndAutSHP61LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_AUTHORITY_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Place
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_PLACE)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndPlaSHP63"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_PLACE
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Place Language
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_PLACE_LNG)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndPlaSHP63LNG"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_PLACE_LNG
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    'Endorsement Country
                    If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_COUNTRY)) > 0 Then
                        Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("EndCouSHP65"))
                        objChildElement2.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).ENDORSEMENT_COUNTRY
                        objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    End If
                    
                    If G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).Containers.Count > 0 Then
                        
                        For lngContCtr = 1 To G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).Containers.Count
                            
                            If Len(Trim$(G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).Containers(lngContCtr).NEW_CONTAINER_NUMBER)) > 0 Then
                                'Containers +
                                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("CONNR3"))
                                objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                    
                                    'Container Number
                                    Set objChildElement3 = objChildElement2.appendChild(objDOM.createElement("ConNumNR31"))
                                    objChildElement3.Text = G_clsEDIArrival.EnRouteEvents(lngctr).Transhipments(1).Containers(lngContCtr).NEW_CONTAINER_NUMBER
                                    objChildElement2.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab & vbTab)
                                                        
                                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                            End If
                            
                        Next
                        
                    End If
                    
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                '********************************************************************************
                
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            
            
            Set objChildElement3 = Nothing
        Next
        
    End If
        
End Sub

