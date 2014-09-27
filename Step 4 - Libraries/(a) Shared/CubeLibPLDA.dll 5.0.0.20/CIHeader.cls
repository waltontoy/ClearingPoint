VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_rstHeaderHandelaars As ADODB.Recordset
Public m_rstHeader As ADODB.Recordset
Public m_rstHeaderZegels As ADODB.Recordset
Public m_enuDeclarationMode As DECLARATION_MODE
Public m_lngHeaderSegments As Long

Private m_strSender As String
Private m_strIEReference As String
Private m_blnUsedCurrencyUnknown As Boolean
Private m_strRecipient As String
Private m_strSegmentTemp As String
Private m_lngHeaderSeals As Long


Public Function CreateHeader(ByRef SegmentCount As Long, _
                             ByVal RecordsHeader As ADODB.Recordset, _
                             ByVal RecordsHeaderZegels As ADODB.Recordset, _
                             ByVal RecordsHeaderHandelaars As ADODB.Recordset, _
                             ByVal DeclarationMode As DECLARATION_MODE, _
                             ByVal EDIUserName As String, _
                             ByVal LogTIN As String, _
                             ByVal Recipient As String, _
                             ByVal LastEDIRef As String) As String
                        
    Dim strMessage As String
    
    Set m_rstHeader = RecordsHeader.Clone
    Set m_rstHeaderZegels = RecordsHeaderZegels.Clone
    Set m_rstHeaderHandelaars = RecordsHeaderHandelaars.Clone
    
    m_lngHeaderSegments = 0                 'Number of header segments
    m_enuDeclarationMode = DeclarationMode  'Mode of the Declaration
    m_strSender = EDIUserName               'Sender
    m_strRecipient = Recipient              'Recipient
    
    m_strIEReference = Left(LogTIN, 9) & Format(LastEDIRef, "00000")
    
    '*******************************************************************************************
    'Create the header
    '*******************************************************************************************
    strMessage = vbNullString
    
    strMessage = UNB
    strMessage = strMessage & UNH
    strMessage = strMessage & BGM
    strMessage = strMessage & CST
    strMessage = strMessage & LOC
    strMessage = strMessage & DTM
    strMessage = strMessage & GEI
    strMessage = strMessage & FII
    strMessage = strMessage & MEA
    strMessage = strMessage & EQD
    strMessage = strMessage & SEL
    strMessage = strMessage & FTX
    strMessage = strMessage & SG1
    strMessage = strMessage & SG4
    strMessage = strMessage & SG5
    strMessage = strMessage & SG6
    strMessage = strMessage & SG7
    strMessage = strMessage & SG8
    '*******************************************************************************************
    
    '******************************************************************************************************
    'Set Summary Digital Signature for D1, D2, and D3
    'Based on D1, D2, and D3 in CISummary Segment
    '******************************************************************************************************
    If m_enuDeclarationMode = enuOriginal Then
        'Number of Packages
        AddFieldToDigiSign "D3", m_rstHeader, False
        
        'Gross Weight
        AddFieldToDigiSign "D1", m_rstHeader, False
        
        'Net Weight
        AddFieldToDigiSign "D2", m_rstHeader, False
        
        'Total Number of Seals
        AddFieldToDigiSign "Total Number of Seals", m_rstHeaderZegels, False
    End If
    '******************************************************************************************************
    
    'Set the segment count
    SegmentCount = SegmentCount + m_lngHeaderSegments
    
    'Destroy the instance of the recordsets
    ADORecordsetClose m_rstHeader
    ADORecordsetClose m_rstHeaderHandelaars
    ADORecordsetClose m_rstHeaderZegels
    
    'Return the created header EDI Message
    CreateHeader = strMessage
    
End Function

Friend Property Get UNB() As String
    
    
    UNB = "UNB+UNOA:2+" & ReplaceSpecialCharacters(m_strSender)         'Sender
    UNB = UNB & "+" & ReplaceSpecialCharacters(m_strRecipient)          'Recipient
    UNB = UNB & "+" & Format(Now, "yyMMdd")                             'Date
    UNB = UNB & ":" & Format(Now, "hhmm")                               'Time
    UNB = UNB & "+" & ReplaceSpecialCharacters(m_strIEReference) & "'"  'Interchange Reference
    
End Property

Friend Property Get UNH() As String    'M1

    'Message header
    'UNH+1+CUSDEC:D:04A:UN:IMPORT+<MRN>

    UNH = "UNH+1+CUSDEC:D:04A:UN:IMPORT'"
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(UNH)
    
End Property

Friend Property Get BGM() As String    'M1

    'Beginning of message
    'BGM+:::<type of declaration>+<MRN if for modification>:<language>+<modification mode>
    
    'Type of Declaration
    BGM = "BGM+:::" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A1").Value)) & _
                      ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A2").Value))
        
    AddFieldToDigiSign "A1", m_rstHeader, False
    AddFieldToDigiSign "A2", m_rstHeader, False
    
    BGM = BGM & "+"
        
    'MRN
    If m_enuDeclarationMode <> enuOriginal Then
        BGM = BGM & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("MRN").Value))
    End If
    
    'Language of the Declaration
    BGM = BGM & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A9").Value))
    
    AddFieldToDigiSign "A9", m_rstHeader, False
    
    'Declaration Mode
    BGM = BGM & "+" & ReplaceSpecialCharacters(m_enuDeclarationMode) & "'"
    BGM = TrimSegment(BGM, "BGM")
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(BGM)
    
End Property

Friend Property Get CST() As String    'C1
'Customs status of goods
End Property

Friend Property Get LOC() As String    'C99

    'Place/location identification
    'LOC+<Qualifier>+F<1>
        
    Dim strTemp As String
    
    'Customs Office of Validation
    strTemp = "LOC+41+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A6").Value)) & "'"
    
    AddFieldToDigiSign "A6", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+41")
    
    'Location of Goods
    strTemp = "LOC+43+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D4").Value)) & "'"
    
    AddFieldToDigiSign "D4", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+43")
        
    'Country of Sending
    If CheckRegimeType("J", "<>") And CheckRegimeType("K", "<>") Then
        strTemp = "LOC+35+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("DA").Value)) & "'"
        
        AddFieldToDigiSign "DA", m_rstHeader, False
        
        LOC = LOC & TrimSegment(strTemp, "LOC+35")
    End If
    
    'Place of the Declaration
    strTemp = "LOC+91+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A5").Value)) & "'"
    
    AddFieldToDigiSign "A5", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+91")

    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(LOC)
    
End Property

Friend Property Get DTM() As String    'C15
    
    'Date/time/period
    'DTM+<Qualifier>:F<yyyymmdd>:102
    
    DTM = "DTM+137:" & Format(Now, "yyyyMMdd") & ":102'"
    
    DTM = DTM & TrimSegment("DTM+254:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A4").Value)) & ":102'", "DTM+254::102")
    
    AddFieldToDigiSign "A4", m_rstHeader, False
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(DTM)
    
End Property

Friend Property Get GEI() As String    'C25
    
    'Processing information
    'GEI+2+<Processing indicator>:ZZZ:147
    
    'Manner of Payment
    GEI = "GEI+2+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B1").Value)) & ":ZZZ:147'"
    
    AddFieldToDigiSign "B1", m_rstHeader, False
    
    GEI = TrimSegment(GEI, "GEI+2+:ZZZ:147")
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(GEI)
        
End Property

Friend Property Get FII() As String    'C1

    'Financial institution information
    'FII+AE+<acct49 + customs office>:<TIN>
    
    Dim strTemp As String
    
    If Not (CheckRegimeType("J", "=")) Then
        'Account Number
        strTemp = "FII+AE+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B5").Value))
        
        AddFieldToDigiSign "B5", m_rstHeader, False
        
        m_rstHeaderHandelaars.MoveFirst
        m_rstHeaderHandelaars.Find "XE = '" & enuHeaderDeclarant & "'"
        
        If (m_rstHeaderHandelaars.EOF = False) Then
            
            If UCase(FNullField(m_rstHeaderHandelaars.Fields("X1").Value)) <> UCase(FNullField(m_rstHeader.Fields("B4").Value)) Then
                'Account Holder
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B4").Value))
                
                AddFieldToDigiSign "B4", m_rstHeader, False
            End If
        Else
            'Account Holder
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B4").Value))
            
            AddFieldToDigiSign "B4", m_rstHeader, False
        End If
        
        strTemp = strTemp & "'"
        
        FII = TrimSegment(strTemp, "FII+AE")
        
        m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(FII)
    End If
    
End Property

Friend Property Get MEA() As String    'C5
'Measurements
End Property

Friend Property Get EQD() As String  'C999
'Equipment details
End Property

Friend Property Get SEL() As String    'C999

    'Seal number
    'SEL+0+<seal issuer>:::<seal number>
    
    Dim strTemp As String
    
    With m_rstHeaderZegels
        If .RecordCount > 0 Then
            .Sort = "ORDINAL"
            
            .MoveFirst
            
            Do While Not .EOF
                'Responsible Sealing
                strTemp = "SEL+0+" & ReplaceSpecialCharacters(FNullField(.Fields("E2").Value))
                
                AddFieldToDigiSign "E2", m_rstHeaderZegels, False, , .Fields("Ordinal").Value
                
                'Identification
                strTemp = strTemp & ":::" & ReplaceSpecialCharacters(FNullField(.Fields("E1").Value)) & "'"
                
                AddFieldToDigiSign "E1", m_rstHeaderZegels, False, , .Fields("Ordinal").Value
                
                SEL = SEL & TrimSegment(strTemp, "SEL+0")
                           
                If FNullField(.Fields("E3").Value) = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
    End With
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SEL)
    
End Property

Friend Property Get FTX() As String    'C99

    'Free text
    'FTX+<Qualifier>++<payment code>+<payment literal>
    
    Dim strTemp As String
    
    'Transaction Type
    If Not (CheckRegimeType("K", "=") Or CheckRegimeType("J", "=")) Then
        strTemp = "FTX+AAB++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C7").Value)) & "'"
        
        AddFieldToDigiSign "C7", m_rstHeader, False
        
        FTX = FTX & TrimSegment(strTemp, "FTX+AAB")
    End If
        
    'Manner of Payment VAT
    strTemp = "FTX+PMT++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B2").Value))
    
    AddFieldToDigiSign "B2", m_rstHeader, False
    
    'Postponed payment VAT
    If UCase(FNullField(m_rstHeader.Fields("B2").Value)) = "B" Or _
       UCase(FNullField(m_rstHeader.Fields("B2").Value)) = "L" Or _
       UCase(FNullField(m_rstHeader.Fields("B2").Value)) = "G" Then
            
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("B6").Value))
        
        AddFieldToDigiSign "B6", m_rstHeader, False
    End If
        
    strTemp = strTemp & "'"
    
    FTX = FTX & TrimSegment(strTemp, "FTX+PMT")
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(FTX)
    
End Property

Friend Property Get SG1() As String    'C99
    
    'RFF-DTM-SG2
    '0130    RFF, Reference                      M1  RFF+<qualifier>:<identifier>
    '0140    DTM, Date/time/period               C1
    '0150    Segment group 2:  PAC-SG3           C10
    '    0160    PAC, Package                    M1
    '    0170    Segment group 3:  PCI-FTX       C999
    '        0180    PCI, Package identification M1
    '        0190    FTX, Free text              C1
        
    Dim strTemp As String
    
    'LRN
    strTemp = "RFF+ABE:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A3").Value)) & "'"
    
    AddFieldToDigiSign "A3", m_rstHeader, False
    
    SG1 = SG1 & TrimSegment(strTemp, "RFF+ABE")
    
    If m_enuDeclarationMode = enuOriginal Then
        'Details
        strTemp = "RFF+AKO:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C1").Value)) & "'"
            
        AddFieldToDigiSign "C1", m_rstHeader, False
            
        SG1 = SG1 & TrimSegment(strTemp, "RFF+AKO")
                    
        'Commercial Reference Number
        strTemp = "RFF+ABO:" & ReplaceSpecialCharacters(Left(FNullField(m_rstHeader.Fields("AC").Value), 22)) & "'"
        
        AddFieldToDigiSign "AC", m_rstHeader, False
        
        SG1 = SG1 & TrimSegment(strTemp, "RFF+ABO")
    End If
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG1)
    
End Property

Friend Property Get SG4() As String    'C5
    
    'TDT-TPL
    '0210    TDT, Transport information              M1  TDT+<qualifier>++<transport mode name code>+++++:::<identification name>:<nationality code>
    '0220    TPL, Transport placement                C1

    Dim strTemp As String
    
    'Identity of the means of transport at departure
    If Not (CheckRegimeType("K", "=") Or CheckRegimeType("J", "=")) Then
        strTemp = "TDT+13+++++++:::" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D5").Value)) & "'"
        
        AddFieldToDigiSign "D5", m_rstHeader, False
        
        SG4 = SG4 & TrimSegment(strTemp, "TDT+13")
    End If
    
    'Transport alternative to the border
    strTemp = "TDT+11++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D8").Value))
        
    AddFieldToDigiSign "D8", m_rstHeader, False
        
    'Nationality of the means of transport at the border
    strTemp = strTemp & "+++++::::" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D7").Value)) & "'"
                    
    AddFieldToDigiSign "D7", m_rstHeader, False
                    
    SG4 = SG4 & TrimSegment(strTemp, "TDT+11")
        
    'Internal transport alternative
    If m_enuDeclarationMode = enuOriginal Then
                    
        If LenB(Trim$(FNullField(m_rstHeader.Fields("D9").Value))) > 0 Then
            If Val(m_rstHeader.Fields("D9").Value) <> 0 Then
                strTemp = "TDT+1++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D9").Value)) & "'"
                
                AddFieldToDigiSign "D9", m_rstHeader, False
                
                SG4 = SG4 & TrimSegment(strTemp, "TDT+1")
            End If
        End If
    End If
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG4)
    
End Property

Friend Property Get SG5() As String    'C99

    'DOC-DTM-LOC
    '0240    DOC, Document/message details           M1  DOC+122+<total no. of loading lists>
    
    'Number of Loading lists
    If CheckRegimeType("H", "=") Or CheckRegimeType("I", "=") Then
            
        If LenB(Trim$(FNullField(m_rstHeader.Fields("A8").Value))) > 0 Then
            If Val(m_rstHeader.Fields("A8").Value) <> 0 Then
                SG5 = "DOC+122+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A8").Value)) & "'"
                
                AddFieldToDigiSign "A8", m_rstHeader, False
                
                SG5 = TrimSegment(SG5, "DOC+122")
            End If
        End If
    End If

    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG5)
    
    '0250    DTM, Date/time/period                   C2
    '0260    LOC, Place/location identification      C1

End Property

Friend Property Get SG6() As String    'C99

    'NAD-RFF-CTA-COM
    '0280    NAD, Name and address               M1  NAD+<Qualifier>+<party identifier>++<party name>+<street>+<city name>+:::<country sub-entity name>+<postal code>+<country code>
    '0290    RFF, Reference                      C10 RFF+AHP:<inscription number>
    '0300    [CTA], Contact information          C1  CTA++:<department or employee name>
    '0310    [COM], Communication contact        C5  COM+<contact value>:<communication identifier>
                            
    Dim strTemp As String
    Dim strPrefix As String
    
    Dim blnSendHandelaar As Boolean
    
    With m_rstHeaderHandelaars
        If .RecordCount <> 0 Then
        
            .MoveFirst
            
            Do While Not .EOF
                
                blnSendHandelaar = True
                
                Select Case FNullField(.Fields("XE").Value)
                
                    Case enuHeaderConsignee
                        strPrefix = "NAD+CN"
                
                    Case enuHeaderDeclarant
                        strPrefix = "NAD+AE"
                    
                    Case enuHeaderIntracommunautaireVerwerving
                        strPrefix = "NAD+GS"
                    
                    Case enuHeaderResponsibleRepresentative
                        strPrefix = "NAD+AG"
                    
                    Case Else
                        blnSendHandelaar = False
                        
                End Select
                
                If blnSendHandelaar Then
                
                    strTemp = strPrefix
                    
                    'Venture Number
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X1").Value))
                    
                    AddFieldToDigiSign "X1", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    strTemp = strTemp & "+"
                    
                    'Informant Signature
                    If FNullField(.Fields("XE").Value) = enuHeaderDeclarant Then
                        strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AA").Value))
                    End If
                    
                    'Name
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(Left(FNullField(.Fields("X2").Value), 35))
                                                    
                    If Len(FNullField(.Fields("X2").Value)) > 35 Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(Mid(FNullField(.Fields("X2").Value), 36, 35))
                    Else
                        strTemp = strTemp & ":"
                    End If
                        
                    AddFieldToDigiSign "X2", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                        
                    'Address 1 / Address 2
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X3").Value))
                    strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("X4").Value))
                    
                    AddFieldToDigiSign "X3", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    'Municipality
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X6").Value))
                    
                    AddFieldToDigiSign "X6", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    'State or Province
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X7").Value))
                    
                    AddFieldToDigiSign "X7", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    'Postal Code
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X5").Value))
                    
                    AddFieldToDigiSign "X5", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    'Country
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X8").Value))
                    
                    AddFieldToDigiSign "X8", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                     
                    
                    If FNullField(.Fields("XE").Value) = enuHeaderConsignee And Not VentureNumbersAreSame(FNullField(.Fields("X1").Value), "CN", enuImport) Then
                        strPrefix = strTemp
                    End If
                    
                    strTemp = strTemp & "'"
                    
                    SG6 = SG6 & TrimSegment(strTemp, strPrefix)
                    
                    If TrimSegment(strTemp, strPrefix) <> "" Then
                    
                        If FNullField(.Fields("XE").Value) = enuHeaderDeclarant Then
                            'Registration Number
                            strTemp = "RFF+AHP:" & ReplaceSpecialCharacters(FNullField(.Fields("XF").Value))
                            strTemp = strTemp & "-" & ReplaceSpecialCharacters(FNullField(.Fields("XD").Value)) & "'"
                            
                            AddFieldToDigiSign "XF", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                            AddFieldToDigiSign "XD", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                            
                            SG6 = SG6 & TrimSegment(strTemp, "RFF+AHP")
                                            
                            'Repertory Number
                            strTemp = "RFF+ABE:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AB").Value)) & "'"
                            
                            AddFieldToDigiSign "AB", m_rstHeader, False
                            
                            SG6 = SG6 & TrimSegment(strTemp, "RFF+ABE")
                        End If
                        
                        If FNullField(.Fields("XE").Value) = enuHeaderDeclarant Or _
                           FNullField(.Fields("XE").Value) = enuHeaderConsignee Then
                            'Contact Person
                            strTemp = "CTA++:" & ReplaceSpecialCharacters(FNullField(.Fields("X9").Value)) & "'"
                            
                            AddFieldToDigiSign "X9", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                            
                            SG6 = SG6 & TrimSegment(strTemp, "CTA")
                            
                            If FNullField(.Fields("X9").Value) <> vbNullString Then
                                'Telephone Number
                                strTemp = "COM+" & ReplaceSpecialCharacters(FNullField(.Fields("XA").Value)) & ":TE'"
                                
                                AddFieldToDigiSign "XA", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                                
                                SG6 = SG6 & TrimSegment(strTemp, "COM+:TE")
                                
                                'FAX Number
                                strTemp = "COM+" & ReplaceSpecialCharacters(FNullField(.Fields("XB").Value)) & ":FX'"
                                
                                AddFieldToDigiSign "XB", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                                
                                SG6 = SG6 & TrimSegment(strTemp, "COM+:FX")
                                                            
                                'E-Mail Address
                                strTemp = "COM+" & ReplaceSpecialCharacters(FNullField(.Fields("XC").Value)) & ":EM'"
                                
                                AddFieldToDigiSign "XC", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                                
                                SG6 = SG6 & TrimSegment(strTemp, "COM+:EM")
                            End If
                        End If
                    
                    End If
                    
                End If
                
                .MoveNext
            
            Loop
        
        End If
    End With
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG6)
    
End Property

Friend Property Get SG7() As String    'C1
    
    'TOD-LOC-FTX
    '0330    TOD, Terms of delivery or transport        M1  TOD+6++<code>
    '0340    LOC, Place/location identification         C2  LOC+1+:::<location>
    '0350    FTX, Free text                             C1
    
    Dim strTemp As String
    
    If CheckRegimeType("J", "<>") And CheckRegimeType("K", "<>") Then
        'Terms of delivery - Type
        strTemp = "TOD+6++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C2").Value)) & "'"
                                    
        AddFieldToDigiSign "C2", m_rstHeader, False
                                    
        SG7 = SG7 & TrimSegment(strTemp, "TOD+6")
        
        'Terms of delivery - Place
        If TrimSegment(strTemp, "TOD+6") <> "" Then
            strTemp = "LOC+1+:::" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C3").Value)) & "'"
            
            AddFieldToDigiSign "C3", m_rstHeader, False
            
            SG7 = SG7 & TrimSegment(strTemp, "LOC+1")
        End If
    End If
        
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG7)
    
End Property

Friend Property Get SG8() As String    'C99

    'MOA-SG9
    '0370    MOA, Monetary amount                    M1  MOA+<qualifier>:<amount>:<currency code>:1
    '0380    Segment group 9:  CUX-DTM               C1
    '    0390    CUX, Currencies                     M1  CUX+5:<currency code>:1++<rate>
    '    0400    DTM, Date/time/period               C1
    
    Dim strTemp As String
    
    If CheckRegimeType("J", "<>") Then
                    
        If LenB(Trim$(FNullField(m_rstHeader.Fields("C4").Value))) > 0 Then
            If Val(m_rstHeader.Fields("C4").Value) <> 0 Then
                'Customs Total Cost
                strTemp = "MOA+69:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C4").Value))
                
                AddFieldToDigiSign "C4", m_rstHeader, False
                
                'Currency Code
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C5").Value)) & ":1'"
                
                AddFieldToDigiSign "C5", m_rstHeader, False
                
                SG8 = SG8 & TrimSegment(strTemp, "MOA+69:::1")
            
                If TrimSegment(strTemp, "MOA+69:::1") <> "" Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C5").Value)) & ":1"
                    
                    'Exchange Rate
                    strTemp = strTemp & "++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C6").Value)) & "'"
                    
                    AddFieldToDigiSign "C6", m_rstHeader, False
                    
                    SG8 = SG8 & TrimSegment(strTemp, "CUX+5::1++")
                End If
            End If
        End If
                    
        If LenB(Trim$(FNullField(m_rstHeader.Fields("C9").Value))) > 0 Then
            If Val(FNullField(m_rstHeader.Fields("C9").Value)) <> 0 Then
                'Total VAT Cost
                strTemp = "MOA+72:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C9").Value))
                
                AddFieldToDigiSign "C9", m_rstHeader, False
                
                'Currency Code
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("CA").Value)) & ":1'"
                
                AddFieldToDigiSign "CA", m_rstHeader, False
                
                SG8 = SG8 & TrimSegment(strTemp, "MOA+72:::1")
                
                If TrimSegment(strTemp, "MOA+72:::1") <> "" Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("CA").Value)) & ":1"
                    
                    'Exchange Rate
                    strTemp = strTemp & "++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("CB").Value)) & "'"
                                    
                    AddFieldToDigiSign "CB", m_rstHeader, False
                                    
                    SG8 = SG8 & TrimSegment(strTemp, "CUX+5::1++")
                End If
            End If
        End If
    End If
        
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG8)
    
End Property


