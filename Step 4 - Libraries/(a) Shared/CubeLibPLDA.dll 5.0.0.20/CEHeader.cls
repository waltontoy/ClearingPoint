VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_rstHeaderHandelaars As ADODB.Recordset
Public m_rstHeader As ADODB.Recordset
Public m_rstHeaderZegels As ADODB.Recordset
Public m_rstHeaderZekerheid As ADODB.Recordset
Public m_rstHeaderTransitOffices As ADODB.Recordset
Public m_rstHeaderPLDAMessages As ADODB.Recordset

Public m_rstDetails As ADODB.Recordset

Public m_enuDeclarationMode As DECLARATION_MODE
Public m_lngHeaderSegments As Long

Private m_strSender As String
Private m_strIEReference As String
Private m_strRecipient As String
Private m_blnIsAnEcsDeclaration As Boolean
Private m_Code As String

Public Function CreateHeader(ByRef SegmentCount As Long, _
                             ByVal RecordsHeader As ADODB.Recordset, _
                             ByVal RecordsHeaderZegels As ADODB.Recordset, _
                             ByVal RecordsHeaderHandelaars As ADODB.Recordset, _
                             ByVal RecordsHeaderZekerheid As ADODB.Recordset, _
                             ByVal RecordsHeaderTransitOffices As ADODB.Recordset, _
                             ByVal RecordsHeaderPLDAMessages As ADODB.Recordset, _
                             ByVal RecordsDetails As ADODB.Recordset, _
                             ByVal DeclarationMode As DECLARATION_MODE, _
                             ByVal EDIUserName As String, _
                             ByVal LogTIN As String, _
                             ByVal Recipient As String, _
                             ByVal LastEDIRef, _
                             Optional ByVal Code As String) As String
                        
    Dim strMessage As String
    
    'Create clones of Header Recordsets
    Set m_rstHeader = RecordsHeader.Clone
    Set m_rstHeaderZegels = RecordsHeaderZegels.Clone
    Set m_rstHeaderZekerheid = RecordsHeaderZekerheid.Clone
    Set m_rstHeaderTransitOffices = RecordsHeaderTransitOffices.Clone
    Set m_rstHeaderHandelaars = RecordsHeaderHandelaars.Clone
    Set m_rstHeaderPLDAMessages = RecordsHeaderPLDAMessages.Clone
    
    m_blnIsAnEcsDeclaration = (Len(FNullField(m_rstHeader.Fields("AO").Value)) > 0)
    m_Code = Code
    
    'Create Clone of Detail Recordsets
    Set m_rstDetails = RecordsDetails.Clone
    
    m_lngHeaderSegments = 0
    m_enuDeclarationMode = DeclarationMode
    m_strSender = EDIUserName
    m_strRecipient = Recipient
    m_strIEReference = Left(LogTIN, 9) & Format(LastEDIRef, "00000")
    
    '*******************************************************************************************
    'Create the header
    '*******************************************************************************************
    strMessage = vbNullString
    strMessage = strMessage & UNB
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
    
    'Return Segment Count
    SegmentCount = SegmentCount + m_lngHeaderSegments
    
    'Close Clone Recordsets
    ADORecordsetClose m_rstHeader
    ADORecordsetClose m_rstHeaderZegels
    ADORecordsetClose m_rstHeaderZekerheid
    ADORecordsetClose m_rstHeaderTransitOffices
    ADORecordsetClose m_rstHeaderHandelaars
    ADORecordsetClose m_rstHeaderPLDAMessages
    
    ADORecordsetClose m_rstDetails
    
    'Return Created Header EDI Message String
    CreateHeader = strMessage
    
End Function

Friend Property Get UNB() As String
    
    'UNB = "UNB+UNOA:2+" & m_strEDIUserName & "+" & m_strRecipient & "+" & Format(Now, "yymmdd") & ":" & Format(Now, "hhmm") & "+" & Left(m_strLogTIN, 9) & Format(m_strLastEDIReference, "00000") & "'"
        
    UNB = "UNB+UNOC:3+" & ReplaceSpecialCharacters(m_strSender)           'Sender
    UNB = UNB & "+" & ReplaceSpecialCharacters(m_strRecipient)            'Recipient
    UNB = UNB & "+" & Format(Now, "yyMMdd")     'Date
    UNB = UNB & ":" & Format(Now, "hhmm")       'Time
    
    'Edwin - Oct 18
    'UNB = UNB & "+" & ReplaceSpecialCharacters(m_strIEReference) & "'"    'Interchange Reference
    UNB = UNB & "+196688++CUSDEC++++1" & "'"
    
End Property

Friend Property Get UNH() As String    'M1

    'Message header
    
    'Edwin - Oct 18
    'UNH = "UNH+1+CUSDEC:D:04A:UN:EXPORT'"
    UNH = "UNH+196688+CUSDEC:D:04A:UN:EXPORT+196688'"
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(UNH)
    
End Property

'New segments were added/uncommented/modified for Transit Declaration Type.
Friend Property Get BGM() As String    'M1

    'Beginning of message
    'BGM+:::<type of declaration>+<MRN if for modification>+<modification mode>

    Dim A1 As String
    Dim A2 As String
    Dim AD As String
    
    Dim blnValue As Boolean
    
    With m_rstHeader
        'Declaration Type Part 1
        A1 = ReplaceSpecialCharacters(FNullField(.Fields("A1").Value))
        
        'Declaration Type Part 2
        A2 = ReplaceSpecialCharacters(FNullField(.Fields("A2").Value))
        
        'Declaration Type Part 3
        AD = ReplaceSpecialCharacters(FNullField(.Fields("AD").Value))
        
        '************************************************************
        ' Declaration Type
        '************************************************************
        If CheckRegimeType("A", "=") = True Or _
            CheckRegimeType("B", "=") = True Or _
            CheckRegimeType("C", "=") = True Or _
            CheckRegimeType("D", "=") = True Or _
            CheckRegimeType("E", "=") = True Then
            
            'Digital Signature for Fields A1 and A2 - Type A to E
            AddFieldToDigiSign "A1", m_rstHeader, False
            AddFieldToDigiSign "A2", m_rstHeader, False
            
            BGM = "BGM+:::" & A1 & A2
            BGM = BGM & "+"
            
        ElseIf CheckRegimeType("F", "=") = True Or _
            CheckRegimeType("G", "=") = True Then
            
            'Digital Signature for Field AD - Type F and G
            AddFieldToDigiSign "AD", m_rstHeader, False
            
            BGM = "BGM+:::" & Space(2) & Space(1) & AD
            BGM = BGM & "+"
        
        Else
            'Digital Signature for Fields A1 and A2 - Other values
            AddFieldToDigiSign "A1", m_rstHeader, False
            AddFieldToDigiSign "A2", m_rstHeader, False
            
            BGM = "BGM+:::" & A1 & A2
            BGM = BGM & "+"
        End If
        '************************************************************
            
        'MRN if Declaration is Ammendment
        If m_enuDeclarationMode <> enuOriginal Then
            BGM = BGM & ReplaceSpecialCharacters(FNullField(.Fields("MRN").Value))
        End If
        
        'Language of the Declaration
        BGM = BGM & ":" & ReplaceSpecialCharacters(FNullField(.Fields("A9").Value))
        
        AddFieldToDigiSign "A9", m_rstHeader, False
        
        'Declaration Mode
        BGM = BGM & "+" & ReplaceSpecialCharacters(m_enuDeclarationMode) & "'"
        
        BGM = TrimSegment(BGM, "BGM")
    End With
    
    'Add BGM Segment Count to HeaderSegments
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(BGM)
    
End Property

Friend Property Get CST() As String    'C1
'Customs status of goods
End Property

 'New segments were added/uncommented/modified for Transit Declaration.
Friend Property Get LOC() As String    'C99

    'Place/location identification
    'LOC+<Qualifier>+F<1>
    
    Dim strTemp As String
    Dim lngGroupCtr As Long
    
    ' CSCLP-884 - Start
    Dim blnLOC35Added As Boolean
    
    blnLOC35Added = False
    ' CSCLP-884 - End
    
    If CheckRegimeType("G", "<>") And _
       CheckRegimeType("F", "<>") Then
        
        'Location of Goods
        strTemp = "LOC+43+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("D4").Value)) & "'"
        
        AddFieldToDigiSign "D4", m_rstHeader, False
        
        LOC = LOC & TrimSegment(strTemp, "LOC+43")
        
        'Country of Sending/Export/Dispatch/Loading
        If CheckRegimeType("D", "<>") = True Then
            If Not m_blnIsAnEcsDeclaration Then
                
                ' CSCLP-884 - Start
                If Not blnLOC35Added Then
                    strTemp = "LOC+35+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("DB").Value)) & "'"
                    
                    AddFieldToDigiSign "DB", m_rstHeader, False
                    
                    LOC = LOC & TrimSegment(strTemp, "LOC+35")
                    
                    blnLOC35Added = True
                End If
                ' CSCLP-884 - End
            End If
        End If
        
        'CSCLP-606
        If CheckRegimeType("D", "=") = True Then
            If m_rstDetails.Fields("N1").Value = "31" And m_rstDetails.Fields("N2").Value = "71" Then
            
                ' CSCLP-884 - Start
                If Not blnLOC35Added Then
                    strTemp = "LOC+35+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("DB").Value)) & "'"
                    LOC = LOC & TrimSegment(strTemp, "LOC+35")
                    
                    blnLOC35Added = True
                End If
                ' CSCLP-884 - End
            End If
        End If
        '>>
        
        
        'Office of Entrance/Termination
        strTemp = "LOC+42+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A7").Value)) & "'"
        
        AddFieldToDigiSign "A7", m_rstHeader, False
        
        LOC = LOC & TrimSegment(strTemp, "LOC+42")
    End If
    
    If m_blnIsAnEcsDeclaration And Not RoutingBoxEmpty("AE", m_Code) Then
        strTemp = "LOC+9+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("DB").Value)) & "'"
                
        AddFieldToDigiSign "DB", m_rstHeader, False
        
        LOC = LOC & TrimSegment(strTemp, "LOC+9")
    'sept15
    ElseIf m_blnIsAnEcsDeclaration And RoutingBoxEmpty("AE", m_Code) Then
    
        ' CSCLP-884 - Start
        If Not blnLOC35Added Then
            strTemp = "LOC+35+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("DB").Value)) & "'"
                    
            AddFieldToDigiSign "DB", m_rstHeader, False
            
            LOC = LOC & TrimSegment(strTemp, "LOC+35")
            
            blnLOC35Added = True
        End If
        ' CSCLP-884 - End
    End If
    
    Dim blnRegimeTypeF As Boolean
    
    blnRegimeTypeF = CheckRegimeType("F", "=")
    
    
        With m_rstHeaderTransitOffices
            If .RecordCount > 0 Then
                
                .Sort = "ORDINAL"
                .MoveFirst
                
                Do While Not .EOF
                    lngGroupCtr = lngGroupCtr + 1
                    
                    'Transit Offices
                    If LenB(Trim$(FNullField(.Fields("AE").Value))) > 0 Then
                        If m_blnIsAnEcsDeclaration And Not RoutingBoxEmpty("AE", m_Code) Then
                            strTemp = "LOC+92+" & ReplaceSpecialCharacters(FNullField(.Fields("AE").Value)) & "'"
                            
                            LOC = LOC & TrimSegment(strTemp, "LOC+92")
                        ElseIf blnRegimeTypeF Then
                            strTemp = "LOC+50+" & ReplaceSpecialCharacters(FNullField(.Fields("AE").Value)) & "'"
                            
                            LOC = LOC & TrimSegment(strTemp, "LOC+50")
                        End If
                    End If
                    
                    'Transit Offices
                    If LenB(Trim$(FNullField(.Fields("AF").Value))) > 0 Then
                            If m_blnIsAnEcsDeclaration And Not RoutingBoxEmpty("AF", m_Code) Then
                                strTemp = "LOC+92+" & ReplaceSpecialCharacters(FNullField(.Fields("AF").Value)) & "'"
                                
                                LOC = LOC & TrimSegment(strTemp, "LOC+92")
                            ElseIf blnRegimeTypeF Then
                                strTemp = "LOC+50+" & ReplaceSpecialCharacters(FNullField(.Fields("AF").Value)) & "'"
                                
                                LOC = LOC & TrimSegment(strTemp, "LOC+50")
                            End If
                    End If
                    
                    If m_blnIsAnEcsDeclaration Or blnRegimeTypeF Then
                        AddFieldToDigiSign "AE", m_rstHeaderTransitOffices, False, , , .Fields("Ordinal").Value
                    End If
                    
                    If UCase(Trim(.Fields("AG").Value)) = "E" Then Exit Do
                    
                    .MoveNext
                Loop
                
            End If
        End With
    
    If CheckRegimeType("F", "=") = True Then
        With m_rstHeader
            'Country of Ultimate Destination
            strTemp = "LOC+36+" & ReplaceSpecialCharacters(FNullField(.Fields("AN").Value)) & "'"
            
            AddFieldToDigiSign "AN", m_rstHeader, False
            
            LOC = LOC & TrimSegment(strTemp, "LOC+36")
            
            '***********************************************************************
            'Location of Goods
            '***********************************************************************
            If LenB(Trim$(FNullField(.Fields("DC").Value))) > 0 Then
                strTemp = "LOC+14+" & ReplaceSpecialCharacters(FNullField(.Fields("DC").Value)) & "'"
                LOC = LOC & TrimSegment(strTemp, "LOC+14")
            End If
            
            If LenB(Trim$(FNullField(.Fields("DD").Value))) > 0 Then
                strTemp = "LOC+14+" & ReplaceSpecialCharacters(FNullField(.Fields("DD").Value)) & "'"
                LOC = LOC & TrimSegment(strTemp, "LOC+14")
            End If
            
            If LenB(Trim$(FNullField(.Fields("DE").Value))) > 0 Then
                strTemp = "LOC+14+" & ReplaceSpecialCharacters(FNullField(.Fields("DE").Value)) & "'"
                LOC = LOC & TrimSegment(strTemp, "LOC+14")
            End If
            '***********************************************************************
            
            'Digital Signature for Fields DC, DD, and DE
            If LenB(Trim$(FNullField(.Fields("DC").Value))) > 0 Or _
               LenB(Trim$(FNullField(.Fields("DD").Value))) > 0 Or _
               LenB(Trim$(FNullField(.Fields("DE").Value))) > 0 Then
                AddFieldToDigiSign "DC", m_rstHeader, False
            End If
            
            'Port of Loading
            strTemp = "LOC+9+" & ReplaceSpecialCharacters(FNullField(.Fields("AM").Value)) & "'"
            
            AddFieldToDigiSign "AM", m_rstHeader, False
            
            LOC = LOC & TrimSegment(strTemp, "LOC+9")
        End With
    End If
    
    'Customs Departure Office
    strTemp = "LOC+118+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AH").Value)) & "'"
    
    AddFieldToDigiSign "AH", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+118")
    
    'Place of Declaration
    strTemp = "LOC+91+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A5").Value)) & "'"
    
    AddFieldToDigiSign "A5", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+91")
    
    'Customs Validation Office
    strTemp = "LOC+41+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A6").Value)) & "'"
    
    AddFieldToDigiSign "A6", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+41")
    
    'Customs Office Destination
    strTemp = "LOC+45+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AI").Value)) & "'"
    
    AddFieldToDigiSign "AI", m_rstHeader, False
    
    LOC = LOC & TrimSegment(strTemp, "LOC+45")
    
    'Identification of Warehouse Box 49
    If m_rstDetails.RecordCount > 0 Then
        m_rstDetails.MoveFirst
        
        If LenB(Trim$(FNullField(m_rstDetails.Fields("M3").Value))) > 0 Or _
            LenB(Trim$(FNullField(m_rstDetails.Fields("M4").Value))) > 0 Or _
            LenB(Trim$(FNullField(m_rstDetails.Fields("M5").Value))) > 0 Then
            
            '************************************************************************
            'COMMENTED AS PER CSCLP-505 INSTRUCTION <It should be noted that this existing code
            'does not call AddFieldToDigiSign anymore ???
            '************************************************************************
            'strTemp = "LOC+18+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value)) & Space(1)
            
            'strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M4").Value)) & Space(1)
            
            'strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value))
            
            'strTemp = strTemp & Space(1)
            
            'strTemp = strTemp & ":"
            
            'strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value))
            
            'strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value)) & "'"
            
            'LOC = LOC & TrimSegment(strTemp, "LOC+18+  ::+")
            '************************************************************************
            
            'strTemp = "LOC+18+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value)) & Space(1)
            '
            ''AddFieldToDigiSign "M3", m_rstDetails, False
            '
            'strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value))
            '
            ''AddFieldToDigiSign "M5", m_rstDetails, False
            '
            'strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value)) & Space(1)
            '
            ''AddFieldToDigiSign "M3", m_rstDetails, False
            '
            'strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M4").Value)) & Space(1)
            '
            ''AddFieldToDigiSign "M4", m_rstDetails, False
            '
            'strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value)) & "'"
            '
            ''AddFieldToDigiSign "M5", m_rstDetails, False
            '
            'LOC = LOC & TrimSegment(strTemp, "LOC+18+::  ")
        End If
    End If
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(LOC)
    
End Property

Friend Property Get DTM() As String    'C15

    'Date/time/period
    'DTM+<Qualifier>:F<yyyymmdd>:102
    
    Dim strTemp As String
    
    'Document/message date/time
    strTemp = "DTM+137:" & Format(Now, "yyyyMMdd") & ":102'"
    DTM = DTM & TrimSegment(strTemp, "DTM+137::102")
    
    
    'Registration Date of Previous Customs Declaration
    strTemp = "DTM+254:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A4").Value)) & ":102'"
    
    AddFieldToDigiSign "A4", m_rstHeader, False
    
    DTM = DTM & TrimSegment(strTemp, "DTM+254::102")
    
    'Date Limit
    If CheckRegimeType("F", "=") = True Then
        strTemp = "DTM+268:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AL").Value)) & ":102'"
        
        AddFieldToDigiSign "AL", m_rstHeader, False
        
        DTM = DTM & TrimSegment(strTemp, "DTM+268::102")
    End If
        
    ' Expiry Date
    strTemp = "DTM+36:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AL").Value)) & ":102'"
    
    'AddFieldToDigiSign "AL", m_rstHeader, False
    
    DTM = DTM & TrimSegment(strTemp, "DTM+36::102")
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(DTM)
    
End Property

Friend Property Get GEI() As String    'C25
'Processing information
'GEI+2+<Processing indicator>:ZZZ:147
End Property

Friend Property Get FII() As String    'C1
'Financial institution information
'FII+AE+<acct49 + customs office> :<TIN>
End Property

Friend Property Get MEA() As String    'C5
'Measurements
End Property

Friend Property Get EQD() As String  'C999
'Equipment details
End Property

Friend Property Get SEL() As String    'C999
    
    'Seal number
    'SEL+0+<seal issuer>:::<seal number>+<seal condition code>
        
    Dim strTemp As String
    
    With m_rstHeaderZegels
    
        If .RecordCount > 0 Then
            .Sort = "ORDINAL"
            
            .MoveFirst
            
            Do While Not .EOF
                'Seal Code
                strTemp = "SEL+0+" & ReplaceSpecialCharacters(FNullField(.Fields("E2").Value))
                
                AddFieldToDigiSign "E2", m_rstHeaderZegels, False, , , .Fields("Ordinal").Value
                
                'Seal Identifier/Number
                strTemp = strTemp & ":::" & ReplaceSpecialCharacters(FNullField(.Fields("E1").Value))
                strTemp = strTemp & "'"
                
                AddFieldToDigiSign "E1", m_rstHeaderZegels, False, , , .Fields("Ordinal").Value
                
                SEL = SEL & TrimSegment(strTemp, "SEL+0")
                
                If UCase(Trim(.Fields("E3").Value)) = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
    
    End With
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SEL)
    
End Property

'Added Capacity of declarant, Account number, Investigation Type, and Cancellation reason.
Friend Property Get FTX() As String    'C99

    'Free text
    'FTX+<Qualifier>++<payment code>+<payment literal>
        
    Dim strTemp As String
    
    If CheckRegimeType("B", "<>") = True And _
       CheckRegimeType("D", "<>") = True And _
       CheckRegimeType("G", "<>") = True And _
       CheckRegimeType("F", "<>") = True Then

        'Terms of Payment (Nature of Transaction)
        strTemp = "FTX+AAB++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C7").Value)) & "'"
        
        AddFieldToDigiSign "C7", m_rstHeader, False
        
        'FTX = FTX & TrimSegment(strTemp, "FTX+AAB++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C7").Value)))
        FTX = FTX & TrimSegment(strTemp, "FTX+AAB++")
    End If
    
    If CheckRegimeType("F", "=") = True Then
        With m_rstHeaderHandelaars
            .MoveFirst
            Do While Not .EOF
                If CLng(Trim(.Fields("XE").Value)) = 1 Then
                    'Capacity of the Declarant (Free Text) - XG(1)
                    strTemp = "FTX+ACP++" & ReplaceSpecialCharacters(FNullField(.Fields("XG").Value)) & "'"
                    
                    AddFieldToDigiSign "XG", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    FTX = FTX & TrimSegment(strTemp, "FTX+ACP++")
                    
                    'Account Number
                    strTemp = "FTX+ADE++" & ReplaceSpecialCharacters(FNullField(.Fields("XH").Value)) & "'"
                    
                    AddFieldToDigiSign "XH", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    FTX = FTX & TrimSegment(strTemp, "FTX+ADE+")
                    
                    Exit Do
                End If
                .MoveNext
            Loop
            
            'Investigation Type
            strTemp = "FTX+ADO++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AK").Value)) & "'"
            
            AddFieldToDigiSign "AK", m_rstHeader, False
            
            FTX = FTX & TrimSegment(strTemp, "FTX+ADO++")
        End With
    End If
    
    'Cancellation Reason
    If m_enuDeclarationMode <> enuCancellation Then
        With m_rstHeaderPLDAMessages
            If .RecordCount > 0 Then
                .MoveFirst

                If Not .EOF Then
                    strTemp = "FTX+ACD++" & ReplaceSpecialCharacters(FNullField(.Fields("Message_Reason").Value)) & "'"
                    FTX = FTX & TrimSegment(strTemp, "FTX+ACD++")
                End If
            End If
        End With
    End If
    
    ' Specific Circumstance Indicator
    strTemp = "FTX+SPT++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AO").Value)) & "'"
            
    'AddFieldToDigiSign "AO", m_rstHeader, False
    
    FTX = FTX & TrimSegment(strTemp, "FTX+SPT++")
    
    ' Representative Status Code - Declarant Only
    ' Check DTm+36 and FTX+CUS
'    If Not (FNullField(m_rstHeader.Fields("AL").Value) = "" And _
'        FNullField(m_rstHeader.Fields("AK").Value = "")) Then
            
        With m_rstHeaderHandelaars
            If .RecordCount <> 0 Then
                .MoveFirst
                
                Do While Not .EOF
                    
                    ' Representative Status Code - Declarant Only
                    If FNullField(.Fields("XE").Value) = enuExportHeaderDeclarant Then
                        strTemp = "FTX+AHN++" & ReplaceSpecialCharacters(FNullField(m_rstHeaderHandelaars.Fields("XF").Value)) & "'"
                
                        'AddFieldToDigiSign "XF", m_rstHeader, False
                        
                        FTX = FTX & TrimSegment(strTemp, "FTX+AHN++")
                        
                        Exit Do
                    End If
    
                    .MoveNext
                Loop
            End If
        End With
    'End If

    ' Control Result Code (Box D)
    strTemp = "FTX+CUS++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AK").Value)) & "'"
            
    'AddFieldToDigiSign "AK", m_rstHeader, False
    
    FTX = FTX & TrimSegment(strTemp, "FTX+CUS++")
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(FTX)
    
End Property

'Changes to Guarantee Fields manner of sending - Edwin Dec18
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
    Dim lngGroupCtr As Long
    
    'LRN
    strTemp = "RFF+ABE:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A3").Value)) & "'"
    
    AddFieldToDigiSign "A3", m_rstHeader, False
    
    SG1 = SG1 & TrimSegment(strTemp, "RFF+ABE")
        
    'Commercial Reference
    strTemp = "RFF+ABO:" & ReplaceSpecialCharacters(Left(FNullField(m_rstHeader.Fields("AC").Value), 22)) & "'"
    
    AddFieldToDigiSign "AC", m_rstHeader, False
    
    SG1 = SG1 & TrimSegment(strTemp, "RFF+ABO")
    
    
    With m_rstHeaderZekerheid
        If .RecordCount > 0 Then
            
            .Sort = "ORDINAL"
            .MoveFirst
            
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
                
                'Guarantee Type
                If LenB(Trim$(FNullField(.Fields("E4").Value))) > 0 Then
                    strTemp = "RFF+ABL:" & ReplaceSpecialCharacters(FNullField(.Fields("E4").Value)) & "'"
                    
                    AddFieldToDigiSign "E4", m_rstHeaderZekerheid, False, , , lngGroupCtr
                    
                    SG1 = SG1 & TrimSegment(strTemp, "RFF+ABL")
                End If
                
                If LenB(Trim$(FNullField(.Fields("E4").Value))) > 0 Then
                    Select Case CLng(Trim(FNullField(.Fields("E4").Value)))
                        Case "0", "1", "2", "4", "9"
                            strTemp = "PAC+1+++"
                            
                            'Guarantee Access
                            If Trim(FNullField(.Fields("E5").Value)) <> "" And Not IsNull(.Fields("E5").Value) Then
                                strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("E5").Value))
                                
                                AddFieldToDigiSign "E5", m_rstHeaderZekerheid, False, , , lngGroupCtr
                            End If
                            
                            'Guarantee Number
                            If Trim(FNullField(.Fields("E6").Value)) <> "" And Not IsNull(.Fields("E6").Value) Then
                                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("E6").Value))
                            
                                AddFieldToDigiSign "E6", m_rstHeaderZekerheid, False, , , lngGroupCtr
                            End If
                            
                            'Guarantee ( Other )
                            If Trim(FNullField(.Fields("E7").Value)) <> "" And Not IsNull(.Fields("E7").Value) Then
                                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(.Fields("E7").Value))
                            
                                AddFieldToDigiSign "E7", m_rstHeaderZekerheid, False, , , lngGroupCtr
                            End If
                            
                            strTemp = strTemp & "'"
                            SG1 = SG1 & TrimSegment(strTemp, "PAC+1+++:::")
                    End Select
                End If
                
                If CheckRegimeType("F", "=") = True Or _
                   CheckRegimeType("G", "=") = True Then
                    '***********************************************************************
                    'Other Parties (Only EVA or Visegrad countries are possible)
                    '***********************************************************************
                    strTemp = "PCI+28+"
                    
                    If Trim(FNullField(.Fields("E9").Value)) <> "" And Not IsNull(.Fields("E9").Value) Then
                        strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("E9").Value))
                    End If
                    
                    If Trim(FNullField(.Fields("EA").Value)) <> "" And Not IsNull(.Fields("EA").Value) Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("EA").Value))
                    End If
                    
                    If Trim(FNullField(.Fields("EB").Value)) <> "" And Not IsNull(.Fields("EB").Value) Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("EB").Value))
                    End If
                    
                    If Trim(FNullField(.Fields("EC").Value)) <> "" And Not IsNull(.Fields("EC").Value) Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("EC").Value))
                    End If
                    
                    If Trim(FNullField(.Fields("ED").Value)) <> "" And Not IsNull(.Fields("ED").Value) Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("ED").Value))
                    End If
                    
                    If Trim(FNullField(.Fields("EE").Value)) <> "" And Not IsNull(.Fields("EE").Value) Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("EE").Value))
                    End If
                    
                    AddFieldToDigiSign "E9", m_rstHeaderZekerheid, False, , , lngGroupCtr
                    
                    strTemp = strTemp & "'"
                    SG1 = SG1 & TrimSegment(strTemp, "PCI+28+:::::")
                    '***********************************************************************
                    
                    'Validity of the Guarantee
                    If Trim(FNullField(.Fields("E8").Value)) <> "" And Not IsNull(.Fields("E8").Value) Then
                        strTemp = "PCI+19+" & ReplaceSpecialCharacters(FNullField(.Fields("E8").Value)) & "'"
                        SG1 = SG1 & TrimSegment(strTemp, "PCI+19+")
                        
                        AddFieldToDigiSign "E8", m_rstHeaderZekerheid, False, , , lngGroupCtr
                    End If
                End If
                
                If .Fields("EF").Value = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
    End With
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG1)
    
End Property

Friend Property Get SG4() As String    'C5

    'TDT-TPL
    '0210    TDT, Transport information              M1  TDT+<qualifier>++<transport mode name code>+++++:::<identification name>:<nationality code>
    '0220    TPL, Transport placement                C1
        
    Dim strTemp As String
    
    With m_rstHeader
        If CheckRegimeType("B", "<>") And CheckRegimeType("D", "<>") And CheckRegimeType("G", "<>") Then
            'Identity of Means of Transport at Departure
            strTemp = "TDT+12+++++++:::" & ReplaceSpecialCharacters(FNullField(.Fields("D5").Value))
            
            AddFieldToDigiSign "D5", m_rstHeader, False
            
            'Departure Nationality
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("AJ").Value))
            
            AddFieldToDigiSign "AJ", m_rstHeader, False
            
            strTemp = strTemp & "'"
            
            SG4 = SG4 & TrimSegment(strTemp, "TDT+12")
        End If
        
        strTemp = "TDT+11++"
        
        'Transport Alternative to the Border
        If CheckRegimeType("D", "<>") And CheckRegimeType("G", "<>") Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("D8").Value))
            
            AddFieldToDigiSign "D8", m_rstHeader, False
        End If
        
        strTemp = strTemp & "+"
        
        'Border Type
        If Trim(FNullField(.Fields("DF").Value)) <> "" And Not IsNull(.Fields("DF").Value) Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("DF").Value))
            
            AddFieldToDigiSign "DF", m_rstHeader, False
        End If
        
        strTemp = strTemp & "++++:::"
        
        If CheckRegimeType("B", "<>") And CheckRegimeType("D", "<>") And CheckRegimeType("G", "<>") Then
            'Identity of means of transport crossing border
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("D6").Value))
            
            AddFieldToDigiSign "D6", m_rstHeader, False
            
            'Nationality at crossing border
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("D7").Value))
            
            AddFieldToDigiSign "D7", m_rstHeader, False
        End If
        
        strTemp = strTemp & "'"
        SG4 = SG4 & TrimSegment(strTemp, "TDT+11")
        
        'Internal Transport Alternative
        If CheckRegimeType("D", "<>") And CheckRegimeType("G", "<>") Then
            If LenB(Trim$(FNullField(.Fields("D9").Value))) > 0 And FNullField(.Fields("D9").Value) <> 0 Then
                strTemp = "TDT+1++" & ReplaceSpecialCharacters(FNullField(.Fields("D9").Value)) & "'"
                
                AddFieldToDigiSign "D9", m_rstHeader, False
                
                SG4 = SG4 & TrimSegment(strTemp, "TDT+1")
            End If
        End If
    End With

    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG4)
    
End Property

Friend Property Get SG5() As String    'C99

    'DOC-DTM-LOC
    '0240    DOC, Document/message details           M1  DOC+122+<total no. of loading lists>
    
    'Number of Loading Lists
    If LenB(Trim(FNullField(m_rstHeader.Fields("A8").Value))) > 0 And _
        FNullField(m_rstHeader.Fields("A8").Value) <> 0 Then
        
        SG5 = "DOC+122+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A8").Value)) & "'"
        
        AddFieldToDigiSign "A8", m_rstHeader, False
        
        SG5 = TrimSegment(SG5, "DOC+122")
        
        m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG5)
    End If
    
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
    Dim strVentureNum As String
    Dim strX1 As String
    
    With m_rstHeaderHandelaars
        If .RecordCount <> 0 Then
            .MoveFirst
            
            Do While Not .EOF
                Select Case FNullField(.Fields("XE").Value)
                    Case enuExportHeaderConsignee
                        strPrefix = "NAD+CN"
                    Case enuExportHeaderDeclarant
                        strPrefix = "NAD+AE"
                    Case enuExportHeaderBenificiary
                        strPrefix = "NAD+BE"
                    Case enuExportHeaderResponsibleRepresentative
                        strPrefix = "NAD+AG"
                    Case enuexportHeaderExporter
                        strPrefix = "NAD+EX"
                End Select
                
                strTemp = strPrefix
                   
                strX1 = FNullField(.Fields("X1").Value)
                
                'Venture Number
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("X1").Value))
                
                AddFieldToDigiSign "X1", m_rstHeaderHandelaars, False, , .Fields("XE").Value
                
                strTemp = strTemp & "+"
                
                'Informant Signature - Declarant Only
                If FNullField(.Fields("XE").Value) = enuExportHeaderDeclarant Then
                    m_strDeclarantVentureNumber = strX1
                    strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AA").Value))
                    
                    AddFieldToDigiSign "AA", m_rstHeader, False
                End If
'CSCLP-603
'                If (FNullField(.Fields("XE").Value) = enuExportHeaderDeclarant) Or _
'                   (strX1 = "" Or Left$(strX1, 5) = "BE999" Or Left$(strX1, 2) <> "BE") Then
                    If (FNullField(.Fields("XE").Value) = enuExportHeaderDeclarant) Or _
                       (FNullField(.Fields("XE").Value) = enuExportHeaderBenificiary) Or _
                       (FNullField(.Fields("XE").Value) = enuExportHeaderConsignee And VentureNumbersAreSame(FNullField(strX1), "CN", enuExport) And CheckRegimeType("G", "<>")) Or _
                       (FNullField(.Fields("XE").Value) = enuExportHeaderResponsibleRepresentative And CheckRegimeType("A", "=")) Or _
                       (FNullField(.Fields("XE").Value) = enuexportHeaderExporter And VentureNumbersAreSame(FNullField(strX1), "EX", enuExport)) Then
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
                    End If
'                End If
                
                If FNullField(.Fields("XE").Value) = enuExportHeaderConsignee And Not VentureNumbersAreSame(FNullField(strX1), "CN", enuExport) Then
                    strPrefix = strTemp
                End If
                
                strTemp = strTemp & "'"
                SG6 = SG6 & TrimSegment(strTemp, strPrefix)
                
                If TrimSegment(strTemp, strPrefix) <> "" Then
                
                    If FNullField(.Fields("XE").Value) = enuExportHeaderDeclarant Then
                        'Registration Number
                        strTemp = "RFF+AHP:" & ReplaceSpecialCharacters(FNullField(.Fields("XF").Value)) & "-"
                        strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("XD").Value)) & "'"
                        
                        AddFieldToDigiSign "XF", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                        AddFieldToDigiSign "XD", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                        
                        SG6 = SG6 & TrimSegment(strTemp, "RFF+AHP")
                        
                        'Repertory Number
                        strTemp = "RFF+ABE:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("AB").Value)) & "'"
                        
                        AddFieldToDigiSign "AB", m_rstHeader, False
                        
                        SG6 = SG6 & TrimSegment(strTemp, "RFF+ABE")
                    End If
                    
                    'Contact Person
                    strTemp = "CTA++:" & ReplaceSpecialCharacters(FNullField(.Fields("X9").Value)) & "'"
                    
                    AddFieldToDigiSign "X9", m_rstHeaderHandelaars, False, , FNullField(.Fields("XE").Value)
                    
                    SG6 = SG6 & TrimSegment(strTemp, "CTA")
                    
                    If FNullField(.Fields("X9").Value) <> "" Then
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
                
                .MoveNext
            Loop
        End If
    End With
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG6)
    
'COMMENTED: Glenn 9/6/2006
'    'AE
'    With m_rstHeader
'
'        strTemp = "NAD+AE+" & !X1                       'Venture Number
'        strTemp = strTemp & "+" & !AA                   'Informant's Signature
'
'        If UCase(Left(!X1, 2)) <> "BE" Then
'
'            strTemp = strTemp & "+" & Left(!X2, 35)     'Name
'            strTemp = strTemp & ":" & Mid(!X2, 36, 35)  'Name
'            strTemp = strTemp & "+" & !X3               'Address 1
'            strTemp = strTemp & ":" & !X4               'Address 2
'            strTemp = strTemp & "+" & !X6               'Municipality
'            strTemp = strTemp & "+" & !X7               'State/Province
'            strTemp = strTemp & "+" & !X5               'Postal Code
'            strTemp = strTemp & "+" & !X8               'Country
'
'        End If
'
'        strTemp = strTemp & "'"
'        SG6 = SG6 & TrimSegment(strTemp, "NAD+AE++" & !AA)
'
'        If TrimSegment(strTemp, "NAD+AE++" & !AA) <> "" Then
'
'            strTemp = "RFF+AHP:" & !XD & "'"            'Registration Number
'            SG6 = SG6 & TrimSegment(strTemp, "RFF+AHP")
'
'            strTemp = "RFF+ABE:" & !AB & "'"            'Repertory Number
'            SG6 = SG6 & TrimSegment(strTemp, "RFF+ABE")
'
'            strTemp = "CTA++:" & !X9 & "'"              'Contact Person
'            SG6 = SG6 & TrimSegment(strTemp, "CTA")
'
'            If TrimSegment(strTemp, "CTA") <> "" Then
'
'                strTemp = "COM+" & !XA & ":TE'"         'Telephone Number
'                SG6 = SG6 & TrimSegment(strTemp, "COM+:TE")
'
'                strTemp = "COM+" & !XB & ":FX'"         'FAX
'                SG6 = SG6 & TrimSegment(strTemp, "COM+:FX")
'
'                strTemp = "COM+" & !XC & ":EM'"         'E-Mail
'                SG6 = SG6 & TrimSegment(strTemp, "COM+:EM")
'
'            End If
'
'        End If
'
'    End With
'
'    'CN
'    With m_rstHeader
'
'        If VentureNumbersAreSame(!W1, "CN") And CheckRegimeType("G", "<>") Then '
'
'            If !W1 <> "" Then
'
'                strTemp = "NAD+CN+" & !W1                   'Venture Number
'
'            Else
'
'                strTemp = "NAD+CN+0"
'
'            End If
'
'            If UCase(Left(!W1, 2)) <> "BE" Then
'
'                strTemp = strTemp & "++" & Left(!W2, 35)    'Name
'                strTemp = strTemp & ":" & Mid(!W2, 36, 35)  'Name
'                strTemp = strTemp & "+" & !W3               'Address 1
'                strTemp = strTemp & ":" & !W4               'Address 2
'                strTemp = strTemp & "+" & !W6               'Municipality
'                strTemp = strTemp & "+" & !W7               'State/Province
'                strTemp = strTemp & "+" & !W5               'Postal Code
'                strTemp = strTemp & "+" & !W8               'Country
'
'            End If
'
'            strTemp = strTemp & "'"
'            SG6 = SG6 & TrimSegment(strTemp, "NAD+CN+0")
'
'            '*Wrong : ito ba talaga kukunin na value??
'            'SG6 = SG6 & "RFF+ABE:" & m_rstHeader!W1 & "'"
'
'            If TrimSegment(strTemp, "NAD+CN+0") <> "" Then
'
'                strTemp = "CTA++:" & !W9 & "'"              'Contact Person
'                SG6 = SG6 & TrimSegment(strTemp, "CTA")
'
'                If TrimSegment(strTemp, "CTA") <> "" Then
'
'                    strTemp = "COM+" & !WA & ":TE'"         'Telephone Number
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:TE")
'
'                    strTemp = "COM+" & !WB & ":FX'"         'FAX
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:FX")
'
'                    strTemp = "COM+" & !WC & ":EM'"         'E-Mail
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:EM")
'
'                End If
'
'            End If
'
'        End If
'
'    End With
'
'    'AG
'    With m_rstHeader
'
'        If CheckRegimeType("A", "=") Then
'
'            strTemp = "NAD+AG+" & Space(35)
'            strTemp = strTemp & "++" & Left(!G2, 35)    'Name
'            strTemp = strTemp & ":" & Mid(!G2, 36, 35)  'Name
'            strTemp = strTemp & "+" & !G3               'Address 1
'            strTemp = strTemp & ":" & !G4               'Address 2
'            strTemp = strTemp & "+" & !G6               'Municipality
'            strTemp = strTemp & "+" & !G7               'State/Province
'            strTemp = strTemp & "+" & !G5               'Postal Code
'            strTemp = strTemp & "+" & !G8               'Country
'            strTemp = strTemp & "'"
'            SG6 = SG6 & TrimSegment(strTemp, "NAD+AG+" & Space(35))
'
'            If TrimSegment(strTemp, "NAD+AG+" & Space(35)) <> "" Then
'
'                strTemp = "CTA++:" & !G9 & "'"      'Contact Person
'                SG6 = SG6 & TrimSegment(strTemp, "CTA")
'
'                If TrimSegment(strTemp, "CTA") <> "" Then
'
'                    strTemp = "COM+" & !GA & ":TE'"     'Telephone Number
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:TE")
'
'                    strTemp = "COM+" & !GB & ":FX'"     'FAX
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:FX")
'
'                    strTemp = "COM+" & !GC & ":EM'"     'E-Mail
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:EM")
'
'                End If
'
'            End If
'
'        End If
'
'    End With
'
'    'BE
'    With m_rstHeader
'
'        If VentureNumbersAreSame(!H1, "CN") And CheckRegimeType("G", "<>") Then
'
'            If !H1 <> "" Then
'
'                strTemp = "NAD+BE+" & !H1                   'Venture Number
'
'            Else
'
'                strTemp = "NAD+BE+0"
'
'            End If
'
'            If UCase(Left(!H1, 2)) <> "BE" Then
'
'                strTemp = strTemp & "++" & Left(!H2, 35)    'Name
'                strTemp = strTemp & ":" & Mid(!H2, 36, 35)  'Name
'                strTemp = strTemp & "+" & !H3               'Address 1
'                strTemp = strTemp & ":" & !H4               'Address 2
'                strTemp = strTemp & "+" & !H6               'Municipality
'                strTemp = strTemp & "+" & !H7               'State/Province
'                strTemp = strTemp & "+" & !H5               'Postal Code
'                strTemp = strTemp & "+" & !H8               'Country
'
'            End If
'
'            strTemp = strTemp & "'"
'            SG6 = SG6 & TrimSegment(strTemp, "NAD+BE+0")
'
'            If TrimSegment(strTemp, "NAD+BE+0") <> "" Then
'
'                strTemp = "CTA++:" & !H9 & "'"              'Contact Person
'                SG6 = SG6 & TrimSegment(strTemp, "CTA")
'
'                If TrimSegment(strTemp, "CTA") <> "" Then
'
'                    strTemp = "COM+" & !HA & ":TE'"         'Telephone Number
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:TE")
'
'                    strTemp = "COM+" & !HB & ":FX'"         'FAX
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:FX")
'
'                    strTemp = "COM+" & !HC & ":EM'"         'E-Mail
'                    SG6 = SG6 & TrimSegment(strTemp, "COM+:EM")
'
'                End If
'
'            End If
'
'        End If
'
'    End With
'
'Commented by Migs 06/06/2006 Mali ung box na kinukuhanan nito... wala pang consignor/exporter boxes sa Combined codisheet!
'**************************************************************************************************************************
'**************************************************************************************************************************
'    '*Wrong : to be replaced with procedures that checks if EX in all tariff lines are the same,
'    '         if not then put 00200
'    'Uncommented by Migs 04/05/2006 meron nang proc... ung VentureNumbersAreSame
'    If VentureNumbersAreSame(m_rstHeader!H1, "EX") Then
'
'    '    'EX
'        If m_rstHeader!H1 <> "" Then
'            SG6 = SG6 + "NAD+EX+" & m_rstHeader!H1
'            m_lngHeaderSegments = m_lngHeaderSegments + 1
'            If UCase(Right(m_rstHeader!H1, 2)) <> "BE" Then
'                'Commented by Migs 04/21/2006 replaced with the following code lines
'                'SG6 = SG6 & "++" & m_rstHeader!H2 & "+" & m_rstHeader!H3 & ":" & m_rstHeader!H4 & "+" & m_rstHeader!H6 & "+" & m_rstHeader!H7 & "+" & m_rstHeader!H5 & "+" & m_rstHeader!H8
'                strTemp = ""
'                If m_rstHeader!H2 <> "" Or m_rstHeader!H3 <> "" Or m_rstHeader!H4 <> "" _
'                   Or m_rstHeader!H6 <> "" Or m_rstHeader!H7 <> "" Or m_rstHeader!H5 <> "" _
'                   Or m_rstHeader!H8 <> "" Then
'
'                    strTemp = strTemp & "++" & m_rstHeader!H2
'                    If m_rstHeader!H3 <> "" Or m_rstHeader!H4 <> "" Or m_rstHeader!H6 <> "" _
'                       Or m_rstHeader!H7 <> "" Or m_rstHeader!H5 <> "" Or m_rstHeader!H8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstHeader!H3 & IIf(m_rstHeader!H4 = "", "", ":" & m_rstHeader!H4)
'                        If m_rstHeader!H6 <> "" Or m_rstHeader!H7 <> "" Or m_rstHeader!H5 <> "" _
'                           Or m_rstHeader!H8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstHeader!H6
'                            If m_rstHeader!H7 <> "" Or m_rstHeader!H5 <> "" Or m_rstHeader!H8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstHeader!H7
'                                If m_rstHeader!H5 <> "" Or m_rstHeader!H8 <> "" Then
'
'                                    strTemp = strTemp & "+" & m_rstHeader!H5
'                                    If m_rstHeader!H8 <> "" Then strTemp = strTemp & "+" & m_rstHeader!H8
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'            SG6 = SG6 & strTemp & "'"
'
''Commented by Migs 05/10/2006 temporarily commented dahil unsure pa kung pang NAD+AE lang ito
''            If m_rstHeader!H1 <> "" Then
''                SG6 = SG6 & "RFF+ABE:" & m_rstHeader!H1 & "'"
''                m_lngHeaderSegments = m_lngHeaderSegments + 1
''            End If
'
'            If m_rstHeader!H9 <> "" Then
'                SG6 = SG6 & "CTA+::" & m_rstHeader!H9 & "'"
'                m_lngHeaderSegments = m_lngHeaderSegments + 1
'            End If
'
'            If m_rstHeader!HA <> "" Then
'                SG6 = SG6 & "COM+" & m_rstHeader!HA & ":TE'"
'                m_lngHeaderSegments = m_lngHeaderSegments + 1
'            End If
'
'            If m_rstHeader!HB <> "" Then
'                SG6 = SG6 & "COM+" & m_rstHeader!HB & ":FX'"
'                m_lngHeaderSegments = m_lngHeaderSegments + 1
'            End If
'
'            If m_rstHeader!HC <> "" Then
'                SG6 = SG6 & "COM+" & m_rstHeader!HC & ":EM'"
'                m_lngHeaderSegments = m_lngHeaderSegments + 1
'            End If
'        End If
'    End If
'**************************************************************************************************************************
'**************************************************************************************************************************
                
'Commented by Migs 04/05/2006 commented muna...
    '*Wrong
'    SG6 = SG6 & "NAD+NBB+" & m_rstHeader!W1 & "'"
'    m_lngHeaderSegments = m_lngHeaderSegments + 1
'    SG6 = SG6 & "CTA+::" & m_rstHeader!W9 & "'"
'    m_lngHeaderSegments = m_lngHeaderSegments + 1
'    SG6 = SG6 & "COM+" & m_rstHeader!WA & ":TE'"
'    m_lngHeaderSegments = m_lngHeaderSegments + 1
'    SG6 = SG6 & "COM+" & m_rstHeader!WB & ":FX'"
'    m_lngHeaderSegments = m_lngHeaderSegments + 1
'    SG6 = SG6 & "COM+" & m_rstHeader!WC & ":EM'"
'    m_lngHeaderSegments = m_lngHeaderSegments + 1

End Property

Friend Property Get SG7() As String    'C1

    'TOD-LOC-FTX
    '0330    TOD, Terms of delivery or transport        M1  TOD+6++<code>
    '0340    LOC, Place/location identification         C2  LOC+1+:::<location>
    '0350    FTX, Free text                             C1
    
    Dim strTemp As String
    
    If CheckRegimeType("A", "=") Or _
       CheckRegimeType("C", "=") Or _
       CheckRegimeType("D", "=") Or _
       CheckRegimeType("E", "=") Then
    
        'Delivery Condition Type
        strTemp = "TOD+6++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C2").Value)) & "'"
        
        AddFieldToDigiSign "C2", m_rstHeader, False
        
        SG7 = SG7 & TrimSegment(strTemp, "TOD+6")
        
        'Delivery Condition Place
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
    
    'For testing of sending Exchange Rate in Type B and D declarations - Edwin Dec18
    'If CheckRegimeType("B", "<>") And _
       CheckRegimeType("D", "<>") And _
       CheckRegimeType("G", "<>") And _
       CheckRegimeType("F", "<>") Then
    
    If CheckRegimeType("G", "<>") And _
       CheckRegimeType("F", "<>") Then
    
        If LenB(Trim$(FNullField(m_rstHeader.Fields("C4").Value))) > 0 Then
            'Removed this checking for CSCLP-791
            'If Val(FNullField(m_rstHeader.Fields("C4").Value)) <> 0 Then
                'Invoice Total Amount
                strTemp = "MOA+39:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C4").Value))
                
                AddFieldToDigiSign "C4", m_rstHeader, False
                
                'Currency Code
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C5").Value)) & ":1'"
                
                AddFieldToDigiSign "C5", m_rstHeader, False
                
                SG8 = SG8 & TrimSegment(strTemp, "MOA+39::1")
                
                If TrimSegment(strTemp, "MOA+39::1") <> "" Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C5").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("C6").Value)) & "'"
                    
                    AddFieldToDigiSign "C6", m_rstHeader, False
                    
                    SG8 = SG8 & TrimSegment(strTemp, "CUX+5::1++")

                End If
            'End If
        End If
    End If
    
    m_lngHeaderSegments = m_lngHeaderSegments + CountSegments(SG8)
        
End Property


Private Sub Class_Initialize()
    m_blnIsAnEcsDeclaration = False
End Sub
