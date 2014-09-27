VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_enuDeclarationMode As DECLARATION_MODE
Public m_rstDetailsHandelaars As ADODB.Recordset
Public m_rstDetailsBerekeningsEenheden As ADODB.Recordset

Public m_lngDetailSegments As Long
Public m_lngDetailNo As Long
Public m_rstHeader As ADODB.Recordset
Public m_rstDetails As ADODB.Recordset
Public m_rstDetailsBijzondere As ADODB.Recordset
Public m_rstDetailsContainer As ADODB.Recordset
Public m_rstDetailsDocumenten As ADODB.Recordset
Public m_rstDetailsZelf As ADODB.Recordset

Public Function CreateDetail(ByVal RecordsHeader As ADODB.Recordset, _
                             ByRef RecordsDetails As ADODB.Recordset, _
                             ByRef RecordsDetailsBijzondere As ADODB.Recordset, _
                             ByRef RecordsDetailsContainer As ADODB.Recordset, _
                             ByRef RecordsDetailsDocumenten As ADODB.Recordset, _
                             ByRef RecordsDetailsZelf As ADODB.Recordset, _
                             ByRef RecordsDetailsHandelaars As ADODB.Recordset, _
                             ByRef RecordsDetailsBerekeningsEenheden As ADODB.Recordset, _
                             ByRef SegmentCount As Long, _
                             ByVal DetailNo As Long, _
                             ByVal DecMode As DECLARATION_MODE) As String
    
    Dim strMessage As String
    
    m_lngDetailSegments = 0
    m_lngDetailNo = DetailNo
    m_enuDeclarationMode = DecMode
    
    Set m_rstHeader = RecordsHeader.Clone
    Set m_rstDetails = RecordsDetails.Clone
    Set m_rstDetailsBijzondere = RecordsDetailsBijzondere.Clone
    Set m_rstDetailsContainer = RecordsDetailsContainer.Clone
    Set m_rstDetailsDocumenten = RecordsDetailsDocumenten.Clone
    Set m_rstDetailsZelf = RecordsDetailsZelf.Clone
    Set m_rstDetailsHandelaars = RecordsDetailsHandelaars.Clone
    Set m_rstDetailsBerekeningsEenheden = RecordsDetailsBerekeningsEenheden.Clone
    
    'm_rstDetails.Filter = adFilterNone
    'm_rstDetails.Filter = "Detail = " & DetailNo
    
    '*********************************************************************************************
    'Set Filters for all recordset
    '*********************************************************************************************
    m_rstDetails.Filter = adFilterNone
    m_rstDetails.Filter = "Detail = " & DetailNo
    
    m_rstDetailsBijzondere.Filter = adFilterNone
    m_rstDetailsBijzondere.Filter = "Detail = " & DetailNo
    
    m_rstDetailsContainer.Filter = adFilterNone
    m_rstDetailsContainer.Filter = "Detail = " & DetailNo
    
    m_rstDetailsDocumenten.Filter = adFilterNone
    m_rstDetailsDocumenten.Filter = "Detail = " & DetailNo
    
    m_rstDetailsZelf.Filter = adFilterNone
    m_rstDetailsZelf.Filter = "Detail = " & DetailNo
    
    m_rstDetailsHandelaars.Filter = adFilterNone
    m_rstDetailsHandelaars.Filter = "Detail = " & DetailNo
    
    m_rstDetailsBerekeningsEenheden.Filter = adFilterNone
    m_rstDetailsBerekeningsEenheden.Filter = "Detail = " & DetailNo
    '*********************************************************************************************
    
    If m_lngDetailNo > 1 Then
        strMessage = ""
    Else
        strMessage = UNS
    End If
    
    strMessage = strMessage & SG10
    strMessage = strMessage & SG30
    
    SegmentCount = SegmentCount + m_lngDetailSegments
    
    CreateDetail = strMessage
    
End Function
Friend Property Get UNS() As String    'M1
    
    'Section control
    'UNS+D
    
    UNS = "UNS+D'"
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(UNS)

End Property

Friend Property Get SG10() As String   'C999
'DMS-DTM-MEA-SG11-SG13-SG14-SG16-SG18-SG19-SG21
'0430    DMS, Document/message summary               M1
'0440    DTM, Date/time/period                       C1
'0450    MEA, Measurements                           C5
'0460    Segment group 11: MOA-SG12                  C25
'    0470    MOA, Monetary amount                    M1
'    0480    Segment group 12: CUX-DTM               C1
'        0490    CUX, Currencies                     M1
'        0500    DTM, Date/time/period               C1
'0510    Segment group 13: TOD-LOC-FTX               C1
'    0520    TOD, Terms of delivery or transport     M1
'    0530    LOC, Place/location identification      C2
'    0540    FTX, Free text                          C1
'0550    Segment group 14: NAD-SG15                  C10
'    0560    NAD, Name and address                   M1
'    0570    Segment group 15: DOC-DTM-LOC           C10
'        0580    DOC, Document/message details       M1
'        0590    DTM, Date/time/period               C2
'        0600    LOC, Place/location identification  C1
'0610    Segment group 16: PAC-SG17                  C10
'    0620    PAC, Package                            M1
'    0630    Segment group 17: PCI-FTX-RFF           C999
'        0640    PCI, Package identification         M1
'        0650    FTX, Free text                      C1
'        0660    RFF, Reference                      C1
'0670    Segment group 18: PYT-MOA-PCD-FTX           C5
'    0680    PYT, Payment terms                      M1
'    0690    MOA, Monetary amount                    C1
'    0700    PCD, Percentage details                 C1
'    0710    FTX, Free text                          C1
'0720    Segment group 19: ALC-RTE-MOA-PCD-QTY-SG20  C10
'    0730    ALC, Allowance or charge                M1
'    0740    RTE, Rate details                       C1
'    0750    MOA, Monetary amount                    C1
'    0760    PCD, Percentage details                 C1
'    0770    QTY, Quantity                           C1
'    0780    Segment group 20: CUX-DTM               C1
'        0790    CUX, Currencies                     M1
'        0800    DTM, Date/time/period               C1
'0810    Segment group 21:                           C999
'LIN-PIA-QTY-PRI-PCD-MEA-QVR-MOA-NAD-GIR-GEI-LOC-SG22-SG23-SG25-SG26-SG27-SG28
'    0820    LIN, Line item                          M1
'    0830    PIA, Additional product id              C1
'    0840    QTY, Quantity                           C1
'    0850    PRI, Price details                      C2
'    0860    PCD, Percentage details                 C1
'    0870    MEA, Measurements                       C15
'    0880    QVR, Quantity variances                 C5
'    0890    MOA, Monetary amount                    C15
'    0900    NAD, Name and address                   C10
'    0910    GIR, Related identification numbers     C9999
'    0920    GEI, Processing information             C99
'    0930    LOC, Place/location identification      C1
'    0940    Segment group 22: DOC-DTM               C10
'        0950    DOC, Document/message details       M1
'        0960    DTM, Date/time/period               C2
'    0970    Segment group 23:                       C10
'    ALC-RTE-MOA-PCD-QTY-SG24
'        0980    ALC, Allowance or charge            M1
'        0990    RTE, Rate details                   C1
'        1000    MOA, Monetary amount                C1
'        1010    PCD, Percentage details             C1
'        1020    QTY, Quantity                       C1
'        1030    Segment group 24: CUX-DTM           C1
'            1040    CUX, Currencies                 M1
'            1050    DTM, Date/time/period           C1
'    1060    Segment group 25: TOD-LOC               C1
'        1070    TOD, Terms of delivery or transport M1
'        1080    LOC, Place/location identification  C2
'    1090    Segment group 26: PYT-MOA-PCD-FTX       C5
'        1100    PYT, Payment terms                  M1
'        1110    MOA, Monetary amount                C1
'        1120    PCD, Percentage details             C1
'        1130    FTX, Free text                      C1
'    1140    Segment group 27: IMD-FTX               C5
'        1150    IMD, Item description               M1
'        1160    FTX, Free text                      C1
'    1170    Segment group 28: PAC-SG29              C10
'        1180    PAC, Package                        M1
'        1190    Segment group 29: PCI-FTX-RFF       C999
'            1200    PCI, Package identification     M1
'            1210    FTX, Free text                  C1
'            1220    RFF, Reference                  C1
End Property

Friend Property Get SG30() As String   'C99999
'CST-FTX-LOC-DTM-MEA-NAD-TDT-SG31-SG33-SG35-SG37-SG38-SG39-SG40-SG41-SG42-SG44
'1240    CST, Customs status of goods                    M1  CST+<detail no.>+<1>+<2>+<3>+<4>+<5>
    
    'Property modified by Migs 04/20/2006 added empty string validations (if <field> <> "")
    
'    Dim strTemp As String
'    strTemp = ""  'Added by Migs 04/20/2006
        
'>>> Migs 07-13-2006 <<<
'Commented these. Moved to new property 'CST' -----------------
'    SG30 = "CST+" & m_lngDetailNo
'    If UCase(m_rstDetails!N4) <> "J" Then
'        'Commented by Migs 04/20/2006 replaced with the following code lines
'        'SG30 = SG30 & "+" & m_rstDetails!L1 & "+" & m_rstDetails!L2 & "+" & m_rstDetails!L3 & "+" & m_rstDetails!L4 & "+" & m_rstDetails!L5 & "-" & m_rstDetails!L6
'        If m_rstDetails!L1 <> "" Or m_rstDetails!L2 <> "" Or m_rstDetails!L3 <> "" Or _
'           m_rstDetails!L4 <> "" Or m_rstDetails!L5 <> "" Or m_rstDetails!L6 <> "" Then
'
'            strTemp = strTemp & "+" & m_rstDetails!L1
'            If m_rstDetails!L2 <> "" Or m_rstDetails!L3 <> "" Or m_rstDetails!L4 <> "" Or _
'               m_rstDetails!L5 <> "" Or m_rstDetails!L6 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!L2
'                If m_rstDetails!L3 <> "" Or m_rstDetails!L4 <> "" Or m_rstDetails!L5 <> "" Or _
'                   m_rstDetails!L6 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!L3
'                    If m_rstDetails!L4 <> "" Or m_rstDetails!L5 <> "" Or m_rstDetails!L6 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!L4
'                        If m_rstDetails!L5 <> "" And m_rstDetails!L6 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!L5 & "-" & m_rstDetails!L6    'Modified by Migs 04/29/2006
'                            'If m_rstDetails!L6 <> "" Then strTemp = strTemp & "+" & m_rstDetails!L6
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'    SG30 = SG30 & strTemp & "'"
'    m_lngDetailSegments = m_lngDetailSegments + 1
'    strTemp = ""
'--------------------------------------------------------------

'1250    FTX, Free text                                  C9  FTX+<qualifier>++<code>+<description>+<language code>
    
'>>> Migs 07-13-2006 <<<
'Commented these. Moved to new property 'FTX' -----------------
'    'Goods
'    If m_rstDetails!L8 <> "" Or m_rstHeader!A9 <> "" Then
'        SG30 = SG30 & "FTX+AAA+++" & m_rstDetails!L8 & IIf(m_rstHeader!A9 = "", "", "+" & m_rstHeader!A9) & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'Additional Information
'    m_rstDetailsBijzondere.Filter = adFilterNone
'    m_rstDetailsBijzondere.Filter = "Detail = " & m_lngDetailNo
'    If m_rstDetailsBijzondere.RecordCount > 0 Then
'        m_rstDetailsBijzondere.MoveFirst
'        Do While m_rstDetailsBijzondere.EOF = False
'            strTemp = IIf(IsNull(m_rstDetailsBijzondere!P2), "", Trim(m_rstDetailsBijzondere!P2))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "FTX+ACB++" & m_rstDetailsBijzondere!P1 & "+" & m_rstDetailsBijzondere!P2 & _
'                       IIf(m_rstHeader!A9 = "", "", "+" & m_rstHeader!A9) & "'"   'Modified by Migs 04/20/2006 added IIF statement
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsBijzondere!P4), "", Trim(m_rstDetailsBijzondere!P4))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "FTX+ACB++" & m_rstDetailsBijzondere!P3 & "+" & m_rstDetailsBijzondere!P4 & _
'                       IIf(m_rstHeader!A9 = "", "", "+" & m_rstHeader!A9) & "'"   'Modified by Migs 04/20/2006 added IIF statement
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            m_rstDetailsBijzondere.MoveNext
'        Loop
'    End If
'
'    'Customs Clearance Instruction
'    If m_rstDetails!N4 = "I" Or m_rstDetails!N4 = "H" Or m_rstDetails!N4 = "K" Then
'        If m_rstDetails!O1 <> "" Then
'            SG30 = SG30 & "FTX+CCI++" & m_rstDetails!O1 & "'"   'Modified by Migs 06/01/2006 changed back to CCI 'Modified by Migs 04/24/2006 changed CCI to CII
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'
'    'Terms of payment
'    If m_rstDetails!N9 <> "" And m_rstDetails!NA <> "" Then
'        SG30 = SG30 & "FTX+AAB++" & m_rstDetails!N9 & m_rstDetails!NA & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'Financieel
'    If m_rstDetails!T2 <> "" Or m_rstDetails!T1 <> "" Then
'        SG30 = SG30 & "FTX+PMT++" & m_rstDetails!T2 & IIf(m_rstDetails!T1 = "", "", "+" & m_rstDetails!T1)
'        SG30 = SG30 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!N4 <> "" Then
'        SG30 = SG30 & "FTX+CUS++" & m_rstDetails!N4 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'--------------------------------------------------------------

'1260    LOC, Place/location identification              C25 LOC+<qualifier>+<identification>
    
'>>> Migs 07-13-2006 <<<
'Commented these. Moved to new property 'LOC' -----------------
'    If m_rstDetails!NB <> "" Then
'        SG30 = SG30 & "LOC+27+" & m_rstDetails!NB & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!N4 <> "K" And m_rstDetails!N4 <> "J" Then
'        If m_rstDetails!N7 <> "" Then
'            SG30 = SG30 & "LOC+28+" & m_rstDetails!N7 & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'
'        If m_rstDetails!N8 <> "" Then
'            SG30 = SG30 & "LOC+47+" & m_rstDetails!N8 & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'
'    If m_rstDetails!M3 <> "" And m_rstDetails!M4 <> "" And m_rstDetails!M5 <> "" Then
'        'Migs 06-14-2006 replaced this code line
'        'SG30 = SG30 & "LOC+18+" & m_rstDetails!M3 & "-" & m_rstDetails!M4 & "-" & m_rstDetails!M5 & "'"   'Modified by Migs 06/07/2006 placed "-"    'Modified by Migs 04/28/2006
'        SG30 = SG30 & "LOC+18+" & m_rstDetails!M3 & m_rstDetails!M4 & m_rstDetails!M5 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'--------------------------------------------------------------

'1270    DTM, Date/time/period                           C5
'1280    MEA, Measurements                               C20 MEA+<qualifier>+<measurement code>+<value code>:<measurement>
    
'>>> Migs 07-13-2006 <<<
'Commented these. Moved to new property 'MEA' -----------------
'    If m_rstDetails!L9 <> "" Then
'        SG30 = SG30 & "MEA+WT+AAB+KGM:" & m_rstDetails!L9 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!LA <> "" Then
'        SG30 = SG30 & "MEA+WT+AAA+KGM:" & m_rstDetails!LA & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!M1 <> "" Or m_rstDetails!M2 <> "" Then
'        SG30 = SG30 & "MEA+AAE++" & m_rstDetails!M1 & IIf(m_rstDetails!M2 = "", "", ":" & m_rstDetails!M2) & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If Val(m_rstDetails!T8) <> 0 And Trim(m_rstDetails!T7) <> "" Then
'        SG30 = SG30 & "MEA+AAF++" & m_rstDetails!T7 & ":" & m_rstDetails!T8 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    If Val(m_rstDetails!TA) <> 0 And Trim(m_rstDetails!T9) <> "" Then
'        SG30 = SG30 & "MEA+AAF++" & m_rstDetails!T9 & ":" & m_rstDetails!TA & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    If Val(m_rstDetails!TC) <> 0 And Trim(m_rstDetails!TB) <> "" Then
'        SG30 = SG30 & "MEA+AAF++" & m_rstDetails!TB & ":" & m_rstDetails!TC & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'--------------------------------------------------------------

'1290    NAD, Name and address                           C5  NAD+<qualifier>+<identifier>++<party name>+<street>+<city name>+:::<country sub-entity name>+<postal code>+<country code>

'>>> Migs 07-13-2006 <<<
'Commented these. Moved to new property 'NAD' -----------------
'    'Modified by Migs 06/01/2006 added condition for header venture number (W1)
'    strTemp = ""
'    'SG30 = SG30 & "NAD+CN+" & IIf(m_rstHeader!W1 = "" And (Trim(m_rstDetails!V1) = "" Or Trim(m_rstDetails!V1) = "0"), "000000000", Trim(m_rstDetails!V1))
'    If m_rstHeader!W1 = "" Then
'        strTemp = "NAD+CN+" & IIf(Trim(m_rstDetails!V1) = "" Or Trim(m_rstDetails!V1) = "0", "000000000", Trim(m_rstDetails!V1))
'        m_lngDetailSegments = m_lngDetailSegments + 1
'        If UCase(Left(Trim(m_rstDetails!V1), 2)) <> "BE" Then
'    'Commented by Migs 04/21/2006 replaced with the following code lines
'    '        SG30 = SG30 & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2) & _
'    '            "+" & m_rstDetails!V3 & ":" & m_rstDetails!V4 & "+" & m_rstDetails!V6 & "+" & m_rstDetails!V7 & "+" & m_rstDetails!V5 & "+" & m_rstDetails!V8
'            'strTemp = ""
'            If m_rstDetails!V2 <> "" Or m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" _
'               Or m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'               Or m_rstDetails!V8 <> "" Then
'
'                strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'                If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'                   Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!V3 & IIf(m_rstDetails!V4 = "", "", ":" & m_rstDetails!V4)
'                    If m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'                       Or m_rstDetails!V8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!V6
'                        If m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!V7
'                            If m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!V5
'                                If m_rstDetails!V8 <> "" Then strTemp = strTemp & "+" & m_rstDetails!V8
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'    SG30 = SG30 & strTemp & IIf(strTemp = "", "", "'")
'
'    If m_rstDetails!Y1 <> "" Then
'        SG30 = SG30 & "NAD+AG+" & Trim(m_rstDetails!Y1)
''Commented by Migs 04/21/2006 replaced with the following code lines
''        SG30 = SG30 & "++" & IIf(Len(m_rstDetails!Y2) > 35, Left(m_rstDetails!Y2, 35) & ":" & Mid(m_rstDetails!Y2, 36, 35), m_rstDetails!Y2) & _
''            "+" & m_rstDetails!Y3 & ":" & m_rstDetails!Y4 & "+" & m_rstDetails!Y6 & "+" & m_rstDetails!Y7 & "+" & m_rstDetails!Y5 & "+" & m_rstDetails!Y8
'        strTemp = ""
'        If m_rstDetails!Y2 <> "" Or m_rstDetails!Y3 <> "" Or m_rstDetails!Y4 <> "" _
'           Or m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'           Or m_rstDetails!Y8 <> "" Then
'
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'            If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'               Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!Y3 & IIf(m_rstDetails!Y4 = "", "", ":" & m_rstDetails!Y4)
'                If m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'                   Or m_rstDetails!Y8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!Y6
'                    If m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!Y7
'                        If m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!Y5
'                            If m_rstDetails!Y8 <> "" Then strTemp = strTemp & "+" & m_rstDetails!Y8
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        SG30 = SG30 & strTemp & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!I1 <> "" Then
'        SG30 = SG30 & "NAD+GS+" & m_rstDetails!I1
''Commented by Migs 04/21/2006 replaced with the following code lines
''        SG30 = SG30 & "++" & IIf(Len(m_rstDetails!I2) > 35, Left(m_rstDetails!I2, 35) & ":" & Mid(m_rstDetails!I2, 36, 35), m_rstDetails!I2) & _
''            "+" & m_rstDetails!I3 & ":" & m_rstDetails!I4 & "+" & m_rstDetails!I6 & "+" & m_rstDetails!I7 & "+" & m_rstDetails!I5 & "+" & m_rstDetails!I8
'        strTemp = ""
'        If m_rstDetails!I2 <> "" Or m_rstDetails!I3 <> "" Or m_rstDetails!I4 <> "" _
'           Or m_rstDetails!I6 <> "" Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" _
'           Or m_rstDetails!I8 <> "" Then
'
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!I2) > 35, Left(m_rstDetails!I2, 35) & ":" & Mid(m_rstDetails!I2, 36, 35), m_rstDetails!I2)
'            If m_rstDetails!I3 <> "" Or m_rstDetails!I4 <> "" Or m_rstDetails!I6 <> "" _
'               Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!I3 & IIf(m_rstDetails!I4 = "", "", ":" & m_rstDetails!I4)
'                If m_rstDetails!I6 <> "" Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" _
'                   Or m_rstDetails!I8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!I6
'                    If m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!I7
'                        If m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!I5
'                            If m_rstDetails!I8 <> "" Then strTemp = strTemp & "+" & m_rstDetails!I8
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        SG30 = SG30 & strTemp & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!Z1 <> "" Then
'        SG30 = SG30 & "NAD+WD+" & m_rstDetails!Z1
''Modified by Migs 04/28/2006 replaced Z8 with M5
''Commented by Migs 04/21/2006 replaced with the following code lines
''        SG30 = SG30 & "++" & IIf(Len(m_rstDetails!Z2) > 35, Left(m_rstDetails!Z2, 35) & ":" & Mid(m_rstDetails!Z2, 36, 35), m_rstDetails!Z2) & _
''            "+" & m_rstDetails!Z3 & ":" & m_rstDetails!Z4 & "+" & m_rstDetails!Z6 & "+" & m_rstDetails!Z7 & "+" & m_rstDetails!Z5 & "+" & m_rstDetails!M5
'        strTemp = ""
'        If m_rstDetails!Z2 <> "" Or m_rstDetails!Z3 <> "" Or m_rstDetails!Z4 <> "" _
'           Or m_rstDetails!Z6 <> "" Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" _
'           Or m_rstDetails!M5 <> "" Then
'
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!Z2) > 35, Left(m_rstDetails!Z2, 35) & ":" & Mid(m_rstDetails!Z2, 36, 35), m_rstDetails!Z2)
'            If m_rstDetails!Z3 <> "" Or m_rstDetails!Z4 <> "" Or m_rstDetails!Z6 <> "" _
'               Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!Z3 & IIf(m_rstDetails!Z4 = "", "", ":" & m_rstDetails!Z4)
'                If m_rstDetails!Z6 <> "" Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" _
'                   Or m_rstDetails!M5 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!Z6
'                    If m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!Z7
'                        If m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!Z5
'                            If m_rstDetails!M5 <> "" Then strTemp = strTemp & "+" & m_rstDetails!M5
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        SG30 = SG30 & strTemp & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    strTemp = ""  'Added by Migs 04/21/2006 re-initialize strTemp
'--------------------------------------------------------------

'1300    TDT, Transport information                      C9

'***
'1310    Segment group 31: PAC-SG32                      C10
'    1320    PAC, Package                                M1  PAC+<no. of packages>++<package type>
'    1330    Segment group 32: PCI-FTX                   C999
'        1340    PCI, Package identification             M1  PCI+28+<marks and number>
'        1350    FTX, Free text                          C1

'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG31' ----------------
'    'Modified by Migs 04/27/2006 changed OR to AND
'    If m_rstDetails!S2 <> "" And m_rstDetails!S1 <> "" Then 'Or m_rstDetails!S1 <> "" Then
'        SG30 = SG30 & "PAC+" & m_rstDetails!S2 & "++" & m_rstDetails!S1 & "'"   'Modified by Migs 04/20/2006 changed + to ++
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Migs 06-09-2006 placed these PCI segments inside the If Then statement of the PAC segment
'        If m_rstDetails!S3 <> "" Then
'            If Len(m_rstDetails!S3) > 70 Then
'                SG30 = SG30 & "PCI+28+" & Left(m_rstDetails!S3, 35) & ":" & Mid(m_rstDetails!S3, 36, 35) & ":" & Mid(m_rstDetails!S3, 71, 35) & "'"
'            ElseIf Len(m_rstDetails!S3) > 35 Then
'                SG30 = SG30 & "PCI+28+" & Left(m_rstDetails!S3, 35) & ":" & Mid(m_rstDetails!S3, 36, 35) & "'"
'            Else
'                SG30 = SG30 & "PCI+28+" & m_rstDetails!S3 & "'"
'            End If
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'--------------------------------------------------------------

'1360    Segment group 33: MOA-SG34                      C25
'    1370    MOA, Monetary amount                        M1  MOA+<qualifier>:<amount>:<currency code>::65
'    1380    Segment group 34: CUX-DTM                   C1
'        1390    CUX, Currencies                         M1  CUX+5:<currency code>:1++<exchange rate>
'        1400    DTM, Date/time/period                   C1
    
'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG33' ----------------
'    If m_rstDetails!O2 <> "" Then
'        SG30 = SG30 & "MOA+123:" & m_rstDetails!O2 & ":::65'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If m_rstDetails!O7 <> "" Or m_rstDetails!O8 <> "" Then
'        SG30 = SG30 & "MOA+69:" & m_rstDetails!O7 & ":" & m_rstDetails!O8 & "::65'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Migs 06-09-2006 placed this CUX segment inside the If Then statement of the MOA segment
'        If m_rstDetails!O8 <> "" Then
'            'Commented by Migs 04/21/2006 replaced with the following code line
'            'SG30 = SG30 & "CUX+5:" & m_rstDetails!O8 & ":1++" & IIf(Val(m_rstDetails!OC) <> 0, m_rstDetails!OC, "") & "'"
'            SG30 = SG30 & "CUX+5:" & m_rstDetails!O8 & ":1" & IIf(Val(m_rstDetails!OC) <> 0, "++" & m_rstDetails!OC, "") & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'
'    If m_rstDetails!O9 <> "" Or m_rstDetails!OA <> "" Then
'        SG30 = SG30 & "MOA+72:" & m_rstDetails!O9 & ":" & m_rstDetails!OA & "::65'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Migs 06-09-2006 placed this CUX segment inseide the If Then statement of the MOA segment
'        If m_rstDetails!OA <> "" Then
'            'Commented by Migs 04/21/2006 replaced with the following code line
'            'SG30 = SG30 & "CUX+5:" & m_rstDetails!OA & ":1++" & IIf(Val(m_rstDetails!OD) <> 0, m_rstDetails!OD, "") & "'"
'            SG30 = SG30 & "CUX+5:" & m_rstDetails!OA & ":1" & IIf(Val(m_rstDetails!OD) <> 0, "++" & m_rstDetails!OD, "") & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'    '*Wrong : di ko sure yung condition na to!!     <--- Comment Posted by Migs on 04/04/2006 di ko rin sure pero
'    '                                                    parang ito nga ung condition na yun kya ni-uncomment ko na ung If Then statement
'    If Val(m_rstHeader!C6) <> Val(m_rstDetails!O5) Then
'        If m_rstDetails!O5 <> "" Or m_rstDetails!O6 <> "" Then
'            'Modified by Migs 04/21/2006 added IIF statement
'            SG30 = SG30 & "MOA+38:" & m_rstDetails!O5 & IIf(m_rstDetails!O6 = "", "", ":" & m_rstDetails!O6) & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'            'Migs 06-09-2006 placed this CUX segment inseide the If Then statement of the MOA segment
'            If m_rstDetails!O6 <> "" Then
'                'Commented by Migs 04/21/2006 replaced with the following code line
'                'SG30 = SG30 & "CUX+5:" & m_rstDetails!O6 & ":1++" & IIf(Val(m_rstDetails!OB) <> 0, m_rstDetails!OB, "") & "'"
'                SG30 = SG30 & "CUX+5:" & m_rstDetails!O6 & ":1" & IIf(Val(m_rstDetails!OB) <> 0, "++" & m_rstDetails!OB, "") & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'        End If
'    End If
'--------------------------------------------------------------

'1410    Segment group 35: RFF-DTM-GIN-MOA-SG36          C999
'    1420    RFF, Reference                              M1  RFF+<qualifier>:<identifier>
'    1430    DTM, Date/time/period                       C1
'    1440    GIN, Goods identity number                  C99
'    1450    MOA, Monetary amount                        C1
'    1460    Segment group 36: IMD-FTX                   C1
'        1470    IMD, Item description                   M1
'        1480    FTX, Free text                          C5

'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG35' ----------------
'    '***
'    m_rstDetailsContainer.Filter = adFilterNone
'    m_rstDetailsContainer.Filter = "Detail = " & m_lngDetailNo
'
'    If m_rstDetailsContainer.RecordCount > 0 Then
'        m_rstDetailsContainer.MoveFirst
'        Do While m_rstDetailsContainer.EOF = False
'
'            strTemp = IIf(IsNull(m_rstDetailsContainer!S4), "", Trim(m_rstDetailsContainer!S4))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "RFF+AAQ:" & strTemp & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsContainer!S5), "", Trim(m_rstDetailsContainer!S5))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "RFF+AAQ:" & strTemp & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsContainer!S6), "", Trim(m_rstDetailsContainer!S6))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "RFF+AAQ:" & strTemp & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsContainer!S7), "", Trim(m_rstDetailsContainer!S7))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "RFF+AAQ:" & strTemp & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsContainer!S8), "", Trim(m_rstDetailsContainer!S8))
'            If strTemp <> "" And strTemp <> "0" Then
'                SG30 = SG30 & "RFF+AAQ:" & strTemp & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            m_rstDetailsContainer.MoveNext
'
'        Loop
'    End If
'
'    'Unique Consignment Reference number
'    If m_rstDetails!LC <> "" Then
'        SG30 = SG30 & "RFF+UCN:" & m_rstDetails!LC & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'Quota
'    If m_rstDetails!L7 <> "" Then
'        SG30 = SG30 & "RFF+ABJ:" & m_rstDetails!L7 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'Customs preference number
'    If m_rstDetails!N5 <> "" Or m_rstDetails!N6 <> "" Then
'        SG30 = SG30 & "RFF+AIP:" & m_rstDetails!N5 & m_rstDetails!N6 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'Customs release code
'    If m_rstDetails!N1 <> "" Or m_rstDetails!N2 <> "" Or m_rstDetails!N3 <> "" Then
''Modified by Migs 06/01/2006 changed 2 to 3 for fields ND, NE, NF, NG and NH
'        SG30 = SG30 & "RFF+AHZ:" & m_rstDetails!N1 & Space(2 - Len(m_rstDetails!N1)) & _
'            m_rstDetails!N2 & Space(2 - Len(m_rstDetails!N2)) & _
'            m_rstDetails!N3 & Replace(Space(3 - Len(m_rstDetails!N3)), " ", "0") & _
'            m_rstDetails!ND & Replace(Space(3 - Len(m_rstDetails!ND)), " ", "0") & _
'            m_rstDetails!NE & Replace(Space(3 - Len(m_rstDetails!NE)), " ", "0") & _
'            m_rstDetails!NF & Replace(Space(3 - Len(m_rstDetails!NF)), " ", "0") & _
'            m_rstDetails!NG & Replace(Space(3 - Len(m_rstDetails!NG)), " ", "0") & _
'            m_rstDetails!NH & Replace(Space(3 - Len(m_rstDetails!NH)), " ", "0") & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'--------------------------------------------------------------

'1490    Segment group 37: DOC-DTM-LOC-NAD               C25
'    1500    DOC, Document/message details               M1  DOC+<doc code>:::<doc name>+<doc identifier> ::<doc source desc> :<lang code> :<version> :<revision>
'    1510    DTM, Date/time/period                       C2  DTM+<qualifier>:<date>:102
'    1520    LOC, Place/location identification          C1  LOC+127+<loc code>:::<loc name>
'    1530    NAD, Name and address                       C5

'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG37' ----------------
'    '***
'    '*Wrong
'    m_rstDetailsDocumenten.Filter = adFilterNone
'    m_rstDetailsDocumenten.Filter = "Detail = " & m_lngDetailNo
'    If m_rstDetailsDocumenten.RecordCount > 0 Then
'        m_rstDetailsDocumenten.MoveFirst
'        Do While m_rstDetailsDocumenten.EOF = False
'            If m_rstDetailsDocumenten!Q1 <> "" Or m_rstDetailsDocumenten!Q9 <> "" Then
'                'Modified by Migs 04/25/2006 added '20006001'
'                'Modified by Migs 04/22/2006 changed "+::" & m_rstDetailsDocumenten!Q9 to ":::" & m_rstDetailsDocumenten!Q9 & ":2:1"
'                'Commented by Migs 04/21/2006 replaced with the following code line
'                'SG30 = SG30 & "DOC+190:::" & m_rstDetailsDocumenten!Q1 & "+::" & m_rstDetailsDocumenten!Q9 & "'"
'                SG30 = SG30 & "DOC+190" & IIf(m_rstDetailsDocumenten!Q1 = "", "", ":::" & m_rstDetailsDocumenten!Q1) _
'                            & "+2006001::" & m_rstDetailsDocumenten!Q9 & "::2:1" & "'"
'                            '& IIf(m_rstDetailsDocumenten!Q9 = "", "", "+2006001::" & m_rstDetailsDocumenten!Q9 & "::2:1") & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'
'                'Migs 06-09-2006 placed these segments inside the If Then statement of the DOC segment --------------------------------------------
'                If m_rstDetailsDocumenten!Q3 <> "" Then
'                    SG30 = SG30 & "DTM+137:" & m_rstDetailsDocumenten!Q3 & ":102'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'
'    'Modified by Migs 06/05/2006 this is obsolete sa V1.8 ng Import CUSDEC
'    '            If m_rstDetailsDocumenten!Q4 <> "" Or (m_rstDetailsDocumenten!Q6 <> "" And m_rstDetailsDocumenten!Q5 <> "") Then
'    '                'Commented by Migs 04/21/2006 replaced with the following code line
'    '                'SG30 = SG30 & "LOC+96+:::" & m_rstDetailsDocumenten!Q4 & "+++" & m_rstDetailsDocumenten!Q6 & m_rstDetailsDocumenten!Q5 & "'"
'    '                SG30 = SG30 & "LOC+96+" & IIf(m_rstDetailsDocumenten!Q4 = "", "", ":::" & m_rstDetailsDocumenten!Q4) _
'    '                            & IIf(m_rstDetailsDocumenten!Q6 = "", "", "+++" & m_rstDetailsDocumenten!Q6 & m_rstDetailsDocumenten!Q5) & "'"
'    '                m_lngDetailSegments = m_lngDetailSegments + 1
'    '            End If
'
'                If m_rstDetailsDocumenten!Q2 <> "" Then
'                    SG30 = SG30 & "LOC+91+" & m_rstDetailsDocumenten!Q2
'                    If m_rstDetailsDocumenten!Q7 <> "" And m_rstDetailsDocumenten!Q8 <> "" Then
'                        SG30 = SG30 & "+" & m_rstDetailsDocumenten!Q7 & ":::" & m_rstDetailsDocumenten!Q8 & _
'                               IIf(m_rstDetailsDocumenten!Q6 <> "" And m_rstDetailsDocumenten!Q7 <> "", _
'                                   "++" & m_rstDetailsDocumenten!Q6 & m_rstDetailsDocumenten!Q5, "")
'                    End If
'                    SG30 = SG30 & "'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'                '----------------------------------------------------------------------------------------------------------------------------------
'
'            End If
'            m_rstDetailsDocumenten.MoveNext
'        Loop
'    End If
'
'    '*Wrong : Modify this to support old types of previous docs
'    If m_rstDetails!R1 <> "" And m_rstDetails!R2 <> "" Then
'        SG30 = SG30 & "DOC+998:::" & m_rstDetails!R1 & "-" & m_rstDetails!R2
'        m_lngDetailSegments = m_lngDetailSegments + 1
'        Select Case m_rstDetails!R2
'            Case "705", "820", "821", "822", "I81", "U82", "C8F"
'                If m_rstDetails!R5 <> "" Or m_rstDetails!R6 <> "" Then
'                    'Modified by Migs 04/21/2006 added IIF statement
'                    SG30 = SG30 & "+" & m_rstDetails!R5 & IIf(m_rstDetails!R6 = "", "", "::" & m_rstDetails!R6) '& "'"
'                End If
'            Case "740", "703"
'                If m_rstDetails!R5 <> "" Or m_rstDetails!R8 <> "" Then
'                    'Modified by Migs 04/21/2006 added IIF statement
'                    SG30 = SG30 & "+" & m_rstDetails!R5 & IIf(m_rstDetails!R8 = "", "", "::" & m_rstDetails!R8) '& "'"
'                End If
'            Case Else
'                If m_rstDetails!R5 <> "" Then
'                    SG30 = SG30 & "+" & m_rstDetails!R5 '& "'"
'                End If
'        End Select
'        SG30 = SG30 & "'"   'Added by Migs 04/21/2006
'
'        'Migs 06-09-2006 placed these segments inside the If Then statement of the DOC segment
'        If UCase(Trim(m_rstDetails!R2)) <> "MNS" Then
'            If m_rstDetails!R3 <> "" Then
'                SG30 = SG30 & "DTM+137:" & m_rstDetails!R3 & ":102'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'        End If
'
'        If m_rstDetails!R9 <> "" Then
'            SG30 = SG30 & "LOC+127+" & m_rstDetails!R9 & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'--------------------------------------------------------------

'1540    Segment group 38: TOD-LOC-FTX                   C1
'    1550    TOD, Terms of delivery or transport         M1  TOD+6++<code>
'    1560    LOC, Place/location identification          C2  LOC+1+ :::<loc name>
'    1570    FTX, Free text                              C1
    
'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG38' ----------------
'    If m_rstDetails!O3 <> "" Then
'        SG30 = SG30 & "TOD+6++" & m_rstDetails!O3 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Migs 06-09-2006 placed this LOC segment inseide the If Then statement of the TOD segment
'        If m_rstDetails!O4 <> "" Then
'            SG30 = SG30 & "LOC+1+:::" & m_rstDetails!O4 & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'--------------------------------------------------------------

'1580    Segment group 39: GDS-FTX                       C1
'    1590    GDS, Nature of cargo                        M1
'    1600    FTX, Free text                              C1

'1610    Segment group 40: GEI-PCD-DTM-RFF               C25
'    1620    GEI, Processing information                 M1  GEI+2++<code>
'    1630    PCD, Percentage details                     C1
'    1640    DTM, Date/time/period                       C1  DTM+36:<date>:102
'    1650    RFF, Reference                              C1  RFF+ANV:<GRN>::<guarantee loc>
    
'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG40' ----------------
'    If m_rstDetails!T3 <> "" Then
'        SG30 = SG30 & "GEI+2++" & m_rstDetails!T3 & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Migs 06-09-2006 placed these segments inside the If Then statement of the GEI segment
'        If UCase(m_rstDetails!T3) >= "TVRY4" And UCase(m_rstDetails!T3) <= "TVRY7" Then
'            If m_rstDetails!T5 <> "" Then
'                SG30 = SG30 & "DTM+36:" & m_rstDetails!T5 & ":102'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'
'            If m_rstDetails!T4 <> "" Then
'                SG30 = SG30 & "RFF+ANV:" & m_rstDetails!T4 & "::" & m_rstDetails!T6 & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'        End If
'    End If
'--------------------------------------------------------------

'1660    Segment group 41: TAX-MOA-GEI                   C10
'    1670    TAX, Duty/tax/fee details                   M1  TAX+1+<tax code>:::<tax>
'    1680    MOA, Monetary amount                        C2  MOA+55:<amount>:<currency code>
'    1690    GEI, Processing information                 C1

'>>> Migs 07-14-2006 <<<
'Commented these. Moved to new property 'SG41' ----------------
'    m_rstDetailsZelf.Filter = adFilterNone
'    m_rstDetailsZelf.Filter = "Detail = " & m_lngDetailNo
'
'    If m_rstDetailsZelf.RecordCount > 0 Then
'        m_rstDetailsZelf.MoveFirst
'        Do While m_rstDetailsZelf.EOF = False
'            If m_rstDetailsZelf!U1 <> "" Then
'                SG30 = SG30 & "TAX+1+" & m_rstDetailsZelf!U1 & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
''            End If
'                'Modified by Migs 06/05/2006 placed this segment inside the If Then statement ng TAX segment dahil sa hierarchy
'                'Removed currency code in V1.6
'                If m_rstDetailsZelf!U2 <> "" Then
'                    SG30 = SG30 & "MOA+55:" & m_rstDetailsZelf!U2 & "'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'            End If
'            m_rstDetailsZelf.MoveNext
'        Loop
'    End If
'--------------------------------------------------------------

'1700    Segment group 42: QVR-QTY-SG43                  C1
'    1710    QVR, Quantity variances                     M1
'    1720    QTY, Quantity                               C1
'    1730    Segment group 43: RFF-DTM                   C1
'        1740    RFF, Reference                          M1
'        1750    DTM, Date/time/period                   C1
'1760    Segment group 44:                               C50
'GIR-GEI-NAD-MEA-SG45-SG47-SG48
'    1770    GIR, Related identification numbers         M1
'    1780    GEI, Processing information                 C10
'    1790    NAD, Name and address                       C1
'    1800    MEA, Measurements                           C5
'    1810    Segment group 45: MOA-SG46                  C10
'        1820    MOA, Monetary amount                    M1
'        1830    Segment group 46: CUX-DTM               C1
'            1840    CUX, Currencies                     M1
'            1850    DTM, Date/time/period               C1
'    1860    Segment group 47: TAX-MOA-GEI               C10
'        1870    TAX, Duty/tax/fee details               M1
'        1880    MOA, Monetary amount                    C2
'        1890    GEI, Processing information             C1
'    1900    Segment group 48: DOC-DTM-LOC               C5
'        1910    DOC, Document/message details           M1
'        1920    DTM, Date/time/period                   C2
'        1930    LOC, Place/location identification      C1

    SG30 = SG30 & CST
    
    SG30 = SG30 & FTX
    
    SG30 = SG30 & LOC
    
    SG30 = SG30 & MEA
    
    SG30 = SG30 & NAD
    
    SG30 = SG30 & SG31
    
    SG30 = SG30 & SG33
    
    SG30 = SG30 & SG35
    
    SG30 = SG30 & SG37
    
    SG30 = SG30 & SG38
    
    SG30 = SG30 & SG40
    
    SG30 = SG30 & SG41

End Property

Friend Property Get CST() As String

    '1240    CST, Customs status of goods                    M1  CST+<detail no.>+<1>+<2>+<3>+<4>+<5>
    
    Dim strTemp As String
    
    'Detail Number
    strTemp = "CST+" & ReplaceSpecialCharacters(m_lngDetailNo)
    
    AddFieldToDigiSign "Detail Tab (Tariff Line) Number", m_rstDetails, True, m_lngDetailNo
    
    If m_enuDeclarationMode = enuOriginal Then
        'Commodity Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L1").Value))
        
        AddFieldToDigiSign "L1", m_rstDetails, True, m_lngDetailNo
        
        'Additional Taric Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L2").Value))
                
        AddFieldToDigiSign "L2", m_rstDetails, True, m_lngDetailNo
                
        'Additional Taric Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L3").Value))
                                
        AddFieldToDigiSign "L3", m_rstDetails, True, m_lngDetailNo
                                
        'Additional National Code
        strTemp = strTemp & "+"
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L4").Value))) > 0 Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L4").Value))
            
            AddFieldToDigiSign "L4", m_rstDetails, True, m_lngDetailNo
        End If
        
        'Additional National Code
        strTemp = strTemp & "+"
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L5").Value))) > 0 Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L5").Value))
        End If

        'Additional National Code
        strTemp = strTemp & "-"
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L6").Value))) > 0 Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L6").Value))
        End If
            
        'Digital Signature for both L5 and L6
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L5").Value))) > 0 Or _
           LenB(Trim$(FNullField(m_rstDetails.Fields("L6").Value))) > 0 Then
                AddFieldToDigiSign "L5", m_rstDetails, True, m_lngDetailNo
        End If
    End If
    
    strTemp = strTemp & "'"
    
    CST = TrimSegment(strTemp, "CST")
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(CST)

End Property

Friend Property Get FTX() As String

    '1250    FTX, Free text                                  C9  FTX+<qualifier>++<code>+<description>+<language code>
    
    Dim strTemp As String
    Dim lngGroupCtr As Long
    
    If m_enuDeclarationMode = enuOriginal Then
        'Description of Goods
        strTemp = "FTX+AAA+++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L8").Value))
            
        AddFieldToDigiSign "L8", m_rstDetails, True, m_lngDetailNo
            
        'Language of the declaration
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A9").Value)) & "'"
        
        FTX = FTX & TrimSegment(strTemp, "FTX+AAA")
    End If
    
    'Additional Information
    With m_rstDetailsBijzondere
    
        .Filter = adFilterNone
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            .Sort = "Ordinal ASC"
            
            .MoveFirst
            
            Do While Not .EOF
                            
                lngGroupCtr = lngGroupCtr + 1
                
                'Type
                strTemp = "FTX+ACB++" & ReplaceSpecialCharacters(FNullField(.Fields("P1").Value))
                
                AddFieldToDigiSign "P1", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                'Particular Indications
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("P2").Value))
                
                AddFieldToDigiSign "P2", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                'Language of the declaration
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstHeader.Fields("A9").Value))
                
                strTemp = strTemp & "'"
                
                FTX = FTX & TrimSegment(strTemp, "FTX+ACB++++" & FNullField(m_rstHeader.Fields("A9").Value))
                
                If FNullField(.Fields("P5").Value) = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
    
        .Filter = adFilterNone
    End With
    
    If m_enuDeclarationMode = enuOriginal Then
    
        If m_rstDetails.Fields("N4").Value = "I" Or _
           m_rstDetails.Fields("N4").Value = "H" Or _
           m_rstDetails.Fields("N4").Value = "K" Then
            'Method of Valuation
            strTemp = "FTX+CCI++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O1").Value)) & "'"
            
            AddFieldToDigiSign "O1", m_rstDetails, True, m_lngDetailNo
            
            FTX = FTX & TrimSegment(strTemp, "FTX+CCI")
        End If
        
        'Type of Operation
        strTemp = "FTX+AAB++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N9").Value)) & "'"
        
        AddFieldToDigiSign "N9", m_rstDetails, True, m_lngDetailNo
        
        FTX = FTX & TrimSegment(strTemp, "FTX+AAB")
        
        'Manner of Payment VAT
        strTemp = "FTX+PMT++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T2").Value))
        
        AddFieldToDigiSign "T2", m_rstDetails, True, m_lngDetailNo
        
        'Postpone Payment VAT
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T1").Value)) & "'"
        
        AddFieldToDigiSign "T1", m_rstDetails, True, m_lngDetailNo
        
        FTX = FTX & TrimSegment(strTemp, "FTX+PMT")
        
        'Declaration Type
        strTemp = "FTX+CUS++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N4").Value)) & "'"
        
        AddFieldToDigiSign "N4", m_rstDetails, True, m_lngDetailNo
        
        FTX = FTX & TrimSegment(strTemp, "FTX+CUS")
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(FTX)
    
End Property

Friend Property Get LOC() As String

    '1260    LOC, Place/location identification              C25 LOC+<qualifier>+<identification>
    
    Dim strTemp As String
    
    'Country of Origin
    strTemp = "LOC+27+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("NB").Value)) & "'"
        
    AddFieldToDigiSign "NB", m_rstDetails, True, m_lngDetailNo
        
    LOC = LOC & TrimSegment(strTemp, "LOC+27")
    
    If FNullField(m_rstDetails.Fields("N4").Value) <> "K" And _
       FNullField(m_rstDetails.Fields("N4").Value) <> "J" Then
        'Country of Destination
                
        strTemp = "LOC+28+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N7").Value)) & "'"
        
        AddFieldToDigiSign "N7", m_rstDetails, True, m_lngDetailNo
        
        LOC = LOC & TrimSegment(strTemp, "LOC+28")
        
        'Region of Destination
        strTemp = "LOC+47+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N8").Value)) & "'"
            
        AddFieldToDigiSign "N8", m_rstDetails, True, m_lngDetailNo
            
        LOC = LOC & TrimSegment(strTemp, "LOC+47")
    End If
    
    'Warehouses
    If LenB(Trim$(FNullField(m_rstDetails.Fields("M3").Value))) > 0 Or _
       LenB(Trim$(FNullField(m_rstDetails.Fields("M4").Value))) > 0 Or _
       LenB(Trim$(FNullField(m_rstDetails.Fields("M5").Value))) > 0 Then
        
        strTemp = "LOC+18+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value))
        strTemp = strTemp & Space(1) & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M4").Value))
        strTemp = strTemp & Space(1) & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value)) & "'"
        
        AddFieldToDigiSign "M3", m_rstDetails, True, m_lngDetailNo
        AddFieldToDigiSign "M4", m_rstDetails, True, m_lngDetailNo
        AddFieldToDigiSign "M5", m_rstDetails, True, m_lngDetailNo
        
        LOC = LOC & TrimSegment(strTemp, "LOC+18")
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(LOC)
    
End Property


Friend Property Get MEA() As String

    '1280    MEA, Measurements                               C20 MEA+<qualifier>+<measurement code>+<value code>:<measurement>
    
    Dim strTemp As String
    Dim lngGroupCtr As Long
    
    'Gross Weight
    strTemp = "MEA+WT+AAB+KGM:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L9").Value)) & "'"
        
    AddFieldToDigiSign "L9", m_rstDetails, True, m_lngDetailNo
        
    MEA = MEA & TrimSegment(strTemp, "MEA+WT+AAB+KGM")
    
    'Net Weight
    strTemp = "MEA+WT+AAA+KGM:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("LA").Value)) & "'"
    
    AddFieldToDigiSign "LA", m_rstDetails, True, m_lngDetailNo
    
    MEA = MEA & TrimSegment(strTemp, "MEA+WT+AAA+KGM")
           
    'Code Statistic Entity
    strTemp = "MEA+AAE++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M1").Value))
            
    AddFieldToDigiSign "M1", m_rstDetails, True, m_lngDetailNo
            
    If LenB(Trim$(FNullField(m_rstDetails.Fields("M2").Value))) > 0 Then
        If Val(m_rstDetails.Fields("M2").Value) <> 0 Then
            'Statistic Unit
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M2").Value)) & "'"
            
            AddFieldToDigiSign "M2", m_rstDetails, True, m_lngDetailNo
        End If
    End If
    
    MEA = MEA & TrimSegment(strTemp, "MEA+AAE")
                
    With m_rstDetailsBerekeningsEenheden
    
        If .RecordCount <> 0 Then
            .Sort = "ORDINAL"
            .MoveFirst
            
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
                
                'Calculation Code
                strTemp = "MEA+AAF++" & ReplaceSpecialCharacters(FNullField(.Fields("TZ").Value))
                    
                AddFieldToDigiSign "TZ", m_rstDetailsBerekeningsEenheden, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                'Calculation Entity
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("T8").Value)) & "'"
                
                AddFieldToDigiSign "T8", m_rstDetailsBerekeningsEenheden, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                If LenB(Trim$(FNullField(.Fields("TZ").Value))) <= 0 Then
                    strTemp = "MEA+AAF'"
                End If
                
                MEA = MEA & TrimSegment(strTemp, "MEA+AAF")
                
                If .Fields("T9").Value = "E" Then Exit Do
                
                .MoveNext
            
            Loop
            
        End If
    End With
            
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(MEA)

End Property

Friend Property Get NAD() As String

    '1290    NAD, Name and address                           C5  NAD+<qualifier>+<identifier>++<party name>+<street>+<city name>+:::<country sub-entity name>+<postal code>+<country code>
    
    Dim strTemp As String
    Dim strPrefix As String
    
    With m_rstDetailsHandelaars
    
        If .RecordCount <> 0 Then
        
            .MoveFirst
            
            Do While Not .EOF
            
                Select Case .Fields("VE").Value
                
                    Case enuDetailConsignee
                        strPrefix = "NAD+CN"
                
                    Case enuDetailResponsibleRepresentative
                        strPrefix = "NAD+AG"
                    
                    Case enuDetailIntracommunautaireVerwerving
                        strPrefix = "NAD+GS"
                    
                    Case enuDetailWarehouseDepositor
                        strPrefix = "NAD+WD"
                
                End Select
                
                strTemp = strPrefix
                
                'Venture Number
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V1").Value))
                
                AddFieldToDigiSign "V1", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                
                If UCase(Left(FNullField(.Fields("V1").Value), 2)) <> "BE" Or _
                   UCase(Left(FNullField(.Fields("V1").Value), 5)) = "BE999" Then
                    
                    'Name
                    strTemp = strTemp & "++" & ReplaceSpecialCharacters(Left(FNullField(.Fields("V2").Value), 35))
                    
                    If Len(FNullField(.Fields("V2").Value)) > 35 Then
                        strTemp = strTemp & ":" & ReplaceSpecialCharacters(Mid(FNullField(.Fields("V2").Value), 36, 35))
                    Else
                        strTemp = strTemp & ":"
                    End If
                        
                    AddFieldToDigiSign "V2", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                        
                    'Address 1 / Address 2
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V3").Value))
                    strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(.Fields("V4").Value))
                    
                    AddFieldToDigiSign "V3", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                    
                    'Municipality
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V6").Value))
                    
                    AddFieldToDigiSign "V6", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                    
                    'State or Province
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V7").Value))
                    
                    AddFieldToDigiSign "V7", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                    
                    'Postal Code
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V5").Value))
                    
                    AddFieldToDigiSign "V5", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                    
                    'Country
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V8").Value))
                    
                    AddFieldToDigiSign "V8", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                End If
                
                strTemp = strTemp & "'"
                
                NAD = NAD & TrimSegment(strTemp, strPrefix)
                                
                .MoveNext
            Loop
        End If
    End With
            
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(NAD)
    
'    If m_rstHeader!W1 = "" Then
'
'        If Trim(m_rstDetails!V1) = "" Or Trim(m_rstDetails!V1) = "0" Then
'
'            strTemp = strTemp & "NAD+CN+000000000"
'
'        Else
'
'            strTemp = strTemp & "NAD+CN+" & Trim(m_rstDetails!V1)   'Venture Number of the Consignee
'
'        End If
'
'        If UCase(Left(Trim(m_rstDetails!V1), 2)) <> "BE" Then
'
'            If m_rstDetails!V2 <> "" Or m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" _
'               Or m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'               Or m_rstDetails!V8 <> "" Then
'
'                'Name
'                strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'
'                If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'                   Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!V3   'Address 1
'
'                    If m_rstDetails!V4 <> "" Then
'
'                        strTemp = strTemp & ":" & m_rstDetails!V4   'Address 2
'
'                    End If
'
'                    If m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'                       Or m_rstDetails!V8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!V6   'Municipality
'
'                        If m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!V7   'State or Province
'
'                            If m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!V5   'Postal Code
'
'                                If m_rstDetails!V8 <> "" Then
'
'                                    strTemp = strTemp & "+" & m_rstDetails!V8   'Country
'
'                                End If
'
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'    NAD = NAD & strTemp & "'"
'
'    m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    strTemp = ""
'
'    If m_rstDetails!Y1 <> "" Then
'
'        NAD = NAD & "NAD+AG+" & Trim(m_rstDetails!Y1)   'Venture Number of the Resposible Representative
'
'        If m_rstDetails!Y2 <> "" Or m_rstDetails!Y3 <> "" Or m_rstDetails!Y4 <> "" _
'           Or m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'           Or m_rstDetails!Y8 <> "" Then
'
'            'Name
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'
'            If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'               Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!Y3   'Address 1
'
'                If m_rstDetails!Y4 <> "" Then
'
'                    strTemp = strTemp & ":" & m_rstDetails!Y4   'Address 2
'
'                End If
'
'                If m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'                   Or m_rstDetails!Y8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!Y6   'Municipality
'
'                    If m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!Y7   'State or Province
'
'                        If m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!Y5   'Postal Code
'
'                            If m_rstDetails!Y8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!Y8   'Country
'
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'        NAD = NAD & strTemp & "'"
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    strTemp = ""
'
'    If m_rstDetails!I1 <> "" Then
'
'        NAD = NAD & "NAD+GS+" & m_rstDetails!I1     'Venture Number of the Intracommunautaire Verwerving
'
'        If m_rstDetails!I2 <> "" Or m_rstDetails!I3 <> "" Or m_rstDetails!I4 <> "" _
'           Or m_rstDetails!I6 <> "" Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" _
'           Or m_rstDetails!I8 <> "" Then
'
'            'Name
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!I2) > 35, Left(m_rstDetails!I2, 35) & ":" & Mid(m_rstDetails!I2, 36, 35), m_rstDetails!I2)
'
'            If m_rstDetails!I3 <> "" Or m_rstDetails!I4 <> "" Or m_rstDetails!I6 <> "" _
'               Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!I3   'Address 1
'
'                If m_rstDetails!I4 <> "" Then
'
'                    strTemp = strTemp & ":" & m_rstDetails!I4   'Address 2
'
'                End If
'
'                If m_rstDetails!I6 <> "" Or m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" _
'                   Or m_rstDetails!I8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!I6   'Municipality
'
'                    If m_rstDetails!I7 <> "" Or m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!I7   'State or Province
'
'                        If m_rstDetails!I5 <> "" Or m_rstDetails!I8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!I5   'Postal Code
'
'                            If m_rstDetails!I8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!I8   'Country
'
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'        NAD = NAD & strTemp & "'"
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    strTemp = ""
'
'    If m_rstDetails!Z1 <> "" Then
'
'        NAD = NAD & "NAD+WD+" & m_rstDetails!Z1     'Venture Number of the Warehouse Depositor
'
'        If m_rstDetails!Z2 <> "" Or m_rstDetails!Z3 <> "" Or m_rstDetails!Z4 <> "" _
'           Or m_rstDetails!Z6 <> "" Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" _
'           Or m_rstDetails!M5 <> "" Then
'
'            'Name
'            strTemp = strTemp & "++" & IIf(Len(m_rstDetails!Z2) > 35, Left(m_rstDetails!Z2, 35) & ":" & Mid(m_rstDetails!Z2, 36, 35), m_rstDetails!Z2)
'
'            If m_rstDetails!Z3 <> "" Or m_rstDetails!Z4 <> "" Or m_rstDetails!Z6 <> "" _
'               Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                strTemp = strTemp & "+" & m_rstDetails!Z3   'Address 1
'
'                If m_rstDetails!Z4 <> "" Then
'
'                    strTemp = strTemp & ":" & m_rstDetails!Z4   'Address 2
'
'                End If
'
'                If m_rstDetails!Z6 <> "" Or m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" _
'                   Or m_rstDetails!M5 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!Z6   'Municipality
'
'                    If m_rstDetails!Z7 <> "" Or m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!Z7   'State or Province
'
'                        If m_rstDetails!Z5 <> "" Or m_rstDetails!M5 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!Z5   'Postal Code
'
'                            If m_rstDetails!M5 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!M5   'Country
'
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'        NAD = NAD & strTemp & "'"
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If

End Property

Friend Property Get SG31() As String

    '1310    Segment group 31: PAC-SG32                      C10
    '    1320    PAC, Package                                M1  PAC+<no. of packages>++<package type>
    '    1330    Segment group 32: PCI-FTX                   C999
    '        1340    PCI, Package identification             M1  PCI+28+<marks and number>
    '        1350    FTX, Free text                          C1
        
    Dim strTemp As String
    
    'Total Number of Packages
    strTemp = "PAC+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("S2").Value))
    
    AddFieldToDigiSign "S2", m_rstDetails, True, m_lngDetailNo
    
    'Type of Packages
    strTemp = strTemp & "++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("S1").Value)) & "'"
    
    AddFieldToDigiSign "S1", m_rstDetails, True, m_lngDetailNo
    
    SG31 = SG31 & TrimSegment(strTemp, "PAC")
    
    'Goods Description
    If TrimSegment(strTemp, "PAC") <> "" Then
    
        strTemp = "PCI+28+" & ReplaceSpecialCharacters(Left(FNullField(m_rstDetails.Fields("S3").Value), 35))
        
        If Len(FNullField(m_rstDetails.Fields("S3").Value)) > 35 Then
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(Mid(FNullField(m_rstDetails.Fields("S3").Value), 36, 35))
        Else
            strTemp = strTemp & ":"
        End If
        
        If Len(FNullField(m_rstDetails.Fields("S3").Value)) > 70 Then
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(Mid(FNullField(m_rstDetails.Fields("S3").Value), 71, 35)) & "'"
        Else
            strTemp = strTemp & ":"
        End If
        
        AddFieldToDigiSign "S3", m_rstDetails, True, m_lngDetailNo
        
        SG31 = SG31 & TrimSegment(strTemp, "PCI+28")
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG31)
    
End Property

Friend Property Get SG33() As String

    '1360    Segment group 33: MOA-SG34                      C25
    '    1370    MOA, Monetary amount                        M1  MOA+<qualifier>:<amount>:<currency code>::65
    '    1380    Segment group 34: CUX-DTM                   C1
    '        1390    CUX, Currencies                         M1  CUX+5:<currency code>:1++<exchange rate>
    '        1400    DTM, Date/time/period                   C1
       
    Dim strTemp As String
    
    
    If (LenB(Trim$(FNullField(m_rstDetails.Fields("O2").Value))) > 0) Then
        If (Val(FNullField(m_rstDetails.Fields("O2").Value)) <> 0) Then
            'Adaptation Statistical Value
            strTemp = "MOA+123:" & ReplaceSpecialCharacters(Abs(Val(FNullField(m_rstDetails.Fields("O2").Value)))) & ":"
            
            AddFieldToDigiSign "O2", m_rstDetails, True, m_lngDetailNo
            
            'Adaptation Statistical Value Currency
            strTemp = strTemp & "EUR"
                        
            If (Val(FNullField(m_rstDetails.Fields("O2").Value)) < 0) Then
                
                ' Amount is negative, so append code 65
                strTemp = strTemp & "::65'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+123::::65")
                
                If (LenB(TrimSegment(strTemp, "MOA+123")) > 0) Then
                    ' As per June 28, 2007, this is required to be
                    ' hard-coded to EUR and exchange rate of 1 all the time
                    'If (LenB(Trim$(FNullField(m_rstDetails!O6))) > 0) Then
                    '    strTemp = "CUX+5:" & ReplaceSpecialCharacters(m_rstDetails!O6)                'Currency Code
                    '    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(m_rstDetails!OB) & "'"  'Exchange Rate
                    'Else
                        strTemp = "CUX+5:" & "EUR"                          'Currency Code
                        strTemp = strTemp & ":1++" & "1" & "'"              'Exchange Rate
                    'End If
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            Else
                
                strTemp = strTemp & ":'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+123")
                
                If (LenB(TrimSegment(strTemp, "MOA+123")) > 0) Then
                    ' As per June 28, 2007, this is required to be
                    ' hard-coded to EUR and exchange rate of 1 all the time
                    'If (LenB(Trim$(FNullField(m_rstDetails!O6))) > 0) Then
                    '    strTemp = "CUX+5:" & ReplaceSpecialCharacters(m_rstDetails!O6)                'Currency Code
                    '    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(m_rstDetails!OB) & "'"  'Exchange Rate
                    'Else
                        strTemp = "CUX+5:" & "EUR"                          'Currency Code
                        strTemp = strTemp & ":1++" & "1" & "'"              'Exchange Rate
                    'End If
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            End If
        End If
    End If
    
    
    If (LenB(Trim$(FNullField(m_rstDetails.Fields("O7").Value))) > 0) Then
        If (Val(m_rstDetails.Fields("O7").Value) <> 0) Then
            'Adaptation Charge Rate
            strTemp = "MOA+69:" & ReplaceSpecialCharacters(Abs(Val(FNullField(m_rstDetails.Fields("O7").Value))))
            
            AddFieldToDigiSign "O7", m_rstDetails, True, m_lngDetailNo
            
            'Currency Code
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O8").Value))
            
            AddFieldToDigiSign "O8", m_rstDetails, True, m_lngDetailNo
            
            If (Val(m_rstDetails.Fields("O7").Value) < 0) Then
                
                ' Amount is negative, so append 65
                strTemp = strTemp & "::65'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+69::::65")
                
                If (LenB(TrimSegment(strTemp, "MOA+69::::65")) > 0) Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O8").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OC").Value)) & "'"
                        
                    AddFieldToDigiSign "OC", m_rstDetails, True, m_lngDetailNo
                        
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            Else
                
                strTemp = strTemp & ":'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+69")
                
                If (LenB(TrimSegment(strTemp, "MOA+69")) > 0) Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O8").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OC").Value)) & "'"
                    
                    AddFieldToDigiSign "OC", m_rstDetails, True, m_lngDetailNo
                    
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            End If
        End If
    End If
    
    
    If (LenB(Trim$(FNullField(m_rstDetails.Fields("O9").Value))) > 0) Then
        If (Val(m_rstDetails.Fields("O9").Value) <> 0) Then
            'Adaptation VAT- Value
            strTemp = "MOA+72:" & ReplaceSpecialCharacters(Abs(Val(FNullField(m_rstDetails.Fields("O9").Value))))
            
            AddFieldToDigiSign "O9", m_rstDetails, True, m_lngDetailNo
            
            'Currency Code
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OA").Value))
            
            AddFieldToDigiSign "OA", m_rstDetails, True, m_lngDetailNo
            
            If (Val(FNullField(m_rstDetails.Fields("O9").Value)) < 0) Then
                ' Amount is negative, so append code 65
                strTemp = strTemp & "::65'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+72::::65")
                
                If (LenB(TrimSegment(strTemp, "MOA+72::::65")) > 0) Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OA").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OD").Value)) & "'"
                    
                    AddFieldToDigiSign "OD", m_rstDetails, True, m_lngDetailNo
                    
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            Else
                
                strTemp = strTemp & ":'"
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+72")
                
                If (LenB(TrimSegment(strTemp, "MOA+72::::1")) > 0) Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OA").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OD").Value)) & "'"
                    
                    AddFieldToDigiSign "OD", m_rstDetails, True, m_lngDetailNo
                    
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                End If
                
            End If
        End If
    End If
    
    
    If (LenB(Trim$(FNullField(m_rstDetails.Fields("O5").Value))) > 0) Then
        If (Val(m_rstDetails.Fields("O5").Value) <> 0) Then
            'Invoice Price
            strTemp = "MOA+38:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O5").Value))
            
            AddFieldToDigiSign "O5", m_rstDetails, True, m_lngDetailNo
            
            'Currency Code
            strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O6").Value)) & "'"
            
            AddFieldToDigiSign "O6", m_rstDetails, True, m_lngDetailNo
            
            SG33 = SG33 & TrimSegment(strTemp, "MOA+38")
            
            If (LenB(TrimSegment(strTemp, "MOA+38")) > 0) Then
                'Currency Code
                strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O6").Value))
                
                'Exchange Rate
                strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OB").Value)) & "'"
                
                AddFieldToDigiSign "OB", m_rstDetails, True, m_lngDetailNo
                
                SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
            End If
        End If
    End If
    
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG33)
    
End Property

Friend Property Get SG35() As String

    '1410    Segment group 35: RFF-DTM-GIN-MOA-SG36          C999
    '    1420    RFF, Reference                              M1  RFF+<qualifier>:<identifier>
    '    1430    DTM, Date/time/period                       C1
    '    1440    GIN, Goods identity number                  C99
    '    1450    MOA, Monetary amount                        C1
    '    1460    Segment group 36: IMD-FTX                   C1
    '        1470    IMD, Item description                   M1
    '        1480    FTX, Free text                          C5
    
    Dim strTemp As String
    Dim lngGroupCtr As Long
    
    With m_rstDetailsContainer
    
        .Filter = adFilterNone
        
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            
            .Sort = "ORDINAL"
            .MoveFirst
            
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
                
                'Container Number
                strTemp = "RFF+AAQ:" & ReplaceSpecialCharacters(FNullField(.Fields("S4").Value)) & "'"
                                              
                SG35 = SG35 & TrimSegment(strTemp, "RFF+AAQ")
                
                'Container Number
                strTemp = "RFF+AAQ:" & ReplaceSpecialCharacters(FNullField(.Fields("S5").Value)) & "'"
                      
                SG35 = SG35 & TrimSegment(strTemp, "RFF+AAQ")
                
                AddFieldToDigiSign "S4", m_rstDetailsContainer, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                If .Fields("S6").Value = "E" Then Exit Do
                
                .MoveNext
                
            Loop
            
        End If
    
    End With
    
    If m_enuDeclarationMode = enuOriginal Then
        'Unique Consignment Reference Number
        strTemp = "RFF+UCN:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("LC").Value)) & "'"
        
        AddFieldToDigiSign "LC", m_rstDetails, True, m_lngDetailNo
        
        SG35 = SG35 & TrimSegment(strTemp, "RFF+UCN")
                   
        'Quota
        strTemp = "RFF+ABJ:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L7").Value)) & "'"
        
        AddFieldToDigiSign "L7", m_rstDetails, True, m_lngDetailNo
        
        SG35 = SG35 & TrimSegment(strTemp, "RFF+ABJ")
                                    
        'Customs preference number
        strTemp = "RFF+AIP:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N5").Value)) & "'"
        
        AddFieldToDigiSign "N5", m_rstDetails, True, m_lngDetailNo
        
        SG35 = SG35 & TrimSegment(strTemp, "RFF+AIP")
                   
        'Customs release code
        With m_rstDetails
        
            strTemp = "RFF+AHZ:"
            
            'Regime
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("N1").Value) & String(2 - Len(FNullField(.Fields("N1").Value)), "0"))
            
            AddFieldToDigiSign "N1", m_rstDetails, True, m_lngDetailNo
            
            'Previous Regime
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("N2").Value) & String(2 - FNullField(Len(.Fields("N2").Value)), "0"))
            
            AddFieldToDigiSign "N2", m_rstDetails, True, m_lngDetailNo
            
            'Community Additional Regulation 1
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("N3").Value) & String(3 - Len(FNullField(.Fields("N3").Value)), "0"))
            
            'Community Additional Regulation 2
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("ND").Value) & String(3 - Len(FNullField(.Fields("ND").Value)), "0"))
            
            'Community Additional Regulation 3
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("NE").Value) & String(3 - Len(FNullField(.Fields("NE").Value)), "0"))
            
            'Digital Signature for N3, ND, and NE
            AddFieldToDigiSign "N3", m_rstDetails, True, m_lngDetailNo
            
            'National Additional Code 1
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("NF").Value) & String(3 - Len(FNullField(.Fields("NF").Value)), "0"))
            
            'National Additional Code 2
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("NG").Value) & String(3 - Len(FNullField(.Fields("NG").Value)), "0"))
            
            'National Additional Code 3
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("NH").Value) & String(3 - Len(FNullField(.Fields("NH").Value)), "0"))
                
            'Digital Signature for NF, NG, and NH
            AddFieldToDigiSign "NF", m_rstDetails, True, m_lngDetailNo
            
            strTemp = strTemp & "'"
            
            SG35 = SG35 & TrimSegment(strTemp, "RFF+AHZ")
        End With
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG35)
    
End Property

Friend Property Get SG37() As String

    '1490    Segment group 37: DOC-DTM-LOC-NAD               C25
    '    1500    DOC, Document/message details               M1  DOC+<doc code>:::<doc name>+<doc identifier> ::<doc source desc> :<lang code> :<version> :<revision>
    '    1510    DTM, Date/time/period                       C2  DTM+<qualifier>:<date>:102
    '    1520    LOC, Place/location identification          C1  LOC+127+<loc code>:::<loc name>
    '    1530    NAD, Name and address                       C5
    
    Dim strTemp As String
    Dim lngGroupCtr As Long
    
    With m_rstDetailsDocumenten
    
        .Filter = adFilterNone
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            
            .Sort = "ORDINAL"
            
            .MoveFirst
            
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
                
                '****************************************************************************************
                'CSCLP-232 IF Q5 = 99 then it is an authorisation document and must be skipped
                '****************************************************************************************
                If Trim$(FNullField(.Fields("Q5").Value)) <> "99" Then
                    'Code
                    strTemp = "DOC+190:::" & ReplaceSpecialCharacters(FNullField(.Fields("Q1").Value))
                    
                    AddFieldToDigiSign "Q1", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    'Reference
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("Q2").Value)) & "::"
                                
                    AddFieldToDigiSign "Q2", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                                
                    'Additional Information
                    strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("Q9").Value)) & "::2:1'"
                    
                    AddFieldToDigiSign "Q9", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    'SG37 = SG37 & TrimSegment(strTemp, "DOC+190:::+2006001::::2:1")
                    SG37 = SG37 & TrimSegment(strTemp, "DOC+190:::+::::2:1")
                    
                    If TrimSegment(strTemp, "DOC+190:::+::::2:1") <> "" Then
                        'Date
                        strTemp = "DTM+137:" & ReplaceSpecialCharacters(FNullField(.Fields("Q3").Value)) & ":102'"
                        
                        AddFieldToDigiSign "Q3", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        
                        SG37 = SG37 & TrimSegment(strTemp, "DTM+137::102")
                        
                        'Office of Validation
                        strTemp = "LOC+91+" & ReplaceSpecialCharacters(FNullField(.Fields("Q4").Value))
                                                                                                            
                        AddFieldToDigiSign "Q4", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                                                                                                            
                        'Quantity Code
                        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(FNullField(.Fields("QB").Value)))
                        
                        ' MUCP-162 - Start
                        AddFieldToDigiSign "QB", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        ' MUCP-162 - End
                        
                        'Quantity
                        strTemp = strTemp & ":::" & ReplaceSpecialCharacters(FNullField(.Fields("QC").Value))
                                                        
                        AddFieldToDigiSign "QC", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                           
                        ' MUCP-162 - Start
                        'Identification Support and Indicator
                        'strTemp = strTemp & "++" & ReplaceSpecialCharacters(FNullField(.Fields("Q5").Value)) & "'"
                        '
                        'AddFieldToDigiSign "Q5", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        ' MUCP-162 - End
                        
                        SG37 = SG37 & TrimSegment(strTemp, "LOC+91")
                        
                        ' MUCP-162 - Start
                        If LenB(Trim$(FNullField(.Fields("Q5").Value))) > 0 And _
                            LenB(Trim$(FNullField(.Fields("Q7").Value))) > 0 Then
                            'Identification Support and Indicator
                            strTemp = "NAD+COT+" & ReplaceSpecialCharacters(FNullField(.Fields("Q5").Value))
                            
                            AddFieldToDigiSign "Q5", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                            
                            'Related Location
                            strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("Q7").Value)) & "'"
                            
                            AddFieldToDigiSign "Q7", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                            
                            SG37 = SG37 & TrimSegment(strTemp, "NAD+COT++")
                        End If
                        ' MUCP-162 - End
                    End If
                End If
                '****************************************************************************************
                
                If .Fields("QA").Value = "E" Then Exit Do
                
                .MoveNext
            
            Loop
            
        End If
                    
    End With
    
    If m_enuDeclarationMode = enuOriginal Then
    
        'Category
        strTemp = "DOC+998:::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R1").Value))
        
        AddFieldToDigiSign "R1", m_rstDetails, True, m_lngDetailNo
        
        'Type
        strTemp = strTemp & "-" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R2").Value))
        
        AddFieldToDigiSign "R2", m_rstDetails, True, m_lngDetailNo
        
        Select Case FNullField(m_rstDetails.Fields("R2").Value)
            
            Case "705", "820", "821", "822", "I81", "U82", "C8F"
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R5").Value))
                    
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                    
                'Article Previous Customs Authorities Regulation
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R6").Value))
                
                AddFieldToDigiSign "R6", m_rstDetails, True, m_lngDetailNo
                
            Case "740", "703"
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R5").Value))
                    
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                    
                'Waybill Number
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R8").Value))
                
                AddFieldToDigiSign "R8", m_rstDetails, True, m_lngDetailNo
                
            Case Else
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R5").Value))
                
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                
                'Article Previous Customs Authorities Regulation
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R6").Value))
                
                AddFieldToDigiSign "R6", m_rstDetails, True, m_lngDetailNo
                
        End Select
        
        strTemp = strTemp & "'"
        
        SG37 = SG37 & TrimSegment(strTemp, "DOC+998")

        If TrimSegment(strTemp, "DOC+998") <> "" Then
        
            'CSCLP-231
            'Added R5 != NIHIL condition
            If UCase(FNullField(m_rstDetails.Fields("R2").Value)) <> "MNS" And UCase(FNullField(m_rstDetails.Fields("R5").Value)) <> "NIHIL" Then
                'Date
                strTemp = "DTM+137:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R3").Value)) & ":102'"
                
                AddFieldToDigiSign "R3", m_rstDetails, True, m_lngDetailNo
                
                SG37 = SG37 & TrimSegment(strTemp, "DTM+137::102")
                
            End If
            
            'Location
            strTemp = "LOC+127+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("R9").Value)) & "'"
                
            AddFieldToDigiSign "R9", m_rstDetails, True, m_lngDetailNo
                
            SG37 = SG37 & TrimSegment(strTemp, "LOC+127")
        
        End If
    
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG37)
    
End Property

Friend Property Get SG38() As String

    '1540    Segment group 38: TOD-LOC-FTX                   C1
    '    1550    TOD, Terms of delivery or transport         M1  TOD+6++<code>
    '    1560    LOC, Place/location identification          C2  LOC+1+ :::<loc name>
    '    1570    FTX, Free text                              C1
    
    Dim strTemp As String
    
    'Supply Conditions - Type
    strTemp = "TOD+6++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O3").Value)) & "'"
        
    AddFieldToDigiSign "O3", m_rstDetails, True, m_lngDetailNo
        
    SG38 = SG38 & TrimSegment(strTemp, "TOD+6")
        
    'Supply Conditions - Place
    strTemp = "LOC+1+:::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O4").Value)) & "'"
    
    AddFieldToDigiSign "O4", m_rstDetails, True, m_lngDetailNo
    
    SG38 = SG38 & TrimSegment(strTemp, "LOC+1")
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG38)
    
End Property

Friend Property Get SG40() As String

    '1610    Segment group 40: GEI-PCD-DTM-RFF               C25
    '    1620    GEI, Processing information                 M1  GEI+2++<code>
    '    1630    PCD, Percentage details                     C1
    '    1640    DTM, Date/time/period                       C1  DTM+36:<date>:102
    '    1650    RFF, Reference                              C1  RFF+ANV:<GRN>::<guarantee loc>
    
    Dim strTemp As String
    
    'Type
    strTemp = "GEI+2++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T3").Value)) & "'"
    
    AddFieldToDigiSign "T3", m_rstDetails, True, m_lngDetailNo
    
    SG40 = SG40 & TrimSegment(strTemp, "GEI+2")
    
    ' Modifed Jan 17, 2008 - Reims. MIG incorrect according to Frank.
    ' Must include TVRY2 and TVRY3
    'If TrimSegment(strTemp, "GEI+2") <> "" And _
       UCase(m_rstDetails.Fields("T3").Value) >= "TVRY4" And _
       UCase(m_rstDetails.Fields("T3").Value) <= "TVRY7" Then
    If TrimSegment(strTemp, "GEI+2") <> "" And _
       ((UCase(FNullField(m_rstDetails.Fields("T3").Value)) >= "TVRY2" And UCase(FNullField(m_rstDetails.Fields("T3").Value)) <= "TVRY7") Or _
         UCase(FNullField(m_rstDetails.Fields("T3").Value)) = "TVRY0") Then
       
       'Date
        If Trim$(FNullField(m_rstDetails.Fields("T5").Value)) <> "" Then
            strTemp = "DTM+36:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T5").Value)) & ":102'"
                
            AddFieldToDigiSign "T5", m_rstDetails, True, m_lngDetailNo
                
            SG40 = SG40 & TrimSegment(strTemp, "DTM+36::102")
        End If
        
        'Reference
        strTemp = "RFF+ANV:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T4").Value))
            
        AddFieldToDigiSign "T4", m_rstDetails, True, m_lngDetailNo
            
        'Office
        strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("T6").Value)) & "'"
            
        AddFieldToDigiSign "T6", m_rstDetails, True, m_lngDetailNo
            
        SG40 = SG40 & TrimSegment(strTemp, "RFF+ANV")
        
    End If

    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG40)
    
End Property

Friend Property Get SG41() As String

    '1660    Segment group 41: TAX-MOA-GEI                   C10
    '    1670    TAX, Duty/tax/fee details                   M1  TAX+1+<tax code>:::<tax>
    '    1680    MOA, Monetary amount                        C2  MOA+55:<amount>:<currency code>
    '    1690    GEI, Processing information                 C1

    Dim strTemp As String
    
    Dim lngGroupCtr As Long
    
    With m_rstDetailsZelf
    
        .Filter = adFilterNone
        
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            
            .Sort = "ORDINAL"
            .MoveFirst
            
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
                
                'Type
                strTemp = "TAX+1+" & ReplaceSpecialCharacters(FNullField(.Fields("U1").Value)) & "'"
                        
                AddFieldToDigiSign "U1", m_rstDetailsZelf, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        
                SG41 = SG41 & TrimSegment(strTemp, "TAX+1")
                
                If TrimSegment(strTemp, "TAX+1") <> "" Then
                    'Amount
                    strTemp = "MOA+55:" & ReplaceSpecialCharacters(FNullField(.Fields("U2").Value)) & "'"
                    
                    AddFieldToDigiSign "U2", m_rstDetailsZelf, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    SG41 = SG41 & TrimSegment(strTemp, "MOA+55")
                
                End If
                
                .MoveNext
            
            Loop
            
        End If

    End With
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG41)
    
End Property
