VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_enuDeclarationMode As DECLARATION_MODE
Public m_lngDetailSegments As Long
Public m_lngDetailNo As Long
Public m_strLang As String
Public m_strDeclarant As String
Public m_rstHeader As ADODB.Recordset
Public m_rstDetails As ADODB.Recordset
Public m_rstDetailsBijzondere As ADODB.Recordset
Public m_rstDetailsContainer As ADODB.Recordset
Public m_rstDetailsDocumenten As ADODB.Recordset
Public m_rstDetailsHandelaars As ADODB.Recordset
Public m_rstDetailsSensitiveGoods As ADODB.Recordset

Private m_blnCurrencyUnknown As Boolean


Private m_blnIsAnEcsDeclaration As Boolean
Private m_Code As String

Public Function CreateDetail(ByVal RecordsHeader As ADODB.Recordset, _
                             ByVal RecordsDetails As ADODB.Recordset, _
                             ByVal RecordsDetailsBijzondere As ADODB.Recordset, _
                             ByVal RecordsDetailsContainer As ADODB.Recordset, _
                             ByVal RecordsDetailsDocumenten As ADODB.Recordset, _
                             ByRef RecordsDetailsHandelaars As ADODB.Recordset, _
                             ByRef RecordsDetailsSensitiveGoods As ADODB.Recordset, _
                             ByRef SegmentCount As Long, _
                             ByVal DetailNo As Long, _
                             ByVal Lang As String, _
                             ByVal DecMode As DECLARATION_MODE, _
                             Optional ByVal Code As String) As String

    Dim strMessage As String
    
    m_lngDetailSegments = 0
    m_lngDetailNo = DetailNo
    m_strLang = Lang
    m_enuDeclarationMode = DecMode
    
    'Create Clone of Detail Recordsets
    Set m_rstHeader = RecordsHeader.Clone
    Set m_rstDetails = RecordsDetails.Clone
    Set m_rstDetailsBijzondere = RecordsDetailsBijzondere.Clone
    Set m_rstDetailsContainer = RecordsDetailsContainer.Clone
    Set m_rstDetailsDocumenten = RecordsDetailsDocumenten.Clone
    Set m_rstDetailsSensitiveGoods = RecordsDetailsSensitiveGoods.Clone
    Set m_rstDetailsHandelaars = RecordsDetailsHandelaars.Clone
    
    m_blnIsAnEcsDeclaration = (Len(FNullField(m_rstHeader.Fields("AO").Value)) > 0)
    m_Code = Code
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
    
    m_rstDetailsSensitiveGoods.Filter = adFilterNone
    m_rstDetailsSensitiveGoods.Filter = "Detail = " & DetailNo
    
    m_rstDetailsHandelaars.Filter = adFilterNone
    m_rstDetailsHandelaars.Filter = "Detail = " & DetailNo
    '*********************************************************************************************
    
    If m_lngDetailNo > 1 Then
        strMessage = ""
    Else
        strMessage = UNS
    End If

    strMessage = strMessage & SG10
    strMessage = strMessage & SG30
    
    'Return Segment Count
    SegmentCount = SegmentCount + m_lngDetailSegments
    
    'Return Detail EDI Message String
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
    
    'Property modified by Migs 04/21/2006 added empty string validations (if <field> <> "")
    
    '>> Migs 06-23-2006 <<
    'Code format changed for better comprehension
    
'    Dim strTemp As String
'    Dim astrTemp() As String
'    Dim blnEmpty As Boolean
'    Dim intCtr, intLastPos As Integer
'
'    strTemp = ""  'Added by Migs 04/21/2006
'
'    SG30 = "CST+" & m_lngDetailNo   'Detail Number
'
'    blnEmpty = False
'
'    'Migs 06-09-2006 Uncommented. Code done
'    'Migs 06-08-2006 unfinished code... This will replace the following code for CST segment. Comment muna
'    strTemp = IIf(m_rstDetails!L1 = "", "0+", "1+") & IIf(m_rstDetails!L2 = "", "0+", "1+") & _
'              IIf(m_rstDetails!L3 = "", "0+", "1+") & IIf(m_rstDetails!L4 = "", "0+", "1+") & _
'              IIf(m_rstDetails!L5 = "", "0", "1")
'
'    blnEmpty = IIf(InStr(strTemp, "1") > 0, False, True)
'    strTemp = IIf(blnEmpty, "", strTemp)
'
'    If Not blnEmpty Then
'        astrTemp = Split(strTemp, "+")
'
'        For intCtr = UBound(astrTemp) To LBound(astrTemp) Step -1
'            If astrTemp(intCtr) = "1" Then
'                intLastPos = intCtr
'                Exit For
'            End If
'        Next intCtr
'
'        ReDim Preserve astrTemp(LBound(astrTemp) To intLastPos)
'
'        For intCtr = LBound(astrTemp) To intLastPos
'            Select Case intCtr
'                Case 0    'L1
'                    If m_rstDetails!N4 <> "F" And m_rstDetails!N4 <> "G" Then
'                        astrTemp(intCtr) = m_rstDetails!L1
'                    Else
'                        astrTemp(intCtr) = ""
'                    End If
'                Case Else   'L2, L3, L4, L5
'                    If m_rstDetails!N4 <> "C" And m_rstDetails!N4 <> "D" And m_rstDetails!N4 <> "E" And _
'                       m_rstDetails!N4 <> "F" And m_rstDetails!N4 <> "G" Then
'
'                        astrTemp(intCtr) = Choose(intCtr, m_rstDetails!L2, m_rstDetails!L3, m_rstDetails!L4, m_rstDetails!L5)
'                    Else
'                        astrTemp(intCtr) = ""
'                    End If
'            End Select
'        Next intCtr
'
'        strTemp = "+" & Join(astrTemp, "+")
'    End If
'
'    SG30 = SG30 & strTemp & "'"
'    m_lngDetailSegments = m_lngDetailSegments + 1
'    Erase astrTemp
'    strTemp = ""
    
'Migs 06-09-2006 replaced this code with the code above
'    If m_rstDetails!N4 <> "F" And m_rstDetails!N4 <> "G" Then
'        '8 lang ang hinihingi for PLDA Export
'        SG30 = SG30 & "+" & Left(m_rstDetails!L1, 8)
'
'        'Migs 06-08-2006 added condition for types F and G
'        If m_rstDetails!N4 <> "C" And m_rstDetails!N4 <> "D" And m_rstDetails!N4 <> "E" And m_rstDetails!N4 <> "F" And m_rstDetails!N4 <> "G" Then
'            strTemp = "+" & m_rstDetails!L2 & "+" & m_rstDetails!L3 & "+" & m_rstDetails!L4
'        Else
'            blnEmpty = True
'        End If
'    Else
'        blnEmpty = True
'    End If
'
'    'Migs 06-08-2006 added condition for types F and G
'    If m_rstDetails!N4 <> "C" And m_rstDetails!N4 <> "D" And m_rstDetails!N4 <> "E" And m_rstDetails!N4 <> "F" And m_rstDetails!N4 <> "G" Then
'        SG30 = SG30 & strTemp & IIf(blnEmpty, "+++", "") & "+" & m_rstDetails!L5 & IIf(Trim(m_rstDetails!L6) <> "", "-" & m_rstDetails!L6, "")
'    End If
'    SG30 = SG30 & "'"
'    m_lngDetailSegments = m_lngDetailSegments + 1
    
'1250    FTX, Free text                                  C9  FTX+<qualifier>++<code>+<description>+<language code>
    
'    'Goods
'    If m_rstDetails!L8 <> "" Then
'
'        SG30 = SG30 & "FTX+AAA+++" & m_rstDetails!L8    'Goods Description
'
'        SG30 = SG30 & "+" & m_strLang       'Language
'
'        'Migs 06-09-2006 This is obsolete in MIG_Export V1.7
'        'If combined then
'        'SG30 = SG30 & "+12"    Commented dahil export lang
'
'        SG30 = SG30 & "'"
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    '******************************************
'    'Migs 06-09-2006 +12 is Obsolete in MIG_Export V1.7
'    '*Wrong : check the +12
'    'Additional Information
'    m_rstDetailsBijzondere.Filter = adFilterNone
'
'    m_rstDetailsBijzondere.Filter = "Detail = " & m_lngDetailNo
'
'    If m_rstDetailsBijzondere.RecordCount > 0 Then
'
'        m_rstDetailsBijzondere.MoveFirst
'
'        Do While m_rstDetailsBijzondere.EOF = False
'
'            strTemp = IIf(IsNull(m_rstDetailsBijzondere!P2), "", Trim(m_rstDetailsBijzondere!P2))
'
'            If strTemp <> "" And strTemp <> "0" Then
'
'                SG30 = SG30 & "FTX+ACB++" & m_rstDetailsBijzondere!P1   'Type
'
'                SG30 = SG30 & "+" & m_rstDetailsBijzondere!P2   'Particular Indication
'
'                SG30 = SG30 & "+" & m_strLang & "'"     'Language
'
'                'SG30 = SG30 & "+12" if transit
'
'                m_lngDetailSegments = m_lngDetailSegments + 1
'
'            End If
'
'            strTemp = IIf(IsNull(m_rstDetailsBijzondere!P4), "", Trim(m_rstDetailsBijzondere!P4))
'
'            If strTemp <> "" And strTemp <> "0" Then
'
'                SG30 = SG30 & "FTX+ACB++" & m_rstDetailsBijzondere!P3   'Type
'
'                SG30 = SG30 & "+" & m_rstDetailsBijzondere!P4   'Particular Indication
'
'                SG30 = SG30 & "+" & m_strLang & "'"     'Language
'
'                'SG30 = SG30 & "+12" if transit
'
'                m_lngDetailSegments = m_lngDetailSegments + 1
'
'            End If
'
'            m_rstDetailsBijzondere.MoveNext
'
'        Loop
'
'    End If
'    '******************************************
'
'    'Nature of transaction
'    If m_rstDetails!N9 <> "" And m_rstDetails!NA <> "" Then
'
'        SG30 = SG30 & "FTX+AAB++" & m_rstDetails!N9 & m_rstDetails!NA   'Type of Operation
'
'        'Migs 06-09-2006 This is obsolete in MIG_Export V1.7
'        'Commented by Migs 04/05/2006 temporarily commented dahil export muna ang pagaganahin
'        '    'If combined then
'        '        SG30 = SG30 & "+12"
'
'        SG30 = SG30 & "'"
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    If m_rstDetails!N4 <> "" Then
'
'        SG30 = SG30 & "FTX+CUS++" & m_rstDetails!N4 & "'"   'Declaration Type
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
    
'1260    LOC, Place/location identification              C25 LOC+<qualifier>+<identification>
    
'    If m_rstDetails!NB <> "" Then
'
'        SG30 = SG30 & "LOC+27+" & m_rstDetails!NB & "'"     'Country of Origin
'
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'    End If
'
'    If m_rstDetails!N4 <> "D" And m_rstDetails!N4 <> "G" Then
'
'        If m_rstDetails!N7 <> "" Then
'
'            SG30 = SG30 & "LOC+28+" & m_rstDetails!N7 & "'"     'Country of Destination
'
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'        End If
'
'    End If
'
'    If Not (m_rstDetails!N4 = "E" Or m_rstDetails!N4 = "F" Or m_rstDetails!N4 = "G") Then
'
'        If m_rstDetails!M3 <> "" And m_rstDetails!M4 <> "" And m_rstDetails!M5 <> "" Then
'
'            'Migs 06-09-2006 Added spaces in between the fields
'            SG30 = SG30 & "LOC+18+" & m_rstDetails!M3 & " " & m_rstDetails!M4 & " " & m_rstDetails!M5 & "'"     'Warehouse
'
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'        End If
'
'    End If
'
'    If Not (UCase(m_rstDetails!N4) = "B" Or UCase(m_rstDetails!N4) = "D" Or UCase(m_rstDetails!N4) = "G" Or UCase(m_rstDetails!N4) = "F") Then
'
'        If m_rstDetails!NC <> "" Then
'
'            SG30 = SG30 & "LOC+106+" & m_rstDetails!NC & "'"    'Region of Orgin
'
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'        End If
'
'    End If
    
'1270    DTM, Date/time/period                           C5
'1280    MEA, Measurements                               C20 MEA+<qualifier>+<measurement code>+<value code>:<measurement>
    
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
'    If Not (m_rstDetails!N4 = "F" Or m_rstDetails!N4 = "G") Then
'
'        'Migs 06-09-2006 removed M2 from the condition. Only M1 is required
'        If m_rstDetails!M1 <> "" Then 'Or m_rstDetails!M2 <> "" Then
'            'Modified by Migs 04/21/2006 added IIF statement
'            SG30 = SG30 & "MEA+AAE++" & m_rstDetails!M1 & IIf(m_rstDetails!M2 = "", "", ":" & m_rstDetails!M2) & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'
'    End If
    
'1290    NAD, Name and address                           C5  NAD+<qualifier>+<identifier>++<party name>+<street>+<city name>+:::<country sub-entity name>+<postal code>+<country code>
    
'    If UCase(m_rstDetails!N4) = "F" Or UCase(m_rstDetails!N4) = "B" Then
'        If m_rstDetails!V1 <> "" Then
'            SG30 = SG30 & "NAD+CN+" & Trim(m_rstDetails!V1)
'            If UCase(Left(m_rstDetails!V1, 2)) <> "BE" Or Trim(m_rstDetails!V1) Then
'    'Commented by Migs 04/21/2006 replaced with the following code lines
'    '            SG30 = SG30 & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2) & _
'    '                "+" & m_rstDetails!V3 & ":" & m_rstDetails!V4 & "+" & m_rstDetails!V6 & "+" & m_rstDetails!V7 & "+" & m_rstDetails!V5 & "+" & m_rstDetails!V8
'                strTemp = ""
'                If m_rstDetails!V2 <> "" Or m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" _
'                   Or m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'                   Or m_rstDetails!V8 <> "" Then
'
'                    strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'                    If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'                       Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!V3 & IIf(m_rstDetails!V4 = "", "", ":" & m_rstDetails!V4)
'                        If m_rstDetails!V6 <> "" Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" _
'                           Or m_rstDetails!V8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!V6
'                            If m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!V7
'                                If m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                                    strTemp = strTemp & "+" & m_rstDetails!V5
'                                    If m_rstDetails!V8 <> "" Then strTemp = strTemp & "+" & m_rstDetails!V8
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'            SG30 = SG30 & strTemp & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'        End If
'    End If
'
'    If m_rstDetails!Y1 <> "" Then
'        SG30 = SG30 & "NAD+EX+" & Trim(m_rstDetails!Y1)
'        If UCase(Left(m_rstDetails!Y1, 2)) <> "BE" Or Trim(m_rstDetails!Y1) <> "" Then
'    'Commented by Migs 04/21/2006 replaced with the following code lines
'    '        SG30 = SG30 & "++" & IIf(Len(m_rstDetails!Y2) > 35, Left(m_rstDetails!Y2, 35) & ":" & Mid(m_rstDetails!Y2, 36, 35), m_rstDetails!Y2) & _
'    '            "+" & m_rstDetails!Y3 & ":" & m_rstDetails!Y4 & "+" & m_rstDetails!Y6 & "+" & m_rstDetails!Y7 & "+" & m_rstDetails!Y5 & "+" & m_rstDetails!Y8
'            strTemp = ""
'            If m_rstDetails!Y2 <> "" Or m_rstDetails!Y3 <> "" Or m_rstDetails!Y4 <> "" _
'               Or m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'               Or m_rstDetails!Y8 <> "" Then
'                'Modified by Migs 04/24/2006 replaced with the next code line
'                'strTemp = strTemp & "++" & IIf(Len(m_rstDetails!V2) > 35, Left(m_rstDetails!V2, 35) & ":" & Mid(m_rstDetails!V2, 36, 35), m_rstDetails!V2)
'                strTemp = strTemp & "++" & IIf(Len(m_rstDetails!Y2) > 35, Left(m_rstDetails!Y2, 35) & ":" & Mid(m_rstDetails!Y2, 36, 35), m_rstDetails!Y2)
'                If m_rstDetails!V3 <> "" Or m_rstDetails!V4 <> "" Or m_rstDetails!V6 <> "" _
'                   Or m_rstDetails!V7 <> "" Or m_rstDetails!V5 <> "" Or m_rstDetails!V8 <> "" Then
'
'                    strTemp = strTemp & "+" & m_rstDetails!Y3 & IIf(m_rstDetails!Y4 = "", "", ":" & m_rstDetails!Y4)
'                    If m_rstDetails!Y6 <> "" Or m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" _
'                       Or m_rstDetails!Y8 <> "" Then
'
'                        strTemp = strTemp & "+" & m_rstDetails!Y6
'                        If m_rstDetails!Y7 <> "" Or m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                            strTemp = strTemp & "+" & m_rstDetails!Y7
'                            If m_rstDetails!Y5 <> "" Or m_rstDetails!Y8 <> "" Then
'
'                                strTemp = strTemp & "+" & m_rstDetails!Y5
'                                If m_rstDetails!Y8 <> "" Then strTemp = strTemp & "+" & m_rstDetails!Y8
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        SG30 = SG30 & strTemp & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    '*Wrong : Exporter => Declarant?? Mukhang X1 yung declarant if to be compared from EDI NCTS  <---X1 naman ung nakalagay sa m_strDeclarant!
'    If UCase(m_strDeclarant) <> UCase(m_rstDetails!Z1) And Trim(m_rstDetails!Z1) <> "" Then
'        SG30 = SG30 & "NAD+WD+" & m_rstDetails!Z1
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
    
'1300    TDT, Transport information                      C9

'***
'1310    Segment group 31: PAC-SG32                      C10
'    1320    PAC, Package                                M1  PAC+<no. of packages>++<package type>
'    1330    Segment group 32: PCI-FTX                   C999
'        1340    PCI, Package identification             M1  PCI+28+<marks and number>
'        1350    FTX, Free text                          C1
    
'    If m_rstDetails!S2 <> "" Or m_rstDetails!S1 <> "" Then
'        'Modified by Migs 04/21/2006 added IIF statement
'        SG30 = SG30 & "PAC+" & m_rstDetails!S2 & IIf(m_rstDetails!S1 = "", "", "++" & m_rstDetails!S1) & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        'Modified by Migs 05/05/2006 placed these codes inside the End..If statement of the PAC segment dahil the PAC segment must exist
'        'for the PCI segment to exist
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
    
'1360    Segment group 33: MOA-SG34                      C25
'    1370    MOA, Monetary amount                        M1  MOA+<qualifier>:<amount>:<currency code>::65
'    1380    Segment group 34: CUX-DTM                   C1
'        1390    CUX, Currencies                         M1  CUX+5:<currency code>:1++<exchange rate>
'        1400    DTM, Date/time/period                   C1

'    '*Wrong : saan ko kukunin yung currency code <--- see added code
'    If Not (m_rstDetails!N4 = "D" Or m_rstDetails!N4 = "F" Or m_rstDetails!N4 = "G") Then
'        If m_rstDetails!O2 <> "" Then
'            SG30 = SG30 & "MOA+123:" & m_rstDetails!O2 & ":" & m_rstDetails!O6 & "'"    'Migs 06-13-2006 added O6 (Currency code)
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'            'Added by Migs 05/12/2006 ito na ung currency OK!
'            If m_rstDetails!O6 <> "" Then
'                SG30 = SG30 & "CUX+5:" & m_rstDetails!O6 & ":1" & IIf(m_rstDetails!OB = "", "", "++" & m_rstDetails!OB) & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'        End If
'    End If
                    
'1410    Segment group 35: RFF-DTM-GIN-MOA-SG36          C999
'    1420    RFF, Reference                              M1  RFF+<qualifier>:<identifier>
'    1430    DTM, Date/time/period                       C1
'    1440    GIN, Goods identity number                  C99
'    1450    MOA, Monetary amount                        C1
'    1460    Segment group 36: IMD-FTX                   C1
'        1470    IMD, Item description                   M1
'        1480    FTX, Free text                          C5

'    '***
'
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
'
'            m_rstDetailsContainer.MoveNext
'
'        Loop
'    End If
'
'    If m_rstDetails!LC <> "" Then
'        SG30 = SG30 & "RFF+UCN:" & m_rstDetails!LC & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    If Not (m_rstDetails!N4 = "F" Or m_rstDetails!N4 = "G") Then
'        'Customs release code
''Modified by Migs 06/06/2006 binalik sa dati then pinalitan ng 3 ung 2 sa N3, ND to NH
''Modified by Migs 06/01/2006 modified this code... see code replacement
'        SG30 = SG30 & "RFF+AHZ:" & m_rstDetails!N1 & Space(2 - Len(m_rstDetails!N1)) & _
'            m_rstDetails!N2 & Space(2 - Len(m_rstDetails!N2)) & _
'            m_rstDetails!N3 & Replace(Space(3 - Len(m_rstDetails!N3)), " ", "0") & _
'            m_rstDetails!ND & Replace(Space(3 - Len(m_rstDetails!ND)), " ", "0") & _
'            m_rstDetails!NE & Replace(Space(3 - Len(m_rstDetails!NE)), " ", "0") & _
'            m_rstDetails!NF & Replace(Space(3 - Len(m_rstDetails!NF)), " ", "0") & _
'            m_rstDetails!NG & Replace(Space(3 - Len(m_rstDetails!NG)), " ", "0") & _
'            m_rstDetails!NH & Replace(Space(3 - Len(m_rstDetails!NH)), " ", "0") & "'"
''        SG30 = SG30 & "RFF+AHZ:" & m_rstDetails!N1 & m_rstDetails!N2 & m_rstDetails!N3 & String(3 - Len(m_rstDetails!N3), "0") & String(15, "0") & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'
'    'UNDG code
'    If m_rstDetails!SA <> "" Then
'        SG30 = SG30 & "RFF+UN:" & m_rstDetails!SA & "'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
    
'1490    Segment group 37: DOC-DTM-LOC-NAD               C25
'    1500    DOC, Document/message details               M1  DOC+<doc code>:::<doc name>+<doc identifier> ::<doc source desc> :<lang code> :<version> :<revision>
'    1510    DTM, Date/time/period                       C2  DTM+<qualifier>:<date>:102
'    1520    LOC, Place/location identification          C1  LOC+127+<loc code>:::<loc name>
'    1530    NAD, Name and address                       C5
        
'    '***
'    m_rstDetailsDocumenten.Filter = adFilterNone
'    m_rstDetailsDocumenten.Filter = "Detail = " & m_lngDetailNo
'    If m_rstDetailsDocumenten.RecordCount > 0 Then
'        m_rstDetailsDocumenten.MoveFirst
'        Do While m_rstDetailsDocumenten.EOF = False
'
'            '*************************************************************
'            '*Wrong => if combined then "m_rstDetailsBijgevoegde!Q1+:12"
'            If m_rstDetailsDocumenten!Q1 <> "" Then
'                SG30 = SG30 & "DOC+190:::" & m_rstDetailsDocumenten!Q1 & ":1:::" & m_strLang & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'
'                'Modified by Migs 05/10/2006 placed code for DTM and LOC inside the IF..THEN statement of DOC
'                If m_rstDetailsDocumenten!Q3 <> "" Then
'                    SG30 = SG30 & "DTM+137:" & m_rstDetailsDocumenten!Q3 & ":102'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'
'                If m_rstDetailsDocumenten!Q2 <> "" Then
'                    SG30 = SG30 & "LOC+91+" & m_rstDetailsDocumenten!Q2 & "'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'
'                If m_rstDetailsDocumenten!Q4 <> "" Or (m_rstDetailsDocumenten!Q6 <> "" And m_rstDetailsDocumenten!Q5 <> "") Then
'                    'Migs 06-08-2006 replaced these codes
'                    'Modified by Migs 04/21/2006 added IIF statements
''                    SG30 = SG30 & "LOC+96+" & IIf(m_rstDetailsDocumenten!Q4 = "", "", ":::" & m_rstDetailsDocumenten!Q4 & _
''                           IIf(m_rstDetailsDocumenten!Q6 <> "" And m_rstDetailsDocumenten!Q5 <> "", "+++" & _
''                               m_rstDetailsDocumenten!Q6 & m_rstDetailsDocumenten!Q5, "")) & "'"
'                    SG30 = SG30 & "LOC+96+" & IIf(m_rstDetailsDocumenten!Q4 = "", "", ":::" & m_rstDetailsDocumenten!Q4) & _
'                           IIf(m_rstDetailsDocumenten!Q6 <> "" And m_rstDetailsDocumenten!Q5 <> "", "+++" & _
'                               m_rstDetailsDocumenten!Q6 & m_rstDetailsDocumenten!Q5, "") & "'"
'                    m_lngDetailSegments = m_lngDetailSegments + 1
'                End If
'
'            End If
'            m_rstDetailsDocumenten.MoveNext
'            '*************************************************************
'
'        Loop
'    End If
'
'    'Removed 325, and 380 in Export V1.5 => tax declaration document (acc4), Pro forma invoice, commercial invoice
'    If m_rstDetails!R1 <> "" And m_rstDetails!R2 <> "" Then
'        SG30 = SG30 & "DOC+998:::" & m_rstDetails!R1 & m_rstDetails!R2
'        m_lngDetailSegments = m_lngDetailSegments + 1
'
'        Select Case m_rstDetails!R2
'            Case "705"
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
'        SG30 = SG30 & "'"   'aDDED BY mIGS 04/21/2006
'
'        'Migs 06-09-2006 placed code for DTM and LOC inside the IF..THEN statement of DOC
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
    
'1540    Segment group 38: TOD-LOC-FTX                   C1
'    1550    TOD, Terms of delivery or transport         M1  TOD+6++<code>
'    1560    LOC, Place/location identification          C2  LOC+1+ :::<loc name>
'    1570    FTX, Free text                              C1

'    '*Wrong : Kulang ng mga related locations <--- meron naman na! ung IIF statement...
'    If Not (m_rstDetails!N4 = "A" Or m_rstDetails!N4 = "C" Or m_rstDetails!N4 = "D") Then
'        If m_rstDetails!O3 <> "" Then
'            SG30 = SG30 & "TOD+6++" & m_rstDetails!O3 & "'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'            'Modified by Migs 05/10/2006 placed this code inside the IF..THEN statement of the TOD segment
'            If m_rstDetails!O4 <> "" Then
'                SG30 = SG30 & "LOC+1+:::" & m_rstDetails!O4 & IIf(Trim(m_rstDetails!OE) <> "", "+" & m_rstDetails!OE & "+" & m_rstDetails!OF, "") & "'"
'                m_lngDetailSegments = m_lngDetailSegments + 1
'            End If
'        End If
'    End If
    
'1580    Segment group 39: GDS-FTX                       C1
'    1590    GDS, Nature of cargo                        M1
'    1600    FTX, Free text                              C1

'1610    Segment group 40: GEI-PCD-DTM-RFF               C25
'    1620    GEI, Processing information                 M1  GEI+2++<code>
'    1630    PCD, Percentage details                     C1
'    1640    DTM, Date/time/period                       C1  DTM+36:<date>:102
'    1650    RFF, Reference                              C1  RFF+ANV:<GRN>::<guarantee loc>

'    SG30 = SG30 & "GEI+2++" & m_rstDetails!T3 & "'"
'    m_lngDetailSegments = m_lngDetailSegments + 1
'
'    SG30 = SG30 & "DTM+36:" & Format(m_rstDetails!T5, "yyyymmdd") & ":102'"
'    m_lngDetailSegments = m_lngDetailSegments + 1
'
'    SG30 = SG30 & "RFF+ANV:" & m_rstDetails!T4 & "::" & m_rstDetails!T6 & "'"
'    m_lngDetailSegments = m_lngDetailSegments + 1

'1660    Segment group 41: TAX-MOA-GEI                   C10
'    1670    TAX, Duty/tax/fee details                   M1  TAX+1+<tax code>:::<tax>
'    1680    MOA, Monetary amount                        C2  MOA+55:<amount>:<currency code>
'    1690    GEI, Processing information                 C1

'    m_rstDetailsZelf.Filter = adFilterNone
'    m_rstDetailsZelf.Filter = "Detail = " & m_lngDetailNo
'
'    If m_rstDetailsZelf.RecordCount > 0 Then
'        m_rstDetailsZelf.MoveFirst
'        Do While m_rstDetailsZelf.EOF = False
'            SG30 = SG30 & "TAX+1+" & m_rstDetails!U1 & ":::" & m_rstDetails!U3
'            m_lngDetailSegments = m_lngDetailSegments + 1
'
'            'Removed
'            'saan ko kukunin currency?
'            SG30 = SG30 & "MOA+55:" & m_rstDetails!U2 & "'" '":EUR'"
'            m_lngDetailSegments = m_lngDetailSegments + 1
'            m_rstDetailsZelf.MoveNext
'        Loop
'    End If
    
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
    
'Commented by Migs 04/05/2006 temporarily commented dahil export muna ang pagaganahin
'    'For combined only
'    If Val(m_rstDetails!SB) > 0 Then
'        SG30 = SG30 & "GIR+3+" & m_rstDetails!SB & "+1+1'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    If Val(m_rstDetails!SC) > 0 Then
'        SG30 = SG30 & "GIR+3+" & m_rstDetails!SC & "+1+1'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    If Val(m_rstDetails!SD) > 0 Then
'        SG30 = SG30 & "GIR+3+" & m_rstDetails!SD & "+1+1'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If
'    If Val(m_rstDetails!SE) > 0 Then
'        SG30 = SG30 & "GIR+3+" & m_rstDetails!SE & "+1+1'"
'        m_lngDetailSegments = m_lngDetailSegments + 1
'    End If

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
    SG30 = SG30 & SG44
End Property

Friend Property Get CST() As String

    '1240    CST, Customs status of goods                    M1  CST+<detail no.>+<1>+<2>+<3>+<4>+<5>

    'If N4 = C or D then L1 is not mandatory
    'if N4 = C, D, E, F or G then L2 to L6 are not needed
    
    Dim strTemp As String

    'Detail Number
    strTemp = "CST+" & ReplaceSpecialCharacters(m_lngDetailNo)
    
    AddFieldToDigiSign "Detail Tab (Tariff Line) Number", m_rstDetails, True, m_lngDetailNo
    
    If (m_enuDeclarationMode = enuOriginal) Then
        'Commodity Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(Left(FNullField(m_rstDetails.Fields("L1").Value), 8))
        
        AddFieldToDigiSign "L1", m_rstDetails, True, m_lngDetailNo
        
        'First Additional Taric Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L2").Value))
        
        AddFieldToDigiSign "L2", m_rstDetails, True, m_lngDetailNo
        
        'Second Additional Taric Code
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L3").Value))
        
        AddFieldToDigiSign "L3", m_rstDetails, True, m_lngDetailNo
        
        'First Additional National Code
        strTemp = strTemp & "+"
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L4").Value))) > 0 Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L4").Value))
            
            AddFieldToDigiSign "L4", m_rstDetails, True, m_lngDetailNo
        End If
        
        'Second Additional National Code
        strTemp = strTemp & "+"
        If LenB(Trim$(FNullField(m_rstDetails.Fields("L5").Value))) > 0 Then
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L5").Value))
        End If
        
        'Third Additional National Code
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
    
    If m_enuDeclarationMode = enuOriginal Then

        'Goods Description
        strTemp = "FTX+AAA+++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L8").Value))
        
        AddFieldToDigiSign "L8", m_rstDetails, True, m_lngDetailNo
        
        'Language
        strTemp = strTemp & "+" & ReplaceSpecialCharacters(m_strLang) & "'"
        FTX = FTX & TrimSegment(strTemp, "FTX+AAA++++" & m_strLang)
            
        'Type of Operation
        strTemp = "FTX+AAB++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N9").Value)) & "'"
        
        AddFieldToDigiSign "N9", m_rstDetails, True, m_lngDetailNo
        
        FTX = FTX & TrimSegment(strTemp, "FTX+AAB")
                    
        'Declaration Type
        strTemp = "FTX+CUS++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N4").Value)) & "'"
        
        AddFieldToDigiSign "N4", m_rstDetails, True, m_lngDetailNo
        
        FTX = FTX & TrimSegment(strTemp, "FTX+CUS")
    End If
    
    With m_rstDetailsBijzondere

        .Filter = adFilterNone
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            .Sort = "Ordinal ASC"
            
            .MoveFirst
            
            Do While Not .EOF
                'Type
                strTemp = "FTX+ACB++" & ReplaceSpecialCharacters(FNullField(.Fields("P1").Value))
                
                AddFieldToDigiSign "P1", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                'Particular Indication
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("P2").Value))
                
                AddFieldToDigiSign "P2", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                'Language
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(m_strLang) & "'"
                
                FTX = FTX & TrimSegment(strTemp, "FTX+ACB++++" & m_strLang)
                
                If .Fields("P5").Value = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
        
        .Filter = adFilterNone
    End With
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(FTX)
    
End Property

Friend Property Get LOC() As String

    '1260    LOC, Place/location identification              C25 LOC+<qualifier>+<identification>
    
    Dim strTemp As String
    Dim strN4 As String
    
    strN4 = UCase$(Trim$(FNullField(m_rstDetails.Fields("N4").Value)))
    
    'Country of Origin
    strTemp = "LOC+27+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("NB").Value)) & "'"
    
    AddFieldToDigiSign "NB", m_rstDetails, True, m_lngDetailNo
    
    LOC = LOC & TrimSegment(strTemp, "LOC+27")
    
    'Country of Destination
    If Not strN4 = "G" Then
        If m_blnIsAnEcsDeclaration And Not RoutingBoxEmpty("AE", m_Code) Then
            strTemp = "LOC+11+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N7").Value)) & "'"
        Else
            strTemp = "LOC+36+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N7").Value)) & "'"
        End If
        
        AddFieldToDigiSign "N7", m_rstDetails, True, m_lngDetailNo
        
        If m_blnIsAnEcsDeclaration And Not RoutingBoxEmpty("AE", m_Code) Then
            LOC = LOC & TrimSegment(strTemp, "LOC+11")
        Else
            LOC = LOC & TrimSegment(strTemp, "LOC+36")
        End If
        
    ElseIf m_blnIsAnEcsDeclaration Then
        strTemp = "LOC+11+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("N7").Value)) & "'"
        
        AddFieldToDigiSign "N7", m_rstDetails, True, m_lngDetailNo
                
        LOC = LOC & TrimSegment(strTemp, "LOC+11")
    End If
            
    'Warehouse
    If Not (strN4 = "E" Or strN4 = "F" Or strN4 = "G") Then
        
        If LenB(Trim$(FNullField(m_rstDetails.Fields("M3").Value))) > 0 Or _
            LenB(Trim$(FNullField(m_rstDetails.Fields("M4").Value))) > 0 Or _
            LenB(Trim$(FNullField(m_rstDetails.Fields("M5").Value))) > 0 Then
        
            strTemp = "LOC+18+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M3").Value)) & Space(1)
            
            AddFieldToDigiSign "M3", m_rstDetails, True, m_lngDetailNo
            
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M4").Value)) & Space(1)
            
            AddFieldToDigiSign "M4", m_rstDetails, True, m_lngDetailNo
            
            strTemp = strTemp & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M5").Value)) & "'"
            
            AddFieldToDigiSign "M5", m_rstDetails, True, m_lngDetailNo
            
            LOC = LOC & TrimSegment(strTemp, "LOC+18+  ")
        End If
    End If
    
    'Region of Origin
    If Not (strN4 = "B" Or strN4 = "D" Or strN4 = "G" Or strN4 = "F") Then
    
        strTemp = "LOC+106+" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("NC").Value)) & "'"
        
        AddFieldToDigiSign "NC", m_rstDetails, True, m_lngDetailNo
        
        LOC = LOC & TrimSegment(strTemp, "LOC+106")
    End If

    m_lngDetailSegments = m_lngDetailSegments + CountSegments(LOC)
    
End Property

Friend Property Get MEA() As String

    '1280    MEA, Measurements                               C20 MEA+<qualifier>+<measurement code>+<value code>:<measurement>
    
    Dim strTemp As String
    
    'Gross Weight
    strTemp = "MEA+WT+AAB+KGM:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("L9").Value)) & "'"
    
    AddFieldToDigiSign "L9", m_rstDetails, True, m_lngDetailNo
    
    MEA = MEA & TrimSegment(strTemp, "MEA+WT+AAB+KGM")
    
    'Net Weight
    strTemp = "MEA+WT+AAA+KGM:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("LA").Value)) & "'"
    
    AddFieldToDigiSign "LA", m_rstDetails, True, m_lngDetailNo
    
    MEA = MEA & TrimSegment(strTemp, "MEA+WT+AAA+KGM")
            
    If Not (FNullField(m_rstDetails.Fields("N4").Value) = "F" Or FNullField(m_rstDetails.Fields("N4").Value) = "G") Then
        
        'Code Statistic Entity
        strTemp = "MEA+AAE++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M1").Value))
        
        AddFieldToDigiSign "M1", m_rstDetails, True, m_lngDetailNo
        
        'Statistic Entities
        If LenB(Trim$(FNullField(m_rstDetails.Fields("M2").Value))) > 0 Then
            If Val(m_rstDetails.Fields("M2").Value) <> 0 Then
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("M2").Value)) & "'"
                
                AddFieldToDigiSign "M2", m_rstDetails, True, m_lngDetailNo
            End If
        End If
        
        MEA = MEA & TrimSegment(strTemp, "MEA+AAE")
    End If

    m_lngDetailSegments = m_lngDetailSegments + CountSegments(MEA)
    
End Property

Friend Property Get NAD() As String

    '1290    NAD, Name and address                           C5  NAD+<qualifier>+<identifier>++<party name>+<street>+<city name>+:::<country sub-entity name>+<postal code>+<country code>
    
    Dim strTemp As String
    Dim strPrefix As String
    Dim strV1 As String
    
    With m_rstDetailsHandelaars
        If .RecordCount <> 0 Then
            .MoveFirst
            
            Do While Not .EOF
                Select Case FNullField(.Fields("VE").Value)
                    Case enuExportDetailExporter
                        strPrefix = "NAD+EX"
                    Case enuExportDetailConsignee
                        strPrefix = "NAD+CN"
                    Case enuExportDetailWarehouseDepositor
                        strPrefix = "NAD+WD"
                End Select
                
                'Venture Number
                strTemp = strPrefix
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("V1").Value))
                   
                AddFieldToDigiSign "V1", m_rstDetailsHandelaars, True, m_lngDetailNo, FNullField(.Fields("VE").Value)
                   
                strV1 = FNullField(.Fields("V1").Value)

                If (strV1 = "" Or Left$(strV1, 5) = "BE999" Or Left$(strV1, 2) <> "BE") Then
                    If (FNullField(.Fields("VE").Value) = enuExportDetailConsignee And (CheckRegimeType("B", "=") Or CheckRegimeType("F", "="))) Or _
                       (FNullField(.Fields("VE").Value) = enuExportDetailExporter And (m_strDeclarantVentureNumber <> "" And Not VentureNumbersAreSame(m_strDeclarantVentureNumber, "EX", enuExport))) Or _
                       (FNullField(.Fields("VE").Value) = enuExportDetailWarehouseDepositor And (m_strDeclarantVentureNumber <> "" And m_strDeclarantVentureNumber = strV1)) Then
                        
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
                        
                        AddFieldToDigiSign "V3", m_rstDetailsHandelaars, True, m_lngDetailNo, .Fields("VE").Value
                        
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
                End If
                
                strTemp = strTemp & "'"
                NAD = NAD & TrimSegment(strTemp, strPrefix)
                                
                .MoveNext
            Loop
        End If
    End With
            
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(NAD)
    
    
   'COMMENTED: Glenn
    'CN
   ' With m_rstDetails
   '
   '     If UCase(!N4) = "B" Or UCase(!N4) = "F" Then
   '
   '         strTemp = "NAD+CN+" & !V1                       'Venture Number of the Consignee
   '
   '         If !V1 = "" Or UCase(Left(!V1, 2)) <> "BE" Then
   '
   '             strTemp = strTemp & "++" & Left(!V2, 35)    'Name
   '             strTemp = strTemp & ":" & Mid(!V2, 36, 35)  'Name
   '             strTemp = strTemp & "+" & !V3               'Address 1
   '             strTemp = strTemp & ":" & !V4               'Address 2
   '             strTemp = strTemp & "+" & !V6               'Municipality
   '             strTemp = strTemp & "+" & !V7               'State or Province
   '             strTemp = strTemp & "+" & !V5               'Postal Code
   '             strTemp = strTemp & "+" & !V8               'Country
   '
   '         End If
   '
   '         strTemp = strTemp & "'"
   '         NAD = NAD & TrimSegment(strTemp, "NAD+CN")
   '
   '     End If
   '
   ' End With
   '
   ' 'EX
   ' With m_rstDetails
   '
   '     strTemp = "NAD+EX+" & !Y1                       'Venture Number of the Exporter
   '
   '     If !Y1 = "" Or UCase(Left(!Y1, 2)) <> "BE" Then
   '
   '         strTemp = strTemp & "++" & Left(!Y2, 35)    'Name
   '         strTemp = strTemp & ":" & Mid(!Y2, 36, 35)  'Name
   '         strTemp = strTemp & "+" & !Y3               'Address 1
   '         strTemp = strTemp & ":" & !Y4               'Address 2
   '         strTemp = strTemp & "+" & !Y6               'Municipality
   '         strTemp = strTemp & "+" & !Y7               'State or Province
   '         strTemp = strTemp & "+" & !Y5               'Postal Code
   '         strTemp = strTemp & "+" & !Y8               'Country
   '
   '     End If
   '
   '     strTemp = strTemp & "'"
   '     NAD = NAD & TrimSegment(strTemp, "NAD+EX")
   '
   ' End With
   '
   ' 'WD
   ' With m_rstDetails
   '
   '     strTemp = "NAD+WD+" & !Z1                   'Venture Number of the Warehouse Depositor
   '     strTemp = strTemp & "++" & Left(!Z2, 35)    'Name
   '     strTemp = strTemp & ":" & Mid(!Z2, 36, 35)  'Name
   '     strTemp = strTemp & "+" & !Z3               'Address 1
   '     strTemp = strTemp & ":" & !Z4               'Address 2
   '     strTemp = strTemp & "+" & !Z6               'Municipality
   '     strTemp = strTemp & "+" & !Z7               'State or Province
   '     strTemp = strTemp & "+" & !Z5               'Postal Code
   '     strTemp = strTemp & "'"
   '     NAD = NAD & TrimSegment(strTemp, "NAD+WD")
   '
   ' End With
   '
    'm_lngDetailSegments = m_lngDetailSegments + CountSegments(NAD)

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
        
    If TrimSegment(strTemp, "PAC") <> "" Then
    
        'Goods Description
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
    Dim strN4  As String
    
    strN4 = UCase(FNullField(m_rstDetails.Fields("N4").Value))
    
    If Not (strN4 = "D" Or strN4 = "F" Or strN4 = "G") Then
        
        If LenB(Trim$(FNullField(m_rstDetails.Fields("O2").Value))) > 0 Then
            If Val(m_rstDetails.Fields("O2").Value) <> 0 Then
                'Statistic Value
                strTemp = "MOA+123:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O2").Value))
                
                AddFieldToDigiSign "O2", m_rstDetails, True, m_lngDetailNo
                
                'Currency Code
                strTemp = strTemp & ":" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O6").Value)) & "'"
                
                AddFieldToDigiSign "O6", m_rstDetails, True, m_lngDetailNo
                
                SG33 = SG33 & TrimSegment(strTemp, "MOA+123")
                
                If TrimSegment(strTemp, "MOA+123") <> "" Then
                    'Currency Code
                    strTemp = "CUX+5:" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O6").Value))
                    
                    'Exchange Rate
                    strTemp = strTemp & ":1++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("OB").Value)) & "'"
                    
                    AddFieldToDigiSign "OB", m_rstDetails, True, m_lngDetailNo
                    
                    SG33 = SG33 & TrimSegment(strTemp, "CUX+5::1++")
                
                End If
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
    '    1470    IMD, Item description                       M1
    '    1480    FTX, Free text                              C5
    
    Dim strTemp As String
    
    With m_rstDetailsContainer
        .Filter = adFilterNone
        .Filter = "Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            .Sort = "Ordinal ASC"
            
            .MoveFirst
            
            Do While Not .EOF
                'Container Number
                ' CSCLP-965 - Start
                strTemp = "RFF+AAQ:" & ReplaceSpecialCharacters(FNullField(.Fields("S4").Value)) & "'"
                ' CSCLP-965 - End
                
                AddFieldToDigiSign "S4", m_rstDetailsContainer, True, m_lngDetailNo, , .Fields("Ordinal").Value
                
                SG35 = SG35 & TrimSegment(strTemp, "RFF+AAQ")
                
                'Container Number
                ' CSCLP-965 - Start
                strTemp = "RFF+AAQ:" & ReplaceSpecialCharacters(FNullField(.Fields("S5").Value)) & "'"
                'strTemp = "RFF+AAQ:" & ReplaceSpecialCharacters(.Fields("S5").Value) & "'"
                ' CSCLP-965 - End
                
                SG35 = SG35 & TrimSegment(strTemp, "RFF+AAQ")

                If .Fields("S6").Value = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
        
        .Filter = adFilterNone
    End With
    
    If m_enuDeclarationMode = enuOriginal Then
        With m_rstDetails
            'Unique Consignment Reference Number
            strTemp = "RFF+UCN:" & ReplaceSpecialCharacters(.Fields("LC").Value) & "'"
            
            AddFieldToDigiSign "LC", m_rstDetails, True, m_lngDetailNo
            
            SG35 = SG35 & TrimSegment(strTemp, "RFF+UCN")
                    
            If CheckRegimeType("F", "<>") And CheckRegimeType("G", "<>") Then
            
                'Customs release code
                strTemp = "RFF+AHZ:"
                
                'Regime
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("N1").Value & String(2 - Len(.Fields("N1").Value), "0"))
                
                AddFieldToDigiSign "N1", m_rstDetails, True, m_lngDetailNo
                
                'Previous Regime
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("N2").Value & String(2 - Len(.Fields("N2").Value), "0"))
                
                AddFieldToDigiSign "N2", m_rstDetails, True, m_lngDetailNo
                
                'Community Additional Regulation 1
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("N3").Value & String(3 - Len(.Fields("N3").Value), "0"))
                
                'Community Additional Regulation 2
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("ND").Value & String(3 - Len(.Fields("ND").Value), "0"))
                
                'Community Additional Regulation 3
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("NE").Value & String(3 - Len(.Fields("NE").Value), "0"))
                
                'Digital Signature for Fields N3, ND, and NE
                AddFieldToDigiSign "N3", m_rstDetails, True, m_lngDetailNo
                
                'National Additional Code 1
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("NF").Value & String(3 - Len(.Fields("NF").Value), "0"))
                
                'National Additional Code 2
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("NG").Value & String(3 - Len(.Fields("NG").Value), "0"))
                
                'National Additional Code 3
                strTemp = strTemp & ReplaceSpecialCharacters(.Fields("NH").Value & String(3 - Len(.Fields("NH").Value), "0"))
                
                'Digital Signature for Fields NF, NG, and NH
                AddFieldToDigiSign "NF", m_rstDetails, True, m_lngDetailNo
                
                strTemp = strTemp & "'"
                SG35 = SG35 & TrimSegment(strTemp, "RFF+AHZ:0000000000000000000000")
            End If
        End With
    End If
    
    'UNDG Code
    With m_rstDetailsSensitiveGoods
        .Filter = adFilterNone
        .Filter = " Detail = " & m_lngDetailNo
        
        If .RecordCount > 0 Then
            .Sort = "Ordinal ASC"
            
            .MoveFirst
            
            Do While Not .EOF
                If LenB(Trim$(FNullField(.Fields("SD").Value))) > 0 Then
                    strTemp = "RFF+UN:" & .Fields("SD").Value & "'"
                    
                    AddFieldToDigiSign "SD", m_rstDetailsSensitiveGoods, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    SG35 = SG35 & TrimSegment(strTemp, "RFF+UN")
                    
                    Exit Do
                End If
                
                If .Fields("SE").Value = "E" Then Exit Do
                
                .MoveNext
            Loop
        End If
    End With
    
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
                
                '****************************************************************************************
                'CSCLP-232 IF Q5 = 99 then it is an authorisation document and must be skipped
                '****************************************************************************************
                If Trim$(FNullField(.Fields("Q5").Value)) <> "99" Then
                    'Code
                    strTemp = "DOC+190:::" & ReplaceSpecialCharacters(FNullField(.Fields("Q1").Value))
                    
                    AddFieldToDigiSign "Q1", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    'Reference
                    strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("Q2").Value)) & "'"
                    
                    AddFieldToDigiSign "Q2", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    'strTemp = strTemp & ":1:::" & ReplaceSpecialCharacters(m_strLang) & "'"   'Language
                    
                    SG37 = SG37 & TrimSegment(strTemp, "DOC+190")
                    
                    If TrimSegment(strTemp, "DOC+190:::") <> "" Then
                        'Date
                        strTemp = "DTM+137:" & ReplaceSpecialCharacters(FNullField(.Fields("Q3").Value)) & ":102'"
                        
                        AddFieldToDigiSign "Q3", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        
                        SG37 = SG37 & TrimSegment(strTemp, "DTM+137")
                                            
                        'Office of Validation
                        strTemp = "LOC+91+" & ReplaceSpecialCharacters(FNullField(.Fields("Q4").Value))
                        
                        AddFieldToDigiSign "Q4", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        
                        'Quantity Code
                        strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("QB").Value))
                        
                        AddFieldToDigiSign "QB", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                            
                        'Quantity
                        strTemp = strTemp & ":::" & ReplaceSpecialCharacters(FNullField(.Fields("QC").Value)) & "'"
                        
                        AddFieldToDigiSign "QC", m_rstDetailsDocumenten, True, m_lngDetailNo, , .Fields("Ordinal").Value
                        
                        SG37 = SG37 & TrimSegment(strTemp, "LOC+91")
                        
                        'Edwin - Oct 18
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
                    End If
                End If
                '****************************************************************************************
                
                If FNullField(.Fields("QA").Value) = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
            
        .Filter = adFilterNone
    End With
    
    If m_enuDeclarationMode = enuOriginal Then
        With m_rstDetails
            'Category - Type
            strTemp = "DOC+998:::" & ReplaceSpecialCharacters(FNullField(.Fields("R1").Value)) & "-" & _
                                     ReplaceSpecialCharacters(FNullField(.Fields("R2").Value))
            
            AddFieldToDigiSign "R1", m_rstDetails, True, m_lngDetailNo
            AddFieldToDigiSign "R2", m_rstDetails, True, m_lngDetailNo
            
            If FNullField(.Fields("R2").Value) = "705" Then
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("R5").Value))
                
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                
                'Article Previous Customs Authorities Regulation
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(.Fields("R6").Value))
                
                AddFieldToDigiSign "R6", m_rstDetails, True, m_lngDetailNo
                
            ElseIf FNullField(.Fields("R2").Value) = "740" Or FNullField(.Fields("R2").Value) = "703" Then
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("R5").Value))
                
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                
                'Waybill Number
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(.Fields("R8").Value))
                
                AddFieldToDigiSign "R8", m_rstDetails, True, m_lngDetailNo
            Else
                'Flight Number + Flight Date\FIRMA Code\Verblijfsnummer + IMO-Nummer + Agentcode
                strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("R5").Value))
                
                AddFieldToDigiSign "R5", m_rstDetails, True, m_lngDetailNo
                
                'Article Previous Customs Authorities Regulation
                strTemp = strTemp & "::" & ReplaceSpecialCharacters(FNullField(.Fields("R6").Value))
                
                AddFieldToDigiSign "R6", m_rstDetails, True, m_lngDetailNo
            End If
            
            strTemp = strTemp & "'"
            SG37 = SG37 & TrimSegment(strTemp, "DOC+998")
            
            If TrimSegment(strTemp, "DOC+998") <> "" Then
            
                'CSCLP-231
                'Added R5 != NIHIL condition
                If UCase(FNullField(.Fields("R2").Value)) <> "MNS" And UCase(FNullField(.Fields("R5").Value)) <> "NIHIL" Then
                    'Date
                    strTemp = "DTM+137:" & ReplaceSpecialCharacters(FNullField(.Fields("R3").Value)) & ":102'"
                    
                    AddFieldToDigiSign "R3", m_rstDetails, True, m_lngDetailNo
                    
                    SG37 = SG37 & TrimSegment(strTemp, "DTM+137")

                End If
                
                'Location
                strTemp = "LOC+127+" & ReplaceSpecialCharacters(FNullField(.Fields("R9").Value)) & "'"
                
                AddFieldToDigiSign "R9", m_rstDetails, True, m_lngDetailNo
                
                SG37 = SG37 & TrimSegment(strTemp, "LOC+127")
            
            End If
        
        End With
    
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG37)
    
End Property

Friend Property Get SG38() As String

    '1540    Segment group 38: TOD-LOC-FTX                   C1
    '    1550    TOD, Terms of delivery or transport         M1  TOD+6++<code>
    '    1560    LOC, Place/location identification          C2  LOC+1+ :::<loc name>
    '    1570    FTX, Free text                              C1

    Dim strTemp As String
    Dim strN4  As String
    
    Dim lngGroupCtr As Long
    
    With m_rstDetailsBijzondere
        .Filter = adFilterNone
        .Filter = "Detail = " & m_lngDetailNo
        
        .Sort = "ORDINAL"
        
        ' CSCLP-970 - Start
        '.MoveFirst
        
        If .RecordCount > 0 Then
            .MoveFirst
        ' CSCLP-970 - End
        
            Do While Not .EOF
                lngGroupCtr = lngGroupCtr + 1
            
                'Export Indicator EU
                strTemp = "TOD+2++"
                If Trim(FNullField(.Fields("P3").Value)) <> "" And Not IsNull(.Fields("P3").Value) Then
                    strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("P3").Value))
                    
                    AddFieldToDigiSign "P3", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    strTemp = strTemp & TrimSegment(strTemp, "TOD+2+")
                End If
                
                'Export Indicator Country
                strTemp = strTemp & "+"
                If TrimSegment(strTemp, "TOD+2++") <> "" Then
                    strTemp = strTemp & ReplaceSpecialCharacters(FNullField(.Fields("P4").Value)) & "'"
                    
                    AddFieldToDigiSign "P4", m_rstDetailsBijzondere, True, m_lngDetailNo, , .Fields("Ordinal").Value
                    
                    SG38 = SG38 & TrimSegment(strTemp, "TOD+2+")
                End If
                
                If FNullField(.Fields("P5").Value) = "E" Then Exit Do
                
                .MoveNext
            Loop
            
        End If
        
        .Filter = adFilterNone
    End With
    
    strN4 = UCase(FNullField(m_rstDetails.Fields("N4").Value))
    
    '*Wrong : Kulang ng mga related locations
    If Not (strN4 = "A" Or strN4 = "C" Or strN4 = "D") Then
            
        'Supply Conditions - Type
        strTemp = "TOD+6++" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O3").Value)) & "'"
        
        AddFieldToDigiSign "O3", m_rstDetails, True, m_lngDetailNo
        
        SG38 = SG38 & TrimSegment(strTemp, "TOD+6")
        
        If TrimSegment(strTemp, "TOD+6") <> "" Then
            'Supply Conditions - Place
            strTemp = "LOC+1+:::" & ReplaceSpecialCharacters(FNullField(m_rstDetails.Fields("O4").Value))
            strTemp = strTemp & "'"
            
            AddFieldToDigiSign "O4", m_rstDetails, True, m_lngDetailNo
            
            SG38 = SG38 & TrimSegment(strTemp, "LOC+1")
        End If
    End If

    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG38)
    
End Property

Friend Property Get SG44() As String
    
    '5.1.14.11   Segment Group 44: GIR_GEI-NAD-MEA-SG45-SG47-SG48
    
    Dim strTemp As String
    
    Dim lngGroupCtr As Long
    
    If CheckRegimeType("F", "=") Or _
       CheckRegimeType("G", "=") Then
        
        With m_rstDetailsSensitiveGoods
            .Sort = "ORDINAL"
            
            ' CSCLP-970 - Start
            '.MoveFirst
            
            If .RecordCount > 0 Then
                .MoveFirst
            ' CSCLP-970 - Start
            
                Do While Not .EOF
                    lngGroupCtr = lngGroupCtr + 1
                    
                    'Changes to structure of GIR+3 Segment - Edwin Jan16
                    If IsNull(.Fields("SC").Value) = False Then
                        If Trim(FNullField(.Fields("SC").Value)) <> "" Then
                            strTemp = "GIR+3+" & ReplaceSpecialCharacters(FNullField(.Fields("SC").Value))
                            
                            AddFieldToDigiSign "SC", m_rstDetailsSensitiveGoods, True, m_lngDetailNo, , .Fields("Ordinal").Value
                            
                            strTemp = strTemp & "+" & ReplaceSpecialCharacters(FNullField(.Fields("SB").Value)) & "'"
                            
                            AddFieldToDigiSign "SB", m_rstDetailsSensitiveGoods, True, m_lngDetailNo, , .Fields("Ordinal").Value
                            
                            SG44 = SG44 & TrimSegment(strTemp, "GIR+3+")
                        End If
                    End If
                    
                    If .Fields("SE").Value = "E" Then Exit Do
                    
                    .MoveNext
                Loop
                
            End If
        End With
    End If
    
    m_lngDetailSegments = m_lngDetailSegments + CountSegments(SG44)
    
End Property

Private Sub Class_Initialize()
    m_blnIsAnEcsDeclaration = False
End Sub
