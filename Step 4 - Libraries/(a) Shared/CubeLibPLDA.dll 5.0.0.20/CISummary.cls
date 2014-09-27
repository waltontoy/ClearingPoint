VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CISummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_rstSummary As ADODB.Recordset
Public m_lngDType As Long
Public m_lngSummarySegments As Long
Public m_lngNumberOfDetails As Long
Public m_lngSealCount As Long
Public m_enuDeclarationMode As DECLARATION_MODE

Private m_dblNetTotal As Double
Private m_strIEReference As String
Private m_strRegimeType As String
Private m_blnDigisignOn As Boolean


Public Function CreateSummary(ByRef SegmentCount As Long, _
                              ByVal RecordsSummary As ADODB.Recordset, _
                              ByVal SealCount As Long, _
                              ByVal NetTotal As Double, _
                              ByVal LogTIN As String, _
                              ByVal DType As Long, _
                              ByVal LastEDIRef, _
                              ByVal NumberOfDetails As Long, _
                              ByVal DecMode As DECLARATION_MODE) As String
        
    Dim strMessage As String
    
    Set m_rstSummary = RecordsSummary.Clone
        
    m_lngSummarySegments = SegmentCount
    m_lngSealCount = SealCount
    m_dblNetTotal = NetTotal
    m_lngDType = DType
    
    m_strIEReference = Left(LogTIN, 9) & Format(LastEDIRef, "00000")
    m_lngNumberOfDetails = NumberOfDetails
    m_enuDeclarationMode = DecMode
    
    strMessage = UNS
    
    strMessage = strMessage & CNT
    
    strMessage = strMessage & SG49
    
    strMessage = strMessage & SG50
    
    strMessage = strMessage & UNT
    
    m_blnDigisignOn = False
    
    'Digital Signature Segment
    If g_intDigitalSignatureType <> DigitalSignatureType.None Then
        If Trim(g_strCertificateToUse) <> "" Then
            If g_clsSignData.ValidSigner(g_strCertificateToUse) Then
                strMessage = strMessage & DIGISIGN
                m_blnDigisignOn = True
            End If
        End If
    End If
    
    strMessage = strMessage & UNZ

    SegmentCount = m_lngSummarySegments
    
    Set m_rstSummary = Nothing
            
    'Return the Created Summary EDI Message
    CreateSummary = strMessage
    
End Function

Friend Property Get DIGISIGN() As String
                
    Dim lngDigiSignSegmentCount As Long
    Dim lngCtr As Long
    
    'Digital Signature Segment
    DIGISIGN = "UNH+2+RDRMES:D:04A:UN'"
    DIGISIGN = DIGISIGN & "BGM+++6'"
    DIGISIGN = DIGISIGN & "DTM+ZZZ:" & Format(Now, "DDMMYY") & ":2'"
    DIGISIGN = DIGISIGN & "IDE+9+" & ReplaceSpecialCharacters(FNullField(m_rstSummary.Fields("A3").Value)) & "'"
    DIGISIGN = DIGISIGN & "SCD+9'"
    
    'Send Signature in 512 sizes
    lngDigiSignSegmentCount = Len(g_clsSignData.SignedString) / 512
    
    If (Len(g_clsSignData.SignedString) Mod 512) > 0 Then
        lngDigiSignSegmentCount = lngDigiSignSegmentCount + 1
    End If
    
    For lngCtr = 1 To lngDigiSignSegmentCount
        DIGISIGN = DIGISIGN & "FTX+ZZZ++" & lngCtr & "+" & Mid(g_clsSignData.SignedString, 512 * (lngCtr - 1) + 1, 512) & "'"
    Next
    
    DIGISIGN = DIGISIGN & "UNT+" & CountSegments(DIGISIGN) + 1 & "+2'"
    
End Property

Friend Property Get UNS() As String    'M1

    'Section control
    'UNS+S
    
    UNS = "UNS+S'"
    
    m_lngSummarySegments = m_lngSummarySegments + CountSegments(UNS)
    
End Property

Friend Property Get CNT() As String    'C5

    'Control total
    'CNT+<qualifier>:<amount>
    
    Dim strTemp As String
    
    'If m_enuDeclarationMode = enuOriginal Then 'Change to below to ensure CNT segments are added for ammendment request
    If m_enuDeclarationMode = enuOriginal Then
        'Number of Details
        strTemp = "CNT+5:" & ReplaceSpecialCharacters(m_lngNumberOfDetails) & "'"
        CNT = CNT & TrimSegment(strTemp, "CNT+5")
        
        ' Reims - Jan 18, 2008; Declaraion is Rejected for F and G. This could be the reason.
        'If Not (m_lngDType = 18 And (CheckRegimeType("G", "=") Or CheckRegimeType("F", "="))) Then
            'Number of Packages
            strTemp = "CNT+10:" & ReplaceSpecialCharacters(FNullField(m_rstSummary.Fields("D3").Value)) & "'"
            CNT = CNT & TrimSegment(strTemp, "CNT+10")
        'End If
        
        'Gross Weight
        strTemp = "CNT+26:" & ReplaceSpecialCharacters(FNullField(m_rstSummary.Fields("D1").Value)) & "'"
        CNT = CNT & TrimSegment(strTemp, "CNT+26")
        
        'Net Weight
        strTemp = "CNT+18:" & ReplaceSpecialCharacters(FNullField(m_rstSummary.Fields("D2").Value)) & "'"
        CNT = CNT & TrimSegment(strTemp, "CNT+18")
    End If
        
    'Number of Seals
    strTemp = "CNT+16:" & ReplaceSpecialCharacters(m_lngSealCount) & "'"
    CNT = CNT & TrimSegment(strTemp, "CNT+16:0")
    
    m_lngSummarySegments = m_lngSummarySegments + CountSegments(CNT)
    
End Property

Friend Property Get SG49() As String   'C50
'TAX-MOA-GEI
'1970    TAX, Duty/tax/fee details               M1
'1980    MOA, Monetary amount                    C2
'1990    GEI, Processing information             C1
End Property

Friend Property Get SG50() As String   'C1
'AUT-DTM
'2010    AUT, Authentication result              M1
'2020    DTM, Date/time/period                   C1
End Property

Friend Property Get UNT() As String    'M1

    'Message trailer
    'UNT+<total number of segments>+1
    
    'UNT is included in the Segment Count
    m_lngSummarySegments = m_lngSummarySegments + 1
                                                        
    If m_lngDType = 18 Then
        UNT = "UNT+" & m_lngSummarySegments & "+196688'"
    Else
        UNT = "UNT+" & m_lngSummarySegments & "+1'"
    End If
    
End Property

Friend Property Get UNZ() As String    'M1
    
    Dim lngInterchangeCount As Long
    
    If m_blnDigisignOn = True Then
        lngInterchangeCount = 2
    Else
        lngInterchangeCount = 1
    End If
    
    If m_lngDType = 18 Then
        UNZ = "UNZ+" & lngInterchangeCount & "+196688'"
    Else
        UNZ = "UNZ+" & lngInterchangeCount & "+" & ReplaceSpecialCharacters(m_strIEReference) & "'"
    End If
    
End Property



