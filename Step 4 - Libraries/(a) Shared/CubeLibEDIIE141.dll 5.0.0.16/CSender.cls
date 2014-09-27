VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_lngSegmentCount As Long
Public m_strTIN As String


Public Function CreateFollowUpRequestEDI(ByRef DataSourceProperties As CDataSourceProperties, _
                                         ByVal UniqueCode As String, _
                                         ByVal DBPath As String, _
                                         ByRef CallingForm As Object) As String
                                 
    
    On Error GoTo ErrHandler
    
    Set G_CallingForm = CallingForm
    
    Call InitializeRecordsets(DataSourceProperties, UniqueCode, DBPath)
    
    CreateFollowUpRequestEDI = CreateEDI()
    
    Call DestroyRecordsets

ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            G_CallingForm.AddtoTraceFollowUpReq "Function: CreateFollowUpRequestEDI" & vbCrLf, True
            G_CallingForm.AddtoTraceFollowUpReq "Error: " & Err.Number & " - " & Err.Description & vbCrLf, True
            
    End Select

End Function


Private Function CreateEDI() As String
    
    Dim strEDIMessage As String
    
    On Error GoTo ErrHandler
    
    m_lngSegmentCount = 0
    m_strTIN = GetMapFunctionValue("F<1 TIN REF>")
    
    strEDIMessage = vbNullString
    strEDIMessage = strEDIMessage & SEGMENTSTART
    strEDIMessage = strEDIMessage & HEADER
    strEDIMessage = strEDIMessage & ENQUIRY
    strEDIMessage = strEDIMessage & COMPETENTAUTHORITY
    strEDIMessage = strEDIMessage & PRESENTATIONOFFICE
    strEDIMessage = strEDIMessage & PRINCIPAL
    strEDIMessage = strEDIMessage & CONSIGNEE
    strEDIMessage = strEDIMessage & SEGMENTEND
    
    CreateEDI = strEDIMessage

ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            G_CallingForm.AddtoTraceFollowUpReq "Function: CreateFollowUpRequestEDI" & vbCrLf, True
            G_CallingForm.AddtoTraceFollowUpReq "Error: " & Err.Number & " - " & Err.Description & vbCrLf, True
            
    End Select

End Function


Friend Property Get SEGMENTSTART() As String
    
    SEGMENTSTART = "UNB+UNOC:3+"
    SEGMENTSTART = SEGMENTSTART & GetMapFunctionValue("F<RECEIVE QUEUE>") & EDI_SEP_DATA_ELEMENT
    SEGMENTSTART = SEGMENTSTART & GetMapFunctionValue("F<RECIPIENT>") & EDI_SEP_DATA_ELEMENT
    SEGMENTSTART = SEGMENTSTART & GetMapFunctionValue("F<DATE, YYMMDD>") & EDI_SEP_COMPOSITE_DATA_ELEMENT
    SEGMENTSTART = SEGMENTSTART & GetMapFunctionValue("F<TIME, HHMM>") & EDI_SEP_DATA_ELEMENT
    SEGMENTSTART = SEGMENTSTART & m_strTIN & "++++++0" + EDI_SEP_SEGMENT
    
    SEGMENTSTART = SEGMENTSTART & "UNH+1+CUSRES:D:96B:UN:BE141A" & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get HEADER() As String
    
    HEADER = "BGM++" & Trim$(FNullField(G_rstFollowUpRequest.Fields("MRN").Value)) & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get PRINCIPAL() As String
    
    PRINCIPAL = "NAD+AF" & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X4") & EDI_SEP_DATA_ELEMENT & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X1") & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X2") & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X3") & EDI_SEP_DATA_ELEMENT & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X6") & EDI_SEP_DATA_ELEMENT
    PRINCIPAL = PRINCIPAL & GetValueFromNCTSFollowUpRequest("X5") & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get COMPETENTAUTHORITY() As String
    
    COMPETENTAUTHORITY = "LOC+120" & EDI_SEP_DATA_ELEMENT
    COMPETENTAUTHORITY = COMPETENTAUTHORITY & GetValueFromNCTSFollowUpRequest("A4") & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get PRESENTATIONOFFICE() As String
    
    PRESENTATIONOFFICE = "LOC+22" & EDI_SEP_DATA_ELEMENT
    PRESENTATIONOFFICE = PRESENTATIONOFFICE & GetValueFromNCTSFollowUpRequest("C3") & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get ENQUIRY() As String
    
    ENQUIRY = "FTX+RET" & EDI_SEP_DATA_ELEMENT
    ENQUIRY = ENQUIRY & GetValueFromNCTSFollowUpRequest("AR") + EDI_SEP_DATA_ELEMENT
    
    If GetValueFromNCTSFollowUpRequest("AR") = "0" Then
        ENQUIRY = ENQUIRY & ":0++" & EDI_SEP_SEGMENT
    Else
        ENQUIRY = ENQUIRY & GetValueFromNCTSFollowUpRequest("AO") & EDI_SEP_COMPOSITE_DATA_ELEMENT
        ENQUIRY = ENQUIRY & GetValueFromNCTSFollowUpRequest("AN") & EDI_SEP_DATA_ELEMENT
        ENQUIRY = ENQUIRY & GetValueFromNCTSFollowUpRequest("AP") & EDI_SEP_DATA_ELEMENT
        ENQUIRY = ENQUIRY & GetValueFromNCTSFollowUpRequest("AQ") & EDI_SEP_SEGMENT
    End If
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get CONSIGNEE() As String
    
    CONSIGNEE = "NAD+CM" & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W6") & EDI_SEP_DATA_ELEMENT & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W1") & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W2") & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W3") & EDI_SEP_DATA_ELEMENT & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W4") & EDI_SEP_DATA_ELEMENT
    CONSIGNEE = CONSIGNEE & GetValueFromNCTSFollowUpRequest("W5") & EDI_SEP_SEGMENT
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
End Property


Friend Property Get SEGMENTEND() As String
    
    m_lngSegmentCount = m_lngSegmentCount + 1
    
    SEGMENTEND = "UNT+" & m_lngSegmentCount & "+1'"
    SEGMENTEND = SEGMENTEND & "UNZ+1+" & m_strTIN & EDI_SEP_SEGMENT
    
End Property

Private Sub InitializeRecordsets(ByRef DataSourceProperties As CDataSourceProperties, _
                                 ByVal UniqueCode As String, _
                                 ByVal DBPath As String)

    Dim strCommand As String
    
    ADOConnectDB G_conEdifact, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    ADOConnectDB G_conData, DataSourceProperties, DBInstanceType_DATABASE_DATA
    ADOConnectDB G_conSadbel, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
    
    'ConnectDB G_conEdifact, DBPath, "edifact.mdb"
    'ConnectDB G_conData, DBPath, "mdb_data.mdb"
    'ConnectDB G_conSadbel, DBPath, "mdb_sadbel.mdb"
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[NCTS DEPARTURE FOLLOW UP REQUEST] "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "Code = '" & ProcessQuotes(UniqueCode) & "' "
    
    ADORecordsetOpen strCommand, G_conEdifact, G_rstFollowUpRequest, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conEdifact, G_rstFollowUpRequest, adOpenKeyset, adLockOptimistic
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "[DATA_NCTS].[LOGID DESCRIPTION] as LogIDDescription, "
    strCommand = strCommand & "[DATA_NCTS].[CODE] as Code, "
    strCommand = strCommand & "[DATA_NCTS].[DOCUMENT NAME] as DocumentName, "
    strCommand = strCommand & "[DATA_NCTS].[MRN] as MRN, "
    strCommand = strCommand & "[DATA_NCTS].[TYPE] as SendMode "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[DATA_NCTS] "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "[DATA_NCTS].[Code] = '" & ProcessQuotes(UniqueCode) & "' "
    
    ADORecordsetOpen strCommand, G_conEdifact, G_rstNCTSData, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conEdifact, G_rstNCTSData, adOpenKeyset, adLockOptimistic
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "[MASTEREDINCTS] "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "Code = '" & ProcessQuotes(UniqueCode) & "' "
    
    ADORecordsetOpen strCommand, G_conData, G_rstMasterEDI, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, G_conData, G_rstMasterEDI, adOpenKeyset, adLockOptimistic
    
End Sub


Private Sub DestroyRecordsets()
    
    ADORecordsetClose G_rstFollowUpRequest
    ADORecordsetClose G_rstMasterEDI
    ADORecordsetClose G_rstNCTSData
    
    'RstClose G_rstFollowUpRequest
    'RstClose G_rstMasterEDI
    'RstClose G_rstNCTSData
    
    ADODisconnectDB G_conEdifact
    ADODisconnectDB G_conData

End Sub

Private Sub Class_Terminate()
    
    Set G_CallingForm = Nothing
    
End Sub

