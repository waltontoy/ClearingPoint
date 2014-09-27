Attribute VB_Name = "modQueries"
Option Explicit
    ' QUERY FOR NCTS DATA
    Const QRY_EDI_NCTS_RECORD_SELECT = "SELECT NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, * "
    Const QRY_EDI_NCTS_RECORD_FROM = "FROM NCTS_IEM INNER JOIN (DATA_NCTS INNER JOIN DATA_NCTS_MESSAGES ON DATA_NCTS.DATA_NCTS_ID = DATA_NCTS_MESSAGES.DATA_NCTS_ID) ON DATA_NCTS_MESSAGES.NCTS_IEM_ID = NCTS_IEM.NCTS_IEM_ID "
    Const QRY_EDI_NCTS_RECORD_WHERE = "WHERE DATA_NCTS.Code = "
    Const QRY_EDI_NCTS_RECORD_WHERE_STATUS_TYPE = " AND DATA_NCTS_MESSAGES.DATA_NCTS_MSG_StatusType = "
    Const QRY_EDI_NCTS_RECORD_WHERE_MESSAGE_ID = " AND DATA_NCTS_MESSAGES.EDI_NCTS_MSG_ID = "
    
    ' QUERY FOR NCTS DATA MESSAGE
    Const QRY_EDI_NCTS_MESSAGE_SELECT = "SELECT NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, DATA_NCTS_MESSAGES.NCTS_IEM_ID AS NCTS_IEM_ID, * "
    Const QRY_EDI_NCTS_MESSAGE_FROM = "FROM DATA_NCTS_MESSAGES INNER JOIN NCTS_IEM ON NCTS_IEM.NCTS_IEM_ID = DATA_NCTS_MESSAGES.NCTS_IEM_ID "
    'Const QRY_EDI_NCTS_MESSAGE_WHERE_STATUS_TYPE = "WHERE DATA_NCTS_MESSAGES.DATA_NCTS_MSG_StatusType = "
    Const QRY_EDI_NCTS_MESSAGE_WHERE_MESSAGE_ID = "WHERE DATA_NCTS_MESSAGES.DATA_NCTS_MSG_ID = "
    
    
    ' QUERY FOR NCTS MESSAGE SEGMENTS
    Global Const QRY_NCTS_MESSAGE_SEGMENTS_SELECT = "SELECT NCTS_IEM_TMS.NCTS_IEM_ID AS NCTS_IEM_ID, NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, NCTS_IEM_TMS.NCTS_IEM_TMS_ParentID, EDI_TMS.EDI_TMS_Occurrence, EDI_TMS.EDI_TMS_Usage, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Description, EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Tag, EDI_TMS.EDI_TMS_Level, EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_ID, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal, NCTS_IEM_TMS.NCTS_IEM_TMS_Occurrence, NCTS_IEM_TMS.NCTS_IEM_TMS_Usage, NCTS_IEM_TMS.NCTS_IEM_TMS_RemarksQualifier "
    Global Const QRY_NCTS_MESSAGE_SEGMENTS_FROM = "FROM NCTS_IEM INNER JOIN (EDI_TMS_SEGMENTS INNER JOIN (NCTS_IEM_TMS INNER JOIN EDI_TMS ON EDI_TMS.EDI_TMS_ID = NCTS_IEM_TMS.EDI_TMS_ID) ON EDI_TMS.EDI_TMS_SEG_ID = EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID) ON NCTS_IEM.NCTS_IEM_ID = NCTS_IEM_TMS.NCTS_IEM_ID "
    Global Const QRY_NCTS_MESSAGE_SEGMENTS_WHERE = "WHERE NCTS_IEM.NCTS_IEM_Name = "
    Global Const QRY_NCTS_MESSAGE_SEGMENTS_ORDER_BY = " ORDER BY EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal "
    
    ' QUERY FOR NCTS DATA MESSAGE STRUCTURE
    Global Const QRY_NCTS_MESSAGE_DATA_SELECT = "SELECT DISTINCT NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, NCTS_IEM_TMS.NCTS_IEM_TMS_ParentID, EDI_TMS.EDI_TMS_Occurrence, EDI_TMS.EDI_TMS_Usage, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Description, EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Tag, EDI_TMS.EDI_TMS_Level, EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_ID, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal, NCTS_IEM_TMS.NCTS_IEM_TMS_Occurrence, NCTS_IEM_TMS.NCTS_IEM_TMS_Usage, NCTS_IEM_TMS.NCTS_IEM_TMS_RemarksQualifier "
    Global Const QRY_NCTS_MESSAGE_DATA_FROM = "FROM DATA_NCTS_MESSAGES INNER JOIN (NCTS_IEM INNER JOIN (EDI_TMS_SEGMENTS INNER JOIN (NCTS_IEM_TMS INNER JOIN EDI_TMS ON EDI_TMS.EDI_TMS_ID = NCTS_IEM_TMS.EDI_TMS_ID) ON EDI_TMS.EDI_TMS_SEG_ID = EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID) ON NCTS_IEM.NCTS_IEM_ID = NCTS_IEM_TMS.NCTS_IEM_ID) ON DATA_NCTS_MESSAGES.NCTS_IEM_ID = NCTS_IEM_TMS.NCTS_IEM_ID "
    Global Const QRY_NCTS_MESSAGE_DATA_WHERE = "WHERE DATA_NCTS_MESSAGES.DATA_NCTS_MSG_ID = "
    Global Const QRY_NCTS_MESSAGE_DATA_ORDER_BY = " ORDER BY EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal "
    
    
    ' QUERY FOR NCTS MESSAGE DATA TABLES
    Global Const QRY_NCTS_MESSAGE_DATA_TABLES_SELECT = "SELECT DISTINCT NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Tag AS SegmentTag "
    Global Const QRY_NCTS_MESSAGE_DATA_TABLES_FROM = "FROM NCTS_IEM INNER JOIN (EDI_TMS_SEGMENTS INNER JOIN (NCTS_IEM_TMS INNER JOIN EDI_TMS ON EDI_TMS.EDI_TMS_ID = NCTS_IEM_TMS.EDI_TMS_ID) ON EDI_TMS.EDI_TMS_SEG_ID = EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID) ON NCTS_IEM.NCTS_IEM_ID = NCTS_IEM_TMS.NCTS_IEM_ID "
    Global Const QRY_NCTS_MESSAGE_DATA_TABLES_WHERE = "WHERE NCTS_IEM.NCTS_IEM_Name = "
    Global Const QRY_NCTS_MESSAGE_DATA_TABLES_ORDER_BY = " ORDER BY EDI_TMS_SEGMENTS.EDI_TMS_SEG_Tag "
    

        
    ' QUERY FOR NCTS BOX ITEM MAP
    Global Const QRY_NCTS_MESSAGE_BOX_ITEM_MAP_SELECT = "SELECT NCTS_IEM.NCTS_IEM_Code AS NCTS_IEM_Code, NCTS_IEM_MAP.NCTS_IEM_MAP_Source, NCTS_IEM_MAP.NCTS_IEM_MAP_StartPosition, NCTS_IEM_MAP.NCTS_IEM_MAP_Length, NCTS_IEM_MAP.NCTS_IEM_MAP_Qualifier, EDI_TMS_SEGMENTS.EDI_TMS_SEG_Tag, NCTS_IEM_MAP.NCTS_IEM_MAP_ITM_ID "
    Global Const QRY_NCTS_MESSAGE_BOX_ITEM_MAP_FROM = "FROM NCTS_IEM INNER JOIN (NCTS_IEM_MAP INNER JOIN EDI_TMS_SEGMENTS ON EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID = NCTS_IEM_MAP.EDI_TMS_SEG_ID) ON NCTS_IEM.NCTS_IEM_ID = NCTS_IEM_MAP.NCTS_IEM_ID "
    Global Const QRY_NCTS_MESSAGE_BOX_ITEM_MAP_WHERE = "WHERE NCTS_IEM.NCTS_IEM_Name = "

        
    ' QUERY FOR NCTS MESSAGE SEGMENT ITEMS
    Global Const QRY_NCTS_MESSAGE_ITEMS_SELECT = "SELECT EDI_TMS_SEGMENTS.EDI_TMS_SEG_TAG, EDI_TMS.EDI_TMS_Level, EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal, NCTS_IEM_TMS.NCTS_IEM_TMS_Occurrence, NCTS_IEM_TMS.NCTS_IEM_TMS_Usage, NCTS_IEM_TMS.NCTS_IEM_TMS_RemarksQualifier, EDI_TMS_ITEMS.EDI_TMS_ITM_Ordinal, "
    Global Const QRY_NCTS_MESSAGE_ITEMS_SELECT_TAG_1 = "NCTS_ITM_***.NCTS_ITM_***_ID, NCTS_ITM_***.NCTS_ITM_***_Description, NCTS_ITM_***.NCTS_ITM_***_Value, NCTS_ITM_***.NCTS_ITM_***_Codelist, NCTS_ITM_***.NCTS_ITM_***_DataType, NCTS_ITM_***.NCTS_ITM_***_Usage "
    Global Const QRY_NCTS_MESSAGE_ITEMS_FROM_TAG_1 = "FROM NCTS_IEM INNER JOIN (EDI_TMS_ITEMS INNER JOIN (NCTS_ITM_*** INNER JOIN "
    Global Const QRY_NCTS_MESSAGE_ITEMS_FROM = "(EDI_TMS_SEGMENTS INNER JOIN (NCTS_IEM_TMS INNER JOIN EDI_TMS ON EDI_TMS.EDI_TMS_ID = NCTS_IEM_TMS.EDI_TMS_ID) ON EDI_TMS.EDI_TMS_SEG_ID = EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID) "
    Global Const QRY_NCTS_MESSAGE_ITEMS_FROM_TAG_2 = "ON NCTS_IEM_TMS.NCTS_IEM_TMS_ID = NCTS_ITM_***.NCTS_IEM_TMS_ID) ON EDI_TMS_ITEMS.EDI_TMS_ITM_ID = NCTS_ITM_***.EDI_TMS_ITM_ID) ON NCTS_IEM.NCTS_IEM_ID = NCTS_IEM_TMS.NCTS_IEM_ID "
    Global Const QRY_NCTS_MESSAGE_ITEMS_WHERE = "WHERE NCTS_IEM.NCTS_IEM_Name = "
    Global Const QRY_NCTS_MESSAGE_ITEMS_ORDER_BY = " ORDER BY EDI_TMS.EDI_TMS_Sequence, NCTS_IEM_TMS.NCTS_IEM_TMS_Ordinal, EDI_TMS_ITEMS.EDI_TMS_ITM_Ordinal "
    
    ' QUERY FOR NCTS MESSAGE TYPES
    Global Const QRY_NCTS_MESSAGE_TYPE_SELECT = "SELECT NCTS_IEM.NCTS_IEM_Name "
    Global Const QRY_NCTS_MESSAGE_TYPE_FROM = "FROM NCTS_IEM "

Public Function GetQryNCTSMessageType(ByVal InternalCode As String, _
                                        ByVal MessageID As Long) _
                                        As String
                                
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_NCTS_MESSAGE_TYPE_SELECT
    strDummy = strDummy & QRY_NCTS_MESSAGE_TYPE_FROM
    
    GetQryNCTSMessageType = strDummy
End Function


Public Function GetQryNCTSData(ByVal InternalCode As String, ByVal MessageID As Long) As String
    Dim strReturnValue As String
    
    strReturnValue = QRY_EDI_NCTS_RECORD_SELECT & _
                     QRY_EDI_NCTS_RECORD_FROM & _
                     QRY_EDI_NCTS_RECORD_WHERE & Chr(39) & InternalCode & Chr(39)
    If MessageID <= 0 Then
        strReturnValue = strReturnValue & QRY_EDI_NCTS_RECORD_WHERE_STATUS_TYPE & Chr(39) & GetMessageStatusType(EMsgStatusType_Document) & Chr(39)
    Else
        strReturnValue = strReturnValue & QRY_EDI_NCTS_RECORD_WHERE_MESSAGE_ID & MessageID
    End If
    GetQryNCTSData = strReturnValue
End Function

Public Function GetQryNCTSDataMessage(ByVal MessageID As Long) As String
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_EDI_NCTS_MESSAGE_SELECT
    strDummy = strDummy & QRY_EDI_NCTS_MESSAGE_FROM
    'strDummy = strDummy & QRY_EDI_NCTS_MESSAGE_WHERE_STATUS_TYPE & Chr(39) & GetMessageStatusType(EMsgStatusType_Document) & Chr(39)
    strDummy = strDummy & QRY_EDI_NCTS_MESSAGE_WHERE_MESSAGE_ID & MessageID
    
    GetQryNCTSDataMessage = strDummy
End Function


Public Function GetQryNCTSTMSDataMessage(ByVal MessageID As Long) As String
    Dim strReturnValue As String
    
    strReturnValue = QRY_NCTS_MESSAGE_DATA_SELECT & _
                     QRY_NCTS_MESSAGE_DATA_FROM & _
                     QRY_NCTS_MESSAGE_DATA_WHERE & MessageID & _
                     QRY_NCTS_MESSAGE_DATA_ORDER_BY
    
    GetQryNCTSTMSDataMessage = strReturnValue
End Function

Public Function GetQryDataNCTSMessages(ByVal DataNCTSID As Long, Optional ByVal MessageStatus As String = vbNullString) As String
    Dim strReturnValue As String
    'strReturnValue = QRY_NCTS_MESSAGE_DATA_SELECT & _
                     QRY_NCTS_MESSAGE_DATA_FROM & _
                     QRY_NCTS_MESSAGE_DATA_WHERE & MessageID & " "
    strReturnValue = QRY_NCTS_MESSAGE_DATA_SELECT & _
                     QRY_NCTS_MESSAGE_DATA_FROM & _
                     "WHERE DATA_NCTS_MESSAGES.DATA_NCTS_ID = " & DataNCTSID & " "
    Select Case MessageStatus
        Case MESSAGE_STATUS_DOCUMENT, MESSAGE_STATUS_QUEUED, MESSAGE_STATUS_RECEIVED, MESSAGE_STATUS_SENT
            strReturnValue = strReturnValue & "AND DATA_NCTS_MESSAGES.DATA_NCTS_MSG_StatusType = '" & MessageStatus & "' "
        Case vbNullString
            '----->  APPEND NOTHING
        Case Else
            '----->  Unidentified message status
            Debug.Assert False
    End Select
    strReturnValue = strReturnValue & QRY_NCTS_MESSAGE_DATA_ORDER_BY
    GetQryDataNCTSMessages = strReturnValue
End Function

Public Function GetQueryDataNCTSMessage(ByVal DataNCTSMessageID As Long)
    Dim strReturnValue As String
    
'    strReturnValue = "SELECT * FROM
    GetQueryDataNCTSMessage = strReturnValue
End Function

Public Function GetQryMessageTechnicalStructure(ByVal NCTSMessageType As ENCTSMessageType)
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_NCTS_MESSAGE_SEGMENTS_SELECT
    strDummy = strDummy & QRY_NCTS_MESSAGE_SEGMENTS_FROM
    strDummy = strDummy & QRY_NCTS_MESSAGE_SEGMENTS_WHERE & Chr(39) & GetNCTSIEMessageCode(NCTSMessageType) & Chr(39)
    strDummy = strDummy & QRY_NCTS_MESSAGE_SEGMENTS_ORDER_BY
    
    GetQryMessageTechnicalStructure = strDummy
End Function


Public Function GetQryMessageDataTables(ByVal NCTSMessageType As ENCTSMessageType)
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_NCTS_MESSAGE_DATA_TABLES_SELECT
    strDummy = strDummy & QRY_NCTS_MESSAGE_DATA_TABLES_FROM
    strDummy = strDummy & QRY_NCTS_MESSAGE_DATA_TABLES_WHERE & Chr(39) & GetNCTSIEMessageCode(NCTSMessageType) & Chr(39)
    strDummy = strDummy & QRY_NCTS_MESSAGE_DATA_TABLES_ORDER_BY
    
    GetQryMessageDataTables = strDummy
End Function

Public Function GetQryMessageItems(ByVal NCTSMessageType As ENCTSMessageType, _
                                    ByVal SegmentType As ESegmentType) As String
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_SELECT
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_SELECT_TAG_1
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_FROM_TAG_1
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_FROM
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_FROM_TAG_2
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_WHERE & Chr(39) & GetNCTSIEMessageCode(NCTSMessageType) & Chr(39)
    strDummy = strDummy & QRY_NCTS_MESSAGE_ITEMS_ORDER_BY
    
    strDummy = Replace(strDummy, "_***", "_" & GetSegmentTag(SegmentType))
    
    GetQryMessageItems = strDummy
End Function

Public Function GetQryMessageBoxItemMap(ByVal NCTSMessageType As ENCTSMessageType)
    Dim strDummy As String
    
    strDummy = ""
    strDummy = strDummy & QRY_NCTS_MESSAGE_BOX_ITEM_MAP_SELECT
    strDummy = strDummy & QRY_NCTS_MESSAGE_BOX_ITEM_MAP_FROM
    strDummy = strDummy & QRY_NCTS_MESSAGE_BOX_ITEM_MAP_WHERE & Chr(39) & GetNCTSIEMessageCode(NCTSMessageType) & Chr(39)
    
    GetQryMessageBoxItemMap = strDummy
End Function

Public Function GetSegmentTag(ByVal SegmentType As ESegmentType) As String
    Select Case SegmentType
        Case ESegment_Type_BGM
            GetSegmentTag = "BGM"
        Case ESegment_Type_CNT
            GetSegmentTag = "CNT"
        Case ESegment_Type_CST
            GetSegmentTag = "CST"
        Case ESegment_Type_DOC
            GetSegmentTag = "DOC"
        Case ESegment_Type_DTM
            GetSegmentTag = "DTM"
        Case ESegment_Type_FTX
            GetSegmentTag = "FTX"
        Case ESegment_Type_GIR
            GetSegmentTag = "GIR"
        Case ESegment_Type_GIS
            GetSegmentTag = "GIS"
        Case ESegment_Type_LOC
            GetSegmentTag = "LOC"
        Case ESegment_Type_MEA
            GetSegmentTag = "MEA"
        Case ESegment_Type_NAD
            GetSegmentTag = "NAD"
        Case ESegment_Type_PAC
            GetSegmentTag = "PAC"
        Case ESegment_Type_PCI
            GetSegmentTag = "PCI"
        Case ESegment_Type_RFF
            GetSegmentTag = "RFF"
        Case ESegment_Type_SEL
            GetSegmentTag = "SEL"
        Case ESegment_Type_TDT
            GetSegmentTag = "TDT"
        Case ESegment_Type_TOD
            GetSegmentTag = "TOD"
        Case ESegment_Type_TPL
            GetSegmentTag = "TPL"
        Case ESegment_Type_UNB
            GetSegmentTag = "UNB"
        Case ESegment_Type_UNH
            GetSegmentTag = "UNH"
        Case ESegment_Type_UNS
            GetSegmentTag = "UNS"
        Case ESegment_Type_UNT
            GetSegmentTag = "UNT"
        Case ESegment_Type_UNZ
            GetSegmentTag = "UNZ"
    End Select
End Function

Public Function GetNCTSIEMessageCode(ByVal NCTSMessageType As ENCTSMessageType) As String
    Select Case NCTSMessageType
        Case EMsg_IE04
            GetNCTSIEMessageCode = "IE04"
        Case EMsg_IE05
            GetNCTSIEMessageCode = "IE05"
        Case EMsg_IE07
            GetNCTSIEMessageCode = "IE07"
        Case EMsg_IE08
            GetNCTSIEMessageCode = "IE08"
        Case EMsg_IE09
            GetNCTSIEMessageCode = "IE09"
        Case EMsg_IE13
            GetNCTSIEMessageCode = "IE13"
        Case EMsg_IE14
            GetNCTSIEMessageCode = "IE14"
        Case EMsg_IE15
            GetNCTSIEMessageCode = "IE15"
        Case EMsg_IE16
            GetNCTSIEMessageCode = "IE16"
        Case EMsg_IE19
            GetNCTSIEMessageCode = "IE19"
        Case EMsg_IE21
            GetNCTSIEMessageCode = "IE21"
        Case EMsg_IE23
            GetNCTSIEMessageCode = "IE23"
        Case EMsg_IE25
            GetNCTSIEMessageCode = "IE25"
        Case EMsg_IE28
            GetNCTSIEMessageCode = "IE28"
        Case EMsg_IE29
            GetNCTSIEMessageCode = "IE29"
        Case EMsg_IE43
            GetNCTSIEMessageCode = "IE43"
        Case EMsg_IE44
            GetNCTSIEMessageCode = "IE44"
        Case EMsg_IE45
            GetNCTSIEMessageCode = "IE45"
        Case EMsg_IE51
            GetNCTSIEMessageCode = "IE51"
        Case EMsg_IE54
            GetNCTSIEMessageCode = "IE54"
        Case EMsg_IE58
            GetNCTSIEMessageCode = "IE58"
        Case EMsg_IE60
            GetNCTSIEMessageCode = "IE60"
        Case EMsg_IE62
            GetNCTSIEMessageCode = "IE62"
        Case EMsg_IE100
            GetNCTSIEMessageCode = "IE100"
        Case EMsg_IE904
            GetNCTSIEMessageCode = "IE904"
        Case EMsg_IE905
            GetNCTSIEMessageCode = "IE905"
        Case EMsg_IE906
            GetNCTSIEMessageCode = "IE906"
        Case EMsg_IE907
            GetNCTSIEMessageCode = "IE907"
        Case EMsg_CODEM
            GetNCTSIEMessageCode = "CODEM"
        Case EMsg_IE917
            GetNCTSIEMessageCode = "IE917"
        Case EMsg_IE55  'IAN 06-04-2005 for IE55 support
            GetNCTSIEMessageCode = "IE55"
        Case EMsg_IE34 'added by Rachelle on Aug312005 for Guarantee Status Request
            GetNCTSIEMessageCode = "IE34"
    End Select
End Function


Public Function GetMessageType(ByVal IEMessageType As String) As ENCTSMessageType
    Dim enuReturnValue As ENCTSMessageType
    
    Select Case Mid(IEMessageType, 3, 3)
        Case "004"
            enuReturnValue = EMsg_IE04
        Case "005"
            enuReturnValue = EMsg_IE05
        Case "007"
            enuReturnValue = EMsg_IE07
        Case "008"
            enuReturnValue = EMsg_IE08
        Case "009"
            enuReturnValue = EMsg_IE09
        Case "013"
            enuReturnValue = EMsg_IE13
        Case "014"
            enuReturnValue = EMsg_IE14
        Case "015"
            enuReturnValue = EMsg_IE15
        Case "016"
            enuReturnValue = EMsg_IE16
        Case "019"
            enuReturnValue = EMsg_IE19
        Case "021"
            enuReturnValue = EMsg_IE21
        Case "023"
            enuReturnValue = EMsg_IE23
        Case "025"
            enuReturnValue = EMsg_IE25
        Case "028"
            enuReturnValue = EMsg_IE28
        Case "029"
            enuReturnValue = EMsg_IE29
        Case "043"
            enuReturnValue = EMsg_IE43
        Case "044"
            enuReturnValue = EMsg_IE44
        Case "045"
            enuReturnValue = EMsg_IE45
        Case "051"
            enuReturnValue = EMsg_IE51
        Case "054"
            enuReturnValue = EMsg_IE54
        Case "058"
            enuReturnValue = EMsg_IE58
        Case "060"
            enuReturnValue = EMsg_IE60
        Case "062"
            enuReturnValue = EMsg_IE62
        Case "100"
            enuReturnValue = EMsg_IE100
        Case "904"
            enuReturnValue = EMsg_IE904
        Case "905"
            enuReturnValue = EMsg_IE905
        Case "906"
            enuReturnValue = EMsg_IE906
        Case "907"
            enuReturnValue = EMsg_IE907
        Case "917"
            enuReturnValue = EMsg_IE917
        Case "055"
            enuReturnValue = EMsg_IE55  'IAN 06-04-2005 for IE55 support
    End Select
    
    GetMessageType = enuReturnValue
End Function
