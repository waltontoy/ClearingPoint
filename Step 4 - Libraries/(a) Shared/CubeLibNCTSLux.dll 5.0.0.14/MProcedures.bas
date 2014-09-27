Attribute VB_Name = "MProcedures"
Option Explicit

'021808
Private m_strMRN As String
Public Const CONST_NCTS_IEM_ID_MSG_IE44 = 11
Public Const CONST_NCTS_IEM_ID_MSG_IE43 = 10

Private m_clsEDINCTSIE44Message As PCubeLibEDIArrivals.cpiIE44Message
Private m_clsEDINCTSIE44Messages As PCubeLibEDIArrivals.cpiIE44Messages

Private m_clsMASTEREDINCTSIE44 As PCubeLibEDIMaster.cpiMASTEREDINCTSIE44
Private m_clsMASTEREDINCTSIE44s As PCubeLibEDIMaster.cpiMASTEREDINCTSIE44s

Private m_conDataDB As ADODB.Connection
Private m_conEdifactHistoryYearDB As ADODB.Connection

Private m_clsEdiNcts2Record As PCubeLibEDIMaster.cpiMasterEdiNcts2
Private m_clsEdiNcts2Records As PCubeLibEDIMaster.cpiMasterEdiNcts2s
Private m_strEdifactHistoryYear As String


'Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Function GetTemporaryFilename() As String
    
    Dim strTemporaryFileName As String
    Dim strWindowsTemporaryPath As String
    
    strWindowsTemporaryPath = WindowsTempPath
    
    ' Create a buffer
    strTemporaryFileName = String(260, 0)
    ' Get a temporary filename
    GetTempFileName AddBackSlashOnPath(strWindowsTemporaryPath), "nct", 0, strTemporaryFileName
        
    ' Remove all the unnecessary chr$(0)'s
    strTemporaryFileName = Left$(strTemporaryFileName, InStr(1, strTemporaryFileName, Chr$(0)) - 1)
    
    ' Set the file attributes
    SetFileAttributes strTemporaryFileName, FILE_ATTRIBUTE_TEMPORARY
        
    GetTemporaryFilename = strTemporaryFileName
    
End Function

' Returns the Windows temporary folder
Public Function WindowsTempPath() As String

    Dim strTempPath As String

    ' Create a buffer
    strTempPath = String(200, Chr(0))

    ' Get the temporary path
    Call GetTempPath(200, strTempPath)

    ' Strip the rest of the buffer
    strTempPath = Left(strTempPath, InStr(strTempPath, Chr(0)) - 1)

    WindowsTempPath = strTempPath



End Function

Public Function CreateXML(ByRef DataSourceProperties As CDataSourceProperties, _
                          ByVal MessageType As NCTS_IEM_IDs) As String
    
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
    
    Select Case MessageType
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE15 'IE15 - Departure
            Set objParentNode = objDOM.appendChild(objDOM.createElement("CC015A"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateXMLMessageIE15(DataSourceProperties, objDOM, objParentNode, objChildNode)
            
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE14 'IE14 - Cancellation Request
            Set objParentNode = objDOM.appendChild(objDOM.createElement("CC014A"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateXMLMessageIE14(DataSourceProperties, objDOM, objParentNode, objChildNode)
            
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE07 'IE07 - Arrival Notification
            Set objParentNode = objDOM.appendChild(objDOM.createElement("CC007A"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateXMLMessageIE07(DataSourceProperties, objDOM, objParentNode, objChildNode)
        
        'p4tric 021408
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE44 'IE44 - Unloading Remarks
            Set objParentNode = objDOM.appendChild(objDOM.createElement("CC044A"))
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            Call CreateXMLMessageIE44(DataSourceProperties, objDOM, objParentNode, objChildNode)

        Case Else
            Debug.Assert False
            
    End Select
    
    CreateXML = objDOM.xml
    
    Set objChildNode = Nothing
    Set objParentNode = Nothing
    Set objDOM = Nothing
    
End Function

Public Sub PrepareEDIDepartureClass(ByRef DataSourceProperties As CDataSourceProperties, _
                                    ByVal UniqueCode As String, _
                                    ByVal MessageID As Long)
        
    Set G_clsEDIDeparture = New EdifactMessage

    'G_clsEDIDeparture.DBLocation = DataSourceProperties.DataSource
    'G_clsEDIDeparture.DBName = "EDIFACT.MDB"
    G_clsEDIDeparture.ConnectEDIDB DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    
    'Prepare Message from Database
    G_clsEDIDeparture.PrepareMessageFromDatabase MessageID, UniqueCode
    
    'Create EDI Message for referencing
    G_strEDIMessage = G_clsEDIDeparture.CreateEDIMessage
    
    'Open Map Recordset
    OpenMapRecordset g_conEdifact, G_rstDepartureMap
    
End Sub


Public Sub PrepareEDICancellationClass(ByRef DataSourceProperties As CDataSourceProperties, _
                                       ByVal UniqueCode As String, _
                                       ByVal DATA_NCTS_ID As Long, _
                                       ByVal CancelationMsgID As Long)
    
    Dim strCommand As String
    Dim rstTemp As ADODB.Recordset
    
    Dim lngMessageID As Long
    Dim lngCurrentIEMID As Long
    
    Dim clsEDICancellation As EdifactMessage
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT * "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "DATA_NCTS_MESSAGES "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "DATA_NCTS_MESSAGES.DATA_NCTS_ID = " & DATA_NCTS_ID & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "(DATA_NCTS_MESSAGES.NCTS_IEM_ID = 3 "
        strCommand = strCommand & "OR "
        strCommand = strCommand & "DATA_NCTS_MESSAGES.NCTS_IEM_ID = 5) "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "DATA_NCTS_MESSAGES.DATA_NCTS_MSG_StatusType = 'Document' "
    ADORecordsetOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    lngCurrentIEMID = 0
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveFirst
        
        Do Until rstTemp.EOF
            If rstTemp.Fields("NCTS_IEM_ID").Value = 5 Then
                lngMessageID = rstTemp.Fields("DATA_NCTS_MSG_ID").Value
            End If
            
            rstTemp.MoveNext
        Loop
    End If
    
    '***************************************************************************************
    'Generate Class to be use to get data for cancellation Message
    '***************************************************************************************
    Set G_clsEDIDeparture = New EdifactMessage

    'G_clsEDIDeparture.DBLocation = G_strMdbPath
    'G_clsEDIDeparture.DBName = "EDIFACT.MDB"
    G_clsEDIDeparture.ConnectEDIDB DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    
    G_clsEDIDeparture.PrepareMessageFromDatabase lngMessageID, UniqueCode
    
    G_strEDIMessage = G_clsEDIDeparture.CreateEDIMessage
    
    OpenMapRecordset g_conEdifact, G_rstDepartureMap
    '***************************************************************************************
    
    '***************************************************************************************
    'Generate EDI Message for Cancellation for referencing
    '***************************************************************************************
    Set clsEDICancellation = New EdifactMessage
    
    'clsEDICancellation.DBLocation = G_strMdbPath
    'clsEDICancellation.DBName = "EDIFACT.MDB"
    clsEDICancellation.ConnectEDIDB DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    
    clsEDICancellation.PrepareMessageFromDatabase CancelationMsgID, UniqueCode
    
    G_strEDICancellation = clsEDICancellation.CreateEDIMessage
    '***************************************************************************************
    
End Sub

Private Sub LoadIE44(ByRef EdifactDB As ADODB.Connection, _
                        ByRef DataDB As ADODB.Connection, _
                        ByRef EdifactHistoryYear As ADODB.Connection, _
                        ByRef MasterEDINCTS2 As PCubeLibEDIMaster.cpiMasterEdiNcts2, _
                        ByVal UniqueCode As String, _
                        Optional ByVal EDIHistoryYear As String = "")
    
    Dim blnFound As Boolean
    Dim lngLoopCtr As Long
    Dim lngNCTSIEMID As Long
        
        
    Set m_clsMASTEREDINCTSIE44 = New cpiMASTEREDINCTSIE44
    Set m_clsMASTEREDINCTSIE44s = New cpiMASTEREDINCTSIE44s
        
    Set m_clsEDINCTSIE44Messages = New PCubeLibEDIArrivals.cpiIE44Messages
    
    m_clsMASTEREDINCTSIE44.FIELD_CODE = UniqueCode
    
    'If EDIHistoryYear = "" Then
    '    blnFound = m_clsMASTEREDINCTSIE44s.SearchRecord(DataDB, _
                                            "CODE", UniqueCode)
    'Else
    '    blnFound = m_clsMASTEREDINCTSIE44s.SearchRecord(EdifactHistoryYear, _
                                            "CODE", UniqueCode)
    'End If
            
    'If (blnFound = True) Then
'       IAN 07-29-2005 For opening from history purposes
     '   If EDIHistoryYear = "" Then
            m_clsMASTEREDINCTSIE44s.GetRecord DataDB, m_clsMASTEREDINCTSIE44
        'Else
        '    m_clsMASTEREDINCTSIE44s.GetRecord EdifactHistoryYear, m_clsMASTEREDINCTSIE44
        'End If
    
    'ElseIf (blnFound = False) Then
    
     '   MapIE07ToIE44 MasterEDINCTS2, m_clsMASTEREDINCTSIE44
    'End If

    Set m_clsEDINCTSIE44Message = m_clsEDINCTSIE44Messages.Add(UniqueCode & "_1-" & CStr(m_clsEDINCTSIE44Messages.Count + 1), UniqueCode)
 
' ... NEED TO INSERT CODE HERE
    
    'Select Case Not blnFound
     '   Case False
            lngNCTSIEMID = CONST_NCTS_IEM_ID_MSG_IE44
     '   Case True
     '       lngNCTSIEMID = CONST_NCTS_IEM_ID_MSG_IE43
    'End Select
    
    GetTags g_conEdifact, m_clsMASTEREDINCTSIE44, G_clsIE44Arrival, UniqueCode, G_clsIE44Arrival.Headers(G_strHeaderKey).MOVEMENT_REFERENCE_NUMBER, lngNCTSIEMID, Not blnFound
    
    Set m_clsEDINCTSIE44Messages = Nothing
    
    Set m_clsEDINCTSIE44Messages = Nothing
End Sub

Private Function GetTags(ByRef EdifactDB As ADODB.Connection, _
                            ByRef MasterEDINCTSIE44 As PCubeLibEDIMaster.cpiMASTEREDINCTSIE44, _
                            ByRef UnloadingRemarksMessage As PCubeLibEDIArrivals.cpiIE44Message, _
                             ByVal UniqueCode As String, _
                             ByVal MRN As String, _
                             ByVal NCTS_IEM_ID As Long, _
                             ByVal NewCode As Boolean) As Boolean
'
    Dim strCommand As String

    Dim lngDATA_NCTS_MSG_ID As Long

    Dim clsDataNctsTable As PCubeLibEDIDataNCTS.cpiDataNctsTable
    Dim clsDataNctsTables As PCubeLibEDIDataNCTS.cpiDataNctsTables
    Dim clsDataNctsMessage As PCubeLibEDIDataNCTS.cpiDataNctsMessage
    Dim clsDataNctsMessages As PCubeLibEDIDataNCTS.cpiDataNctsMessages

    Set clsDataNctsTables = New PCubeLibEDIDataNCTS.cpiDataNctsTables
    Set clsDataNctsMessages = New PCubeLibEDIDataNCTS.cpiDataNctsMessages

        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[DATA_NCTS] "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "[CODE] = " & Chr(39) & ProcessQuotes(UniqueCode) & Chr(39) & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "[MRN] = " & Chr(39) & ProcessQuotes(MRN) & Chr(39) & " "
    ADORecordsetOpen strCommand, EdifactDB, clsDataNctsTables.Recordset, adOpenKeyset, adLockOptimistic
    'Set clsDataNctsTables.Recordset = EdifactDB.Execute(strCommand)

    If (clsDataNctsTables.Recordset.EOF = False) Then

        Set clsDataNctsTable = clsDataNctsTables.GetClassRecord(clsDataNctsTables.Recordset)

        ' get data_ncts_id
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "* "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[DATA_NCTS_MESSAGES] "
            strCommand = strCommand & "WHERE "
            strCommand = strCommand & "[NCTS_IEM_ID] = " & CStr(NCTS_IEM_ID) & " "
            strCommand = strCommand & "AND "
            strCommand = strCommand & "[DATA_NCTS_ID] = " & CStr(clsDataNctsTable.FIELD_DATA_NCTS_ID) & " "
            strCommand = strCommand & "AND "
            strCommand = strCommand & "[DATA_NCTS_MSG_StatusType] = 'Queued'"
        ADORecordsetOpen strCommand, EdifactDB, clsDataNctsMessages.Recordset, adOpenKeyset, adLockOptimistic
        'Set clsDataNctsMessages.Recordset = EdifactDB.Execute(strCommand)

        If (clsDataNctsMessages.Recordset.EOF = False) Then

            Set clsDataNctsMessage = clsDataNctsMessages.GetClassRecord(clsDataNctsMessages.Recordset)

            lngDATA_NCTS_MSG_ID = clsDataNctsMessage.FIELD_DATA_NCTS_MSG_ID

            GetData EdifactDB, lngDATA_NCTS_MSG_ID, MasterEDINCTSIE44, UnloadingRemarksMessage, NCTS_IEM_ID, NewCode

            GetTags = True

        ElseIf (clsDataNctsMessages.Recordset.EOF = True) Then
            GetTags = False
        End If

    ElseIf (clsDataNctsTables.Recordset.EOF = True) Then
        GetTags = False
    End If

    Set clsDataNctsTable = Nothing
    Set clsDataNctsTables = Nothing

    Set clsDataNctsMessages = Nothing
    Set clsDataNctsMessage = Nothing

End Function

Public Sub PrepareIE44ArrivalClass(ByRef DataSourceProperties As CDataSourceProperties, _
                                    ByVal UniqueCode As String)
    
    
    Dim clsIE44ArrivalMessages As New PCubeLibEDIArrivals.cpiIE44Messages
    Set G_clsIE44Arrival = New PCubeLibEDIArrivals.cpiIE44Message
            
    'Pass the Unique Code
    G_clsIE44Arrival.CODE_FIELD = UniqueCode
      
    'Initialize Keys
    G_clsIE44Arrival.Key = UniqueCode & "_1-1"
    G_strHeaderKey = G_clsIE44Arrival.Key & "_1.1-1"
    G_strCustomOfcKey = G_clsIE44Arrival.Key & "_1.3-1"
    G_strTraderKey = G_clsIE44Arrival.Key & "_1.4-1"
    
    'Populate with values for Arrivals
    clsIE44ArrivalMessages.GetRecord g_conEdifact, G_clsIE44Arrival
    
    Set m_conDataDB = New ADODB.Connection
    Set m_conEdifactHistoryYearDB = New ADODB.Connection
    Set m_clsEdiNcts2Record = New PCubeLibEDIMaster.cpiMasterEdiNcts2
    Set m_clsEdiNcts2Records = New PCubeLibEDIMaster.cpiMasterEdiNcts2s
        
    'm_clsEdiNcts2Record.CODE_FIELD = UniqueCode
    'm_clsEdiNcts2Record.MR_FIELD = G_clsIE44Arrival.Headers(G_strHeaderKey).MOVEMENT_REFERENCE_NUMBER
    
    ADOConnectDB m_conDataDB, DataSourceProperties, DBInstanceType_DATABASE_DATA
    'ConnectDB m_conDataDB, G_strMdbPath, "mdb_data.mdb"
    
    m_clsEdiNcts2Records.GetRecord m_conDataDB, m_clsEdiNcts2Record
   
              
    '***************************************************************************************
    'Handle Empty Segments
    '***************************************************************************************
    'Header
    If G_clsIE44Arrival.Headers Is Nothing = True Then
        Set G_clsIE44Arrival.Headers = New PCubeLibEDIArrivals.cpiHeaders
        G_clsIE44Arrival.Headers.Add G_strHeaderKey, UniqueCode
    End If

    'Trader
    If G_clsIE44Arrival.Traders Is Nothing = True Then
        Set G_clsIE44Arrival.Traders = New PCubeLibEDIArrivals.cpiTraders
        G_clsIE44Arrival.Traders.Add G_strTraderKey, UniqueCode
    End If

    'Customs Office
    If G_clsIE44Arrival.CustomOffices Is Nothing = True Then
        Set G_clsIE44Arrival.CustomOffices = New PCubeLibEDIArrivals.cpiCustomOffices
        G_clsIE44Arrival.CustomOffices.Add G_strCustomOfcKey, UniqueCode
    End If

    'unloading remarks
    If G_clsIE44Arrival.UnloadingRemarks Is Nothing = True Then
        Set G_clsIE44Arrival.UnloadingRemarks = New PCubeLibEDIArrivals.cpiUnloadingRemarks
        G_clsIE44Arrival.UnloadingRemarks.Add G_strHeaderKey, UniqueCode
    End If
'
    'result of control
    If G_clsIE44Arrival.ResultOfControls Is Nothing = True Then
        Set G_clsIE44Arrival.ResultOfControls = New PCubeLibEDIArrivals.cpiResultOfControls
        G_clsIE44Arrival.ResultOfControls.Add G_strHeaderKey, UniqueCode
    End If
'
   'seals info
    If G_clsIE44Arrival.SealInfos Is Nothing = True Then
        Set G_clsIE44Arrival.SealInfos = New PCubeLibEDIArrivals.cpiOldSeals_Tables
        G_clsIE44Arrival.SealInfos.Add G_strHeaderKey, UniqueCode, 1
    End If
'
    'goods items
    If G_clsIE44Arrival.GoodsItems Is Nothing = True Then
        Set G_clsIE44Arrival.GoodsItems = New PCubeLibEDIArrivals.cpiGoodsItems
        G_clsIE44Arrival.GoodsItems.Add G_strHeaderKey, UniqueCode
    End If
    
    LoadIE44 g_conEdifact, m_conDataDB, m_conEdifactHistoryYearDB, m_clsEdiNcts2Record, UniqueCode, m_strEdifactHistoryYear

    Set clsIE44ArrivalMessages = Nothing

End Sub
'>>>>>>>>>>>>

Public Sub PrepareEDIArrivalClass(ByVal UniqueCode As String)
    
    Dim clsEDIArrivalMessages As New PCubeLibEDIArrivals.cpiMessages
    
    Set G_clsEDIArrival = New PCubeLibEDIArrivals.cpiMessage
    
    'Pass the Unique Code
    G_clsEDIArrival.CODE_FIELD = UniqueCode
    
    'Initialize Keys
    G_clsEDIArrival.Key = UniqueCode & "_1-1"
    G_strHeaderKey = G_clsEDIArrival.Key & "_1.1-1"
    G_strCustomOfcKey = G_clsEDIArrival.Key & "_1.3-1"
    G_strTraderKey = G_clsEDIArrival.Key & "_1.4-1"
    
    'Populate with values for Arrivals
    clsEDIArrivalMessages.GetRecord g_conEdifact, G_clsEDIArrival
    
    '***************************************************************************************
    'Handle Empty Segments
    '***************************************************************************************
    'Header
    If G_clsEDIArrival.Headers Is Nothing = True Then
        Set G_clsEDIArrival.Headers = New PCubeLibEDIArrivals.cpiHeaders
        G_clsEDIArrival.Headers.Add G_strHeaderKey, UniqueCode
    End If
    
    'Trader
    If G_clsEDIArrival.Traders Is Nothing = True Then
        Set G_clsEDIArrival.Traders = New PCubeLibEDIArrivals.cpiTraders
        G_clsEDIArrival.Traders.Add G_strTraderKey, UniqueCode
    End If
    
    'Customs Office
    If G_clsEDIArrival.CustomOffices Is Nothing = True Then
        Set G_clsEDIArrival.CustomOffices = New PCubeLibEDIArrivals.cpiCustomOffices
        G_clsEDIArrival.CustomOffices.Add G_strCustomOfcKey, UniqueCode
    End If
    
    'EnRouteEvents
    If G_clsEDIArrival.EnRouteEvents Is Nothing = True Then
        Set G_clsEDIArrival.EnRouteEvents = New PCubeLibEDIArrivals.cpiEnRouteEvents
    End If
    '***************************************************************************************
    
    ' Fetch Group Terminator Boxes (like T7: Next/End of Group)
    GetValuesforNonSegmentBox_Detail g_conEdifact, G_clsEDIArrival
    
    Set clsEDIArrivalMessages = Nothing
    
End Sub


Public Sub OpenMapRecordset(ByRef ADOActiveConnection As ADODB.Connection, _
                            ByRef MapRecordset As ADODB.Recordset)
    
    Dim strCommand As String
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT * "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "NCTS_IEM_MAP "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "NCTS_IEM_ID = " & CStr(NCTS_IEM_ID_IE15)
    
    ADORecordsetOpen strCommand, ADOActiveConnection, MapRecordset, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, ADOActiveConnection, MapRecordset, adOpenKeyset, adLockReadOnly
    
End Sub

Public Function GetValueFromClass(ByVal EDIClass As EdifactMessage, _
                                  ByVal MapRecordset As ADODB.Recordset, _
                                  ByVal IE29Value As IE29Values, _
                                  ByVal ValueSource As String, _
                                  ByVal TabNumber As Long) As Variant
    
    Dim varReturnValue                  As Variant
    
    Dim strSegmentKey                   As String
    Dim strSegmentKeyCST                As String
    Dim strBoxCodeValue                 As String
    Dim strLeftSubString                As String
    Dim strRightSubString               As String
    Dim strMiddleSubString              As String
    Dim strSpaces                       As String
    Dim astrSegmentValues()             As String
    
    Dim lngValuesCount                  As Long
    Dim lngBoxCodeInstance              As Long
    Dim lngSegmentInstance              As Long
    Dim lngCST_NCTS_IEM_TMS_ID          As Long
    Dim lngNCTS_IEM_TMS_ID              As Long
    Dim lngNCTS_IEM_TMS_IDAlternative   As Long
    Dim lngDataItemOrdinal              As Long
    Dim lngDataItemsCount               As Long
    
    Dim blnValueIsComplete              As Boolean
    Dim blnContinueLoop                 As Boolean
    Dim blnIsNumeric                    As Boolean
    Dim blnSegmentInHeader              As Boolean
    
    ReDim astrSegmentValues(0)
    astrSegmentValues(0) = ""
    lngValuesCount = 0
    lngDataItemsCount = 1
    
    Call GetValueFromClassAssertions(EDIClass, IE29Value)
    
    Select Case IE29Value
        Case IE29Values.enuIE29Val_NotFromIE29
            lngBoxCodeInstance = 1 'TabNumber
            
            MapRecordset.Filter = "NCTS_IEM_MAP_Source = #" & ValueSource & "#"
            If MapRecordset.RecordCount > 0 Then
                lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                Select Case GetTabType(G_CONST_EDINCTS1_TYPE, ValueSource)
                    Case eTabType.eTab_Header
                        Do
                            If Not MapRecordset.BOF And Not MapRecordset.EOF Then
                                lngSegmentInstance = GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance)
                                blnValueIsComplete = False
                                strSegmentKey = "S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(lngSegmentInstance)
                                If EDIClass.GetSegmentIndex(strSegmentKey) > 0 Then
                                    lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                                    blnValueIsComplete = (MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0)
                                    If MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0 Then
                                        strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                        blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                    Else
                                        Debug.Assert MapRecordset.Fields("NCTS_IEM_MAP_Length").Value > 0
                                        strMiddleSubString = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                        blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                        Debug.Assert Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1
                                        If Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1 Then
                                            strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1) - Len(strMiddleSubString))
                                        Else
                                            strSpaces = ""
                                        End If
                                        strMiddleSubString = strMiddleSubString & strSpaces
                                        strLeftSubString = Left(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1)
                                        strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1) - Len(strLeftSubString))
                                        strLeftSubString = strLeftSubString & strSpaces
                                        strRightSubString = Mid(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_Length").Value + 1)
                                        strBoxCodeValue = strLeftSubString & strMiddleSubString & strRightSubString
                                    End If
                                End If
                            End If
                            
                            If Not blnValueIsComplete Then
                                MapRecordset.MoveNext
                                blnValueIsComplete = MapRecordset.EOF
                                If blnValueIsComplete Then
                                    MapRecordset.MoveFirst
                                End If
                            End If
                            
                            If blnValueIsComplete Then
                                ReDim Preserve astrSegmentValues(lngValuesCount)
                                If blnIsNumeric Then
                                    If Trim(strBoxCodeValue) = vbNullString Then
                                        astrSegmentValues(lngValuesCount) = "0"
                                    Else
                                        astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                    End If
                                Else
                                    astrSegmentValues(lngValuesCount) = strBoxCodeValue
                                End If
                                lngValuesCount = lngValuesCount + 1
                                lngBoxCodeInstance = lngBoxCodeInstance + 1
                            End If
                            
                            blnContinueLoop = (EDIClass.GetSegmentIndex("S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance))) > 0)
                            
                        Loop While blnContinueLoop
                        
                    Case eTabType.eTab_Detail
                        lngBoxCodeInstance = TabNumber
                        Select Case EDIClass.MessageType
                            Case ENCTSMessageType.EMsg_IE15
                                lngCST_NCTS_IEM_TMS_ID = 33
                            Case ENCTSMessageType.EMsg_IE29
                                lngCST_NCTS_IEM_TMS_ID = 126
                            Case ENCTSMessageType.EMsg_IE43
                                lngCST_NCTS_IEM_TMS_ID = 260
                                Debug.Assert False
                            Case ENCTSMessageType.EMsg_IE51
                                lngCST_NCTS_IEM_TMS_ID = 187
                                Debug.Assert False
                            Case ENCTSMessageType.EMsg_IE13
                                lngCST_NCTS_IEM_TMS_ID = 442
                            Case Else
                                Debug.Assert False
                        End Select
                        
                        strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(TabNumber)
                        Debug.Assert EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0
                        If EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0 Then
                            Do
                                If Not MapRecordset.BOF And Not MapRecordset.EOF Then
                                    lngSegmentInstance = GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance)
                                    blnValueIsComplete = False
                                    strSegmentKey = "S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(lngSegmentInstance)
                                    If EDIClass.GetSegmentIndex(strSegmentKey) > 0 Then
                                        If IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or strSegmentKeyCST = strSegmentKey Then
                                            lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                                            blnValueIsComplete = (MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0)
                                            If MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0 Then
                                                strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                                blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                            Else
                                                Debug.Assert MapRecordset.Fields("NCTS_IEM_MAP_Length").Value
                                                strMiddleSubString = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                                blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                                Debug.Assert Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1
                                                If Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1 Then
                                                    strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1) - Len(strMiddleSubString))
                                                Else
                                                    strSpaces = ""
                                                End If
                                                strMiddleSubString = strMiddleSubString & strSpaces
                                                strLeftSubString = Left(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1)
                                                strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1) - Len(strLeftSubString))
                                                strLeftSubString = strLeftSubString & strSpaces
                                                strRightSubString = Mid(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_Length").Value + 1)
                                                strBoxCodeValue = strLeftSubString & strMiddleSubString & strRightSubString
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If Not blnValueIsComplete Then
                                    MapRecordset.MoveNext
                                    blnValueIsComplete = MapRecordset.EOF
                                    If blnValueIsComplete Then
                                        MapRecordset.MoveFirst
                                    End If
                                End If
                                If blnValueIsComplete Then
                                    If IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or strSegmentKeyCST = strSegmentKey Then
                                        ReDim Preserve astrSegmentValues(lngValuesCount)
                                        If blnIsNumeric Then
                                            If Trim(strBoxCodeValue) = vbNullString Then
                                                astrSegmentValues(lngValuesCount) = "0"
                                            Else
                                                astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                            End If
                                        Else
                                            astrSegmentValues(lngValuesCount) = strBoxCodeValue
                                        End If
                                        lngValuesCount = lngValuesCount + 1
                                    End If
                                    lngBoxCodeInstance = lngBoxCodeInstance + 1
                                End If
                                
                                blnContinueLoop = (EDIClass.GetSegmentIndex("S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance))) > 0)
                            Loop Until Not blnContinueLoop
                        End If
                End Select
            Else
                Select Case ValueSource
                    Case "F<DETAIL COUNT>"
                        lngCST_NCTS_IEM_TMS_ID = 33
                        lngSegmentInstance = 0
                        strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance + 1)
                        Do While EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0
                            lngSegmentInstance = lngSegmentInstance + 1
                            strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance + 1)
                        Loop
                        ReDim Preserve astrSegmentValues(0)
                        astrSegmentValues(0) = lngSegmentInstance
                    Case Else
                        Debug.Assert False
                End Select
            End If
        
        Case Else
            Select Case IE29Value
                Case enuIEVal_IE43_Marks_And_Numbers To enuIEVal_IE43_Commodity_Code
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE43(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
                Case enuIE29Val_MessageIdentification To enuIE29Val_ControlledBy
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE29(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
                Case enuIEVal_IE28_TPTIN To enuIEVal_IE28_TPCountry
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE28(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
            End Select
            
    End Select
    
    Dim blnSegmentExists As Boolean
    
    If IE29Value <> enuIE29Val_NotFromIE29 Then
        Select Case GetTabTypeNonIE15(IE29Value)
            Case eTabType.eTab_Header
                lngValuesCount = 0
                lngSegmentInstance = 1
                blnIsNumeric = False
                Do
                    strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                    blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    blnSegmentInHeader = (blnSegmentExists And (lngNCTS_IEM_TMS_IDAlternative > 0))
                    If Not blnSegmentExists And lngNCTS_IEM_TMS_IDAlternative <> 0 Then
                        blnSegmentInHeader = False
                        lngNCTS_IEM_TMS_ID = lngNCTS_IEM_TMS_IDAlternative
                        strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                        blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    End If
                    If blnSegmentExists Then
                            strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                            blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                            ReDim Preserve astrSegmentValues(lngValuesCount)
                            If blnIsNumeric Then
                                If Trim(strBoxCodeValue) = vbNullString Then
                                    astrSegmentValues(lngValuesCount) = "0"
                                Else
                                    astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                End If
                            Else
                                astrSegmentValues(lngValuesCount) = strBoxCodeValue
                            End If
                            lngValuesCount = lngValuesCount + 1
                    End If
                    lngSegmentInstance = lngSegmentInstance + 1
                Loop Until Not blnSegmentExists

                
                
            Case eTabType.eTab_Detail
                Select Case EDIClass.MessageType
                    Case ENCTSMessageType.EMsg_IE15
                        Debug.Assert False
                    Case ENCTSMessageType.EMsg_IE29
                        lngCST_NCTS_IEM_TMS_ID = 126
                    Case ENCTSMessageType.EMsg_IE43
                        lngCST_NCTS_IEM_TMS_ID = 260
                    Case ENCTSMessageType.EMsg_IE51
                        lngCST_NCTS_IEM_TMS_ID = 187
                        Debug.Assert False
                    Case ENCTSMessageType.EMsg_IE28
                        lngCST_NCTS_IEM_TMS_ID = 0
                    Case Else
                        Debug.Assert False
                End Select
                If lngCST_NCTS_IEM_TMS_ID = 0 Then
                    strSegmentKeyCST = ""
                Else
                    strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(TabNumber)
                End If
                lngValuesCount = 0
                lngSegmentInstance = 1
                blnIsNumeric = False
                Do
                    strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                    blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    blnSegmentInHeader = (blnSegmentExists And (lngNCTS_IEM_TMS_IDAlternative > 0))
                    If Not blnSegmentExists And lngNCTS_IEM_TMS_IDAlternative <> 0 Then
                        blnSegmentInHeader = False
                        lngNCTS_IEM_TMS_ID = lngNCTS_IEM_TMS_IDAlternative
                        strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                        blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    End If
                    If blnSegmentExists Then
                        If (IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or (strSegmentKeyCST = strSegmentKey) Or (strSegmentKeyCST = "")) Or blnSegmentInHeader Then
                            strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                            blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                            ReDim Preserve astrSegmentValues(lngValuesCount)
                            If blnIsNumeric Then
                                If Trim(strBoxCodeValue) = vbNullString Then
                                    astrSegmentValues(lngValuesCount) = "0"
                                Else
                                    astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                End If
                            Else
                                astrSegmentValues(lngValuesCount) = strBoxCodeValue
                            End If
                            lngValuesCount = lngValuesCount + 1
                        End If
                    End If
                    lngSegmentInstance = lngSegmentInstance + 1
                Loop Until Not blnSegmentExists
        End Select
    End If
    
    varReturnValue = astrSegmentValues
    GetValueFromClass = varReturnValue
    
End Function


Private Function GetValueFromClassAssertions(ByVal EDIClass As EdifactMessage, _
                                             ByVal IEMValue As IE29Values) As Boolean
    Select Case IEMValue
        Case enuIE29Val_NotFromIE29
            Debug.Assert EDIClass.MessageType = EMsg_IE15
        Case enuIEVal_IE43_Marks_And_Numbers, _
             enuIEVal_IE43_Number_of_Packages, enuIEVal_IE43_Kind_of_Packages, _
             enuIEVal_IE43_Container_Numbers, _
             enuIEVal_IE43_Description_of_Goods, _
             enuIEVal_IE43_Sensitivity_Code, enuIEVal_IE43_Sensitive_Quantity, _
             enuIEVal_IE43_Country_of_Dispatch_Export, enuIEVal_IE43_Country_of_Destination, _
             enuIEVal_IE43_CO_Departure, _
             enuIEVal_IE43_Gross_Mass, enuIEVal_IE43_Net_Mass, _
             enuIEVal_IE43_Additional_Information, _
             enuIEVal_IE43_Consignor_TIN, enuIEVal_IE43_Consignor_Name, enuIEVal_IE43_Consignor_Street_And_Number, enuIEVal_IE43_Consignor_Postal_Code, enuIEVal_IE43_Consignor_City, enuIEVal_IE43_Consignor_Country, _
             enuIEVal_IE43_Consignee_TIN, enuIEVal_IE43_Consignee_Name, enuIEVal_IE43_Consignee_Street_And_Number, enuIEVal_IE43_Consignee_Postal_Code, enuIEVal_IE43_Consignee_City, enuIEVal_IE43_Consignee_Country, _
             enuIEVal_IE43_Document_Type, enuIEVal_IE43_Document_Reference, enuIEVal_IE43_Document_Complement_Information, _
             enuIEVal_IE43_Detail_Number, _
             enuIEVal_IE43_Commodity_Code
            Debug.Assert EDIClass.MessageType = EMsg_IE43
        Case enuIE29Val_MessageIdentification, _
             enuIE29Val_ReferenceNumber, _
             enuIE29Val_AuthorizedLocationOfGoods, _
             enuIE29Val_DeclarationPlace, _
             enuIE29Val_COReferencNumber, enuIE29Val_COName, enuIE29Val_COCountry, enuIE29Val_COStreetAndNumber, enuIE29Val_COPostalCode, enuIE29Val_COCity, enuIE29Val_COLanguage, _
             enuIE29Val_DateApproval, enuIE29Val_DateIssuance, enuIE29Val_DateControl, enuIEVal_IE29_DateLimitTransit, _
             enuIE29Val_ReturnCopy, _
             enuIE29Val_BindingItinerary, _
             enuIE29Val_NotValidForEC, _
             enuIE29Val_TPName, enuIE29Val_TPStreetAndNumber, enuIE29Val_TPCity, enuIE29Val_TPPostalCode, enuIE29Val_TPCountry, _
             enuIE29Val_ControlledBy
            Debug.Assert EDIClass.MessageType = EMsg_IE29
        Case enuIEVal_IE28_TPTIN, enuIEVal_IE28_TPName, enuIEVal_IE28_TPStreetAndNumber, enuIEVal_IE28_TPCity, enuIEVal_IE28_TPPostalCode, enuIEVal_IE28_TPCountry
            Debug.Assert EDIClass.MessageType = EMsg_IE28
    End Select
    
End Function


Public Function GetTabType(ByVal DocumentTypeIdentifier As String, ByVal BoxCode As String) As eTabType
    
    Dim enuReturnValue As eTabType
    
    Select Case DocumentTypeIdentifier
        Case G_CONST_EDINCTS1_TYPE
            Select Case BoxCode
                Case "A4", "A5", "A6", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF", _
                     "B7", "B1", "B8", "B2", "B3", "B9", "BA", "B5", _
                     "C2", "C3", "C4", "C5", _
                     "X4", "X5", "X1", "X2", "X6", "X3", "X7", "X8", _
                     "E1", "EJ", "E3", "EK", "E4", "E5", "E6", "E7", "EM", "EN", "EO", "E8", "EA", "EC", "EE", "EG", "EI"
                    enuReturnValue = eTab_Header
                    
                Case "U6", "U2", "U3", "U4", "U8", "U7", _
                     "W6", "W7", "W1", "W2", "W4", "W3", "W5", _
                     "L7", "L1", "L8", _
                     "M1", "M2", "M9", _
                     "S1", "S2", "S4", "S3", "S5", "S6", "S7", "S8", "S9", "SA", "SB", _
                     "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", _
                     "Y1", "Y2", "Y3", "Y4", "Y5", _
                     "Z1", "Z2", "Z3", "Z4", _
                     "T7"
                    enuReturnValue = eTab_Detail
                    
                Case Else
                    Debug.Assert False
            
            End Select
        
        Case Else
            Debug.Assert False
    
    End Select
    
    GetTabType = enuReturnValue
    
End Function


Private Function GetSegmentInstance(ByVal BoxCode As String, _
                                    ByVal BoxCodeInstance As Long) As Long
    Dim lngReturnValue As Long

    Select Case BoxCode
        Case "A4", "A5", "A6", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF"
            lngReturnValue = BoxCodeInstance
        Case "B7", "B1", "B8", "B2", "B3", "B9", "BA", "B5"
            lngReturnValue = BoxCodeInstance
        Case "C2", "C3", "C4", "C5"
            lngReturnValue = BoxCodeInstance
        Case "X4", "X5", "X1", "X2", "X6", "X3", "X7", "X8"
            lngReturnValue = BoxCodeInstance
        Case "E8", "EA", "EC", "EE", "EG", "EI"
            lngReturnValue = ((BoxCodeInstance - 1) * 6) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Header) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "E8", eTab_Header) + 1)
        Case "E1", "E3", "EK"
            lngReturnValue = BoxCodeInstance
        Case "E4", "E5", "E6", "E7", "EM", "EN"
            lngReturnValue = ((BoxCodeInstance - 1) * 6) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Header) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "E4", eTab_Header) + 1)
        Case "W1" To "W6", "U1" To "U8"
            lngReturnValue = BoxCodeInstance
        Case "V1", "V3", "V5", "V7"
            lngReturnValue = ((BoxCodeInstance - 1) * 4) + (((IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "V1", eTab_Detail)) / 2) + 1)
        Case "V2", "V4", "V6", "V8"
            lngReturnValue = ((BoxCodeInstance - 1) * 4) + (((IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "V2", eTab_Detail)) / 2) + 1)
        Case "L1", "L7", "L8", "M1", "M2", "M9"
            lngReturnValue = BoxCodeInstance
        Case "S1"
            lngReturnValue = BoxCodeInstance
        Case "S2", "S3", "S4"
            lngReturnValue = BoxCodeInstance
        Case "S6", "S7", "S8", "S9", "SA"
            lngReturnValue = ((BoxCodeInstance - 1) * 5) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "S6", eTab_Detail) + 1)
        Case "Y2", "Y3", "Y4"
            lngReturnValue = BoxCodeInstance
        Case "Z1", "Z2", "Z3"
            lngReturnValue = BoxCodeInstance
    End Select
    
    GetSegmentInstance = lngReturnValue
    
End Function


Private Function IsDecendant(ByVal EDIClass As EdifactMessage, _
                             ByVal RootSegmentKey As String, _
                             DecendantSegmentKey As String) As Boolean
    
    Dim blnReturnValue As Boolean
    Dim blnContinueSearch As Boolean
    
    blnReturnValue = False
    
    Debug.Assert EDIClass.GetSegmentIndex(DecendantSegmentKey) > 0
    
    If EDIClass.GetSegmentIndex(DecendantSegmentKey) > 0 Then
        If Trim(EDIClass.Segments(DecendantSegmentKey).KeyParent) <> vbNullString Then
            If EDIClass.Segments(DecendantSegmentKey).KeyParent = RootSegmentKey Then
                blnReturnValue = True
            Else
                blnReturnValue = IsDecendant(EDIClass, RootSegmentKey, EDIClass.Segments(DecendantSegmentKey).KeyParent)
            End If
        Else
            blnReturnValue = False
        End If
    End If
    
    IsDecendant = blnReturnValue
    
End Function


Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE43(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    
    NCTS_IEM_TMS_IDAlternative = 0
    
    Select Case IEMessageValue
        Case enuIEVal_IE43_Marks_And_Numbers                'PCI+28(2,3)
            NCTS_IEM_TMS_ID = 268
            DataItemOrdinal = 2
            DataItemsCount = 2
        Case enuIEVal_IE43_Number_of_Packages, enuIEVal_IE43_Kind_of_Packages
            NCTS_IEM_TMS_ID = 267
            Select Case IEMessageValue
                Case enuIEVal_IE43_Number_of_Packages               'PAC+6(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Kind_of_Packages                 'PAC+6(9)
                    DataItemOrdinal = 9
            End Select
        Case enuIEVal_IE43_Container_Numbers                'RFF+AAQ
            NCTS_IEM_TMS_ID = 269
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Description_of_Goods             'FTX+AAA(6,7,8,9)
            NCTS_IEM_TMS_ID = 261
            DataItemOrdinal = 6
            DataItemsCount = 4
        Case enuIEVal_IE43_Sensitivity_Code, enuIEVal_IE43_Sensitive_Quantity
            NCTS_IEM_TMS_ID = 273
            Select Case IEMessageValue
                Case enuIEVal_IE43_Sensitivity_Code                 'GIR+3+AP(5)
                    DataItemOrdinal = 5
                Case enuIEVal_IE43_Sensitive_Quantity               'GIR+3+AP(2)
                    DataItemOrdinal = 2
            End Select
        Case enuIEVal_IE43_Country_of_Dispatch_Export       'LOC+35(2)
            NCTS_IEM_TMS_ID = 246
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Country_of_Destination           'LOC+36(2)
            NCTS_IEM_TMS_ID = 247
            DataItemOrdinal = 2
        Case enuIEVal_IE43_CO_Departure                     'LOC+118(2)
            NCTS_IEM_TMS_ID = 244
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Gross_Mass                       'MEA+WT+AAB+KGM(7)
            NCTS_IEM_TMS_ID = 263
            DataItemOrdinal = 7
        Case enuIEVal_IE43_Net_Mass                         'MEA+WT+AAA+KGM(7)
            NCTS_IEM_TMS_ID = 264
            DataItemOrdinal = 7
        Case enuIEVal_IE43_Additional_Information           'FTX+ACB(11)
            NCTS_IEM_TMS_ID = 272
            DataItemOrdinal = 11
        Case enuIEVal_IE43_Consignor_TIN, enuIEVal_IE43_Consignor_Name, enuIEVal_IE43_Consignor_Street_And_Number, _
             enuIEVal_IE43_Consignor_Postal_Code, enuIEVal_IE43_Consignor_City, enuIEVal_IE43_Consignor_Country
            NCTS_IEM_TMS_ID = 258
            NCTS_IEM_TMS_IDAlternative = 266
            Select Case IEMessageValue
                Case enuIEVal_IE43_Consignor_TIN                    'NAD+CZ(2)
                    DataItemOrdinal = 2
                Case enuIEVal_IE43_Consignor_Name                   'NAD+CZ(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Consignor_Street_And_Number      'NAD+CZ(16)
                    DataItemOrdinal = 16
                Case enuIEVal_IE43_Consignor_Postal_Code            'NAD+CZ(22)
                    DataItemOrdinal = 22
                Case enuIEVal_IE43_Consignor_City                   'NAD+CZ(20)
                    DataItemOrdinal = 20
                Case enuIEVal_IE43_Consignor_Country                'NAD+CZ(23)
                    DataItemOrdinal = 23
            End Select
        Case enuIEVal_IE43_Consignee_TIN, enuIEVal_IE43_Consignee_Name, enuIEVal_IE43_Consignee_Street_And_Number, _
             enuIEVal_IE43_Consignee_Postal_Code, enuIEVal_IE43_Consignee_City, enuIEVal_IE43_Consignee_Country
            NCTS_IEM_TMS_ID = 256
            NCTS_IEM_TMS_IDAlternative = 265
            Select Case IEMessageValue
                Case enuIEVal_IE43_Consignee_TIN                    'NAD+CN(2)
                    DataItemOrdinal = 2
                Case enuIEVal_IE43_Consignee_Name                   'NAD+CN(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Consignee_Street_And_Number      'NAD+CN(16)
                    DataItemOrdinal = 16
                Case enuIEVal_IE43_Consignee_Postal_Code            'NAD+CN(22)
                    DataItemOrdinal = 22
                Case enuIEVal_IE43_Consignee_City                   'NAD+CN(20)
                    DataItemOrdinal = 20
                Case enuIEVal_IE43_Consignee_Country                'NAD+CN(23)
                    DataItemOrdinal = 23
            End Select
        Case enuIEVal_IE43_Document_Type, enuIEVal_IE43_Document_Reference, enuIEVal_IE43_Document_Complement_Information
            NCTS_IEM_TMS_ID = 270
            Select Case IEMessageValue
                Case enuIEVal_IE43_Document_Type                    'DOC+916(4)
                    DataItemOrdinal = 4
                Case enuIEVal_IE43_Document_Reference               'DOC+916(5)
                    DataItemOrdinal = 5
                Case enuIEVal_IE43_Document_Complement_Information  'DOC+916(7)
                    DataItemOrdinal = 7
            End Select
        Case enuIEVal_IE43_Detail_Number                    'CST(1)
            NCTS_IEM_TMS_ID = 260
            DataItemOrdinal = 1
        Case enuIEVal_IE43_Commodity_Code                   'CST(2)
            NCTS_IEM_TMS_ID = 260
            DataItemOrdinal = 2
        Case Else
            Debug.Assert False
    
    End Select
    
End Sub

Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE29(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    NCTS_IEM_TMS_IDAlternative = 0
    
    Select Case IEMessageValue
        Case enuIE29Val_MessageIdentification               'UNH(1)
            NCTS_IEM_TMS_ID = 83
            DataItemOrdinal = 1
        Case enuIE29Val_ReferenceNumber                     'BGM(5)
            NCTS_IEM_TMS_ID = 84
            DataItemOrdinal = 5
        Case enuIE29Val_AuthorizedLocationOfGoods           'LOC+14(6)
            NCTS_IEM_TMS_ID = 86
            DataItemOrdinal = 6
        Case enuIE29Val_DeclarationPlace                    'LOC+91(5)
            NCTS_IEM_TMS_ID = 93
            DataItemOrdinal = 5
        Case enuIE29Val_COReferencNumber, enuIE29Val_COName, enuIE29Val_COCountry, enuIE29Val_COStreetAndNumber, enuIE29Val_COPostalCode, enuIE29Val_COCity, enuIE29Val_COLanguage
            NCTS_IEM_TMS_ID = 87
            Select Case IEMessageValue
                Case enuIE29Val_COReferencNumber                    'LOC+168(2) - CO = Customs Office
                    DataItemOrdinal = 2
                Case enuIE29Val_COName                              'LOC+168(5) - CO = Customs Office
                    DataItemOrdinal = 5
                Case enuIE29Val_COCountry                           'LOC+168(6) - CO = Customs Office
                    DataItemOrdinal = 6
                Case enuIE29Val_COStreetAndNumber                   'LOC+168(9) - CO = Customs Office
                    DataItemOrdinal = 9
                Case enuIE29Val_COPostalCode                        'LOC+168(10) - CO = Customs Office
                    DataItemOrdinal = 10
                Case enuIE29Val_COCity                              'LOC+168(13) - CO = Customs Office
                    DataItemOrdinal = 13
                Case enuIE29Val_COLanguage                          'LOC+168(14) - CO = Customs Office
                    DataItemOrdinal = 14
            End Select
        Case enuIE29Val_DateApproval                         'DTM+148(2)
            NCTS_IEM_TMS_ID = 95
            DataItemOrdinal = 2
        Case enuIE29Val_DateIssuance                         'DTM+182(2)
            NCTS_IEM_TMS_ID = 96
            DataItemOrdinal = 2
        Case enuIE29Val_DateControl                          'DTM+9(2)
            NCTS_IEM_TMS_ID = 98
            DataItemOrdinal = 2
        Case enuIEVal_IE29_DateLimitTransit                  'DTM+268(2)
            NCTS_IEM_TMS_ID = 97
            DataItemOrdinal = 2
        Case enuIE29Val_ReturnCopy                           'GIS 62(2)
            NCTS_IEM_TMS_ID = 100
            DataItemOrdinal = 2
        Case enuIE29Val_BindingItinerary                     'FTX+ABL(6)
            NCTS_IEM_TMS_ID = 103
            DataItemOrdinal = 6
        Case enuIE29Val_NotValidForEC                        'PCI+19(2)
            NCTS_IEM_TMS_ID = 112
            DataItemOrdinal = 2
        Case enuIE29Val_TPName, enuIE29Val_TPStreetAndNumber, enuIE29Val_TPCity, enuIE29Val_TPPostalCode, enuIE29Val_TPCountry
            NCTS_IEM_TMS_ID = 119
            Select Case IEMessageValue
                Case enuIE29Val_TPName                               'NAD+AF(10) - TP = Transit Principal
                    DataItemOrdinal = 10
                Case enuIE29Val_TPStreetAndNumber                    'NAD+AF(16) - TP = Transit Principal
                    DataItemOrdinal = 16
                Case enuIE29Val_TPCity                               'NAD+AF(20) - TP = Transit Principal
                    DataItemOrdinal = 20
                Case enuIE29Val_TPPostalCode                         'NAD+AF(22) - TP = Transit Principal
                    DataItemOrdinal = 22
                Case enuIE29Val_TPCountry                            'NAD+AF(23) - TP = Transit Principal
                    DataItemOrdinal = 23
            End Select
        
        Case enuIE29Val_ControlledBy                         'NAD+EI(2)
            NCTS_IEM_TMS_ID = 123
            DataItemOrdinal = 2
    
    End Select
    
End Sub

Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE28(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    NCTS_IEM_TMS_IDAlternative = 0
    NCTS_IEM_TMS_ID = 78
    
    Select Case IEMessageValue
        Case enuIEVal_IE28_TPTIN                             'NAD+AF(2)  - TP = Transit Principal
            DataItemOrdinal = 2
        Case enuIEVal_IE28_TPName                            'NAD+AF(10) - TP = Transit Principal
            DataItemOrdinal = 10
        Case enuIEVal_IE28_TPStreetAndNumber                 'NAD+AF(16) - TP = Transit Principal
            DataItemOrdinal = 16
        Case enuIEVal_IE28_TPCity                            'NAD+AF(20) - TP = Transit Principal
            DataItemOrdinal = 20
        Case enuIEVal_IE28_TPPostalCode                      'NAD+AF(22) - TP = Transit Principal
            DataItemOrdinal = 22
        Case enuIEVal_IE28_TPCountry                         'NAD+AF(23) - TP = Transit Principal
            DataItemOrdinal = 23
    End Select
    
End Sub


Private Function GetTabTypeNonIE15(IEMessageValue As IE29Values) As eTabType
    
    Dim enuReturnValue As eTabType
    
    Select Case IEMessageValue
        Case enuIEVal_IE43_Marks_And_Numbers
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Number_of_Packages
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Kind_of_Packages
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Container_Numbers
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Description_of_Goods
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Sensitivity_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Sensitive_Quantity
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Country_of_Dispatch_Export
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Country_of_Destination
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_CO_Departure
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Gross_Mass
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Net_Mass
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Additional_Information
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_TIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Name
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Street_And_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Postal_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_City
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Country
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_TIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Name
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Street_And_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Postal_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_City
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Country
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Type
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Reference
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Complement_Information
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Detail_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Commodity_Code
            enuReturnValue = eTab_Detail
        
        Case enuIE29Val_MessageIdentification
            enuReturnValue = eTab_Detail
        Case enuIE29Val_ReferenceNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_AuthorizedLocationOfGoods
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DeclarationPlace
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COReferencNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COName
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COCountry
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COPostalCode
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COCity
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COLanguage
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateApproval
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateIssuance
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateControl
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE29_DateLimitTransit
            'enuReturnValue = eTab_Detail
            enuReturnValue = eTab_Header
        Case enuIE29Val_ReturnCopy
            enuReturnValue = eTab_Detail
        Case enuIE29Val_BindingItinerary
            enuReturnValue = eTab_Detail
        Case enuIE29Val_NotValidForEC
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPName
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPCity
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPPostalCode
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPCountry
            enuReturnValue = eTab_Detail
        Case enuIE29Val_ControlledBy
            enuReturnValue = eTab_Detail
        
        Case enuIEVal_IE28_TPTIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPName
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPCity
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPPostalCode
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPCountry
            enuReturnValue = eTab_Detail
    End Select
    
    GetTabTypeNonIE15 = enuReturnValue
    
End Function


Public Function IndexTextNCTS(ByVal DocumentTypeIdentifier As String, _
                              ByVal BoxCode As String, _
                              ByVal TabType As eTabType, _
                     Optional ByRef blnNotFound As Boolean) As Integer
        
    Select Case DocumentTypeIdentifier
        Case G_CONST_NCTS1_TYPE, G_CONST_EDINCTS1_TYPE

                If TabType = eTab_Header Then
                    Select Case BoxCode
                        Case "A4": IndexTextNCTS = 0
                        Case "A5": IndexTextNCTS = 1
                        Case "A6": IndexTextNCTS = 2
                        Case "A8": IndexTextNCTS = 3
                        Case "A9": IndexTextNCTS = 4
                        Case "AA": IndexTextNCTS = 5
                        Case "AB": IndexTextNCTS = 6
                        Case "AC": IndexTextNCTS = 7
                        Case "AD": IndexTextNCTS = 8
                        Case "AE": IndexTextNCTS = 9
                        Case "AF": IndexTextNCTS = 10
                        
                        Case "B7": IndexTextNCTS = 11
                        Case "B1": IndexTextNCTS = 12
                        Case "B8": IndexTextNCTS = 13
                        Case "B2": IndexTextNCTS = 14
                        Case "B3": IndexTextNCTS = 15
                        Case "B9": IndexTextNCTS = 16
                        Case "BA": IndexTextNCTS = 17
                        Case "B5": IndexTextNCTS = 18
                                    
                        Case "C2": IndexTextNCTS = 19
                        Case "C3": IndexTextNCTS = 20
                        Case "C4": IndexTextNCTS = 21
                        Case "C5": IndexTextNCTS = 22
                    
                        Case "X4": IndexTextNCTS = 23
                        Case "X5": IndexTextNCTS = 24
                        Case "X1": IndexTextNCTS = 25
                        Case "X2": IndexTextNCTS = 26
                        Case "X6": IndexTextNCTS = 27
                        Case "X3": IndexTextNCTS = 28
                        Case "X7": IndexTextNCTS = 29
                        Case "X8": IndexTextNCTS = 30
            
                        Case "E1": IndexTextNCTS = 31
                        Case "EJ": IndexTextNCTS = 32
                        Case "E3": IndexTextNCTS = 33
                        Case "EK": IndexTextNCTS = 34
                        Case "E4": IndexTextNCTS = 35
                        Case "E5": IndexTextNCTS = 36
                        Case "E6": IndexTextNCTS = 37
                        Case "E7": IndexTextNCTS = 38
                        Case "EM": IndexTextNCTS = 39
                        Case "EN": IndexTextNCTS = 40
                        Case "EO": IndexTextNCTS = 41
                        Case "E8": IndexTextNCTS = 42
                        Case "EA": IndexTextNCTS = 43
                        Case "EC": IndexTextNCTS = 44
                        Case "EE": IndexTextNCTS = 45
                        Case "EG": IndexTextNCTS = 46
                        Case "EI": IndexTextNCTS = 47
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                    End Select
                Else
                    Select Case BoxCode
                        Case "U6": IndexTextNCTS = 0
                        Case "U2": IndexTextNCTS = 1
                        Case "U3": IndexTextNCTS = 2
                        Case "U4": IndexTextNCTS = 3
                        Case "U8": IndexTextNCTS = 4
                        Case "U7": IndexTextNCTS = 5
                        
                        Case "W6": IndexTextNCTS = 6
                        Case "W7": IndexTextNCTS = 7
                        Case "W1": IndexTextNCTS = 8
                        Case "W2": IndexTextNCTS = 9
                        Case "W4": IndexTextNCTS = 10
                        Case "W3": IndexTextNCTS = 11
                        Case "W5": IndexTextNCTS = 12
                        
                        Case "L7": IndexTextNCTS = 13
                        Case "L1": IndexTextNCTS = 14
                        Case "L8": IndexTextNCTS = 15
                        
                        Case "M1": IndexTextNCTS = 16
                        Case "M2": IndexTextNCTS = 17
                        Case "M9": IndexTextNCTS = 18
                        
                        Case "S1": IndexTextNCTS = 19
                        Case "S2": IndexTextNCTS = 20
                        Case "S4": IndexTextNCTS = 21
                        Case "S3": IndexTextNCTS = 22
                        Case "S5": IndexTextNCTS = 23
                        Case "S6": IndexTextNCTS = 24
                        Case "S7": IndexTextNCTS = 25
                        Case "S8": IndexTextNCTS = 26
                        Case "S9": IndexTextNCTS = 27
                        Case "SA": IndexTextNCTS = 28
                        Case "SB": IndexTextNCTS = 29
                        
                        Case "V1": IndexTextNCTS = 30
                        Case "V2": IndexTextNCTS = 31
                        Case "V3": IndexTextNCTS = 32
                        Case "V4": IndexTextNCTS = 33
                        Case "V5": IndexTextNCTS = 34
                        Case "V6": IndexTextNCTS = 35
                        Case "V7": IndexTextNCTS = 36
                        Case "V8": IndexTextNCTS = 37
            
                        Case "Y1": IndexTextNCTS = 38
                        Case "Y2": IndexTextNCTS = 39
                        Case "Y3": IndexTextNCTS = 40
                        Case "Y4": IndexTextNCTS = 41
                        Case "Y5": IndexTextNCTS = 42
                    
                        Case "Z1": IndexTextNCTS = 43
                        Case "Z2": IndexTextNCTS = 44
                        Case "Z3": IndexTextNCTS = 45
                        Case "Z4": IndexTextNCTS = 46
                    
                        Case "T7": IndexTextNCTS = 47
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                    End Select
                End If
                

        Case G_CONST_NCTS2_TYPE

                If TabType = eTab_Header Then
                    Select Case BoxCode
                        Case "A1": IndexTextNCTS = 0
                        Case "A2": IndexTextNCTS = 1
                        Case "A4": IndexTextNCTS = 2
                        Case "A5": IndexTextNCTS = 3
                        Case "A6": IndexTextNCTS = 4
                        Case "A7": IndexTextNCTS = 5
                        Case "A8": IndexTextNCTS = 6
                        Case "A9": IndexTextNCTS = 7
                        Case "AA": IndexTextNCTS = 8
                        Case "AB": IndexTextNCTS = 9
                        Case "AC": IndexTextNCTS = 10
                        Case "AD": IndexTextNCTS = 11
                        Case "AE": IndexTextNCTS = 12
                        Case "AF": IndexTextNCTS = 13
                    
                        Case "B7": IndexTextNCTS = 14
                        Case "B1": IndexTextNCTS = 15
                        Case "B8": IndexTextNCTS = 16
                        Case "B2": IndexTextNCTS = 17
                        Case "B3": IndexTextNCTS = 18
                        Case "B9": IndexTextNCTS = 19
                        Case "B4": IndexTextNCTS = 20
                        Case "BA": IndexTextNCTS = 21
                        Case "B5": IndexTextNCTS = 22
                        Case "B6": IndexTextNCTS = 23
                                                           
                        Case "C1": IndexTextNCTS = 24
                        Case "C2": IndexTextNCTS = 25
                        Case "C3": IndexTextNCTS = 26
                        Case "C4": IndexTextNCTS = 27
                        Case "C5": IndexTextNCTS = 28
                    
                        Case "D1": IndexTextNCTS = 29
                        Case "D2": IndexTextNCTS = 30
                        Case "D3": IndexTextNCTS = 31
                        Case "D4": IndexTextNCTS = 32
                        Case "D5": IndexTextNCTS = 33
                        Case "D6": IndexTextNCTS = 34
                        Case "D7": IndexTextNCTS = 35
                    
                        Case "F1": IndexTextNCTS = 36
                        Case "G1": IndexTextNCTS = 37
                        Case "H1": IndexTextNCTS = 38
                        Case "J1": IndexTextNCTS = 39
                        Case "F2": IndexTextNCTS = 40
                        Case "G2": IndexTextNCTS = 41
                        Case "H2": IndexTextNCTS = 42
                        Case "J2": IndexTextNCTS = 43
                        Case "F3": IndexTextNCTS = 44
                        Case "G3": IndexTextNCTS = 45
                        Case "H3": IndexTextNCTS = 46
                        Case "J3": IndexTextNCTS = 47
                    
                        Case "K1": IndexTextNCTS = 48
                        Case "K2": IndexTextNCTS = 49
                        Case "K3": IndexTextNCTS = 50
                        Case "K4": IndexTextNCTS = 51
                        Case "K5": IndexTextNCTS = 52
                        Case "K6": IndexTextNCTS = 53
                    
                        Case "X4": IndexTextNCTS = 54
                        Case "X5": IndexTextNCTS = 55
                        Case "X1": IndexTextNCTS = 56
                        Case "X2": IndexTextNCTS = 57
                        Case "X6": IndexTextNCTS = 58
                        Case "X3": IndexTextNCTS = 59
                        Case "X7": IndexTextNCTS = 60
                        Case "X8": IndexTextNCTS = 61
            
                        Case "E1": IndexTextNCTS = 62
                        Case "EJ": IndexTextNCTS = 63
                        Case "E3": IndexTextNCTS = 64
                        Case "EK": IndexTextNCTS = 65
                        Case "E4": IndexTextNCTS = 66
                        Case "E5": IndexTextNCTS = 67
                        Case "E6": IndexTextNCTS = 68
                        Case "E7": IndexTextNCTS = 69
                        Case "EM": IndexTextNCTS = 70
                        Case "EN": IndexTextNCTS = 71
                        Case "EO": IndexTextNCTS = 72
                        Case "E8": IndexTextNCTS = 73
                        Case "EA": IndexTextNCTS = 74
                        Case "EC": IndexTextNCTS = 75
                        Case "EE": IndexTextNCTS = 76
                        Case "EG": IndexTextNCTS = 77
                        Case "EI": IndexTextNCTS = 78
                        
                        Case Else: blnNotFound = True
                    End Select
                Else
                    Select Case BoxCode
                        Case "U6": IndexTextNCTS = 0
                        Case "U7": IndexTextNCTS = 1
                        Case "U2": IndexTextNCTS = 2
                        Case "U3": IndexTextNCTS = 3
                        Case "U4": IndexTextNCTS = 4
                        Case "U8": IndexTextNCTS = 5
                        Case "U5": IndexTextNCTS = 6

                        Case "W6": IndexTextNCTS = 7
                        Case "W7": IndexTextNCTS = 8
                        Case "W1": IndexTextNCTS = 9
                        Case "W2": IndexTextNCTS = 10
                        Case "W4": IndexTextNCTS = 11
                        Case "W3": IndexTextNCTS = 12
                        Case "W5": IndexTextNCTS = 13
                                            
                        Case "L1": IndexTextNCTS = 14
                        Case "L2": IndexTextNCTS = 15
                        Case "L3": IndexTextNCTS = 16
                        Case "L4": IndexTextNCTS = 17
                        Case "L5": IndexTextNCTS = 18
                        Case "L6": IndexTextNCTS = 19
                        Case "L8": IndexTextNCTS = 20
                        
                        Case "M1": IndexTextNCTS = 21
                        Case "M2": IndexTextNCTS = 22
                        Case "M9": IndexTextNCTS = 23
                        Case "M3": IndexTextNCTS = 24
                        Case "M4": IndexTextNCTS = 25
                        Case "M5": IndexTextNCTS = 26
                    
                        Case "S1": IndexTextNCTS = 27
                        Case "S2": IndexTextNCTS = 28
                        Case "S4": IndexTextNCTS = 29
                        Case "S3": IndexTextNCTS = 30
                        Case "S5": IndexTextNCTS = 31
                        Case "S6": IndexTextNCTS = 32
                        Case "S7": IndexTextNCTS = 33
                        Case "S8": IndexTextNCTS = 34
                        Case "S9": IndexTextNCTS = 35
                        Case "SA": IndexTextNCTS = 36
                        Case "SB": IndexTextNCTS = 37
                    
                        Case "V1": IndexTextNCTS = 38
                        Case "V2": IndexTextNCTS = 39
                        Case "V3": IndexTextNCTS = 40
                        Case "V4": IndexTextNCTS = 41
                        Case "V5": IndexTextNCTS = 42
                        Case "V6": IndexTextNCTS = 43
                        Case "V7": IndexTextNCTS = 44
                        Case "V8": IndexTextNCTS = 45

                        Case "Y1": IndexTextNCTS = 46
                        Case "Y2": IndexTextNCTS = 47
                        Case "Y3": IndexTextNCTS = 48
                        Case "Y4": IndexTextNCTS = 49
                        Case "Y5": IndexTextNCTS = 50
                    
                        Case "Z1": IndexTextNCTS = 51
                        Case "Z2": IndexTextNCTS = 52
                        Case "Z3": IndexTextNCTS = 53
                        Case "Z4": IndexTextNCTS = 54

                        Case "M6": IndexTextNCTS = 55
                        Case "M7": IndexTextNCTS = 56
                        Case "M8": IndexTextNCTS = 57
                
                        Case "N1": IndexTextNCTS = 58
                        Case "O1": IndexTextNCTS = 59
                        Case "P1": IndexTextNCTS = 60
                        Case "Q1": IndexTextNCTS = 61
                        Case "N2": IndexTextNCTS = 62
                        Case "O2": IndexTextNCTS = 63
                        Case "P2": IndexTextNCTS = 64
                        Case "Q2": IndexTextNCTS = 65
                        Case "N3": IndexTextNCTS = 66
                        Case "O3": IndexTextNCTS = 67
                        Case "P3": IndexTextNCTS = 68
                        Case "Q3": IndexTextNCTS = 69
                    
                        Case "R1": IndexTextNCTS = 70
                        Case "R2": IndexTextNCTS = 71
                        Case "R3": IndexTextNCTS = 72
                        Case "R4": IndexTextNCTS = 73
                        Case "R5": IndexTextNCTS = 74
                        Case "R6": IndexTextNCTS = 75
                        Case "R7": IndexTextNCTS = 76
                        Case "R8": IndexTextNCTS = 77
                        Case "R9": IndexTextNCTS = 78
                        Case "RA": IndexTextNCTS = 79
                        
                        Case "T1": IndexTextNCTS = 80
                        Case "T2": IndexTextNCTS = 81
                        Case "T3": IndexTextNCTS = 82
                        Case "T4": IndexTextNCTS = 83
                        Case "T5": IndexTextNCTS = 84
                        Case "T6": IndexTextNCTS = 85
                        Case "T7": IndexTextNCTS = 86
                        
                        Case Else: blnNotFound = True
                    End Select
                End If
    End Select
    
End Function


Public Function GetMapFunctionValue(ByVal MapFunction As String) As String

    Dim strReturnValue As String
    Dim rstTempADORecordset As ADODB.Recordset    'Dim rstTempADORecordset As DAO.Recordset
    Dim lngLastEDIReference As Long
    Dim strTIN As String
    
    Dim strAE_Value As String
    Dim strAF_Value As String
    
    Dim lngSealsCount As Long
    Dim lngSealsIndex As Long
    
    Dim lngTempSealCtr As Long
    Dim lngTempSealStart As Long
    Dim lngTempSealEnd As Long
    Dim strTempSealStart As String
    Dim strTempSealEnd As String
    Dim strTempZeros As String
    Dim strSealFixedSubstring As String
    
    strReturnValue = vbNullString

    Select Case MapFunction
        Case "F<AE..AF>"
            strAE_Value = GetAEandAF("AE")
            strAF_Value = GetAEandAF("AF")
            
            Select Case GetSealsCreationMode(strAE_Value, strAF_Value)
                Case SealsCreationModes.SealsCreationMode_InvalidInput
                    strReturnValue = strAE_Value & "*****" & strAF_Value & "&&&&&" & CStr(2)
                
                Case SealsCreationModes.SealsCreationMode_NoSeal
                    strReturnValue = "&&&&&" & "0"
                
                Case SealsCreationModes.SealsCreationMode_AEValueOnly
                    strReturnValue = strAE_Value & "&&&&&" & "1"
                
                Case SealsCreationModes.SealsCreationMode_Repetition
                    Debug.Assert IsNumeric(Mid(strAF_Value, 2))
                    If IsNumeric(Mid(strAF_Value, 2)) Then
                        lngSealsCount = CLng(Mid(strAF_Value, 2))
                        For lngSealsIndex = 0 To lngSealsCount - 1
                            strReturnValue = strReturnValue & "*****" & strAE_Value
                        Next lngSealsIndex
                        If InStr(1, strReturnValue, "*****") > 0 Then
                            strReturnValue = Right(strReturnValue, Len(strReturnValue) - 5)
                        End If
                        strReturnValue = strReturnValue & "&&&&&" & lngSealsCount
                    End If
                
                Case SealsCreationModes.SealsCreationMode_Increment
                    
                    If Len(strAE_Value) = Len(strAF_Value) Then
                    
                        For lngTempSealCtr = 1 To Len(strAE_Value)
                            If UCase(Left(strAE_Value, lngTempSealCtr)) <> UCase(Left(strAF_Value, lngTempSealCtr)) Then
                                Exit For
                            End If
                        Next
                        
                        If lngTempSealCtr = Len(strAE_Value) + 1 Then
                            strReturnValue = strAE_Value & "&&&&&" & CStr(1)
                        ElseIf lngTempSealCtr > Len(strAE_Value) + 1 Then
                            strReturnValue = strAE_Value & "*****" & strAF_Value & "&&&&&" & CStr(2)
                        Else
                            strSealFixedSubstring = Left(strAE_Value, lngTempSealCtr - 1)
                            strTempSealStart = Right(strAE_Value, Len(strAE_Value) - Len(strSealFixedSubstring))
                            strTempSealEnd = Right(strAF_Value, Len(strAF_Value) - Len(strSealFixedSubstring))
                            strTempZeros = Space(Len(strAE_Value) - Len(strSealFixedSubstring))
                            strTempZeros = Replace(strTempZeros, " ", "0")
                            
                            If IsNumeric(strTempSealEnd) = False Or IsNumeric(strTempSealStart) = False Then
                                strReturnValue = strAE_Value & "*****" & strAF_Value & "&&&&&" & CStr(2)
                            Else
                                strReturnValue = strAE_Value
                                lngTempSealStart = Val(strTempSealStart)
                                lngTempSealEnd = Val(strTempSealEnd)
                                lngSealsCount = 1
                                
                                For lngTempSealCtr = lngTempSealStart + 1 To lngTempSealEnd
                                    lngSealsCount = lngSealsCount + 1
                                    strReturnValue = strReturnValue & "*****" & strSealFixedSubstring & Format(CStr(lngTempSealCtr), strTempZeros)
                                    If lngSealsCount = 99 Then
                                        Exit For
                                    End If
                                Next
                                
                                strReturnValue = strReturnValue & "&&&&&" & CStr(lngSealsCount)
                            End If
                            
                        End If
                    
                    End If
            End Select
        
        Case "F<DATE, YYMMDD>"
            strReturnValue = Format(Date, "YYMMDD")
        
        Case "F<DATE, YYYYMMDD>"
            strReturnValue = Format(Date, "YYYYMMDD")
        
        Case "F<SEGMENT COUNT>"

        Case "F<MESSAGE REFERENCE>"
            strReturnValue = "1"
        
        Case "F<INTERCHANGE CONTROL COUNT>"
            strReturnValue = "1"
        
        Case "F<RECEIVE QUEUE>"
            '>>BAGOBAGO
            'COMMENTED FOR CSCLP-353
'            Set rstTempADORecordset = G_datSADBEL.OpenRecordset("SELECT * FROM [LOGICAL ID] WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39))
'            If G_strSendMode = "O" Then
'                If Not IsNull(rstTempADORecordset.Fields("SEND NCTS SENDER OPERATIONAL").Value) Then
'                    strReturnValue = rstTempADORecordset.Fields("SEND NCTS SENDER OPERATIONAL").Value
'                End If
'            Else
'                If Not IsNull(rstTempADORecordset.Fields("SEND NCTS SENDER TEST").Value) Then
'                    strReturnValue = rstTempADORecordset.Fields("SEND NCTS SENDER TEST").Value
'                End If
'            End If
'            rstTempADORecordset.Close
'            Set rstTempADORecordset = Nothing
            'CSCLP -353
            Dim strSQL As String
            
                strSQL = vbNullString
                strSQL = strSQL & "SELECT QueueProperties.QueueProp_QueueName "
                strSQL = strSQL & "FROM QueueProperties INNER JOIN [LOGICAL ID] ON "
                strSQL = strSQL & "QueueProperties.QueueProp_Code=[LOGICAL ID].[NCTS_QueuePropCode] WHERE "
                strSQL = strSQL & "QueueProperties.QueueProp_Type = 2 "
                strSQL = strSQL & "AND [LOGICAL ID].[LOGID DESCRIPTION]= " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39)
            ADORecordsetOpen strSQL, g_conSADBEL, rstTempADORecordset, adOpenKeyset, adLockOptimistic
            'Set rstTempADORecordset = G_datSADBEL.OpenRecordset(strSQL)
            If rstTempADORecordset.RecordCount > 0 Then
                rstTempADORecordset.MoveFirst
                If Not IsNull(rstTempADORecordset.Fields("QueueProp_QueueName").Value) Then
                    strReturnValue = rstTempADORecordset.Fields("QueueProp_QueueName").Value
                End If
            End If
            ADORecordsetClose rstTempADORecordset
            'rstTempADORecordset.Close
            'Set rstTempADORecordset = Nothing
            
            '>>
        'original
'            Set rstTempADORecordset = G_datScheduler.OpenRecordset("SELECT EDIPROP_QueueName FROM EDIProperties WHERE EDIPROP_Type = 2")
'            If rstTempADORecordset.RecordCount > 0 Then
'                If Not IsNull(rstTempADORecordset.Fields("EDIPROP_QueueName").Value) Then
'                    strReturnValue = rstTempADORecordset.Fields("EDIPROP_QueueName").Value
'                End If
'            End If
'            rstTempADORecordset.Close
'            Set rstTempADORecordset = Nothing
        
        Case "F<RECIPIENT>"
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "* "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "[LOGICAL ID] "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "[LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39) & " "
            ADORecordsetOpen strSQL, g_conSADBEL, rstTempADORecordset, adOpenKeyset, adLockOptimistic
            'Set rstTempADORecordset = G_datSADBEL.OpenRecordset("SELECT * FROM [LOGICAL ID] WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39))
            If Not (rstTempADORecordset.EOF And rstTempADORecordset.BOF) Then
                rstTempADORecordset.MoveFirst
                
                If G_strSendMode = "O" Then
                    If Not IsNull(rstTempADORecordset.Fields("SEND EDI RECIPIENT OPERATIONAL").Value) Then
                        strReturnValue = rstTempADORecordset.Fields("SEND EDI RECIPIENT OPERATIONAL").Value
                    End If
                Else
                    If Not IsNull(rstTempADORecordset.Fields("SEND EDI RECIPIENT TEST").Value) Then
                        strReturnValue = rstTempADORecordset.Fields("SEND EDI RECIPIENT TEST").Value
                    End If
                End If
            End If
            
            ADORecordsetClose rstTempADORecordset
            'rstTempADORecordset.Close
            'Set rstTempADORecordset = Nothing
        
        Case "F<TIME, HHMM>"
            strReturnValue = Format(Time, "HHMM")

        Case "F<1 TIN REF>"
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "* "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "[LOGICAL ID] "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "[LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39) & " "
            ADORecordsetOpen strSQL, g_conSADBEL, rstTempADORecordset, adOpenKeyset, adLockOptimistic
            'Set rstTempADORecordset = G_datSADBEL.OpenRecordset("SELECT * FROM [LOGICAL ID] WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39))
            If Not (rstTempADORecordset.EOF And rstTempADORecordset.BOF) Then
                rstTempADORecordset.MoveFirst
                
                If Not IsNull(rstTempADORecordset.Fields("TIN").Value) Then
                    strTIN = rstTempADORecordset.Fields("TIN").Value
                Else
                    strTIN = ""
                End If
                If Not IsNull(rstTempADORecordset.Fields("LAST EDI REFERENCE").Value) Then
                    lngLastEDIReference = Val(rstTempADORecordset.Fields("LAST EDI REFERENCE").Value)
                Else
                    lngLastEDIReference = 0
                End If
                lngLastEDIReference = lngLastEDIReference + 1
                If lngLastEDIReference > 99999 Then
                    lngLastEDIReference = 0
                End If
                
                
                'rstTempADORecordset.Edit
                rstTempADORecordset.Fields("LAST EDI REFERENCE").Value = CStr(lngLastEDIReference)
                rstTempADORecordset.Update
                rstTempADORecordset.Close
                
                UpdateRecordset g_conSADBEL, rstTempADORecordset, "LOGID DESCRIPTION"
            End If
            
            ADORecordsetClose rstTempADORecordset
            'Set rstTempADORecordset = Nothing

            strReturnValue = Format(Left(strTIN, 9) & Format(lngLastEDIReference, "00000"), Replace(Space(14), " ", "0"))

        Case "F<2 TIN REF>"
                
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "* "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "[LOGICAL ID] "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "[LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39) & " "
            ADORecordsetOpen strSQL, g_conEdifact, rstTempADORecordset, adOpenKeyset, adLockOptimistic
            'Set rstTempADORecordset = G_datSADBEL.OpenRecordset("SELECT * FROM [LOGICAL ID] WHERE [LOGID DESCRIPTION] = " & Chr(39) & ProcessQuotes(G_strLogicalId) & Chr(39))
            
            If Not (rstTempADORecordset.EOF And rstTempADORecordset.BOF) Then
                rstTempADORecordset.MoveFirst
                
                If Not IsNull(rstTempADORecordset.Fields("TIN").Value) Then
                    strTIN = rstTempADORecordset.Fields("TIN").Value
                Else
                    strTIN = ""
                End If
                If Not IsNull(rstTempADORecordset.Fields("LAST EDI REFERENCE").Value) Then
                    lngLastEDIReference = Val(rstTempADORecordset.Fields("LAST EDI REFERENCE").Value)
                Else
                    lngLastEDIReference = 0
                End If
            End If
            
            ADORecordsetClose rstTempADORecordset
            'rstTempADORecordset.Close
            'Set rstTempADORecordset = Nothing

            strReturnValue = Format(Left(strTIN, 9) & Format(lngLastEDIReference, "00000"), Replace(Space(14), " ", "0"))

        Case Else
            Debug.Assert False
            
    End Select
    
    GetMapFunctionValue = strReturnValue
    
End Function


'Function use to compare if FieldToCheck is the same for all Details
Public Function VentureNumberAreTheSame(ByVal FieldToCheck As String) As Boolean
    
    Dim strTemp As String
    
    Dim lngDetailCount As Long
    Dim lngctr As Long
    
    VentureNumberAreTheSame = True
    strTemp = vbNullString
    
    'Get Number of Details
    lngDetailCount = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "F<DETAIL COUNT>", 0)(0)
    
    If lngDetailCount = 1 Then
        VentureNumberAreTheSame = True
    Else
        strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, FieldToCheck, 1)(0)
        
        For lngctr = 1 To lngDetailCount
            If UCase(Trim(strTemp)) <> GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, FieldToCheck, lngDetailCount)(0) Then
                VentureNumberAreTheSame = False
                Exit For
            End If
        Next lngctr
    End If

End Function


Public Function GetSealsCreationMode(ByVal AEValue As String, _
                                     ByVal AFValue As String) As SealsCreationModes
    
    Dim enuReturnValue As SealsCreationModes
    
    Dim strAFValueSubstring As String
    Dim strSealsMax As String
    Dim strSealsMin As String
    
    Dim lngIndex As Long
    Dim lngIndexMax As Long
    
    Dim blnContinueComparison As Boolean
    
    AEValue = Trim(AEValue)
    AFValue = Trim(AFValue)
    
    enuReturnValue = SealsCreationMode_InvalidInput
    If AEValue = vbNullString Or AEValue = "0" Then
        '----->  No Seals
        enuReturnValue = SealsCreationMode_NoSeal
    Else
        If AFValue = vbNullString Then
            '----->  One seal only - AE Value
            enuReturnValue = SealsCreationMode_AEValueOnly
        Else
            If Left(AFValue, 1) = "*" Then
                strAFValueSubstring = Mid(AFValue, 2)
                If IsNumeric(strAFValueSubstring) Then
                    '----->  takes care of decimal numbers
                    If CLng(strAFValueSubstring) = CDbl(strAFValueSubstring) Then
                        If CLng(strAFValueSubstring) >= 1 And CLng(strAFValueSubstring) <= 99 Then
                            '----->  Repeat Seals
                            enuReturnValue = SealsCreationMode_Repetition
                        End If
                    End If
                End If
            Else
                If Len(AEValue) = Len(AFValue) Then
                    lngIndexMax = Len(AEValue)
                    lngIndex = 0
                    blnContinueComparison = True
                    Do
                        lngIndex = lngIndex + 1
                        blnContinueComparison = ((lngIndex <= lngIndexMax) And (Mid(AEValue, lngIndex, 1) = Mid(AFValue, lngIndex, 1)))
                    Loop Until Not blnContinueComparison
                    If lngIndex <= lngIndexMax Then
                        strSealsMin = Mid(AEValue, lngIndex)
                        strSealsMax = Mid(AFValue, lngIndex)
                        If AllCharactersAreNumeric(strSealsMin) And AllCharactersAreNumeric(strSealsMax) Then
                            If CLng(strSealsMax) - CLng(strSealsMin) > 0 And (CLng(strSealsMax) - CLng(strSealsMin) + 1) <= 99 Then
                                enuReturnValue = SealsCreationMode_Increment
                            End If
                        End If
                    End If
                Else
                    If IsNumeric(AFValue) Then
                        If CDbl(AFValue) = 0 Then
                            enuReturnValue = SealsCreationMode_AEValueOnly
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    GetSealsCreationMode = enuReturnValue
    
End Function

Public Function AllCharactersAreNumeric(ByVal SourceString As String) As Boolean
    
    Dim blnReturnValue As Boolean
    Dim blnContinueLoop As Boolean
    
    Dim lngIndex As Long
    Dim lngSourceStringLength As String
    Dim strCharCheck As String
    
    blnReturnValue = True
    lngIndex = 0
    lngSourceStringLength = Len(SourceString)
    blnContinueLoop = SourceString <> vbNullString
    
    Do While blnContinueLoop
        lngIndex = lngIndex + 1
        strCharCheck = Mid(SourceString, lngIndex, 1)
        blnReturnValue = blnReturnValue And (Asc(strCharCheck) >= Asc("0")) And (Asc(strCharCheck) <= Asc("9"))
        blnContinueLoop = blnReturnValue And (lngIndex < lngSourceStringLength)
    Loop
    
    AllCharactersAreNumeric = blnReturnValue
    
End Function


Public Function GetSealsNumberAndIdentity(ByRef SealsNumber As Long, ByRef SealsIdentity As Variant)
    
    Dim strTemp As String
    Dim arrTemp() As String
    Dim arrReturnValue() As String
    
    Dim lngctr As Long
    
    strTemp = GetMapFunctionValue("F<AE..AF>")
    
    '<SealsIdentity>&&&&&<SealsNumber>
    arrTemp = Split(strTemp, "&&&&&")
    
    If UBound(arrTemp) = 1 Then
        If arrTemp(1) = 0 Then
            SealsNumber = 0
        ElseIf arrTemp(1) = 1 Then
            SealsNumber = 1
            SealsIdentity = arrTemp(0)
        Else
            SealsNumber = arrTemp(1)
            
            '<SealsIdentity1>*****<SealsIdentity2>
            arrTemp = Split(arrTemp(0), "*****")
            
            For lngctr = LBound(arrTemp) To UBound(arrTemp)
                ReDim Preserve arrReturnValue(lngctr)
                arrReturnValue(lngctr) = arrTemp(lngctr)
            Next
            
            SealsIdentity = arrReturnValue
        End If
    End If
    
End Function

Private Sub GetValuesforNonSegmentBox_Detail(ByRef EdifactDB As ADODB.Connection, _
                                             ByRef EDIArrival As PCubeLibEDIArrivals.cpiMessage)
    Dim rstTemp As ADODB.Recordset
    Dim strCommand As String
    
    Dim i As Long
    Dim j As Long
    
    Set rstTemp = New ADODB.Recordset
    
    For i = 1 To EDIArrival.EnRouteEvents.Count
        
        '********************************************************************************************
        'BOX SB
        '********************************************************************************************
            strCommand = vbNullString
            strCommand = strCommand & "SELECT * "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "DATA_NCTS_DETAIL_CONTAINER "
            strCommand = strCommand & "WHERE CODE = '" & EDIArrival.CODE_FIELD & "' "
            strCommand = strCommand & "AND DETAIL = " & i & " "
            strCommand = strCommand & "ORDER BY ORDINAL"
        ADORecordsetOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
        
        For j = 1 To EDIArrival.EnRouteEvents(i).Transhipments(1).Containers.Count / 5
            
            rstTemp.Find "ORDINAL = " & j
            
            If Not rstTemp.EOF Then
                EDIArrival.EnRouteEvents(i).Transhipments(1).Containers(((j - 1)) * 5 + 1).FIELD_SB = IIf(IsNull(rstTemp.Fields("SB").Value), "", Trim(rstTemp.Fields("SB").Value))
            End If
            
        Next
        
        ADORecordsetClose rstTemp
        'RstClose rstTemp
        '********************************************************************************************
        
        '********************************************************************************************
        'BOX T7
        '********************************************************************************************
            strCommand = vbNullString
            strCommand = strCommand & "SELECT * "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "DATA_NCTS_DETAIL "
            strCommand = strCommand & "WHERE "
            strCommand = strCommand & "CODE = '" & EDIArrival.CODE_FIELD & "' "
            strCommand = strCommand & "AND "
            strCommand = strCommand & "DETAIL = " & i & " "
                
        ADORecordsetOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
        'RstOpen strCommand, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
                
        If Not rstTemp.EOF Then
            EDIArrival.EnRouteEvents(i).Ctl_Controls(1).FIELD_T7 = IIf(IsNull(rstTemp.Fields("T7").Value), "", Trim(rstTemp.Fields("T7").Value))
        End If
        
        ADORecordsetClose rstTemp
        'RstClose rstTemp
        '********************************************************************************************
        
    Next
    
    Set rstTemp = Nothing
    
End Sub


Public Function GetAEandAF(ByVal BoxCode As String) As String
    
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "DATA_NCTS_HEADER "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
        
    ADORecordsetOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
    
    On Error Resume Next
    
    If rstTemp.RecordCount > 0 Then
        GetAEandAF = IIf(IsNull(rstTemp.Fields(BoxCode).Value), "", Trim$(rstTemp.Fields(BoxCode).Value))
    Else
        GetAEandAF = vbNullString
    End If
    
    If Err.Number <> 0 Then
        GetAEandAF = vbNullString
    End If
    
    On Error GoTo 0
    
End Function


Public Function GetGroupRecordsFromDataNCTSTables(ByVal TableName As String, _
                                                  ByRef RecordSetToReturn As ADODB.Recordset, _
                                         Optional ByVal DetailNumber As Long = 0)

    Dim strSQL As String
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & TableName & " "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
    
    If DetailNumber > 0 Then
        strSQL = strSQL & "AND Detail = " & DetailNumber
    End If
        
    ADORecordsetOpen strSQL, g_conEdifact, RecordSetToReturn, adOpenKeyset, adLockOptimistic
    'RstOpen strSQL, g_conEdifact, RecordSetToReturn, adOpenKeyset, adLockOptimistic, , True
    
    On Error Resume Next
    RecordSetToReturn.Sort = "Ordinal ASC"
    On Error GoTo 0
    
End Function

Public Function GetSegmentOptionForLanguage(ByVal TagName As String) As Boolean
    
    If g_rstSegment.EOF Or g_rstSegment.BOF Then Exit Function
    
    g_rstSegment.MoveFirst
    g_rstSegment.Find "Segment_TagName = '" & UCase$(TagName) & "' "
    
    'g_rstSegment.Index = "Segment_TagName"
    'g_rstSegment.Seek "=", UCase$(TagName)
    
    If Not g_rstSegment.EOF Then
    'If g_rstSegment.NoMatch = False Then
        GetSegmentOptionForLanguage = g_rstSegment.Fields("Segment_ExcludeLANGWhenEmpty").Value
    Else
        GetSegmentOptionForLanguage = False
    End If
    
End Function

Public Sub InitializeRecorsetsForXML(ByRef DataSourceProperties As CDataSourceProperties)
    Dim strCommand As String
    Dim lngDetailCount As Long
    
    
    ADOConnectDB g_conEdifact, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    'OpenDAODatabase EdifactDAO, G_strMdbPath, "Edifact.mdb"
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT * "
        strCommand = strCommand & "FROM DATA_NCTS "
        strCommand = strCommand & "WHERE CODE = '" & G_strUniqueCode & "'"
            
    ADORecordsetOpen strCommand, g_conEdifact, g_rstEdifact, adOpenKeyset, adLockOptimistic
    'Set rstEdifact = EdifactDAO.OpenRecordset(strCommand)
    
    If Not (g_rstEdifact.EOF And g_rstEdifact.BOF) Then
        strDATA_NCTS_ID = g_rstEdifact.Fields("DATA_NCTS_ID").Value
    End If
    
    ADORecordsetClose g_rstEdifact
    'Set g_rstEdifact = Nothing
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT * "
        strCommand = strCommand & "FROM DATA_NCTS_MESSAGES "
        strCommand = strCommand & "WHERE DATA_NCTS_ID = " & strDATA_NCTS_ID & ""
    ADORecordsetOpen strCommand, g_conEdifact, g_rstEdifact, adOpenKeyset, adLockOptimistic
    'Set g_rstEdifact = EdifactDAO.OpenRecordset(strCommand)
    If Not (g_rstEdifact.EOF And g_rstEdifact.BOF) Then
        strDATA_NCTS_MSG_ID = g_rstEdifact.Fields("DATA_NCTS_MSG_ID").Value
    End If
    
    '###############################################################################HEADER
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_UNB where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & ""
    ADORecordsetOpen strCommand, g_conEdifact, rstUNB, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstUNB, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_UNH where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & ""
    ADORecordsetOpen strCommand, g_conEdifact, rstUNH, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstUNH, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_RFF where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=17 or NCTS_IEM_TMS_ID=18)"
    ADORecordsetOpen strCommand, g_conEdifact, rstRFF, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstRFF, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_BGM where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=3"
    ADORecordsetOpen strCommand, g_conEdifact, rstBGM, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstBGM, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_LOC where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=4 or NCTS_IEM_TMS_ID=5 or NCTS_IEM_TMS_ID=6 or NCTS_IEM_TMS_ID=7 or NCTS_IEM_TMS_ID=57 or NCTS_IEM_TMS_ID=58 or NCTS_IEM_TMS_ID=59 or NCTS_IEM_TMS_ID=60)"
    ADORecordsetOpen strCommand, g_conEdifact, rstLOC, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstLOC, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_TDT where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=22 or NCTS_IEM_TMS_ID=23 or NCTS_IEM_TMS_ID=25)"
    ADORecordsetOpen strCommand, g_conEdifact, rstTDT, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstTDT, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_GIS where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=10"
    ADORecordsetOpen strCommand, g_conEdifact, rstGIS, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstGIS, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_FTX where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=13 or NCTS_IEM_TMS_ID=14 or NCTS_IEM_TMS_ID=15 or NCTS_IEM_TMS_ID=16)"
    ADORecordsetOpen strCommand, g_conEdifact, rstFTX, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstFTX, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_CNT where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=51 or NCTS_IEM_TMS_ID=50 or NCTS_IEM_TMS_ID=53)"
    ADORecordsetOpen strCommand, g_conEdifact, rstCNT, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstCNT, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_MEA where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=11 or NCTS_IEM_TMS_ID=8)"
    ADORecordsetOpen strCommand, g_conEdifact, rstMEA, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstMEA, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_Nad where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=27 or NCTS_IEM_TMS_ID=28 or NCTS_IEM_TMS_ID=29 or NCTS_IEM_TMS_ID=30 or NCTS_IEM_TMS_ID=39)"
    ADORecordsetOpen strCommand, g_conEdifact, rstNAD, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstNAD, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_DTM where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=9 or NCTS_IEM_TMS_ID=8)"
    ADORecordsetOpen strCommand, g_conEdifact, rstDTM, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDTM, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_PAC where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=19"
    ADORecordsetOpen strCommand, g_conEdifact, rstPAC, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstPAC, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_PCI where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=21 order by DATA_NCTS_PCI_Instance ASC"  'allan ncts not sure
    ADORecordsetOpen strCommand, g_conEdifact, rstPCI, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstPCI, adOpenKeyset, adodb.adLockOptimistic, , True
    '###############################################################################HEADER END
    
    '###############################################################################DETAIL
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT TOP 1 "
        strCommand = strCommand & "[DETAIL] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "DATA_NCTS_DETAIL "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "CODE= '" & G_strUniqueCode & "' "
        strCommand = strCommand & "ORDER BY DETAIL DESC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailCount, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailCount, adOpenKeyset, adodb.adLockOptimistic
    
    If Not (rstDetailCount.EOF And rstDetailCount.BOF) Then
        lngDetailCount = rstDetailCount.Fields("DETAIL").Value
    End If
    
    'For lngCtr = 1 To lngDetailCount
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_CST where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=33 order by DATA_NCTS_CST_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailCST, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailCST, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_FTX where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=34 or NCTS_IEM_TMS_ID=47) order by DATA_NCTS_FTX_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailFTX, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailFTX, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_MEA where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=11 or NCTS_IEM_TMS_ID=37 or NCTS_IEM_TMS_ID=38) order by DATA_NCTS_MEA_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailMEA, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailMEA, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_DOC where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=44 order by DATA_NCTS_DOC_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailDOC, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailDOC, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_TOD where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=47 or NCTS_IEM_TMS_ID=46) order by DATA_NCTS_TOD_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailTOD, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailTOD, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_RFF where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and NCTS_IEM_TMS_ID=43 order by DATA_NCTS_RFF_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailRFF, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailRFF, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_PCI where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=42 or NCTS_IEM_TMS_ID=41) order by DATA_NCTS_PCI_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailPCI, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailPCI, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_PAC where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=42 or NCTS_IEM_TMS_ID=41) order by DATA_NCTS_PAC_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailPAC, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailPAC, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_LOC where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=36 or NCTS_IEM_TMS_ID=35 or NCTS_IEM_TMS_ID=4 or NCTS_IEM_TMS_ID=7) order by DATA_NCTS_LOC_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailLOC, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailLOC, adOpenKeyset, adodb.adLockOptimistic, , True
    
        strCommand = vbNullString
        strCommand = strCommand & "Select * from DATA_NCTS_GIR where DATA_NCTS_MSG_ID=" & strDATA_NCTS_MSG_ID & " and (NCTS_IEM_TMS_ID=48) order by DATA_NCTS_GIR_Instance ASC"
    ADORecordsetOpen strCommand, g_conEdifact, rstDetailGIR, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conEdifact, rstDetailGIR, adOpenKeyset, adodb.adLockOptimistic, , True
    
    'Next
    
    '###############################################################################DETAIL END

End Sub

Public Function GetValueIfNotNull(ByVal DataFromDBS As Variant, Optional ByVal BoxCode As String) As String
    
    If IsNull(DataFromDBS) Then
        If BoxCode <> "" And (BoxCode = "V2" Or BoxCode = "V4" Or _
            BoxCode = "V6" Or BoxCode = "V8") Then
            GetValueIfNotNull = "0"
        Else
            GetValueIfNotNull = ""
        End If
    Else
        GetValueIfNotNull = DataFromDBS
    End If
    
End Function

Public Function GetValueForSegment(ByRef rstToFilter As ADODB.Recordset, ByVal Segment As String, ByVal DataNctsIEMTMSID As String, ByVal SequenceNum As String, Optional blnWithInstance As Boolean = False, Optional lngInstance As Long, Optional strBox As String) As String

'#####################################################################
'##                 Get Value From Recordsets                      '##
'##                         CSCLP-439                              '##
'#####################################################################
    
    
    If blnWithInstance = True Then
        
        rstToFilter.Filter = adFilterNone
        rstToFilter.Filter = "[NCTS_IEM_TMS_ID] = " & DataNctsIEMTMSID & " and [DATA_NCTS_" & Segment & "_Instance]= " & lngInstance & ""
        
        If IsNumeric(SequenceNum) Then
            GetValueForSegment = GetValueIfNotNull(rstToFilter.Fields("DATA_NCTS_" & Segment & "_Seq" & SequenceNum).Value, strBox)   'allan ncts
        Else
            GetValueForSegment = GetValueIfNotNull(rstToFilter.Fields("DATA_NCTS_" & Segment & "_" & SequenceNum).Value)   'allan ncts
        End If
    Else
        
        rstToFilter.Filter = adFilterNone
        rstToFilter.Filter = "[NCTS_IEM_TMS_ID] = " & DataNctsIEMTMSID & ""
        
        If IsNumeric(SequenceNum) Then
            GetValueForSegment = GetValueIfNotNull(rstToFilter.Fields("DATA_NCTS_" & Segment & "_Seq" & SequenceNum).Value)   'allan ncts
        Else
            GetValueForSegment = GetValueIfNotNull(rstToFilter.Fields("DATA_NCTS_" & Segment & "_" & SequenceNum).Value)   'allan ncts
        End If
    End If

End Function

