VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReleaseMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CONST_MESSAGE_DELIMETER As String = "UNH"
Private Const CONST_SECTION_DELIMETER As String = "UNS"
Private Const CONST_SEGMENT_DELIMETER As String = "'"
Private Const CONST_DATA_ITEM_DELIMETER As String = "+"
Private Const CONST_COMPOSITE_DATA_ITEM_DELIMETER As String = ":"

Private Enum Message_Section
    enuHeader
    enuDetail
    enuSummary
End Enum

'Class properties
Public ReleaseMessage As String     'The EDI string for release
Public DType As Long                'Import = 14, Combined = 18

Private m_rstHeader As ADODB.Recordset
Private m_rstHeaderZegels As ADODB.Recordset
Private m_rstHeaderHandelaars As ADODB.Recordset
Private m_rstDetails As ADODB.Recordset
Private m_rstDetailsBijzondere As ADODB.Recordset
Private m_rstDetailsContainer As ADODB.Recordset
Private m_rstDetailsDocumenten As ADODB.Recordset
Private m_rstDetailsZelf As ADODB.Recordset
Private m_rstDetailsHandelaars As ADODB.Recordset
Private m_rstDetailsBerekeningsEenheden As ADODB.Recordset
Private m_strLastGroupRootSegment As String
Private m_lngDetailNo As Long

Public Sub SaveReleaseMeassageToDB(ByRef DataSourceProperties As CDataSourceProperties, _
                                   ByVal Code As String)
    
    Dim astrPerMessage() As String      'Used to store several messages (multiple UNH segments)
    Dim astrPerSection() As String      'Used to store the header, details and summary sections per message
    Dim astrPerSegment() As String      'Used to store the segments per Section
    
    Dim strCUSDEC As String             'Used to store the CUSDEC from the release message
    Dim lngKownterPerSection As Long
    Dim lngKownterPerSegment As Long
           
    Dim strSQLH As String
    Dim strSQLHZegels As String
    Dim strSQLD As String
    Dim strSQLDBijzondere As String
    Dim strSQLDContainer As String
    Dim strSQLDDocumenten As String
    Dim strSQLDZelf As String
    Dim strSQLDBijgevoegde As String
    Dim strSQLHHandelaars As String
    Dim strSQLDHandelaars As String
    Dim strSQLDBerekeningsEenheden As String
    
    
    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
    
Step_1: 'Check if the properties are set
    If Len(ReleaseMessage) = 0 Then GoTo Step_4
    If DType = 0 Then GoTo Step_4
    
Step_2: 'Open Recordsets
    Set m_rstHeader = New ADODB.Recordset
    Set m_rstHeaderZegels = New ADODB.Recordset
    Set m_rstHeaderHandelaars = New ADODB.Recordset
    Set m_rstDetails = New ADODB.Recordset
    Set m_rstDetailsBijzondere = New ADODB.Recordset
    Set m_rstDetailsContainer = New ADODB.Recordset
    Set m_rstDetailsDocumenten = New ADODB.Recordset
    Set m_rstDetailsZelf = New ADODB.Recordset
    Set m_rstDetailsHandelaars = New ADODB.Recordset
    Set m_rstDetailsBerekeningsEenheden = New ADODB.Recordset
    
    Select Case DType
        Case 14
            strSQLH = "SELECT * FROM [PLDA Import Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Import Header Zegels] WHERE Code = '" & Code & "'"
            strSQLD = "SELECT * FROM [PLDA Import Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Import Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Import Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Import Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDZelf = "SELECT * FROM [PLDA Import Detail Zelf] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLHHandelaars = "SELECT * FROM [PLDA Import Header Handelaars] WHERE Code = '" & Code & "'"
            strSQLDHandelaars = "SELECT * FROM [PLDA Import Detail Handelaars] WHERE Code = '" & Code & "'"
            strSQLDBerekeningsEenheden = "SELECT * FROM [PLDA Import Detail Berekenings Eenheden] WHERE Code = '" & Code & "'"
            
            ADORecordsetOpen strSQLDBerekeningsEenheden, g_conSADBEL, m_rstDetailsBerekeningsEenheden, adOpenKeyset, adLockOptimistic
            ADORecordsetOpen strSQLDZelf, g_conSADBEL, m_rstDetailsZelf, adOpenKeyset, adLockOptimistic
            
            'm_rstDetailsBerekeningsEenheden.Open strSQLDBerekeningsEenheden, g_conSADBEL, adOpenKeyset, adLockOptimistic
            'm_rstDetailsZelf.Open strSQLDZelf, g_conSADBEL, adOpenKeyset, adLockOptimistic
            
        Case 18
            
            strSQLH = "SELECT * FROM [PLDA Combined Header] WHERE Code = '" & Code & "'"
            strSQLHZegels = "SELECT * FROM [PLDA Combined Header Zegels] WHERE Code = '" & Code & "'"
            strSQLD = "SELECT * FROM [PLDA Combined Detail] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDBijzondere = "SELECT * FROM [PLDA Combined Detail Bijzondere] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDContainer = "SELECT * FROM [PLDA Combined Detail Container] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLDDocumenten = "SELECT * FROM [PLDA Combined Detail Documenten] WHERE Code = '" & Code & "' ORDER BY Detail"
            strSQLHHandelaars = "SELECT * FROM [PLDA Combined Header Handelaars] WHERE Code = '" & Code & "'"
            strSQLDHandelaars = "SELECT * FROM [PLDA Combined Detail Handelaars] WHERE Code = '" & Code & "'"
                
    End Select
    
    
    ADORecordsetOpen strSQLH, g_conSADBEL, m_rstHeader, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHZegels, g_conSADBEL, m_rstHeaderZegels, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLHHandelaars, g_conSADBEL, m_rstHeaderHandelaars, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLD, g_conSADBEL, m_rstDetails, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDBijzondere, g_conSADBEL, m_rstDetailsBijzondere, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDContainer, g_conSADBEL, m_rstDetailsContainer, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDDocumenten, g_conSADBEL, m_rstDetailsDocumenten, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen strSQLDHandelaars, g_conSADBEL, m_rstDetailsHandelaars, adOpenKeyset, adLockOptimistic
    
    
Step_3: 'Split the release message and save each segment
    astrPerMessage = Split(ReleaseMessage, CONST_MESSAGE_DELIMETER)   'astrPerMessage(0) and astrPerMessage(1) are the CUSDEC
    strCUSDEC = astrPerMessage(0) & CONST_MESSAGE_DELIMETER & astrPerMessage(1)
    
    astrPerSection = Split(strCUSDEC, CONST_SECTION_DELIMETER)
    
    'astrPerSection(0) = header
    'astrPerSection(1) = details
    'astrPerSection(2) = summary
    
    For lngKownterPerSection = LBound(astrPerSection) To UBound(astrPerSection)
        astrPerSegment = Split(astrPerSection(lngKownterPerSection), CONST_SEGMENT_DELIMETER)
        
        For lngKownterPerSegment = LBound(astrPerSegment) To UBound(astrPerSegment)
            SaveSegmentToDB astrPerSegment(lngKownterPerSegment), lngKownterPerSection
        Next lngKownterPerSegment
    
    Next lngKownterPerSection
    
Step_4:  'Clean up then exit

    Erase astrPerMessage
    Erase astrPerSection
    Erase astrPerSegment
End Sub

Private Sub SaveSegmentToDB(ByVal Segment As String, ByVal Section As Message_Section)

    Dim astrPerDataItem() As String             'Used to store the data items per segment
    
    Dim strSegmentName As String
    
    strSegmentName = Left(Segment, 3)   'Get the segment name
    
Step_1: 'Validations
    If Len(Segment) = 0 Then GoTo Step_Exit
    If Not HasFieldToSave(strSegmentName, Section) Then GoTo Step_Exit

Step_2: 'Parse the segment into data items
    astrPerDataItem = Split(Segment, CONST_DATA_ITEM_DELIMETER)     'astrPerDataItem(0) is the segment name

    Call SaveDataItemToBoxes(astrPerDataItem, Section)

Step_Exit:

End Sub

Private Sub SaveDataItemToBoxes(ByRef Segment() As String, ByVal Section As Message_Section)
    
    Dim IsDetail As Boolean
    Dim intDIOrdinal As Integer
    Dim intCDIOrdinal As Integer
    Dim strQualifier As String
    Dim strSegmentName As String
    Dim strBox As String
    Dim strGroupQualifier As String
    Dim astrPerCompositeDataItem() As String    'Used to store the composite data items per data item group
        
    strSegmentName = Segment(0)
    
    Select Case Section
        Case enuHeader
            IsDetail = False
            Select Case strSegmentName
                Case "RFF", "TDT", "DOC", "NAD", "TOD", "MOA"
                    If InStr(1, Segment(1), ":") = 0 Then
                        m_strLastGroupRootSegment = strSegmentName & "+" & Segment(1)
                    Else
                        m_strLastGroupRootSegment = strSegmentName & "+" & Left(Segment(1), InStr(1, Segment(1), ":") - 1)
                    End If
            End Select
            
        Case enuDetail
            IsDetail = True
            Select Case strSegmentName
                Case "PAC", "RFF", "DOC", "TOD", "MOA", "GEI", "TAX"
                    If InStr(1, Segment(1), ":") = 0 Then
                        m_strLastGroupRootSegment = strSegmentName & "+" & Segment(1)
                    Else
                        m_strLastGroupRootSegment = strSegmentName & "+" & Left(Segment(1), InStr(1, Segment(1), ":") - 1)
                    End If
            End Select
            
        Case enuSummary: IsDetail = False
    End Select
    
    For intDIOrdinal = 1 To UBound(Segment)
        
        astrPerCompositeDataItem = Split(Segment(intDIOrdinal), CONST_COMPOSITE_DATA_ITEM_DELIMETER)
                
        If UBound(astrPerCompositeDataItem) <> -1 Then 'Edwin Sept 17
            If intDIOrdinal = 1 Then strQualifier = astrPerCompositeDataItem(0)
        End If
        
        If strSegmentName = "CST" Then m_lngDetailNo = CLng(strQualifier)
        
        For intCDIOrdinal = 1 To UBound(astrPerCompositeDataItem) + 1
            
            strGroupQualifier = GetGroupQualifier(strSegmentName, strQualifier)
            strBox = GetBoxCode(intDIOrdinal, intCDIOrdinal, strSegmentName, strQualifier, DType, strGroupQualifier, IsDetail)

            If Not IsDetail Then
                If strBox = "E1" Or strBox = "E2" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstHeaderZegels, IsDetail)
                
                ElseIf strBox = "B1" Then
                    ' Reims For Now
                    Debug.Assert False ' Must check if a better way can be found
                    Call SaveToBox(strBox, Segment(2), m_rstHeader, IsDetail)

                ElseIf InStr(1, strBox, "X1", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X2", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X3", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X4", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X5", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X6", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X7", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X8", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "X9", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "XA", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "XB", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "XC", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "XD", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "XF", vbTextCompare) <> 0 Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstHeaderHandelaars, IsDetail)
                Else
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstHeader, IsDetail)
                End If
            ElseIf IsDetail Then
                If strBox = "P1" Or strBox = "P2" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsBijzondere, IsDetail)
                ElseIf strBox = "S4,S5" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsContainer, IsDetail)
                ElseIf strBox = "Q1" Or strBox = "Q2" Or strBox = "Q3" Or strBox = "Q5" Or _
                       strBox = "Q9" Or strBox = "QB" Or strBox = "QC" Then 'Or strBox = "R1" Or _
                       strBox = "R2" Or strBox = "R3" Or strBox = "R5" Or strBox = "R6" Or _
                       strBox = "R8" Or strBox = "R9" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsDocumenten, IsDetail)
                ElseIf strBox = "U1" Or strBox = "U2" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsZelf, IsDetail)
                ElseIf InStr(1, strBox, "V1", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V2", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V3", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V4", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V5", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V6", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V7", vbTextCompare) <> 0 Or _
                       InStr(1, strBox, "V8", vbTextCompare) <> 0 Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsHandelaars, IsDetail)
                ElseIf strBox = "TZ" Or strBox = "T8" Then
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetailsBerekeningsEenheden, IsDetail)
                Else
                    Call SaveToBox(strBox, astrPerCompositeDataItem(intCDIOrdinal - 1), m_rstDetails, IsDetail)
                End If
            End If
        
        Next intCDIOrdinal
        
        Erase astrPerCompositeDataItem
        
    Next intDIOrdinal

End Sub

Private Sub SaveToBox(ByVal BoxCode As String, ByVal BoxValue As String, _
                      ByRef Recordset As ADODB.Recordset, ByVal IsDetail As Boolean)
    
    Dim astrTemp() As String
    Dim lngCtr As Long
    Dim lngInitial As Long
    Dim lngStep As Long
    
    On Error Resume Next
    
    If Len(BoxValue) = 0 Then Exit Sub
    If Len(BoxCode) = 0 Then Exit Sub
    
    If InStr(1, BoxCode, "*") <> 0 Then
    
        If BoxCode = "A1*A2" Then
            Recordset.Fields("A1").Value = Left(BoxValue, 2)
            Recordset.Fields("A2").Value = Right(BoxValue, 1)
            Recordset.Update
             
            ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
            
        ElseIf BoxCode = "M3*M4*M5" Then
            Recordset.Filter = adFilterNone
            Recordset.Filter = "Detail = " & m_lngDetailNo
            
            If IsNull(BoxValue) = True Then
                If LenB(Trim$(BoxValue)) > 0 Then
                    
                    astrTemp = Split(BoxValue, " ")
                    
                    If UBound(astrTemp) <> 2 Then
                        astrTemp = Split(BoxValue, "-")
                    End If
                      
                    lngInitial = 2
                    lngStep = 1
                    
                    For lngCtr = LBound(astrTemp) To UBound(astrTemp) Step lngStep
                        Select Case (lngInitial + lngStep)
                            Case 3 To 5
                                'lngInitial = lngInitial + lngStep - Edwin Sept 24
                                Recordset.Fields("M" & lngInitial).Value = astrTemp(lngCtr)
                                Recordset.Update
                                
                                ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
                                
                                lngInitial = lngInitial + lngStep
                        End Select
                    Next lngCtr
                End If
            End If
            
            Recordset.Filter = adFilterNone
        
        ElseIf BoxCode = "XF(2)*XD(2)" Then
            Recordset.Filter = adFilterNone
            Recordset.Filter = "XE = '2'"
            
            astrTemp = Split(BoxValue, "-")
            
            Recordset.Fields("XF").Value = astrTemp(0)
            Recordset.Fields("XD").Value = astrTemp(1)
            Recordset.Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
            
            Recordset.Filter = adFilterNone
        
        ElseIf BoxCode = "R1*R2" Then
            Recordset.Filter = adFilterNone
            Recordset.Filter = "Detail = " & m_lngDetailNo
            
            astrTemp = Split(BoxValue, "-")
            
            Recordset.Fields("R1").Value = astrTemp(0)
            Recordset.Fields("R2").Value = astrTemp(1)
            Recordset.Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
            
            Recordset.Filter = adFilterNone
            
        ElseIf BoxCode = "L5*L6" Then
            Recordset.Filter = adFilterNone
            Recordset.Filter = "Detail = " & m_lngDetailNo
            
            astrTemp = Split(BoxValue, "-")
            
            Recordset.Fields("L5").Value = astrTemp(0)
            Recordset.Fields("L6").Value = astrTemp(1)
            Recordset.Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
            
            Recordset.Filter = adFilterNone
        End If
    
    ElseIf InStr(1, BoxCode, "(") <> 0 Then
        
        Dim strBox As String
        Dim strFilterElement As String
        
        strBox = Left(BoxCode, 2)
        strFilterElement = Replace(Right(BoxCode, 2), "(", "")
        strFilterElement = Replace(strFilterElement, ")", "")
        
        Recordset.Filter = adFilterNone
        If IsDetail Then
            Recordset.Filter = "Detail = " & m_lngDetailNo & " AND " & Left(strBox, 1) & "E = '" & strFilterElement & "' "
        Else
            Recordset.Filter = Left(strBox, 1) & "E = '" & strFilterElement & "' "
        End If
        
        Recordset.Fields(strBox).Value = BoxValue
        Recordset.Update
        
        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
        
        Recordset.Filter = adFilterNone
    
    ElseIf BoxCode = "S4,S5" Then
        'Recordset.Filter = "Detail = " & m_lngDetailNo
                
    Else
        
        If BoxCode = "U1" Or BoxCode = "U2" Then Exit Sub

        Recordset.Filter = adFilterNone

        If IsDetail Then
            Recordset.Filter = "Detail = " & m_lngDetailNo
        End If
        
        Recordset.Fields(BoxCode).Value = BoxValue
        Recordset.Update
        
        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSADBEL, Recordset, GetTableToUpdateFromActiveCommand(Recordset)
        
        Recordset.Filter = adFilterNone
    End If
    
    Erase astrTemp
    
End Sub

Private Function GetGroupQualifier(ByVal SegmentName As String, ByVal Qualifier As String) As String

    Select Case SegmentName
        Case "BGM", "GEI", "FII", "SEL", "FTX", "TDT", "DOC", "NAD", _
             "TOD", "MOA", "CST", "MEA", "PAC", "PCI", "TAX", "CNT"
            GetGroupQualifier = ""
            
        Case "LOC"
            Select Case Qualifier
                Case "41", "43", "35", "27", "28", "47", "18", "42", "106"
                    GetGroupQualifier = ""
                Case "91"
                    If m_strLastGroupRootSegment = "DOC+190" Then
                        GetGroupQualifier = "190"
                    Else
                        GetGroupQualifier = ""
                    End If
                Case "127"
                    If m_strLastGroupRootSegment = "DOC+998" Then
                        GetGroupQualifier = "998"
                    Else
                        GetGroupQualifier = ""
                    End If
                Case "1"
                    If m_strLastGroupRootSegment = "TOD+6" Then
                        GetGroupQualifier = "6"
                    Else
                        GetGroupQualifier = ""
                    End If
               End Select
                
        Case "DTM"
            Select Case Qualifier
                Case "254", "36"
                    GetGroupQualifier = ""
                Case "137"
                    If m_strLastGroupRootSegment = "DOC+190" Then
                        GetGroupQualifier = "190"
                    ElseIf m_strLastGroupRootSegment = "DOC+998" Then
                        GetGroupQualifier = "998"
                    End If
            End Select
        
        Case "RFF"
            Select Case Qualifier
                Case "ABE"
                    If m_strLastGroupRootSegment = "NAD+AE" Then GetGroupQualifier = "AE"
                Case Else
                    GetGroupQualifier = ""
            End Select

        Case "CUX"
            If m_strLastGroupRootSegment = "MOA+69" Then GetGroupQualifier = "69"
            If m_strLastGroupRootSegment = "MOA+72" Then GetGroupQualifier = "72"
            If m_strLastGroupRootSegment = "MOA+39" Then GetGroupQualifier = "39"
            If m_strLastGroupRootSegment = "MOA+38" Then GetGroupQualifier = "38"
            If m_strLastGroupRootSegment = "MOA+123" Then GetGroupQualifier = "123"
                            
    End Select

End Function

Private Function HasFieldToSave(ByVal SegmentName As String, ByVal Section As Message_Section) As Boolean

    Select Case Section
        Case enuHeader
        
            Select Case SegmentName
                Case "BGM", "LOC", "DTM", "GEI", "FII", "MEA", "SEL", "FTX", "CST", _
                     "RFF", "TDT", "DOC", "NAD", "CTA", "COM", "TOD", "MOA", "CUX"
                    
                    HasFieldToSave = True
                    
                Case "UNB", "UNH", "EQD", "UNS", "UNT", "UNZ"
                
                    HasFieldToSave = False
                    
                Case Else
                
                    HasFieldToSave = False
                    
            End Select
        
        Case enuDetail
        
            Select Case SegmentName
                Case "CST", "FTX", "LOC", "MEA", "NAD", "PAC", "PCI", "MOA", _
                     "CUX", "RFF", "DOC", "DTM", "TOD", "GEI", "TAX"
                    
                    HasFieldToSave = True
                                       
                Case Else
                
                    HasFieldToSave = False
                    
            End Select
        
        Case enuSummary
        
            Select Case SegmentName
                Case "CNT"
                    
                    HasFieldToSave = True
                                       
                Case Else
                
                    HasFieldToSave = False
                    
            End Select
        
    End Select
    
End Function

Private Sub Class_Initialize()
    DType = 0   'Initialize the DType
End Sub

Private Sub Class_Terminate()
    
    'Clean up
    
    ADORecordsetClose m_rstHeader
    ADORecordsetClose m_rstHeaderZegels
    ADORecordsetClose m_rstDetails
    ADORecordsetClose m_rstDetailsBijzondere
    ADORecordsetClose m_rstDetailsContainer
    ADORecordsetClose m_rstDetailsDocumenten
    ADORecordsetClose m_rstDetailsZelf
    ADORecordsetClose m_rstHeaderHandelaars
    ADORecordsetClose m_rstDetailsHandelaars
    
    
'    If Not m_rstHeader Is Nothing Then
'        If m_rstHeader.State = adStateOpen Then
'            m_rstHeader.Close
'        End If
'        Set m_rstHeader = Nothing
'    End If
'
'    If Not m_rstHeaderZegels Is Nothing Then
'        If m_rstHeaderZegels.State = adStateOpen Then
'            m_rstHeaderZegels.Close
'        End If
'        Set m_rstHeaderZegels = Nothing
'    End If
'
'    If Not m_rstDetails Is Nothing Then
'        If m_rstDetails.State = adStateOpen Then
'            m_rstDetails.Close
'        End If
'        Set m_rstDetails = Nothing
'    End If
'
'    If Not m_rstDetailsBijzondere Is Nothing Then
'        If m_rstDetailsBijzondere.State = adStateOpen Then
'            m_rstDetailsBijzondere.Close
'        End If
'        Set m_rstDetailsBijzondere = Nothing
'    End If
'
'    If Not m_rstDetailsContainer Is Nothing Then
'        If m_rstDetailsContainer.State = adStateOpen Then
'            m_rstDetailsContainer.Close
'        End If
'        Set m_rstDetailsContainer = Nothing
'    End If
'
'    If Not m_rstDetailsDocumenten Is Nothing Then
'        If m_rstDetailsDocumenten.State = adStateOpen Then
'            m_rstDetailsDocumenten.Close
'        End If
'        Set m_rstDetailsDocumenten = Nothing
'    End If
'
'    If Not m_rstDetailsZelf Is Nothing Then
'        If m_rstDetailsZelf.State = adStateOpen Then
'            m_rstDetailsZelf.Close
'        End If
'        Set m_rstDetailsZelf = Nothing
'    End If
'
'    If Not m_rstHeaderHandelaars Is Nothing Then
'        If m_rstHeaderHandelaars.State = adStateOpen Then
'            m_rstHeaderHandelaars.Close
'        End If
'        Set m_rstHeaderHandelaars = Nothing
'    End If
'
'
'    If Not m_rstDetailsHandelaars Is Nothing Then
'        If m_rstDetailsHandelaars.State = adStateOpen Then
'            m_rstDetailsHandelaars.Close
'        End If
'        Set m_rstDetailsHandelaars = Nothing
'    End If
    
    ADODisconnectDB g_conSADBEL
End Sub
