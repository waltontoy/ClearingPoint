VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdifactMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SegmentCounts"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5CF029E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Attribute VB_Ext_KEY = "Member1" ,"Segments"
Option Explicit


Private mvarDAOConnection As ADODB.Connection

Private mconADOConnectionStandard As ADODB.Connection
Private mconADOConnectionShaping As ADODB.Connection

Private rstNCTSDataMessage As ADODB.Recordset
Private rstNCTSDataTable As ADODB.Recordset
Private rstNCTSDataTables As ADODB.Recordset
Private mrstDataNCTSMessage As ADODB.Recordset
Private mrstTechnicalMessageStructure As ADODB.Recordset
Private mvarrstMessageTypes As ADODB.Recordset

Private mvarSegments As Segments

' Reims = Change to ENCTSMessageType
Private mvarMessageType As ENCTSMessageType

Private mvarSendLocation As String
Private mvarReceiveLocation As String
Private mvarDBName As String 'local copy
'Private mvarDBLocation As String 'local copy
Private mvarSegmentCounts As SegmentCounts
Private mvarEDIMessage As String
Private mvarNCTS_IEM_ID As Long 'local copy

Private mstrUniqueCode As String
Private mvarLRN As String 'local copy

Friend Property Let LRN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LRN = 5
    mvarLRN = vData
End Property

Public Property Get LRN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LRN
    LRN = mvarLRN
End Property

Public Property Get NCTSTMSRecordset() As ADODB.Recordset
    Set NCTSTMSRecordset = mrstTechnicalMessageStructure
End Property
    
Public Property Get DataNCTSMessageRecordset() As ADODB.Recordset
    Set DataNCTSMessageRecordset = mrstDataNCTSMessage
End Property
    
Public Property Get UniqueCode() As String
    UniqueCode = mstrUniqueCode
End Property

Public Property Let UniqueCode(ByVal strUniqueCode As String)
    mstrUniqueCode = strUniqueCode
End Property
    
Public Property Get NCTS_IEM_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCTS_IEM_ID
    NCTS_IEM_ID = mvarNCTS_IEM_ID
End Property

Public Property Get ReferenceNumber() As String
    Dim strReturnValue As String
    
    Dim lngSegmentIndex As Long
    
    Dim blnFinishedSearch As Boolean
    Dim blnFoundSegment As Boolean
    
    lngSegmentIndex = 0
    blnFinishedSearch = (Segments.Count = 0)
    Do While Not blnFinishedSearch
        lngSegmentIndex = lngSegmentIndex + 1
        blnFoundSegment = (Segments(lngSegmentIndex).SegmentTag = "RFF") And _
                          (Segments(lngSegmentIndex).SDataItems(1).Value = "ABE")
        blnFinishedSearch = (blnFoundSegment Or (lngSegmentIndex = Segments.Count))
    Loop
    If blnFoundSegment Then
        strReturnValue = Segments(lngSegmentIndex).SDataItems(2).Value
    Else
        strReturnValue = vbNullString
    End If
    
    mvarLRN = strReturnValue
    ReferenceNumber = strReturnValue
End Property

Public Property Get MRN() As String
    Dim strReturnValue As String
    
    Dim lngSegmentIndex As Long
    
    Dim blnFinishedSearch As Boolean
    Dim blnFoundSegment As Boolean
    
    lngSegmentIndex = 0
    blnFinishedSearch = (Segments.Count = 0)

    Do While Not blnFinishedSearch
        lngSegmentIndex = lngSegmentIndex + 1
        blnFoundSegment = Segments(lngSegmentIndex).SegmentTag = "BGM"

        blnFinishedSearch = (blnFoundSegment Or (lngSegmentIndex = Segments.Count))
    Loop
    
    If blnFoundSegment Then
        strReturnValue = Segments(lngSegmentIndex).SDataItems(5).Value
    Else
        strReturnValue = vbNullString
    End If
        
    MRN = strReturnValue
End Property

Public Function GetSegmentCountsIndex(ByVal SearchSegmentCountsKey As String) As Long
    Dim lngReturnValue As Long
    
    Dim blnFoundKey As Boolean
    Dim blnContinueLoop As Boolean
    Dim lngSegmentCountsIndex As Long
    
    lngReturnValue = 0
    lngSegmentCountsIndex = 0
    blnFoundKey = False
    blnContinueLoop = (SegmentCounts.Count > 0)
    Do While blnContinueLoop
        lngSegmentCountsIndex = lngSegmentCountsIndex + 1
        blnFoundKey = (SegmentCounts(lngSegmentCountsIndex).Key = SearchSegmentCountsKey)
        blnContinueLoop = Not (blnFoundKey Or (SegmentCounts.Count = lngSegmentCountsIndex))
    Loop
    If blnFoundKey Then
        lngReturnValue = lngSegmentCountsIndex
    End If
    GetSegmentCountsIndex = lngReturnValue
End Function

Public Property Get SegmentCounts() As SegmentCounts
'    If mvarSegmentCounts Is Nothing Then
'        Set mvarSegmentCounts = New SegmentCounts
'    End If
    Set SegmentCounts = mvarSegmentCounts
End Property

Public Property Set SegmentCounts(vData As SegmentCounts)
    Set mvarSegmentCounts = vData
End Property

'Public Property Let DBLocation(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.DBLocation = 5
'    mvarDBLocation = vData
'End Property

Public Property Get EDIMessage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DBLocation
    EDIMessage = mvarEDIMessage
End Property

'Public Property Get DBLocation() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.DBLocation
'    DBLocation = mvarDBLocation
'End Property
'
'Public Property Let DBName(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.DBName = 5
'    mvarDBName = vData
'End Property
'
'Public Property Get DBName() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.DBName
'    DBName = mvarDBName
'End Property
    
Public Property Get ReceiveLocation() As String
    ReceiveLocation = mvarReceiveLocation
End Property

Public Property Let ReceiveLocation(ByVal vNewValue As String)
    mvarReceiveLocation = vNewValue
End Property

Public Property Get SendLocation() As String
    SendLocation = mvarSendLocation
End Property

Public Property Let SendLocation(ByVal vNewValue As String)
    mvarSendLocation = vNewValue
End Property

Private Function ReduceSegment(ByVal SegmentData As String) As String
    Dim strReturnValue As String
    
    Dim strReverseSegmentData As String
    Dim strReverseCompositeData As String
    
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
    
    lngStartIndex = 1
    '----->  Invert the string because the process will start from the end of the string
    strReverseSegmentData = StrReverse(SegmentData)
    
    lngStartIndex = 1
    
    Do
        '----->  GET THE DESIRED COMPOSITE DATA STRING
        '----->  Get the position of the next data element separator
        lngEndIndex = InStr(lngStartIndex, strReverseSegmentData, EDI_SEP_DATA_ELEMENT)
        If lngEndIndex = 0 Then
            lngEndIndex = Len(strReverseSegmentData)
        End If
        '----->    and get the desired substring starting from lngStartIndex to lngEndIndex
        strReverseCompositeData = Mid(strReverseSegmentData, lngStartIndex, lngEndIndex - lngStartIndex + 1)
        
        '----->  REDUCE THE COMPOSITE DATA STRING IF POSSIBLE
        
        '----->  Remove all empty simple data elements from the end
        Do While Left(strReverseCompositeData, 1) = EDI_SEP_COMPOSITE_DATA_ELEMENT
            strReverseCompositeData = Right(strReverseCompositeData, Len(strReverseCompositeData) - 1)
        Loop
        strReturnValue = strReturnValue & strReverseCompositeData
        lngStartIndex = lngEndIndex + 1
    Loop Until lngEndIndex = Len(strReverseSegmentData)
    
    '----->  Remove all empty composite data elements from the end
    Do While Left(strReturnValue, 1) = EDI_SEP_DATA_ELEMENT
        strReturnValue = Right(strReturnValue, Len(strReturnValue) - 1)
    Loop
    
    ReduceSegment = StrReverse(strReturnValue)
End Function

Private Sub AddMissingRequiredSegments()
    Dim strSegmentKeySubString As String
    Dim lngSegmentIndex As Long
    
    Dim blnFoundRequiredSegment As Boolean
    Dim blnFinishedSearched As Boolean
    
    mrstTechnicalMessageStructure.MoveFirst
    Do While Not mrstTechnicalMessageStructure.EOF
        If mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_Usage").Value = EDI_USAGE_REQUIRED Then
            lngSegmentIndex = 0
            blnFoundRequiredSegment = False
            blnFinishedSearched = False
            Do While Not blnFinishedSearched
                lngSegmentIndex = lngSegmentIndex + 1
                strSegmentKeySubString = "S_" & mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value & "_"
                blnFoundRequiredSegment = (InStr(1, Segments(lngSegmentIndex).Key, strSegmentKeySubString) > 0)
                blnFinishedSearched = (lngSegmentIndex = Segments.Count) Or blnFoundRequiredSegment
            Loop
            If Not blnFoundRequiredSegment Then
                Call AddEmptySegment(mrstTechnicalMessageStructure, mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value, 1, 1)
            End If
        End If
        mrstTechnicalMessageStructure.MoveNext
    Loop
    
    ' Reims - hardcoded key for now
    mrstTechnicalMessageStructure.MoveFirst
    'Segments.Item("S2_1_P_0_1").SDataItems("S6").Value = mrstTechnicalMessageStructure.Fields("NCTS_IEM_Code").Value
End Sub

Private Function GetBoxCodeGroup(ByVal BoxCode As String) As String
    Dim strReturnValue As String
    
    Select Case UCase(BoxCode)
        Case "W7", "T7", "MEMO FIELD"
            strReturnValue = BOX_GROUP_DETAIL
        Case "Z4"
            strReturnValue = BOX_GROUP_DETAIL_BIJZONDERE
        Case "S5"
            strReturnValue = BOX_GROUP_DETAIL_COLLI
        Case "SB"
            strReturnValue = BOX_GROUP_DETAIL_CONTAINER
        Case "Y1", "Y5"
            strReturnValue = BOX_GROUP_DETAIL_DOCUMENTEN
        Case "AE", "AF", "BA", "DOC NUMBER"
            strReturnValue = BOX_GROUP_HEADER
        Case "EJ", "EO"
            strReturnValue = BOX_GROUP_HEADER_ZEKERHEID
        Case "AC", "B9", "S2", "S4", "Z1", "Z2"
            strReturnValue = vbNullString
        Case "IT" ' Means the value of the current box code is to be checked
            strReturnValue = vbNullString
        Case Else
            If Left(BoxCode, 4) = "MID(" Then
                strReturnValue = vbNullString
            ElseIf IsNumeric(BoxCode) Then
                strReturnValue = vbNullString
            Else
                strReturnValue = vbNullString
                'Debug.Assert False
            End If
    End Select
    
    GetBoxCodeGroup = strReturnValue
End Function

Private Function GetConditionTokens(ByVal SourceString As String) As Variant
    Dim varReturnValue As Variant
    Dim varConditions  As Variant
    
    Dim lngConditionIndex As Long
    Dim lngConditionsCount As Long
    Dim lngLengthOperator As Long
    
    Dim strCondition As String
    Dim strOperator As String
    Dim strRightOperand As String
    Dim strLogicOperator As String
    
    If InStr(1, SourceString, OPERATOR_AND) > 0 Then
        '----->  AND
        varConditions = Split(SourceString, OPERATOR_AND)
        strLogicOperator = OPERATOR_AND
    ElseIf InStr(1, SourceString, OPERATOR_OR) > 0 Then
        '----->  OR
        varConditions = Split(SourceString, OPERATOR_OR)
        strLogicOperator = OPERATOR_OR
    Else
        '----->  ONE CONDITION ONLY
        varConditions = Array(SourceString)
        strLogicOperator = OPERATOR_OR
    End If
    
    lngConditionsCount = UBound(varConditions)
    ReDim varReturnValue(lngConditionsCount, 2)
    
    For lngConditionIndex = LBound(varConditions) To lngConditionsCount
        strCondition = Trim(varConditions(lngConditionIndex))
        Select Case Mid(strCondition, 2, 1)
            Case OPERATOR_LESS_THAN
                Debug.Assert Mid(strCondition, 1, 1) = OPERATOR_EQUAL
                lngLengthOperator = 2
            Case OPERATOR_GREATER_THAN
                Debug.Assert Mid(strCondition, 1, 1) = OPERATOR_EQUAL Or Mid(strCondition, 1, 1) = OPERATOR_LESS_THAN
                lngLengthOperator = 2
            
            ' MUCP-65 - Start
            Case OPERATOR_LEFT
                Debug.Assert False
            ' MUCP-65 - End
            
            Case Else
                lngLengthOperator = 1
        End Select
        strOperator = Left(strCondition, lngLengthOperator)
        strRightOperand = Trim(Mid(strCondition, lngLengthOperator + 1))
        varReturnValue(lngConditionIndex, 0) = strOperator
        varReturnValue(lngConditionIndex, 1) = strRightOperand
        varReturnValue(lngConditionIndex, 2) = strLogicOperator
    Next
    '----->  Set operator of first element to OR always
    varReturnValue(0, 2) = OPERATOR_OR
    GetConditionTokens = varReturnValue
End Function

Private Function IsLogicalOperator(ByVal Character As String) As Boolean
    Dim blnReturnValue As Boolean
    
    ' MUCP-65 - Start
    Debug.Assert Character <> OPERATOR_LEFT
    blnReturnValue = ((Character = OPERATOR_EQUAL) Or _
                      (Character = OPERATOR_GREATER_THAN) Or _
                      (Character = OPERATOR_LESS_THAN) Or _
                      (Character = OPERATOR_LEFT))
    
    'blnReturnValue = ((Character = OPERATOR_EQUAL) Or _
                      (Character = OPERATOR_GREATER_THAN) Or _
                      (Character = OPERATOR_LESS_THAN))
    ' MUCP-65 - End
    
    IsLogicalOperator = blnReturnValue
End Function

Private Function ParseConditionTokens(ByVal SourceString As String) As Variant
    Dim varReturnValue          As Variant
    
    Dim lngCharacterIndex       As Long
    Dim lngSourceStringLength   As Long
    Dim strCharacter            As String
    Dim blnContinueLoop         As Boolean
    Dim strLeftToken            As String
    Dim strOperator             As String
    Dim strRightOperator        As String
    
    Debug.Assert SourceString <> "%'240220'"
    lngSourceStringLength = Len(SourceString)
    lngCharacterIndex = 0
    '----->  LEFT OPERAND
    strLeftToken = vbNullString
    blnContinueLoop = True
    Do
        lngCharacterIndex = lngCharacterIndex + 1
        strCharacter = Mid(SourceString, lngCharacterIndex, 1)
        If IsLogicalOperator(strCharacter) Or strCharacter = vbNullString Then
            blnContinueLoop = False
        Else
            strLeftToken = strLeftToken & strCharacter
        End If
    Loop Until Not blnContinueLoop
    '----->  OPERAND
    lngCharacterIndex = lngCharacterIndex - 1
    strOperator = vbNullString
    blnContinueLoop = True
    Do
        lngCharacterIndex = lngCharacterIndex + 1
        strCharacter = Mid(SourceString, lngCharacterIndex, 1)
        If IsLogicalOperator(strCharacter) Then
            strOperator = strOperator & strCharacter
        Else
            blnContinueLoop = False
        End If
    Loop Until Not blnContinueLoop
    '----->  RIGHT OPERAND
    lngCharacterIndex = lngCharacterIndex - 1
    strRightOperator = vbNullString
    blnContinueLoop = True
    Do
        lngCharacterIndex = lngCharacterIndex + 1
        strCharacter = Mid(SourceString, lngCharacterIndex, 1)
        If IsLogicalOperator(strCharacter) Or strCharacter = vbNullString Then
            blnContinueLoop = False
        Else
            strRightOperator = strRightOperator & strCharacter
        End If
    Loop Until Not blnContinueLoop
    varReturnValue = Array(strLeftToken, strOperator, strRightOperator)
    ParseConditionTokens = varReturnValue
End Function

Private Function ParseCondition(ByVal SourceString As String) As Variant
    Dim varReturnValue          As Variant
    Dim varConditions           As Variant
    Dim varConditionTokens      As Variant
    
    Dim lngConditionIndex       As Long
    Dim lngConditionsCount      As Long
'    Dim lngLengthOperator       As Long    'IAN unused
    Dim lngAndLoc As Long   'IAN
    Dim lngOrLoc As Long    'IAN
    
'    Dim strCondition            As String  'IAN unused
'    Dim strOperator             As String  'IAN unused
'    Dim strRightOperand         As String  'IAN unused
'    Dim strLogicOperator        As String  'IAN
    Dim strArgument As String   'IAN 06-07-2005 temporary storage of Map condition
    Dim astrLogicalOperator() As String    'IAN 06-07-2005 Handle of logical operators for map conditions
        
    'IAN 06-07-2005
'    If InStr(1, SourceString, OPERATOR_AND) > 0 Then
'        '----->  AND
'        varConditions = Split(SourceString, OPERATOR_AND)
'        strLogicOperator = OPERATOR_AND
'    ElseIf InStr(1, SourceString, OPERATOR_OR) > 0 Then
'        '----->  OR
'        varConditions = Split(SourceString, OPERATOR_OR)
'        strLogicOperator = OPERATOR_OR
'    Else
'        '----->  ONE CONDITION ONLY
'        varConditions = Array(SourceString)
'        strLogicOperator = OPERATOR_OR
'    End If
    
    'IAN 06-07-2005 Above-commented group of code was replaced by this to handle combination of OR and AND in map conditions
    If InStr(1, SourceString, OPERATOR_AND) > 0 Or InStr(1, SourceString, OPERATOR_OR) > 0 Then
    
        strArgument = SourceString
        
        ReDim varConditions(0)
        ReDim astrLogicalOperator(0)
        
        astrLogicalOperator(0) = OPERATOR_OR
        
        Do While InStr(strArgument, OPERATOR_OR) > 0 Or InStr(strArgument, OPERATOR_AND) > 0
            
            lngOrLoc = InStr(strArgument, OPERATOR_OR)
            lngAndLoc = InStr(strArgument, OPERATOR_AND)
                
            ReDim Preserve astrLogicalOperator(UBound(astrLogicalOperator) + 1)
            If (UBound(varConditions) = 0 And varConditions(0) = "") = False Then
                ReDim Preserve varConditions(UBound(varConditions) + 1)
            End If
            
            If lngOrLoc = 0 Then
                varConditions(UBound(varConditions)) = Left(strArgument, lngAndLoc - 1)
                strArgument = Mid(strArgument, lngAndLoc + 1)
                astrLogicalOperator(UBound(astrLogicalOperator)) = OPERATOR_AND
            ElseIf lngAndLoc = 0 Then
                varConditions(UBound(varConditions)) = Left(strArgument, lngOrLoc - 1)
                strArgument = Mid(strArgument, lngOrLoc + 1)
                astrLogicalOperator(UBound(astrLogicalOperator)) = OPERATOR_OR
            Else
                If lngOrLoc < lngAndLoc Then
                    varConditions(UBound(varConditions)) = Left(strArgument, lngOrLoc - 1)
                    strArgument = Mid(strArgument, lngOrLoc + 1)
                    astrLogicalOperator(UBound(astrLogicalOperator)) = OPERATOR_OR
                Else
                    varConditions(UBound(varConditions)) = Left(strArgument, lngAndLoc - 1)
                    strArgument = Mid(strArgument, lngAndLoc + 1)
                    astrLogicalOperator(UBound(astrLogicalOperator)) = OPERATOR_AND
                End If
            End If
            
        Loop
        
        ReDim Preserve varConditions(UBound(varConditions) + 1)
        varConditions(UBound(varConditions)) = strArgument
        
    Else
        '----->  ONE CONDITION ONLY
        varConditions = Array(SourceString)
        
        'IAN 06-07-2005 replaced by astrLogicalOperator
'        strLogicOperator = OPERATOR_OR
        ReDim astrLogicalOperator(0)
        astrLogicalOperator(0) = OPERATOR_OR
        
    End If
    
    lngConditionsCount = UBound(varConditions)
    ReDim varReturnValue(lngConditionsCount, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator)
    
    For lngConditionIndex = LBound(varConditions) To lngConditionsCount
        
        varConditionTokens = ParseConditionTokens(Trim(varConditions(lngConditionIndex)))
        varReturnValue(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand) = varConditionTokens(ConditionTokenIndexes.ConditionTokenIndex_LeftOperand)
        varReturnValue(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_Operator) = varConditionTokens(ConditionTokenIndexes.ConditionTokenIndex_Operator)
        varReturnValue(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_RightOperand) = varConditionTokens(ConditionTokenIndexes.ConditionTokenIndex_RightOperand)
        
        'IAN 06-07-2005 replaced so that multiple operators will be supported
'        varReturnValue(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator) = strLogicOperator
        varReturnValue(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator) = astrLogicalOperator(lngConditionIndex)
        
    Next
    '----->  Set operator of first element to OR always
    varReturnValue(0, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator) = OPERATOR_OR
    ParseCondition = varReturnValue
End Function

Private Sub SetInstances(ByRef SegmentInstance As Long, _
                         ByRef SegmentInstanceParent As Long, _
                         ByVal BoxCode As String, _
                         ByVal HeaderNumber As Long, _
                         ByVal DetailNumber As Long, _
                         ByVal SegmentOccurrence As Long)
    Dim lngSegmentInstance As Long
    Dim lngSegmentInstanceParent As Long

    Select Case BoxCode
        Case "IT"
        Case "B9"
            SegmentInstance = HeaderNumber
            SegmentInstanceParent = SegmentOccurrence
        Case "EJ"
        Case "S2"
            lngSegmentInstance = 1
            lngSegmentInstanceParent = SegmentOccurrence
        Case "W7"
        Case "Y1"
        Case Else
            '----->  Not yet handled
            Debug.Assert IsNumeric(BoxCode)
    End Select

    SegmentInstance = lngSegmentInstance
    SegmentInstanceParent = lngSegmentInstanceParent
End Sub

Private Function GetLeftOperand(ByVal MapConditionsRecordset As ADODB.Recordset, _
                                ByVal UnmappedValuesRecordset As ADODB.Recordset, _
                                ByVal SourceBoxCode As String, _
                                ByVal BoxCodeGroup As String, _
                                ByVal SegmentKey As String, _
                                ByVal HeaderNumber As Long, _
                                ByVal DetailNumber As Long, _
                                ByRef OperandsAreNumeric As Boolean) As String
    Dim strReturnValue                  As String
    
    Dim lngSegmentInstance              As Long
    'Dim lngParentSegmentInstance        As Long
    Dim lngDetailNumber                 As Long
    Dim lngOrdinal                      As Long
    Dim lngStart                        As Long
    Dim lngLength                       As Long
    
    Dim lngDocumentenGroupIndex         As Long
    
    Dim strSegmentKey                   As String
    
    Dim varMidParameters                As Variant
    
    Dim rstMapConditionsRecordsetClone  As ADODB.Recordset
    Dim rstGroups                       As ADODB.Recordset
    
    If BoxCodeGroup = vbNullString Then
        '----->  Take note that MapConditionsRecordset is filtered but rstMapConditionsRecordsetClone is NOT
        Set rstMapConditionsRecordsetClone = MapConditionsRecordset.Clone
        
        'If IsInRecordset(rstMapConditionsRecordsetClone, "NCTS_IEM_MAP_Source = '" & MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value & "'") Then
        If IsInRecordset(rstMapConditionsRecordsetClone, "NCTS_IEM_MAP_Source = '" & SourceBoxCode & "'") Then
            'Select Case MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value
            Select Case SourceBoxCode
                Case "B9", "AC", "A6"
                    lngSegmentInstance = HeaderNumber
                    'lngParentSegmentInstance = 1
                    strSegmentKey = "S_" & rstMapConditionsRecordsetClone.Fields("NCTS_IEM_TMS_ID").Value & "_" & CStr(lngSegmentInstance)
'                Case "S2"
'                    strSegmentKey = SegmentKey
                Case "S2", "S4"
                    lngSegmentInstance = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                    strSegmentKey = "S_" & rstMapConditionsRecordsetClone.Fields("NCTS_IEM_TMS_ID").Value & "_" & CStr(lngSegmentInstance)
                Case "Z1", "Z2" 'IAN 06-07-2005 Added Z2 so it will be included in the EDI message
                    lngSegmentInstance = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                    'lngParentSegmentInstance = 1
                    strSegmentKey = "S_" & rstMapConditionsRecordsetClone.Fields("NCTS_IEM_TMS_ID").Value & "_" & CStr(lngSegmentInstance)
                ' MUCP-65 - Start
                Case "L1"
                    lngSegmentInstance = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                    strSegmentKey = "S_" & rstMapConditionsRecordsetClone.Fields("NCTS_IEM_TMS_ID").Value & "_" & CStr(lngSegmentInstance)
                ' MUCP-65 - End
                
                Case Else
                    '-----> ADD CODE TO HANDLE THIS BOX CODE
                    Debug.Assert False
                    strSegmentKey = vbNullString
            End Select
            
            If GetSegmentIndex(strSegmentKey) > 0 Then
                strReturnValue = Segments(strSegmentKey).SDataItems(rstMapConditionsRecordsetClone.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value).Value
                OperandsAreNumeric = (UCase(Left(Segments(strSegmentKey).SDataItems(rstMapConditionsRecordsetClone.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value).NCTSDataFormat, 1)) = "N")
            Else
                Debug.Assert False
            End If
        'ElseIf Left(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value, 4) = "MID(" And _
               Right(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value, 1) = ")" Then
        ElseIf Left(SourceBoxCode, 4) = "MID(" And _
               Right(SourceBoxCode, 1) = ")" Then
                'varMidParameters = Split(Mid(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value, 5, Len(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value) - 5), ",")
                varMidParameters = Split(Mid(SourceBoxCode, 5, Len(SourceBoxCode) - 5), ",")
                Debug.Assert IsArray(varMidParameters)
                Debug.Assert LBound(varMidParameters) = 0
                Debug.Assert UBound(varMidParameters) = 2
                If IsArray(varMidParameters) Then
                    If LBound(varMidParameters) = 0 And UBound(varMidParameters) = 2 Then
                        lngOrdinal = CLng(varMidParameters(0))
                        lngStart = CLng(varMidParameters(1))
                        lngLength = CLng(varMidParameters(2))
                    End If
                End If
                If GetSegmentIndex(SegmentKey) > 0 Then
                    strReturnValue = Segments(SegmentKey).SDataItems(lngOrdinal).Value
                    OperandsAreNumeric = (UCase(Left(Segments(SegmentKey).SDataItems(lngOrdinal).NCTSDataFormat, 1)) = "N")
                Else
                    Debug.Assert False
                End If
        'ElseIf MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value = "IT" Then
        ElseIf SourceBoxCode = "IT" Then
            If GetSegmentIndex(SegmentKey) > 0 Then
                strReturnValue = Segments(SegmentKey).SDataItems(MapConditionsRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value).Value
                OperandsAreNumeric = (UCase(Left(Segments(SegmentKey).SDataItems(MapConditionsRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value).NCTSDataFormat, 1)) = "N")
            Else
                Debug.Assert False
            End If
        'ElseIf IsNumeric(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value) Then
        ElseIf IsNumeric(SourceBoxCode) Then
            If GetSegmentIndex(SegmentKey) > 0 Then
                'strReturnValue = Segments(SegmentKey).SDataItems(CLng(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value)).Value
                strReturnValue = Segments(SegmentKey).SDataItems(CLng(SourceBoxCode)).Value
                'OperandsAreNumeric = (UCase(Left(Segments(SegmentKey).SDataItems(CLng(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value)).NCTSDataFormat, 1)) = "N")
                OperandsAreNumeric = (UCase(Left(Segments(SegmentKey).SDataItems(CLng(SourceBoxCode)).NCTSDataFormat, 1)) = "N")
            Else
                Debug.Assert False
            End If
        Else
            Debug.Assert False
        End If
    Else
        Set rstGroups = UnmappedValuesRecordset.Fields(BoxCodeGroup).Value
        Debug.Assert rstGroups.RecordCount > 0
        If rstGroups.RecordCount > 0 Then
            Select Case BoxCodeGroup
                Case BOX_GROUP_DETAIL
                    'Select Case MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value
                    Select Case SourceBoxCode
                        Case "W7"
                            lngDetailNumber = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                            Debug.Assert lngDetailNumber = DetailNumber
                            rstGroups.Filter = "HEADER = " & CStr(HeaderNumber) & " AND DETAIL = " & CStr(lngDetailNumber)
                        Case Else
                            Debug.Assert False
                    End Select
                Case BOX_GROUP_DETAIL_BIJZONDERE, BOX_GROUP_DETAIL_COLLI, BOX_GROUP_DETAIL_CONTAINER
                    Debug.Assert False
                Case BOX_GROUP_DETAIL_DOCUMENTEN
                    'Select Case MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value
                    Select Case SourceBoxCode
                        Case "Y1"
                            'lngParentSegmentInstance = CLng(Split(Segments(SegmentKey).KeyParent, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                            lngSegmentInstance = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                            rstGroups.Filter = adFilterNone
                            rstGroups.MoveFirst
                            rstGroups.Sort = "HEADER ASC, DETAIL ASC, ORDINAL ASC"
                            '----->  loop up to instance of segment minus 1
                            '----->  This is because for every detail in document, ordinal of Y1 is reset to 1
                            '----->  but it does not work this way for the other box codes
                            Debug.Assert lngSegmentInstance <= rstGroups.RecordCount
                            For lngDocumentenGroupIndex = 1 To lngSegmentInstance - 1
                                Debug.Assert Not rstGroups.EOF
                                If Not rstGroups.EOF Then
                                    rstGroups.MoveNext
                                End If
                            Next
                            Debug.Assert Not rstGroups.EOF
                            'rstGroups.Filter = "HEADER = " & CStr(HeaderNumber) & " AND DETAIL = " & CStr(lngParentSegmentInstance) & " AND ORDINAL = " & CStr(lngSegmentInstance)
                        Case Else
                            Debug.Assert False
                    End Select
                Case BOX_GROUP_HEADER
                    Debug.Assert False
                    rstGroups.Filter = "HEADER = " & CStr(HeaderNumber)
                Case BOX_GROUP_HEADER_ZEKERHEID
                    'Select Case MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value
                    Select Case SourceBoxCode
                        Case "EJ"
                            lngOrdinal = CLng(Split(SegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                            lngOrdinal = ((lngOrdinal - 1) \ 6) + 1
                            rstGroups.Filter = "HEADER = " & CStr(HeaderNumber) & " AND ORDINAL = " & CStr(lngOrdinal)
                        Case Else
                            Debug.Assert False
                    End Select
            End Select
            
            If rstGroups.RecordCount > 0 Then
                Debug.Assert Not rstGroups.EOF
                If Not rstGroups.EOF Then
                    'strReturnValue = rstGroups.Fields(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value).Value
                    strReturnValue = rstGroups.Fields(SourceBoxCode).Value
                End If
            Else
                Debug.Assert rstGroups.RecordCount > 0
            End If
            rstGroups.Filter = adFilterNone
            rstGroups.Sort = vbNullString
        Else
            strReturnValue = vbNullString
        End If
    End If
    
    ADORecordsetClose rstMapConditionsRecordsetClone

    GetLeftOperand = strReturnValue
End Function

Private Function ConditionApplies(ByVal Operator As String, _
                                  ByVal LeftOperand As String, _
                                  ByVal RightOperand As String, _
                                  ByVal OperandsAreNumeric As Boolean) As Boolean
    Dim blnReturnValue As Boolean
    
    Select Case Operator
        Case OPERATOR_EQUAL
            If OperandsAreNumeric Then
                Debug.Assert IsNumeric(RightOperand)
                If IsNumeric(LeftOperand) Then
                    blnReturnValue = (CLng(LeftOperand) = CLng(RightOperand))
                Else
                    blnReturnValue = False
                End If
            Else
                blnReturnValue = ("'" & LeftOperand & "'" = RightOperand)
            End If
        Case OPERATOR_GREATER_THAN
            If OperandsAreNumeric Then
                Debug.Assert IsNumeric(RightOperand)
                If IsNumeric(LeftOperand) Then
                    blnReturnValue = (CLng(LeftOperand) > CLng(RightOperand))
                Else
                    blnReturnValue = False
                End If
            Else
                blnReturnValue = ("'" & LeftOperand & "'" > RightOperand)
            End If
        Case OPERATOR_LESS_THAN
            If OperandsAreNumeric Then
                Debug.Assert IsNumeric(RightOperand)
                If IsNumeric(LeftOperand) Then
                    blnReturnValue = (CLng(LeftOperand) < CLng(RightOperand))
                Else
                    blnReturnValue = False
                End If
            Else
                blnReturnValue = ("'" & LeftOperand & "'" < RightOperand)
            End If
        Case OPERATOR_NOT_EQUAL
            If OperandsAreNumeric Then
                Debug.Assert IsNumeric(RightOperand)
                If IsNumeric(LeftOperand) Then
                    blnReturnValue = (CLng(LeftOperand) <> CLng(RightOperand))
                Else
                    blnReturnValue = False
                End If
            Else
                blnReturnValue = ("'" & LeftOperand & "'" <> RightOperand)
            End If
        ' MUCP-65 - Start
        Case OPERATOR_LEFT
            If OperandsAreNumeric Then
                blnReturnValue = (Left$(LeftOperand, 6) = RightOperand)
            Else
                blnReturnValue = ("'" & Left$(LeftOperand, 6) & "'" = RightOperand)
            End If
            Debug.Assert False
        ' MUCP-65 - End
        
        Case Else
            '----->  Unknow operator
            Debug.Assert False
    End Select
    ConditionApplies = blnReturnValue
End Function

Private Function ConditionIsSatisfied(ByVal MapConditionsRecordset As ADODB.Recordset, _
                                      ByVal UnmappedValuesRecordset As ADODB.Recordset, _
                                      ByVal NCTS_IEM_TMS_ID As Long, _
                                      ByVal DataItemOrdinal As Long, _
                                      ByVal SegmentKey As String, _
                                      ByVal HeaderNumber As Long, _
                                      ByVal DetailNumber As Long, _
                                      ByVal GroupNumber As Long) As Boolean
    Dim blnReturnValue As Boolean
    
    Dim blnContinueLoopThroughMapConditions As Boolean
    Dim blnFoundMapCondition As Boolean
    Dim blnConditionIsSatisfied As Boolean
    Dim blnOperandsAreNumeric As Boolean
    
    Dim strOperator As String
    Dim strLeftOperand As String
    Dim strRightOperand As String
    Dim strBoxCodeGroups As String
    
    Dim varConditions As Variant
    Dim lngConditionIndex As Long
    
    Dim rstMapConditionsRecordsetClone As ADODB.Recordset
    
    blnConditionIsSatisfied = False
    '----->  Search MapConditionsRecordset based on NCTS_IEM_TMS_ID DataItemOrdinal
    
    '----->  Check if NCTS_IEM_TMS_ID is controlled by certain conditions
    If MapConditionsRecordset.RecordCount > 0 Then
        MapConditionsRecordset.MoveFirst
        MapConditionsRecordset.Find "NCTS_IEM_TMS_ID = " & CStr(NCTS_IEM_TMS_ID), , adSearchForward
    End If
    If Not MapConditionsRecordset.EOF Then
        blnContinueLoopThroughMapConditions = True
        Do
            blnFoundMapCondition = (DataItemOrdinal = MapConditionsRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value)
            If blnFoundMapCondition Then
                blnContinueLoopThroughMapConditions = False
            Else
                MapConditionsRecordset.MoveNext
                If MapConditionsRecordset.EOF Then
                    blnContinueLoopThroughMapConditions = False
                Else
                    blnContinueLoopThroughMapConditions = (MapConditionsRecordset.Fields("NCTS_IEM_TMS_ID").Value = NCTS_IEM_TMS_ID)
                End If
            End If
        Loop Until Not blnContinueLoopThroughMapConditions
        If blnFoundMapCondition Then
            ' ********** Modified October 28, 2003 **********
            ' ********** Restricting the NCTS_IEM_MCN_ID values to 14 or less would leave future additional records unaccounted for.
            ' If MapConditionsRecordset.Fields("NCTS_IEM_MCN_ID").Value <= 14 Then
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V1"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V2"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V3"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V4"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V5"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V6"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V7"
            Debug.Assert MapConditionsRecordset.Fields("NCTS_IEM_MAP_Source").Value <> "V8"
            
            If Len(Trim(IIf(IsNull(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value), "", MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value))) Then
            ' ********** End Modify *************************
                If UnmappedValuesRecordset.RecordCount > 0 Then
                    strBoxCodeGroups = GetBoxCodeGroup(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value)
                    strLeftOperand = GetLeftOperand(MapConditionsRecordset, _
                                                    UnmappedValuesRecordset, _
                                                    MapConditionsRecordset.Fields("NCTS_IEM_MCN_Source").Value, _
                                                    strBoxCodeGroups, _
                                                    SegmentKey, _
                                                    HeaderNumber, _
                                                    DetailNumber, _
                                                    blnOperandsAreNumeric)
                    
                    Debug.Assert UCase$(Trim$(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value)) <> "Z1='DG0'|Z1='DG1'&IT=1"
                    Debug.Assert UCase$(Trim$(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value)) <> "Z1='DG0'|Z1='DG1'&IT<>'0'&IT<>'1'"
                    
                    '----->  code in sub temp
                    varConditions = GetConditionTokens(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value)
                    blnConditionIsSatisfied = False
                    
                    Debug.Assert varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand) <> "V1"
                    Debug.Assert varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand) <> "V"
                    
                    For lngConditionIndex = LBound(varConditions) To UBound(varConditions)
                        Select Case varConditions(lngConditionIndex, 2)
                            Case OPERATOR_AND
                                blnConditionIsSatisfied = blnConditionIsSatisfied And ConditionApplies(varConditions(lngConditionIndex, 0), strLeftOperand, varConditions(lngConditionIndex, 1), blnOperandsAreNumeric)
                            Case OPERATOR_OR
                                blnConditionIsSatisfied = blnConditionIsSatisfied Or ConditionApplies(varConditions(lngConditionIndex, 0), strLeftOperand, varConditions(lngConditionIndex, 1), blnOperandsAreNumeric)
                        End Select
                        
                    Next
                Else
                    Debug.Assert UnmappedValuesRecordset.RecordCount > 0
                    blnConditionIsSatisfied = True
                End If
            Else
                If UnmappedValuesRecordset.RecordCount > 0 Then
                    'Debug.Assert UCase$(Trim$(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value)) <> "Z1='DG0'|Z1='DG1'&IT=1"
                    'Debug.Assert UCase$(Trim$(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value)) <> "Z1='DG0'|Z1='DG1'&IT<>'0'&IT<>'1'"
                    
                    varConditions = ParseCondition(IIf(IsNull(MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value), "", MapConditionsRecordset.Fields("NCTS_IEM_MCN_Condition").Value))
                    Debug.Assert varConditions(0, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator) = OPERATOR_OR
                    blnConditionIsSatisfied = False
                    Debug.Assert varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand) <> "V1"
                    Debug.Assert varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand) <> "V2"
                    For lngConditionIndex = LBound(varConditions) To UBound(varConditions)
                        strBoxCodeGroups = GetBoxCodeGroup(varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand))
                        strLeftOperand = GetLeftOperand(MapConditionsRecordset, _
                                                        UnmappedValuesRecordset, _
                                                        varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LeftOperand), _
                                                        strBoxCodeGroups, _
                                                        SegmentKey, _
                                                        HeaderNumber, _
                                                        DetailNumber, _
                                                        blnOperandsAreNumeric)
                        Select Case varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_LogicalOperator)
                            Case OPERATOR_AND
                                blnConditionIsSatisfied = blnConditionIsSatisfied _
                                                          And ConditionApplies(varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_Operator), _
                                                          strLeftOperand, _
                                                          varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_RightOperand), _
                                                          blnOperandsAreNumeric)
                            Case OPERATOR_OR
                                blnConditionIsSatisfied = blnConditionIsSatisfied _
                                                          Or ConditionApplies(varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_Operator), _
                                                          strLeftOperand, _
                                                          varConditions(lngConditionIndex, ConditionTokenIndexes.ConditionTokenIndex_RightOperand), _
                                                          blnOperandsAreNumeric)
                        End Select
                    Next
                Else
                    Debug.Assert UnmappedValuesRecordset.RecordCount > 0
                    blnConditionIsSatisfied = True
                End If
            End If
        Else
            blnConditionIsSatisfied = True
        End If
    Else
        blnConditionIsSatisfied = True
    End If
    
    blnReturnValue = blnConditionIsSatisfied
    ConditionIsSatisfied = blnReturnValue
End Function

Private Function FormatNumericValues(ByVal TargetString As String) As String
    Dim strReturnValue As String
    Dim strLeftOfDecimalPoint As String
    Dim strRightOfDecimalPoint As String
    Dim blnTargetStringWasNullString As Boolean
    
    strReturnValue = Trim(TargetString)
    blnTargetStringWasNullString = (strReturnValue = vbNullString)
    
    '----->  Remove leading zeroes
    Do While Left(strReturnValue, 1) = "0"
        strReturnValue = Mid(strReturnValue, 2)
    Loop
    If strReturnValue = vbNullString And Not blnTargetStringWasNullString Then
        strReturnValue = "0"
    End If
    '----->  Does a decimal point exist?
    If InStr(1, strReturnValue, ".") > 0 Then
        '----->  YES, remove all instances of commas
        strReturnValue = Replace(strReturnValue, ",", vbNullString)
    Else
        '----->  NO, Replace all instances of commas to decimal points
        strReturnValue = Replace(strReturnValue, ",", ".")
    End If
    
    If InStr(1, strReturnValue, ".") > 0 Then
        strLeftOfDecimalPoint = Left(strReturnValue, InStr(1, strReturnValue, ".") - 1)
        strRightOfDecimalPoint = Mid(strReturnValue, InStr(1, strReturnValue, ".") + 1)
        strRightOfDecimalPoint = Replace(strRightOfDecimalPoint, ".", vbNullString)
        strReturnValue = strLeftOfDecimalPoint & "." & strRightOfDecimalPoint
    End If
    FormatNumericValues = strReturnValue
End Function

Private Function ReplaceSpecialCharacters(ByVal SourceString As String) As String
    Dim strReturnValue As String
    strReturnValue = Trim(SourceString)
    If strReturnValue <> vbNullString Then
        '----->  RELEASE CHARACTER
        strReturnValue = Replace(strReturnValue, EDI_SEP_RELEASE_CHARACTER, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER)
        '----->  SEGMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_SEGMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
        '----->  COMPOSITE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_COMPOSITE_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT)
        '----->  SIMPLE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT)
    End If
    ReplaceSpecialCharacters = strReturnValue
End Function

Private Function CreateEDIMessageSegmentData(ByVal SourceSegment As Segment, _
                                             ByVal TMSRecordset As ADODB.Recordset, _
                                             ByVal MapConditionsRecordset As ADODB.Recordset, _
                                             ByVal UnmappedValuesRecordset As ADODB.Recordset, _
                                             ByVal HeaderNumber As Long, _
                                             ByVal DetailNumber As Long, _
                                             ByVal GroupNumber As Long) As String
    Dim strReturnValue                      As String
    
    Dim blnEmptySegment                     As Boolean
    
    Dim lngPreviousGroupID                  As Long
    Dim lngSDItemsIndex                     As Long
    
    Dim strSegmentKey                       As String
    Dim strSegmentData                      As String
    Dim strDataItemValue                    As String
    
    Dim colSDItems                          As SDataItems
    
    Dim strTempSegmentTag As String 'IAN 06-07-2005 Added to be in conditions for empty segment in TOD
    
    strReturnValue = vbNullString
    
    Debug.Assert SourceSegment.SegmentTag <> "GIR"
    
    blnEmptySegment = True
    strSegmentKey = SourceSegment.Key
    strSegmentData = SourceSegment.SegmentTag & EDI_SEP_DATA_ELEMENT
    Set colSDItems = SourceSegment.SDataItems
    
    blnEmptySegment = blnEmptySegment And (colSDItems(1).NCTSDataFormat = "Q")
    
    lngPreviousGroupID = colSDItems(1).EDIGroupID
    
    lngSDItemsIndex = 1
    If ConditionIsSatisfied(MapConditionsRecordset, _
                            UnmappedValuesRecordset, _
                            CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID)), _
                            lngSDItemsIndex, _
                            strSegmentKey, HeaderNumber, DetailNumber, GroupNumber) Then
        strDataItemValue = ProcessValue(colSDItems(lngSDItemsIndex).Value, colSDItems(lngSDItemsIndex).NCTSDataFormat)
'''''            '----->  RICHARD - 06272003
                 '----->  Removed since all leading zeroes in numeric values are removed,
                 '----->  INCLUDING THE DATES (e.g. 030627 => 30627) which is WRONG
'''''            If UCase(Left(colSDItems(lngSDItemsIndex).NCTSDataFormat, 1)) = "N" Then
'''''                strDataItemValue = FormatNumericValues(strDataItemValue)
'''''            End If
        strDataItemValue = ReplaceSpecialCharacters(strDataItemValue)
        strSegmentData = strSegmentData & strDataItemValue
    End If
    
    For lngSDItemsIndex = 2 To colSDItems.Count
        If lngPreviousGroupID = colSDItems(lngSDItemsIndex).EDIGroupID And colSDItems(lngSDItemsIndex).EDIGroupID <> 0 Then
            strSegmentData = strSegmentData & EDI_SEP_COMPOSITE_DATA_ELEMENT
        Else
            strSegmentData = strSegmentData & EDI_SEP_DATA_ELEMENT
        End If
        If ConditionIsSatisfied(MapConditionsRecordset, _
                                UnmappedValuesRecordset, _
                                CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID)), _
                                lngSDItemsIndex, _
                                strSegmentKey, HeaderNumber, DetailNumber, GroupNumber) Then
            'strDataItemValue = colSDItems(lngSDItemsIndex).Value
            strDataItemValue = ProcessValue(colSDItems(lngSDItemsIndex).Value, colSDItems(lngSDItemsIndex).NCTSDataFormat)
            '----->  08282003, RICHARD
            '----->  This code-segment is in compliance with Andrei's request to send only the
            '----->  first 8 of the 10 characters in box code L1.
            If lngSDItemsIndex = 2 And SourceSegment.SegmentTag = "CST" Then
                strDataItemValue = Left(strDataItemValue, 8)
            End If
            '----->  End of code-segment
            
'''''            '----->  RICHARD - 06272003
                 '----->  Removed since all leading zeroes in numeric values are removed,
                 '----->  INCLUDING THE DATES (e.g. 030627 => 30627) which is WRONG
'''''            If UCase(Left(colSDItems(lngSDItemsIndex).NCTSDataFormat, 1)) = "N" Then
'''''                strDataItemValue = FormatNumericValues(strDataItemValue)
'''''            End If
            
            'IAN 03-14-05
            'This group of codes were added to solve the issue in sending Non-conform
            'Unloading Remarks. This will remove the control description and the
            'description language when the control indicator is 'DI'
            If UCase(SourceSegment.SegmentTag) = "FTX" Then
                If UCase(ProcessValue(colSDItems(3).Value, colSDItems(3).NCTSDataFormat)) <> "OT" And _
                    UCase(ProcessValue(colSDItems(1).Value, colSDItems(1).NCTSDataFormat)) = "ABV" Then

                    If lngSDItemsIndex = 11 Or lngSDItemsIndex = 7 Then
                        strDataItemValue = ""
                    End If

                End If
            End If
            
            strDataItemValue = ReplaceSpecialCharacters(strDataItemValue)
            strSegmentData = strSegmentData & strDataItemValue
            
            'Rachelle - for IE34 - 091205
            If SourceSegment.SegmentTag <> "TOD" And SourceSegment.SegmentTag <> "SEQ" Then
                blnEmptySegment = blnEmptySegment And (colSDItems(lngSDItemsIndex).NCTSDataFormat = "Q" Or Trim(colSDItems(lngSDItemsIndex).Value) = "")
            ElseIf SourceSegment.SegmentTag = "SEQ" Then
                blnEmptySegment = False
            Else
                strTempSegmentTag = GetLeftOperand(MapConditionsRecordset, _
                                                UnmappedValuesRecordset, _
                                                "Z1", _
                                                "", _
                                                strSegmentKey, _
                                                HeaderNumber, _
                                                DetailNumber, _
                                                False)
                'blnEmptySegment = UCase(strTempSegmentTag) <> "DG2" And UCase(strTempSegmentTag) <> "SM1" And UCase(strTempSegmentTag) <> "CAL" And blnEmptySegment And (colSDItems(lngSDItemsIndex).NCTSDataFormat = "Q" Or Trim(colSDItems(lngSDItemsIndex).Value) = "")
                blnEmptySegment = UCase(strTempSegmentTag) <> "NPS" And UCase(strTempSegmentTag) <> "DG2" And UCase(strTempSegmentTag) <> "SM1" And UCase(strTempSegmentTag) <> "CAL" And blnEmptySegment And (colSDItems(lngSDItemsIndex).NCTSDataFormat = "Q" Or Trim(colSDItems(lngSDItemsIndex).Value) = "")
            End If
                        
        End If
        lngPreviousGroupID = colSDItems(lngSDItemsIndex).EDIGroupID
    Next
    
    blnEmptySegment = blnEmptySegment And (colSDItems(1).NCTSDataFormat = "Q")
    'blnEmptySegment = blnEmptySegment And (UCase(SourceSegment.EDIUsage) <> "R")
    
    'CSCLP-206: Temporary solution for NAD+TD Segment
    If strSegmentKey <> "S_31_1" Then
        blnEmptySegment = blnEmptySegment And (UCase(SourceSegment.EDIUsage) <> "R")
    End If
    
    strSegmentData = ReduceSegment(strSegmentData)
    
    ' CSCLP-957 - Start
    If UCase$(Trim$(Left$(strSegmentData, 15))) = "MEA+WT+AAA+KGM:" Then
        
        If Len(Trim$(strSegmentData)) > 15 Then
            If Val(Mid(Trim$(strSegmentData), 16)) = 0 Then
                blnEmptySegment = True
            End If
        End If
    End If
    ' CSCLP-957 - End
    
    If strSegmentData <> SourceSegment.SegmentTag And Not blnEmptySegment Then
        strReturnValue = strSegmentData & EDI_SEP_SEGMENT
    End If
    CreateEDIMessageSegmentData = strReturnValue
    
End Function

Private Function CreateEDIMessageSegmentDataOfDecendants(ByVal SourceSegment As Segment, _
                                                         ByVal TMSRecordset As ADODB.Recordset, _
                                                         ByVal MapConditionsRecordset As ADODB.Recordset, _
                                                         ByVal UnmappedValuesRecordset As ADODB.Recordset, _
                                                         ByVal HeaderNumber As Long, _
                                                         ByVal DetailNumber As Long, _
                                                         ByVal GroupNumber As Long, _
                                                         ByVal PutConsignorInHeader As Boolean, _
                                                         ByVal PutConsigneeInHeader As Boolean, _
                                                         ByRef NewSegmentsCount As Long) As String
    Dim strReturnValue As String
    
    Dim strSegmentData As String
    
    Dim astrChildrenKeys() As String
    
    Dim lngChildKeyIndex As Long
    Dim lngNewSegmentsCountToPass As Long
    Dim lngNewSegmentsCountToReturn As Long
    Dim blnIncludeSegmentInMessage As Boolean
    
    Dim lngNCTSIEMTMSID29 As Long
    Dim lngNCTSIEMTMSID30 As Long
    Dim lngNCTSIEMTMSID39 As Long
    Dim lngNCTSIEMTMSID40 As Long
    Dim lngNCTSIEMTMSID770 As Long
    Dim lngNCTSIEMTMSID771 As Long
    Dim lngNCTSIEMTMSID781 As Long
    Dim lngNCTSIEMTMSID782 As Long
        
    'IAN 08-08-2005 => This will get ID for data segments that were hard-coded before
    GetIE13IE15CNCZ lngNCTSIEMTMSID29, lngNCTSIEMTMSID30, lngNCTSIEMTMSID39, _
                    lngNCTSIEMTMSID40, lngNCTSIEMTMSID770, lngNCTSIEMTMSID771, _
                    lngNCTSIEMTMSID781, lngNCTSIEMTMSID782
                    
    strReturnValue = vbNullString
    lngNewSegmentsCountToReturn = 0
    
    Select Case CLng(Split(SourceSegment.Key, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID))
'        Case 29, 771 'IE15 and IE13 respectively
        Case lngNCTSIEMTMSID29, lngNCTSIEMTMSID771 'IE15 and IE13 respectively
            '----->  NAD+CN in header
            Debug.Assert False
            blnIncludeSegmentInMessage = PutConsigneeInHeader
'        Case 39, 781 'IE15 and IE13 respectively
        Case lngNCTSIEMTMSID39, lngNCTSIEMTMSID781 'IE15 and IE13 respectively
            '----->  NAD+CN in detail
            blnIncludeSegmentInMessage = Not PutConsigneeInHeader
'        Case 30, 770 'IE15 and IE13 respectively
        Case lngNCTSIEMTMSID30, lngNCTSIEMTMSID770 'IE15 and IE13 respectively
            '----->  NAD+CZ in header
            Debug.Assert False
            blnIncludeSegmentInMessage = PutConsignorInHeader
'        Case 40, 782 'IE15 and IE13 respectively
        Case lngNCTSIEMTMSID40, lngNCTSIEMTMSID782 'IE15 and IE13 respectively
            '----->  NAD+CZ in detail
            blnIncludeSegmentInMessage = Not PutConsignorInHeader
        Case Else
            blnIncludeSegmentInMessage = True
    End Select
    
    If blnIncludeSegmentInMessage Then
        strSegmentData = CreateEDIMessageSegmentData(SourceSegment, _
                                                     TMSRecordset, _
                                                     MapConditionsRecordset, _
                                                     UnmappedValuesRecordset, _
                                                     HeaderNumber, _
                                                     DetailNumber, _
                                                     GroupNumber)
    End If
    If strSegmentData <> vbNullString Then
        strReturnValue = strSegmentData
        lngNewSegmentsCountToReturn = lngNewSegmentsCountToReturn + 1
        If SourceSegment.KeysChildren <> vbNullString Then
            astrChildrenKeys = Split(SourceSegment.KeysChildren, "|")
            
            For lngChildKeyIndex = LBound(astrChildrenKeys) To UBound(astrChildrenKeys)
                Debug.Assert GetSegmentIndex(astrChildrenKeys(lngChildKeyIndex)) > 0
                If GetSegmentIndex(astrChildrenKeys(lngChildKeyIndex)) > 0 Then
                
                    strSegmentData = CreateEDIMessageSegmentDataOfDecendants(Segments(astrChildrenKeys(lngChildKeyIndex)), _
                                                                             TMSRecordset, _
                                                                             MapConditionsRecordset, _
                                                                             UnmappedValuesRecordset, _
                                                                             HeaderNumber, _
                                                                             DetailNumber, _
                                                                             GroupNumber, _
                                                                             PutConsignorInHeader, _
                                                                             PutConsigneeInHeader, _
                                                                             lngNewSegmentsCountToPass)
                    strReturnValue = strReturnValue & strSegmentData
                    lngNewSegmentsCountToReturn = lngNewSegmentsCountToReturn + lngNewSegmentsCountToPass
                End If
            Next
        End If
    End If
    
    NewSegmentsCount = lngNewSegmentsCountToReturn
    CreateEDIMessageSegmentDataOfDecendants = strReturnValue
    
End Function

Private Function ShallNADSegmentBeInHeader(ByVal NCTS_IEM_TMS_ID As Long) As Boolean
    Dim blnReturnValue As Boolean
    
    Dim lngSegmentInstance As Long
    Dim strTIN As String
    '----->  To be used to store values to compare.
    '----->  Data type is Variant to make it flexible for possible extensions in the future
    '----->  e.g. name, address, etc may be compared in the future
    Dim varPropertiesToCompare As Variant
    Dim blnCompareTIN As Boolean
    
    blnReturnValue = True
    strTIN = vbNullString
    If GetSegmentIndex("S_" & CStr(NCTS_IEM_TMS_ID) & "_1") > 0 Then
        '----->  The second data item is the TIN
        strTIN = Segments("S_" & CStr(NCTS_IEM_TMS_ID) & "_1").SDataItems(2).Value
    Else
        Debug.Assert GetSegmentIndex("S_" & CStr(NCTS_IEM_TMS_ID) & "_1") > 0
    End If
    If strTIN = vbNullString Or strTIN = "0" Then
        ReDim varPropertiesToCompare(0)
        '----->  The 16th data item is the address
        varPropertiesToCompare(0) = Segments("S_" & CStr(NCTS_IEM_TMS_ID) & "_1").SDataItems(16).Value
        blnCompareTIN = False
    Else
        ReDim varPropertiesToCompare(0)
        varPropertiesToCompare(0) = strTIN
        blnCompareTIN = True
    End If
    lngSegmentInstance = 2
    Do While GetSegmentIndex("S_" & CStr(NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)) > 0 And blnReturnValue
        blnReturnValue = (strTIN = Segments("S_" & CStr(NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)).SDataItems(2).Value)
        If blnCompareTIN Then
            blnReturnValue = blnReturnValue And (varPropertiesToCompare(0) = Segments("S_" & CStr(NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)).SDataItems(2).Value)
        Else
            blnReturnValue = blnReturnValue And (varPropertiesToCompare(0) = Segments("S_" & CStr(NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)).SDataItems(16).Value)
        End If
        lngSegmentInstance = lngSegmentInstance + 1
    Loop
    ShallNADSegmentBeInHeader = blnReturnValue
End Function

Private Function GetNCTS_IEM_ID(ByVal MessageType As ENCTSMessageType) As Long
    Dim lngReturnValue As Long
    
'            EMsg_IE18 = 11
'            EMsg_IE25 = 15
    
    lngReturnValue = 0
    Select Case MessageType
        Case ENCTSMessageType.EMsg_IE04
'            EMsg_IE04 = 1
            lngReturnValue = 1
        Case ENCTSMessageType.EMsg_IE07
'            EMsg_IE07 = 3
            lngReturnValue = 2
        Case ENCTSMessageType.EMsg_IE13
'            EMsg_IE13 = 6
            lngReturnValue = 3
        Case ENCTSMessageType.EMsg_IE14
'            EMsg_IE14 = 7
            lngReturnValue = 4
        Case ENCTSMessageType.EMsg_IE15
'            EMsg_IE15 = 8
            lngReturnValue = 5
        Case ENCTSMessageType.EMsg_IE19
'            EMsg_IE19 = 12
            lngReturnValue = 6
        Case ENCTSMessageType.EMsg_IE21
'            EMsg_IE21 = 13
            lngReturnValue = 7
        Case ENCTSMessageType.EMsg_IE23
'            EMsg_IE23 = 14
            lngReturnValue = 8
        Case ENCTSMessageType.EMsg_IE29
'            EMsg_IE29 = 17
            lngReturnValue = 9
        Case ENCTSMessageType.EMsg_IE43
'            EMsg_IE43 = 18
            lngReturnValue = 10
        Case ENCTSMessageType.EMsg_IE44
'            EMsg_IE44 = 19
            lngReturnValue = 11
        Case ENCTSMessageType.EMsg_IE45
'            EMsg_IE45 = 20
            lngReturnValue = 12
        Case ENCTSMessageType.EMsg_IE51
'            EMsg_IE51 = 21
            lngReturnValue = 13
        Case ENCTSMessageType.EMsg_IE54
'            EMsg_IE54 = 22
            lngReturnValue = 14
        Case ENCTSMessageType.EMsg_IE100
'            EMsg_IE100 = 26
            lngReturnValue = 15
        Case ENCTSMessageType.EMsg_IE09
'            EMsg_IE09 = 5
            lngReturnValue = 16
        Case ENCTSMessageType.EMsg_IE25
'            EMsg_IE17 = 10
            lngReturnValue = 17
        Case ENCTSMessageType.EMsg_IE28
'            EMsg_IE28 = 16
            lngReturnValue = 18
        Case ENCTSMessageType.EMsg_IE60
'            EMsg_IE60 = 24
            lngReturnValue = 19
        Case ENCTSMessageType.EMsg_IE05
'            EMsg_IE05 = 2
            lngReturnValue = 20
        Case ENCTSMessageType.EMsg_IE08
'            EMsg_IE08 = 4
            lngReturnValue = 21
        Case ENCTSMessageType.EMsg_IE16
'            EMsg_IE16 = 9
            lngReturnValue = 22
        Case ENCTSMessageType.EMsg_IE58
'            EMsg_IE58 = 23
            lngReturnValue = 23
        Case ENCTSMessageType.EMsg_IE62
'            EMsg_IE62 = 25
            lngReturnValue = 24
        Case ENCTSMessageType.EMsg_IE906
'            EMsg_IE906 = 29
            lngReturnValue = 25
        Case ENCTSMessageType.EMsg_IE907
'            EMsg_IE907 = 30
            lngReturnValue = 26
        Case ENCTSMessageType.EMsg_IE904
'            EMsg_IE904 = 27
            lngReturnValue = 27
        Case ENCTSMessageType.EMsg_IE905
'            EMsg_IE905 = 28
            lngReturnValue = 28
        Case ENCTSMessageType.EMsg_CODEM
            lngReturnValue = 29
        Case ENCTSMessageType.EMsg_IE917
            lngReturnValue = 30
        Case ENCTSMessageType.EMsg_IE34 'Rachelle - for IE34 - 091205
            lngReturnValue = 32
    End Select
    
    GetNCTS_IEM_ID = lngReturnValue
End Function

Private Function GetMessageTypeFromNCTS_IEM_ID(ByVal NCTS_IEM_ID As NCTS_IEM_IDs) As ENCTSMessageType
    Dim enuReturnValue As ENCTSMessageType
    
'            EMsg_IE18 = 11
'            EMsg_IE25 = 15
    
    enuReturnValue = 0
    Select Case NCTS_IEM_ID
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE04
            enuReturnValue = EMsg_IE04
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE07
            enuReturnValue = EMsg_IE07
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE13
            enuReturnValue = EMsg_IE13
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE14
            enuReturnValue = EMsg_IE14
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE15
            enuReturnValue = EMsg_IE15
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE19
            enuReturnValue = EMsg_IE19
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE21
            enuReturnValue = EMsg_IE21
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE23
            enuReturnValue = EMsg_IE23
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE29
            enuReturnValue = EMsg_IE29
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE43
            enuReturnValue = EMsg_IE43
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE44
            enuReturnValue = EMsg_IE44
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE45
            enuReturnValue = EMsg_IE45
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE51
            enuReturnValue = EMsg_IE51
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE54
            enuReturnValue = EMsg_IE54
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE100
            enuReturnValue = EMsg_IE100
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE09
            enuReturnValue = EMsg_IE09
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE25
            enuReturnValue = EMsg_IE25
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE28
            enuReturnValue = EMsg_IE28
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE60
            enuReturnValue = EMsg_IE60
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE05
            enuReturnValue = EMsg_IE05
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE08
            enuReturnValue = EMsg_IE08
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE16
            enuReturnValue = EMsg_IE16
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE58
            enuReturnValue = EMsg_IE58
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE62
            enuReturnValue = EMsg_IE62
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE906
            enuReturnValue = EMsg_IE906
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE907
            enuReturnValue = EMsg_IE907
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE904
            enuReturnValue = EMsg_IE904
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE905
            enuReturnValue = EMsg_IE905
        Case NCTS_IEM_IDs.NCTS_IEM_ID_CODEM
            enuReturnValue = EMsg_CODEM
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE917
            enuReturnValue = EMsg_IE917
        Case NCTS_IEM_IDs.NCTS_IEM_ID_IE34 'Rachelle - for IE34 - 091205
            enuReturnValue = EMsg_IE34
    End Select
    
    GetMessageTypeFromNCTS_IEM_ID = enuReturnValue
End Function

'IAN 08-08-2005 => This will get ID for data segments that were hard-coded before
Private Sub GetIE13IE15CNCZ(ByRef lngNCTSIEMTMSID29 As Long, ByRef lngNCTSIEMTMSID30 As Long, ByRef lngNCTSIEMTMSID39 As Long, _
                    ByRef lngNCTSIEMTMSID40 As Long, ByRef lngNCTSIEMTMSID770 As Long, ByRef lngNCTSIEMTMSID771 As Long, _
                    ByRef lngNCTSIEMTMSID781 As Long, ByRef lngNCTSIEMTMSID782 As Long)
                    
    Dim rstNCTSIEMTMS As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT NCTS_IEM_TMS_ID, " & _
        "NCTS_IEM_TMS_RemarksIEMName, " & _
        "NCTS_IEM_TMS_RemarksSegmentTag, " & _
        "NCTS_IEM_TMS_RemarksTMSSequence, " & _
        "NCTS_IEM_TMS_RemarksQualifier " & _
        "FROM NCTS_IEM_TMS " & _
        "WHERE (NCTS_IEM_TMS_RemarksQualifier='CN' OR NCTS_IEM_TMS_RemarksQualifier='CZ') AND " & _
        "(NCTS_IEM_TMS_RemarksTMSSequence=97 OR NCTS_IEM_TMS_RemarksTMSSequence=22) AND " & _
        "(NCTS_IEM_TMS_RemarksSegmentTag='NAD') AND " & _
        "(NCTS_IEM_TMS_RemarksIEMName='IE13' OR NCTS_IEM_TMS_RemarksIEMName='IE15') " & _
        "ORDER BY NCTS_IEM_TMS_ID"
    
    ADORecordsetOpen strSQL, mconADOConnectionStandard, rstNCTSIEMTMS, adOpenKeyset, adLockOptimistic
    'Set rstNCTSIEMTMS = New ADODB.Recordset
    'rstNCTSIEMTMS.Open strSQL, mconADOConnectionStandard, adOpenKeyset, adLockBatchOptimistic
    
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE15' AND NCTS_IEM_TMS_RemarksQualifier='CN' AND NCTS_IEM_TMS_RemarksTMSSequence=22"
    lngNCTSIEMTMSID29 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE15' AND NCTS_IEM_TMS_RemarksQualifier='CZ' AND NCTS_IEM_TMS_RemarksTMSSequence=22"
    lngNCTSIEMTMSID30 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE15' AND NCTS_IEM_TMS_RemarksQualifier='CN' AND NCTS_IEM_TMS_RemarksTMSSequence=97"
    lngNCTSIEMTMSID39 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE15' AND NCTS_IEM_TMS_RemarksQualifier='CZ' AND NCTS_IEM_TMS_RemarksTMSSequence=97"
    lngNCTSIEMTMSID40 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE13' AND NCTS_IEM_TMS_RemarksQualifier='CZ' AND NCTS_IEM_TMS_RemarksTMSSequence=22"
    lngNCTSIEMTMSID770 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE13' AND NCTS_IEM_TMS_RemarksQualifier='CN' AND NCTS_IEM_TMS_RemarksTMSSequence=22"
    lngNCTSIEMTMSID771 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE13' AND NCTS_IEM_TMS_RemarksQualifier='CN' AND NCTS_IEM_TMS_RemarksTMSSequence=97"
    lngNCTSIEMTMSID781 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Filter = adFilterNone
    rstNCTSIEMTMS.Filter = "NCTS_IEM_TMS_RemarksIEMName='IE13' AND NCTS_IEM_TMS_RemarksQualifier='CZ' AND NCTS_IEM_TMS_RemarksTMSSequence=97"
    lngNCTSIEMTMSID782 = rstNCTSIEMTMS!NCTS_IEM_TMS_ID
    
    rstNCTSIEMTMS.Close
    Set rstNCTSIEMTMS = Nothing
                    
End Sub

Public Function CreateEDIMessage() As String
    Dim strReturnValue                          As String
    Dim strSegmentData                          As String
    Dim strBuffer                               As String
    Dim strSegmentKey                           As String
    Dim strSegmentKeySubString                  As String
    Dim SQLMapConditions                        As String
    Dim SQLUnmappedValues                       As String
    
    Dim lngSegmentIndex                         As Long
    Dim lngSegmentKeySubStringLength            As Long
    Dim lngSegmentCount                         As Long
    Dim lngSealsCount                           As Long
    Dim lngHeaderNumber                         As Long
    Dim lngDetailNumber                         As Long
    Dim lngNewSegmentsCount                     As Long
    Dim lngStartIndex                           As Long
    Dim lngEndIndex                             As Long
    
    Dim blnPutConsignorInHeader                 As Boolean
    Dim blnPutConsigneeInHeader                 As Boolean
    Dim blnIncludeSegmentInMessage              As Boolean
    
    Dim colSDItems                              As SDataItems
    
    Dim rstMapConditions                        As ADODB.Recordset
    Dim rstUnmappedValues                       As ADODB.Recordset
    
    Dim lngNCTSIEMTMSID29 As Long
    Dim lngNCTSIEMTMSID30 As Long
    Dim lngNCTSIEMTMSID39 As Long
    Dim lngNCTSIEMTMSID40 As Long
    Dim lngNCTSIEMTMSID770 As Long
    Dim lngNCTSIEMTMSID771 As Long
    Dim lngNCTSIEMTMSID781 As Long
    Dim lngNCTSIEMTMSID782 As Long
    
    On Error GoTo ErrHandler
    
    'IAN 08-08-2005 => This will get ID for data segments that were hard-coded before
    GetIE13IE15CNCZ lngNCTSIEMTMSID29, lngNCTSIEMTMSID30, lngNCTSIEMTMSID39, _
                    lngNCTSIEMTMSID40, lngNCTSIEMTMSID770, lngNCTSIEMTMSID771, _
                    lngNCTSIEMTMSID781, lngNCTSIEMTMSID782
                    
10  strReturnValue = vbNullString
    
15  SQLMapConditions = "SELECT NCTS_IEM_MAP.NCTS_IEM_MAP_Source, NCTS_IEM_MAP.NCTS_IEM_TMS_ID, NCTS_IEM_MAP.NCTS_IEM_MAP_ParentID, NCTS_IEM_MAP.NCTS_IEM_MAP_EDI_ITM_ORDINAL, " & _
                       "NCTS_IEM_MAP.NCTS_IEM_MCN_ID, NCTS_IEM_MAP_CONDITIONS.NCTS_IEM_MCN_SOURCE, NCTS_IEM_MAP_CONDITIONS.NCTS_IEM_MCN_Condition " & _
                       "FROM NCTS_IEM_MAP LEFT OUTER JOIN NCTS_IEM_MAP_CONDITIONS ON NCTS_IEM_MAP.NCTS_IEM_MCN_ID = NCTS_IEM_MAP_CONDITIONS.NCTS_IEM_MCN_ID " & _
                       "WHERE NCTS_IEM_ID = " & CStr(GetNCTS_IEM_ID(mvarMessageType)) & " " & _
                       "ORDER BY NCTS_IEM_MAP.NCTS_IEM_TMS_ID"
    
20  ADORecordsetOpen SQLMapConditions, mconADOConnectionShaping, rstMapConditions, adOpenKeyset, adLockReadOnly, , True
    
25  rstMapConditions.Filter = "NCTS_IEM_MCN_ID > 0"
    
    '                   "SHAPE {SELECT * FROM DATA_NCTS WHERE CODE = '" & mstrUniqueCode & "'} "
30  SQLUnmappedValues = "SHAPE {SELECT '" & mstrUniqueCode & "' AS CODE} " & _
                        "APPEND " & _
                        "({SELECT * FROM DATA_NCTS_HEADER} AS " & BOX_GROUP_HEADER & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_HEADER_ZEKERHEID} AS " & BOX_GROUP_HEADER_ZEKERHEID & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_DETAIL} AS " & BOX_GROUP_DETAIL & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_DETAIL_BIJZONDERE} AS " & BOX_GROUP_DETAIL_BIJZONDERE & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_DETAIL_COLLI} AS " & BOX_GROUP_DETAIL_COLLI & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_DETAIL_CONTAINER} AS " & BOX_GROUP_DETAIL_CONTAINER & " RELATE CODE TO CODE), " & _
                        "({SELECT * FROM DATA_NCTS_DETAIL_DOCUMENTEN} AS " & BOX_GROUP_DETAIL_DOCUMENTEN & " RELATE CODE TO CODE)"
    
35  ADORecordsetOpen SQLUnmappedValues, mconADOConnectionShaping, rstUnmappedValues, adOpenKeyset, adLockReadOnly, , True

40  Call AddMissingRequiredSegments
    
45  lngSegmentCount = 0
50  lngSealsCount = 0
55  lngHeaderNumber = 1
60  lngDetailNumber = 1
    
    'IAN 08-08-2005 => replaced contants values with values contained from database
65  If MessageType = EMsg_IE15 Then
'70      blnPutConsigneeInHeader = ShallNADSegmentBeInHeader(39)
70      blnPutConsigneeInHeader = ShallNADSegmentBeInHeader(lngNCTSIEMTMSID39)
'75      blnPutConsignorInHeader = ShallNADSegmentBeInHeader(40)
75      blnPutConsignorInHeader = ShallNADSegmentBeInHeader(lngNCTSIEMTMSID40)
80  ElseIf MessageType = EMsg_IE13 Then
'85      blnPutConsigneeInHeader = ShallNADSegmentBeInHeader(781)
85      blnPutConsigneeInHeader = ShallNADSegmentBeInHeader(lngNCTSIEMTMSID781)
'90      blnPutConsignorInHeader = ShallNADSegmentBeInHeader(782)
90      blnPutConsignorInHeader = ShallNADSegmentBeInHeader(lngNCTSIEMTMSID782)
    Else
95      blnPutConsigneeInHeader = False
100     blnPutConsignorInHeader = False
    End If
    
105 mrstTechnicalMessageStructure.MoveFirst
    
    Dim lngLastDetailWithF As Long
    lngLastDetailWithF = GetLastDetailwithF(mstrUniqueCode)
    
110 Do While Not mrstTechnicalMessageStructure.EOF
115     strSegmentKeySubString = "S_" & CStr(mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value) & "_"
            
        'IAN 08-08-2005 => replaced contants values with values contained from database
120     Select Case mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value
'            Case 29, 771 'IE15 and IE13 respectively
            Case lngNCTSIEMTMSID29, lngNCTSIEMTMSID771 'IE15 and IE13 respectively
                '----->  NAD+CN in header
125             blnIncludeSegmentInMessage = blnPutConsigneeInHeader
'130         Case 39, 781 'IE15 and IE13 respectively
130         Case lngNCTSIEMTMSID39, lngNCTSIEMTMSID781 'IE15 and IE13 respectively
                '----->  NAD+CN in detail
135             blnIncludeSegmentInMessage = Not blnPutConsigneeInHeader
'140         Case 30, 770 'IE15 and IE13 respectively
140         Case lngNCTSIEMTMSID30, lngNCTSIEMTMSID770 'IE15 and IE13 respectively
                '----->  NAD+CZ in header
145             blnIncludeSegmentInMessage = blnPutConsignorInHeader
'150         Case 40, 782 'IE15 and IE13 respectively
150         Case lngNCTSIEMTMSID40, lngNCTSIEMTMSID782 'IE15 and IE13 respectively
                '----->  NAD+CZ in detail
155             blnIncludeSegmentInMessage = Not blnPutConsignorInHeader
160         Case Else
165             blnIncludeSegmentInMessage = True
        End Select
        
170     If blnIncludeSegmentInMessage Then
175         lngSegmentKeySubStringLength = Len(strSegmentKeySubString)

180         For lngSegmentIndex = 1 To Segments.Count
185             If Left(Segments(lngSegmentIndex).Key, lngSegmentKeySubStringLength) = strSegmentKeySubString Then
190                 Set colSDItems = Segments(lngSegmentIndex).SDataItems
                    
195                 If Segments(lngSegmentIndex).SegmentTag = "UNT" Then
200                     lngSegmentCount = 0
205                     lngStartIndex = 1
210                     lngEndIndex = 1
                        
215                     strBuffer = Replace(strReturnValue, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, vbNullString)
                        
                        Do
220                         lngEndIndex = InStr(lngStartIndex, strBuffer, EDI_SEP_SEGMENT)
                            
225                         If lngEndIndex > 0 Then
230                             lngSegmentCount = lngSegmentCount + 1
235                             lngStartIndex = lngEndIndex + 1
                            End If
240                     Loop Until lngEndIndex = 0
                        
245                     colSDItems(1).Value = CStr(lngSegmentCount)
                    End If
                    
250                 If Segments(lngSegmentIndex).SegmentTag = "CNT" And Segments(lngSegmentIndex).SDataItems(1).Value = "16" Then
255                     colSDItems(2).Value = CStr(lngSealsCount)
                    End If
                    
260                 If Segments(lngSegmentIndex).SegmentTag = "CST" Then
265                     lngDetailNumber = Val(colSDItems(1).Value)
                    End If
                    '-------------------------------------------------------------------------'
                    '     NOTE: ONLY SEGMENTS WITHOUT PARENTS ARE PROCESSED AT THIS LEVEL     '
                    '-------------------------------------------------------------------------'
270                 If Segments(lngSegmentIndex).KeysChildren = vbNullString And Segments(lngSegmentIndex).KeyParent = vbNullString And IIf(Segments(lngSegmentIndex).SegmentTag = "CST", lngDetailNumber <= lngLastDetailWithF, True) Then
275                     strSegmentData = CreateEDIMessageSegmentData(Segments(lngSegmentIndex), _
                                                                     mrstTechnicalMessageStructure, _
                                                                     rstMapConditions, _
                                                                     rstUnmappedValues, _
                                                                     lngHeaderNumber, _
                                                                     lngDetailNumber, _
                                                                     1)
280                     If strSegmentData <> vbNullString Then
285                         strReturnValue = strReturnValue & strSegmentData
290                         lngSegmentCount = lngSegmentCount + 1
                            
295                         If Segments(lngSegmentIndex).SegmentTag = "SEL" Then
300                             lngSealsCount = lngSealsCount + 1
                            End If
                        End If
305                 ElseIf Segments(lngSegmentIndex).KeysChildren <> vbNullString And Segments(lngSegmentIndex).KeyParent = vbNullString And IIf(Segments(lngSegmentIndex).SegmentTag = "CST", lngDetailNumber <= lngLastDetailWithF, True) Then
310                     strSegmentData = CreateEDIMessageSegmentDataOfDecendants(Segments(lngSegmentIndex), _
                                                                                 mrstTechnicalMessageStructure, _
                                                                                 rstMapConditions, _
                                                                                 rstUnmappedValues, _
                                                                                 lngHeaderNumber, _
                                                                                 lngDetailNumber, _
                                                                                 1, _
                                                                                 blnPutConsignorInHeader, blnPutConsigneeInHeader, lngNewSegmentsCount)
315                     If strSegmentData <> vbNullString Then
320                         strReturnValue = strReturnValue & strSegmentData
325                         lngSegmentCount = lngSegmentCount + lngNewSegmentsCount
                            
330                         If Segments(lngSegmentIndex).SegmentTag = "SEL" Then
335                             lngSealsCount = lngSealsCount + 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
340     mrstTechnicalMessageStructure.MoveNext
    Loop
    
345 ADORecordsetClose rstMapConditions
350 ADORecordsetClose rstUnmappedValues
    
355 CreateEDIMessage = strReturnValue
    
360 mvarEDIMessage = strReturnValue
    
    Exit Function
    
ErrHandler:
    
    Err.Raise Err.Number, "prjEdifactDll.EdifactMessage.CreateEDIMessage", Err.Source & " (" & Erl & "): " & Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub CommitToFile(ByVal FilePath As String)
    Dim intFileHandle As Integer
    
    intFileHandle = FreeFile
    Open FilePath For Output As intFileHandle
    Print #intFileHandle, CreateEDIMessage;
    Close intFileHandle
End Sub

Public Function Commit(ByVal MessageStatus As MessageStatusTypes, ByVal DataNCTSID As Long, ByVal UserID As Long, Optional ByVal ErrorReply As String) As Long
    Dim lngReturnValue As Long
    
    lngReturnValue = 0
    
    Select Case MessageStatus
        Case MessageStatusTypes.MsgStatus_Document
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_DOCUMENT, UserID)
        Case MessageStatusTypes.MsgStatus_QueuedOnly
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_QUEUED, UserID)
        Case MessageStatusTypes.MsgStatus_Queued
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_DOCUMENT, UserID)
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_QUEUED, UserID)
        Case MessageStatusTypes.MsgStatus_Received
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_RECEIVED, UserID, ErrorReply)
        Case MessageStatusTypes.MsgStatus_Sent
            lngReturnValue = CommitToDatabase(DataNCTSID, MESSAGE_STATUS_SENT, UserID)
    End Select
    
    Commit = lngReturnValue
End Function

Public Property Let MessageType(ByVal vData As ENCTSMessageType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MessageType = 5
    mvarMessageType = vData
End Property


Public Property Get MessageType() As ENCTSMessageType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MessageType
    MessageType = mvarMessageType
End Property

Public Property Get Segments() As Segments
'    If mvarSegments Is Nothing Then
'        Set mvarSegments = New Segments
'    End If
    Set Segments = mvarSegments
End Property

Public Property Set Segments(vData As Segments)
    Set mvarSegments = vData
End Property

Private Sub Class_Initialize()
    Set mvarSegments = New Segments
    Set mvarSegmentCounts = New SegmentCounts
End Sub

Private Sub Class_Terminate()
    ADORecordsetClose mrstTechnicalMessageStructure
    
    Set mvarSegmentCounts = Nothing
    Set mvarSegments = Nothing
    
    ADODisconnectDB mconADOConnectionShaping
End Sub

Public Property Get EDIConnection() As ADODB.Connection
    Set EDIConnection = mconADOConnectionStandard
End Property

Public Property Set EDIConnection(ByVal vNewValue As ADODB.Connection)
    Set mconADOConnectionStandard = vNewValue
End Property
            
Public Function SaveDataNCTSMessage(ByRef EDIConnection As ADODB.Connection, _
                                    ByRef DataNCTSMessagesDB As ADODB.Connection, _
                                    ByRef DataNCTS As CDataNCTS, _
                                    Optional ByVal MessageID As Long = 0) _
                                    As Boolean
    Dim lngDataMessageID As Long
    Dim clsSegment As Segment
    Dim clsDataItem As SDataItem
    
    Dim strSegmentTag As String
    Dim lngNCTS_IEM_TMS_ID As Long
    Dim strCommand As String
    Dim blnNoMatch As Boolean
    
    
    SaveDataNCTSMessage = False
    
    ' Open Currentlt Edited Document
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "NCTS_IEM_ID, "
        strCommand = strCommand & "Data_NCTS_ID, "
        strCommand = strCommand & "DATA_NCTS_MSG_StatusType, "
        strCommand = strCommand & "User_ID, "
        strCommand = strCommand & "DATA_NCTS_MSG_Date "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "Data_NCTS_Messages "
    ADORecordsetOpen strCommand, DataNCTSMessagesDB, rstNCTSDataMessage, adOpenKeyset, adLockOptimistic
    'Call DAOOpenTable(DAODatabase, rstNCTSDataMessage, "Data_NCTS_Messages")
    With rstNCTSDataMessage
        
        blnNoMatch = (.EOF And .BOF)
        
        If Not blnNoMatch Then
            .MoveFirst
            .Find "DATA_NCTS_MSG_ID = " & MessageID, , adSearchForward
            
            blnNoMatch = .EOF
        End If
        
        If blnNoMatch Then
            .AddNew
            !NCTS_IEM_ID = G_IE15_NCTS_IEM_ID
            !Data_NCTS_ID = DataNCTS.Data_NCTS_ID
                        
            ' Add Code - if Sending then use --- EMsgStatusType_Send
            !DATA_NCTS_MSG_StatusType = GetMessageStatusType(EMsgStatusType_Document)
            !User_ID = DataNCTS.UserID
        End If
                
        !DATA_NCTS_MSG_Date = Now
        .Update
        
        If blnNoMatch Then
            InsertRecordset DataNCTSMessagesDB, rstNCTSDataMessage, "Data_NCTS_Messages"
        Else
            UpdateRecordset DataNCTSMessagesDB, rstNCTSDataMessage, "Data_NCTS_Messages"
        End If
            
'''''            .Index = "DATA_NCTS_MSG_ID"
'''''            .Seek "=", MessageID
'''''
'''''            If .NoMatch Then
'''''                .AddNew
'''''                !NCTS_IEM_ID = G_IE15_NCTS_IEM_ID
'''''                !Data_NCTS_ID = DataNCTS.Data_NCTS_ID
'''''
'''''                ' Add Code - if Sending then use --- EMsgStatusType_Send
'''''                !DATA_NCTS_MSG_StatusType = GetMessageStatusType(EMsgStatusType_Document)
'''''                !User_ID = DataNCTS.UserID
'''''            End If
'''''
'''''            !DATA_NCTS_MSG_Date = Now
'''''            .Update
    End With
    
    lngDataMessageID = rstNCTSDataMessage.Fields("DATA_NCTS_MSG_ID").Value
    
    ADORecordsetClose rstNCTSDataMessage
    'rstNCTSDataMessage.Close
    
    With Me
        For Each clsSegment In .Segments
            strSegmentTag = clsSegment.SegmentTag
            lngNCTS_IEM_TMS_ID = clsSegment.NCTS_IEM_TMS_ID
                   
                       
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "DATA_NCTS_MSG_ID, "
                strCommand = strCommand & "NCTS_ITM_" & strSegmentTag & "_ID, "
            
                For Each clsDataItem In clsSegment.SDataItems
                    strCommand = strCommand & "DATA_NCTS_" & strSegmentTag & "_Seq" & clsDataItem.EDIOrdinal & ", "
                Next
                
                strCommand = Trim$(strCommand)
                strCommand = Left$(strCommand, Len(strCommand) - 1)
                strCommand = strCommand & " "
                
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "Data_NCTS_" & strSegmentTag & " "
            ADORecordsetOpen strCommand, DataNCTSMessagesDB, rstNCTSDataTable, adOpenKeyset, adLockOptimistic
            'Call DAOOpenTable(DAODatabase, rstNCTSDataTable, "Data_NCTS_" & strSegmentTag)
            
                strCommand = vbNullString
                strCommand = strCommand & "DELETE "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "Data_NCTS_" & strSegmentTag & " "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "DATA_NCTS_MSG_ID = " & lngDataMessageID & " "
            ExecuteNonQuery DataNCTSMessagesDB, strCommand
            
            'rstNCTSDataTable.Index = "DATA_NCTS_MSG_ID"
            'rstNCTSDataTable.Seek "=", lngDataMessageID
            'If Not rstNCTSDataTable.NoMatch Then
            '    Do While rstNCTSDataTable!DATA_NCTS_MSG_ID = lngDataMessageID
            '        rstNCTSDataTable.Delete
            '    Loop
            'End If
                                   
            rstNCTSDataTable.AddNew
            rstNCTSDataTable.Fields("DATA_NCTS_MSG_ID").Value = lngDataMessageID
            rstNCTSDataTable.Fields("NCTS_ITM_" & strSegmentTag & "_ID").Value = lngNCTS_IEM_TMS_ID
            
            For Each clsDataItem In clsSegment.SDataItems
                rstNCTSDataTable.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & clsDataItem.EDIOrdinal).Value = clsDataItem.Value
            Next
            rstNCTSDataTable.Update
            
            InsertRecordset DataNCTSMessagesDB, rstNCTSDataTable, "Data_NCTS_" & strSegmentTag
            
            ADORecordsetClose rstNCTSDataTable
            'rstNCTSDataTable.Close
        Next
    End With

    SaveDataNCTSMessage = True
End Function

Public Sub PrepareMessageEmpty(ByVal NCTSMessageType As ENCTSMessageType, ByVal PrepareSegments As Boolean)
    Dim SQLTechnicalMessageStructure As String
    
    SQLTechnicalMessageStructure = "SHAPE {" & GetQryMessageTechnicalStructure(NCTSMessageType) & "} " & _
                                   "APPEND " & PrepareMessageDataItems(NCTSMessageType) & _
                                   "({SELECT * FROM EDI_TMS_ITEMS ORDER BY EDI_TMS_ITM_Ordinal} AS EDI_TMS_ITEMS RELATE EDI_TMS_SEG_ID TO EDI_TMS_SEG_ID) "
    ADORecordsetOpen SQLTechnicalMessageStructure, mconADOConnectionShaping, mrstTechnicalMessageStructure, adOpenKeyset, adLockReadOnly, , True
    Set mvarSegments.TechnicalMessageStructure = mrstTechnicalMessageStructure
    
    mvarMessageType = NCTSMessageType
    mvarNCTS_IEM_ID = mrstTechnicalMessageStructure.Fields("NCTS_IEM_ID").Value
End Sub

Private Function PrepareMessageDataItems(ByVal NCTSMessageType As ENCTSMessageType) As String
    Dim strDummy As String
    Dim strSegmentTag As String
    
    Dim rstMessageDataTables As ADODB.Recordset
    
    strDummy = GetQryMessageDataTables(NCTSMessageType)
    ADORecordsetOpen strDummy, mconADOConnectionShaping, rstMessageDataTables, adOpenKeyset, adLockOptimistic, , True
    
    If rstMessageDataTables.RecordCount > 0 Then
        strDummy = ""
        rstMessageDataTables.MoveFirst
        Do Until rstMessageDataTables.EOF
            strSegmentTag = rstMessageDataTables.Fields("SegmentTag").Value
            strDummy = strDummy & "({SELECT * FROM NCTS_ITM_" & strSegmentTag & "} AS " & strSegmentTag & " RELATE NCTS_IEM_TMS_ID TO NCTS_IEM_TMS_ID), "
            rstMessageDataTables.MoveNext
        Loop
    End If
    
    PrepareMessageDataItems = strDummy
End Function

Public Sub AddEmptySegment(ByRef TMSShapeRecords As ADODB.Recordset, _
                           ByVal NCTS_IEM_TMS_ID As Long, _
                           ByVal SegmentInstance As Long, _
                           ByVal ParentInstance As Long)
    Debug.Assert TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value = NCTS_IEM_TMS_ID
    SegmentCounts.Add "S" & TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value
    If SegmentInstance = 0 Then
        Segments.Add TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value, SegmentCounts("S" & TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value).Count, ParentInstance
    Else
        Segments.Add TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value, SegmentInstance, ParentInstance
    End If
End Sub

Public Function PrepareDataNCTSMessage(ByVal NCTSMessageType As ENCTSMessageType, _
                                       ByVal InternalCode As String, _
                                       ByVal MessageID As Long) _
                                       As String
    Dim rstMessageSheet As ADODB.Recordset
    Dim SQLTechnicalMessageStructure As String
    
    SQLTechnicalMessageStructure = "SHAPE {" & GetQryNCTSTMSDataMessage(MessageID) & "} " & _
                                   "APPEND " & _
                                   PrepareDataNCTSMessageItems(EMsg_IE15, MessageID) & _
                                   "({SELECT * FROM EDI_TMS_ITEMS ORDER BY EDI_TMS_ITM_Ordinal} AS EDI_TMS_ITEMS RELATE EDI_TMS_SEG_ID TO EDI_TMS_SEG_ID) "
    ADORecordsetOpen SQLTechnicalMessageStructure, mconADOConnectionShaping, mrstTechnicalMessageStructure, adOpenKeyset, adLockReadOnly, , True
    Set mvarSegments.TechnicalMessageStructure = mrstTechnicalMessageStructure
    
    mvarMessageType = NCTSMessageType
    
    If mrstTechnicalMessageStructure.RecordCount > 0 Then
        Do Until mrstTechnicalMessageStructure.EOF
            AddDataNCTSMessageSegments mrstTechnicalMessageStructure.Fields("EDI_TMS_SEG_Tag").Value, mrstTechnicalMessageStructure
            mrstTechnicalMessageStructure.MoveNext
        Loop
    End If
End Function

Private Function PrepareDataNCTSMessageItems(ByVal NCTSMessageType As ENCTSMessageType, _
                                                ByVal MessageID As Long) _
                                            As String
    Dim strDummy As String
    Dim strSegmentTag As String
    
    Dim rstMessageDataTables As ADODB.Recordset
    
    strDummy = GetQryMessageDataTables(NCTSMessageType)
    ADORecordsetOpen strDummy, mconADOConnectionShaping, rstMessageDataTables, adOpenKeyset, adLockOptimistic, , True
    
    If rstMessageDataTables.RecordCount > 0 Then
        strDummy = ""
        rstMessageDataTables.MoveFirst
        Do Until rstMessageDataTables.EOF
            strSegmentTag = rstMessageDataTables.Fields("SegmentTag").Value
            strDummy = strDummy & "((SHAPE {SELECT * FROM DATA_NCTS_" & strSegmentTag & " WHERE DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = " & MessageID & "} AS DATA_NCTS_" & strSegmentTag & " APPEND ({SELECT * FROM NCTS_ITM_" & strSegmentTag & "} AS NCTS_ITM_" & strSegmentTag & " RELATE NCTS_IEM_TMS_ID TO NCTS_IEM_TMS_ID)) RELATE NCTS_IEM_TMS_ID TO NCTS_IEM_TMS_ID), "
            rstMessageDataTables.MoveNext
        Loop
    End If
    
    ADORecordsetClose rstMessageDataTables
    PrepareDataNCTSMessageItems = strDummy
End Function

Private Function PrepareDataNCTSSegments(ByVal NCTSMessageType As ENCTSMessageType, _
                                         ByVal MessageID As Long) As String
    Dim strReturnValue As String
    
    Dim SQLMessageDataTables As String
    Dim strSegmentTag As String
    
    Dim rstMessageDataTables As ADODB.Recordset
    
    SQLMessageDataTables = GetQryMessageDataTables(NCTSMessageType)
    ADORecordsetOpen SQLMessageDataTables, mconADOConnectionShaping, rstMessageDataTables, adOpenKeyset, adLockOptimistic, , True
    
    Debug.Assert rstMessageDataTables.RecordCount > 0
    If rstMessageDataTables.RecordCount > 0 Then
        strReturnValue = vbNullString
        rstMessageDataTables.MoveFirst
        Do Until rstMessageDataTables.EOF
            strSegmentTag = rstMessageDataTables.Fields("SegmentTag").Value
            strReturnValue = strReturnValue & "({SELECT DATA_NCTS_" & strSegmentTag & ".*, EDI_TMS_SEGMENTS.EDI_TMS_SEG_TAG AS PARENT_SEGMENT_TAG "
            strReturnValue = strReturnValue & "FROM (DATA_NCTS_" & strSegmentTag & " INNER JOIN NCTS_IEM_TMS ON DATA_NCTS_" & strSegmentTag & ".NCTS_IEM_TMS_ID = NCTS_IEM_TMS.NCTS_IEM_TMS_ID) LEFT OUTER JOIN (NCTS_IEM_TMS AS PARENT LEFT OUTER JOIN (EDI_TMS LEFT OUTER JOIN EDI_TMS_SEGMENTS ON EDI_TMS.EDI_TMS_SEG_ID = EDI_TMS_SEGMENTS.EDI_TMS_SEG_ID) ON PARENT.EDI_TMS_ID = EDI_TMS.EDI_TMS_ID) ON NCTS_IEM_TMS.NCTS_IEM_TMS_ParentID = PARENT.NCTS_IEM_TMS_ID "
            strReturnValue = strReturnValue & "WHERE DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = " & MessageID & " "
            strReturnValue = strReturnValue & "ORDER BY DATA_NCTS_" & strSegmentTag & ".NCTS_IEM_TMS_ID, DATA_NCTS_" & strSegmentTag & "_ParentID, DATA_NCTS_" & strSegmentTag & "_Instance} "
            strReturnValue = strReturnValue & "AS DATA_NCTS_" & strSegmentTag & " "
            strReturnValue = strReturnValue & "RELATE DATA_NCTS_MSG_ID TO DATA_NCTS_MSG_ID)"
            rstMessageDataTables.MoveNext
            If Not rstMessageDataTables.EOF Then
                strReturnValue = strReturnValue & ", "
            End If
        Loop
    End If
    
    ADORecordsetClose rstMessageDataTables
    
    PrepareDataNCTSSegments = strReturnValue
End Function

Public Sub AddDataNCTSMessageSegments(ByVal SegmentTag As String, _
                                    ByRef TMSShapeRecords As ADODB.Recordset)


'''''
'''''    Dim rstNCTSItems As ADODB.Recordset
'''''    Dim rstEDIItems As ADODB.Recordset
'''''    Dim rstNCTSDataItems As ADODB.Recordset
'''''
'''''    Dim strSegmentKey As String
'''''    Dim lngNCTS_IEM_TMS_ID As Long
'''''    Dim lngNCTS_ITM_ID As Long
'''''    Dim lngItemCodelistID As Long
'''''    Dim strDataItemValue As String
'''''    Dim lngDataItemOrdinal As Long
'''''
'''''    Dim colDataItems As SDataItems
'''''
'''''    Set rstEDIItems = TMSShapeRecords.Fields("EDI_TMS_ITEMS").Value
'''''    Set rstNCTSDataItems = TMSShapeRecords.Fields("DATA_NCTS_" & SegmentTag).Value
'''''
'''''    lngNCTS_IEM_TMS_ID = TMSShapeRecords.Fields("NCTS_IEM_TMS_ID").Value
'''''
'''''    If rstNCTSDataItems.RecordCount > 0 Then
'''''        rstNCTSDataItems.MoveFirst
'''''        Do Until rstNCTSDataItems.EOF
'''''            SegmentCounts.Add "S" & lngNCTS_IEM_TMS_ID
'''''
'''''            strSegmentKey = "S" & lngNCTS_IEM_TMS_ID & "_" & SegmentCounts.Item("S" & lngNCTS_IEM_TMS_ID).Count
'''''
'''''            Set rstNCTSItems = rstNCTSDataItems.Fields("NCTS_ITM_" & SegmentTag).Value
'''''
'''''            Segments.Add strSegmentKey, _
'''''                        lngNCTS_IEM_TMS_ID, _
'''''                        SegmentTag, _
'''''                        TMSShapeRecords.Fields("EDI_TMS_SEG_Description").Value, _
'''''                        TMSShapeRecords.Fields("EDI_TMS_Usage").Value, _
'''''                        TMSShapeRecords.Fields("EDI_TMS_Occurrence").Value, _
'''''                        TMSShapeRecords.Fields("NCTS_IEM_TMS_Usage").Value, _
'''''                        TMSShapeRecords.Fields("NCTS_IEM_TMS_Occurrence").Value
'''''
'''''
'''''            Set colDataItems = Segments.Item(strSegmentKey).SDataItems
'''''
'''''            rstNCTSItems.MoveFirst
'''''            Do Until rstNCTSItems.EOF
'''''                lngNCTS_ITM_ID = rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_ID").Value
'''''                If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_Codelist").Value) Then
'''''                    lngItemCodelistID = 0
'''''                Else
'''''                    lngItemCodelistID = rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_Codelist").Value
'''''                End If
'''''
'''''                rstEDIItems.MoveFirst
'''''                rstEDIItems.Find "EDI_TMS_ITM_ID = " & rstNCTSItems.Fields("EDI_TMS_ITM_ID").Value, , adSearchForward, 1
'''''
'''''                lngDataItemOrdinal = rstEDIItems.Fields("EDI_TMS_ITM_Ordinal").Value
'''''
'''''                If rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_DataType").Value <> "Q" Then
'''''                    strDataItemValue = rstNCTSDataItems.Fields("DATA_NCTS_" & SegmentTag & "_Seq" & lngDataItemOrdinal).Value
'''''                Else
'''''                    strDataItemValue = rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_Value").Value
'''''                End If
'''''
'''''                colDataItems.Add "S" & lngNCTS_ITM_ID, _
'''''                                        lngNCTS_ITM_ID, _
'''''                                        rstEDIItems.Fields("EDI_TMS_ITM_Description").Value, _
'''''                                        rstEDIItems.Fields("EDI_TMS_ITM_DataType").Value, _
'''''                                        rstEDIItems.Fields("EDI_TMS_ITM_Usage").Value, _
'''''                                        rstEDIItems.Fields("EDI_TMS_ITM_Ordinal").Value, _
'''''                                        rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_Description").Value, _
'''''                                        rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_DataType").Value, _
'''''                                        rstNCTSItems.Fields("NCTS_ITM_" & SegmentTag & "_Usage").Value, _
'''''                                        lngItemCodelistID, _
'''''                                        strDataItemValue
'''''
'''''                rstNCTSItems.MoveNext
'''''            Loop
'''''            rstNCTSDataItems.MoveNext
'''''        Loop
'''''    End If
End Sub

Public Function CommitToDB(Optional ByRef DataNCTS As CDataNCTS, _
                        Optional ByVal MessageID As Long = 0) _
                        As Boolean
                                    
    Dim rstTableMessages As ADODB.Recordset
    Dim rstTableSegments As ADODB.Recordset
    
    Dim rstMessageTables As ADODB.Recordset
    Dim rstMessageNCTS_IEM_TMS_IDs As ADODB.Recordset

    Dim lngDataMessageID As Long
    Dim clsSegment As Segment
    Dim clsDataItem As SDataItem
    Dim clsDataItems As SDataItems
        
    Dim lngSegmentTypeCount As Long
    Dim lngSegmentTypeCtr As Long
    Dim lngSegmentTypeInstanceCount As Long
    Dim lngSegmentTypeInstanceCtr As Long
    Dim lngDataItemOrdinal As Long
    
    Dim lngNCTS_IEM_TMS_ID As Long
    
    Dim strSegmentTag As String
    Dim strSegmentKey As String
    
    Dim blnAddNew As Boolean
    
    Dim strCommand As String
    
    CommitToDB = False
    
    ' Open Currentlt Edited Document
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "NCTS_IEM_ID, "
        strCommand = strCommand & "Data_NCTS_ID, "
        strCommand = strCommand & "DATA_NCTS_MSG_StatusType, "
        strCommand = strCommand & "User_ID, "
        strCommand = strCommand & "DATA_NCTS_MSG_Date "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "Data_NCTS_Messages "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "DATA_NCTS_MSG_ID = " & MessageID
    ADORecordsetOpen strCommand, mvarDAOConnection, rstNCTSDataMessage, adOpenKeyset, adLockOptimistic
    'Call DAOOpenTable(mvarDAOConnection, rstNCTSDataMessage, "Data_NCTS_Messages")
    With rstNCTSDataMessage
        
        '.Index = "DATA_NCTS_MSG_ID"
        '.Seek "=", MessageID
        'blnAddNew = False
        'If .NoMatch Then
        If (.EOF And .BOF) Then
            blnAddNew = True
            .AddNew
            !NCTS_IEM_ID = G_IE15_NCTS_IEM_ID
            !Data_NCTS_ID = DataNCTS.Data_NCTS_ID
            ' Add Code - if Sending then use --- EMsgStatusType_Send
            !DATA_NCTS_MSG_StatusType = GetMessageStatusType(EMsgStatusType_Document)
            !User_ID = DataNCTS.UserID
        End If
        !DATA_NCTS_MSG_Date = Now
        'lngDataMessageID = rstNCTSDataMessage.Fields("DATA_NCTS_MSG_ID").Value
'        .Update
        RecordLockingWorkAround rstNCTSDataMessage, IIf(blnAddNew, enumLockAdd, enumLockModify)

        If blnAddNew Then
            lngDataMessageID = InsertRecordset(mvarDAOConnection, rstNCTSDataMessage, "Data_NCTS_Messages")
        Else
            UpdateRecordset mvarDAOConnection, rstNCTSDataMessage, "Data_NCTS_Messages"
        End If
    End With
    
    ADORecordsetClose rstNCTSDataMessage
    'rstNCTSDataMessage.Close
    
    mvarMessageType = EMsg_IE15
    
    G_strQuery = GetQryMessageDataTables(mvarMessageType)
    ADORecordsetOpen G_strQuery, mconADOConnectionStandard, rstMessageTables, adOpenKeyset, adLockReadOnly
    
    G_strQuery = GetQryMessageTechnicalStructure(mvarMessageType)
    ADORecordsetOpen G_strQuery, mconADOConnectionStandard, rstMessageNCTS_IEM_TMS_IDs, adOpenKeyset, adLockReadOnly
    
    If rstMessageTables.RecordCount > 0 Then
        rstMessageTables.MoveFirst
        Do Until rstMessageTables.EOF
            strSegmentTag = rstMessageTables.Fields("SegmentTag").Value
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "Data_NCTS_" & strSegmentTag & " "
            ADORecordsetOpen strCommand, mvarDAOConnection, rstNCTSDataMessage, adOpenKeyset, adLockOptimistic
            'Call DAOOpenTable(mvarDAOConnection, rstNCTSDataMessage, "Data_NCTS_" & strSegmentTag)
            
'''''            ' Reims = Do only if Type is 'Document'
'''''            rstNCTSDataMessage.Index = "DATA_NCTS_MSG_ID"
'''''            rstNCTSDataMessage.Seek "=", MessageID
'''''
'''''            If Not rstNCTSDataMessage.NoMatch Then
'''''                Do While rstNCTSDataMessage.Fields("DATA_NCTS_MSG_ID").Value = MessageID
''''''                    rstNCTSDataMessage.Delete
''''''                    'RecordLockingWorkAroundForDAO rstNCTSDataMessage, enumLockDelete
'''''                    rstNCTSDataMessage.MoveNext
'''''                Loop
'''''            End If
            
            rstMessageNCTS_IEM_TMS_IDs.Filter = adFilterNone
            rstMessageNCTS_IEM_TMS_IDs.Filter = "EDI_TMS_SEG_Tag = " & Chr(39) & strSegmentTag & Chr(39)
            
            If rstMessageNCTS_IEM_TMS_IDs.RecordCount > 0 Then
                rstMessageNCTS_IEM_TMS_IDs.MoveFirst
                Do Until rstMessageNCTS_IEM_TMS_IDs.EOF
                    lngNCTS_IEM_TMS_ID = rstMessageNCTS_IEM_TMS_IDs.Fields("NCTS_IEM_TMS_ID").Value
                    
                    
                    On Error Resume Next
                    lngSegmentTypeInstanceCount = mvarSegmentCounts.Item("S" & lngNCTS_IEM_TMS_ID).Count
                    If Err.Number = 5 Then
                        Err.Clear

                    ElseIf Err.Number = 0 Then
                        On Error GoTo 0
                        
                        For lngSegmentTypeInstanceCtr = 1 To lngSegmentTypeInstanceCount
                            strSegmentKey = "S" & lngNCTS_IEM_TMS_ID & "_" & lngSegmentTypeInstanceCtr
                            Set clsDataItems = Segments.Item(strSegmentKey).SDataItems
                        
                            rstNCTSDataMessage.AddNew
                            rstNCTSDataMessage.Fields("DATA_NCTS_MSG_ID").Value = lngDataMessageID
                            rstNCTSDataMessage.Fields("NCTS_IEM_TMS_ID").Value = lngNCTS_IEM_TMS_ID
                                                                            
                            For Each clsDataItem In clsDataItems
                                lngDataItemOrdinal = clsDataItem.EDIOrdinal
                                
                                rstNCTSDataMessage.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & lngDataItemOrdinal).Value = clsDataItem.Value
                            Next
                            rstNCTSDataMessage.Update
                            
                            InsertRecordset mvarDAOConnection, rstNCTSDataMessage, "Data_NCTS_Messages"
                            
                            'RecordLockingWorkAroundForDAO rstNCTSDataMessage, enumLockAdd
                        Next
                    ElseIf Err.Number > 0 Then
                        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                    End If
                    On Error GoTo 0
                    rstMessageNCTS_IEM_TMS_IDs.MoveNext
                Loop
            End If
            rstMessageTables.MoveNext
        Loop
    End If
    rstMessageNCTS_IEM_TMS_IDs.Filter = adFilterNone
End Function

Public Sub ConnectEDIDB(ByRef DataSourceProperties As CDataSourceProperties, _
                        ByVal InitialCatalog As DBInstanceType, _
               Optional ByVal InitialCatalogYear As String = vbNullString)

    Dim strDummy As String
    If LenB(Trim$(InitialCatalogYear)) > 0 Then
        ADOConnectDB mconADOConnectionStandard, DataSourceProperties, InitialCatalog, InitialCatalogYear, , False
        ADOConnectDB mconADOConnectionShaping, DataSourceProperties, InitialCatalog, InitialCatalogYear, , True
        ADOConnectDB mvarDAOConnection, DataSourceProperties, InitialCatalog, InitialCatalogYear, , True
    Else
        ADOConnectDB mconADOConnectionStandard, DataSourceProperties, InitialCatalog, , , False
        ADOConnectDB mconADOConnectionShaping, DataSourceProperties, InitialCatalog, , , True
        '<<< dandan 112306
        '<<< Update with database password
        'Set mvarDAOConnection = OpenDatabase(mvarDBLocation & "\" & mvarDBName)
        
        ADOConnectDB mvarDAOConnection, DataSourceProperties, InitialCatalog, , , True
        'DAOConnectDB mvarDAOConnection, mvarDBLocation, mvarDBName
    End If
End Sub

Public Sub DisConnectEDIDB()
    
    ADODisconnectDB mconADOConnectionStandard
    ADODisconnectDB mconADOConnectionShaping
    ADODisconnectDB mvarDAOConnection
    
    'Set mconADOConnectionStandard = Nothing
    'Set mconADOConnectionShaping = Nothing
    'Set mvarDAOConnection = Nothing
End Sub

Private Function GetMessageHeaderProperties(ByVal IEMSource As IEMessageSources, ByVal BufferString As String) As Variant
    Dim varReturnValue As Variant
    
    Dim intFileHandle As Integer
    Dim lngUNHSegmentPosition As Long
    Dim strEDIMessage As String
    Dim strUNHSegment As String
    Dim strSeparator As String
    Dim strMsgHdrPropSegmentTag As String
    Dim strMsgHdrPropMessageIdentification  As String
    Dim strMsgHdrPropEDIMessageType As String
    Dim strMsgHdrPropMessageVersionNumber As String
    Dim strMsgHdrPropMessageReleaseNumber As String
    Dim strMsgHdrPropControllingAgency As String
    Dim strMsgHdrPropIEMessageType As String
    
    Dim clsParser As EdifactParser
    
    Select Case IEMSource
        Case IEMessageSources.IEMessageSource_Constant
            strEDIMessage = BufferString
        Case IEMessageSources.IEMessageSource_File
            '----->  Read EDI Message
            intFileHandle = FreeFile
            Open BufferString For Input As intFileHandle
            Line Input #intFileHandle, strEDIMessage
            Close intFileHandle
    End Select
    
    '----->  Get UNH Segment
    lngUNHSegmentPosition = InStr(1, strEDIMessage, "UNH")
    strUNHSegment = Mid(strEDIMessage, lngUNHSegmentPosition, InStr(lngUNHSegmentPosition, strEDIMessage, EDI_SEP_SEGMENT) - lngUNHSegmentPosition + 1)
    
    Set clsParser = New EdifactParser
    
    Call clsParser.SetParseString(IEMessageSource_Constant, strUNHSegment)
    strMsgHdrPropSegmentTag = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropMessageIdentification = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropEDIMessageType = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropMessageVersionNumber = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropMessageReleaseNumber = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropControllingAgency = clsParser.GetNextToken(0, 0)
    strSeparator = clsParser.GetNextToken(0, 0)
    strMsgHdrPropIEMessageType = clsParser.GetNextToken(0, 0)
    Set clsParser = Nothing
    varReturnValue = Array(strMsgHdrPropMessageIdentification, _
                           strMsgHdrPropEDIMessageType, _
                           strMsgHdrPropMessageVersionNumber, _
                           strMsgHdrPropMessageReleaseNumber, _
                           strMsgHdrPropControllingAgency, _
                           strMsgHdrPropIEMessageType)
    GetMessageHeaderProperties = varReturnValue
End Function

Private Function MapEDIMessageToSegmentDetails(Parser As EdifactParser, SegmentRecordset As ADODB.Recordset) As Variant
    Dim varReturnValue As Variant
    Dim strToken As String
    Dim strSeparator As String
    Dim blnGetNextToken As Boolean
    Dim lngPreviousGroupID As Long
    Dim astrTokens() As String
    Dim lngArrayIndex As Long
    
    strSeparator = Parser.GetNextToken(0, 0)
    Debug.Assert strSeparator = EDI_SEP_DATA_ELEMENT
    If strSeparator = EDI_SEP_DATA_ELEMENT Then
        Debug.Assert SegmentRecordset.RecordCount > 0
        If SegmentRecordset.RecordCount > 0 Then
            ReDim astrTokens(SegmentRecordset.RecordCount - 1)
            
            lngPreviousGroupID = 0
            lngArrayIndex = 0
            SegmentRecordset.MoveFirst
            Do While Not SegmentRecordset.EOF
                blnGetNextToken = False
                strToken = vbNullString
                If SegmentRecordset!EDI_TMS_GRP_ID = 0 Then
                    Debug.Assert (strSeparator = EDI_SEP_DATA_ELEMENT) Or (strSeparator = EDI_SEP_SEGMENT)
                    blnGetNextToken = (strSeparator = EDI_SEP_DATA_ELEMENT)
                Else
                    If lngPreviousGroupID = SegmentRecordset!EDI_TMS_GRP_ID Then
                        blnGetNextToken = (strSeparator = EDI_SEP_COMPOSITE_DATA_ELEMENT)
                    Else
                        blnGetNextToken = (strSeparator = EDI_SEP_DATA_ELEMENT)
                    End If
                End If
                If blnGetNextToken Then
                    strToken = Parser.GetNextToken(0, 0)
                    If IsEDISeparator(strToken) Then
                        strSeparator = strToken
                        strToken = vbNullString
                    Else
                        strSeparator = Parser.GetNextToken(0, 0)
                    End If
                End If
                lngPreviousGroupID = SegmentRecordset!EDI_TMS_GRP_ID
                astrTokens(lngArrayIndex) = strToken
                lngArrayIndex = lngArrayIndex + 1
                SegmentRecordset.MoveNext
            Loop
        End If
    End If
    varReturnValue = astrTokens
    MapEDIMessageToSegmentDetails = varReturnValue
End Function

Private Function ParentExists(ByVal SearchParentID As Long) As Boolean
    Dim blnReturnValue As Boolean
    Dim lngSegmentIndex As Long
    Dim blnFoundSegment As Boolean
    Dim blnFinishedSearch As Boolean
    Dim strSegmentKey As String
    
    lngSegmentIndex = 0
    blnFinishedSearch = False
    Do While Not blnFinishedSearch
        lngSegmentIndex = lngSegmentIndex + 1
        strSegmentKey = Segments(lngSegmentIndex).Key
        blnFoundSegment = (CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID)) = SearchParentID) '(CLng(Mid(strSegmentKey, 2, InStr(2, strSegmentKey, "_") - 2)) = SearchParentID)
        blnFinishedSearch = blnFoundSegment Or (lngSegmentIndex = Segments.Count)
    Loop
    blnReturnValue = blnFoundSegment
    ParentExists = blnReturnValue
End Function

Public Function GetSegmentIndex(ByVal SegmentKey As String) As Long
    Dim lngReturnValue As Long
    Dim lngSegmentIndex As Long
    
    Dim blnFoundSegment As Boolean
    Dim blnFinishedSearch As Boolean
    
    lngReturnValue = 0
    lngSegmentIndex = 0
    blnFinishedSearch = (Segments.Count = 0)
    Do While Not blnFinishedSearch
        lngSegmentIndex = lngSegmentIndex + 1
        blnFoundSegment = (Segments(lngSegmentIndex).Key = SegmentKey)
        blnFinishedSearch = (blnFoundSegment Or (lngSegmentIndex = Segments.Count))
    Loop
    If blnFoundSegment Then
        lngReturnValue = lngSegmentIndex
    Else
        lngReturnValue = 0
    End If
    GetSegmentIndex = lngReturnValue
End Function

Private Function FoundSegment(ByVal SegmentTag As String, ByRef TMSRecordset As ADODB.Recordset, ByRef SegmentDetailsValues As Variant, ByVal SearchDetailsOnly As Boolean) As Boolean
    Dim blnReturnValue                  As Boolean
    Dim blnFoundSegment                 As Boolean
    Dim blnIsQualifierSatisfied         As Boolean
    Dim blnSearchedFromStart            As Boolean
    Dim blnSearchedEntireRecordset      As Boolean
    Dim blnInDetailsSectionAlready      As Boolean
    
    Dim lngEDISegmentItemsIndex         As Long
    
    Dim rstEDISegmentItems              As ADODB.Recordset
    Dim rstNCTSSegmentItems             As ADODB.Recordset
    
    blnFoundSegment = False
    blnSearchedFromStart = False
    blnSearchedEntireRecordset = False
    blnIsQualifierSatisfied = False
    blnInDetailsSectionAlready = False
    TMSRecordset.MoveFirst
    Do
        If TMSRecordset.Fields("EDI_TMS_SEG_Tag").Value = "UNS" Then
            Set rstNCTSSegmentItems = TMSRecordset.Fields("UNS").Value
            blnInDetailsSectionAlready = (blnInDetailsSectionAlready Or rstNCTSSegmentItems.Fields("NCTS_ITM_UNS_Value").Value = "D")
        End If
        
        If ((TMSRecordset.Fields("EDI_TMS_SEG_Tag").Value = SegmentTag) And Not SearchDetailsOnly) Or (SearchDetailsOnly And blnInDetailsSectionAlready) Then
            Set rstEDISegmentItems = TMSRecordset.Fields("EDI_TMS_ITEMS").Value
            Set rstNCTSSegmentItems = TMSRecordset.Fields(SegmentTag).Value
            lngEDISegmentItemsIndex = 0
            blnIsQualifierSatisfied = (rstNCTSSegmentItems.RecordCount > 0)
            Do While Not rstEDISegmentItems.EOF And blnIsQualifierSatisfied
                If IsInRecordset(rstNCTSSegmentItems, "EDI_TMS_ITM_ID = " & rstEDISegmentItems.Fields("EDI_TMS_ITM_ID").Value) Then
                    If rstNCTSSegmentItems.Fields(PREFIX_NCTS_ITEMS_TABLE & "_" & SegmentTag & "_DataType").Value = EDI_DATATYPE_QUALIFIER Then
                        blnIsQualifierSatisfied = blnIsQualifierSatisfied And (rstNCTSSegmentItems.Fields(PREFIX_NCTS_ITEMS_TABLE & "_" & SegmentTag & "_Value").Value = SegmentDetailsValues(lngEDISegmentItemsIndex))
                    End If
                End If
                lngEDISegmentItemsIndex = lngEDISegmentItemsIndex + 1
                rstEDISegmentItems.MoveNext
            Loop
            If TMSRecordset.Fields("NCTS_IEM_TMS_ParentID").Value Then
                If blnIsQualifierSatisfied Then
                    blnIsQualifierSatisfied = ParentExists(TMSRecordset.Fields("NCTS_IEM_TMS_ParentID").Value)
                End If
            End If
        End If
        blnFoundSegment = blnIsQualifierSatisfied
        If Not blnFoundSegment Then
            TMSRecordset.MoveNext
            If TMSRecordset.EOF Then
                TMSRecordset.MoveFirst
                
                If blnSearchedFromStart Then
                    blnSearchedEntireRecordset = True
                Else
                    blnSearchedFromStart = True
                End If
            End If
        End If
    Loop Until blnFoundSegment Or blnSearchedEntireRecordset
    blnReturnValue = blnFoundSegment
    FoundSegment = blnReturnValue
End Function

'Rachelle - for IE34 - 091205 : added optional variable blnISIE34
Public Sub PrepareMessageFromDatabase(ByVal DataNCTSMessageID As Long, ByVal UniqueCode As String, Optional ByVal blnISIE34 As Boolean)
    Dim SQLTechnicalMessageStructure As String
    Dim SQLDataNCTSMessage As String
    
    Dim rstDataNCTSMessage As ADODB.Recordset
    Dim rstDataNCTSSegment As ADODB.Recordset
    Dim rstDataNCTSMessageParent As ADODB.Recordset
    Dim rstDataNCTSSegmentParent As ADODB.Recordset
    
    Dim strSegmentTag As String
    Dim lngDataIndex As Long
    
    On Error GoTo ErrHandler
    
10  Do While Segments.Count > 0
15      Segments.Remove Segments.Count
    Loop
    
20  Do While SegmentCounts.Count > 0
25      SegmentCounts.Remove SegmentCounts.Count
    Loop
    
30  mstrUniqueCode = UniqueCode
    
35  SQLDataNCTSMessage = "SHAPE {" & GetQryNCTSDataMessage(DataNCTSMessageID) & "} AS DATA_NCTS_MESSAGES "
40  SQLDataNCTSMessage = SQLDataNCTSMessage & "APPEND "
    
    'Rachelle - for IE34 - 091205
    If blnISIE34 = False Then
45      SQLDataNCTSMessage = SQLDataNCTSMessage & PrepareDataNCTSSegments(EMsg_IE15, DataNCTSMessageID)
    Else
        SQLDataNCTSMessage = SQLDataNCTSMessage & PrepareDataNCTSSegments(EMsg_IE34, DataNCTSMessageID)
    End If

50  ADORecordsetOpen SQLDataNCTSMessage, mconADOConnectionShaping, mrstDataNCTSMessage, adOpenKeyset, adLockReadOnly, , True
    
55  Set rstDataNCTSMessage = mrstDataNCTSMessage.Clone(adLockReadOnly)
60  Set rstDataNCTSMessageParent = mrstDataNCTSMessage.Clone(adLockReadOnly)
    
65  mvarMessageType = GetMessageTypeFromNCTS_IEM_ID(mrstDataNCTSMessage.Fields("NCTS_IEM_ID").Value)
70  mvarNCTS_IEM_ID = rstDataNCTSMessage.Fields("NCTS_IEM_ID").Value
    
75  SQLTechnicalMessageStructure = "SHAPE {" & GetQryMessageTechnicalStructure(mvarMessageType) & "} " & _
                                   "APPEND " & _
                                   PrepareMessageDataItems(mvarMessageType) & _
                                   "({SELECT * FROM EDI_TMS_ITEMS ORDER BY EDI_TMS_ITM_Ordinal} AS EDI_TMS_ITEMS RELATE EDI_TMS_SEG_ID TO EDI_TMS_SEG_ID) "
    
80  ADORecordsetOpen SQLTechnicalMessageStructure, mconADOConnectionShaping, mrstTechnicalMessageStructure, adOpenKeyset, adLockReadOnly, , True
    
85  Set mvarSegments.TechnicalMessageStructure = mrstTechnicalMessageStructure
    
90  mrstTechnicalMessageStructure.MoveFirst
    
95  Do While Not mrstTechnicalMessageStructure.EOF
100     strSegmentTag = mrstTechnicalMessageStructure.Fields("EDI_TMS_SEG_Tag").Value
105     Set rstDataNCTSSegment = rstDataNCTSMessage.Fields("DATA_NCTS_" & strSegmentTag).Value
        
110     If rstDataNCTSSegment.RecordCount > 0 Then
115         rstDataNCTSSegment.MoveFirst
            
120         Do While Not rstDataNCTSSegment.EOF
125             If rstDataNCTSSegment.Fields("NCTS_IEM_TMS_ID").Value = mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value Then
130                 If Not IsNull(rstDataNCTSSegment.Fields("PARENT_SEGMENT_TAG").Value) Then
135                     Set rstDataNCTSSegmentParent = rstDataNCTSMessageParent.Fields("DATA_NCTS_" & rstDataNCTSSegment.Fields("PARENT_SEGMENT_TAG").Value).Value
                        
140                     rstDataNCTSSegmentParent.MoveFirst
145                     rstDataNCTSSegmentParent.Find "DATA_NCTS_" & rstDataNCTSSegment.Fields("PARENT_SEGMENT_TAG").Value & "_ID = " & rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_ParentID").Value
                        
                        ' Checking whether a record has been found or not avoids run-time error 3021 for ADODB.Field:
                        ' [Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.]
150                     If Not rstDataNCTSSegmentParent.EOF Then
155                         Call AddEmptySegment(mrstTechnicalMessageStructure, mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value, rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_Instance").Value, _
                                                 rstDataNCTSSegmentParent.Fields("DATA_NCTS_" & rstDataNCTSSegment.Fields("PARENT_SEGMENT_TAG").Value & "_Instance").Value)
                        Else
160                         Call AddEmptySegment(mrstTechnicalMessageStructure, mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value, rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_Instance").Value, 1)
                        End If
                    Else
165                     Call AddEmptySegment(mrstTechnicalMessageStructure, mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value, rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_Instance").Value, 1)
                    End If
                    
170                 For lngDataIndex = 1 To Segments(Segments.Count).SDataItems.Count
175                     If Not IsNull(rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value) Then
180                         Segments(Segments.Count).SDataItems(lngDataIndex).Value = rstDataNCTSSegment.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value
                        Else
185                         Segments(Segments.Count).SDataItems(lngDataIndex).Value = vbNullString
                        End If
                    Next
                End If
                
190             rstDataNCTSSegment.MoveNext
            Loop
        End If
        
195     mrstTechnicalMessageStructure.MoveNext
    Loop
    
200 UpdateChildrenKeys
    
205 ADORecordsetClose rstDataNCTSMessage
210 ADORecordsetClose rstDataNCTSSegment
215 ADORecordsetClose rstDataNCTSMessageParent
220 ADORecordsetClose rstDataNCTSSegmentParent
    
    Exit Sub
    
ErrHandler:
    
    Err.Raise Err.Number, "prjEdifactDll.EdifactMessage.PrepareMessageFromDatabase", Err.Source & " (" & Erl & "): " & Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub PrepareMessageFromIEMessage(ByVal IEMSource As IEMessageSources, ByVal BufferString As String)
    Dim varMessageHeaderProperties As Variant
    Dim enuMessageType As ENCTSMessageType
    Dim clsEdifactParser As EdifactParser
    Dim strSegmentTag As String
    Dim blnSearchDetailsOnly As Boolean
    Dim varSegmentDetailValues As Variant
    
    Dim lngDataItemsIndex As Long
    
    Dim lngSegmentInstance As Long
    Dim lngSegmentInstanceParent As Long
    
    On Error GoTo ErrHandler
    
10  Do While Segments.Count > 0
15      Segments.Remove Segments.Count
    Loop
    
20  Do While SegmentCounts.Count > 0
25      SegmentCounts.Remove SegmentCounts.Count
    Loop
    
30  BufferString = StripNullTerminator(BufferString)
    
35  varMessageHeaderProperties = GetMessageHeaderProperties(IEMSource, BufferString)
40  enuMessageType = GetMessageType(varMessageHeaderProperties(G_IE_MESSAGE_TYPE_ARRAY_INDEX))
45  mvarMessageType = enuMessageType
    
50  G_strQuery = "SHAPE {" & GetQryMessageTechnicalStructure(enuMessageType) & "} " & _
                 "APPEND " & _
                 PrepareMessageDataItems(enuMessageType) & _
                 "({SELECT * FROM EDI_TMS_ITEMS ORDER BY EDI_TMS_ITM_Ordinal} AS EDI_TMS_ITEMS RELATE EDI_TMS_SEG_ID TO EDI_TMS_SEG_ID) "
    
55  ADORecordsetOpen G_strQuery, mconADOConnectionShaping, mrstTechnicalMessageStructure, adOpenKeyset, adLockReadOnly, , True
    
60  Set mvarSegments.TechnicalMessageStructure = mrstTechnicalMessageStructure
    
65  Set clsEdifactParser = New EdifactParser
    
70  mvarNCTS_IEM_ID = mrstTechnicalMessageStructure.Fields("NCTS_IEM_ID").Value
    
75  clsEdifactParser.SetParseString IEMSource, BufferString
    
80  mvarEDIMessage = BufferString
    
85  blnSearchDetailsOnly = False
    
    Do
90      strSegmentTag = clsEdifactParser.GetNextToken(0, 0)
        AddToTrace "CubeLibEdifact.EdifactMessage.PrepareMessageFromIEMessage(EDI_TMS_SEG_Tag: " & strSegmentTag & ")"
95      If IsInRecordset(mrstTechnicalMessageStructure, "EDI_TMS_SEG_Tag = '" & strSegmentTag & "'") Then
100         varSegmentDetailValues = MapEDIMessageToSegmentDetails(clsEdifactParser, mrstTechnicalMessageStructure.Fields("EDI_TMS_ITEMS").Value)
            
105         If FoundSegment(strSegmentTag, mrstTechnicalMessageStructure, varSegmentDetailValues, blnSearchDetailsOnly) Then
110             If IsNull(mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ParentID").Value) Then
115                 lngSegmentInstanceParent = 1
                Else
120                 If mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ParentID").Value = 0 Then
125                     lngSegmentInstanceParent = 1
                    Else
130                     lngSegmentInstanceParent = SegmentCounts("S" & mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ParentID").Value).Count
                    End If
                End If
                
135             SegmentCounts.Add "S" & mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value
                
140             lngSegmentInstance = SegmentCounts("S" & mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value).Count
                
145             Segments.Add mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value, lngSegmentInstance, lngSegmentInstanceParent
                
150             If IsNull(mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value) Then
155                 lngSegmentInstanceParent = 1
                Else
160                 If mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value = 0 Then
165                     lngSegmentInstanceParent = 1
                    Else
170                     lngSegmentInstanceParent = SegmentCounts("S" & mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ID").Value).Count
                    End If
                End If
                
175             For lngDataItemsIndex = 1 To Segments(Segments.Count).SDataItems.Count
180                 Segments(Segments.Count).SDataItems(lngDataItemsIndex).Value = varSegmentDetailValues(lngDataItemsIndex - 1)
                Next
            End If
            
185         blnSearchDetailsOnly = (blnSearchDetailsOnly Or ((strSegmentTag = "UNS") And (varSegmentDetailValues(0) = "D")))
        End If
190 Loop Until strSegmentTag = vbNullString
    
195 UpdateChildrenKeys
    
200 Set clsEdifactParser = Nothing
    
    Exit Sub
    
ErrHandler:
    AddToTrace "Error in CubeLibEdifact.EdifactMessage.PrepareMessageFromIEMessage(IEMSource: " & IEMSource & " BufferString: " & BufferString & ")"
    Err.Raise Err.Number, "prjEdifactDll.EdifactMessage.PrepareMessageFromIEMessage", Err.Source & " (" & Erl & "): " & Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Function AllSegmentsAreSavedToDatabase() As Boolean
    Dim blnReturnValue As Boolean
    Dim lngSegmentIndex As Long
    
    blnReturnValue = True
    For lngSegmentIndex = 1 To Segments.Count
        blnReturnValue = blnReturnValue And (Segments(lngSegmentIndex).DataNCTSSegmentID > 0)
    Next
    
    AllSegmentsAreSavedToDatabase = blnReturnValue
End Function

Private Function CommitToDatabase(ByVal DataNCTSID As Long, ByVal MessageStatus As String, ByVal UserID As Long, Optional ByVal ErrorReply As String) As Long
    Dim lngReturnValue As Long
    
    Dim rstDataNCTSMessage              As ADODB.Recordset
    Dim SQLDataNCTSMessage              As String
    Dim rstSegments                     As ADODB.Recordset
    Dim SQLSegments                     As String
    
    Dim strSegmentTag                   As String
    Dim strSegmentKey                   As String
    
    Dim lngSegmentIndex                 As Long
    Dim lngDataIndex                    As Long
    Dim lngDataNCTSMessageID            As Long
    Dim lngDataNCTSSegmentID            As Long
    Dim lngDataNCTSMessageIDOld         As Long
    Dim lngParentSegmentID              As Long
    
    Dim blnParentIsAlreadySaved         As Boolean
    Dim blnSegmentHasNoParent           As Boolean
    Dim blnReplaceOldRecords            As Boolean
    Dim blnAddNewRecord                 As Boolean
    
    Dim clsSegment                      As Segment
    Dim strCommand As String
    
    Call AddMissingRequiredSegments
    
    SQLDataNCTSMessage = "SELECT * FROM DATA_NCTS_MESSAGES"
    
    If MessageStatus = MESSAGE_STATUS_SENT Then
        '----->  If message status = "Sent" then find the queued document
        SQLDataNCTSMessage = SQLDataNCTSMessage & " WHERE DATA_NCTS_ID = " & CStr(DataNCTSID) & " AND DATA_NCTS_MSG_StatusType = '" & MESSAGE_STATUS_QUEUED & "'"
    Else
        SQLDataNCTSMessage = SQLDataNCTSMessage & " WHERE DATA_NCTS_ID = " & CStr(DataNCTSID) & " AND DATA_NCTS_MSG_StatusType = '" & MessageStatus & "'"
    End If
    
    ADORecordsetOpen SQLDataNCTSMessage, mconADOConnectionStandard, rstDataNCTSMessage, adOpenKeyset, adLockOptimistic
    
    If MessageStatus = MESSAGE_STATUS_DOCUMENT Or MessageStatus = MESSAGE_STATUS_QUEUED Then
        '----->  Maximum number of records found must be 1
        Debug.Assert rstDataNCTSMessage.RecordCount = 0 Or rstDataNCTSMessage.RecordCount = 1
        
        blnReplaceOldRecords = (rstDataNCTSMessage.RecordCount > 0)
        
        If blnReplaceOldRecords Then
            '----->  Save the DATA_NCTS_MSG_ID because it will be used as a reference when
            '----->  deleting old records in the DATA_NCTS_XXX tables
            lngDataNCTSMessageIDOld = rstDataNCTSMessage.Fields("DATA_NCTS_MSG_ID").Value
        Else
            lngDataNCTSMessageIDOld = 0
        End If
        
        blnAddNewRecord = True
    ElseIf MessageStatus = MESSAGE_STATUS_SENT Then
        '----->  Exactly ONE record must exist
        Debug.Assert rstDataNCTSMessage.RecordCount = 1
        
        blnAddNewRecord = False
    ElseIf MessageStatus = MESSAGE_STATUS_RECEIVED Then
        blnReplaceOldRecords = False
        blnAddNewRecord = True
    Else
        '----->  Unidentified message status
        Debug.Assert False
    End If
    
    If blnReplaceOldRecords Then
            strCommand = vbNullString
            strCommand = strCommand & "DELETE "
            strCommand = strCommand & "* "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "DATA_NCTS_MESSAGES "
            strCommand = strCommand & "WHERE "
            strCommand = strCommand & "DATA_NCTS_ID = " & CStr(DataNCTSID) & " "
            strCommand = strCommand & "AND "
            If MessageStatus = MESSAGE_STATUS_SENT Then
                strCommand = strCommand & "DATA_NCTS_MSG_StatusType = '" & MESSAGE_STATUS_QUEUED & "'"
            Else
                strCommand = strCommand & "DATA_NCTS_MSG_StatusType = '" & MessageStatus & "'"
            End If
        ExecuteNonQuery mconADOConnectionStandard, strCommand
        
'        rstDataNCTSMessage.Delete
        RecordLockingWorkAround rstDataNCTSMessage, enumLockDelete
    End If
    
    'If rstDataNCTSMessage.RecordCount = 0 Or MessageStatus = MESSAGE_STATUS_RECEIVED Then
    If blnAddNewRecord Then
        rstDataNCTSMessage.AddNew
    End If
    
    'lngDataNCTSMessageID = rstDataNCTSMessage.Fields("DATA_NCTS_MSG_ID").Value
    'lngReturnValue = lngDataNCTSMessageID
    
    If Len(ErrorReply) Then
        rstDataNCTSMessage.Fields("NCTS_IEM_ID").Value = NCTS_IEM_ID_CODEM
        rstDataNCTSMessage.Fields("DATA_NCTS_MSG_Message").Value = ErrorReply
    Else
        rstDataNCTSMessage.Fields("NCTS_IEM_ID").Value = mvarNCTS_IEM_ID
        rstDataNCTSMessage.Fields("DATA_NCTS_MSG_Message").Value = mvarEDIMessage
    End If
    
    rstDataNCTSMessage.Fields("DATA_NCTS_MSG_Reference").Value = ReferenceNumber
    rstDataNCTSMessage.Fields("DATA_NCTS_ID").Value = DataNCTSID
    rstDataNCTSMessage.Fields("DATA_NCTS_MSG_Date").Value = Now
    rstDataNCTSMessage.Fields("DATA_NCTS_MSG_StatusType").Value = MessageStatus
    rstDataNCTSMessage.Fields("User_ID").Value = UserID
    
'    rstDataNCTSMessage.Update
    RecordLockingWorkAround rstDataNCTSMessage, enumLockAdd

    lngDataNCTSMessageID = InsertRecordset(mconADOConnectionStandard, rstDataNCTSMessage, "DATA_NCTS_MESSAGES")
    lngReturnValue = lngDataNCTSMessageID
    
    ADORecordsetClose rstDataNCTSMessage
    
    If Len(ErrorReply) = 0 Then
        '----->  Reset DATA_NCTS_XXX IDs
        For Each clsSegment In Segments
            clsSegment.DataNCTSSegmentID = 0
        Next
        
        Do
            For lngSegmentIndex = 1 To Segments.Count
                If (Segments(lngSegmentIndex).DataNCTSSegmentID = 0) Then
                    '----->  Get necessary segment properties
                    strSegmentTag = Segments(lngSegmentIndex).SegmentTag
                    strSegmentKey = Segments(lngSegmentIndex).Key
                    
                    '----->  Check if segment has a parent
                    blnSegmentHasNoParent = (Segments(lngSegmentIndex).KeyParent = vbNullString)
                    
                    If Not blnSegmentHasNoParent Then
                        lngParentSegmentID = GetSegmentIndex(Segments(lngSegmentIndex).KeyParent)
                        
                        Debug.Assert lngParentSegmentID > 0
                        
                        If lngParentSegmentID > 0 Then
                            blnParentIsAlreadySaved = (Segments(lngParentSegmentID).DataNCTSSegmentID > 0)
                        End If
                    End If
                    
                    If blnParentIsAlreadySaved Or blnSegmentHasNoParent Then
                        SQLSegments = "SELECT * " & _
                                      "FROM DATA_NCTS_" & strSegmentTag & " " & _
                                      "WHERE DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = "
                        
                        If blnReplaceOldRecords Then
                            ADORecordsetOpen SQLSegments & lngDataNCTSMessageIDOld, mconADOConnectionStandard, rstSegments, adOpenKeyset, adLockOptimistic
                            
                            Do While rstSegments.RecordCount > 0
'                                rstSegments.Delete
                                RecordLockingWorkAround rstSegments, enumLockDelete

                                    strCommand = vbNullString
                                    strCommand = strCommand & "DELETE "
                                    strCommand = strCommand & "* "
                                    strCommand = strCommand & "FROM "
                                    strCommand = strCommand & "DATA_NCTS_" & strSegmentTag & " "
                                    strCommand = strCommand & "WHERE "
                                    strCommand = strCommand & "DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = " & lngDataNCTSMessageIDOld
                                ExecuteNonQuery mconADOConnectionStandard, strCommand
                                
                                rstSegments.MoveNext
                            Loop
                        End If
                        
                        '----->  Open corresponding DATA_NCTS_XXX table
                        ADORecordsetOpen SQLSegments & lngDataNCTSMessageID, mconADOConnectionStandard, rstSegments, adOpenKeyset, adLockOptimistic
                        
                        '----->  Save segment to database
                        rstSegments.AddNew
                        '----->  Get ID of new segment (to be saved to class for reference)
                        lngDataNCTSSegmentID = rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ID").Value
                        
                        '----->  Save necessary data to database
                        rstSegments.Fields("DATA_NCTS_MSG_ID").Value = lngDataNCTSMessageID
                        rstSegments.Fields("NCTS_IEM_TMS_ID").Value = CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID))
                        
                        If blnSegmentHasNoParent Then
                            rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ParentID").Value = 0
                        ElseIf blnParentIsAlreadySaved Then
                            rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ParentID").Value = Segments(lngParentSegmentID).DataNCTSSegmentID
                        Else
                            '----->  Segment's parent must be saved to database already or it has no parent at all
                            Debug.Assert False
                        End If
                        
                        rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Instance").Value = CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                        
                        For lngDataIndex = 1 To Segments(lngSegmentIndex).SDataItems.Count
                            If Segments(lngSegmentIndex).SDataItems(lngDataIndex).Value <> vbNullString Then
                                rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value = Segments(lngSegmentIndex).SDataItems(lngDataIndex).Value
                            Else
                                rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value = Null
                            End If
                        Next
                        
'                        rstSegments.Update
                        RecordLockingWorkAround rstSegments, enumLockAdd
                        
                        lngDataNCTSSegmentID = InsertRecordset(mconADOConnectionStandard, rstSegments, "DATA_NCTS_" & strSegmentTag)
                        
                        Segments(lngSegmentIndex).DataNCTSSegmentID = lngDataNCTSSegmentID
                        
                        ADORecordsetClose rstSegments
                        
                        
                    End If
                End If
            Next
        Loop Until AllSegmentsAreSavedToDatabase
    End If
    
    CommitToDatabase = lngReturnValue
End Function

Public Sub CommitToOtherDatabase(ByVal DataNCTSID As Long, ByVal lngNewDataNCTSID As Long, ByVal DataNCTSMessageID As Long, ByRef conEDIHistory As ADODB.Connection)
        
    Dim rstDataNCTSMessage              As ADODB.Recordset
    Dim rstHistoryDataNCTSMessage       As ADODB.Recordset
    Dim SQLDataNCTSMessage              As String
    Dim rstSegments                     As ADODB.Recordset
    Dim SQLSegments                     As String
    
    Dim strSegmentTag                   As String
    Dim strSegmentKey                   As String
    Dim strFieldName                    As String
    
    Dim lngSegmentIndex                 As Long
    Dim lngDataIndex                    As Long
    Dim lngDataNCTSMessageID            As Long
    Dim lngDataNCTSSegmentID            As Long
    Dim lngDataNCTSMessageIDOld         As Long
    Dim lngParentSegmentID              As Long
    Dim lngCtr                          As Long
    
    Dim blnParentIsAlreadySaved         As Boolean
    Dim blnSegmentHasNoParent           As Boolean
    Dim blnReplaceOldRecords            As Boolean
    Dim blnAddNewRecord                 As Boolean
    
    Dim clsSegment                      As Segment
    
    Call AddMissingRequiredSegments
    
    Dim strCommand As String
    
    SQLDataNCTSMessage = "SELECT * FROM DATA_NCTS_MESSAGES"
    SQLDataNCTSMessage = SQLDataNCTSMessage & " WHERE DATA_NCTS_ID = " & CStr(DataNCTSID) & " AND Data_NCTS_MSG_ID = " & DataNCTSMessageID
        
    ADORecordsetOpen SQLDataNCTSMessage, mconADOConnectionStandard, rstDataNCTSMessage, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen SQLDataNCTSMessage, conEDIHistory, rstHistoryDataNCTSMessage, adOpenKeyset, adLockOptimistic
        
    lngDataNCTSMessageIDOld = rstDataNCTSMessage.Fields("DATA_NCTS_MSG_ID").Value
    blnAddNewRecord = True
    blnReplaceOldRecords = True
    
    rstHistoryDataNCTSMessage.AddNew
    
    lngDataNCTSMessageID = rstHistoryDataNCTSMessage.Fields("DATA_NCTS_MSG_ID").Value
    
    For lngCtr = 0 To rstHistoryDataNCTSMessage.Fields.Count - 1
        strFieldName = rstHistoryDataNCTSMessage.Fields(lngCtr).Name
        rstHistoryDataNCTSMessage.Fields(strFieldName).Value = rstDataNCTSMessage.Fields(strFieldName).Value
    Next
    rstHistoryDataNCTSMessage.Fields("DATA_NCTS_ID").Value = lngNewDataNCTSID
    rstHistoryDataNCTSMessage.Fields("DATA_NCTS_MSG_ID").Value = lngDataNCTSMessageID
    
    rstDataNCTSMessage.Delete
    rstDataNCTSMessage.Update
    
        strCommand = vbNullString
        strCommand = strCommand & "DELETE "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "DATA_NCTS_MESSAGES "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "DATA_NCTS_ID = " & CStr(DataNCTSID) & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "Data_NCTS_MSG_ID = " & DataNCTSMessageID
    ExecuteNonQuery mconADOConnectionStandard, strCommand
    
    rstHistoryDataNCTSMessage.Update
    
    InsertRecordset conEDIHistory, rstHistoryDataNCTSMessage, "DATA_NCTS_MESSAGES"
    
    ADORecordsetClose rstDataNCTSMessage
    ADORecordsetClose rstHistoryDataNCTSMessage
    
    '----->  Reset DATA_NCTS_XXX IDs
    For Each clsSegment In Segments
        clsSegment.DataNCTSSegmentID = 0
    Next
    
    Do
        For lngSegmentIndex = 1 To Segments.Count
            If (Segments(lngSegmentIndex).DataNCTSSegmentID = 0) Then
                '----->  Get necessary segment properties
                strSegmentTag = Segments(lngSegmentIndex).SegmentTag
                strSegmentKey = Segments(lngSegmentIndex).Key
                
                '----->  Check if segment has a parent
                blnSegmentHasNoParent = (Segments(lngSegmentIndex).KeyParent = vbNullString)
                
                If Not blnSegmentHasNoParent Then
                    lngParentSegmentID = GetSegmentIndex(Segments(lngSegmentIndex).KeyParent)
                    
                    Debug.Assert lngParentSegmentID > 0
                    
                    If lngParentSegmentID > 0 Then
                        blnParentIsAlreadySaved = (Segments(lngParentSegmentID).DataNCTSSegmentID > 0)
                    End If
                End If
                
                If blnParentIsAlreadySaved Or blnSegmentHasNoParent Then
                    SQLSegments = "SELECT * " & _
                                  "FROM DATA_NCTS_" & strSegmentTag & " " & _
                                  "WHERE DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = "
                    
                    If blnReplaceOldRecords Then
                        ADORecordsetOpen SQLSegments & lngDataNCTSMessageIDOld, mconADOConnectionStandard, rstSegments, adOpenKeyset, adLockOptimistic
                        
                        Do While rstSegments.RecordCount > 0
                            
                            rstSegments.Delete

                            rstSegments.MoveNext
                        Loop
                        
                            strCommand = vbNullString
                            strCommand = strCommand & "DELETE "
                            strCommand = strCommand & "* "
                            strCommand = strCommand & "FROM "
                            strCommand = strCommand & "DATA_NCTS_" & strSegmentTag & " "
                            strCommand = strCommand & "WHERE "
                            strCommand = strCommand & "DATA_NCTS_" & strSegmentTag & ".DATA_NCTS_MSG_ID = " & lngDataNCTSMessageIDOld
                        ExecuteNonQuery mconADOConnectionStandard, strCommand
                    End If
                    
                    '----->  Open corresponding DATA_NCTS_XXX table
                    ADORecordsetOpen SQLSegments & lngDataNCTSMessageID, conEDIHistory, rstSegments, adOpenKeyset, adLockOptimistic
                    '----->  Save segment to database
                    rstSegments.AddNew
                    '----->  Get ID of new segment (to be saved to class for reference)
                    lngDataNCTSSegmentID = rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ID").Value
                    
                    '----->  Save necessary data to database
                    rstSegments.Fields("DATA_NCTS_MSG_ID").Value = lngDataNCTSMessageID
                    rstSegments.Fields("NCTS_IEM_TMS_ID").Value = CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_NCTS_IEM_TMS_ID))
                    
                    If blnSegmentHasNoParent Then
                        rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ParentID").Value = 0
                    ElseIf blnParentIsAlreadySaved Then
                        rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_ParentID").Value = Segments(lngParentSegmentID).DataNCTSSegmentID
                    Else
                        '----->  Segment's parent must be saved to database already or it has no parent at all
                        Debug.Assert False
                    End If
                    
                    rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Instance").Value = CLng(Split(strSegmentKey, "_")(SegmentKeyIndexes.SegmentKeyIndex_SegmentInstance))
                    
                    For lngDataIndex = 1 To Segments(lngSegmentIndex).SDataItems.Count
                        If Segments(lngSegmentIndex).SDataItems(lngDataIndex).Value <> vbNullString Then
                            rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value = Segments(lngSegmentIndex).SDataItems(lngDataIndex).Value
                        Else
                            rstSegments.Fields("DATA_NCTS_" & strSegmentTag & "_Seq" & CStr(lngDataIndex)).Value = Null
                        End If
                    Next
                    
                    rstSegments.Update
                    
                    lngDataNCTSSegmentID = InsertRecordset(conEDIHistory, rstSegments, "DATA_NCTS_" & strSegmentTag)
                    
                    Segments(lngSegmentIndex).DataNCTSSegmentID = lngDataNCTSSegmentID
                    
                    ADORecordsetClose rstSegments
                End If
            End If
        Next
    Loop Until AllSegmentsAreSavedToDatabase
    
End Sub

Public Sub UpdateChildrenKeys()
    Dim clsSegment As Segment
    
    For Each clsSegment In Segments
        '----->  Does segment have a parent?
        If clsSegment.KeyParent <> vbNullString Then
            '----->  YES
            If Trim(Segments(clsSegment.KeyParent).KeysChildren) = vbNullString Then
                Segments(clsSegment.KeyParent).KeysChildren = clsSegment.Key
            Else
                Segments(clsSegment.KeyParent).KeysChildren = Segments(clsSegment.KeyParent).KeysChildren & "|" & clsSegment.Key
            End If
        End If
    Next
End Sub

Private Function ProcessValue(ByVal strValue As String, ByVal strDataType As String) As String
Dim strTemp As String
'requirements: 1. strip leading zeroes
'              2. decimal separator should be "."

    '===== for stripping leading zeroes ============='
    If UCase(Left(strDataType, 3)) = "N.." Then
        'If InStr(3, strDataType, ",") > 0 And IsNumeric(strValue) Then 'with decimal?
        If IsNumeric(strValue) Then 'with decimal?
            strTemp = Replace(Val(strValue), ",", ".")
            'ProcessValue = GetWholeNumber(strTemp) & GetDecimalNumber(strTemp)
            ProcessValue = GetWholeNumber(strTemp) & GetDecimalNumber(strValue)
        ElseIf IsNumeric(Replace(strValue, ".", ",")) Then  'For language other than english IAN 04-07-05
            strTemp = Replace(Val(strValue), ",", ".")
            ProcessValue = GetWholeNumber(strTemp) & GetDecimalNumber(strValue)
        Else
            ProcessValue = strValue
        End If
    Else
        ProcessValue = strValue
    End If
    '================================================'
    
End Function

Public Function GetDataNCTSRecordset(ByVal DataNCTSMessageID As Long, ByVal UniqueCode As String) As ADODB.Recordset
Dim strSqlDataNctsMessage As String
Dim rstDataNCTSMessage As ADODB.Recordset

    strSqlDataNctsMessage = "SHAPE {" & GetQryNCTSDataMessage(DataNCTSMessageID) & "} AS DATA_NCTS_MESSAGES "
    strSqlDataNctsMessage = strSqlDataNctsMessage & "APPEND "
    strSqlDataNctsMessage = strSqlDataNctsMessage & PrepareDataNCTSSegments(EMsg_IE15, DataNCTSMessageID)
        
    ADORecordsetOpen strSqlDataNctsMessage, mconADOConnectionShaping, rstDataNCTSMessage, adOpenKeyset, adLockReadOnly, , True

    Set GetDataNCTSRecordset = rstDataNCTSMessage.Clone
    
    ADORecordsetClose rstDataNCTSMessage
    
End Function

Private Function GetLastDetailwithF(ByVal Code As String) As Long
    Dim rstDetail As ADODB.Recordset
    Dim strSQL As String
    Dim lngTemp As Long
    
    strSQL = "Select MIN(Detail) as Min_Detail from DATA_NCTS_DETAIL where T7 = 'F' and Code = '" & Code & "'"
    
    
    ADORecordsetOpen strSQL, mconADOConnectionStandard, rstDetail, adOpenKeyset, adLockReadOnly
    
    If Not rstDetail.BOF And Not rstDetail.EOF Then
        
        If IsNull(rstDetail!Min_Detail) Then
            ADORecordsetClose rstDetail
            
            strSQL = "Select Max(Detail) as Max_Detail from DATA_NCTS_DETAIL where Code = '" & Code & "'"
            ADORecordsetOpen strSQL, mconADOConnectionStandard, rstDetail, adOpenKeyset, adLockReadOnly
            
            If Not rstDetail.BOF And Not rstDetail.EOF Then
                If IsNull(rstDetail!Max_Detail) Then
                    lngTemp = 0
                Else
                    lngTemp = rstDetail!Max_Detail
                End If
            End If
            
            ADORecordsetClose rstDetail
        Else
            lngTemp = rstDetail!Min_Detail
            ADORecordsetClose rstDetail
        End If
        
    Else
    
        ADORecordsetClose rstDetail
    End If
    
    GetLastDetailwithF = lngTemp
    
End Function
