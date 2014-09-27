Attribute VB_Name = "MProcedures"
Option Explicit


Private Function IsEmptySegment(ByVal strSegment As String, strPrefix As String) As Boolean
    
    Dim strTrimSegment As String
    Dim strTrimPrefix As String
    Const CONST_DUMMY = "^&*("
    
    ' Replace Special Characters with distinct characters (DUMMY)
    ' so that it will be detected as non-empty
    '   RELEASE CHARACTER
    strSegment = Replace(strSegment, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER, CONST_DUMMY)
    '   SEGMENT SEPARATOR
    strSegment = Replace(strSegment, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, CONST_DUMMY)
    '   COMPOSITE DATA ELEMENT SEPARATOR
    strSegment = Replace(strSegment, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT, CONST_DUMMY)
    '   SIMPLE DATA ELEMENT SEPARATOR
    strSegment = Replace(strSegment, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT, CONST_DUMMY)
    
    strTrimSegment = Replace(strSegment, "+", "")
    strTrimSegment = UCase(Replace(strTrimSegment, ":", ""))
    
    strTrimPrefix = Replace(strPrefix, "+", "")
    strTrimPrefix = UCase(Replace(strTrimPrefix, ":", ""))
    
    IsEmptySegment = (strTrimSegment = strTrimPrefix)
    
End Function

'Added by Migs on 08-10-2006.
'This procedure will remove unnecessaary field separators ('+' and ':')
'Example: "NAD+AE+BE0000428730003:+++::++this is a test+++::::+" ---> "NAD+AE+BE0000428730003+++++this is a test"
'Notice the FOR loop 'For lngMAX = 1 to 10'. This was placed to remove the
'unnecessary colons in field groups. This is done 10 times since in the MIG v0.9.0.0,
'a field group is separated by colons in no more than 10 times. Notice also that
'we can change the constant 10 to any number higher than this, but 10 is the optimal
'choice since placing a large value may affect the speed of processing. Change this considerably.
Public Function TrimSegment(ByVal Segment As String, ByVal Prefix As String) As String

    Dim strTemp As String
    Dim strChar As String
    Dim lngCtr As Long
    Dim lngMAX As Long
    'EDI_SEP_SEGMENT
    'EDI_SEP_COMPOSITE_DATA_ELEMENT
    'EDI_SEP_DATA_ELEMENT
    'EDI_SEP_RELEASE_CHARACTER

    Const CONST_REPLACE_PREFIX = "^&*("
    
    Const CONST_REPLACE_RELEASE_CHARACTER = "!"
    Const CONST_REPLACE_SEP_SEGMENT = "@"
    Const CONST_REPLACE_SEP_COMPOSITE_DATA_ELEMENT = "#"
    Const CONST_REPLACE_SEP_DATA_ELEMENT = "$"
    
    ' 1234  6789  --->  numeric
    ' !@#$  ^&*(  --->  numeric + SHIFT
    
    
    strTemp = Left(Segment, Len(Segment) - 1)
    
    ' ------------ '
    ' START of (A)
    ' ------------ '
    ' Replace "?" & EDI_SEP_RELEASE_CHARACTER with "^&*(" & "!"
    strTemp = Replace(strTemp, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER, CONST_REPLACE_PREFIX & CONST_REPLACE_RELEASE_CHARACTER)
    
    ' Replace "?" & EDI_SEP_SEGMENT with "^&*(" & "@"
    strTemp = Replace(strTemp, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_SEGMENT)
    
    ' Replace "?" & EDI_SEP_COMPOSITE_DATA_ELEMENT with "^&*(" & "#"
    strTemp = Replace(strTemp, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_COMPOSITE_DATA_ELEMENT)
    
    ' Replace "?" & EDI_SEP_DATA_ELEMENT with "^&*(" & "$"
    strTemp = Replace(strTemp, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_DATA_ELEMENT)
    ' ------------ '
    ' END of (A)
    ' ------------ '
    
    
    ' We are sure at this point that there are no special characters prefixed with a release character;
    ' all that remains are official separation (special) characters
    For lngCtr = 1 To Len(Segment)
        strChar = Right(strTemp, 1)
        If (strChar = ":" Or strChar = "+" Or strChar = "-") Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        Else
            Exit For
        End If
    Next
    
    For lngMAX = 1 To 10
        For lngCtr = 1 To Len(strTemp)
            strChar = Mid(strTemp, lngCtr, 1)
            
            ' Check if ":+"
            If strChar = ":" Then
                If InStr(lngCtr, strTemp, "+") = lngCtr + 1 Then
                    ' Remove ":" from ":+"
                    strTemp = Mid(strTemp, 1, lngCtr - 1) & Mid(strTemp, lngCtr + 1)
                End If
            End If
        Next lngCtr
    Next lngMAX
    
    ' Revert replaced released special characters to original
    ' This must be in reverse order of the above (A)
    
    ' Replace "^&*(" & "$" with "?" & EDI_SEP_DATA_ELEMENT
    strTemp = Replace(strTemp, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT)
    
    ' Replace "^&*(" & "#" with "?" & EDI_SEP_COMPOSITE_DATA_ELEMENT
    strTemp = Replace(strTemp, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_COMPOSITE_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT)
    
    ' Replace "^&*(" & "@" with "?" & EDI_SEP_SEGMENT
    strTemp = Replace(strTemp, CONST_REPLACE_PREFIX & CONST_REPLACE_SEP_SEGMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
    
    ' Replace "^&*(" & "!" with "?" & EDI_SEP_RELEASE_CHARACTER
    strTemp = Replace(strTemp, CONST_REPLACE_PREFIX & CONST_REPLACE_RELEASE_CHARACTER, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER)

    If IsEmptySegment(strTemp, Prefix) Then
        TrimSegment = ""
    Else
        TrimSegment = strTemp & "'"
    End If

End Function

Public Function CountSegments(ByVal Segment As String) As Long

    Dim astrTemp() As String
    Dim strTemporarySegment As String
    
    Const CONST_DUMMY = "^&*("
    
    ' Replace Special Characters with distinct characters (DUMMY)
    ' so that it will not be detected as a separator
    '   RELEASE CHARACTER
    strTemporarySegment = Segment
    
    strTemporarySegment = Replace(strTemporarySegment, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT, CONST_DUMMY)
    
    astrTemp = Split(strTemporarySegment, EDI_SEP_SEGMENT)
    
    If UBound(astrTemp) = -1 Then
    
        CountSegments = 0
    
    Else
    
        CountSegments = UBound(astrTemp)
    
    End If
    
End Function

'Function Modified by Migs 04/04/2006
'This will check the declaration type of all the details and return a boolean
'indicating the occurence of the strDocType in all details whether it is needed '='
'or not needed '<>' as specified by strCondition
Public Function CheckRegimeType(ByVal strDocType As String, ByVal strCondition As String) As Boolean
    
    Dim intCtr As Integer
    Dim blnFuncVal As Boolean
    Dim rstCloneDetails As ADODB.Recordset
    
    Set rstCloneDetails = g_rstDetails.Clone
    
    'g_rstDetails.MoveFirst
    rstCloneDetails.MoveFirst
    
    If strCondition = "<>" Then
        
        blnFuncVal = True
        
        'For intCtr = 1 To g_rstDetails.RecordCount
        '    blnFuncVal = blnFuncVal And (strDocType = g_rstDetails!N4)
        '    g_rstDetails.MoveNext
        'Next intCtr
        For intCtr = 1 To rstCloneDetails.RecordCount
            blnFuncVal = blnFuncVal And (UCase(Trim(strDocType)) = UCase(Trim(FNullField(rstCloneDetails.Fields("N4").Value))))
            rstCloneDetails.MoveNext
        Next intCtr
        
        CheckRegimeType = Not blnFuncVal
        
    ElseIf strCondition = "=" Then
    
        blnFuncVal = False
        
        For intCtr = 1 To rstCloneDetails.RecordCount
            blnFuncVal = blnFuncVal Or (UCase(Trim(strDocType)) = UCase(Trim(FNullField(rstCloneDetails.Fields("N4").Value))))
            rstCloneDetails.MoveNext
        Next intCtr
        'For intCtr = 1 To g_rstDetails.RecordCount
        '    blnFuncVal = blnFuncVal Or (strDocType = g_rstDetails!N4)
        '    g_rstDetails.MoveNext
        'Next intCtr
        
        CheckRegimeType = blnFuncVal
    
    End If
    
    'g_rstDetails.MoveFirst 'Added by Migs 04/24/2006 para hindi magka-error sa pag access ng recordset na to after magamit dito sa proc
    rstCloneDetails.MoveFirst 'Added by Migs 04/24/2006 para hindi magka-error sa pag access ng recordset na to after magamit dito sa proc
    ' Actually no need for this code anymore after replacing g_rstDetails with rstCloneDetails
End Function

'Added by Migs 04/04/2006
'This will check CN for all details and if their Venture Nos.
'specified by the parameter strVentureNumber are equal, the function
'returns true otherwise, false
'
'Modifed by Migs 04/05/2006
'Changed the name of the proc to accomodate din ung EX bukod sa CN
'para hindi na gumawa pa ng isa pang separate proc...
'then added another parameter strQualifier to specify the type of checking
'same pa rin ang logic ng proc
Public Function VentureNumbersAreSame(ByVal strVentureNumber As String, ByVal strQualifier As String, ByVal DeclarationType As DECLARATION_TYPE) As Boolean
    
    Dim intCtr As Integer
    Dim blnFuncVal As Boolean
        
    blnFuncVal = True
    Dim rstHandelaarsClone As ADODB.Recordset
    
    Set rstHandelaarsClone = g_rstDetailsHandelaars.Clone
    
    If strQualifier = "CN" Then
    
        'If Not g_rstDetailsHandelaars Is Nothing Then
        If DeclarationType = enuImport Then     'Glenn
            'Added by Migs on 08-09-2006. This for checking the consignees in the new table
            'g_rstDetailsHandelaars.Filter = "VE = '" & enuDetailConsignee & "'"
            rstHandelaarsClone.Filter = "VE = '" & enuDetailConsignee & "'"
        Else
            'Glenn
            'g_rstDetailsHandelaars.Filter = "VE = '" & enuExportDetailConsignee & "'"
            rstHandelaarsClone.Filter = "VE = '" & enuExportDetailConsignee & "'"
        End If
        
        If rstHandelaarsClone.RecordCount > 0 Then
            rstHandelaarsClone.MoveFirst
            
            For intCtr = 1 To rstHandelaarsClone.RecordCount
                blnFuncVal = blnFuncVal And (strVentureNumber = FNullField(rstHandelaarsClone.Fields("V1").Value) Or (Len(FNullField(rstHandelaarsClone.Fields("V1").Value)) = 0))
                rstHandelaarsClone.MoveNext
            Next intCtr
            
            rstHandelaarsClone.Filter = adFilterNone
            rstHandelaarsClone.MoveFirst
        End If
        
        'If g_rstDetailsHandelaars.RecordCount > 0 Then
        '    g_rstDetailsHandelaars.MoveFirst
       '
       '     For intCtr = 1 To g_rstDetailsHandelaars.RecordCount
       '         blnFuncVal = blnFuncVal And (strVentureNumber = g_rstDetailsHandelaars!V1 Or (Len(g_rstDetailsHandelaars!V1) = 0))
       '         g_rstDetailsHandelaars.MoveNext
       '     Next intCtr
       '
       '     g_rstDetailsHandelaars.Filter = adFilterNone
       '     g_rstDetailsHandelaars.MoveFirst
       ' End If
        
       ' Else
       '
       '     g_rstDetails.MoveFirst
       '
       '     For intCtr = 1 To g_rstDetails.RecordCount
       '         blnFuncVal = blnFuncVal And (strVentureNumber = g_rstDetails!V1)
       '         g_rstDetails.MoveNext
       '     Next intCtr
       '
       '     g_rstDetails.MoveFirst
       '
       ' End If
        
    End If
    
    If strQualifier = "EX" Then
    
        'Glenn 9/7/2006
        'g_rstDetailsHandelaars.Filter = "VE = '" & enuExportDetailExporter & "'"
        rstHandelaarsClone.Filter = "VE = '" & enuExportDetailExporter & "'"
        
        If rstHandelaarsClone.RecordCount > 0 Then
            rstHandelaarsClone.MoveFirst
        
            For intCtr = 1 To rstHandelaarsClone.RecordCount
                blnFuncVal = blnFuncVal And (strVentureNumber = FNullField(rstHandelaarsClone.Fields("V1").Value) Or (Len(FNullField(rstHandelaarsClone.Fields("V1").Value)) = 0))
                rstHandelaarsClone.MoveNext
            Next intCtr
            
            rstHandelaarsClone.Filter = adFilterNone
            rstHandelaarsClone.MoveFirst
        End If
        'If g_rstDetailsHandelaars.RecordCount > 0 Then
        '    g_rstDetailsHandelaars.MoveFirst
       '
       '     For intCtr = 1 To g_rstDetailsHandelaars.RecordCount
       '         blnFuncVal = blnFuncVal And (strVentureNumber = g_rstDetailsHandelaars!V1 Or (Len(g_rstDetailsHandelaars!V1) = 0))
       '         g_rstDetailsHandelaars.MoveNext
       '     Next intCtr
       '
       '     g_rstDetailsHandelaars.Filter = adFilterNone
       '     g_rstDetailsHandelaars.MoveFirst
       ' End If
        'g_rstDetails.MoveFirst
       '
       ' For intCtr = 1 To g_rstDetails.RecordCount
       '     blnFuncVal = blnFuncVal And (strVentureNumber = g_rstDetails!Y1)
       '     g_rstDetails.MoveNext
       ' Next intCtr
        
       ' g_rstDetails.MoveFirst 'Added by Migs 04/24/2006 para hindi magka-error sa pag access ng recordset na to after magamit dito sa proc
        
        
    End If
    
    VentureNumbersAreSame = blnFuncVal
    
End Function

Public Function GetBoxCode(ByVal DataItemOrdinalNumber As Integer, _
                           ByVal CompositeDataItemOrdinalNumber As Integer, _
                           ByVal SegmentName As String, _
                           ByVal Qualifier As String, _
                           ByVal DType As Long, _
                  Optional ByVal GroupQualifier As String = "", _
                  Optional ByVal IsDetail As Boolean = False) As String
                            
    Select Case SegmentName
        Case "BGM"
            Select Case DataItemOrdinalNumber
                Case 1
                    If IsDetail = False Then
                        GetBoxCode = "A1*A2"
                    End If
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            GetBoxCode = "MRN"
                        Case 2
                            If IsDetail = False Then GetBoxCode = "A9"
                            
                    End Select
            End Select
                
        Case "LOC"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "41"
                                    If IsDetail = False Then GetBoxCode = "A6"
                                    
                                Case "43"
                                    If IsDetail = False Then GetBoxCode = "D4"
                                    
                                Case "35", "9"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "DA"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "DB"
                                        
                                    End If
                                    
                                Case "91"
                                    Select Case GroupQualifier
                                        Case ""
                                            If IsDetail = False Then GetBoxCode = "A5"
                                            
                                        Case "190"
                                            If IsDetail = True Then GetBoxCode = "Q2"
                                            
                                    End Select
                                Case "42"
                                    If IsDetail = False Then GetBoxCode = "A7"
                                    
                                Case "27"
                                    If IsDetail = True Then GetBoxCode = "NB"
                                    
                                    
                                Case "28", "11"
                                    If IsDetail = True Then GetBoxCode = "N7"
                                    
                                Case "47"
                                    If IsDetail = True Then GetBoxCode = "N8"
                                    
                                Case "18"
                                    If IsDetail = True Then GetBoxCode = "M3*M4*M5"
                                    
                                Case "127"
                                    Select Case GroupQualifier
                                        Case "998"
                                            If IsDetail = True Then GetBoxCode = "R9"

                                    End Select
                                Case "106"
                                    If IsDetail = True Then GetBoxCode = "NC"
                                    
                            End Select
                            
                        Case 4
                            Select Case Qualifier
                                Case "1"
                                    If Not IsDetail Then
                                        GetBoxCode = "C3"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "O4"
                                    End If
                            End Select
                            
                    End Select
                    
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "91"
                                    Select Case GroupQualifier
                                        Case "190"
                                            If IsDetail = True Then GetBoxCode = "QB"
                                            
                                    End Select
                            End Select
                        Case 4
                            Select Case Qualifier
                                Case "91"
                                    Select Case GroupQualifier
                                        Case "190"
                                            If IsDetail = True Then GetBoxCode = "QC"
                                            
                                    End Select
                            End Select
                    End Select
                    
                Case 5
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "91"
                                    Select Case GroupQualifier
                                        Case "190"
                                            If IsDetail = True Then GetBoxCode = "Q5"
                                            
                                    End Select
                            End Select
                    End Select
            End Select
        
        Case "DTM"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 2
                            Select Case Qualifier
                                Case "254"
                                    If IsDetail = False Then GetBoxCode = "A4"
                                    
                                Case "137"
                                    Select Case GroupQualifier
                                        Case "190"
                                            If IsDetail = True Then GetBoxCode = "Q3"
                                            
                                        Case "998"
                                            If IsDetail = True Then GetBoxCode = "R3"
                                            
                                    End Select
                                Case "36"
                                    If IsDetail = True Then GetBoxCode = "T5"
                                    
                            End Select
                    End Select
            End Select
               
        Case "GEI"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "2"
                                    If Not IsDetail Then GetBoxCode = "B1"
                            End Select
                    End Select
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "2"
                                    If IsDetail Then GetBoxCode = "T3"
                            End Select
                    End Select
            End Select
       
        Case "FII"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If IsDetail = False Then GetBoxCode = "B5"
                                    
                            End Select
                        Case 2
                            Select Case Qualifier
                                Case "AE"
                                    If IsDetail = False Then GetBoxCode = "B4"
                                    
                            End Select
                    End Select
            End Select
                    
        Case "SEL"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = False Then GetBoxCode = "E2"
                            
                        Case 4
                            If IsDetail = False Then GetBoxCode = "E1"
                            
                    End Select
            End Select
                    
        Case "FTX"
            Select Case DataItemOrdinalNumber
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AAB"
                                    If Not IsDetail Then
                                        GetBoxCode = "C7"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "N9"
                                    End If
                                    
                                Case "PMT"
                                    If Not IsDetail Then
                                        GetBoxCode = "B2"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "T2"
                                    End If
                                    
                                Case "ACB"
                                    If IsDetail = True Then GetBoxCode = "P1"
                                    
                                Case "CCI"
                                    If IsDetail = True Then GetBoxCode = "O1"
                                    
                                Case "CUS"
                                    If IsDetail = True Then GetBoxCode = "N4"
                                    
                            End Select
                    End Select
                Case 4
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "PMT"
                                    If Not IsDetail Then
                                        GetBoxCode = "B6"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "T1"
                                    End If
                                Case "AAA"
                                    If IsDetail = True Then GetBoxCode = "L8"
                                    
                                Case "ACB"
                                    If IsDetail = True Then GetBoxCode = "P2"
                                    
                            End Select
                    End Select
            End Select
            
        Case "RFF"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 2
                            Select Case Qualifier
                                Case "ABE"
                                    Select Case GroupQualifier
                                        Case ""
                                            If IsDetail = False Then GetBoxCode = "A3"
                                            
                                        Case "AE"
                                            If IsDetail = False Then GetBoxCode = "AB"
                                            
                                    End Select
                                Case "ABO"
                                    If IsDetail = False Then GetBoxCode = "AC"
                                    
                                Case "AKO"
                                    If IsDetail = False Then GetBoxCode = "C1"
                                    
                                Case "AHP"
                                    If IsDetail = False Then GetBoxCode = "XF(2)*XD(2)"
                                    
                                Case "AAQ"
                                    If IsDetail = True Then GetBoxCode = "S4,S5"
                                    
                                Case "UCN"
                                    If IsDetail = True Then GetBoxCode = "LC"
                                    
                                Case "ABJ"
                                    If IsDetail = True Then GetBoxCode = "L7"
                                    
                                Case "AIP"
                                    If IsDetail = True Then GetBoxCode = "N5"
                                    
                                Case "ANV"
                                    If IsDetail = True Then GetBoxCode = "T4"
                                    
                            End Select
                        Case 4
                            Select Case Qualifier
                                Case "ANV"
                                    If IsDetail = False Then GetBoxCode = "T6"
                                    
                            End Select
                    End Select
            End Select
            
        Case "TDT"
            Select Case DataItemOrdinalNumber
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "11"
                                    If IsDetail = False Then GetBoxCode = "D8"
                                Case "1"
                                    If IsDetail = False Then GetBoxCode = "D9"
                            End Select
                    End Select
                Case 8
                    Select Case CompositeDataItemOrdinalNumber
                        Case 4
                            Select Case Qualifier
                                Case "13", "12"
                                    If IsDetail = False Then GetBoxCode = "D5"
                                Case "11"
                                    If IsDetail = False Then GetBoxCode = "D6"
                            End Select
                        Case 5
                            Select Case Qualifier
                                Case "11"
                                    If IsDetail = False Then GetBoxCode = "D7"
                            End Select
                    End Select
            End Select
            
        Case "DOC"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 4
                            Select Case Qualifier
                                Case "190"
                                    If IsDetail = True Then GetBoxCode = "Q1"
                                    
                                Case "998"
                                    If IsDetail = True Then GetBoxCode = "R1*R2"
                                    
                            End Select
                    End Select
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "122"
                                    If IsDetail = False Then GetBoxCode = "A8"
                                    
                                Case "998"
                                    If IsDetail = True Then GetBoxCode = "R5"
                                    
                            End Select
                        Case 3
                            Select Case Qualifier
                                Case "190"
                                    If IsDetail = True Then GetBoxCode = "Q9"
                                    
                                Case "998"
                                    If IsDetail = True Then GetBoxCode = "R6"
                                    
                            End Select
                    End Select
            End Select
            
        Case "NAD"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X1(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X1(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X1(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V1(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X1(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V1(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X1(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V1(3)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X1(2)"
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X1(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V1(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X1(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X1(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V1(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V1(4)"
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V1(3)"
                                    End If
                            End Select
                    End Select
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If IsDetail = False Then GetBoxCode = "AA"
                                    
                            End Select
                    End Select
                Case 4
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X2(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X2(1)"
                                        
                                    End If
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X2(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V2(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X2(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V2(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X2(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V2(3)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X2(2)"
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X2(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V2(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X2(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X2(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V2(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V2(4)"
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V2(3)"
                                    End If
                                    
                            End Select
                    End Select
                Case 5
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X3(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X3(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X3(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V3(1)"
                                        End If
                                        
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X3(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V3(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X3(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V3(3)"
                                        End If
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X3(2)"
                                        
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X3(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V3(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X3(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X3(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V3(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V3(4)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V3(3)"
                                        
                                    End If
                                    
                            End Select
                            
                        Case 2
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X4(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X4(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X4(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V4(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X4(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V4(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X4(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V4(3)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X4(2)"
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X4(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V4(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X4(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X4(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V4(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V4(4)"
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V4(3)"
                                    End If
                                    
                            End Select
                    End Select
                    
                Case 6
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X6(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X6(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X6(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V6(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X6(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V6(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X6(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V6(3)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X6(2)"
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X6(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V6(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X6(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X6(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V6(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V6(4)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V6(3)"
                                        
                                    End If
                                    
                            End Select
                    End Select
                Case 7
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X7(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X7(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X7(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V7(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X7(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V7(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X7(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V7(3)"
                                        End If
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X7(2)"
                                        
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X7(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V7(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X7(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X7(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V7(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V7(4)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V7(3)"
                                        
                                    End If
                                    
                            End Select
                    End Select
                Case 8
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        GetBoxCode = "X5(2)"
                                    ElseIf DType = 18 Then
                                        GetBoxCode = "X5(1)"
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X5(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V5(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X5(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V5(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X5(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V5(3)"
                                        End If
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X5(2)"
                                        
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X5(4)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V5(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X5(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X5(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V5(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V5(4)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V5(3)"
                                        
                                    End If
                                    
                            End Select
                    End Select
                Case 9
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AE"
                                    If DType = 14 Then
                                        If IsDetail = False Then GetBoxCode = "X8(2)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X8(1)"
                                        
                                    End If
                                    
                                Case "CN"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X8(1)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V8(1)"
                                        End If
                                    ElseIf DType = 18 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X8(4)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V8(2)"
                                        End If
                                    End If
                                    
                                Case "AG"
                                    If DType = 14 Then
                                        If Not IsDetail Then
                                            GetBoxCode = "X8(3)"
                                        ElseIf IsDetail Then
                                            GetBoxCode = "V8(3)"
                                        End If
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = False Then GetBoxCode = "X8(2)"
                                        
                                    End If
                                    
                                Case "GS"
                                    If Not IsDetail Then
                                        GetBoxCode = "X8(4)"
                                    ElseIf Not IsDetail Then
                                        GetBoxCode = "V8(2)"
                                    End If
                                    
                                Case "BE"
                                    If IsDetail = False Then GetBoxCode = "X8(3)"
                                    
                                Case "EX"
                                    If Not IsDetail Then
                                        GetBoxCode = "X8(5)"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "V8(1)"
                                    End If
                                    
                                Case "WD"
                                    If DType = 14 Then
                                        If IsDetail = True Then GetBoxCode = "V8(4)"
                                        
                                    ElseIf DType = 18 Then
                                        If IsDetail = True Then GetBoxCode = "V8(3)"
                                        
                                    End If
                                    
                            End Select
                    End Select
            End Select
        
        Case "TOD"
            Select Case DataItemOrdinalNumber
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "6"
                                    If Not IsDetail Then
                                        GetBoxCode = "C2"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "O3"
                                    End If
                            End Select
                    End Select
            End Select
        
        Case "MOA"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 2
                            Select Case Qualifier
                                Case "69", "39"
                                    If Not IsDetail Then
                                        GetBoxCode = "C4"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "O7"
                                    End If
                                    
                                Case "72"
                                    If Not IsDetail Then
                                        GetBoxCode = "C9"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "O9"
                                    End If
                                    
                                Case "123"
                                    If IsDetail = True Then GetBoxCode = "O2"
                                    
                                Case "38"
                                    If IsDetail = True Then GetBoxCode = "O5"
                                    
                                Case "55"
                                    If IsDetail = True Then GetBoxCode = "U2"
                                    
                            End Select
                            
                        Case 3
                            Select Case Qualifier
                                Case "69", "39"
                                    If Not IsDetail Then
                                        GetBoxCode = "C5"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "O8"
                                    End If
                                    
                                Case "72"
                                    If Not IsDetail Then
                                        GetBoxCode = "CA"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "OA"
                                    End If
                                    
                                Case "38"
                                    If IsDetail = True Then GetBoxCode = "O6"
                                    
                            End Select
                    End Select
            End Select
        
        Case "CUX"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 2
                            Select Case GroupQualifier
                                Case "69", "39"
                                    If IsDetail = False Then GetBoxCode = "C5"
                                    
                                Case "72"
                                    If IsDetail = False Then GetBoxCode = "CA"
                                    
                            End Select
                    End Select
                    
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case GroupQualifier
                                Case "69", "39"
                                    If Not IsDetail Then
                                        GetBoxCode = "C6"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "OC"
                                    End If
                                Case "72"
                                    If Not IsDetail Then
                                        GetBoxCode = "CB"
                                    ElseIf IsDetail Then
                                        GetBoxCode = "OD"
                                    End If
                                Case "38", "123"
                                    If IsDetail = True Then GetBoxCode = "OB"
                            End Select
                    End Select
            End Select
        
        Case "CST"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "L1"
                    End Select
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "L2"
                    End Select
                Case 4
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "L3"
                    End Select
                Case 5
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "L4"
                    End Select
                Case 6
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "L5*L6"
                    End Select
            End Select
        
        Case "MEA"
            Select Case DataItemOrdinalNumber
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "AAE"
                                    If IsDetail = True Then GetBoxCode = "M1"
                                    
                                Case "AAF"
                                    If IsDetail = True Then GetBoxCode = "TZ"
                                    
                            End Select
                            
                        Case 2
                            Select Case Qualifier
                                Case "AAB"
                                    If IsDetail = True Then GetBoxCode = "L9"
                                Case "AAA"
                                    If IsDetail = True Then GetBoxCode = "LA"
                                Case "AAE"
                                    If IsDetail = True Then GetBoxCode = "M2"
                                Case "AAF"
                                    If IsDetail = True Then GetBoxCode = "T8"
                            End Select
                    End Select
            End Select
        
        Case "PAC"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "S2"
                    End Select
                Case 3
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            If IsDetail = True Then GetBoxCode = "S1"
                    End Select
            End Select
        
        Case "PCI"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1, 2, 3
                            Select Case Qualifier
                                Case "28"
                                    If IsDetail = True Then GetBoxCode = "S3"
                            End Select
                    End Select
            End Select
        
        Case "TAX"
            Select Case DataItemOrdinalNumber
                Case 2
                    Select Case CompositeDataItemOrdinalNumber
                        Case 1
                            Select Case Qualifier
                                Case "1"
                                    If IsDetail = True Then GetBoxCode = "U1"
                            End Select
                    End Select
            End Select
        
        Case "CNT"
            Select Case DataItemOrdinalNumber
                Case 1
                    Select Case CompositeDataItemOrdinalNumber
                        Case 2
                            Select Case Qualifier
                                Case "10"
                                    If IsDetail = False Then GetBoxCode = "D3"
                                Case "26"
                                    If IsDetail = False Then GetBoxCode = "D1"
                                Case "18"
                                    If IsDetail = False Then GetBoxCode = "D2"
                            End Select
                    End Select
            End Select
        
    End Select
                    
End Function


Public Function AddFieldToDigiSign(ByVal FieldName As String, _
                                   ByVal FieldRecordset As ADODB.Recordset, _
                          Optional ByVal FieldIsDetail As Boolean, _
                          Optional ByVal DetailNumber As Long, _
                          Optional ByVal HandelaarOrdinal As Long = 0, _
                          Optional ByVal GroupOrdinal As Long)

    Dim strFieldOrdinal As String
    Dim arrFieldOrdinal() As String
    
    Dim strTableName As String
    Dim strFieldValue As String
    Dim lngDetailNumber As Long
    
    Dim lngCtr As Long
    
    Dim rstTemp As ADODB.Recordset
    
    '*******************************************************************************************************************
    'Funtion to Build Records to be Digitally Signed
    '       FieldName           - Field to add to the digital string
    '       FieldRecordset      - Recordset to use to build the digital string for the field
    '       DetailNumber        - A long signifying the Detail number of the field
    '       HandelaarOrdinal    - Handelaar Type based on field "XE"
    '*******************************************************************************************************************
    'Bypass DigiSign Procedure when DigiSign Indicator is False
    If g_blnDigiSignActivated = False Then Exit Function
    
    Set rstTemp = FieldRecordset.Clone
    
    'Check if field is to be digitally signed
    strFieldOrdinal = GetFieldOrdinal(FieldName, HandelaarOrdinal)
    
    If Trim$(strFieldOrdinal) = vbNullString Then Exit Function
    
    If InStr(1, strFieldOrdinal, ",") > 0 Then
        arrFieldOrdinal = Split(strFieldOrdinal, ",")
    Else
        ReDim arrFieldOrdinal(0)
        arrFieldOrdinal(0) = strFieldOrdinal
    End If
    
    For lngCtr = LBound(arrFieldOrdinal) To UBound(arrFieldOrdinal)
        lngDetailNumber = DetailNumber
        
        'Get field value and DetailNumber in case of header/summary fields
        GetFieldValueAndDetailNumber strFieldValue, _
                                     lngDetailNumber, _
                                     CLng(arrFieldOrdinal(lngCtr)), _
                                     rstTemp, _
                                     FieldIsDetail, _
                                     GroupOrdinal, _
                                     strTableName
        
        'Add record to packing recordset
        If Trim$(strFieldValue) <> vbNullString Then
            Select Case UCase$(strTableName)
                Case "PLDA IMPORT HEADER", "PLDA IMPORT DETAIL", "PROCEDURE", _
                     "PLDA COMBINED HEADER", "PLDA COMBINED DETAIL", _
                     "PLDA IMPORT HEADER HANDELAARS", "PLDA IMPORT DETAIL HANDELAARS", _
                     "PLDA COMBINED HEADER HANDELAARS", "PLDA COMBINED DETAIL HANDELAARS"

                    g_rstDigiSignData.AddNew
                    g_rstDigiSignData.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDigiSignData.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDigiSignData.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDigiSignData.Fields("Value").Value = strFieldValue
                    g_rstDigiSignData.Update
                
                Case "PLDA IMPORT HEADER ZEGELS", "PLDA COMBINED HEADER ZEGELS"
                    
                    g_rstHeaderSeals.AddNew
                    g_rstHeaderSeals.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstHeaderSeals.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstHeaderSeals.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstHeaderSeals.Fields("Value").Value = strFieldValue
                    g_rstHeaderSeals.Update
                    
                Case "PLDA IMPORT DETAIL BIJZONDERE", "PLDA COMBINED DETAIL BIJZONDERE"
                    
                    g_rstDetailBijzondere.AddNew
                    g_rstDetailBijzondere.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailBijzondere.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailBijzondere.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailBijzondere.Fields("Value").Value = strFieldValue
                    g_rstDetailBijzondere.Update
                
                Case "PLDA IMPORT DETAIL BEREKENINGS EENHEDEN"
                    
                    g_rstDetailBerekenings.AddNew
                    g_rstDetailBerekenings.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailBerekenings.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailBerekenings.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailBerekenings.Fields("Value").Value = strFieldValue
                    g_rstDetailBerekenings.Update
                    
                Case "PLDA IMPORT DETAIL DOCUMENTEN", "PLDA COMBINED DETAIL DOCUMENTEN"
                    
                    g_rstDetailDocumenten.AddNew
                    g_rstDetailDocumenten.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailDocumenten.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailDocumenten.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailDocumenten.Fields("Value").Value = strFieldValue
                    g_rstDetailDocumenten.Update
                    
                Case "PLDA IMPORT DETAIL ZELF"
                    
                    g_rstDetailZelf.AddNew
                    g_rstDetailZelf.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailZelf.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailZelf.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailZelf.Fields("Value").Value = strFieldValue
                    g_rstDetailZelf.Update
                
                'CSCLP-95
                Case "PLDA IMPORT DETAIL CONTAINER", "PLDA COMBINED DETAIL CONTAINER"
                'Case "PLDA IMPORT DETAIL CONTAINER", "PLDA EXPORT DETAIL CONTAINER"
                    
                    g_rstDetailContainer.AddNew
                    g_rstDetailContainer.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailContainer.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailContainer.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailContainer.Fields("Value").Value = strFieldValue
                    g_rstDetailContainer.Update
                    
                Case "PLDA COMBINED HEADER TRANSIT OFFICES"
                    
                    g_rstHeaderTransitOffices.AddNew
                    g_rstHeaderTransitOffices.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstHeaderTransitOffices.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstHeaderTransitOffices.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstHeaderTransitOffices.Fields("Value").Value = strFieldValue
                    g_rstHeaderTransitOffices.Update
                    
                Case "PLDA COMBINED HEADER ZEKERHEID"
                    
                    g_rstHeaderGuarantee.AddNew
                    g_rstHeaderGuarantee.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstHeaderGuarantee.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstHeaderGuarantee.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstHeaderGuarantee.Fields("Value").Value = strFieldValue
                    g_rstHeaderGuarantee.Update
                    
                Case "PLDA COMBINED DETAIL SENSITIVE GOODS"
                    
                    g_rstDetailSensitiveGoods.AddNew
                    g_rstDetailSensitiveGoods.Fields("GroupOrdinal").Value = GroupOrdinal
                    g_rstDetailSensitiveGoods.Fields("DigiSignOrdinal").Value = arrFieldOrdinal(lngCtr)
                    g_rstDetailSensitiveGoods.Fields("DetailNumber").Value = lngDetailNumber
                    g_rstDetailSensitiveGoods.Fields("Value").Value = strFieldValue
                    g_rstDetailSensitiveGoods.Update
                    
                Case Else
                    Debug.Assert False
                    
            End Select
                
        End If
    
    Next lngCtr
    
    'Destroy Temporary Copy of Recordset
    ADORecordsetClose rstTemp
    
End Function


Public Function BuildDigisignFieldVerifier(ByRef conConnection As ADODB.Connection, _
                                           ByVal DType As Long, _
                                           ByRef CallingForm As Object)
    
    'Dim rstDigiSign As ADODB.Recordset
    
    Dim strCommand As String
    
    Dim strTableToQuery As String
    Dim strTableOrdinal As String
    Dim strSourceTable As String
    Dim strDigiSignField As String
    Dim strDigiSignDataType As String
    Dim strDigiSignDetailNumber As String
    
    On Error GoTo ErrHandler
    
    'Initialize field verifier
    g_strDigiSignFieldVerifier = vbNullString
    
    'Get field names and table to query based on DType(Document Type)
    Select Case DType
        Case 18
            strTableToQuery = "DIGISIGN_PLDA_COMBINED"
            strTableOrdinal = "DigiSignCombined_Ordinal"
            strSourceTable = "DigiSignCombined_SourceTable"
            strDigiSignField = "DigiSignCombined_Fields"
            strDigiSignDataType = "DigiSignCombined_DataType"
            strDigiSignDetailNumber = "DigiSignCombined_DetailNumber"
            
        Case 14
            strTableToQuery = "DIGISIGN_PLDA_IMPORT"
            strTableOrdinal = "DigiSignImport_Ordinal"
            strSourceTable = "DigiSignImport_SourceTable"
            strDigiSignField = "DigiSignImport_Fields"
            strDigiSignDataType = "DigiSignImport_DataType"
            strDigiSignDetailNumber = "DigiSignImport_DetailNumber"
        
        Case Else
            CallingForm.DigiSignAddToTrace "Function: BuildDigisignFieldVerifier"
            CallingForm.DigiSignAddToTrace "Error: Dtype = " & DType & " is not valid."
            
            Exit Function
            
    End Select
    
    'Open DigiSign Table
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & strTableOrdinal & " as Ordinal, "
        strCommand = strCommand & strSourceTable & " as SourceTable, "
        strCommand = strCommand & strDigiSignField & " as Fields, "
        strCommand = strCommand & strDigiSignDataType & " as Datatype, "
        strCommand = strCommand & strDigiSignDetailNumber & " as DetailNumber "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & IIf(DType = 14, "DIGISIGN_PLDA_IMPORT", "DIGISIGN_PLDA_COMBINED") & " "
        strCommand = strCommand & "ORDER BY " & strTableOrdinal
    ADORecordsetOpen strCommand, conConnection, g_rstDigiSign, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, conConnection, g_rstDigiSign, adOpenKeyset, adLockOptimistic, , True
    
    'Build field verifier
    With g_rstDigiSign
    
        If .RecordCount > 0 Then
            .MoveFirst
            
            Do While Not .EOF
                If Not IsNull(.Fields("Fields").Value) And _
                   Not IsNull(.Fields("SourceTable").Value) And _
                   Not IsNull(.Fields("Datatype").Value) Then
                   
                    If Trim$(.Fields("Fields").Value) <> vbNullString And _
                       Trim$(.Fields("SourceTable").Value) <> vbNullString And _
                       Trim$(.Fields("Datatype").Value) <> vbNullString Then
                        
                        g_strDigiSignFieldVerifier = g_strDigiSignFieldVerifier & Trim$(.Fields("Fields").Value) & "@"
                        g_strDigiSignFieldVerifier = g_strDigiSignFieldVerifier & .Fields("Ordinal").Value & "*"
                        
                    End If
                End If
                
                .MoveNext
            Loop
            
        End If
        
    End With

ErrHandler:
    Select Case Err.Number
        Case 0
            'Do nothing
            
        Case Else
            CallingForm.DigiSignAddToTrace "Function: BuildDigisignFieldVerifier"
            CallingForm.DigiSignAddToTrace "Error: " & Err.Number & " - " & Err.Description
            
    End Select
    
End Function


Public Function InitializeRecordsToDigiSign(ByRef CallingForm As Object)
    
    On Error GoTo ErrHandler
    
    Set g_rstDigiSignData = New ADODB.Recordset
    Set g_rstHeaderSeals = New ADODB.Recordset
    Set g_rstHeaderTransitOffices = New ADODB.Recordset
    Set g_rstHeaderGuarantee = New ADODB.Recordset
    Set g_rstDetailBijzondere = New ADODB.Recordset
    Set g_rstDetailBerekenings = New ADODB.Recordset
    Set g_rstDetailDocumenten = New ADODB.Recordset
    Set g_rstDetailZelf = New ADODB.Recordset
    Set g_rstDetailContainer = New ADODB.Recordset
    Set g_rstDetailSensitiveGoods = New ADODB.Recordset
    
    'Header and Detail
    g_rstDigiSignData.CursorLocation = adUseClient
    
    g_rstDigiSignData.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDigiSignData.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDigiSignData.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDigiSignData.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDigiSignData.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDigiSignData.Open
    
    'Seals
    g_rstHeaderSeals.CursorLocation = adUseClient
    
    g_rstHeaderSeals.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstHeaderSeals.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderSeals.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderSeals.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstHeaderSeals.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstHeaderSeals.Open
    
    'Transit offices
    g_rstHeaderTransitOffices.CursorLocation = adUseClient
    
    g_rstHeaderTransitOffices.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstHeaderTransitOffices.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderTransitOffices.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderTransitOffices.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstHeaderTransitOffices.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstHeaderTransitOffices.Open
    
    'Zekerheid
    g_rstHeaderGuarantee.CursorLocation = adUseClient
    
    g_rstHeaderGuarantee.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstHeaderGuarantee.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderGuarantee.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstHeaderGuarantee.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstHeaderGuarantee.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstHeaderGuarantee.Open
    
    'Bijzondere
    g_rstDetailBijzondere.CursorLocation = adUseClient
    
    g_rstDetailBijzondere.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailBijzondere.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailBijzondere.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailBijzondere.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailBijzondere.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailBijzondere.Open
    
    
    'Berekenings
    g_rstDetailBerekenings.CursorLocation = adUseClient
    
    g_rstDetailBerekenings.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailBerekenings.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailBerekenings.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailBerekenings.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailBerekenings.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailBerekenings.Open
    
    'Documenten
    g_rstDetailDocumenten.CursorLocation = adUseClient
    
    g_rstDetailDocumenten.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailDocumenten.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailDocumenten.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailDocumenten.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailDocumenten.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailDocumenten.Open
    
    'Zelf
    g_rstDetailZelf.CursorLocation = adUseClient
    
    g_rstDetailZelf.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailZelf.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailZelf.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailZelf.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailZelf.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailZelf.Open
    
    'Container
    g_rstDetailContainer.CursorLocation = adUseClient
    
    g_rstDetailContainer.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailContainer.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailContainer.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailContainer.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailContainer.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailContainer.Open
    
    'Sensitive Goods
    g_rstDetailSensitiveGoods.CursorLocation = adUseClient
    
    g_rstDetailSensitiveGoods.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    g_rstDetailSensitiveGoods.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailSensitiveGoods.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    g_rstDetailSensitiveGoods.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    g_rstDetailSensitiveGoods.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    g_rstDetailSensitiveGoods.Open
    
ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            CallingForm.DigiSignAddToTrace "Function: InitializeRecordsToDigiSign"
            CallingForm.DigiSignAddToTrace "Error: " & Err.Number & " - " & Err.Description
            
    End Select

End Function


Public Function GetFieldOrdinal(ByVal FieldName As String, _
                                ByVal GroupIndex As Long) As String
    
    Dim lngPos As Long
    Dim lngStartOfOrdinalPos As Long
    Dim lngEndOfOrdinalPos As Long
    
    Dim strTemp As String
    
    GetFieldOrdinal = vbNullString
    
    Select Case FieldName
        Case "C7", "N5", "Q5"
            'C7[1,1] / N5[1,1]
            lngPos = InStr(1, g_strDigiSignFieldVerifier, FieldName)
            
            If lngPos > 0 Then
                lngStartOfOrdinalPos = InStr(lngPos, g_strDigiSignFieldVerifier, "@") + 1
                lngEndOfOrdinalPos = InStr(lngStartOfOrdinalPos, g_strDigiSignFieldVerifier, "*") - 1
                strTemp = Mid(g_strDigiSignFieldVerifier, lngStartOfOrdinalPos, lngEndOfOrdinalPos - lngStartOfOrdinalPos + 1)
                GetFieldOrdinal = strTemp
            End If
                        
            'C7[2,1] / N5[2,1]
            lngPos = InStr(lngEndOfOrdinalPos, g_strDigiSignFieldVerifier, FieldName)

            If lngPos > 0 Then
                lngStartOfOrdinalPos = InStr(lngPos, g_strDigiSignFieldVerifier, "@") + 1
                lngEndOfOrdinalPos = InStr(lngStartOfOrdinalPos, g_strDigiSignFieldVerifier, "*") - 1
                strTemp = Mid(g_strDigiSignFieldVerifier, lngStartOfOrdinalPos, lngEndOfOrdinalPos - lngStartOfOrdinalPos + 1)
                GetFieldOrdinal = GetFieldOrdinal & "," & strTemp
            End If
            
        Case "R5", "R6"
            lngPos = InStr(1, g_strDigiSignFieldVerifier, FieldName)
            
            Do While lngPos > 0
                lngStartOfOrdinalPos = InStr(lngPos, g_strDigiSignFieldVerifier, "@") + 1
                lngEndOfOrdinalPos = InStr(lngStartOfOrdinalPos, g_strDigiSignFieldVerifier, "*") - 1
                strTemp = Mid(g_strDigiSignFieldVerifier, lngStartOfOrdinalPos, lngEndOfOrdinalPos - lngStartOfOrdinalPos + 1)
                
                If Len(GetFieldOrdinal) > 0 Then
                    GetFieldOrdinal = GetFieldOrdinal & "," & strTemp
                Else
                    GetFieldOrdinal = strTemp
                End If
                
                lngPos = InStr(lngEndOfOrdinalPos, g_strDigiSignFieldVerifier, FieldName)
            Loop
            
        Case Else
            If GroupIndex > 0 Then
                lngPos = InStr(1, g_strDigiSignFieldVerifier, FieldName & "(" & GroupIndex & ")")
            Else
                lngPos = InStr(1, g_strDigiSignFieldVerifier, FieldName)
            End If
            
            If lngPos > 0 Then
                lngStartOfOrdinalPos = InStr(lngPos, g_strDigiSignFieldVerifier, "@") + 1
                lngEndOfOrdinalPos = InStr(lngStartOfOrdinalPos, g_strDigiSignFieldVerifier, "*") - 1
                strTemp = Mid(g_strDigiSignFieldVerifier, lngStartOfOrdinalPos, lngEndOfOrdinalPos - lngStartOfOrdinalPos + 1)
                GetFieldOrdinal = strTemp
            End If
    
    End Select
    
End Function

'CSCLP-95
'Added rstTemp for cloning the recordset
Public Function GetFieldValueAndDetailNumber(ByRef Value As String, _
                                             ByRef DetailNumber As Long, _
                                             ByVal FieldOrdinal As Long, _
                                             ByVal RecordsetToUse As ADODB.Recordset, _
                                             ByVal FieldIsDetail As Boolean, _
                                             ByVal GroupOrdinal As Long, _
                                             ByRef TableName As String)

    Dim strFieldName As String
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = RecordsetToUse.Clone

    g_rstDigiSign.Filter = adFilterNone
    g_rstDigiSign.Filter = "Ordinal = " & FieldOrdinal
    
    If FieldIsDetail = True Then
        rstTemp.Filter = adFilterNone
        rstTemp.Filter = "Detail  = " & DetailNumber
    End If
    
    If g_rstDigiSign.RecordCount = 1 Then
        
        g_rstDigiSign.MoveFirst
        
        'Field Value
        Select Case Trim$(g_rstDigiSign.Fields("SourceTable").Value)
            Case "Procedure"
                Value = GetProcedureData(g_rstDigiSign.Fields("Fields").Value, _
                                         rstTemp, _
                                         DetailNumber, _
                                         g_rstDigiSign.Fields("Datatype").Value)
            Case Else
                If g_rstDigiSign.Fields("Fields").Value = "R5" Then
                    Value = GetR5Data(FieldOrdinal, _
                                      rstTemp, _
                                      g_rstDigiSign.Fields("Datatype").Value)
                    
                ElseIf g_rstDigiSign.Fields("Fields").Value = "R6" Then
                    Value = GetR6Data(FieldOrdinal, _
                                      rstTemp, _
                                      g_rstDigiSign.Fields("Datatype").Value)
                
                ElseIf g_rstDigiSign.Fields("Fields").Value = "SD" Then
                    Value = GetSDData(rstTemp, _
                                      g_rstDigiSign.Fields("Datatype").Value)
                
                Else
                    Value = GetData(g_rstDigiSign.Fields("Fields").Value, _
                                    g_rstDigiSign.Fields("Datatype").Value, _
                                    rstTemp, _
                                    GroupOrdinal)
                End If
                                
        End Select
        
        'Detail Number
        If FieldIsDetail = False Then
            DetailNumber = g_rstDigiSign.Fields("DetailNumber").Value
        End If
        
        'Digital Table
        TableName = g_rstDigiSign.Fields("SourceTable").Value
    End If
    
    ADORecordsetClose rstTemp
    
End Function


Public Function GetData(ByVal sField As String, _
                        ByVal FieldType As Long, _
                        ByVal RecordsetToUse As ADODB.Recordset, _
                        ByVal GroupOrdinal As Long) As String
                                
    Dim tmpFieldData As String
    Dim arrFields() As String
    Dim lngCtr As Long
    Dim tmpSearchField As String
    Dim strOrdinal As String
    Dim strSQL As String
    
    Dim lngGroupCtr As Long
    Dim lngDataPos As Long
    Dim lngDataLength As Long
    
    Dim strFieldName As String
    
    tmpFieldData = vbNullString
    
    On Error Resume Next
    Dim x As Long
    '>>
    For x = 0 To RecordsetToUse.Fields.Count - 1
        If UCase(RecordsetToUse.Fields(x).Name) = UCase("Ordinal") Then
            RecordsetToUse.Sort = "Ordinal ASC"
        End If
    Next x
    '>>
    'original
    'RecordsetToUse.Sort = "Ordinal ASC"
    
    On Error GoTo 0
    
    'No need to filter for the Code field since the data is already filtered
    RecordsetToUse.MoveFirst
    
    'Check if there are records in the recordset
    If RecordsetToUse.RecordCount <= 0 Then Exit Function
    
    ' Check if multiple fields
    If InStr(Trim(sField), ";") > 0 Then
        
        arrFields = Split(sField, ";")
        
        For lngCtr = 0 To UBound(arrFields)
            
            ' Check if field has specific ordinal id
            If InStr(Trim(arrFields(lngCtr)), "(") > 0 Then
            
                tmpSearchField = Left(Trim(arrFields(lngCtr)), InStr(Trim(arrFields(lngCtr)), "(") - 1)
                strOrdinal = Mid(Trim(arrFields(lngCtr)), InStr(Trim(arrFields(lngCtr)), "(") + 1)
                strOrdinal = Replace(strOrdinal, ")", "")
                                
                Do While Not RecordsetToUse.EOF
                    If CStr(FNullField(RecordsetToUse.Fields("Ordinal").Value)) = strOrdinal Then
                        'CSCLP-222
                        'From Allan
                        If Not IsNull(RecordsetToUse.Fields(tmpSearchField).Value) Then
                        'tmpFieldData = tmpFieldData & _
                                       IIf(IsNull(RecordsetToUse.Fields(tmpSearchField).Value), "", _
                                       FormatData(Trim(RecordsetToUse.Fields(tmpSearchField).Value), FieldType))
                            tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(tmpSearchField).Value), FieldType)
                        End If
                    End If
                    RecordsetToUse.MoveNext
                Loop
            
            Else
                'Loop thru all the records in the case of group fields
                Do While Not RecordsetToUse.EOF
                    If GroupOrdinal = 0 Then
                    
                        'CSCLP-222
                        'From Allan
                        If Not IsNull(RecordsetToUse.Fields(arrFields(lngCtr)).Value) Then
                        'tmpFieldData = tmpFieldData & _
                                       IIf(IsNull(RecordsetToUse.Fields(arrFields(lngCtr)).Value), "", _
                                       FormatData(Trim(RecordsetToUse.Fields(arrFields(lngCtr)).Value), FieldType))
                            tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(arrFields(lngCtr)).Value), FieldType)
                        End If
                        
                    Else
                        If CStr(FNullField(RecordsetToUse.Fields("Ordinal").Value)) = GroupOrdinal Then
                            'CSCLP-222
                            'From Allan
                            If Not IsNull(RecordsetToUse.Fields(arrFields(lngCtr)).Value) Then
                            'tmpFieldData = tmpFieldData & _
                                           IIf(IsNull(RecordsetToUse.Fields(arrFields(lngCtr)).Value), "", _
                                           FormatData(Trim(RecordsetToUse.Fields(arrFields(lngCtr)).Value), FieldType))
                                tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(arrFields(lngCtr)).Value), FieldType)
                            End If
                        End If
                    End If
                    
                    RecordsetToUse.MoveNext
                Loop
                
            End If
                    
        Next
        
    Else
        
        ' Check if field has specific ordinal id
        If InStr(Trim(sField), "(") > 0 Then
                        
            tmpSearchField = Left(Trim(sField), InStr(Trim(sField), "(") - 1)
            strOrdinal = Mid(Trim(sField), InStr(Trim(sField), "(") + 1)
            strOrdinal = Replace(strOrdinal, ")", "")
                        
            'Loop thru all the records in the case of group fields
            Do While Not RecordsetToUse.EOF
                If CStr(FNullField(RecordsetToUse.Fields("Ordinal").Value)) = strOrdinal Then
                    'CSCLP-222
                    'From Allan
                    If Not IsNull(RecordsetToUse.Fields(tmpSearchField).Value) Then
                    'tmpFieldData = tmpFieldData & _
                                   IIf(IsNull(RecordsetToUse.Fields(tmpSearchField).Value), "", _
                                   FormatData(Trim(RecordsetToUse.Fields(tmpSearchField).Value), FieldType))
                        tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(tmpSearchField).Value), FieldType)
                    End If
                End If
                
                RecordsetToUse.MoveNext
            Loop
            
        'Check if the field has a specific part of the data to use
        ElseIf InStr(Trim(sField), "[") > 0 Then
            
            lngDataPos = Trim(Mid(Trim(sField), InStr(Trim(sField), "[") + 1, (InStr(Trim(sField), ",") - 1) - (InStr(Trim(sField), "[") + 1) + 1))
            lngDataLength = Trim(Mid(Trim(sField), InStr(Trim(sField), ",") + 1, (InStr(Trim(sField), "]") - 1) - (InStr(Trim(sField), ",") + 1) + 1))
            
            strFieldName = Mid(sField, 1, 2)
            
            'Loop thru all the records in the case of group fields
            Do While Not RecordsetToUse.EOF
                If GroupOrdinal = 0 Then
                    'CSCLP-222
                    'From Allan
                    If Not IsNull(RecordsetToUse.Fields(strFieldName).Value) Then
                    'tmpFieldData = tmpFieldData & _
                                   IIf(IsNull(RecordsetToUse.Fields(strFieldName).Value), "", _
                                   FormatData(Trim(RecordsetToUse.Fields(strFieldName).Value), FieldType, lngDataPos, lngDataLength))
                        tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(strFieldName).Value), FieldType, lngDataPos, lngDataLength)
                    End If
                Else
                    If CStr(FNullField(RecordsetToUse.Fields("Ordinal").Value)) = GroupOrdinal Then
                        'CSCLP-222
                        'From Allan
                        If Not IsNull(RecordsetToUse.Fields(strFieldName).Value) Then
                        'tmpFieldData = tmpFieldData & _
                                       IIf(IsNull(RecordsetToUse.Fields(strFieldName).Value), "", _
                                       FormatData(Trim(RecordsetToUse.Fields(strFieldName).Value), FieldType, lngDataPos, lngDataLength))
                            tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(strFieldName).Value), FieldType, lngDataPos, lngDataLength)
                        End If
                    End If
                    
                End If
                
                RecordsetToUse.MoveNext
            Loop
        Else
            'Loop thru all the records in the case of group fields
            Do While Not RecordsetToUse.EOF
                If GroupOrdinal = 0 Then
                    'CSCLP-222
                    'From Allan
                    If Not IsNull(RecordsetToUse.Fields(sField).Value) Then
                    'tmpFieldData = tmpFieldData & _
                                   IIf(IsNull(RecordsetToUse.Fields(sField).Value), "", _
                                   FormatData(Trim(RecordsetToUse.Fields(sField).Value), FieldType))
                         tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(sField).Value), FieldType)
                    End If
                Else
                    If CStr(FNullField(RecordsetToUse.Fields("Ordinal").Value)) = GroupOrdinal Then
                        'CSCLP-222
                        'From Allan
                        If Not IsNull(RecordsetToUse.Fields(sField).Value) Then
                        'tmpFieldData = tmpFieldData & _
                                       IIf(IsNull(RecordsetToUse.Fields(sField).Value), "", _
                                       FormatData(Trim(RecordsetToUse.Fields(sField).Value), FieldType))
                             tmpFieldData = tmpFieldData & FormatData(Trim(RecordsetToUse.Fields(sField).Value), FieldType)
                        End If
                    End If
                    
                End If
                
                RecordsetToUse.MoveNext
            Loop
            
        End If
        
    
    End If
                                                                                
    GetData = Trim(tmpFieldData)
    
    'Destroy Copy of Temporary Recordset
    ADORecordsetClose RecordsetToUse
    
End Function


Public Function FormatData(ByVal sData As String, _
                           ByVal FieldType As Long, _
                  Optional ByVal DataPos As Long, _
                  Optional ByVal DataLength As Long) As String
    
    Dim strNewData As String
    Dim strDecimalData As String
    Dim strWholeNumberData As String
    
    strNewData = vbNullString
    strDecimalData = vbNullString
    strWholeNumberData = vbNullString
    
    Select Case FieldType
        Case 1 '"Amount"
        
            If IsNumeric(sData) Then
                'If no hyphen found
                If InStr(1, sData, Chr(45)) = 0 Then
                    If Trim(sData) <> "" Then
                        sData = Replace(sData, GetRegionalSetting(LOCALE_STHOUSAND), "")
                        
                        'Check if decimal separator is already present
                        If InStr(1, sData, GetRegionalSetting(LOCALE_SDECIMAL)) > 0 Then
                                                                                
                            'Parse data for zero padding
                            strWholeNumberData = Mid(Trim(sData), 1, InStr(1, sData, GetRegionalSetting(LOCALE_SDECIMAL)) - 1)
                            strDecimalData = Mid(Trim(sData), InStrRev(sData, GetRegionalSetting(LOCALE_SDECIMAL)) + 1, Len(sData))
                            
                            strWholeNumberData = Pad(48, strWholeNumberData, 12, enuLead)
                            
                            'if decimal value is more than 8 chars, get the first 8 chars only
                            strDecimalData = IIf(Len(strDecimalData) > 8, Left$(strDecimalData, 8), Trim(strDecimalData))
                            strDecimalData = Pad(48, strDecimalData, 8, enuTrail)
                            
                            strNewData = strWholeNumberData & "," & strDecimalData
                        Else
                            strNewData = Pad(48, sData, 12, enuLead) & "," & String(8, "0")
                        End If
                    End If
                End If
            End If
                                                                                    
        Case 2 '"Date"
                        
            sData = Mid(sData, 1, 4) & "/" & Mid(sData, 5, 2) & "/" & Mid(sData, 7, 2)
            If IsDate(sData) Then
                strNewData = Format(sData, "yyyyMMdd")
            End If
        
        Case 3 '"Text"
            
            If DataPos > 0 And DataLength > 0 Then
                strNewData = Mid(sData, DataPos, DataLength)
                
            Else
                strNewData = Trim(sData)
                
            End If
            
    End Select
    
    FormatData = strNewData

End Function


Public Function GetProcedureData(ByVal Field As String, _
                                 ByVal RecordsetToUse As ADODB.Recordset, _
                                 ByVal DetailNumber As Long, _
                                 ByVal FieldType As Long) As String
    
    Dim lngCtr As Long
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = RecordsetToUse.Clone
    
    Select Case UCase$(Trim(Field))
        Case "DETAIL TAB (TARIFF LINE) NUMBER"
            GetProcedureData = FormatData(Trim(DetailNumber), FieldType)
            
        Case "TOTAL NUMBER OF SEALS"
            If rstTemp.RecordCount > 0 Then
                rstTemp.MoveFirst
                
                Do While Not rstTemp.EOF
                    'If E1 or E2 is filled count Seal
                    If (Not IsNull(rstTemp.Fields("E1").Value) And _
                       Trim$(rstTemp.Fields("E1").Value) <> vbNullString) Or _
                       (Not IsNull(rstTemp.Fields("E2").Value) And _
                       Trim$(rstTemp.Fields("E2").Value) <> vbNullString) Then
                               
                        lngCtr = lngCtr + 1
                               
                    End If
                    
                    rstTemp.MoveNext
                Loop
            End If
            
            GetProcedureData = FormatData(Trim(lngCtr), FieldType)
            
        Case "TOTAL NUMBER OF DETAILS"
            GetProcedureData = FormatData(Trim(rstTemp.RecordCount), FieldType)
        
        Case Else
            Debug.Assert False
        
    End Select
    
    'Destroy Temporary Copy of Recordset
    ADORecordsetClose rstTemp
    
End Function


Public Function GetR5Data(ByVal Ordinal As Long, _
                          ByVal RecordsetToUse As ADODB.Recordset, _
                          ByVal FieldType As Long) As String
    
    Dim strTemp As String
    Dim strR5Data As String
    
    strTemp = Trim$(FNullField(RecordsetToUse.Fields("R5").Value))
    
    'Check if there is a valid data
    If strTemp = vbNullString Then Exit Function
    
    If LenB(Trim$(FNullField(RecordsetToUse.Fields("R1").Value))) > 0 And _
       LenB(Trim$(FNullField(RecordsetToUse.Fields("R2").Value))) > 0 Then
    
        Select Case g_lngDType
            Case 14
                Select Case Ordinal
                    Case 167
                        'SG37/DOC/1004 - 998 + 705 Pos 15-20
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 20 Then
                                strR5Data = Mid(strTemp, 15, 6)
                            End If
                        End If
                        
                    Case 168
                        'SG37/DOC/1004 - 998 + 703 Magcode
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "703" Then
                            strR5Data = strTemp
                        End If
                        
                    Case 169
                        'SG37/DOC/1004 - 998 + 705 Pos 7-14
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 14 Then
                                strR5Data = Mid(strTemp, 7, 8)
                            End If
                        End If
                        
                    Case 170
                        'SG37/DOC/1004 - 998 + 705 Pos 1-6
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 6 Then
                                strR5Data = Left(strTemp, 6)
                            End If
                        End If
                        
                    Case 171
                        'SG37/DOC/1004 - 998 + 705 Pos 6-11
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 11 Then
                                strR5Data = Mid(strTemp, 6, 6)
                            End If
                        End If
                        
                    Case 172
                        'SG37/DOC/1004 - 998 + 703 Number
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "703" Then
                            strR5Data = strTemp
                        End If
                        
                    Case 173
                        'SG37/DOC/1004 - 998 + other
                        strR5Data = strTemp
                        
                    Case Else
                        Debug.Assert False
                        
                End Select
                
            Case 18
                Select Case Ordinal
                    Case 170
                        'SG37/DOC/1004 - 998 + 705 Pos 15-20
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 20 Then
                                strR5Data = Mid(strTemp, 15, 6)
                            End If
                        End If
                        
                    Case 171
                        'SG37/DOC/1004 - 998 + 703 Magcode
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "703" Then
                            strR5Data = strTemp
                        End If
                        
                    Case 172
                        'SG37/DOC/1004 - 998 + 705 Pos 7-14
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 14 Then
                                strR5Data = Mid(strTemp, 7, 8)
                            End If
                        End If
                        
                    Case 173
                        'SG37/DOC/1004 - 998 + 705 Pos 1-6
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 6 Then
                                strR5Data = Left(strTemp, 6)
                            End If
                        End If
                        
                    Case 174
                        'SG37/DOC/1004 - 998 + 705 Pos 6-11
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 11 Then
                                strR5Data = Mid(strTemp, 6, 6)
                            End If
                        End If
                        
                    Case 175
                        'SG37/DOC/1004 - 998 + 740 Pos 1-5
                        'SG37/DOC/1004 - 998 + 703 Number
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "703" Then
                            strR5Data = strTemp
                        End If
                        
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "740" Then
                            If LenB(strTemp) >= 5 Then
                                strR5Data = Left(strTemp, 5)
                            End If
                        End If
                        
                    Case 176
                        'SG37/DOC/1004 - 998 + other
                        strR5Data = strTemp
                        
                    Case Else
                        Debug.Assert False
                        
                End Select
            
            Case Else
                Debug.Assert False
                
        End Select
        
    End If
    
    If Trim(strR5Data) = vbNullString Then
        strR5Data = strTemp
    End If
    
    'Return R5 Data
    GetR5Data = FormatData(Trim(strR5Data), FieldType)
    
End Function


Public Function GetR6Data(ByVal Ordinal As Long, _
                          ByVal RecordsetToUse As ADODB.Recordset, _
                          ByVal FieldType As Long) As String
    
    Dim strTemp As String
    Dim strR6Data As String
    
    strTemp = Trim$(FNullField(RecordsetToUse.Fields("R6").Value))
    
    'Check if there is a valid data
    If strTemp = vbNullString Then Exit Function
    
    If LenB(Trim$(FNullField(RecordsetToUse.Fields("R1").Value))) > 0 And _
       LenB(Trim$(FNullField(RecordsetToUse.Fields("R2").Value))) > 0 Then
    
        Select Case g_lngDType
            Case 14
                Select Case Ordinal
                    Case 175
                        'SG37/DOC/1366 - 998 + 705 pos 1-4
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 4 Then
                                strR6Data = Left(strTemp, 4)
                            End If
                        End If
                        
                    Case 179
                        'SG37/DOC/1366 - 998 + 705 pos23-26
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 26 Then
                                strR6Data = Mid(strTemp, 23, 4)
                            End If
                        Else
                            strR6Data = strTemp 'Temporary Solution for Segment 179(Import)
                            Debug.Assert False
                        End If
                        
                    Case Else
                        Debug.Assert False
                        
                End Select
                
            Case 18
                Select Case Ordinal
                    Case 178
                        'SG37/DOC/1366 - 998 + 705 pos 1-4
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 4 Then
                                strR6Data = Left(strTemp, 4)
                            End If
                        End If
                        
                    Case 182
                        'SG37/DOC/1366 - 998 + 705 pos23-26
                        If Trim$(FNullField(RecordsetToUse.Fields("R2").Value)) = "705" Then
                            If LenB(strTemp) >= 26 Then
                                strR6Data = Mid(strTemp, 23, 4)
                            End If
                        End If
                    
                    Case Else
                        Debug.Assert False
                        
                End Select
            
            Case Else
                Debug.Assert False
                
        End Select
        
    End If
    
    'Return R6 Data
    GetR6Data = FormatData(Trim(strR6Data), FieldType)
    
End Function


Public Function GetSDData(ByVal RecordsetToUse As ADODB.Recordset, _
                          ByVal FieldType As Long) As String
    
    Dim strTemp As String
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = RecordsetToUse.Clone
    
    With rstTemp
        .Sort = "Ordinal ASC"
        
        .MoveFirst
                    
        Do While Not .EOF
            If LenB(Trim$(FNullField(.Fields("SD").Value))) > 0 Then
                strTemp = Trim$(.Fields("SD").Value)
                
                Exit Do
            End If
            
            If .Fields("SE").Value = "E" Then Exit Do
            
            .MoveNext
        Loop
    End With
    
    'Return SD Data
    GetSDData = FormatData(Trim(strTemp), FieldType)
    
End Function


Public Function BuildWholeDigitalSignature(ByVal DType As Long)
    
    Dim rstDigiSignData As New ADODB.Recordset
    Dim lngDetailNumber As Long
    Dim lngCtr As Long
    
    'Header and Detail
    rstDigiSignData.CursorLocation = adUseClient
    
    rstDigiSignData.Fields.Append "AutoID", adBigInt, , adFldIsNullable
    rstDigiSignData.Fields.Append "GroupOrdinal", adBigInt, , adFldIsNullable
    rstDigiSignData.Fields.Append "DigiSignOrdinal", adBigInt, , adFldIsNullable
    rstDigiSignData.Fields.Append "DetailNumber", adBigInt, , adFldIsNullable
    rstDigiSignData.Fields.Append "Value", adVarWChar, 1000, adFldIsNullable
    rstDigiSignData.Open
    
    '*********************************************************************************
    'Build Recordset for Digital Signature
    '*********************************************************************************
    Select Case DType
        Case 18
            '1-12
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal < 13"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
            
            '13
            g_rstHeaderTransitOffices.Sort = "GroupOrdinal, DigiSignOrdinal"
                
            If Not (g_rstHeaderTransitOffices.EOF And g_rstHeaderTransitOffices.BOF) Then
                g_rstHeaderTransitOffices.MoveFirst
            End If
                
            Do While Not g_rstHeaderTransitOffices.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstHeaderTransitOffices.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstHeaderTransitOffices.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstHeaderTransitOffices.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstHeaderTransitOffices.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstHeaderTransitOffices.MoveNext
            Loop
            
            '14
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal = 14"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
            
            '15-16
            g_rstHeaderSeals.Sort = "GroupOrdinal, DigiSignOrdinal"
                
            If Not (g_rstHeaderSeals.EOF And g_rstHeaderSeals.BOF) Then
                g_rstHeaderSeals.MoveFirst
            End If
                
            Do While Not g_rstHeaderSeals.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstHeaderSeals.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstHeaderSeals.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstHeaderSeals.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstHeaderSeals.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstHeaderSeals.MoveNext
            Loop
            
            '17-20
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal > 16 AND DigiSignOrdinal < 21"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
            
            '21-25
            g_rstHeaderGuarantee.Sort = "GroupOrdinal, DigiSignOrdinal"
                
            If Not (g_rstHeaderGuarantee.EOF And g_rstHeaderGuarantee.BOF) Then
                g_rstHeaderGuarantee.MoveFirst
            End If
                
            Do While Not g_rstHeaderGuarantee.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstHeaderGuarantee.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstHeaderGuarantee.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstHeaderGuarantee.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstHeaderGuarantee.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstHeaderGuarantee.MoveNext
            Loop
            
            '26-50
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal > 25 AND DigiSignOrdinal < 51"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
            
            'Detail Count
            g_rstDigiSignData.Filter = adFilterNone
            
            g_rstDigiSignData.Sort = "DetailNumber DESC"
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
                
                lngDetailNumber = g_rstDigiSignData.Fields("DetailNumber").Value
            Else
                Exit Function
            End If
                
            For lngCtr = 1 To lngDetailNumber
                '51 - 57
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 50 AND DigiSignOrdinal < 58 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '58 - 61
                g_rstDetailBijzondere.Filter = adFilterNone
                g_rstDetailBijzondere.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailBijzondere.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailBijzondere.EOF And g_rstDetailBijzondere.BOF) Then
                    g_rstDetailBijzondere.MoveFirst
                End If
                    
                Do While Not g_rstDetailBijzondere.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailBijzondere.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailBijzondere.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailBijzondere.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailBijzondere.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailBijzondere.MoveNext
                Loop
                
                '62 - 81
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 61 AND DigiSignOrdinal < 82 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '82
                g_rstDetailContainer.Filter = adFilterNone
                g_rstDetailContainer.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailContainer.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailContainer.EOF And g_rstDetailContainer.BOF) Then
                    g_rstDetailContainer.MoveFirst
                End If
                    
                Do While Not g_rstDetailContainer.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailContainer.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailContainer.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailContainer.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailContainer.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailContainer.MoveNext
                Loop
                
                '83-87
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 82 AND DigiSignOrdinal < 88 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '88
                g_rstDetailSensitiveGoods.Filter = adFilterNone
                g_rstDetailSensitiveGoods.Filter = "DetailNumber = " & lngCtr & " AND DigiSignOrdinal = 88"
                
                g_rstDetailSensitiveGoods.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailSensitiveGoods.EOF And g_rstDetailSensitiveGoods.BOF) Then
                    g_rstDetailSensitiveGoods.MoveFirst
                End If
                    
                Do While Not g_rstDetailSensitiveGoods.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailSensitiveGoods.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailSensitiveGoods.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailSensitiveGoods.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailSensitiveGoods.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailSensitiveGoods.MoveNext
                Loop
                
                '89 - 98
                g_rstDetailDocumenten.Filter = adFilterNone
                g_rstDetailDocumenten.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailDocumenten.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailDocumenten.EOF And g_rstDetailDocumenten.BOF) Then
                    g_rstDetailDocumenten.MoveFirst
                End If
                    
                Do While Not g_rstDetailDocumenten.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailDocumenten.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailDocumenten.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailDocumenten.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailDocumenten.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailDocumenten.MoveNext
                Loop
                
                '99-100
                g_rstDetailSensitiveGoods.Filter = adFilterNone
                g_rstDetailSensitiveGoods.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailSensitiveGoods.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailSensitiveGoods.EOF And g_rstDetailSensitiveGoods.BOF) Then
                    g_rstDetailSensitiveGoods.MoveFirst
                End If
                    
                Do While Not g_rstDetailSensitiveGoods.EOF
                    If g_rstDetailSensitiveGoods.Fields("GroupOrdinal").Value = 99 Or _
                       g_rstDetailSensitiveGoods.Fields("GroupOrdinal").Value = 100 Then
                        rstDigiSignData.AddNew
                        rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailSensitiveGoods.Fields("GroupOrdinal").Value
                        rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailSensitiveGoods.Fields("DigiSignOrdinal").Value
                        rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailSensitiveGoods.Fields("DetailNumber").Value
                        rstDigiSignData.Fields("Value").Value = g_rstDetailSensitiveGoods.Fields("Value").Value
                        rstDigiSignData.Update
                    End If
                        
                    g_rstDetailSensitiveGoods.MoveNext
                Loop
                
                '101-103
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 100 AND DigiSignOrdinal < 104 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
            Next lngCtr
            
        Case 14
            '1-9
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal < 10"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
                
            '10 - 11
            g_rstHeaderSeals.Sort = "GroupOrdinal, DigiSignOrdinal"
                
            If Not (g_rstHeaderSeals.EOF And g_rstHeaderSeals.BOF) Then
                g_rstHeaderSeals.MoveFirst
            End If
                
            Do While Not g_rstHeaderSeals.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstHeaderSeals.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstHeaderSeals.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstHeaderSeals.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstHeaderSeals.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstHeaderSeals.MoveNext
            Loop
                
            '12 - 40
            g_rstDigiSignData.Filter = adFilterNone
            g_rstDigiSignData.Filter = "DigiSignOrdinal > 11 AND DigiSignOrdinal < 41"
                
            g_rstDigiSignData.Sort = "DigiSignOrdinal"
                
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
            End If
                
            Do While Not g_rstDigiSignData.EOF
                rstDigiSignData.AddNew
                rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                rstDigiSignData.Update
                    
                g_rstDigiSignData.MoveNext
            Loop
                
            'Detail Count
            g_rstDigiSignData.Filter = adFilterNone
            
            g_rstDigiSignData.Sort = "DetailNumber DESC"
            If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                g_rstDigiSignData.MoveFirst
                
                lngDetailNumber = g_rstDigiSignData.Fields("DetailNumber").Value
            Else
                Exit Function
            End If
                
            For lngCtr = 1 To lngDetailNumber
                
                '41 - 46
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 40 AND DigiSignOrdinal < 47 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '47 - 48
                g_rstDetailBijzondere.Filter = adFilterNone
                g_rstDetailBijzondere.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailBijzondere.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailBijzondere.EOF And g_rstDetailBijzondere.BOF) Then
                    g_rstDetailBijzondere.MoveFirst
                End If
                    
                Do While Not g_rstDetailBijzondere.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailBijzondere.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailBijzondere.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailBijzondere.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailBijzondere.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailBijzondere.MoveNext
                Loop
                
                '49-62
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 48 AND DigiSignOrdinal < 63 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '63 - 64
                g_rstDetailBerekenings.Filter = adFilterNone
                g_rstDetailBerekenings.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailBerekenings.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailBerekenings.EOF And g_rstDetailBerekenings.BOF) Then
                    g_rstDetailBerekenings.MoveFirst
                End If
                    
                Do While Not g_rstDetailBerekenings.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailBerekenings.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailBerekenings.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailBerekenings.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailBerekenings.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailBerekenings.MoveNext
                Loop
                
                '65 - 83
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 64 AND DigiSignOrdinal < 84 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '84
                g_rstDetailContainer.Filter = adFilterNone
                g_rstDetailContainer.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailContainer.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailContainer.EOF And g_rstDetailContainer.BOF) Then
                    g_rstDetailContainer.MoveFirst
                End If
                    
                Do While Not g_rstDetailContainer.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailContainer.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailContainer.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailContainer.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailContainer.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailContainer.MoveNext
                Loop
                
                '85 - 88
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 84 AND DigiSignOrdinal < 89 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '89 - 98
                g_rstDetailDocumenten.Filter = adFilterNone
                g_rstDetailDocumenten.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailDocumenten.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailDocumenten.EOF And g_rstDetailDocumenten.BOF) Then
                    g_rstDetailDocumenten.MoveFirst
                End If
                    
                Do While Not g_rstDetailDocumenten.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailDocumenten.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailDocumenten.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailDocumenten.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailDocumenten.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailDocumenten.MoveNext
                Loop
                
                '99 - 104
                g_rstDigiSignData.Filter = adFilterNone
                g_rstDigiSignData.Filter = "DigiSignOrdinal > 98 AND DigiSignOrdinal < 105 AND DetailNumber = " & lngCtr
                    
                g_rstDigiSignData.Sort = "DigiSignOrdinal"
                    
                If Not (g_rstDigiSignData.EOF And g_rstDigiSignData.BOF) Then
                    g_rstDigiSignData.MoveFirst
                End If
                    
                Do While Not g_rstDigiSignData.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDigiSignData.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDigiSignData.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDigiSignData.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDigiSignData.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDigiSignData.MoveNext
                Loop
                
                '105 - 106
                g_rstDetailZelf.Filter = adFilterNone
                g_rstDetailZelf.Filter = "DetailNumber = " & lngCtr
                
                g_rstDetailZelf.Sort = "GroupOrdinal, DigiSignOrdinal"
                    
                If Not (g_rstDetailZelf.EOF And g_rstDetailZelf.BOF) Then
                    g_rstDetailZelf.MoveFirst
                End If
                    
                Do While Not g_rstDetailZelf.EOF
                    rstDigiSignData.AddNew
                    rstDigiSignData.Fields("GroupOrdinal").Value = g_rstDetailZelf.Fields("GroupOrdinal").Value
                    rstDigiSignData.Fields("DigiSignOrdinal").Value = g_rstDetailZelf.Fields("DigiSignOrdinal").Value
                    rstDigiSignData.Fields("DetailNumber").Value = g_rstDetailZelf.Fields("DetailNumber").Value
                    rstDigiSignData.Fields("Value").Value = g_rstDetailZelf.Fields("Value").Value
                    rstDigiSignData.Update
                        
                    g_rstDetailZelf.MoveNext
                Loop
                
            Next lngCtr
            
        Case Else
            Debug.Assert False
            
    End Select
    '*********************************************************************************
        
    Set g_rstDigiSignData = rstDigiSignData
        
End Function

'CSCLP-499
Public Function RoutingBoxEmpty(ByVal strBox As String, ByVal Code As String) As Boolean
    
    Dim conTemp As ADODB.Connection
    Dim rstTemp As ADODB.Recordset
    Dim strCommand As String
    Dim strTableToUse As String
    
    RoutingBoxEmpty = True
    
    Select Case strBox
        Case "AE", "AF"
            strTableToUse = "[PLDA COMBINED HEADER TRANSIT OFFICES]"
        Case "DB"
            strTableToUse = "[PLDA COMBINED HEADER]"
        Case "N7"
            strTableToUse = "[PLDA COMBINED DETAIL]"
    End Select
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT * "
    strCommand = strCommand & "FROM " & strTableToUse
    strCommand = strCommand & " WHERE "
    strCommand = strCommand & "CODE = '" & Code & "'"
        
    ADORecordsetOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, g_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
     
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.MoveFirst
        Do While Not rstTemp.EOF
            Select Case strBox
                Case "AE", "AF"
                    If Len(Trim(rstTemp.Fields("AF").Value)) <> 0 Or Len(Trim(rstTemp.Fields("AF").Value)) <> 0 Then
                        RoutingBoxEmpty = False
                        Exit Do
                    Else
                        RoutingBoxEmpty = True
                    End If
                Case Else
                    If Len(Trim(rstTemp.Fields(strBox).Value)) <> 0 Then
                        RoutingBoxEmpty = False
                        Exit Do
                    Else
                        RoutingBoxEmpty = True
                    End If
            End Select
            rstTemp.MoveNext
        Loop
    End If
    
    ADORecordsetClose rstTemp
End Function

Public Function GetTableToUpdateFromActiveCommand(ByRef ActiveRecordset As ADODB.Recordset) As String
    Dim strActiveCommand As String
    Dim strTableToUpdate As String
    Dim lngPos As Long
    Dim blnSquareBracket As Boolean
    
    strActiveCommand = UCase$(Trim$(ActiveRecordset.ActiveCommand))
    
    lngPos = InStr(1, strActiveCommand, "FROM")
    
    If lngPos > 0 Then
        blnSquareBracket = False
        
        lngPos = lngPos + 4
        
        Do While Mid(strActiveCommand, lngPos) = " "
            
            lngPos = lngPos + 1
        Loop
        
        blnSquareBracket = (Mid(strActiveCommand, lngPos) = "[")
        
        If blnSquareBracket Then
            strTableToUpdate = Mid(strActiveCommand, lngPos + 1)
            strTableToUpdate = Left$(strTableToUpdate, InStr(1, strTableToUpdate, "]") - 1)
        Else
            strTableToUpdate = Mid(strActiveCommand, lngPos)
            strTableToUpdate = Left$(strTableToUpdate, InStr(1, strTableToUpdate, " ") - 1)
        End If
    End If
    
    GetTableToUpdateFromActiveCommand = strTableToUpdate
End Function
