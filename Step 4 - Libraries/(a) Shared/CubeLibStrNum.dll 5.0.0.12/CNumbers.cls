VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumbers"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function GetDecimalSeparator() As String
    
    Dim lngLCID As Long
    Dim strCodePage As String
    
    ' Use user locale default settings still for the separators.
    lngLCID = GetThreadLocale()
    
    ' ********** Added August 10, 2000 **********
    strCodePage = Space(16)
    
    GetLocaleInfo lngLCID, LOCALE_SDECIMAL Or LOCALE_NOUSEROVERRIDE, strCodePage, Len(strCodePage)
    
    strCodePage = StripNullTerminator_F(strCodePage)
    
    GetDecimalSeparator = strCodePage
    
End Function

'Public Sub GetSystemSeparators(DecSep As String, DigSep As String)
Public Sub GetNumericSeparators(DecSep As String, DigSep As String)
    Dim dblSample As Double
    Dim strDecimalSeparator As String
    Dim strTemp As String
    
    dblSample = 1234.56
    strTemp = Format(dblSample, "Standard")
    
    DigSep = Mid(strTemp, 2, 1)
    DecSep = Mid(strTemp, 6, 1)
End Sub

Private Function IsExponential(strNum As String) As Boolean
    IsExponential = False
    If IsNumeric(strNum) And InStr(UCase(strNum), "E") Then
        IsExponential = True
    End If
End Function

Private Sub RemoveCharFromStr(strSource As String, Char As String)
    Dim strTemp As String
    Dim i As Integer
    
    strTemp = ""
    For i = 1 To Len(strSource)
        If Mid(strSource, i, 1) <> Char Then
            strTemp = strTemp & Mid(strSource, i, 1)
        End If
    Next i
    strSource = strTemp
End Sub

Public Function cpiCDbl(varNum As Variant) As Double
    Dim strDecSep As String * 1 ' Contains the Decimal Separator Character
    Dim strDigSep As String * 1 ' Contains the Digit/Thousand Separator Character
    
    Dim lStart As Long
    Dim lEnd As Long
    
    Dim iTempLen As Integer
    
    Dim strDigPart As String
    Dim strDecPart As String
    
    Dim strNum As String
    
    Dim blnExponential As Boolean
    Dim strExponent As String
    Dim strLeftExponent As String
    
    If IsNull(varNum) Then
        cpiCDbl = 0
        Exit Function
    End If
    
    strNum = CStr(varNum)
        
    strNum = Trim(strNum)
    If strNum = "" Then
        cpiCDbl = 0
        Exit Function
    End If
    
    If Not cpiIsNumeric(strNum) Then
        cpiCDbl = 0
        Exit Function
    End If
        
    ' Check if string is in exponential format
    If InStr(UCase(strNum), "E") Then
        blnExponential = True
        strLeftExponent = Left(strNum, InStr(UCase(strNum), "E") - 1)
        strExponent = Mid(strNum, InStr(UCase(strNum), "E") + 1)
    Else
        blnExponential = False
        strLeftExponent = strNum
        strExponent = 0
    End If
    
    Call GetNumericSeparators(strDecSep, strDigSep)
    
    ' Search for a Decimal Separator
    iTempLen = InStr(1, strLeftExponent, strDecSep)
    
    '  Divide number into its Decimal and Whole/Digit Parts
    If iTempLen = 0 Then 'No Decimal Separator
        strDigPart = Mid(strLeftExponent, 1)
        strDecPart = "0"
    Else 'With Decimal Separator
        strDigPart = Mid(strLeftExponent, 1, IIf(iTempLen > 0, iTempLen - 1, 0))
        strDecPart = Mid(strLeftExponent, iTempLen + 1)
    End If
    
    ' Check for any Decimal Separator in the Digit/Whole Part
    If InStr(1, strDigPart, strDecSep) <> 0 Then
        cpiCDbl = 0
        Exit Function
    End If
    
    ' Check for any Decimal Separator in the Decimal Part
    If InStr(1, strDecPart, strDecSep) <> 0 Then
        cpiCDbl = 0
        Exit Function
    End If
    
    ' Remove all Digital Separators existing in Decimal and Digit/Whole Parts
    Call RemoveCharFromStr(strDecPart, strDigSep)
    Call RemoveCharFromStr(strDigPart, strDigSep)
    
    If strDecPart = "" Then strDecPart = "0"
    If strDigPart = "" Then strDigPart = "0"
    
    On Error GoTo DOUBLE_RANGE_HANDLER
        cpiCDbl = CDbl(strDigPart)
        If CDbl(strLeftExponent) < 0 Then
            cpiCDbl = cpiCDbl - (CDbl(strDecPart) / (10 ^ Len(strDecPart)))
        Else
            cpiCDbl = cpiCDbl + (CDbl(strDecPart) / (10 ^ Len(strDecPart)))
        End If
        cpiCDbl = cpiCDbl * CDbl("1E" & strExponent)
    On Error GoTo 0

    Exit Function

DOUBLE_RANGE_HANDLER:

    Select Case Err.Number
        Case 6      ' Exceeds double data type range values
            Err.Raise Err.Number
        Case Else
            Err.Raise Err.Number
    End Select

    On Error GoTo 0
End Function

Public Function cpiIsNumeric(varNum As Variant) As Boolean
    Dim clsStrings As CStrings
    
    Dim strDecSep As String * 1 ' Contains the Decimal Separator Character
    Dim strDigSep As String * 1 ' Contains the Digit/Thousand Separator Character
            
    Dim strNum As String
    
    Dim intDecSepLoopCtr As Integer
    Dim intStartSearch As Integer
    Dim intFirstDecSep As Integer
    Dim intSecondDecSep As Integer
    
    Dim strLeftExponent As String
    Dim strRightExponent As String
    Dim intFirstExponent As Integer
    Dim intSecondExponent As Integer
    Dim intDecimalPlace As Integer
    Dim intLastDigitPlace2 As Integer
    Dim strDummy As String
                
    Set clsStrings = New CStrings
    
    If IsNull(varNum) Then
        cpiIsNumeric = False
        Exit Function
    End If
    
    strNum = CStr(varNum)
    
    Call GetNumericSeparators(strDecSep, strDigSep)
    
    
    '----->Check if in Exponential form
    If InStr(UCase(strNum), "E") Then
        strLeftExponent = Left(strNum, InStr(UCase(strNum), "E") - 1)
    Else
        strLeftExponent = strNum
    End If
    
    '----->Check for Duplicate of Decimal Separators
    intFirstDecSep = InStr(1, strLeftExponent, strDecSep)
    intDecimalPlace = intFirstDecSep
    If intFirstDecSep > 0 Then
        intSecondDecSep = InStr(intFirstDecSep + 1, strLeftExponent, strDecSep)
        If intSecondDecSep > 0 Then
            cpiIsNumeric = False
            Exit Function
        End If
    End If
    
    '----->Check for Spacing of Digit Separators
    intStartSearch = 1
    intFirstDecSep = InStr(intStartSearch, strLeftExponent, strDigSep)
    
    If intFirstDecSep >= 1 Then
        If Val(Left(strLeftExponent, intFirstDecSep)) = 0 Then
            cpiIsNumeric = False
            Exit Function
        End If
    End If
    
    If (intDecimalPlace = 0 And Left(strLeftExponent, 1) = "0") Then
        strDummy = clsStrings.TrimLeadTrail(strLeftExponent, Asc("0"), enuLead)
        If strDummy <> "" Then
            cpiIsNumeric = False
            Exit Function
        End If
    ElseIf intDecimalPlace > 0 Then
        If Val(Left(strLeftExponent, intDecimalPlace - 1)) > 0 And Left(strLeftExponent, 1) = "0" Then
            cpiIsNumeric = False
            Exit Function
        End If
    End If
    
    If intFirstDecSep > 1 Then
        intLastDigitPlace2 = intFirstDecSep
        Do While intStartSearch
            intSecondDecSep = InStr(intFirstDecSep + 1, strLeftExponent, strDigSep)
            
            If intSecondDecSep > 0 Then
                intLastDigitPlace2 = intSecondDecSep
                If intSecondDecSep - intFirstDecSep <> 4 Then
                    cpiIsNumeric = False
                    Exit Function
                End If
                intFirstDecSep = intSecondDecSep
                intStartSearch = intSecondDecSep + 1
            Else
                Exit Do
            End If
        Loop
        
        intSecondDecSep = InStr(intFirstDecSep + 1, strLeftExponent, strDecSep)
        If intSecondDecSep > 0 Then
            If intSecondDecSep - intFirstDecSep <> 4 Then
                cpiIsNumeric = False
                Exit Function
            End If
        Else
            If Len(strLeftExponent) - intLastDigitPlace2 <> 3 Then
                cpiIsNumeric = False
                Exit Function
            End If
        End If
    ElseIf intFirstDecSep = 1 Then
        cpiIsNumeric = False
        Exit Function
    End If
    
    Set clsStrings = Nothing
    
    cpiIsNumeric = IsNumeric(strLeftExponent)
End Function

Public Function IsValidValue(ByVal Quantity As String, _
                                        Optional ByVal LowerBound As String = "None", _
                                        Optional ByVal InclusiveLBound As Boolean = False, _
                                        Optional ByVal UpperBound As String = "None", _
                                        Optional ByVal InclusiveUBound As Boolean = True, _
                                        Optional ByVal PromptMessage As Boolean = False, _
                                        Optional ByVal AllowDecimal As Boolean = False, _
                                        Optional ByVal AllowNegative As Boolean = False, _
                                        Optional ByVal AllowSpace As Boolean = False) _
                                        As Boolean

On Error GoTo DOUBLE_RANGE_HANDLER

    IsValidValue = True

    If cpiIsNumeric(Quantity) = False Then
        If AllowSpace = False Then
            If PromptMessage = True Then
                MsgBox "The quantity is not a valid number.", vbInformation, "Invalid Input (3001)"
            End If
            IsValidValue = False
        ElseIf Trim(Quantity) <> "" Then
            If PromptMessage = True Then
                MsgBox "The quantity is not a valid number.", vbInformation, "Invalid Input (3002)"
            End If
            IsValidValue = False
        End If
    End If
    If AllowDecimal = False Then
        If Fix(cpiCDbl(Quantity)) <> cpiCDbl(Quantity) Then
            If PromptMessage = True Then
                MsgBox "The quantity could not contain a decimal.", vbInformation, "Invalid Input (3003)"
            End If
            IsValidValue = False
        End If
    End If
    If LowerBound <> "None" Then
        If InclusiveLBound = False Then
            If cpiCDbl(Quantity) <= cpiCDbl(LowerBound) Then
                If PromptMessage = True Then
                    MsgBox "The quantity could not be less than or equal to " & LowerBound & ".", vbInformation, "Invalid Input (3004)"
                End If
                IsValidValue = False
            End If
        Else
            If cpiCDbl(Quantity) < cpiCDbl(LowerBound) Then
                If PromptMessage = True Then
                    MsgBox "The quantity could not be less than " & LowerBound & ".", vbInformation, "Invalid Input (3005)"
                End If
                IsValidValue = False
            End If
        End If
    End If
    If UpperBound <> "None" Then
        If InclusiveUBound = True Then
            If cpiCDbl(Quantity) > cpiCDbl(UpperBound) Then
                If PromptMessage = True Then
                    MsgBox "The quantity could not be greater than or equal to " & UpperBound & ".", vbInformation, "Invalid Input (3006)"
                End If
                IsValidValue = False
            End If
        Else
            If cpiCDbl(Quantity) >= cpiCDbl(UpperBound) Then
                If PromptMessage = True Then
                    MsgBox "The quantity could not be greater than " & UpperBound & ".", vbInformation, "Invalid Input (3007)"
                End If
                IsValidValue = False
            End If
        End If
    End If

    Exit Function

DOUBLE_RANGE_HANDLER:

    Select Case Err.Number
        Case 6      ' Exceeds double data type range values
            IsValidValue = False
        Case Else
            Err.Raise Err.Number
    End Select

    On Error GoTo 0
    
End Function

Public Function MyDoubleFormat(ByVal InputVal As Variant) As String
    Dim strNumber As String
    Dim strLeftOfDecimal As String
    Dim strDecimalPart As String
    Dim strWholePart As String
    Dim strDecSeparator As String
    Dim strDigSeparator As String
    Dim sglSegments As Single
    Dim strSegment As String
    Dim strHundreds As String
    Dim intSegmentCtr As Integer
    Dim blnIsNegative As Boolean
    Dim lngCarryOverRound As Long
    
    Const EVEN_DIGITS = "24680"
    
    strWholePart = ""
    blnIsNegative = False
    
    Call GetNumericSeparators(strDecSeparator, strDigSeparator)
    
    strNumber = Trim(CStr(InputVal))
    
    Call RemoveCharFromStr(strNumber, strDigSeparator)
    
    If InStr(1, strNumber, strDecSeparator) > 0 Then
        strLeftOfDecimal = Left(strNumber, InStr(1, strNumber, strDecSeparator) - 1)
        strDecimalPart = Mid(strNumber, InStr(1, strNumber, strDecSeparator) + 1)
    ElseIf InStr(1, strNumber, strDecSeparator) = 1 Then
        strLeftOfDecimal = "0"
        strDecimalPart = Mid(strNumber, InStr(1, strNumber, strDecSeparator) + 1)
    Else
        strLeftOfDecimal = strNumber
        strDecimalPart = "00"
    End If
    
    If Len(strLeftOfDecimal) > 0 Then
        If Left(strLeftOfDecimal, 1) = "-" Then
            strLeftOfDecimal = Mid(strLeftOfDecimal, 2)
            blnIsNegative = True
        End If
    End If
    
    sglSegments = Len(strLeftOfDecimal) \ 3
    If Len(strLeftOfDecimal) Mod 3 > 0 Then
        sglSegments = sglSegments + 1
    End If
    
    For intSegmentCtr = 1 To sglSegments
        If intSegmentCtr = sglSegments Then
            strWholePart = strLeftOfDecimal & strWholePart
        Else
            strSegment = Right(strLeftOfDecimal, 3)
            strLeftOfDecimal = Left(strLeftOfDecimal, Len(strLeftOfDecimal) - 3)
            strWholePart = strDigSeparator & strSegment & strWholePart
        End If
    Next
        
    Select Case Len(strDecimalPart)
        Case 0
            strDecimalPart = "00"
        Case 1
            strDecimalPart = strDecimalPart & "0"
        Case 2
            strDecimalPart = strDecimalPart
        Case Else
            If Val(Mid(strDecimalPart, 3, 1)) >= 5 Then
                strHundreds = CStr(Val(Mid(strDecimalPart, 2, 1)) + 1)
                If strHundreds = "10" Then
                    strDecimalPart = CStr(Val(Left(strDecimalPart, 1)) + 1) & "0"
                Else
                    strDecimalPart = Left(strDecimalPart, 1) & strHundreds
                End If
            Else
                strDecimalPart = Left(strDecimalPart, 1) & Mid(strDecimalPart, 2, 1)
            End If
    End Select
    
    If Left(Trim(strWholePart), 1) = strDigSeparator Then
        strWholePart = Mid(Trim(strWholePart), 2, Len(Trim(strWholePart)))
    End If
    
    MyDoubleFormat = IIf(blnIsNegative, "-", "") & CStr(CLng(strWholePart) + lngCarryOverRound) & strDecSeparator & strDecimalPart
End Function

Public Function ConvertBase(ByVal Source As String, ByVal BaseIn As Integer, ByVal BaseOut As Integer) As String
    Dim intCharCtr As Integer
    Dim CurrentCharacter As String
    Dim CharacterValue As Integer
    Dim PlaceValue As Integer
    Dim RunningTotal As Double
    Dim Remainder As Double

    '----->Validate Input Parameters
    If Source = "" Or BaseIn < 1 Or BaseIn > 36 Or BaseOut < 1 Or BaseOut > 36 Then
        ConvertBase = "Error"
        Exit Function
    End If

    '----->Convert Source into Decimal
    PlaceValue = Len(Source)

    For intCharCtr = 1 To Len(Source)
        PlaceValue = PlaceValue - 1
        CurrentCharacter = Mid$(Source, intCharCtr, 1)
        CharacterValue = 0
        If Asc(CurrentCharacter) > 64 And _
            Asc(CurrentCharacter) < 91 Then
            
            CharacterValue = Asc(CurrentCharacter) - 55
        End If

        If CharacterValue = 0 Then
            If Asc(CurrentCharacter) < 48 Or _
                Asc(CurrentCharacter) > 57 Then
                
                '------>Ensure Source is correct
                ConvertBase = "Error"
                Exit Function
            Else
                CharacterValue = Val(CurrentCharacter)
            End If
        End If

        If CharacterValue < 0 Or _
            CharacterValue > BaseIn - 1 Then
            
            '----->Ensure Source is correct
            ConvertBase = "Error"
            Exit Function
        End If

        RunningTotal = RunningTotal + CharacterValue * (BaseIn ^ PlaceValue)
    Next intCharCtr

    '----->Convert Decimal Number into the desired base using Repeated Division
    Do
        Remainder = (RunningTotal Mod BaseOut)
        RunningTotal = (RunningTotal - Remainder) / BaseOut


        If Remainder >= 10 Then
            CurrentCharacter = Chr$(Remainder + 55)
        Else
            CurrentCharacter = Right$(str$(Remainder), Len(str$(Remainder)) - 1)
        End If

        ConvertBase = CurrentCharacter + ConvertBase
    Loop While RunningTotal > 0
End Function
