VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConversion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Sub GetSystemSeparators(DecSep As String, DigSep As String)
    Dim dblSample As Double
    Dim strDecimalSeparator As String
    Dim strTemp As String
    
    dblSample = 1234.56
    strTemp = Format(dblSample, "Standard")
    
    DigSep = Mid(strTemp, 2, 1)
    DecSep = Mid(strTemp, 6, 1)
End Sub

Private Function MyCDbl(ByVal InputVal As Variant) As Double
    On Error GoTo modMyCDbl_MyDoubleFormat_ErrorHandler
    
    Dim strNumber As String
    Dim strLeftOfDecimal As String
    Dim strDecimalPart As String
    Dim strWholePart As String
    Dim strDecSeparator As String
    Dim strDigSeparator As String
    Dim sglSegments As Single
    Dim strSegment As String
    Dim intSegmentCtr As Integer
    
    Const EVEN_DIGITS = "24680"
    
1000    strWholePart = ""
1010    Call GetSystemSeparators(strDecSeparator, strDigSeparator)
1020    strNumber = CStr(InputVal)
1030    Call RemoveCharFromStr(strNumber, strDigSeparator)
1040    If InStr(1, strNumber, strDecSeparator) > 0 Then
1050        strLeftOfDecimal = Left(strNumber, InStr(1, strNumber, strDecSeparator) - 1)
1060        strDecimalPart = Mid(strNumber, InStr(1, strNumber, strDecSeparator) + 1)
1070    ElseIf InStr(1, strNumber, strDecSeparator) = 1 Then
1080        strLeftOfDecimal = "0"
1090        strDecimalPart = Mid(strNumber, InStr(1, strNumber, strDecSeparator) + 1)
1100    Else
1110        strLeftOfDecimal = strNumber
1120        strDecimalPart = "00"
1130    End If
        
1140    sglSegments = Len(strLeftOfDecimal) \ 3
1150    If Len(strLeftOfDecimal) Mod 3 > 0 Then
1160        sglSegments = sglSegments + 1
1170    End If
    
1180    For intSegmentCtr = 1 To sglSegments
1190        If intSegmentCtr = sglSegments Then
1200            strWholePart = strLeftOfDecimal & strWholePart
1210        Else
1220            strSegment = Right(strLeftOfDecimal, 3)
1230            strLeftOfDecimal = Left(strLeftOfDecimal, Len(strLeftOfDecimal) - 3)
1240            strWholePart = strDigSeparator & strSegment & strWholePart
1250        End If
1260    Next
        
        '----->  3
3000    Select Case Len(strDecimalPart)
            Case 0
3010            strDecimalPart = "00"
            Case 1
3120            strDecimalPart = strDecimalPart & "0"
            Case 2
3130            strDecimalPart = strDecimalPart
            Case Else
3140            If InStr(1, EVEN_DIGITS, Mid(strDecimalPart, 2, 1)) > 0 Then
3150                If Val(Mid(strDecimalPart, 3, 1)) > 5 Then
3160                    strDecimalPart = Left(strDecimalPart, 1) & CStr(Val(Mid(strDecimalPart, 2, 1)) + 1)
3170                Else
3180                    strDecimalPart = Left(strDecimalPart, 1) & Mid(strDecimalPart, 2, 1)
3190                End If
3210            Else
3220                If Val(Mid(strDecimalPart, 3, 1)) >= 5 Then
3230                    strDecimalPart = Left(strDecimalPart, 1) & CStr(Val(Mid(strDecimalPart, 2, 1)) + 1)
3240                Else
3250                    strDecimalPart = Left(strDecimalPart, 1) & Mid(strDecimalPart, 2, 1)
3260                End If
                End If
        End Select
        MyCDbl = CDbl(strWholePart & strDecSeparator & strDecimalPart)
    
    On Error GoTo 0
    Exit Function
    
modMyCDbl_MyDoubleFormat_ErrorHandler:
    Err.Clear
    On Error GoTo 0
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    
End Function

Private Sub RemoveCharFromStr(strSource As String, Char As String)
    On Error GoTo modMyCDbl_RemoveCharFromStr_ErrorHandler
    
    Dim strTemp As String
    Dim i As Integer
    
1000    strTemp = ""
1010    For i = 1 To Len(strSource)
1020        If Mid(strSource, i, 1) <> Char Then
1030            strTemp = strTemp & Mid(strSource, i, 1)
1040        End If
1050    Next i
1060    strSource = strTemp
    
    On Error GoTo 0
    
    Exit Sub
    
modMyCDbl_RemoveCharFromStr_ErrorHandler:
    Err.Clear
    On Error GoTo 0
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    
End Sub

