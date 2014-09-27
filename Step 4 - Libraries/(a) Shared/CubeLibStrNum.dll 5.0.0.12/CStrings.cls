VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStrings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, _
     ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Public Enum ReplaceLoc
    enuLead = 0
    enuTrail = 1
End Enum

Public ResourceHandler As Long

Public Function ReplicateString(ByVal StringToReplicate As String, ByVal HowManyTimes As Integer) As String
    Dim intCtr As Integer
    Dim strReplicatedString As String
    
    For intCtr = 1 To HowManyTimes
        strReplicatedString = strReplicatedString & StringToReplicate
    Next
    
    ReplicateString = strReplicatedString
End Function

Public Function TrimLeadTrail(ByVal Data As String, ByVal AsciiCode As Integer, Location As ReplaceLoc) As String
    Dim strData As String
    
    strData = Data
    
    Do While True
        If Len(strData) > 0 Then
            Select Case Location
                '----->Leading Characters
                Case 0
                    If AsciiCode = Asc(Left(strData, 1)) Then
                        If Len(strData) > 1 Then
                            strData = Mid(strData, 2)
                        Else
                            strData = ""
                        End If
                    Else
                        Exit Do
                    End If
                    
                '----->Trailing Characters
                Case 1
                    If AsciiCode = Asc(Right(strData, 1)) Then
                        If Len(strData) > 1 Then
                            strData = Left(strData, Len(strData) - 1)
                        Else
                            strData = ""
                        End If
                    Else
                        Exit Do
                    End If
            End Select
        Else
            Exit Do
        End If
    Loop
    
    TrimLeadTrail = strData
End Function

Public Function Pad(AsciiCode As Integer, SourceString As String, TargetLength As Byte, Location As ReplaceLoc) As String
    Dim byteLen As Byte
    Dim byteIndex As Byte
    Dim TargetString As String
    Dim strCharToAppend As String
    
    strCharToAppend = Chr(AsciiCode)
    
    TargetString = SourceString
    byteLen = Len(TargetString)
    For byteIndex = byteLen To TargetLength - 1
        If Location = enuLead Then
            TargetString = strCharToAppend & TargetString
        Else
            TargetString = TargetString & strCharToAppend
        End If
    Next byteIndex
    
    Pad = TargetString
End Function

Public Function CropMultiLine(ByVal MultiLineString As String, ByVal LineLimit As Integer, ByVal CharLimit As Integer) As String
    Dim arrLines
    Dim intLineCtr As Integer
    Dim strAccumulateLines As String
    Dim intAccumulateCtr As Integer
    Dim lngAsc As Long
    
    If InStr(1, MultiLineString, Chr(13)) = 0 Then
        CropMultiLine = IIf(Len(MultiLineString) > CharLimit, Mid(Trim(MultiLineString), 1, CharLimit), Trim(MultiLineString))
    Else
        strAccumulateLines = ""
        intAccumulateCtr = 0
        For intLineCtr = 1 To Len(MultiLineString)
            lngAsc = Asc(Mid(MultiLineString, 1, 1))
            If Mid(MultiLineString, 1, 1) = Chr(13) Or Mid(MultiLineString, 1, 1) = Chr(10) Then
                MultiLineString = Mid(MultiLineString, 2)
            Else
                Exit For
            End If
        Next
        
        arrLines = Split(MultiLineString, Chr(13) + Chr(10))
        
        For intLineCtr = 0 To UBound(arrLines)
            If Trim(CStr(arrLines(intLineCtr))) <> "" And intAccumulateCtr < LineLimit Then
                strAccumulateLines = strAccumulateLines & IIf(Len(CStr(arrLines(intLineCtr))) > CharLimit, Mid(Trim(CStr(arrLines(intLineCtr))), 1, CharLimit), Trim(CStr(arrLines(intLineCtr)))) & Chr(10)
                intAccumulateCtr = intAccumulateCtr + 1
            End If
        Next
        
        CropMultiLine = strAccumulateLines
    End If
End Function

Public Function Parse(ByVal Source As String, ByVal Delimiter As String) As Variant
    
    Dim arrElements
            
    If Trim(Source) <> "" Then
        If InStr(1, Source, Delimiter) > 0 Then
            arrElements = Split(Source, Delimiter)
        Else
            ReDim arrElements(0)
            arrElements(0) = Source
        End If
    Else
        arrElements = Empty
    End If
            
    Parse = arrElements

End Function

Public Function Encrypt(WordToEncrypt, Password As String) As String
    Dim lngI As Long
    Dim lngJ As Long
   
    Dim strWord As String
   
    If Trim(Password) = "" Then
        Encrypt = WordToEncrypt
        Exit Function
    End If
    
    For lngI = 1 To Len(WordToEncrypt)
        If lngI > Len(Password) Then
            lngJ = 1
        Else
            lngJ = lngI
        End If
      
        If Mid(WordToEncrypt, lngI, 1) <> ":" Then
            strWord = strWord + Chr(Asc(Mid(WordToEncrypt, lngI, 1)) + Asc(Mid(Password, lngJ, 1)))
        Else
            strWord = strWord + ":"
        End If
    Next
    Encrypt = strWord
    
End Function


Public Function Decrypt(ByVal WordToDEcrypt, ByVal Password As String) As String
    Dim lngI As Long
    Dim lngJ As Long
   
    Dim strWord As String
   
    If Trim(Password) = "" Then
        Decrypt = WordToDEcrypt
        Exit Function
    End If
    
    WordToDEcrypt = Trim$(WordToDEcrypt)
   
    For lngI = 1 To Len(WordToDEcrypt)
        If lngI > Len(Password) Then
            lngJ = 1
        Else
            lngJ = lngI
        End If
            
        If Mid(WordToDEcrypt, lngI, 1) <> ":" Then
            strWord = strWord + Chr$(Asc(Mid(WordToDEcrypt, lngI, 1)) - Asc(Mid(Password, lngJ, 1)))
        Else
            strWord = strWord + ":"
        End If
    Next
    Decrypt = strWord
   
End Function

Public Function StripNullTerminator(ByVal sCP As String) As String
            
    StripNullTerminator = StripNullTerminator_F(sCP)
    
End Function

Public Function Translate(ByVal StringToTranslate As Variant, _
                            Optional ByVal ReturnStringToTranslate As Boolean = True) _
                            As String

    Dim cTranslated As String * 520
    
    If (IsNumeric(StringToTranslate) = True) Then
        LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
        cTranslated = StripNullTerminator(cTranslated)
        Translate = RTrim$(cTranslated)
        
        If LenB(Translate) = 0 Then
            If ReturnStringToTranslate Then
                Translate = StringToTranslate
            End If
        End If
        
    Else
        Translate = StringToTranslate
    End If
    
End Function

