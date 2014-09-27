VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdifactParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim mstrParseString As String
Dim mlngStartIndex As Long

Friend Sub SetStartIndex(ByVal Value As Long)
    mlngStartIndex = Value
End Sub

Public Function GetNextToken(ByVal TokenStartIndex As Long, ByRef TokenEndIndex As Long) As String
    Dim strReturnValue As String
    Dim strCharacter As String * 1
    Dim lngIndex As Long
    Dim blnContinueLoop As Boolean
    Dim lngParseStringLength As Long
    
    If TokenStartIndex = 0 Then
        lngIndex = mlngStartIndex
    Else
        lngIndex = TokenStartIndex
    End If
    
    lngParseStringLength = Len(mstrParseString)
    '----->  Is mlngStartIndex still pointing to a character in mstrParseString?
    If lngIndex <= lngParseStringLength Then
        strReturnValue = vbNullString
        strCharacter = Mid(mstrParseString, lngIndex, 1)
        
        If IsEDISeparator(strCharacter) And strCharacter <> EDI_SEP_RELEASE_CHARACTER Then
            strReturnValue = strCharacter
            lngIndex = lngIndex + 1
        Else
            blnContinueLoop = True
            Do While blnContinueLoop 'Not IsEDISeparator(strCharacter)
                If strCharacter = EDI_SEP_RELEASE_CHARACTER Then
                    lngIndex = lngIndex + 1
                    strCharacter = Mid(mstrParseString, lngIndex, 1)
                End If
                strReturnValue = strReturnValue & strCharacter
                lngIndex = lngIndex + 1
                strCharacter = Mid(mstrParseString, lngIndex, 1)
                'blnContinueLoop = (Not (IsEDISeparator(strCharacter) And (strCharacter <> EDI_SEP_RELEASE_CHARACTER)) And (lngIndex <= lngParseStringLength))
                blnContinueLoop = (Not (IsEDISeparator(strCharacter) And (strCharacter <> EDI_SEP_RELEASE_CHARACTER)) And (lngIndex <= lngParseStringLength) And strCharacter <> vbNullString)
            Loop
        End If
    End If

    '----->  Update character pointer for mstrParseString
    mlngStartIndex = lngIndex
    TokenEndIndex = mlngStartIndex
    GetNextToken = strReturnValue
End Function

Private Function GetParseStringFromFile(ByVal FilePath As String) As String
    Dim strReturnValue As String
    
    Dim lngFileHandle As Long
    
    lngFileHandle = FreeFile
    Open FilePath For Input As #lngFileHandle
    
    Line Input #lngFileHandle, strReturnValue
    
    Close #lngFileHandle
    GetParseStringFromFile = strReturnValue
End Function

Public Sub SetParseString(ByVal IEMSource As IEMessageSources, ByVal BufferString As String)
    Select Case IEMSource
        Case IEMessageSources.IEMessageSource_Constant
            mstrParseString = BufferString
        Case IEMessageSources.IEMessageSource_File
            mstrParseString = GetParseStringFromFile(BufferString)
    End Select
End Sub

Private Sub Class_Initialize()
    mlngStartIndex = 1
End Sub
