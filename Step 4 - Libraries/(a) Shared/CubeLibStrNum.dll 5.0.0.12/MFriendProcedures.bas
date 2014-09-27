Attribute VB_Name = "MFriendProcedures"
Option Explicit

Public Function StripNullTerminator_F(ByVal sCP As String) As String
        
    Dim posNull As Long
    
    posNull = InStr(sCP, Chr$(0))
    If posNull > 0 Then
        StripNullTerminator_F = Left$(sCP, posNull - 1)
    Else
        StripNullTerminator_F = sCP
    End If
    
End Function

Public Function NoBackSlash_F(ByVal cString As String) As String
    Do While Right(cString, 1) = "\"
        cString = Left(cString, Len(cString) - 1)
    Loop
    
    NoBackSlash_F = cString
End Function

Public Function AddBackSlashOnPath_F(ByVal Path As String) As String
    Dim strPathWithBackslash As String
    
    strPathWithBackslash = Path + String(100, 0)

    PathAddBackslash strPathWithBackslash

    strPathWithBackslash = StripNullTerminator_F(strPathWithBackslash)
    
    AddBackSlashOnPath_F = strPathWithBackslash
End Function

Public Function GetTemporaryPath_F() As String
    Dim rc As Long
    Dim lpBuffer As String
    Dim nSize As Long
    
    nSize = 255
    lpBuffer = Space$(nSize)
    rc = GetTempPath(nSize, lpBuffer)
    
    If rc <> 0 Then
        GetTemporaryPath_F = Left$(lpBuffer, rc)
    Else
        GetTemporaryPath_F = ""
    End If
End Function

Public Function IsFileOpenedExclusively_F(ByVal pathName As String, ByVal FileName As String) As Boolean

    Dim strNewName As String
    Dim strPathFileName As String
    Dim intDotPos As Integer
    Dim strPathWithBackslash As String
   
    IsFileOpenedExclusively_F = False
    
    strPathWithBackslash = pathName + String(100, 0)
    PathAddBackslash strPathWithBackslash

    strPathFileName = FileName
    intDotPos = InStrRev(FileName, ".")
    
    If intDotPos Then
        strNewName = Left(FileName, intDotPos - 1) & "x" & Mid(FileName, intDotPos)
    Else
        strNewName = FileName & "x"
    End If
    
    On Error GoTo Error_Handler
        Name strPathFileName As strNewName
        Name strNewName As strPathFileName
    On Error GoTo 0
    
    Exit Function
    
Error_Handler:
    
    IsFileOpenedExclusively_F = True
End Function

