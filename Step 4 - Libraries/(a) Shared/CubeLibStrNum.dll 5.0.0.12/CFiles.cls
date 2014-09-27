VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFiles"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim arrInfo() As Byte

Public Function IsFileExisting(ByVal filefullPath As String) As Boolean
    IsFileExisting = CBool(PathFileExists(filefullPath))
End Function

Public Function GetFileTitle(ByVal filefullPath As String) As String
    GetFileTitle = GetFileSystemObjectTitle(filefullPath)
End Function

Public Function GetFileInformation(ByVal filefullPath As String) As CFileInfo
    Dim objFileInfo As CFileInfo
    
    Set objFileInfo = New CFileInfo
    
    Set objFileInfo = GetFileInformation_F(filefullPath)
    
    Set GetFileInformation = objFileInfo
    
    Set objFileInfo = Nothing
End Function

Public Sub CreateLogFile(Path As String, Filename As String, FileExtension As String)
    Dim F
    On Error Resume Next
    F = Dir(Path + "\" + Trim(Filename) + "." + FileExtension)
    If Len(Trim(F & " ")) = 0 Then
        Open Path + "\" + Trim(Filename) + "." + FileExtension For Output As #1 ' Open file for output.
        Print #1, "warehouse by cubepoint, inc." & Trim(Filename)
        Close #1
    End If
End Sub

Public Function GetTemporaryFilename() As String
    Dim strTemporaryFileName As String
    Dim strWindowsTemporaryPath As String
    
    On Error GoTo ErrorHandler

    strWindowsTemporaryPath = GetTemporaryPath_F
    
    ' Create a buffer
    strTemporaryFileName = String(260, 0)
    
    ' Get a temporary filename
    GetTempFileName AddBackSlashOnPath_F(strWindowsTemporaryPath), "pld", 0, strTemporaryFileName
        
    ' Remove all the unnecessary chr$(0)'s
    strTemporaryFileName = Left$(strTemporaryFileName, InStr(1, strTemporaryFileName, Chr$(0)) - 1)
    
    ' Set the file attributes
    SetFileAttributes strTemporaryFileName, FILE_ATTRIBUTE_TEMPORARY
    
        
    GetTemporaryFilename = strTemporaryFileName
    
ErrorHandler:
        Select Case Err.Number
              Case 0
              
            Case Else
                'AddToTrace "Error in Main, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                  
        End Select

End Function


Public Function GetFileVersion(ByVal FilePathName As String, Optional ByVal PromptError As Boolean = False) As String
    Dim mFileVersion As String
    Dim mresult As Variant
    Dim lpHandle As Long
    Dim lpInfoBlock As Long
    Dim lInfosize As Long
    Dim lngTemp As Long
    Dim strLang As String
    Dim arrTemp(4) As Byte

    
    lpHandle = 0
    lInfosize = GetFileVersionInfoSize(FilePathName, lpHandle)
        
    If lInfosize = 0 Then
        If PromptError = True Then
            MsgBox "Cannot proceed due to invalid info size.", vbInformation, "Invalid Info (3008)"
        End If
    End If
    
    ReDim arrInfo(lInfosize)
    mresult = GetFileVersionInfo(FilePathName, 0&, lInfosize, arrInfo(0))
    
    If mresult = 0 Then
        If PromptError = True Then
            MsgBox "File version not found.", vbInformation, "Invalid Version (3009)"
        End If
    End If
    
    mresult = VerQueryValue(arrInfo(0), "\VarFileInfo\Translation", lpInfoBlock, lInfosize)
    
    If mresult = 0 Then
        If PromptError = True Then
            MsgBox "Cannot proceed.", vbInformation, "Invalid Handle (3010)"
        End If
    End If
    
    ' To access the block info, we first move the obtained block info from
    ' strFileInfoString pointed by lpVerBlocke into an array.
    'Dim arrTemp(4) As Byte
    MoveMemory arrTemp(0), lpInfoBlock, lInfosize
    
    ' Word of language + word of codepage in the lang structure. High-low.
    lngTemp = arrTemp(2) + arrTemp(3) * &H100 + arrTemp(0) * &H10000 + arrTemp(1) * &H1000000
    strLang = Hex(lngTemp)
    
    ' Length of lang-codepage is 8; pad zeros in front in case length < 8
    ' e.g. "04094E4" becomes "004094E4"
    Do While Len(strLang) < 8
        strLang = "0" & strLang
    Loop
    
    '----------------------------------------------------------------
    ' "\StringFileInfo\lang-codepage\string-name" pieces of data:
    '----------------------------------------------------------------
    mFileVersion = GetInfoBlock("FileVersion", strLang, lpInfoBlock, lInfosize)
End Function

Private Function GetInfoBlock(ByVal inStringName As String, ByRef Lang As String, ByRef InfoBlock As Long, ByRef InfoSize As Long) As String
    Dim strFileInfoString As String
    Dim varResult As Variant
    
    GetInfoBlock = ""
    
    ' We don't know how long, fill the string with some char 0 first
    strFileInfoString = String(128, 0)
    varResult = VerQueryValue(arrInfo(0), "\StringFileInfo\" & Lang & "\" & inStringName, InfoBlock, InfoSize)
    
    If varResult > 0 Then
        lstrcpy strFileInfoString, InfoBlock
        GetInfoBlock = Mid(strFileInfoString, 1, InStr(strFileInfoString, vbNullChar) - 1)
    End If
End Function

Public Function Copy(ByVal Filename As String, ByVal SourceDir As String, ByVal DestinationDir As String, Optional ByVal DestinationFile As String = "") As Boolean
    Dim intFileAttributes As Integer
    Dim blnFileExists As Boolean
    Dim strTargetFile As String
                    
    Copy = False
    blnFileExists = False
    
    
    If DestinationFile <> "" Then
        strTargetFile = DestinationFile
    Else
        strTargetFile = Filename
    End If
    
    Load frmDummy
        
    frmDummy.flbDummy.Hidden = True
    frmDummy.flbDummy.Archive = True
    frmDummy.flbDummy.System = True
    frmDummy.flbDummy.Pattern = Filename
    
    On Error GoTo ERR_PATH_NOT_FOUND
        frmDummy.flbDummy.Path = SourceDir
    On Error GoTo 0
    
    If frmDummy.flbDummy.ListCount <= 0 Then
        GoTo ERR_SOURCE_FILE_MISSING
    End If
    
    On Error GoTo ERR_PATH_NOT_FOUND
        frmDummy.flbDummy.Path = DestinationDir
    On Error GoTo 0
    
    If frmDummy.flbDummy.ListCount > 0 Then
        intFileAttributes = GetAttr(DestinationDir & "\" & strTargetFile)
        blnFileExists = True
    End If
    
    If blnFileExists Then
        If intFileAttributes And vbReadOnly Then
            SetAttr DestinationDir & "\" & strTargetFile, intFileAttributes - vbReadOnly
            FileCopy SourceDir & "\" & Filename, DestinationDir & "\" & strTargetFile
            SetAttr DestinationDir & "\" & strTargetFile, intFileAttributes
        Else
            FileCopy SourceDir & "\" & Filename, DestinationDir & "\" & strTargetFile
        End If
    Else
        FileCopy SourceDir & "\" & Filename, DestinationDir & "\" & strTargetFile
    End If

    Copy = True
    
    Exit Function

ERR_SOURCE_FILE_MISSING:
    
    'Err.Raise 8003, , "Source file not found."
    
    Exit Function
    
ERR_PATH_NOT_FOUND:
    
    Select Case Err.Number
        Case 76
            On Error GoTo 0
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Function

Public Function IsCleanFileName(ByVal Filename As String) As Boolean

    If InStr(1, Filename, "/") > 0 Or _
       InStr(1, Filename, "\") > 0 Or _
       InStr(1, Filename, ":") > 0 Or _
       InStr(1, Filename, "*") > 0 Or _
       InStr(1, Filename, "?") > 0 Or _
       InStr(1, Filename, Chr(34)) > 0 Or _
       InStr(1, Filename, "<") > 0 Or _
       InStr(1, Filename, ">") > 0 Or _
       InStr(1, Filename, "|") > 0 Then
    
        IsCleanFileName = False
    Else
        IsCleanFileName = True
    End If
End Function

Public Function GetCleanFileName(ByVal Filename As String) As String
    'Removes illegal characters to complement CheckFileName Function.
    Filename = Replace$(Filename, "/", "")      ' /
    Filename = Replace$(Filename, "\", "")      ' \
    Filename = Replace$(Filename, ":", "")      ' :
    Filename = Replace$(Filename, "*", "")      ' *
    Filename = Replace$(Filename, "?", "")      ' ?
    Filename = Replace$(Filename, Chr(34), "")  ' "
    Filename = Replace$(Filename, "<", "")      ' <
    Filename = Replace$(Filename, ">", "")      ' >
    Filename = Replace$(Filename, "|", "")      ' |
    
    GetCleanFileName = Filename
End Function

Public Function GetValidNonExistingFileName(ByVal DestinationPath As String, _
                                            ByVal FileNameOnly As String, _
                                            ByVal FileNameExt) As String
    Dim lngCounter As Long
    Dim strFileNameOnly As String
    Dim strFileNameExt As String
    Dim strDestinationPath As String
    
    strDestinationPath = DestinationPath
    strFileNameOnly = GetCleanFileName(FileNameOnly)
    strFileNameExt = FileNameExt

    strDestinationPath = IIf(Right(strDestinationPath, 1) <> "\", strDestinationPath & "\", strDestinationPath)
    If Len(Dir(strDestinationPath & strFileNameOnly & "." & strFileNameExt)) > 0 Then
        lngCounter = 1
        Do While Len(Dir(strDestinationPath & strFileNameOnly & "(" & lngCounter & ")." & strFileNameExt)) > 0
            lngCounter = lngCounter + 1
        Loop

        GetValidNonExistingFileName = strFileNameOnly & "(" & lngCounter & ")." & strFileNameExt
        'g_strDocName = g_strDocName & "***" & strFileNameOnly & "(" & lngCounter & ").pdf"
    Else
        GetValidNonExistingFileName = strFileNameOnly & "." & strFileNameExt
        'g_strDocName = g_strDocName & "***" & strFileNameOnly & ".pdf"
    End If
End Function

Public Function CheckFileName(ByVal cFilename As String, ByVal cExtName As String, Optional ByVal blnWarnUser As Boolean) As String
    Dim intPos As Integer
    Dim intFreeFile As Integer
    Dim strWorkingFileName As String
    Dim strStrippedFileName As String
    
    CheckFileName = "F" & cFilename
    
    strWorkingFileName = cFilename
    
    If InStr(1, strWorkingFileName, "\") > 0 Then
        intPos = InStrRev(strWorkingFileName, "\")
        
        If Len(strWorkingFileName) - intPos = 0 Then GoTo 5    ' (no filename)
        
        strStrippedFileName = Left(strWorkingFileName, intPos)
        
        If Len(strStrippedFileName) = 1 Then GoTo 5        ' (invalid filename, e.g. "\gemma")
        
        strWorkingFileName = Mid(strWorkingFileName, intPos + 1)
    End If
    
    If Not IsCleanFileName(strWorkingFileName) Then GoTo 5    ' (illegal characters)
    
    On Error GoTo 5
    
    If Len(strStrippedFileName) > 0 Then
        intFreeFile = FreeFile()
        
        Open strStrippedFileName & "trial.txt" For Output As #intFreeFile
        Close #intFreeFile
        
        Kill strStrippedFileName & "trial.txt"
    End If
    
    intPos = InStrRev(strWorkingFileName, ".")
    
    If intPos > 0 Then
        If Len(strWorkingFileName) - intPos = 0 Then
            GoTo 5    ' (invalid filename, e.g. "gemma.")
        End If
        
        If Len(strWorkingFileName) - intPos >= 3 Then
            strWorkingFileName = Left(strWorkingFileName, intPos + 3)
        End If
    Else
        strWorkingFileName = strWorkingFileName & "." & Left(cExtName, 3)
    End If
    
    CheckFileName = Trim(strStrippedFileName & strWorkingFileName)
    
    Exit Function
    
5:
    
    'If blnWarnUser = True Then
    '    MsgBox Left(Translate(572), Len(Translate(572)) - 1) & " " & cFileName & ".", vbInformation
    'End If
    
    CheckFileName = vbNullString
End Function

Public Function FormatSize(lngSize As Long) As String
    FormatSize = FormatSize_F(lngSize)
End Function

Public Function IsFileOpenedExclusively(ByVal FilePath As String, ByVal Filename As String) As Boolean
    IsFileOpenedExclusively = IsFileOpenedExclusively_F(FilePath, Filename)
End Function

