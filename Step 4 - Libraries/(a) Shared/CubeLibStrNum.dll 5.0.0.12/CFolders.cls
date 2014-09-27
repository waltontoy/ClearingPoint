VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFolders"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum FileAttributeConstant
    [Archive]
    [Directory]
    [Hidden]
    [Normal]
    [Read Only]
    [System]
    [Temporary]
    [Compressed]
End Enum

Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal strPath As String)

Public Function NoBackSlash(ByVal cString As String) As String
    
    NoBackSlash = NoBackSlash_F(cString)
End Function

Public Function GetShortPath(ByVal FilePathName As String) As String
    Dim lngResult As Long
    Dim strBufferPath As String
    
    ' Create a Buffer
    strBufferPath = String$(255, 0)
    
    ' Retrieve the Short Pathname
    lngResult = GetShortPathName(FilePathName, strBufferPath, 254)
        
    ' Remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strBufferPath, lngResult)
End Function

Public Function StripPath(ByVal FullyQualifiedPath As String) As String
    Dim strFullyQualifiedPath As String
    
    strFullyQualifiedPath = FullyQualifiedPath

    PathStripPath strFullyQualifiedPath
    
    StripPath = strFullyQualifiedPath
End Function

Public Function AddBackSlashOnPath(ByVal Path As String) As String

    AddBackSlashOnPath = AddBackSlashOnPath_F(Path)

End Function

Public Function GetTemporaryPath() As String
    
    GetTemporaryPath = GetTemporaryPath_F
End Function

Public Function GetSystemPath() As String
    Dim rc As Long
    Dim lpBuffer As String
    Dim nSize As Long
    
    nSize = 255
    lpBuffer = Space$(nSize)
    rc = GetSystemDirectory(lpBuffer, nSize)
    
    If rc <> 0 Then
        GetSystemPath = Left$(lpBuffer, rc)
    Else
        GetSystemPath = ""
    End If
End Function

Public Function IsPathAUNCPath(ByVal FolderPath As String) As Boolean
        
    If PathIsUNC(FolderPath) Then
        IsPathAUNCPath = True
    Else
        IsPathAUNCPath = False
    End If
End Function

Public Function GetPathLogicalDrive(ByVal FolderPath As String) As String
    Dim strFolderPath As String
    
    strFolderPath = FolderPath
    PathStripToRoot strFolderPath
    strFolderPath = StripNullTerminator_F(strFolderPath)
    
    GetPathLogicalDrive = strFolderPath
End Function

Public Function IsPathANetworkPath(ByVal FolderPath As String) As Boolean
    IsPathANetworkPath = CBool(PathIsNetworkPath(FolderPath))
        
End Function

Public Function IsPathADirectory(ByVal FolderPath As String) As Boolean
    IsPathADirectory = CBool(PathIsDirectory(FolderPath))
End Function

Public Function IsFolderAttribute(ByVal FolderPath As String, _
                                  ByVal FileAttribute As FileAttributeConstant) As Boolean
    Dim lngFileAttributes As Long
    
    lngFileAttributes = GetFileAttributes(FolderPath)
    
    Select Case FileAttribute
        Case FileAttributeConstant.[Archive]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_ARCHIVE)
        Case FileAttributeConstant.[Directory]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
        Case FileAttributeConstant.[Hidden]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_HIDDEN)
        Case FileAttributeConstant.[Normal]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_NORMAL)
        Case FileAttributeConstant.[Read Only]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_READONLY)
        Case FileAttributeConstant.[System]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_SYSTEM)
        Case FileAttributeConstant.[Temporary]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_TEMPORARY)
        Case FileAttributeConstant.[Compressed]
            IsFolderAttribute = (lngFileAttributes And FILE_ATTRIBUTE_COMPRESSED)
        Case Else
            IsFolderAttribute = False
    End Select
End Function

Public Function CreateDirectory(ByVal Directory As String) As Boolean
    
    CreateDirectory = False
    
    On Error GoTo ERROR_DRIVE_NOT_FOUND
        MakeSureDirectoryPathExists Directory & "\"
    On Error GoTo 0
    
    Load frmDummy
    
    frmDummy.flbDummy.Hidden = True
    frmDummy.flbDummy.Archive = True
    frmDummy.flbDummy.System = True
    frmDummy.flbDummy.Pattern = ""
    
    On Error GoTo ERROR_DRIVE_NOT_FOUND
        frmDummy.flbDummy.Path = Directory
    On Error GoTo 0
    
    CreateDirectory = True
    
    Exit Function
    
ERROR_DRIVE_NOT_FOUND:

    Select Case Err.Number
        Case 76, 68
            On Error GoTo 0
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Function

