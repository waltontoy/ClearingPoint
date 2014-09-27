VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarServer As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPassword As String 'local copy

Private mvarCurrentFolder As String 'local copy
Private mvarRootFolder As String 'local copy
Private mvarErrorDescription As String 'local copy
Private mvarFiles As clsFiles 'local copy
Private mvarFolders As clsFolders 'local copy
Private mvarConnectionID As Long
Private mvarConnectSuccess As Boolean 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Response[(arg1, arg2, ... , argn)]
Public Enum enuResponse
    ConnectError
    GetFileError
    
    OpenFolderError
    CreateFolderError
    UploadError
    DeleteFileError
    PasswordError
    CreateFolderOK
    OpenFolderOK
    UploadOK
    GetFileOK
    DeleteFileOK
End Enum
Public Event Response(ResponseType As enuResponse, MessageInfo As String)




Public Function CreateFolder(ByVal strFolderName As String) As Boolean
Dim blnRet As Boolean
    blnRet = FtpCreateDirectory(mvarConnectionID, strFolderName)
    If blnRet Then
        CreateFolder = True
        RaiseEvent Response(CreateFolderOK, strFolderName)
    Else
        CreateFolder = False
        RaiseEvent Response(CreateFolderError, strFolderName)
    End If
    
End Function

Public Sub RefreshContent()
    GetFilesAndFolders "*.*", Me, mvarConnectionID
End Sub

Public Function OpenFolder(ByVal strFolder As String) As Boolean
Dim blnRet As Boolean

    blnRet = FtpSetCurrentDirectory(mvarConnectionID, strFolder)
    If blnRet Then
        mvarCurrentFolder = strFolder
        OpenFolder = True
        RaiseEvent Response(OpenFolderOK, strFolder)
        GetFilesAndFolders "*.*", Me, mvarConnectionID
    Else
        OpenFolder = False
        RaiseEvent Response(OpenFolderError, strFolder)
    End If

End Function

Public Function UploadFile(ByVal szFileLocal As String, ByVal strRemoteFile As String) As Boolean
Dim blnRet As Boolean

    blnRet = FtpPutFile(mvarConnectionID, szFileLocal, strRemoteFile, _
     INTERNET_FLAG_RELOAD, 0)
    
    If blnRet Then
        UploadFile = True
        RaiseEvent Response(UploadOK, strRemoteFile)
        RefreshContent
    Else
        UploadFile = False
        RaiseEvent Response(UploadError, strRemoteFile)
    End If
        
End Function

Public Function GetFile(ByVal FileToDownload As Variant, ByVal strLocalDestination As String) As Boolean
    Dim bRet As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim nPos As Long, nTemp As Long
    'Dim nodX As Node
    
    szTempString = mvarFiles.Item(FileToDownload).FileName
        
    'szTempString = mvarDirectoryName
    szFileRemote = szTempString
    nPos = 0
    nTemp = 0
    Do
        nTemp = InStr(1, szTempString, "/", vbBinaryCompare)
        If nTemp = 0 Then Exit Do
        szTempString = Right(szTempString, Len(szTempString) - nTemp)
        nPos = nTemp + nPos
    Loop
    
    szDirRemote = Left(szFileRemote, nPos)
    szFileRemote = Right(szFileRemote, Len(szFileRemote) - nPos)
    szFileLocal = strLocalDestination
    
    bRet = FtpGetFile(mvarConnectionID, szFileRemote, szFileLocal, False, _
    FILE_ATTRIBUTE_NORMAL, FTP_TRANSFER_TYPE_BINARY + INTERNET_FLAG_RELOAD, 0)
    
    If bRet Then
        GetFile = True
        RaiseEvent Response(GetFileOK, szTempString)
    Else
        GetFile = False
        RaiseEvent Response(GetFileError, szTempString)
    End If
    

End Function

Public Function BackToRootFolder() As Boolean
Dim blnRet As Boolean
    blnRet = FtpSetCurrentDirectory(mvarConnectionID, mvarRootFolder)
    If blnRet Then
         mvarCurrentFolder = mvarRootFolder
         BackToRootFolder = True
         GetFilesAndFolders "*.*", Me, mvarConnectionID
    Else
        BackToRootFolder = False
    End If
End Function

Public Function DeleteFTPFile(FileToDelete As Variant) As Boolean
Dim bRet As Boolean
Dim strFile As String

    strFile = mvarFiles.Item(FileToDelete).FileName
    bRet = FtpDeleteFile(mvarConnectionID, strFile)
    If bRet Then
        DeleteFTPFile = True
        RaiseEvent Response(DeleteFileOK, CStr(strFile))
    Else
        DeleteFTPFile = False
        RaiseEvent Response(DeleteFileError, CStr(strFile))
    End If
End Function

Public Property Get Folders() As clsFolders
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folders
    Set Folders = mvarFolders
End Property

Friend Property Set Folders(ByVal vData As clsFolders)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Folders = Form1
    Set mvarFolders = vData
End Property



Public Property Get Files() As clsFiles
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Files
    Set Files = mvarFiles
End Property

Friend Property Set Files(ByVal vData As clsFiles)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Files = Form1
    Set mvarFiles = vData
End Property


Public Property Let ErrorDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ErrorDescription = 5
    mvarErrorDescription = vData
End Property


Public Property Get ErrorDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorDescription
    ErrorDescription = mvarErrorDescription
End Property

Public Property Get RootFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootFolder
    RootFolder = mvarRootFolder
End Property

Friend Property Let RootFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootFolder = 5
    mvarRootFolder = vData
End Property

Public Property Get CurrentFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentFolder
    CurrentFolder = mvarCurrentFolder
End Property


Public Property Let CurrentFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentFolder = 5Re
    mvarCurrentFolder = vData
End Property

Public Property Get ConnectionID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentFolder
    ConnectionID = mvarConnectionID
End Property


Public Property Let ConnectionID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentFolder = 5Re
    mvarConnectionID = vData
End Property

Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property



Public Property Let Server(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    mvarServer = vData
End Property


Public Property Get Server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    Server = mvarServer
End Property

Public Sub Disconnect()
    InternetCloseHandle (mvarConnectionID)
    InternetCloseHandle (hOpen)
    
    mvarConnectSuccess = False

End Sub


Public Sub Connect()
Dim sOrgPath As String

    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
    If hOpen <> 0 Then
        mvarConnectionID = InternetConnect(hOpen, mvarServer, INTERNET_INVALID_PORT_NUMBER, _
         mvarUserName, mvarPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
         
         If mvarConnectionID <> 0 Then
            mvarConnectSuccess = True
            
            sOrgPath = String(MAX_PATH, 0)
            FtpGetCurrentDirectory mvarConnectionID, sOrgPath, Len(sOrgPath)
            sOrgPath = Left(sOrgPath, InStr(1, sOrgPath, String(1, 0), vbBinaryCompare) - 1)

            RootFolder = sOrgPath
            CurrentFolder = sOrgPath
            GetFilesAndFolders "*.*", Me, mvarConnectionID
            
        End If
    End If

End Sub
Public Property Let ConnectSuccess(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectSuccess = 5
    mvarConnectSuccess = vData
End Property


Public Property Get ConnectSuccess() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectSuccess
    ConnectSuccess = mvarConnectSuccess
End Property

