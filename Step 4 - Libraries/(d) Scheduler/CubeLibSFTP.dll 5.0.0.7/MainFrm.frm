VERSION 5.00
Object = "{E532970A-FEEB-4A38-A1BB-4E462DDCA8B9}#8.0#0"; "SFTPBBoxCli8.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2A486285-0C52-4069-8D0C-4E5EB6433DE0}#8.0#0"; "BaseBBox8.dll"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sftp Demo"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin SFTPBBoxCli8.ElSimpleSftpClientX ElSimpleSftpClientX 
      Left            =   6480
      Top             =   3240
      NewLineConvention=   $"MainFrm.frx":0000
      ClientUserName  =   ""
      ClientHostName  =   ""
      ForceCompression=   0   'False
      CompressionLevel=   6
      SoftwareName    =   "SecureBlackbox.8"
      Username        =   ""
      Password        =   ""
      Address         =   ""
      Port            =   22
      UseInternalSocket=   -1  'True
      SocketTimeout   =   0
      UseSocks        =   0   'False
      SocksServer     =   ""
      SocksPort       =   1080
      SocksUserCode   =   ""
      SocksPassword   =   ""
      SocksVersion    =   1
      SocksResolveAddress=   0   'False
      SocksAuthentication=   0
      UseWebTunneling =   0   'False
      WebTunnelAddress=   ""
      WebTunnelPort   =   3128
      WebTunnelAuthentication=   0
      WebTunnelUserId =   ""
      WebTunnelPassword=   ""
      SFTPBufferSize  =   131072
      PipelineLength  =   32
      DownloadBlockSize=   8192
      UploadBlockSize =   32768
      CurrentOperationCancel=   0   'False
      ASCIIMode       =   0   'False
      LocalNewLineConvention=   $"MainFrm.frx":0006
      DefaultWindowSize=   2048000
      MinWindowSize   =   2048
      SSHAuthOrder    =   1
      AutoAdjustCiphers=   -1  'True
      AutoAdjustTransferBlock=   -1  'True
      LocalAddress    =   ""
      LocalPort       =   0
      UseUTF8         =   -1  'True
      OperationErrorHandling=   0
      RequestPasswordChange=   0   'False
      AuthAttempts    =   1
      CertAuthMode    =   1
      IncomingSpeedLimit=   0
      OutgoingSpeedLimit=   0
      KeepAlivePeriod =   0
      SocksUseIPv6    =   0   'False
      UseIPv6         =   0   'False
      AdjustFileTimes =   0   'False
      FIPSMode        =   0   'False
      GSSHostName     =   ""
      GSSDelegateCredentials=   0   'False
      UseTruncateFlagOnUpload=   -1  'True
      TreatZeroSizeAsUndefined=   -1  'True
      UseUTF8OnV3     =   0   'False
   End
   Begin VB.Timer TimeOutCounter 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   1200
   End
   Begin BaseBBox8.ElSBLicenseManagerX ElSBLicenseManagerX 
      Left            =   6480
      Top             =   2520
   End
   Begin ComctlLib.ListView LogListView 
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Timestamp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Event"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.CommandButton btnUpdateInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6120
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Refresh"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnPutFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4920
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Upload"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnGetFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3720
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Download selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnDelete 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2520
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":09D2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnRename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":0D14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Rename selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnMkDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":1056
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Make directory"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Edit4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   7215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Permissions"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox editPath 
      Height          =   300
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "."
      Top             =   1800
      Width           =   7215
   End
   Begin MSComDlg.CommonDialog OpenDialog1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog SaveDialog1 
      Left            =   5640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection properties"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox editPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox editUserName 
         Height          =   300
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox editHost 
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Text            =   "192.168.0.1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "User name"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public mvarCSFTP As CSFTPUploadDownload

Private Const FILE_BLOCK_SIZE = 4096
Private Const STATE_OPEN_DIRECTORY_SENT = 1
Private Const STATE_READ_DIRECTORY_SENT = 2
Private Const STATE_CHANGE_DIR = 3
Private Const STATE_MAKE_DIR = 4
Private Const STATE_RENAME = 5
Private Const STATE_REMOVE = 6
Private Const STATE_DOWNLOAD_OPEN = 7
Private Const STATE_DOWNLOAD_RECEIVE = 8
Private Const STATE_UPLOAD_OPEN = 9
Private Const STATE_UPLOAD_SEND = 10
Private Const STATE_CLOSE_HANDLE = 11

Private Const MESSAGE_PREFIX = "SFTP SERVER: "

Private m_strCurrentHandle As String
Private m_strCurrentDir As String
Private m_strRelDir As String
Private m_strCurrentFile As String

Public m_colCurrentFileList As Collection

Private m_lngState As Long
Private m_lngCurrentFileOffset As Long
Private m_lngCurrentFileSize As Long
Private m_lngCurrentFile As Integer

Private m_blnClientDataAvailable As Boolean
Public m_blnDirectoryReadFinished As Boolean
Public m_blnFileUploaded As Boolean
Public m_blnFileDownloaded As Boolean
Public m_blnFileDeleted As Boolean

Public m_blnHasError As Boolean

Private m_blnAuthenticationFailed As Boolean

Public Sub LoadForm(ByRef CSFTP As CSFTPUploadDownload)
    Load Me
    
    Set mvarCSFTP = CSFTP
End Sub

Private Sub Form_Load()
    
    Set m_colCurrentFileList = New Collection
  
    ElSBLicenseManagerX.SetLicenseKey ("90D40DF1DDFEEC8F659575583B2619AE021FB2D4DCAB1F82E429A554D48A77E8FCD05FBB554D713297DEDEEFE375828F822A11D20B2B7A2671A844123C45D8176FA1898EECA5F4401ACAF8999496A60AD7BCEE80B4C2764E534F7215FF42C83FE42CCA6414394BE80394EFE3A67C6DF36494EB440BC16BB62C32194B4AB2E8FAEAAA11F99004851D8DF675F2C33B3F70C4811A487E59D4023E7C31950A4AC948CAEE628EC8134DAFE72F314D88BDCE932328F8D75AD620E169D90348B51FBFD651779357026431BD0235B2F8FBB10F880EBCEDEEE714E88E644082B878E4297917F1336A5DA52446736870F2AAF8C070EEB5A89519583E9B161D50D21B1B1846")
    
    ElSimpleSftpClientX.EnableVersion SB_SFTP_VERSION_3
    ElSimpleSftpClientX.EnableAuthenticationType SSH_AUTH_TYPE_PASSWORD
    ElSimpleSftpClientX.EnableAuthenticationType SSH_AUTH_TYPE_KEYBOARD
End Sub

Private Sub ElSimpleSftpClientX_OnError(ByVal ErrorCode As Long)
    Call TranslateError(ErrorCode)
    
    If Len(ElSimpleSftpClientX.ServerSoftwareName) > 0 Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Server software identified itself as: " & ElSimpleSftpClientX.ServerSoftwareName
    End If
    
    Call mvarCSFTP.DisconnectFromServer(False)
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationStart(ByVal SupportedAuths As Long)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "TCP connection opened..."
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authenticating Username: " & mvarCSFTP.UserName & ", Password: " & mvarCSFTP.Password & " using SSHClient..."
    
    m_blnAuthenticationFailed = False
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationFailed(ByVal AuthenticationType As SSHBBoxCli8.TxSSHAuthenticationType)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication type " & AuthenticationType & " failed."
    m_blnAuthenticationFailed = True
    Call mvarCSFTP.DisconnectFromServer(False)
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationSuccess()
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication succeeded..."
End Sub

Private Sub ElSimpleSftpClientX_OnSend(ByVal Data As Variant)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Sending Data ..."
End Sub

Public Sub DeleteFile(ByVal Name As String)
    m_blnHasError = False
    
    If Not ElSimpleSftpClientX.Active Then
        m_blnHasError = True
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Delete File Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Removing File " & Name
    Call ElSimpleSftpClientX.RemoveFile(m_strCurrentDir & "/" & Name)
    m_blnFileDeleted = True
End Sub

Public Sub DownloadFile(ByVal info As IElSftpFileInfoX)
    m_blnHasError = False
    
    If Not ElSimpleSftpClientX.Active Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "DownloadFile Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Starting file download, " & info.Name
     
    On Error GoTo ErrHandler
    Call ElSimpleSftpClientX.DownloadFile(m_strCurrentDir & "/" & info.Name, mvarCSFTP.MdbPath + "\" + info.Name)
    m_blnFileDownloaded = True
    
    Exit Sub
    
ErrHandler:
    If Err.Number > 0 Then
        m_blnHasError = True
    End If
    
End Sub

Public Sub UploadFile(ByVal LocalFile As String)
    Dim FName As String
    
    If Not ElSimpleSftpClientX.Active Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "UploadFile Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Starting file upload, " & LocalFile
    
    FName = ExtractFileName(LocalFile)
    Dim RemoteName As String
    RemoteName = m_strCurrentDir & "/" & FName
    
    On Error GoTo ErrHandler
    m_blnFileUploaded = False
    Call ElSimpleSftpClientX.UploadFile(LocalFile, RemoteName)
    m_blnFileUploaded = True
        
    Exit Sub
    
ErrHandler:
    If Err.Number > 0 Then
        m_blnHasError = True
    End If
    
End Sub

Function OpenFileForRead(ByRef File As Integer, ByVal FileName As String) As Boolean
    File = FreeFile()
    Open FileName For Binary Access Read As #File
End Function

Function OpenFileForWrite(ByRef File As Integer, ByVal FileName As String) As Boolean
    File = FreeFile()
    Open FileName For Output Access Write As #File
    
    Close File
    Open FileName For Binary Access Write As #File
End Function

Function ExtractFileName(ByVal FileName As String) As String
    Dim ch As String
    Dim i As Integer, Idx As Integer
    
    Idx = 0
    For i = Len(FileName) To 1 Step -1
        ch = Mid(FileName, i, 1)
        
        If (ch = ":") Or (ch = "\") Or (ch = "/") Then
            Idx = i
            Exit For
        End If
    Next
    
    ExtractFileName = Mid(FileName, Idx + 1)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim FileInfo1 As IElSftpFileInfoX
    
    TimeOutCounter.Enabled = False
    
    'Call UnloadControls(Me)
    
    While m_colCurrentFileList.Count > 0
        Set FileInfo1 = m_colCurrentFileList(1)
        Set FileInfo1 = Nothing
        Call m_colCurrentFileList.Remove(1)
    Wend
    
    Set m_colCurrentFileList = Nothing
    
    Unload Me
    
    Set mvarCSFTP = Nothing
    Set MainForm = Nothing
End Sub

Private Sub TimeOutCounter_Timer()
    mvarCSFTP.HasTimeOut = True
End Sub


Public Sub ConnectSFTP()
    If ElSimpleSftpClientX.Active Then
        Call ElSimpleSftpClientX.Close
    End If
    
    ElSimpleSftpClientX.UserName = mvarCSFTP.UserName
    ElSimpleSftpClientX.Password = mvarCSFTP.Password
    ElSimpleSftpClientX.EnableVersion SB_SFTP_VERSION_3
    ElSimpleSftpClientX.Address = mvarCSFTP.HostName
    ElSimpleSftpClientX.Port = mvarCSFTP.PortNumber
    
    mvarCSFTP.TraceText = "Connecting to Hostname: " & mvarCSFTP.HostName & ", PortNumber: " & mvarCSFTP.PortNumber & "..."
    
    On Error GoTo ErrHandler
    Call ElSimpleSftpClientX.Open

ErrHandler:
    Select Case Err.Number
        Case 0
            
        Case Else
            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Connection Error - " & Err.Description & " (" & Err.Number & ") "
            
    End Select
    
End Sub


Public Sub DisconnectSFTP()
    If ElSimpleSftpClientX.Active Then
        Call ElSimpleSftpClientX.Close
    End If
    
    m_blnDirectoryReadFinished = False
    m_blnFileDownloaded = False
    m_blnFileDeleted = False
End Sub


Public Sub RefreshRootDirectoryList()
    Dim Listing As Variant, i As Long
    Dim info As IElSftpFileInfoX
    Dim info_copy As IElSftpFileInfoX
    Dim item As ListItem
    Dim a() As IElSftpFileInfoX

    m_blnHasError = False

    If Not ElSimpleSftpClientX.Active Then
        Exit Sub
    End If

    'Clearing old data
    While m_colCurrentFileList.Count > 0
        m_colCurrentFileList.Remove (1)
    Wend

    On Error GoTo HandleErr
    m_strCurrentDir = vbNullString
    m_strCurrentDir = ElSimpleSftpClientX.RequestAbsolutePath(m_strCurrentDir)

    'Retrieving directory contents
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Retrieving file list..."
    Call ElSimpleSftpClientX.ListDirectory(m_strCurrentDir, Listing)
    For i = LBound(Listing) To UBound(Listing)
        Set info_copy = New ElSftpFileInfoX

        On Error Resume Next
        Set info = Listing(i)
        On Error GoTo 0

        Call info.CopyTo(info_copy)

        If Not info_copy.Attributes.Directory Then
            Call m_colCurrentFileList.Add(info_copy)
        End If
    Next

    m_blnDirectoryReadFinished = True
    Exit Sub

HandleErr:
    m_blnDirectoryReadFinished = True
    m_blnHasError = True
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Refresh Root Directory List Error - " & Err.Description & " (" & Err.Number & ") "

End Sub


Public Sub TranslateError(ByVal ErrorCode As Long)
    Dim strErrorMessage As String
    
    Select Case ErrorCode
        Case ERROR_SSH_INVALID_IDENTIFICATION_STRING
            strErrorMessage = "Invalid identification string of SSH-protocol."
        
        Case ERROR_SSH_INVALID_VERSION
            strErrorMessage = "Invalid or unsupported version."
        
        Case ERROR_SSH_INVALID_MESSAGE_CODE
            strErrorMessage = "Unsupported message code."
        
        Case ERROR_SSH_INVALID_CRC
            strErrorMessage = "Message CRC is invalid."
        
        Case ERROR_SSH_INVALID_PACKET_TYPE
            strErrorMessage = "Invalid (unknown) packet type."
            
        Case ERROR_SSH_INVALID_PACKET_TYPE
            strErrorMessage = "Invalid (unknown) packet type."
        
        Case ERROR_SSH_INVALID_PACKET
            strErrorMessage = "Packet composed incorrectly."
        
        Case ERROR_SSH_UNSUPPORTED_CIPHER
            strErrorMessage = "There is no cipher supported by both: client and server."
        
        Case ERROR_SSH_UNSUPPORTED_AUTH_TYPE
            strErrorMessage = "Authentication type is unsupported."
        
        Case ERROR_SSH_INVALID_RSA_CHALLENGE
            strErrorMessage = "The wrong signature during public key-authentication."
        
        Case ERROR_SSH_AUTHENTICATION_FAILED
            strErrorMessage = "Authentication failed. There could be wrong password or something else."
        
        Case ERROR_SSH_INVALID_PACKET_SIZE
            strErrorMessage = "The packet is too large."
        
        Case ERROR_SSH_HOST_NOT_ALLOWED_TO_CONNECT
            strErrorMessage = "Connection was rejected by remote host."
        
        Case ERROR_SSH_PROTOCOL_ERROR
            strErrorMessage = "Another protocol error."
        
        Case ERROR_SSH_KEY_EXCHANGE_FAILED
            strErrorMessage = "Key exchange failed."
        
        Case ERROR_SSH_INVALID_MAC
            strErrorMessage = "Received packet has invalid MAC."
        
        Case ERROR_SSH_COMPRESSION_ERROR
            strErrorMessage = "Compression or decompression error."
        
        Case ERROR_SSH_SERVICE_NOT_AVAILABLE
            strErrorMessage = "Service (sftp, shell, etc.) is not available."
        
        Case ERROR_SSH_PROTOCOL_VERSION_NOT_SUPPORTED
            strErrorMessage = "Version is not supported."
        
        Case ERROR_SSH_HOST_KEY_NOT_VERIFIABLE
            strErrorMessage = "Server key can not be verified."
        
        Case ERROR_SSH_CONNECTION_LOST
            strErrorMessage = "Connection was lost by some reason."
        
        Case ERROR_SSH_APPLICATION_CLOSED
            strErrorMessage = "User on the other side of connection closed application that led to disconnection."
        
        Case ERROR_SSH_TOO_MANY_CONNECTIONS
            strErrorMessage = "The server is overladen."
        
        Case ERROR_SSH_AUTH_CANCELLED_BY_USER
            strErrorMessage = "User tired of invalid password entering."
        
        Case ERROR_SSH_NO_MORE_AUTH_METHODS_AVAILABLE
            strErrorMessage = "There are no more methods for user authentication."
        
        Case ERROR_SSH_ILLEGAL_USERNAME
            strErrorMessage = "There is no user with specified username on the server."
        
        Case ERROR_SSH_INTERNAL_ERROR
            strErrorMessage = "Internal error of implementation."
        
        Case ERROR_SSH_NOT_CONNECTED
            strErrorMessage = "There is no connection but user tries to send data."
        
        Case ERROR_SSH_CONNECTION_CANCELLED_BY_USER
            strErrorMessage = "The connection was cancelled by user."
        
        Case ERROR_SSH_FORWARD_DISALLOWED
            strErrorMessage = "SSH forward disallowed."
        
        Case ERROR_SSH_ONKEYVALIDATE_NOT_ASSIGNED
            strErrorMessage = "The event handler for OnKeyValidate event, has not been specified by the application."
        
        Case ERROR_SSH_TCP_CONNECTION_FAILED
            strErrorMessage = "TCP connection failed."
        
        Case ERROR_SSH_TCP_BIND_FAILED
            strErrorMessage = "TCP bind failed."
    
        Case Else
            strErrorMessage = "Unknown Error."
            
    End Select
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "SSH error (" & ErrorCode & ") - " & strErrorMessage
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "If you have ensured that all connection parameters are correct and you still can't connect,"
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "please contact CANDS support."
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Remember to provide details about the error that happened."

End Sub
