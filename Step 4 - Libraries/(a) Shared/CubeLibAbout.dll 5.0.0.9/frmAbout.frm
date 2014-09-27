VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Application"
   ClientHeight    =   4290
   ClientLeft      =   5400
   ClientTop       =   840
   ClientWidth     =   5775
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2961.033
   ScaleMode       =   0  'User
   ScaleWidth      =   5423.023
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2064"
   Begin VB.CommandButton cmdTechSupport 
      Caption         =   "&Tech Support"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Tag             =   "2057"
      Top             =   3000
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Tag             =   "178"
      Top             =   3840
      Width           =   1305
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info"
      Height          =   345
      Left            =   4320
      TabIndex        =   1
      Tag             =   "233"
      Top             =   3405
      Width           =   1290
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1320
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   7
      Top             =   1800
      Width           =   4335
      Begin VB.Label lblSerialNumber 
         AutoSize        =   -1  'True
         Caption         =   "Product ID: 01-ASDF-YHTG-PLKJ-HYTU-1"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Tag             =   "2058"
         Top             =   420
         Width           =   3045
      End
      Begin VB.Label lblLicensee 
         AutoSize        =   -1  'True
         Caption         =   "[ This product is not licensed. ]"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Tag             =   "2059"
         Top             =   120
         Width           =   2160
      End
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   150
      Top             =   75
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   112.686
      X2              =   5296.251
      Y1              =   2000.25
      Y2              =   2000.25
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   112.686
      X2              =   5296.251
      Y1              =   1987.826
      Y2              =   1987.826
   End
   Begin VB.Label Label3 
      Caption         =   " All rights reserved."
      Height          =   255
      Left            =   1380
      TabIndex        =   9
      Tag             =   "2061"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Tag             =   "755,756"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label lblProdCaption 
      AutoSize        =   -1  'True
      Caption         =   "This product is licensed to :"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Tag             =   "232"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright YrStart-YrEnd  Cubepoint, Inc."
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Tag             =   "759"
      Top             =   840
      Width           =   4035
   End
   Begin VB.Label lblApplicationName 
      AutoSize        =   -1  'True
      Caption         =   "APPLICATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   195
      Width           =   2010
   End
   Begin VB.Image imgAppIcon 
      Height          =   735
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblVersion 
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   1440
      TabIndex        =   3
      Tag             =   "229"
      Top             =   600
      Width           =   2445
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   150
      Picture         =   "frmAbout.frx":0121
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAbout As CAbout
Private m_frmInvokingForm As Form
Private m_clsTechnicalSupport As CIEExplore


Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFO = "PATH"
Private Const gREGVALSYSINFOLOC = "MSINFO"

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const READ_CONTROL = &H20000
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                               KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                               KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

Private Sub cmdTechSupport_Click()
    m_clsTechnicalSupport.OpenURL m_clsAbout.TechSupportURL, m_frmInvokingForm
End Sub

Private Sub Form_Load()
    
    Set Me.Icon = m_frmInvokingForm.Icon
    
    ConfigureAboutHeader
    
    If (m_clsAbout.IsDemoVersion = False) Then
        ConfigureAboutLicense
    Else
        ConfigureAboutActivateLicense
    End If
    
    Set m_clsTechnicalSupport = New CIEExplore
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_clsTechnicalSupport = Nothing
    Set m_clsAbout = Nothing
    Set m_frmInvokingForm = Nothing
    
End Sub

Public Function SetAbout(ByRef OwnerForm As Object, ByRef AboutInformation As CAbout)
    Dim strLicensee As String
    
    
    If ((TypeOf OwnerForm Is Form) = False) Then
        GoTo ERROR_TYPE_MISMATCH
    End If
    
    
    Set m_clsAbout = AboutInformation
    Set m_frmInvokingForm = OwnerForm
    
    
    Exit Function
    
ERROR_TYPE_MISMATCH:

    Err.Raise 1001, , "Type Mismatch.~Set About~"
    
End Function

Private Sub ConfigureAboutHeader()
    
    ' Set About image icon
    Set imgAppIcon.Picture = m_frmInvokingForm.Icon
    
    
    ' Set copyright details
    lblApplicationName.Caption = m_clsAbout.ApplicationName
    lblVersion.Caption = "Version " & m_clsAbout.VersionMajor & "." & m_clsAbout.VersionMinor & "." & m_clsAbout.VersionRevision
    lblCopyright.Caption = "Copyright " & m_clsAbout.CopyrightStart & "-" & m_clsAbout.CopyrightEnd & " " & m_clsAbout.CopyrightCompany
    
End Sub

Private Sub ConfigureAboutLicense()
    Dim lngCtr As Long
    
    
    'picLicensed.Visible = True
    'picDemo.Visible = False
    
    'picLicensed.Move 410, 1575, 4780, 2880
    
    
    lblLicensee.Caption = m_clsAbout.Licensee
    lblSerialNumber.Caption = "Serial number : " & m_clsAbout.SerialNumber
    
End Sub

Private Sub ConfigureAboutActivateLicense()
    
    lblProdCaption.Caption = "The product is not licensed."
    
    If (m_clsAbout.IsExpired = True) Then
        lblLicensee.Caption = "This demo version has expired."
    Else
        lblLicensee.Caption = "This demo version will expire on " & Format(m_clsAbout.ExpiryDate, "dd mmm yyyy") & "."
    End If
    
    lblSerialNumber.Visible = False
    
End Sub

Private Sub StartSysInfo()
    Dim strSysInfoPath As String
    Dim strFile As String
    
    
    On Error GoTo SysInfoErr
    
    
    ' -----> try to get system info program path\name from registry...
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, strSysInfoPath) = True) Then
    
    ' -----> try to get system info program path only from registry...
    ElseIf (GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, strSysInfoPath) = True) Then
        
        strFile = Dir(strSysInfoPath & "\MSINFO32.EXE")
        
        ' -----> validate existance of known 32-bit file version
        If (strFile <> "") Then
            
            strSysInfoPath = strSysInfoPath & "\MSINFO32.EXE"
            
        ' -----> error - file cannot be found...
        ElseIf (strFile = "") Then
            
            GoTo SysInfoErr
            
        End If
        
    ' -----> error - registry entry cannot be found...
    Else
        
        GoTo SysInfoErr
        
    End If
    
    Call Shell(strSysInfoPath, vbNormalFocus)
    
    Exit Sub
    
    
SysInfoErr:
    MsgBox "System information is unavailable at this time.", vbOKOnly, "System Error. (7001)"
    
End Sub

Private Function GetKeyValue(ByRef KeyRoot As Long, _
                             ByRef KeyName As String, _
                             ByRef SubKeyRef As String, _
                             ByRef KeyVal As String) As Boolean
    
    Dim i As Long           ' loop counter
    Dim rc As Long          ' return code
    Dim HKey As Long        ' handle to an open registry key
    Dim KeyValType As Long  ' data type of a registry key
    Dim tmpVal As String    ' tempory storage for a registry key value
    Dim KeyValSize As Long  ' size of registry key variable
    
    
    ' ------------------------------------------------------------
    ' open RegKey under key root {HKEY_LOCAL_MACHINE...}
    ' ------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, HKey) ' open registry key
    
    If (rc <> ERROR_SUCCESS) Then
    
        ' handle error...
        GoTo GetKeyError
        
    End If
    
    ' allocate variable space
    tmpVal = String$(1024, 0)
    
    ' mark variable size
    KeyValSize = 1024
    
    ' ------------------------------------------------------------
    ' retrieve registry key value...
    ' ------------------------------------------------------------
    
    ' get/create key value
    rc = RegQueryValueEx(HKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
                        
    If (rc <> ERROR_SUCCESS) Then
    
        ' handle errors
        GoTo GetKeyError
        
    End If
    
    ' windows 95 adds null terminated string...
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then
    
        ' null found, extract from string
        tmpVal = Left$(tmpVal, KeyValSize - 1)
        
    ' windows NT does not null terminate string...
    ElseIf (Asc(Mid$(tmpVal, KeyValSize, 1)) <> 0) Then
    
        ' null not found, extract string only
        tmpVal = Left$(tmpVal, KeyValSize)
        
    End If
    
    ' ------------------------------------------------------------
    ' determine key value type for conversion...
    ' ------------------------------------------------------------
    
    ' search data types...
    Select Case KeyValType
    
        ' string registry key data type
        Case REG_SZ
        
            ' copy string value
            KeyVal = tmpVal
        
        ' double word registry key data type
        Case REG_DWORD
        
            ' convert each bit
            For i = Len(tmpVal) To 1 Step -1
            
                ' build value character by character
                KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))
                
            Next i
            
            ' convert double word to string
            KeyVal = Format$("&h" + KeyVal)
            
    End Select
    
    ' return success
    GetKeyValue = True
    
    ' close registry key
    rc = RegCloseKey(HKey)
    
    Exit Function
    
    
' cleanup after an error has occured...
GetKeyError:

    ' set return Val to Empty string
    KeyVal = ""
    
    ' return failure
    GetKeyValue = False
    
    ' close registry key
    rc = RegCloseKey(HKey)
    
End Function
