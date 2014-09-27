VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ClearingPoint"
   ClientHeight    =   6075
   ClientLeft      =   4695
   ClientTop       =   3495
   ClientWidth     =   5610
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleMode       =   0  'User
   ScaleWidth      =   5599.997
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   100
      ScaleHeight     =   1290
      ScaleWidth      =   5415
      TabIndex        =   25
      Top             =   4650
      Width           =   5410
      Begin VB.CommandButton cmdTechSupport 
         Caption         =   "&Tech Support"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Tag             =   "2057"
         Top             =   32
         Width           =   1400
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Tag             =   "178"
         Top             =   882
         Width           =   1400
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&System Info"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Tag             =   "233"
         Top             =   457
         Width           =   1400
      End
      Begin VB.Label Label 
         Caption         =   $"FAbout.frx":058A
         Height          =   1170
         Index           =   3
         Left            =   50
         TabIndex        =   26
         Tag             =   "755,756"
         Top             =   59
         Width           =   3885
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   100
      ScaleHeight     =   1290
      ScaleWidth      =   5415
      TabIndex        =   20
      Top             =   150
      Width           =   5410
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   " All rights reserved."
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Tag             =   "2061"
         Top             =   1040
         Width           =   5385
      End
      Begin VB.Image imgAppIcon 
         Height          =   300
         Left            =   1350
         Picture         =   "FAbout.frx":06AB
         Stretch         =   -1  'True
         Top             =   60
         Width           =   300
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Caption         =   "Version X.X.X"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Tag             =   "229"
         Top             =   420
         Width           =   5385
      End
      Begin VB.Label lblApplicationName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Application"
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
         Left            =   1890
         TabIndex        =   22
         Top             =   30
         Width           =   1605
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         Caption         =   "Copyright 2001-2006 Cubepoint, Inc."
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Tag             =   "2061"
         Top             =   780
         Width           =   5385
      End
   End
   Begin VB.PictureBox picDemo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   100
      ScaleHeight     =   2880
      ScaleMode       =   0  'User
      ScaleWidth      =   5409.341
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1575
      Width           =   5410
      Begin VB.CommandButton cmdActivate 
         Caption         =   "&Activate"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1992
         TabIndex        =   1
         Top             =   2370
         Width           =   1400
      End
      Begin VB.TextBox txtSerialNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1935
         Width           =   3465
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Enter serial number to activate product"
         Height          =   210
         Index           =   2
         Left            =   300
         TabIndex        =   30
         Top             =   1575
         Width           =   4785
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "This demo version will expire on"
         Height          =   210
         Index           =   1
         Left            =   300
         TabIndex        =   29
         Top             =   795
         Width           =   4785
      End
      Begin VB.Label lblExpiryDateDemo 
         Alignment       =   2  'Center
         Caption         =   "17 October 2006"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   300
         TabIndex        =   28
         Top             =   1155
         Width           =   4785
      End
      Begin VB.Label lblComputerIDDemo 
         Alignment       =   2  'Center
         Caption         =   "Computer ID : 338689"
         Height          =   210
         Left            =   0
         TabIndex        =   27
         Top             =   75
         Width           =   5385
      End
   End
   Begin VB.PictureBox picLicensed 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   100
      ScaleHeight     =   2880
      ScaleWidth      =   5415
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Width           =   5410
      Begin VB.CommandButton cmdUpdateLicense 
         Caption         =   "&Update License"
         Height          =   375
         Left            =   3895
         TabIndex        =   2
         Top             =   150
         Width           =   1400
      End
      Begin VB.ListBox lsbFeatures 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   810
         Left            =   2025
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1890
         Width           =   3270
      End
      Begin VB.Label Label 
         Caption         =   "Licensed features :"
         Height          =   195
         Index           =   9
         Left            =   495
         TabIndex        =   17
         Tag             =   "2058"
         Top             =   1890
         Width           =   1395
      End
      Begin VB.Label lblExpiryDate 
         Caption         =   "17 October 2007"
         Height          =   195
         Left            =   2025
         TabIndex        =   16
         Tag             =   "2058"
         Top             =   1380
         Width           =   2895
      End
      Begin VB.Label lblLicenseType 
         Caption         =   "Floating (10 network users)"
         Height          =   195
         Left            =   2025
         TabIndex        =   15
         Tag             =   "2058"
         Top             =   1635
         Width           =   2895
      End
      Begin VB.Label lblComputerID 
         Caption         =   "338689"
         Height          =   195
         Left            =   2025
         TabIndex        =   14
         Tag             =   "2058"
         Top             =   1125
         Width           =   2895
      End
      Begin VB.Label lblSerialNumber 
         Caption         =   "Az02Pmf599"
         Height          =   195
         Left            =   2025
         TabIndex        =   13
         Tag             =   "2058"
         Top             =   855
         Width           =   2895
      End
      Begin VB.Label Label 
         Caption         =   "License type :"
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   12
         Tag             =   "2058"
         Top             =   1635
         Width           =   1395
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "This product is licensed to :"
         Height          =   210
         Index           =   4
         Left            =   495
         TabIndex        =   11
         Tag             =   "232"
         Top             =   75
         Width           =   4425
      End
      Begin VB.Label Label 
         Caption         =   "Expiry date :"
         Height          =   195
         Index           =   7
         Left            =   495
         TabIndex        =   10
         Tag             =   "2058"
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label 
         Caption         =   "Computer ID :"
         Height          =   195
         Index           =   6
         Left            =   495
         TabIndex        =   9
         Tag             =   "2058"
         Top             =   1125
         Width           =   1395
      End
      Begin VB.Label lblLicensee 
         AutoSize        =   -1  'True
         Caption         =   "Cubepoint, Inc."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   495
         TabIndex        =   8
         Tag             =   "2059"
         Top             =   360
         Width           =   4425
      End
      Begin VB.Label Label 
         Caption         =   "Serial Number : Az02Pmf599"
         Height          =   195
         Index           =   5
         Left            =   495
         TabIndex        =   7
         Tag             =   "2058"
         Top             =   855
         Width           =   1395
      End
   End
   Begin VB.Line Line 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   99.822
      X2              =   5500.175
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line Line 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   104.813
      X2              =   5505.167
      Y1              =   1500
      Y2              =   1500
   End
End
Attribute VB_Name = "FAbout"
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

Private m_blnActivating As Boolean


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
    
    
    ' Reposition copyright company and application icon
    lblApplicationName.Left = (lblVersion.Width / 2) - (lblApplicationName.Width / 2)
    imgAppIcon.Left = lblApplicationName.Left - imgAppIcon.Width - 50
    
End Sub

Private Sub ConfigureAboutLicense()
    Dim lngCtr As Long
    
    
    picLicensed.Visible = True
    picDemo.Visible = False
    
    lblLicensee.Caption = m_clsAbout.Licensee
    lblSerialNumber.Caption = m_clsAbout.SerialNumber
    lblComputerID.Caption = m_clsAbout.ComputerID
    lblExpiryDate.Caption = Format(m_clsAbout.ExpiryDate, "dd mmm yyyy")
    
    ' License type and number of allowed network users
    If (m_clsAbout.LicenseType = LicenseTypeConstants.Floating) Then
        lblLicenseType.Caption = "Floating (" & m_clsAbout.AllowedUsers & " network users)"
    Else
        lblLicenseType.Caption = "Fixed"
    End If
    
    ' Active features
    lsbFeatures.Clear
    For lngCtr = 1 To m_clsAbout.ActiveFeatures.Count
        lsbFeatures.AddItem m_clsAbout.ActiveFeatures.Item(lngCtr).FeatureName
    Next lngCtr
    
    
    If (m_clsAbout.LicenseExpires = False) Then
        lsbFeatures.Height = lsbFeatures.Height + (lsbFeatures.Top - lblLicenseType.Top)
        lsbFeatures.Top = lblLicenseType.Top
        Label(9).Top = Label(8).Top
        
        lblLicenseType.Top = lblExpiryDate.Top
        Label(8).Top = Label(7).Top
        
        lblExpiryDate.Visible = False
        Label(7).Visible = False
    End If
    
End Sub

Private Sub ConfigureAboutActivateLicense()
    
    'Resolve grammatical error on Demo version About Window License Info
    If m_clsAbout.IsDemoVersion = True Then
        If m_clsAbout.IsExpired = True Then
            Label(1).Caption = "This demo version expired last"
        Else
            Label(1).Caption = "This demo version will expire on"
        End If
    End If
    
    picLicensed.Visible = False
    picDemo.Visible = True
    
    lblComputerIDDemo.Caption = "Computer ID: " & m_clsAbout.ComputerID
    lblExpiryDateDemo.Caption = Format(m_clsAbout.ExpiryDate, "dd mmm yyyy")
    
End Sub

Private Sub Form_Load()
    Dim lngYear As Long
    Dim lngMonth As Long
    Dim lngDay As Long
    
    Set Me.Icon = m_frmInvokingForm.Icon
    
    ConfigureAboutHeader
    
    If (m_clsAbout.IsDemoVersion = False) Then
        ConfigureAboutLicense
    Else
        ConfigureAboutActivateLicense
    End If
    
    Set m_clsTechnicalSupport = New CIEExplore
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If (m_blnActivating = True) Then
        Me.WindowState = vbMinimized
        Cancel = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_clsTechnicalSupport = Nothing
    Set m_clsAbout = Nothing
    Set m_frmInvokingForm = Nothing
    
End Sub

Private Sub cmdActivate_Click()
    Dim blnActivateSuccessful As Boolean
    
    
    m_blnActivating = True
        Screen.MousePointer = vbHourglass
        
        cmdActivate.Enabled = False
        cmdOK.Enabled = False
        cmdSysInfo.Enabled = False
        cmdTechSupport.Enabled = False
        
        blnActivateSuccessful = m_clsAbout.GetLicenseFile(Trim$(txtSerialNumber.Text))
        
        cmdActivate.Enabled = True
        cmdOK.Enabled = True
        cmdSysInfo.Enabled = True
        cmdTechSupport.Enabled = True
        
        Screen.MousePointer = vbDefault
    m_blnActivating = False
    
    
    If (Me.WindowState = vbMinimized) Then
        Me.WindowState = vbNormal
    End If
    
    
    If (blnActivateSuccessful = False) Then
        txtSerialNumber.SelStart = 0
        txtSerialNumber.SelLength = Len(txtSerialNumber.Text)
        txtSerialNumber.SetFocus
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdUpdateLicense_Click()
    
    If (FLicenseUpdate.ShowLicenseUpdate(m_clsAbout)) Then
        Unload Me
    End If
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

Private Sub cmdTechSupport_Click()
    m_clsTechnicalSupport.OpenURL m_clsAbout.TechSupportURL, m_frmInvokingForm
End Sub

Private Sub txtSerialNumber_Change()
    
    If (Len(Trim$(txtSerialNumber.Text)) = 0) Then
        cmdActivate.Enabled = False
    Else
        cmdActivate.Enabled = True
    End If
    
End Sub

Private Sub txtSerialNumber_KeyPress(KeyAscii As Integer)
    
    ' Allow only characters: 0-9, A-Z, a-z
    
    If ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
        (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or _
        (KeyAscii >= Asc("a") And KeyAscii <= Asc("z"))) Or _
        KeyAscii = vbKeyBack Then
        
    Else
        
        KeyAscii = 0
        
    End If
    
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
