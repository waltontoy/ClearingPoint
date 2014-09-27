VERSION 5.00
Begin VB.Form FLicenseUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Activation"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "FLicenseUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLicenseInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   105
      ScaleHeight     =   2640
      ScaleWidth      =   4740
      TabIndex        =   3
      Top             =   975
      Width           =   4740
      Begin VB.CommandButton cmdUpdateLicense 
         Caption         =   "&Update License"
         Default         =   -1  'True
         Height          =   375
         Left            =   3235
         TabIndex        =   15
         Top             =   2175
         Width           =   1400
      End
      Begin VB.ListBox lsbFeatures 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   1770
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1673
         Width           =   2865
      End
      Begin VB.TextBox txtSerialNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   233
         Width           =   1650
      End
      Begin VB.TextBox txtComputerID 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   593
         Width           =   2865
      End
      Begin VB.TextBox txtExpiryDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   953
         Width           =   2865
      End
      Begin VB.TextBox txtLicenseType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1313
         Width           =   2865
      End
      Begin VB.CommandButton cmdEditSerial 
         Caption         =   "&Change"
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   233
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "License expiry date :"
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label Label 
         Caption         =   "Licensed Features :"
         Height          =   210
         Index           =   5
         Left            =   150
         TabIndex        =   13
         Top             =   1725
         Width           =   1500
      End
      Begin VB.Label Label 
         Caption         =   "License Type :"
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   12
         Top             =   1365
         Width           =   1500
      End
      Begin VB.Label Label 
         Caption         =   "Computer ID :"
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   645
         Width           =   1500
      End
      Begin VB.Label Label 
         Caption         =   "Serial Number :"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   285
         Width           =   1500
      End
   End
   Begin VB.PictureBox picLicensee 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   105
      ScaleHeight     =   765
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   75
      Width           =   4740
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "This product is licensed to"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   75
         Width           =   4890
      End
      Begin VB.Label lblLicenseeName 
         Alignment       =   2  'Center
         Caption         =   "Cubepoint, Inc."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   1
         Top             =   435
         Width           =   4890
      End
   End
   Begin VB.Line Line 
      BorderStyle     =   6  'Inside Solid
      X1              =   105
      X2              =   4845
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "FLicenseUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAbout As CAbout
Private m_UpdateSuccessful As Boolean

Private Const MINIMUM_HEIGHT_FEATURES_LSB = 1125

Dim m_blnRunning As Boolean

Public Function ShowLicenseUpdate(ByRef AboutInformation As CAbout) As Boolean
    
    Set m_clsAbout = AboutInformation
    
    Me.Show vbModal
    
    ShowLicenseUpdate = m_UpdateSuccessful
    
End Function

Private Sub SetupLicenseUpdateForm()
    Dim lngCtr As Long
    
    
    lblLicenseeName.Caption = m_clsAbout.Licensee
    
    txtSerialNumber.Text = m_clsAbout.SerialNumber
    txtComputerID.Text = m_clsAbout.ComputerID
    txtExpiryDate.Text = Format(m_clsAbout.ExpiryDate, "dd mmm yyyy")
    
    If (m_clsAbout.LicenseType = Fixed) Then
        txtLicenseType.Text = "Fixed"
    Else
        txtLicenseType.Text = "Floating (" & m_clsAbout.AllowedUsers & " network users)"
    End If
    
    For lngCtr = 1 To m_clsAbout.ActiveFeatures.Count
        lsbFeatures.AddItem m_clsAbout.ActiveFeatures.Item(lngCtr).FeatureName
    Next lngCtr
    
    
    If (lsbFeatures.ListCount * 200 > MINIMUM_HEIGHT_FEATURES_LSB) Then
        lsbFeatures.Height = lsbFeatures.ListCount * 200
    Else
        lsbFeatures.Height = MINIMUM_HEIGHT_FEATURES_LSB
    End If
    
    If (m_clsAbout.LicenseExpires = False) Then
        lsbFeatures.Top = txtLicenseType.Top
        Label(5).Top = Label(4).Top
        
        txtLicenseType.Top = txtExpiryDate.Top
        Label(4).Top = Label(3).Top
        
        txtExpiryDate.Visible = False
        Label(3).Visible = False
    End If
    
    cmdUpdateLicense.Top = lsbFeatures.Top + lsbFeatures.Height + 100
    
    picLicenseInfo.Height = cmdUpdateLicense.Top + cmdUpdateLicense.Height + 100
    
    Me.Height = picLicenseInfo.Top + picLicenseInfo.Height + 600
    
End Sub

Private Sub Form_Load()
    
    SetupLicenseUpdateForm
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If (m_blnRunning = True) Then
        Me.WindowState = vbMinimized
        Cancel = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_clsAbout = Nothing
    
End Sub

Private Sub cmdUpdateLicense_Click()
    
    If (Len(Trim$(txtSerialNumber.Text)) <= 0) Then
        MsgBox "Please specify the serial number.", vbInformation, m_clsAbout.ApplicationName
        
        txtSerialNumber.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    m_blnRunning = True
    
    cmdUpdateLicense.Enabled = False
    cmdEditSerial.Enabled = False
    
    'DoEvents
    If (m_clsAbout.GetLicenseFile(txtSerialNumber.Text) = True) Then
        m_UpdateSuccessful = True
            
        m_blnRunning = False
        
        Screen.MousePointer = vbDefault
        
        Unload Me
        
        Exit Sub
        'If (txtSerialNumber.Locked = False) Then
        '    txtSerialNumber.Locked = True
        '    txtSerialNumber.TabStop = False
        '    txtSerialNumber.BackColor = txtComputerID.BackColor
        
        '    cmdUpdateLicense.Default = True
        '    cmdUpdateLicense.Enabled = False
        '    cmdEditSerial.Enabled = True
        '    cmdEditSerial.SetFocus
        'End If
    Else
        
        cmdEditSerial_Click
        
        'txtSerialNumber.SelStart = 0
        'txtSerialNumber.SelLength = Len(txtSerialNumber.Text)
        'txtSerialNumber.SetFocus
        
        cmdUpdateLicense.Enabled = True
        cmdEditSerial.Enabled = True
    End If
    
    m_blnRunning = False
    
    If (Me.WindowState = vbMinimized) Then
        Me.WindowState = vbNormal
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdEditSerial_Click()
    
    txtSerialNumber.Locked = False
    txtSerialNumber.TabStop = True
    txtSerialNumber.BackColor = vbWhite
    txtSerialNumber.SetFocus
    
    'txtSerialNumber_Change
    
End Sub

Private Sub txtComputerID_GotFocus()
    txtComputerID.SelStart = 0
    txtComputerID.SelLength = Len(txtComputerID.Text)
End Sub

Private Sub txtExpiryDate_GotFocus()
    txtExpiryDate.SelStart = 0
    txtExpiryDate.SelLength = Len(txtExpiryDate.Text)
End Sub

Private Sub txtLicenseType_GotFocus()
    txtLicenseType.SelStart = 0
    txtLicenseType.SelLength = Len(txtLicenseType.Text)
End Sub

Private Sub txtSerialNumber_Change()
    'If Len(txtSerialNumber.Text) = 0 Or txtSerialNumber.Locked = True Then
    '    cmdUpdateLicense.Enabled = False
    'Else
    '    cmdUpdateLicense.Enabled = True
    'End If
End Sub

Private Sub txtSerialNumber_GotFocus()
    
    txtSerialNumber.SelStart = 0
    txtSerialNumber.SelLength = Len(txtSerialNumber.Text)
    
    If (txtSerialNumber.Locked = False) Then
        cmdUpdateLicense.Default = False
    End If
    
End Sub

Private Sub txtSerialNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeyReturn) Then
        cmdUpdateLicense.SetFocus
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

Private Sub txtSerialNumber_LostFocus()
    
    On Error Resume Next
    
    If (txtSerialNumber.Locked = False) Then
        txtSerialNumber.Locked = True
        txtSerialNumber.TabStop = False
        txtSerialNumber.BackColor = txtComputerID.BackColor
        
        cmdUpdateLicense.Default = True
        cmdUpdateLicense.SetFocus
    End If
    
    On Error GoTo 0
    
End Sub
