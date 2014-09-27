VERSION 5.00
Begin VB.Form frmLicenseeReminder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cubepoint Licensing"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      Begin VB.TextBox txtDays 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "30"
         Top             =   188
         Width           =   375
      End
      Begin VB.OptionButton optReminder 
         Caption         =   "Never remind me"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton optReminder 
         Caption         =   "Remind me again after"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "days"
         Height          =   210
         Left            =   2775
         TabIndex        =   6
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblReminderMessage 
      Alignment       =   2  'Center
      Caption         =   $"frmLicenseeReminder.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmLicenseeReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_frmOwnerForm As Form
Private m_strDBPathFileName As String
Private m_strLicFileName As String
Private m_strDBPassword As String

Private m_dteServerDate As Date
Private m_dteReminder As Date
Private m_lngDaysLeft As Long

Private m_objConnection As ADODB.Connection

Public Sub LicenseReminder(ByRef OwnerForm As Form, _
                           ByRef ADOConnection As ADODB.Connection, _
                           ByVal DBPathFileName As String, _
                           ByVal LicFileName As String, _
                           ByVal DBPassword As String)
    
    Set m_objConnection = ADOConnection
    
    Set m_frmOwnerForm = OwnerForm
    m_strDBPathFileName = DBPathFileName
    m_strLicFileName = LicFileName
    m_strDBPassword = DBPassword
    
    
    GetFileLastAccessedDate DBPathFileName, m_dteServerDate
    
    'rachelle 101706
    'If (g_typInterface.ILicense.IsDemo = True And ISLicenseExpiredOrClockTurnedBack(m_dteServerDate) = False) Then
    If ISLicenseExpiredOrClockTurnedBack(m_dteServerDate) = False Then
        If Not (g_typInterface.ILicense.ExpireMode = "N" And g_typInterface.ILicense.ExpireDateSoft = "0/0/0") Then
            m_dteServerDate = Format(m_dteServerDate, vbUseSystem)
            If g_typInterface.ILicense.UserDate(5) = "0/0/0" Then
                g_typInterface.ILicense.LFLock
                g_typInterface.ILicense.UserDate(5) = Format(CDate(Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat)) - 7, g_typInterface.ILicense.DateFormat)
                g_typInterface.ILicense.LFUnlock
            End If
            m_dteReminder = Format(g_typInterface.ILicense.UserDate(5), g_typInterface.ILicense.DateFormat)
            If m_dteServerDate >= m_dteReminder Then
                m_strLicFileName = LicFileName
                Me.Show vbModal
            End If
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.Icon = m_frmOwnerForm.Icon
    
    ' Compute days left
    m_lngDaysLeft = CDate(Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat)) - m_dteServerDate
    
    ' Set message
    lblReminderMessage.Caption = "This version of ClearingPoint is currently a demo version. You have " & IIf(m_lngDaysLeft <= 1, m_lngDaysLeft & " day", m_lngDaysLeft & " days") & " left to use this demo. Please activate your copy of ClearingPoint."
    
    txtDays.Text = 1
    optReminder(0).Value = True
    If (m_lngDaysLeft <= 0) Then
        txtDays.Enabled = False
        Frame1.Visible = False
        cmdActivate.Top = cmdActivate.Top - Frame1.Height
        cmdOk(0).Top = cmdOk(0).Top - Frame1.Height
        Me.Height = Me.Height - Frame1.Height
    End If
    
End Sub

Private Sub Form_Activate()
    If (txtDays.Enabled = True) Then
        txtDays.SetFocus
    End If
    SendKeysEx "{Home}+{End}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim dteNextCheck As Date
    
    
    Set m_frmOwnerForm = Nothing
    
    
    If optReminder(0).Value = True Then
        If Len(Trim(txtDays.Text)) > 0 Then
            
            If m_lngDaysLeft > 0 Then
                ' Checking
                If CLng(txtDays.Text) > m_lngDaysLeft Then
                    MsgBox "Please enter no. of days ranges from 0 to " & m_lngDaysLeft & " only.", vbInformation, Me.Caption
                    txtDays.SelStart = 0
                    txtDays.SelLength = Len(txtDays.Text)
                    txtDays.SetFocus
                    Cancel = True
                    Exit Sub
                End If
                ' Save to lic file
                dteNextCheck = m_dteServerDate + CLng(txtDays.Text)
                g_typInterface.ILicense.LFLock
                g_typInterface.ILicense.UserDate(5) = Format(dteNextCheck, g_typInterface.ILicense.DateFormat)
                g_typInterface.ILicense.LFUnlock
            End If
        Else
            MsgBox "Please enter no. of days from today.", vbInformation, Me.Caption
            txtDays.Text = "0"
            txtDays.SelStart = 0
            txtDays.SelLength = Len(txtDays.Text)
            txtDays.SetFocus
            Cancel = True
        End If
        
    ElseIf optReminder(1).Value = True Then
        g_typInterface.ILicense.LFLock
        g_typInterface.ILicense.UserDate(5) = Format(CDate(Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat)) + 1, g_typInterface.ILicense.DateFormat)
        g_typInterface.ILicense.LFUnlock
    End If
        
End Sub

Private Sub cmdActivate_Click()
    
    ' Place code here to activate product
    If (FActivate.Activate(m_frmOwnerForm, m_objConnection, Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat), m_strDBPathFileName, m_strLicFileName, , , m_strDBPassword) = True) Then
        Unload Me
    End If
    
End Sub

Private Sub cmdOK_Click(Index As Integer)
    
    If (Frame1.Visible = True) Then
        If (txtDays.Enabled = True) Then
            If (Len(Trim(txtDays.Text)) = 0) Then
                txtDays.Text = 0
            ElseIf CInt(Trim(txtDays.Text)) > m_lngDaysLeft Then
                MsgBox "The reminder's number of days cannot be greater than the License's number of days left.", vbInformation, "License Reminder"
                txtDays.SetFocus
                SendKeysEx "{Home}+{End}"
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub optReminder_Click(Index As Integer)
    
    If (optReminder(0).Value = True) Then
        If (m_lngDaysLeft <= 0) Then
            txtDays.Enabled = False
        Else
            txtDays.Enabled = True
        End If
    ElseIf optReminder(1).Value = True Then
        txtDays.Enabled = False
    End If
    
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    
    If Not (IsNumeric(Chr(KeyAscii)) = True Or KeyAscii = vbKeyBack) Then
       KeyAscii = 0
    End If
    
End Sub


