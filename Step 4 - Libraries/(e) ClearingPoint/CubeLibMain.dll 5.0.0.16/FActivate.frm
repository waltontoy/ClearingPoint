VERSION 5.00
Begin VB.Form FActivate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activation Form"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   135
      TabIndex        =   1
      Top             =   60
      Width           =   4215
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtSerialNumber 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   338
         TabIndex        =   0
         Top             =   1515
         Width           =   3465
      End
      Begin VB.CommandButton cmdActivate 
         Caption         =   "&Activate"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblExpiryDate 
         Alignment       =   2  'Center
         Caption         =   "17 October 2006"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   6
         Top             =   585
         Width           =   3990
      End
      Begin VB.Label lblLicenseInfo 
         Alignment       =   2  'Center
         Caption         =   "This demo version will expire on"
         Height          =   210
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   3990
      End
      Begin VB.Label lblSerialNumber 
         Alignment       =   2  'Center
         Caption         =   "Enter serial number to activate product"
         Height          =   210
         Left            =   75
         TabIndex        =   4
         Top             =   1155
         Width           =   3990
      End
   End
End
Attribute VB_Name = "FActivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_blnActivating As Boolean
Private m_clsAbout As CAbout

Private m_frmCallingform As Object
Private m_strApplicationName As String
Private m_dteExpiryDate As Date
Private m_strTempateFileName As String
Private m_strLicenseFileName As String
Private m_blnActivateOnlyComputer As Boolean
Private m_strComputerID As String
Private m_strDBPassword As String

Private m_blnActivateSuccessful As Boolean

Private m_objConnection As ADODB.Connection

Public Function Activate(ByRef CallingForm As Object, _
                         ByRef ADOConnection As ADODB.Connection, _
                         ByVal ExpiryDate As Date, _
                         ByVal TempateFileName As String, _
                         ByVal LicenseFileName As String, _
                Optional ByVal ActivateOnlyComputer As Boolean = False, _
                Optional ByVal ComputerID As String, _
                Optional ByVal DBPassword As String) As Boolean
    
    Set m_objConnection = ADOConnection
    
    Set m_frmCallingform = CallingForm
    m_dteExpiryDate = ExpiryDate
    m_strTempateFileName = TempateFileName
    m_strLicenseFileName = LicenseFileName
    m_blnActivateOnlyComputer = ActivateOnlyComputer
    m_strComputerID = ComputerID
    m_strDBPassword = DBPassword
    
    Me.Show vbModal
    
    Activate = m_blnActivateSuccessful
    
End Function

Private Sub Form_Load()
    
    Set m_clsAbout = New CAbout
    
    If (m_blnActivateOnlyComputer = True) Then
        lblLicenseInfo.Caption = "Computer ID :"
        lblExpiryDate.Caption = m_strComputerID
        lblSerialNumber.Caption = "Enter serial number to activate computer."
    Else
        lblLicenseInfo.Caption = "This application will expire on :"
        lblExpiryDate.Caption = Format(m_dteExpiryDate, "dd mmmm yyyy")
        lblSerialNumber.Caption = "Enter serial number to activate product."
    End If
    
    cmdActivate.Enabled = False
    
    Set Me.Icon = m_frmCallingform.Icon
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If (m_blnActivating = True) Then
        Me.WindowState = vbMinimized
        Cancel = True
    End If
    
End Sub

Private Sub cmdActivate_Click()
    Dim conTemplate As ADODB.Connection
    
    
    m_blnActivating = True
        Screen.MousePointer = vbHourglass
        cmdActivate.Enabled = False
        
        ADOConnectDB m_objConnection, g_objDataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
        
        'ConnectDB conTemplate, m_strTempateFileName, , m_strDBPassword
        
        m_blnActivateSuccessful = m_clsAbout.GetLicenseFile(txtSerialNumber.Text, conTemplate, m_frmCallingform, m_strLicenseFileName)
        
        ADODisconnectDB conTemplate
        
        cmdActivate.Enabled = True
        Screen.MousePointer = vbDefault
    m_blnActivating = False
    
    
    If (Me.WindowState = vbMinimized) Then
        Me.WindowState = vbNormal
    End If
    
    If (m_blnActivateSuccessful = False) Then
        txtSerialNumber.SelStart = 0
        txtSerialNumber.SelLength = Len(txtSerialNumber.Text)
        txtSerialNumber.SetFocus
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdCancel_Click()
    
    m_blnActivateSuccessful = False
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_frmCallingform = Nothing
    Set m_clsAbout = Nothing
    
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
