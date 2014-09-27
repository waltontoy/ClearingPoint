VERSION 5.00
Begin VB.Form frmConfigWizardPage4 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage4.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   3
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2040
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Company Data"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   13
            Top             =   816
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     User Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   12
            Top             =   1392
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Connections"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   11
            Top             =   1968
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Programs"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   10
            Top             =   2544
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "Finish"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   9
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "DEXXDATACOMSYS1"
         Top             =   720
         Width           =   2010
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "172"
         Top             =   390
         Width           =   360
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4920
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "31"
         Top             =   390
         Width           =   360
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "1"
         Top             =   390
         Width           =   360
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "10"
         Top             =   390
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   2010
      End
      Begin VB.ComboBox cboPrinters 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   1440
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gateway Name"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   19
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   18
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      .     .      ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   16
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Printer Name"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   15
         Top             =   1500
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    DefLng A-Z
    
    Implements IWizardPage
    
    Private m_blnBackToThisPage As Boolean

Private Sub Form_Activate()
    Text2(0).SetFocus
End Sub

Private Sub Form_Load()
    Dim prnPrinter As Printer
    
    For Each prnPrinter In Printers
        cboPrinters.AddItem prnPrinter.DeviceName
    Next
    
    If cboPrinters.ListCount Then
        cboPrinters.Text = cboPrinters.List(0)
    End If
    
    m_blnBackToThisPage = False
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
    If NextStep = 5 Then
        If Trim(Text1(1).Text) = "" Then
            MsgBox "Please enter the User Name.", vbInformation + vbOKOnly
            
            Text1(1).Text = ""
            
            Text1(1).SetFocus
            Cancel = True
            m_blnBackToThisPage = True
            
            Exit Sub
        Else
            clsConfigWizard.IPAddress = Trim(Text2(0).Text) & "." & Trim(Text2(1).Text) & "." & Trim(Text2(2).Text) & "." & Trim(Text2(3).Text)
            clsConfigWizard.GatewayName = Trim(Text1(0).Text)
            clsConfigWizard.UserName = Trim(Text1(1).Text)
            clsConfigWizard.PrinterName = cboPrinters.Text
        End If
    End If
End Sub

Private Sub IWizardPage_BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Len(Text1(Index).Text) > 0 Then
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 0 Then Text1(1).SetFocus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 1 And cboPrinters.Visible = True And m_blnBackToThisPage = False Then
        cboPrinters.SetFocus
    ElseIf m_blnBackToThisPage = True Then
        m_blnBackToThisPage = False
        Text1(Index).SetFocus
    End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    If Len(Text2(Index).Text) > 0 Then
        Text2(Index).SelStart = 0
        Text2(Index).SelLength = Len(Text2(Index).Text)
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0
                Text2(1).SetFocus
            Case 1
                Text2(2).SetFocus
            Case 2
                Text2(3).SetFocus
            Case 3
                Text1(0).SetFocus
        End Select
    End If
End Sub
