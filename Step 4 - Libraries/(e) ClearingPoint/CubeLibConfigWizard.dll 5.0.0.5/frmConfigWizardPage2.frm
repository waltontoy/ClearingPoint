VERSION 5.00
Begin VB.Form frmConfigWizardPage2 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage2.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   1
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   3
         Left            =   0
         TabIndex        =   10
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
            Index           =   18
            Left            =   120
            TabIndex        =   16
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   15
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
            Index           =   20
            Left            =   120
            TabIndex        =   14
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   13
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
            Index           =   22
            Left            =   120
            TabIndex        =   12
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
            Index           =   23
            Left            =   120
            TabIndex        =   11
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   4320
         MaxLength       =   40
         TabIndex        =   0
         Top             =   240
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   1
         Top             =   600
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1320
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2760
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   8
         Top             =   3120
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1680
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   4320
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2040
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2400
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customs Reg. No."
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   25
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logical ID"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   24
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   23
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account 49"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   22
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sending Password"
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   21
         Top             =   2820
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Printing Password"
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   20
         Top             =   3180
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customs Office"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   19
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Language of Declaration"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   18
         Top             =   2100
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Place of Loading/Discharge"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   17
         Top             =   2460
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    DefLng A-Z
    
    Implements IWizardPage

Private Sub Form_Activate()
    Text1(0).SetFocus
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
    If NextStep = 3 Then
        If Len(Trim(Text1(0).Text)) = 0 Then
            MsgBox "Please enter the Company Name.", vbExclamation
            Text1(0).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.CompanyName = Trim(Text1(0).Text)
        End If
        
        If Len(Trim(Text1(1).Text)) = 0 Then
            MsgBox "Please enter the Logical ID.", vbExclamation
            Text1(1).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.LogicalID = Trim(Text1(1).Text)
        End If
        
        If Len(Trim(Text1(2).Text)) = 0 Then
            MsgBox "Please enter the Customs Registration Number.", vbExclamation
            Text1(2).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.CustomsRegNo = Trim(Text1(2).Text)
        End If
        
        If Len(Trim(Text1(3).Text)) = 0 Then
            MsgBox "Please enter the Account 49.", vbExclamation
            Text1(3).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.Account49 = Trim(Text1(3).Text)
        End If
        
        If Len(Trim(Text1(4).Text)) > 0 Then
            clsConfigWizard.CustomsOffice = Trim(Text1(4).Text)
        End If
        
        If Len(Trim(Text1(5).Text)) > 0 Then
            clsConfigWizard.LanguageofDeclaration = Trim(Text1(5).Text)
        End If
        
        If Len(Trim(Text1(6).Text)) > 0 Then
            clsConfigWizard.PlaceofLoading = Trim(Text1(6).Text)
        End If
        
        If Len(Trim(Text1(7).Text)) = 0 Then
            MsgBox "Please enter the Sending Password.", vbExclamation
            Text1(7).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.SendingPassword = Trim(Text1(7).Text)
        End If
        
        If Len(Trim(Text1(8).Text)) = 0 Then
            MsgBox "Please enter the Printing Password.", vbExclamation
            Text1(8).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.PrintingPassword = Trim(Text1(8).Text)
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
    If KeyAscii = 13 Then
        Select Case Index
            Case 0
                Text1(1).SetFocus
            Case 1
                Text1(2).SetFocus
            Case 2
                Text1(3).SetFocus
            Case 3
                Text1(4).SetFocus
            Case 4
                Text1(5).SetFocus
            Case 5
                Text1(6).SetFocus
            Case 6
                Text1(7).SetFocus
            Case 7
                Text1(8).SetFocus
            Case 8
                'Text1(9).SetFocus
        End Select
    End If
End Sub
