VERSION 5.00
Begin VB.Form frmConfigWizardPage3 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage3.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   2
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   2
         Left            =   0
         TabIndex        =   4
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
            Index           =   12
            Left            =   120
            TabIndex        =   10
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
            Index           =   13
            Left            =   120
            TabIndex        =   9
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   8
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
            Index           =   15
            Left            =   120
            TabIndex        =   7
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
            Index           =   16
            Left            =   120
            TabIndex        =   6
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
            Index           =   17
            Left            =   120
            TabIndex        =   5
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   4320
         MaxLength       =   25
         TabIndex        =   0
         Top             =   360
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   4320
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4320
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Administrator Name"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   12
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   1140
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage3"
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

Private Sub Form_Load()
    Label1(0).Caption = g_strAdminUserLabel
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
    If NextStep = 4 Then
        If Trim(Text1(1).Text) <> Trim(Text1(2).Text) Then
            MsgBox "The password and confirm password do not match.  Please type them again.", vbExclamation
            
            Text1(1).Text = ""
            Text1(2).Text = ""
            
            Text1(1).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            clsConfigWizard.Password = Trim(Text1(1).Text)
        End If
        
        If Len(Trim(Text1(0).Text)) = 0 Then
            MsgBox "Please enter the Administrator Name.", vbExclamation
            Text1(0).SetFocus
            Cancel = True
            
            Exit Sub
        Else
            If Not ValidUserName Then
                Text1(0).SetFocus
                Cancel = True
                
                Exit Sub
            Else
                clsConfigWizard.AdministratorName = Trim(Text1(0).Text)
            End If
        End If
    End If

End Sub

Private Sub IWizardPage_BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub

Private Function ValidUserName() As Boolean
    With g_rstUser
        If Not (.EOF And .BOF) Then
            .MoveFirst
            .Find "[User_Name] = '" & Text1(0).Text & "' ", , adSearchForward
            If Not .EOF Then
                MsgBox "User Name already exists!", vbInformation
            Else
                ValidUserName = True
            End If
        Else
            ValidUserName = True
        End If
        
'        .Index = "User_Name"
'        .Seek "=", Text1(0).Text
'
'        If Not .NoMatch Then
'            MsgBox "User Name already exists!", vbInformation
'        Else
'            ValidUserName = True
'        End If
    End With
End Function

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
        End Select
    End If
End Sub
