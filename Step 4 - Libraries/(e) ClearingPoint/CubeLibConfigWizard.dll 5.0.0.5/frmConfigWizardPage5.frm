VERSION 5.00
Begin VB.Form frmConfigWizardPage5 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage5.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   4
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame2 
         Caption         =   "Printing Program"
         Height          =   1365
         Index           =   0
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "APRINT / ATPRIN"
            Height          =   240
            Index           =   0
            Left            =   1320
            TabIndex        =   0
            Top             =   315
            Value           =   -1  'True
            Width           =   1740
         End
         Begin VB.OptionButton Option1 
            Caption         =   "BPRINT / BTPRIN"
            Height          =   240
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            Top             =   630
            Width           =   1740
         End
         Begin VB.OptionButton Option1 
            Caption         =   "XPRINT / TPRINT"
            Height          =   240
            Index           =   2
            Left            =   1320
            TabIndex        =   2
            Top             =   930
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   4
         Left            =   0
         TabIndex        =   9
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
            Index           =   0
            Left            =   120
            TabIndex        =   15
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
            Index           =   1
            Left            =   120
            TabIndex        =   14
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
            Index           =   2
            Left            =   120
            TabIndex        =   13
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
            Index           =   3
            Left            =   120
            TabIndex        =   12
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
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
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sending Program"
         Height          =   1230
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   4095
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   14
            Left            =   1320
            TabIndex        =   3
            Text            =   "TESSAD"
            Top             =   315
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   15
            Left            =   1320
            TabIndex        =   4
            Text            =   "TPSAD"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Test"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   8
            Top             =   375
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Operational"
            Height          =   195
            Index           =   9
            Left            =   210
            TabIndex        =   7
            Top             =   780
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    DefLng A-Z
    
    Implements IWizardPage

Private Sub Form_Activate()
    Option1(0).SetFocus
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
    If Option1(0).Value = True Then
        clsConfigWizard.PrintingProgram = 0
    ElseIf Option1(1).Value = True Then
        clsConfigWizard.PrintingProgram = 1
    ElseIf Option1(2).Value = True Then
        clsConfigWizard.PrintingProgram = 2
    End If
End Sub

Private Sub IWizardPage_BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub
