VERSION 5.00
Begin VB.Form FLocatePaths 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paths Not Found"
   ClientHeight    =   2850
   ClientLeft      =   3105
   ClientTop       =   2070
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetPaths 
      Caption         =   "Set Paths..."
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Tag             =   "781"
      Top             =   2400
      Width           =   1355
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "&Retry"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Tag             =   "781"
      Top             =   2400
      Width           =   1355
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Tag             =   "781"
      Top             =   2400
      Width           =   1355
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Required file/s are not available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "  Make sure that:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "3. if the file/s were moved to another location,  the ""Set Paths..."" button can be used to specify the new value/s."
         Height          =   480
         Index           =   4
         Left            =   600
         TabIndex        =   6
         Tag             =   "780"
         Top             =   1560
         Width           =   4635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1. the network server is available."
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Tag             =   "778"
         Top             =   1080
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2. your workstation has the correct drive mapping to the server."
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Tag             =   "779"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FLocatePaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mvarOWnerForm As Object
    Dim mvarApplication As Object
    Dim mvarMissingPathsStream As String
    
    Dim enuReturn As CheckResult
    
Private Sub cmdCancel_Click()
    enuReturn = cpiCancel
    
    Unload Me
End Sub

Private Sub cmdRetry_Click()
    enuReturn = cpiRetry
    
    Unload Me
End Sub

Private Sub cmdSetPaths_Click()
    enuReturn = cpiSetPath
    
    Unload Me
End Sub

Public Function OnLog(ByRef OwnerForm As Object, ByVal Application As Object, ByVal MissingPathsStream As String) As CheckResult
    Set mvarOWnerForm = OwnerForm
    Set mvarApplication = Application
    
    mvarMissingPathsStream = MissingPathsStream
    enuReturn = cpiCancel
    
    Set Me.Icon = OwnerForm.Icon
    Me.Show vbModal
    
    OnLog = enuReturn
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mvarOWnerForm = Nothing
    Set mvarApplication = Nothing

End Sub
