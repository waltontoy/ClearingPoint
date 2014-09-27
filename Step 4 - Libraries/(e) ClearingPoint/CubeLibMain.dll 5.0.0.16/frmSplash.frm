VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   1770
   ClientTop       =   1770
   ClientWidth     =   7590
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar pbrSplash 
      Height          =   120
      Left            =   315
      TabIndex        =   0
      Top             =   4935
      Visible         =   0   'False
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   1
      Max             =   5
   End
   Begin VB.Timer tmrLoad 
      Interval        =   50
      Left            =   315
      Top             =   4305
   End
   Begin VB.Image imgSplash 
      Height          =   5445
      Left            =   0
      Top             =   0
      Width           =   7590
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objSplashForm As Object

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error_Handler

12200   pbrSplash.Min = 0
12300   pbrSplash.Max = 100
12400   pbrSplash.ZOrder 1
12500   Me.Refresh
12600   tmrLoad.Enabled = True
        
        Exit Sub
        
Error_Handler:
    Err.Clear
End Sub

Public Function ShowSplash(ByVal Application As Object, _
                           ByRef SplashImage As Variant) _
                           As Object
    
    On Error GoTo Error_Handler
    
        '--->Load frmSplash
10400   Load Me
    
        Set m_objSplashForm = Me
        
10500   Me.imgSplash.Stretch = True

10600   If (IsObject(SplashImage) = True) Then
10700       If (TypeOf SplashImage Is Image) Then
10800           If (SplashImage.Picture Is Nothing = False) Then
10900               Set imgSplash.Picture = SplashImage
                    imgSplash.Refresh
                End If
            End If
        End If

        Me.Visible = True
        Me.Refresh
        '--->Show splash

11100   Me.Show

        Set ShowSplash = m_objSplashForm
        
        Exit Function

Error_Handler:
    Err.Clear
End Function

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Sub tmrLoad_Timer()
    On Error GoTo Error_Handler
     
    If pbrSplash.Value + 2 >= pbrSplash.Max Then
        tmrLoad.Enabled = False
        Exit Sub
    End If
10000   pbrSplash.Value = pbrSplash.Value + 2
        
        Exit Sub
        
Error_Handler:
    Resume
    Err.Clear
    
End Sub

Public Sub UnloadSplash()
    tmrLoad.Enabled = False
    Unload Me
End Sub
