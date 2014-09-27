VERSION 5.00
Begin VB.Form FCubeLibLinkedDB 
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   4140
   ClientTop       =   2760
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5445
   Begin VB.CommandButton Command1 
      Caption         =   "Create Linked Tables"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FCubeLibLinkedDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' Create links to other databases - linked tables
    Dim clsLinks As ILinkedDB
    
    Dim objConProps As CConnectionProperties
    
    Set objConProps = GetDataSourceProperties(App.Path)
    
    Set clsLinks = New ILinkedDB
    clsLinks.CreateLinkedTables objConProps, 1
    Set clsLinks = Nothing
End Sub
