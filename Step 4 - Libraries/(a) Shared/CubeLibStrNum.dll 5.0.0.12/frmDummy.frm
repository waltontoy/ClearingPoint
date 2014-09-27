VERSION 5.00
Begin VB.Form frmDummy 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin VB.DirListBox dirDummy 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.FileListBox flbDummy 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

