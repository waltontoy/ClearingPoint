VERSION 5.00
Begin VB.Form frmFilePath 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   3480
         Pattern         =   "*.mdb"
         TabIndex        =   7
         Top             =   600
         Width           =   1770
      End
      Begin VB.TextBox txtDBPath 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   4420
      End
      Begin VB.CommandButton cmdCreate 
         Height          =   315
         Left            =   3510
         Picture         =   "frmFilePath.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   390
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   240
         Width           =   2620
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   5415
         TabIndex        =   3
         Tag             =   "119"
         Top             =   1065
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   5415
         TabIndex        =   2
         Tag             =   "426"
         Top             =   630
         Width           =   1200
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   2610
      End
      Begin VB.Image imgFooter 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   2880
         Left            =   1215
         Stretch         =   -1  'True
         Top             =   2550
         Width           =   4032
      End
      Begin VB.Label Label4 
         Caption         =   "Preview"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   2550
         Width           =   990
      End
      Begin VB.Image imgDummy 
         Height          =   555
         Left            =   5670
         Top             =   1545
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Drive :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Folder :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Path :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
