VERSION 5.00
Begin VB.Form FCubeLibPreCheck 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   3930
   ClientTop       =   2700
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6525
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "FCubeLibPreCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    Dim objConProps As CConnectionProperties
    
    Set objConProps = GetDataSourceProperties(App.Path)
    
    If (CheckFiles(App, Me, objConProps, G_CONST_PATH_ARG_1, G_CONST_PATH_ARG_2, G_CONST_PATH_ARG_3, G_CONST_PATH_ARG_4, G_CONST_PATH_ARG_5, G_CONST_PATH_ARG_7, G_CONST_PATH_ARG_8) = False) Then
        GoTo EarlyExit
    End If
    
EarlyExit:
    
End Sub
