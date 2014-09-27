VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "CubeLibWinset"
   ClientHeight    =   4950
   ClientLeft      =   5160
   ClientTop       =   3675
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7635
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objConProps As CConnectionProperties

Private Sub Form_Activate()
    Dim clsWinSettings As PCubeLibWinSet.IWindows
    
    
    
    Set m_objConProps = GetDataSourceProperties(App.Path)
    
    ADOConnectDB g_conTemplateCP, m_objConProps, DBInstanceType_DATABASE_TEMPLATE
    
    Set clsWinSettings = New PCubeLibWinSet.IWindows
    clsWinSettings.LoadWindowSettings g_conTemplateCP, G_CONST_USER_ID, G_CONST_WINDOW_KEY, Me
    Set clsWinSettings = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim clsWinSettings As PCubeLibWinSet.IWindows
        
    Set clsWinSettings = New PCubeLibWinSet.IWindows
    clsWinSettings.SaveWindowSettings g_conTemplateCP, G_CONST_USER_ID, G_CONST_WINDOW_KEY, Me
    Set clsWinSettings = Nothing
    
    ADODisconnectDB g_conTemplateCP
    
    Set m_objConProps = Nothing
End Sub
