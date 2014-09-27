VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_objApplication As Object
Private m_objOwnerForm As Object

Private m_blnSecurityOn As Boolean
Private m_strUser As String

Public Function ShowLogin(ByRef MainProps As CMainControls, _
                            ByVal Application As Object, _
                            ByRef OwnerForm As Object, _
                            ByRef ADOConnection As ADODB.Connection, _
                            ByRef LicFilePath As String) As QueryResultConstants
    
    Set m_objApplication = Application
    Set m_objOwnerForm = OwnerForm
    
    InitializeLogin m_objApplication, OwnerForm
    
    ShowLogin = frmLogin.Login(MainProps, _
                                Application, _
                                OwnerForm, _
                                ADOConnection, _
                                m_blnSecurityOn, _
                                m_strUser)
End Function


Private Sub InitializeLogin(ByVal Application As Object, _
                            ByRef OwnerForm As Object)
    Dim clsRegistry As CRegistry
    
    On Error GoTo Error_Handler
10100   Set clsRegistry = New CRegistry
        
        '--->Check if 'Security' key exists in the application's registry settings
10300   If (clsRegistry.RegistryKeyExists(cpiCurrentUser, Application.ProductName, "Settings", "Security") = True) Then
            
            '--->Get security setting
10600       clsRegistry.GetRegistry cpiCurrentUser, Application.ProductName, "Settings", "Security"
            
            '--->Check if security is enabled
10800       m_blnSecurityOn = IIf(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) = "On", True, False)
            
            '--->Check if 'UserName' key exists in the application's registry settings
11000       If (clsRegistry.RegistryKeyExists(cpiCurrentUser, Application.ProductName, "Settings", "UserName") = True) Then
                
                '--->Retrieve last user who logged in
11200           clsRegistry.GetRegistry cpiCurrentUser, Application.ProductName, "Settings", "UserName"
11300           m_strUser = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
            Else
11500           m_blnSecurityOn = False
                
            End If
            
        Else
            '--->Create 'Security' key in the application's registry settings
11800       clsRegistry.SaveRegistry cpiCurrentUser, Application.ProductName, "Settings", "Security", Encrypt("On", KEY_ENCRYPT)
11900       m_blnSecurityOn = True
            
        End If
        
12000   Set clsRegistry = Nothing

    Exit Sub
        
Error_Handler:
    Err.Clear
End Sub

Public Sub UnloadLogin()
    On Error Resume Next
    
    If Not frmLogin Is Nothing Then
        Unload frmLogin
    End If
    Set frmLogin = Nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Unload frmLogin
    Set frmLogin = Nothing
End Sub
