VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPicture As Variant
Private m_objSplashForm As Object

Public Function Splash(ByVal Application As Object, _
                        ByRef SplashImageControl As Object) _
                        As Object
    Set mvarPicture = SplashImageControl
    
    Set m_objSplashForm = frmSplash.ShowSplash(Application, mvarPicture)
    
    Set Splash = m_objSplashForm
End Function

Public Sub UnloadSplash()
    On Error Resume Next
    If Not frmSplash Is Nothing Then
        Unload frmSplash
    End If
    Set frmSplash = Nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Unload frmSplash
    Set frmSplash = Nothing
End Sub
