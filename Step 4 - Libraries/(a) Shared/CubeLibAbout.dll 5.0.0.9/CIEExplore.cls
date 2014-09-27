VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIEExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim intResult As Long
    
Private Const SW_SHOWNORMAL = 1
    
Public Sub OpenURL(ByVal URL As String, ByVal OwnerForm As Object)
Attribute OpenURL.VB_Description = "Opens a URL."
Attribute OpenURL.VB_HelpID = 11
    
    Dim lngRetVal As Long
    Dim lngOpenResult As Long
    
    lngRetVal = ShellExecute(OwnerForm.hWnd, "OPEN", URL, "", "C:\", SW_SHOWNORMAL)
 
    ' check return value
    If (lngRetVal < 31) Then
    
        ' error code - see api help for more info
        MsgBox App.Title & " had a problem running your web browser." & _
                "You should check that your browser is correctly installed." & _
                "(Error" & Format$(lngRetVal) & ")", vbExclamation, "Browser Unavailable (7011)"
    
    ElseIf (lngRetVal = 32) Then
    
        ' -----> no file association
        MsgBox App.Title & " could not find a file association for " & _
                    URL & " on your system. You should check that your browser" & _
                    "is correctly installed and associated with this type of file.", vbExclamation, "Browser Unavailable (7012)"
    
    Else
        
        ' -----> it worked!
 
    End If

End Sub
