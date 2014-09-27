VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Private mvarMAPIForm As frmMAPI

Public Function Send(ByVal RecipientAddr As String, ByVal Subject As String, ByVal BodyMessage As String, ByVal AttachmentPathName As String) As Boolean
    Dim mvarMapiMessages As MAPIMessages
    Dim mvarMapiSession As MAPISession
    
    Const SESSION_SIGNON = 1
    Const MESSAGE_COMPOSE = 6
    Const ATTACHTYPE_DATA = 0
    Const RECIPTYPE_TO = 1
    Const RECIPTYPE_CC = 2
    Const MESSAGE_RESOLVENAME = 13
    Const MESSAGE_SEND = 3
    Const SESSION_SIGNOFF = 2
    
    On Error GoTo Mailerror
    
    Send = False
    
    Set mvarMapiSession = mvarMAPIForm.MAPISession1
    Set mvarMapiMessages = mvarMAPIForm.MAPIMessages1
    
    mvarMapiSession.Action = SESSION_SIGNON
    
    With mvarMapiMessages
        .MsgIndex = -1
        .SessionID = mvarMapiSession.SessionID
        .Action = MESSAGE_COMPOSE
        
        .RecipIndex = 0
        .RecipType = RECIPTYPE_TO
        .RecipDisplayName = RecipientAddr
                
        .MsgSubject = Subject
        
        .MsgNoteText = BodyMessage


        .AttachmentType = ATTACHTYPE_DATA
        .AttachmentIndex = 0
        .AttachmentPosition = 0

        .AttachmentPathName = AttachmentPathName
        .AttachmentName = AttachmentPathName
           
        .Send True
     End With

    mvarMapiSession.Action = SESSION_SIGNOFF
    
    Set mvarMapiSession = Nothing
    Set mvarMapiMessages = Nothing
    
    Send = True
    
    Exit Function
    
Mailerror:
    Select Case Err.Number
        Case 32001 ' User cancelled process
            ' MsgBox Err.Description
        Case 32003 'login cancelled
            ' MsgBox Err.Description
        Case Else
             MsgBox CStr(Err.Number) + " " + Err.Description
    End Select
    
    If mvarMapiSession.SessionID <> 0 Then
        mvarMapiSession.SignOff   ' Signoff only if currently signed in.
    End If

    On Error GoTo 0
End Function

Private Sub Class_Initialize()
    Set mvarMAPIForm = New frmMAPI
End Sub

Private Sub Class_Terminate()
    Set mvarMAPIForm = Nothing
End Sub
