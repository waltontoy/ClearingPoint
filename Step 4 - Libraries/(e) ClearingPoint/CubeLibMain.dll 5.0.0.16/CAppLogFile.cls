VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAppLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
    Dim mvarblnInLogSession As Boolean
    Dim mvarlngFileNumber As Long
    Dim mvarAppPath As String
    Dim mvarAppName As String
    Dim mvarFileName As String
    Dim mvarUserID As Long
    Dim mvarLogOn As Boolean
    
Public Sub StartLogFile(ByVal AppName As String, ByVal AppPath As String, ByVal UserID As Long, Optional ByVal LogOn As Boolean = False)
    
    If mvarblnInLogSession Then
        Close #mvarlngFileNumber
        
        mvarblnInLogSession = False
    End If
    
    mvarLogOn = LogOn
        
    mvarAppPath = AppPath
    mvarAppName = AppName
    mvarUserID = UserID
    
    If mvarLogOn Then
        'Removed date to prevent creation of too many trace files in the db path
        'mvarFileName = AppPath & "\" & AppName & Format(CStr(Date), "MMDDYY") & "Usr" & UserID & ".trc"
        mvarFileName = AppPath & "\" & AppName & "Usr" & UserID & ".trc"
        
        mvarlngFileNumber = FreeFile
        
            Open mvarFileName For Append Lock Read Write As #mvarlngFileNumber
            
            Print #mvarlngFileNumber, " "
            Print #mvarlngFileNumber, " START OF LOG " & Date & " - " & Time
            Print #mvarlngFileNumber, " "
            
'            If IsUserAnAdmin = 1 Then
'                Print #mvarlngFileNumber, " USER IS AN ADMINISTRATOR = TRUE"
'            Else
'                Print #mvarlngFileNumber, " USER IS AN ADMINISTRATOR = FALSE"
'            End If
        
        mvarblnInLogSession = True
        
        Close #mvarlngFileNumber
    End If
    
End Sub

Private Sub Class_Initialize()
    mvarblnInLogSession = False
    mvarlngFileNumber = -1
    mvarLogOn = False
End Sub

Private Sub Class_Terminate()
    If mvarblnInLogSession Then
        mvarblnInLogSession = False
    End If
End Sub

Public Function AppendLog(ByVal LogTrace As String) As Boolean
    If mvarblnInLogSession = False Then
        AppendLog = False
    ElseIf mvarLogOn Then
        mvarlngFileNumber = FreeFile
        
            Open mvarFileName For Append Lock Read Write As #mvarlngFileNumber
        
            Print #mvarlngFileNumber, LogTrace
        
            Close #mvarlngFileNumber
        
        AppendLog = True
    End If
End Function

Public Sub CloseLogFile()
    If mvarblnInLogSession Then
        mvarlngFileNumber = FreeFile
        
            Open mvarFileName For Append Lock Read Write As #mvarlngFileNumber
            
            Print #mvarlngFileNumber, " "
            Print #mvarlngFileNumber, " END OF LOG " & Date & " - " & Time
            Print #mvarlngFileNumber, " "
            
            Close #mvarlngFileNumber
           
        mvarblnInLogSession = False
    End If
End Sub

