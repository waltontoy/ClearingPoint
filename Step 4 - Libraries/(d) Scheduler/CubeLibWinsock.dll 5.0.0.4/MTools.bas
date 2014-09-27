Attribute VB_Name = "MTools"
Option Explicit

Public g_blnTraceCommEnabled As Boolean
Private mstrTracefilePath As String

Public Enum PostType
    ePostPush = 1
    ePostPull = 2
End Enum

Public Enum TimeoutType
    ePullConnect = 0
    ePushConnect = 1
    eReceiveMessage = 2
    eSendMessage = 3
End Enum

'Send Mode Edwin Dec28
Public Enum SendMode
    eSendTestToTestEnvironment = 0
    eSendTestToPreProduction = 1
    eSendTestAndOperationalToTestEnvironment = 2
End Enum

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub TraceComm2(ByVal TraceString As String)
    Dim intFreeFile As Integer
    
    If g_blnTraceCommEnabled Then
        intFreeFile = FreeFile()
        
        If Len(Dir(mstrTracefilePath & "\TraceComm.txt")) Then
            If FileLen(mstrTracefilePath & "\TraceComm.txt") >= 120000 Then
                Name mstrTracefilePath & "\TraceComm.txt" As mstrTracefilePath & "\TraceComm" & Format(Now, "ddMMyyyyhhmm") & ".txt"
                Open mstrTracefilePath & "\TraceComm.txt" For Output As #intFreeFile
            Else
                Open mstrTracefilePath & "\TraceComm.txt" For Append As #intFreeFile
            End If
        Else
            Open mstrTracefilePath & "\TraceComm.txt" For Output As #intFreeFile
        End If
        
        Print #intFreeFile, Now & ": " & TraceString
        
        Close #intFreeFile
    End If
End Sub

Public Function GetSystemPath() As String
    Dim rc As Long
    Dim lpBuffer As String
    Dim nSize As Long
    
    nSize = 255
    lpBuffer = Space$(nSize)
    rc = GetSystemDirectory(lpBuffer, nSize)
    
    If rc <> 0 Then
        GetSystemPath = Left$(lpBuffer, rc)
    Else
        GetSystemPath = ""
    End If
End Function

'Change argument SendPLDAToTestEnvironment from boolean to long - Edwin Dec28
Public Function PostMessage(ByRef EDIComm As FEDIComm, _
                            ByVal PostType As PostType, _
                            ByVal Queue As String, _
                            ByVal UserName As String, _
                            ByVal Password As String, _
                            ByVal RemoteHost As String, _
                            Optional ByVal Message As String, _
                            Optional ByVal strForPLDAPushQueque As String, _
                            Optional ByVal ForPLDA As Boolean = False, _
                            Optional ByVal PLDATestMessageOnly As Boolean = True, _
                            Optional ByVal SendPLDAToTestEnvironment As Long, _
                            Optional ByVal ForLux As Boolean) As Boolean
    
    Dim strSend As String
    
    On Error GoTo ErrHandler
    
    ' Initialize Return Code
10  PostMessage = False
    
'15  TraceComm "P O S T  :  Start Posting Message"
15    EDIComm.mvarEDISend.TraceText = "P O S T  :  Start Posting Message"
    ' Send Method
20  strSend = "POST /jms http/1.1"
25  EDIComm.wskEDISend.SendData strSend & vbCrLf
'30  TraceComm "P O S T  :  " & strSend
30    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
    ' Send Action
35  Select Case PostType
        Case ePostPush
40          strSend = "X-JMS-Action: push-msg"
45      Case ePostPull
50          strSend = "X-JMS-Action: pull-msg"
    End Select
    
55  EDIComm.wskEDISend.SendData strSend & vbCrLf
' 60  TraceComm "P O S T  :  " & strSend
60    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
    ' Send Version
65  strSend = "X-JMS-Version: jmshttp/1.0"
70  EDIComm.wskEDISend.SendData strSend & vbCrLf
' 75  TraceComm "P O S T  :  " & strSend
75    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    ' Send Queue
80  Select Case PostType
        Case ePostPush
85          strSend = "X-JMS-DestinationQueue: " & Queue
90      Case ePostPull
95          strSend = "X-JMS-ReceiveQueue: " & Queue
    End Select
    
100 EDIComm.wskEDISend.SendData strSend & vbCrLf
'105 TraceComm "P O S T  :  " & strSend
105    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
    
    ' Send Message Type
110 strSend = "X-JMS-MessageType: text"
115 EDIComm.wskEDISend.SendData strSend & vbCrLf
' 120 TraceComm "P O S T  :  " & strSend
120    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
    
    ' Send Username
125 strSend = "X-JMS-User: " & UserName
130 EDIComm.wskEDISend.SendData strSend & vbCrLf
' 135 TraceComm "P O S T  :  " & strSend
135    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    ' Send Password
140 strSend = "X-JMS-Password: " & Password
145 EDIComm.wskEDISend.SendData strSend & vbCrLf
'150 TraceComm "P O S T  :  " & strSend
150    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    ' Send Content Type
155 strSend = "Content-Type: text/plain"
160 EDIComm.wskEDISend.SendData strSend & vbCrLf
'165 TraceComm "P O S T  :  " & strSend
165    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    If strForPLDAPushQueque <> "" Then
        strSend = "queue: " & strForPLDAPushQueque 'Modified by Migs 04/27/2006 changed Queue to queue... case sensitive
        EDIComm.wskEDISend.SendData strSend & vbCrLf
        ' TraceComm "P O S T  :  " & strSend
        EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
    End If
        
    '*********************************************************************************************
    'Modified to accomodate third send mode option - Edwin Dec28
    '*********************************************************************************************
    If ForPLDA Then
        Select Case SendPLDAToTestEnvironment
            Case SendMode.eSendTestToTestEnvironment
                strSend = "testmessage: " & IIf(PLDATestMessageOnly, "true", "false")
                EDIComm.wskEDISend.SendData strSend & vbCrLf
                EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                
                If Not ForLux Then
                    strSend = "pldatestprod: false"
                    
                    EDIComm.wskEDISend.SendData strSend & vbCrLf
                    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                End If
                
            Case SendMode.eSendTestToPreProduction
                strSend = "testmessage: " & IIf(PLDATestMessageOnly, "true", "false")
                EDIComm.wskEDISend.SendData strSend & vbCrLf
                EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                
                If Not ForLux Then
                    If Not PLDATestMessageOnly Then
                        strSend = "pldatestprod: false"
                    Else
                        strSend = "pldatestprod: true"
                    End If
                
                    EDIComm.wskEDISend.SendData strSend & vbCrLf
                    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                End If
                
            Case SendMode.eSendTestAndOperationalToTestEnvironment
                strSend = "testmessage: true"
                EDIComm.wskEDISend.SendData strSend & vbCrLf
                EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                
                If Not ForLux Then
                    strSend = "pldatestprod: false"
                    
                    EDIComm.wskEDISend.SendData strSend & vbCrLf
                    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
                End If
            
        End Select
     
     Else
        'FOR NCTS LUXEMBOURG
        If ForLux Then
            strSend = "testmessage: " & IIf(PLDATestMessageOnly, "true", "false")
            EDIComm.wskEDISend.SendData strSend & vbCrLf
            EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend
        End If
     End If
    '*********************************************************************************************
    
    ' Send Content Remote Host
170 strSend = "Host: " & EDIComm.wskEDISend.RemoteHost & ":" & EDIComm.wskEDISend.RemotePort
175 EDIComm.wskEDISend.SendData strSend & vbCrLf
'180 TraceComm "P O S T  :  " & strSend
180    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    ' Send Content Length
185 strSend = "Content-Length: " & CStr(Len(Message)) ' + 1) Modified by Migs 04/27/2006 removed +1
190 EDIComm.wskEDISend.SendData strSend & vbCrLf
'195 TraceComm "P O S T  :  " & strSend
195    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    
    ' Send Line Feed
200 EDIComm.wskEDISend.SendData vbCrLf
'205 TraceComm "P O S T  :  CRLF"
205    EDIComm.mvarEDISend.TraceText = "P O S T  :  CRLF"

    ' Send Message
210 strSend = Message
215 EDIComm.wskEDISend.SendData strSend
'220 TraceComm "P O S T  :  " & strSend
220    EDIComm.mvarEDISend.TraceText = "P O S T  :  " & strSend

    ' Send Line Feed
225 EDIComm.wskEDISend.SendData vbCrLf
'230 TraceComm "P O S T  :  CRLF"
230    EDIComm.mvarEDISend.TraceText = "P O S T  :  CRLF"

    ' Send Line Feed
235 EDIComm.wskEDISend.SendData vbCrLf
'240 TraceComm "P O S T  :  CRLF"
240    EDIComm.mvarEDISend.TraceText = "P O S T  :  CRLF"

'245 TraceComm "P O S T  :  End Posting Message"
245    EDIComm.mvarEDISend.TraceText = "P O S T  :  End Posting Message"

250 PostMessage = True
    
    Exit Function
    
ErrHandler:
    
    'TraceComm "Error in PostMessage (Number: " & Err.Number & ", Line: " & Erl & ", Source: " & Err.Source & ", Description: " & Err.Description & ", State: " & GetWinsockState(EDIComm.wskEDISend) & ", Queue: " & Queue & ")"
    EDIComm.mvarEDISend.TraceText = "Error in PostMessage (Number: " & Err.Number & ", Line: " & Erl & ", Source: " & Err.Source & ", Description: " & Err.Description & ", State: " & GetWinsockState(EDIComm.wskEDISend) & ", Queue: " & Queue & ")"
    
    Err.Clear
    
    On Error Resume Next
    
    EDIComm.wskEDISend.Close
End Function

Public Function PostPullMessage(ByRef EDIComm As FEDIPull, _
                                ByVal PostType As PostType, _
                                ByVal Queue As String, _
                                ByVal UserName As String, _
                                ByVal Password As String, _
                                ByVal RemoteHost As String, _
                                Optional ByVal Message As String) As Boolean
    
    Dim strSend As String
    
    On Error GoTo ErrHandler
    
    ' Initialize Return Code
10  PostPullMessage = False
    
'15  TraceComm "P O S T  :  Start Posting Message"
15    EDIComm.mvarEDIPull.TraceText = "P O S T  :  Start Posting Message"

    ' Send Method
20  strSend = "POST /jms http/1.1"
25  EDIComm.wskEDIPull.SendData strSend & vbCrLf
'30  TraceComm "P O S T  :  " & strSend
30    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Action
35  Select Case PostType
        Case ePostPush
40          strSend = "X-JMS-Action: push-msg"
45      Case ePostPull
50          strSend = "X-JMS-Action: pull-msg"
    End Select
    
55  EDIComm.wskEDIPull.SendData strSend & vbCrLf
'60  TraceComm "P O S T  :  " & strSend
60    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Version
65  strSend = "X-JMS-Version: jmshttp/1.0"
70  EDIComm.wskEDIPull.SendData strSend & vbCrLf
'75  TraceComm "P O S T  :  " & strSend
75    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Queue
80  Select Case PostType
        Case ePostPush
85          strSend = "X-JMS-DestinationQueue: " & Queue
90      Case ePostPull
95          strSend = "X-JMS-ReceiveQueue: " & Queue
    End Select
    
100 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'105 TraceComm "P O S T  :  " & strSend
105    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Message Type
110 strSend = "X-JMS-MessageType: text"
115 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'120 TraceComm "P O S T  :  " & strSend
120    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Username
125 strSend = "X-JMS-User: " & UserName
130 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'135 TraceComm "P O S T  :  " & strSend
135    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Password
140 strSend = "X-JMS-Password: " & Password
145 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'150 TraceComm "P O S T  :  " & strSend
150    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Content Type
155 strSend = "Content-Type: text/plain"
160 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'165 TraceComm "P O S T  :  " & strSend
165    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Content Remote Host
170 strSend = "Host: " & EDIComm.wskEDIPull.RemoteHost & ":" & EDIComm.wskEDIPull.RemotePort
175 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'180 TraceComm "P O S T  :  " & strSend
180    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Content Length
185 strSend = "Content-Length: " & CStr(Len(Message) + 1)
190 EDIComm.wskEDIPull.SendData strSend & vbCrLf
'195 TraceComm "P O S T  :  " & strSend
195    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Line Feed
200 EDIComm.wskEDIPull.SendData vbCrLf
'205 TraceComm "P O S T  :  CRLF"
205    EDIComm.mvarEDIPull.TraceText = "P O S T  :  CRLF"

    ' Send Message
210 strSend = Message
215 EDIComm.wskEDIPull.SendData strSend
'220 TraceComm "P O S T  :  " & strSend
220    EDIComm.mvarEDIPull.TraceText = "P O S T  :  " & strSend

    ' Send Line Feed
225 EDIComm.wskEDIPull.SendData vbCrLf
'230 TraceComm "P O S T  :  CRLF"
230    EDIComm.mvarEDIPull.TraceText = "P O S T  :  CRLF"

    ' Send Line Feed
235 EDIComm.wskEDIPull.SendData vbCrLf
'240 TraceComm "P O S T  :  CRLF"
240    EDIComm.mvarEDIPull.TraceText = "P O S T  :  CRLF"

'245 TraceComm "P O S T  :  End Posting Message"
245    EDIComm.mvarEDIPull.TraceText = "P O S T  :  End Posting Message"

250 PostPullMessage = True
    
    Exit Function

ErrHandler:
    
    ' TraceComm "Error in PostPullMessage (Number: " & Err.Number & ", Line: " & Erl & ", Source: " & Err.Source & ", Description: " & Err.Description & ", State: " & GetWinsockState(EDIComm.wskEDIPull) & ", Queue: " & Queue & ")"
    EDIComm.mvarEDIPull.TraceText = "Error in PostPullMessage (Number: " & Err.Number & ", Line: " & Erl & ", Source: " & Err.Source & ", Description: " & Err.Description & ", State: " & GetWinsockState(EDIComm.wskEDIPull) & ", Queue: " & Queue & ")"
    
    Err.Clear
    
    On Error Resume Next
    
    EDIComm.wskEDIPull.Close
End Function

Public Function GetWinsockState(ByRef Sock As Winsock) As String
    Select Case Sock.State
        Case sckClosed              ' 0 Default. Closed
            GetWinsockState = "sckClosed"
        Case sckOpen                ' 1 Open
            GetWinsockState = "sckOpen"
        Case sckListening           ' 2 Listening
            GetWinsockState = "sckListening"
        Case sckConnectionPending   ' 3 Connection pending
            GetWinsockState = "sckConnectionPending"
        Case sckResolvingHost       ' 4 Resolving host
            GetWinsockState = "sckResolvingHost"
        Case sckHostResolved        ' 5 Host resolved
            GetWinsockState = "sckHostResolved"
        Case sckConnecting          ' 6 Connecting
            GetWinsockState = "sckConnecting"
        Case sckConnected           ' 7 Connected
            GetWinsockState = "sckConnected"
        Case sckClosing             ' 8 Peer is closing the connection
            GetWinsockState = "sckClosing"
        Case sckError               ' 9 Connection Error
            GetWinsockState = "sckError"
        Case Else                   ' X State Else
            GetWinsockState = "sckStateElse"
    End Select
End Function

Public Property Let TracefilePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TracefilePath = strTracefilePath
    mstrTracefilePath = vData
End Property

Public Property Get TracefilePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TracefilePath
    TracefilePath = mstrTracefilePath
End Property
