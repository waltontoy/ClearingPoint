VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDIPush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private FEDISend As FEDIComm

'local variable(s) to hold property value(s)
Private mvarAction As String 'local copy
Private mvarContentType As String 'local copy
Private mvarDestinationQueueName As String 'local copy
Private mvarDestinationQueuePassword As String 'local copy
Private mvarDestinationQueueRemoteHost As String 'local copy
Private mvarDestinationQueueRemotePort As String 'local copy
Private mvarDestinationQueueUsername As String 'local copy
Private mvarMessageType As String 'local copy
Private mvarMethod As String 'local copy
Private mvarSeparator As String 'local copy
Private mvarTimeoutInterval As Long 'local copy
Private mvarTraceOn As Boolean 'local copy
Private mvarUserName As String 'local copy
Private mvarVersion As String 'local copy
Private mvarPassword As String 'local copy
Private mvarDataReceived As String 'local copy
Private mvarForPLDAPushQueque As String
Private mvarForPLDA As Boolean
Private mvarPLDATestMessageOnly As Boolean

Private mvarSendPLDAToTestEnvironment As Long 'Edwin Dec28
'Private mvarSendPLDAToTestEnvironment As Boolean

Private mvarPLDALux As Boolean

Private mvarMessage As String

Public Event MessageReceived(ByVal Data As String)
Public Event SendExit(ByVal MessageSent As Boolean)
Public Event TraceMessage(ByVal TraceText As String)


' Used to Quit the Sending method
Private Sub ExitPush(ByVal MessageSent As Boolean)
    FEDISend.wskEDISend.Close
    FEDISend.tmrTimeoutSend.Enabled = False
    
    ' Gives information on whether the message has been sent successfully upon exit of sending
    ' TraceComm "S E N D  :  End Sending Message (" & IIf(MessageSent, "Message Sent)", "Message Not Sent)")
    Me.TraceText = "S E N D  :  End Sending Message (" & IIf(MessageSent, "Message Sent)", "Message Not Sent)")
    
    RaiseEvent SendExit(MessageSent)
End Sub

Friend Sub TimeOut()
    With FEDISend
        If .tmrTimeoutSend.Tag = ePushConnect Then
            ' TraceComm "S E N D  :  Connection timed out; no connection was established"
            ' TraceComm "S E N D  :  End EDI Send Connect Timeout"
            Me.TraceText = "S E N D  :  Connection timed out; no connection was established"
            Me.TraceText = "S E N D  :  End EDI Send Connect Timeout"
            
            
            Call ExitPush(False)
        ElseIf .tmrTimeoutSend.Tag = eSendMessage Then
            ' TraceComm "S E N D  :  Sending message timed out"
            Me.TraceText = "S E N D  :  Sending message timed out"
            .tmrTimeoutSend.Enabled = False
            
            ' TraceComm "S E N D  :  End Send Message Timeout"
            Me.TraceText = "S E N D  :  End Send Message Timeout"
            
            If Len(.mstrDataBuffer) Then
                ReceiveMessage .mstrDataBuffer, True
            Else
                Call ExitPush(True)
            End If
        End If
    End With
End Sub

' Used to make messages available received on data arrival
Friend Sub ReceiveMessage(ByVal Message As String, ByVal MessageSent As Boolean)
    ' TraceComm "S E N D  :  Start Processing Message"
    Me.TraceText = "S E N D  :  Start Processing Message"
    
    ' TraceComm "S E N D  :" & vbCrLf & Message
    Me.TraceText = "S E N D  :" & vbCrLf & Message
    
    RaiseEvent MessageReceived(Message)
    
    ' TraceComm "S E N D  :  End Processing Message"
    Me.TraceText = "S E N D  :  End Processing Message"
    
    Call ExitPush(MessageSent)
End Sub

' Method used by Projects using the component to send an EDI message
' Change argument SendPLDAToTestEnvironment from boolean to long
Public Sub SendMessage(ByVal Message As String, _
                        Optional ByVal PLDATestMessageOnly As Boolean = True, _
                        Optional ByVal SendPLDAToTestEnvironment As Long)
                        
    mvarPLDATestMessageOnly = PLDATestMessageOnly
    mvarSendPLDAToTestEnvironment = SendPLDAToTestEnvironment

    ' TraceComm "S E N D  :  Start Sending Message"
    Me.TraceText = "S E N D  :  Start Sending Message"
    
    mvarMessage = Message
    
    ' TraceComm "S E N D  :  Start Connection to Host"
    Me.TraceText = "S E N D  :  Start Connection to Host"
    
    FEDISend.wskEDISend.Connect mvarDestinationQueueRemoteHost, mvarDestinationQueueRemotePort
    
    ' TraceComm "S E N D  :  Start EDI Send Connect Timeout (interval in sec: " & FEDISend.mvarEDISend.TimeoutInterval & ")"
    Me.TraceText = "S E N D  :  Start EDI Send Connect Timeout (interval in sec: " & FEDISend.mvarEDISend.TimeoutInterval & ")"
    
    FEDISend.tmrTimeoutSend.Interval = 1000 * FEDISend.mvarEDISend.TimeoutInterval
    FEDISend.tmrTimeoutSend.Enabled = True
    FEDISend.tmrTimeoutSend.Tag = ePushConnect
End Sub

Public Property Let DataReceived(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataReceived = 5
    mvarDataReceived = vData
End Property

Public Property Get DataReceived() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataReceived
    DataReceived = mvarDataReceived
End Property

Public Property Get Message() As String
    Message = mvarMessage
End Property

Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property

Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property

Public Property Let Version(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Version = 5
    mvarVersion = vData
End Property

Public Property Get Version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property

Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property

Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property

Public Property Let TraceOn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TraceOn = 5
    mvarTraceOn = vData
    g_blnTraceCommEnabled = vData
End Property

Public Property Get TraceOn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TraceOn
    TraceOn = mvarTraceOn
End Property

Public Property Let TimeoutInterval(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeoutInterval = 5
    mvarTimeoutInterval = vData
End Property

Public Property Get TimeoutInterval() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeoutInterval
    TimeoutInterval = mvarTimeoutInterval
End Property

Public Property Let Separator(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Separator = 5
    mvarSeparator = vData
End Property

Public Property Get Separator() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Separator
    Separator = mvarSeparator
End Property

Public Property Let Method(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Method = 5
    mvarMethod = vData
End Property

Public Property Get Method() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Method
    Method = mvarMethod
End Property

Public Property Let MessageType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MessageType = 5
    mvarMessageType = vData
End Property

Public Property Get MessageType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MessageType
    MessageType = mvarMessageType
End Property

Public Property Let DestinationQueueUsername(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationQueueUsername = 5
    mvarDestinationQueueUsername = vData
End Property

Public Property Get DestinationQueueUsername() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationQueueUsername
    DestinationQueueUsername = mvarDestinationQueueUsername
End Property

Public Property Let DestinationQueueRemotePort(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationQueueRemotePort = 5
    mvarDestinationQueueRemotePort = vData
End Property

Public Property Get DestinationQueueRemotePort() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationQueueRemotePort
    DestinationQueueRemotePort = mvarDestinationQueueRemotePort
End Property

Public Property Let DestinationQueueRemoteHost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationQueueRemoteHost = 5
    mvarDestinationQueueRemoteHost = vData
End Property

Public Property Get DestinationQueueRemoteHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationQueueRemoteHost
    DestinationQueueRemoteHost = mvarDestinationQueueRemoteHost
End Property

Public Property Let DestinationQueuePassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationQueuePassword = 5
    mvarDestinationQueuePassword = vData
End Property

Public Property Get DestinationQueuePassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationQueuePassword
    DestinationQueuePassword = mvarDestinationQueuePassword
End Property

Public Property Let DestinationQueueName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationQueueName = 5
    mvarDestinationQueueName = vData
End Property

Public Property Get DestinationQueueName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationQueueName
    DestinationQueueName = mvarDestinationQueueName
End Property

Public Property Let ContentType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContentType = 5
    mvarContentType = vData
End Property

Public Property Get ContentType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContentType
    ContentType = mvarContentType
End Property

Public Property Let Action(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Action = 5
    mvarAction = vData
End Property

Public Property Get Action() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Action
    Action = mvarAction
End Property

Private Sub Class_Initialize()
    Set FEDISend = New FEDIComm
    
    mvarTraceOn = False
    mvarMessage = ""
    mvarForPLDAPushQueque = ""
    mvarForPLDA = False
    
    FEDISend.LoadWinsock Me
End Sub

Public Sub TerminateClass()
    On Error Resume Next
    Unload FEDISend
    On Error GoTo 0
    
    On Error Resume Next
    Set FEDISend = Nothing
    On Error GoTo 0
End Sub

Public Sub SetTracefilePath(ByVal strTracefilePath As String)
    TracefilePath = strTracefilePath
End Sub

''Private Sub ErrorEncountered(ByVal intError As Integer)
        
''    SElect Case
''
''sckOutOfMemory 7 Out of memory
''sckInvalidPropertyValue 380 The property value is invalid.
''sckGetNotSupported 394 The property can't be read.
''sckSetNotSupported 383 The property is read-only.
''sckBadState 40006 Wrong protocol or connection state for the requested transaction or request.
''sckInvalidArg 40014 The argument passed to a function was not in the correct format or in the specified range.
''sckSuccess 40017 Successful.
''sckUnsupported 40018 Unsupported variant type.
''sckInvalidOp 40020 Invalid operation at current state
''sckOutOfRange 40021 Argument is out of range.
''sckWrongProtocol 40026 Wrong protocol for the requested transaction or request
''sckOpCanceled 1004 The operation was canceled.
''sckInvalidArgument 10014 The requested address is a broadcast address, but flag is not set.
''sckWouldBlock 10035 Socket is non-blocking and the specified operation will block.
''sckInProgress 10036 A blocking Winsock operation in progress.
''sckAlreadyComplete 10037 The operation is completed. No blocking operation in progress
''sckNotSocket 10038 The descriptor is not a socket.
''sckMsgTooBig 10040 The datagram is too large to fit into the buffer and is truncated.
''sckPortNotSupported 10043 The specified port is not supported.
''sckAddressInUse 10048 Address in use.
''sckAddressNotAvailable 10049 Address not available from the local machine.
''sckNetworkSubsystemFailed 10050 Network subsystem failed.
''sckNetworkUnreachable 10051 The network cannot be reached from this host at this time.
''sckNetReset 10052 Connection has timed out when SO_KEEPALIVE is set.
''sckConnectAborted 11053 Connection is aborted due to timeout or other failure.
''sckConnectionReset 10054 The connection is reset by remote side.
''sckNoBufferSpace 10055 No buffer space is available.
''sckAlreadyConnected 10056 Socket is already connected.
''sckNotConnected 10057 Socket is not connected.
''sckSocketShutdown 10058 Socket has been shut down.
''sckTimedout 10060 Socket has been shut down.
''sckConnectionRefused 10061 Connection is forcefully rejected.
''sckNotInitialized 10093 WinsockInit should be called first.
''sckHostNotFound 11001 Authoritative answer: Host not found.
''sckHostNotFoundTryAgain 11002 Non-Authoritative answer: Host not found.
''sckNonRecoverableError 11003 Non-recoverable errors.
''sckNoData 11004 Valid name, no data record of requested type.
    

''End Sub

Public Property Let ForPLDAPushQueque(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataReceived = 5
    mvarForPLDAPushQueque = vData
End Property

Public Property Get ForPLDAPushQueque() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataReceived
    ForPLDAPushQueque = mvarForPLDAPushQueque
End Property

Public Property Let ForPLDA(ByVal blnForPLDA As Boolean)
    mvarForPLDA = blnForPLDA
End Property

Public Property Get ForPLDA() As Boolean
    ForPLDA = mvarForPLDA
End Property

Public Property Let TraceText(ByVal vNewValue As String)
    'TraceMessage = vNewValue
    RaiseEvent TraceMessage(vNewValue)
End Property

Public Property Get PLDATestMessageOnly() As Boolean
    PLDATestMessageOnly = mvarPLDATestMessageOnly
End Property

'Change from boolean to long - Edwin Dec28
Public Property Get SendPLDAToTestEnvironment() As Long
    SendPLDAToTestEnvironment = mvarSendPLDAToTestEnvironment
End Property

Public Property Let ForLux(ByVal blnNewValue As Boolean)
    mvarPLDALux = blnNewValue
End Property

Public Property Get ForLux() As Boolean
    ForLux = mvarPLDALux
End Property

