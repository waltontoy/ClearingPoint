VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDIPull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CEDIProperties"
Option Explicit
    
Private FEDIReceive As FEDIPull

Public Enum EPullQueue
    ENone = 0
    EMessages = 1
    EErrors = 2
End Enum

Private mvarAction As String 'local copy
Private mvarContentType As String 'local copy
Private mvarMessageType As String 'local copy
Private mvarMethod As String 'local copy
Private mvarPassword As String 'local copy
Private mvarTimeoutInterval As String 'local copy
Private mvarTraceOn As Boolean 'local copy
Private mvarUserName As String 'local copy
Private mvarVersion As String 'local copy
Private mvarSeparator As String 'local copy
Private mvarActivePullQueue As Long
Private mvarProperties As CEDIProperties
Private mvarSlowPollCycles As Long 'local copy
Private mvarFastPollCycles As Long 'local copy
Private mvarForPLDA As Boolean 'local copy
Private mvarDType As Long

Public Event MessageReceived(ByVal Data As String, ByVal PullQueue As String)
Public Event PullExit(ByVal QueuesEmptied As Boolean)
Public Event TraceMessage(ByVal TraceText As String)

Public Property Let FastPollCycles(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FastPollCycles = 5
    mvarFastPollCycles = vData
End Property

Public Property Get FastPollCycles() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FastPollCycles
    FastPollCycles = mvarFastPollCycles
End Property

Public Property Let SlowPollCycles(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlowPollCycles = 5
    mvarSlowPollCycles = vData
End Property

Public Property Get SlowPollCycles() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlowPollCycles
    SlowPollCycles = mvarSlowPollCycles
End Property

Public Property Get Properties() As CEDIProperties
    If mvarProperties Is Nothing Then
        Set mvarProperties = New CEDIProperties
    End If

    Set Properties = mvarProperties
End Property

Public Property Set Properties(vData As CEDIProperties)
    Set mvarProperties = vData
End Property

Friend Sub TimeOut()
    Dim blnQueueEmptiedAtTimeOut As Boolean
    Dim blnStopPullingMessages As Boolean
    Dim blnAllQueuesEmptied As Boolean
    
        ' Stop Pulling Messages because a timeout occurred
1000    blnStopPullingMessages = True
    
    With FEDIReceive
        ' Check if the Timeout Occurred while Connecting to the Host
        If .tmrTimeoutPull.Tag = ePullConnect Then
            ' TraceComm "R E C E I V E  :  Connection timed out; no connection was established"
            Me.TraceText = "R E C E I V E  :  Connection timed out; no connection was established"
            
            ' TraceComm "R E C E I V E  :  End Timing to Check for EDI Connect Action to Timeout"
            Me.TraceText = "R E C E I V E  :  End Timing to Check for EDI Connect Action to Timeout"
            
            ' Still Attempting to Connect at Timeout,
            ' therefore it is undetermined if all Queues
            ' are empty because they haven't even been checked.
            ' It is assumed here that the Queues are not empty
            blnAllQueuesEmptied = False
            
        ElseIf .tmrTimeoutPull.Tag = eReceiveMessage Then
            ' TraceComm "R E C E I V E  :  Connection timed out while receiving a message."
            Me.TraceText = "R E C E I V E  :  Connection timed out while receiving a message."
            
            'TraceComm "R E C E I V E  :  End Receive Message Timeout"
            Me.TraceText = "R E C E I V E  :  End Receive Message Timeout"
            
            ' Check if a Message has been Received While the Timeout Occurred
            blnQueueEmptiedAtTimeOut = (Len(.mstrDataBuffer) <= 0)
            If Not blnQueueEmptiedAtTimeOut Then
                ' To ensure that no other actions will be done while processing the
                ' Reply Message Receievd at Timeout, the Timout timer is disabled
                .tmrTimeoutPull.Enabled = False
            
                ' Process the Message that has been Received While the Timeout Occurred
                ' Note:
                '   1. blnQueueEmptiedAtTimeOut is FALSE at this Point
                '      because the condition to get here is (If Not blnQueueEmptiedAtTimeOut Then),
                '      i.e. the ActiveQueue is not Empty at Timeout.
                '   2. blnStopPullingMessages is TRUE because Receiving of Messages
                '      will be stopped because a timeout occurred.
                '      Line Code 1000 sets blnStopPullingMessages = True
                ReceiveMessage .mstrDataBuffer, Me.Properties(Me.ActivePullQueue).QueueName, _
                                blnQueueEmptiedAtTimeOut, blnStopPullingMessages
            End If
            
            ' If the ActiveQueue has not been emptied,
            ' then All Queues have not been emptied
            blnAllQueuesEmptied = blnQueueEmptiedAtTimeOut
        End If
        
        ' Stop Pulling Messages from Queues.
        ' Notify Clients if the Queues have been Emptied at timeout
        ExitPull blnAllQueuesEmptied
    End With
End Sub

Friend Sub PullMessage(ByVal NextQueue As Boolean)
    Dim blnAllQueuesEmptied As Boolean
    
    ' Check if the Next Queue Needs to be Activated
    If NextQueue Then
        ' Activate the Next Queue to Receive Messages From
        mvarActivePullQueue = mvarActivePullQueue + 1
    End If
    
    ' Check if the active/activated queue is valid. If it is not,
    ' then the Previous Queue is the Last Queue to Receive Messages From
    blnAllQueuesEmptied = (mvarActivePullQueue > mvarProperties.Count)
    If Not blnAllQueuesEmptied Then
        With mvarProperties.Item(mvarActivePullQueue)
            'TraceComm "R E C E I V E  :  Start Checking Queue (" & .QueueName & ") - " & IIf(.Disabled, "Disabled", "Enabled")
            Me.TraceText = "R E C E I V E  :  Start Checking Queue (" & .QueueName & ") - " & IIf(.Disabled, "Disabled", "Enabled")
            
            Do While True
                ' Check if the Receiving Messages from the Active Queue is Enabled and
                ' if the the Active Queue is a Valid Queue (i.e. tha previous queue is not
                ' the last queue to receive messages from)
                If mvarActivePullQueue <= mvarProperties.Count And _
                    Not .Disabled Then
                    
                    On Error Resume Next
                    FEDIReceive.wskEDIPull.Close
                    On Error GoTo 0
                    
                    ' TraceComm "R E C E I V E  :  Start Connection to Host"
                    Me.TraceText = "R E C E I V E  :  Start Connection to Host"
                    
                    ' Attempt to connect to the host. If successful, the event Connect which
                    ' will be triggered will execute a Pull Message Request sent to the host
                    FEDIReceive.wskEDIPull.Connect .RemoteHost, .RemotePort
                    
                    ' TraceComm "R E C E I V E  :  Start Timing to Check for EDI Connect Action to Timeout (interval in sec: " & FEDIReceive.mvarEDIPull.TimeoutInterval & ")"
                    Me.TraceText = "R E C E I V E  :  Start Timing to Check for EDI Connect Action to Timeout (interval in sec: " & FEDIReceive.mvarEDIPull.TimeoutInterval & ")"
                    
                    FEDIReceive.tmrTimeoutPull.Interval = 1000 * FEDIReceive.mvarEDIPull.TimeoutInterval
                    FEDIReceive.tmrTimeoutPull.Enabled = True
                    FEDIReceive.tmrTimeoutPull.Tag = ePullConnect
                                
                    ' Stop going through the queues to receive messages from and wait for the
                    ' reply from the Pull Message Request called in the event Connect triggered
                    ' by [FEDIReceive.wskEDIPull.Connect .RemoteHost, .RemotePort]
                    Exit Do
                    
                ' Check if the Receiving Messages from the Active Queue is Enabled and
                ' if the the Active Queue is a Valid Queue (i.e. tha previous queue is not
                ' the last queue to receive messages from)
                ElseIf mvarActivePullQueue <= mvarProperties.Count And .Disabled Then
                    ' Receiving Messages from the Active Queue is Disabled,
                    ' therefore, Process Messages from the Next Queue
                    
                    ' TraceComm "R E C E I V E  :  Skip Checking Queue (" & .QueueName & ")"
                    Me.TraceText = "R E C E I V E  :  Skip Checking Queue (" & .QueueName & ")"
                    
                    ' Activate the Next Queue to Receive Messages From
                    mvarActivePullQueue = mvarActivePullQueue + 1
                                
                ' Check if the active/activated queue is valid. If it is not,
                ' then the Previous Queue is the Last Queue to Receive Messages From
                ElseIf mvarActivePullQueue > mvarProperties.Count Then
                    ' The previous queue is the last queue to receive messages from.
                    ' Therefore, all queues have been emptied and receiving messages
                    ' should be stopped, i.e. blnAllQueuesEmptied should be TRUE at this point.
                    blnAllQueuesEmptied = True
                
                    ExitPull blnAllQueuesEmptied
                    
                    Exit Do
                End If
            Loop
        End With
    Else
        ' blnAllQueuesEmptied is TRUE at this point since the
        ' previous queue is the last queue to receive messages from.
        ' Therefore, we need to stop receiving messages.
        ExitPull blnAllQueuesEmptied
    End If
End Sub

Public Sub PullMessages()
    ' TraceComm "R E C E I V E  :  Start Receiving Messages"
    Me.TraceText = "R E C E I V E  :  Start Receiving Messages"
    
    Call PullMessage(False)
End Sub

' Stop Receiving Messages from the Queues
Private Sub ExitPull(ByVal AllQueuesEmptied As Boolean)
    ' Close Communication Port for Receiving Messages from Queues
    FEDIReceive.wskEDIPull.Close
    
    ' Stop Timeout Timer from Running
    FEDIReceive.tmrTimeoutPull.Enabled = False
    
    ' Reset Action Timeout Inidicator
    FEDIReceive.tmrTimeoutPull.Tag = 0
    
    ' TraceComm "R E C E I V E  :  End Receiving Messages"
    Me.TraceText = "R E C E I V E  :  End Receiving Messages"
    
    ' Notify Clients that Receiving Messages from Queues has been Stopped
    RaiseEvent PullExit(AllQueuesEmptied)
End Sub

Friend Sub ReceiveMessage(ByVal Message As String, _
                            ByVal PullQueue As String, _
                            ByVal QueueEmpty As Boolean, _
                            Optional ByVal StopPullingMessages As Boolean = False)
    Dim blnPullFromNextQueue As Boolean
    
    blnPullFromNextQueue = QueueEmpty
    
    ' TraceComm "R E C E I V E  :  Start Processing Message"
    Me.TraceText = "R E C E I V E  :  Start Processing Message"
    
    ' TraceComm "R E C E I V E  :" & vbCrLf & Message
    Me.TraceText = "R E C E I V E  :" & vbCrLf & Message
    
    'RaiseEvent MessageReceived(Message, PullQueue)     Migs 06-20-2006 commented placed at the end of the procedure. This causes the time out interval to reset kaya inilipat sa baba
    
    ' TraceComm "R E C E I V E  :  End Processing Message"
    Me.TraceText = "R E C E I V E  :  End Processing Message"
    
    If Not StopPullingMessages Then
        Select Case blnPullFromNextQueue
            Case True
                ' TraceComm "R E C E I V E  :  End Checking Queue (" & PullQueue & ")"
                Me.TraceText = "R E C E I V E  :  End Checking Queue (" & PullQueue & ")"
                
            Case False
                
        End Select
                
        ' Pull Another Message from the Active Queue if the
        ' Active Queue has not been Emptied Yet
        PullMessage IIf(blnPullFromNextQueue, True, False)
    Else
        ' TraceComm "R E C E I V E  :  End Checking Queue (" & PullQueue & ")"
        Me.TraceText = "R E C E I V E  :  End Checking Queue (" & PullQueue & ")"
        
        ' TraceComm "R E C E I V E  :  End Checking All Queues (" & PullQueue & ")"
        Me.TraceText = "R E C E I V E  :  End Checking All Queues (" & PullQueue & ")"
    End If
    
    RaiseEvent MessageReceived(Message, PullQueue)  'Migs 06-20-2006 this code was moved mula sa taas...
    
End Sub

Public Property Get ActivePullQueue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActivePullQueue
    ActivePullQueue = mvarActivePullQueue
End Property

Public Property Get Separator() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Separator
    Separator = mvarSeparator
End Property

Public Property Get Version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property

Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property

Public Property Let TraceOn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TraceOn = True
    mvarTraceOn = vData
    g_blnTraceCommEnabled = vData
End Property

Public Property Get TraceOn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TraceOn
    TraceOn = mvarTraceOn
End Property

Public Property Let TimeoutInterval(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeoutInterval = "5"
    mvarTimeoutInterval = vData
End Property

Public Property Get TimeoutInterval() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeoutInterval
    TimeoutInterval = mvarTimeoutInterval
End Property

Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property

Public Property Get Method() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Method
    Method = mvarMethod
End Property

Public Property Get MessageType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MessageType
    MessageType = mvarMessageType
End Property

Public Property Get ContentType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContentType
    ContentType = mvarContentType
End Property

Public Property Get Action() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Action
    Action = mvarAction
End Property

Private Sub Class_Initialize()
    Set FEDIReceive = New FEDIPull
    
    mvarActivePullQueue = 1
    mvarTraceOn = False
    mvarForPLDA = False
    mvarDType = 0
    
    FEDIReceive.LoadWinsock Me
End Sub

Public Sub TerminateClass()
    On Error Resume Next
    Unload FEDIReceive
    On Error GoTo 0
    
    On Error Resume Next
    Set FEDIReceive = Nothing
    On Error GoTo 0
    
    Set mvarProperties = Nothing
End Sub

Public Sub SetTracefilePath(ByVal strTracefilePath As String)
    TracefilePath = strTracefilePath
End Sub

Public Property Let ForPLDA(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FastPollCycles = 5
    mvarForPLDA = vData
End Property

Public Property Get ForPLDA() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FastPollCycles
    ForPLDA = mvarForPLDA
End Property

Public Property Let DType(ByVal vData As Long)
    mvarDType = vData
End Property

Public Property Get DType() As Long
    DType = mvarDType
End Property

Public Property Let TraceText(ByVal vNewValue As String)
    'TraceMessage = vNewValue
    RaiseEvent TraceMessage(vNewValue)
End Property
