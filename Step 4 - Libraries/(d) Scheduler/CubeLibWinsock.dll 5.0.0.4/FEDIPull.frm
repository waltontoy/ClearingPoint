VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FEDIPull 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   2355
   ClientTop       =   1935
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3795
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
   Begin VB.Timer tmrTimeoutPull 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   240
   End
   Begin MSWinsockLib.Winsock wskEDIPull 
      Left            =   600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FEDIPull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public mvarEDIPull As CEDIPull
Public mstrDataBuffer As String
Private mlngContentLength As Long
Private mlngContentLengthBuffer As Long

Public Sub LoadWinsock(ByRef EDIComm As CEDIPull)
    Load Me
    
    Set mvarEDIPull = EDIComm
End Sub

Private Sub tmrTimeoutPull_Timer()
    mvarEDIPull.TimeOut
End Sub

Private Sub wskEDIPull_Connect()
    tmrTimeoutPull.Enabled = False
    
    ' TraceComm "R E C E I V E  :  End EDI Pull Connect Timeout"
    mvarEDIPull.TraceText = "R E C E I V E  :  End EDI Pull Connect Timeout"
    
    With mvarEDIPull.Properties(mvarEDIPull.ActivePullQueue)
        PostPullMessage Me, ePostPull, .QueueName, .UserName, .Password, .RemoteHost
    End With
    
    ' TraceComm "R E C E I V E  :  Start Receive Message Timeout (interval in sec: " & mvarEDIPull.TimeoutInterval & ")"
    mvarEDIPull.TraceText = "R E C E I V E  :  Start Receive Message Timeout (interval in sec: " & mvarEDIPull.TimeoutInterval & ")"
    
    tmrTimeoutPull.Interval = 1000 * mvarEDIPull.TimeoutInterval
    tmrTimeoutPull.Enabled = True
    tmrTimeoutPull.Tag = eReceiveMessage
End Sub

Private Sub wskEDIPull_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strDataLines() As String
    Dim vntDataLine As Variant
    
    Dim strStartOfMessage As String
    Dim lngPartialMessageLength As Long
    Dim blnEndOfHeader As Boolean
    Dim blnQueueEmpty As Boolean
    Dim blnStopPullingMessages As Boolean
        
    ' The conditions to stop pulling messages will only be:
    '   1. The timer control 'tmrTimeoutPull' says the timeout
    '      interval has elapsed
    '   2. The previous queue is the last queue to receive messages from
    ' Therefore, stopping of pulling messages should not be handled in
    ' this procedure. Thus, [blnStopPullingMessages = False] = transfered from 1.00.25- joy 05/29/2006
1000    blnStopPullingMessages = False
    
    ' Initialize

    
    blnQueueEmpty = False
    
    wskEDIPull.GetData strData, vbString
    
    'To remove infinite loop when this kind of messages are received
    If (InStr(strData, "The JMS Queue does not exist") > 0) Or (InStr(strData, "SSH-2.0-Maverick_SSHD") > 0) Then
        mvarEDIPull.TraceText = "D A T A   A R R I V A L  :" & vbCrLf & strData
        mvarEDIPull.TimeOut
        Exit Sub
    End If
    
    'TraceComm "strData: " & strData
    'TraceComm "bytesTotal: " & bytesTotal
    
    ' Append data packet to mstrDataBuffer for each trigger of DataArrival event
    mstrDataBuffer = mstrDataBuffer & strData
    
    'TraceComm "mstrDataBuffer: " & mstrDataBuffer
    
    ' TraceComm "D A T A   A R R I V A L  :" & vbCrLf & strData
    mvarEDIPull.TraceText = "D A T A   A R R I V A L  :" & vbCrLf & strData
    
    If InStr(1, strData, "HTTP/1.1 100 Continue", vbTextCompare) > 0 Then
        ' ERE_HTTP_100_Continue
    ElseIf InStr(1, strData, "HTTP/1.1 204 The receive queue is empty, or the request timed out while waiting for the next message.", vbTextCompare) > 0 Then
        ' ERE_HTTP_204_Receive_Queue_Empty
        blnQueueEmpty = True
    ElseIf InStr(1, strData, "Invalid request: null", vbTextCompare) > 0 Then
        ' ERE_HTTP_Invalid_Queue_Name
    ElseIf InStr(1, strData, "Unauthorized", vbTextCompare) > 0 Then
        ' ERE_HTTP_Invalid_User_Name
    ElseIf InStr(1, strData, "HTTP/1.1 401 User not permitted to access this resource", vbTextCompare) > 0 Then
        
    ElseIf InStr(1, strData, "UNB+UNOC:3", vbTextCompare) > 0 Then
        ' ERE_HTTP_100_IE_Message
    ElseIf InStr(1, strData, "HTTP/1.1 200 OK", vbTextCompare) > 0 Then
        ' ERE_HTTP_200_OK
    Else
        ' ERE_Unknown
    End If
    
    strDataLines() = Split(strData, vbCrLf)
    
    ' Initialize Flag that inidicates the end of
    ' the Reply Message Header that arrived- Transferred from 1.00.25- joy 05/28/2006
    blnEndOfHeader = False

    For Each vntDataLine In strDataLines()
        'TraceComm "vntDataLine: " & vntDataLine
        
        If InStr(1, vntDataLine, "Content-Length:", vbTextCompare) Then
            mlngContentLength = CLng(Mid(vntDataLine, InStr(1, vntDataLine, "Content-Length:", vbTextCompare) + Len("Content-Length:") + 1))
            'TraceComm "mlngContentLength: " & mlngContentLength
        End If
        
        strStartOfMessage = vntDataLine
        ' There are 2 extra characters accounted for by the backend system
        ' indicated by the reply message line containing the label "Content-Length:",
        ' therefore we must offset the length of the message by 2.
        lngPartialMessageLength = Len(strStartOfMessage) + 2
        
        ' Check if the end of the Reply Message Header
        ' has been read -transferred from 1.00.25- joy 05/29/2006

        If blnEndOfHeader Then

            Exit For
        End If
        
        ' Check if the end of the Reply Message Header
        ' has not been read and the current line indicates
        ' the end of the Reply Message Header.
        If Not blnEndOfHeader And _
            lngPartialMessageLength = 2 Then
            
            ' Flag that the current line of the Reply Message
            ' is the end of the Reply Message Header-transferred from 1.00.24-joy 05/29/2006
            blnEndOfHeader = True
        End If
    Next
    
    ' Store the length of the complete reply EDI message that has been received in
    ' this data packet in case the EDI message does not fit in a single packet.
    mlngContentLengthBuffer = mlngContentLengthBuffer + lngPartialMessageLength
    
    If mlngContentLength > 0 Then
        ' If content-length is greater than actual length of characters
        ' received beyond header information, wait for next data packet
        ' which will be appended to mstrDataBuffer during next trigger
        ' of DataArrival event
        If mlngContentLength > mlngContentLengthBuffer Then
            vntDataLine = ""
            
            Erase strDataLines()
            
            ' Skip the following:
            '   1. stopping the timeout timer
            '   2. processing of the message received;
            '      and starting to issue another Pull Message request
            '   3. Re-initializing the following:
            '       a. mstrDataBuffer = ""
            '       b. mlngContentLength = 0
            '       c. mlngContentLengthBuffer = 0
            ' and wait for the next data packet which will be appended to
            ' mstrDataBuffer during next trigger of DataArrival event
            Exit Sub
        End If
    End If
    
    tmrTimeoutPull.Enabled = False
    
    ' TraceComm "R E C E I V E  :  A Reply has been received without timing out."
    mvarEDIPull.TraceText = "R E C E I V E  :  A Reply has been received without timing out."
    
    ' Note that line code 1000 sets blnStopPullingMessages = False
    ' and that blnQueueEmpty is True only when the data/reply that arrive is
    ' "HTTP/1.1 204 The receive queue is empty, or the request timed out while waiting for the next message."
    mvarEDIPull.ReceiveMessage mstrDataBuffer, mvarEDIPull.Properties(mvarEDIPull.ActivePullQueue).QueueName, _
                                blnQueueEmpty, blnStopPullingMessages
    
    ' Initialize for next message
    mstrDataBuffer = ""
    vntDataLine = ""
    
    mlngContentLength = 0
    mlngContentLengthBuffer = 0
    
    Erase strDataLines()
    
End Sub

Private Sub wskEDIPull_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    
    ' TraceComm "R E C E I V E  :  Number: " & Number & ", Description: " & Description & ", Scode: " & Scode & ", Source: " & Source
    mvarEDIPull.TraceText = "R E C E I V E  :  Number: " & Number & ", Description: " & Description & ", Scode: " & Scode & ", Source: " & Source
End Sub
