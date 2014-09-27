VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FEDIComm 
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   3330
   ClientTop       =   3075
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   2430
   Visible         =   0   'False
   Begin VB.Timer tmrTimeoutSend 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wskEDISend 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FEDIComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public mvarEDISend As CEDIPush
Public mstrDataBuffer As String
Private mlngContentLength As Long
Private mlngContentLengthBuffer As Long

Public Sub LoadWinsock(ByRef EDIComm As CEDIPush)
    Load Me
    
    Set mvarEDISend = EDIComm
End Sub

Private Sub tmrTimeoutSend_Timer()
    mvarEDISend.TimeOut
End Sub

Private Sub wskEDISend_Connect()
    tmrTimeoutSend.Enabled = False
    
    ' TraceComm "S E N D  :  End EDI Send Connect Timeout"
    mvarEDISend.TraceText = "S E N D  :  End EDI Send Connect Timeout"
    
    With mvarEDISend
        Call PostMessage(Me, ePostPush, .DestinationQueueName, .DestinationQueueUsername, _
                            .DestinationQueuePassword, .DestinationQueueRemoteHost, .Message, _
                            .ForPLDAPushQueque, .ForPLDA, .PLDATestMessageOnly, .SendPLDAToTestEnvironment, .ForLux)
        
        ' TraceComm "S E N D  :  Start Send Message Timeout (interval in sec: " & .TimeoutInterval & ")"
        mvarEDISend.TraceText = "S E N D  :  Start Send Message Timeout (interval in sec: " & .TimeoutInterval & ")"
        
        tmrTimeoutSend.Interval = 1000 * .TimeoutInterval
        tmrTimeoutSend.Enabled = True
        tmrTimeoutSend.Tag = eSendMessage
    End With
End Sub

Private Sub wskEDISend_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strDataLines() As String
    Dim vntDataLine As Variant
    
    Dim strStartOfMessage As String
    Dim lngPartialMessageLength As Long
    Dim blnEndOfHeader As Boolean
    Dim blnMessageSent As Boolean
    
    blnMessageSent = False
    
    wskEDISend.GetData strData, vbString
    
    ' Append data packet to mstrDataBuffer for each trigger of DataArrival event
    mstrDataBuffer = mstrDataBuffer & strData
    
    ' TraceComm "D A T A   A R R I V A L :" & vbCrLf & strData
    mvarEDISend.TraceText = "D A T A   A R R I V A L :" & vbCrLf & strData
    
    If InStr(1, strData, "HTTP/1.1 100 Continue", vbTextCompare) > 0 Then
        ' ERE_HTTP_100_Continue
    ElseIf InStr(1, strData, "HTTP/1.1 204 The receive queue is empty, or the request timed out while waiting for the next message.", vbTextCompare) > 0 Then
        ' ERE_HTTP_204_Receive_Queue_Empty
    ElseIf InStr(1, strData, "Invalid request: null", vbTextCompare) > 0 Then
        ' ERE_HTTP_Invalid_Queue_Name
    ElseIf InStr(1, strData, "Unauthorized", vbTextCompare) > 0 Then
        ' ERE_HTTP_Invalid_User_Name
    ElseIf InStr(1, strData, "HTTP/1.1 401 User not permitted to access this resource", vbTextCompare) > 0 Then
        
    ElseIf InStr(1, strData, "UNB+UNOC:3", vbTextCompare) > 0 Then
        ' ERE_HTTP_100_IE_Message
    ElseIf InStr(1, strData, "HTTP/1.1 200 OK", vbTextCompare) > 0 Then
        ' ERE_HTTP_200_OK
        blnMessageSent = True
        'wskEDISend.Close
    Else
        ' ERE_Unknown
    End If
    
    strDataLines() = Split(strData, vbCrLf)
    
    For Each vntDataLine In strDataLines()
        If InStr(1, vntDataLine, "Content-Length:", vbTextCompare) Then
            mlngContentLength = CLng(Mid(vntDataLine, InStr(1, vntDataLine, "Content-Length:", vbTextCompare) + Len("Content-Length:") + 1))
        End If
        
        strStartOfMessage = vntDataLine
        lngPartialMessageLength = Len(strStartOfMessage)
        
        If blnEndOfHeader Then
            Exit For
        End If
        
        If Not blnEndOfHeader And lngPartialMessageLength = 0 Then
            blnEndOfHeader = True
        End If
    Next
    
    mlngContentLengthBuffer = mlngContentLengthBuffer + lngPartialMessageLength
    
    If mlngContentLength > 0 Then
        ' If content-length is greater than actual length of characters received beyond header information,
        ' wait for next data packet which will be appended to mstrDataBuffer during next trigger of DataArrival event
        If mlngContentLength > mlngContentLengthBuffer Then
            vntDataLine = ""
            
            Erase strDataLines()
            
            Exit Sub
        End If
    End If
    
    tmrTimeoutSend.Enabled = False
    
    ' TraceComm "S E N D  :  End Send Message Timeout"
    mvarEDISend.TraceText = "S E N D  :  End Send Message Timeout"
    
    'mvarEDISend.ReceiveMessage strData, blnMessageSent
    mvarEDISend.ReceiveMessage mstrDataBuffer, blnMessageSent
    
    ' Initialize for next message
    mstrDataBuffer = ""
    vntDataLine = ""
    
    mlngContentLength = 0
    mlngContentLengthBuffer = 0
    
    Erase strDataLines()
End Sub

Private Sub wskEDISend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    
    ' TraceComm "S E N D  :  Number: " & Number & ", Description: " & Description & ", Scode: " & Scode & ", Source: " & Source
    mvarEDISend.TraceText = "S E N D  :  Number: " & Number & ", Description: " & Description & ", Scode: " & Scode & ", Source: " & Source
End Sub
