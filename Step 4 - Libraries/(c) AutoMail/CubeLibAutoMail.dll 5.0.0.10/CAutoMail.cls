VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoMail"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SendBlat Lib "blat.dll" Alias "Send" (ByVal sCmd As String) As Integer

Private m_strSubject As String
Private m_strBody As String
Private m_strRecipients As String
Private m_strSMTPServer As String
Private m_strUserName As String
Private m_strPassWord As String
Private m_blnAuthentication As Boolean

'CSCLP-392
Private m_strEmail As String

'CSCLP-570
Private m_strSMTPPort As String

'CSCLP-570
Property Let Mail_SMTPPort(vSMTPPort As String)
    m_strSMTPPort = vSMTPPort
End Property

'CSCLP-392
Property Let Mail_Email(vEmail As String)
    m_strEmail = vEmail
End Property

Property Let Mail_Authentication(vAuthentication As Boolean)
    m_blnAuthentication = vAuthentication
End Property

Property Let Mail_Subject(vSubject As String)
    m_strSubject = vSubject
End Property

Property Let Mail_Body(vMessageBody As String)
    m_strBody = vMessageBody
End Property

Property Let Mail_Recipients(vMailRecipients As String)
    m_strRecipients = vMailRecipients
End Property

Property Let Mail_SMTPServer(vSMTPServer As String)
    m_strSMTPServer = vSMTPServer
End Property

Property Let Mail_UserName(vUserName As String)
    m_strUserName = vUserName
End Property

Property Let Mail_UserPassword(vMailPassword As String)
    m_strPassWord = vMailPassword
End Property

Public Function ReportAutoMail(ByVal strFileName As String, ByRef strErrDesc As String, Optional ByVal strTempPath As String = "") As Boolean
    
    Dim blnBodied As Boolean
    Dim strCmdLine As String
    Dim astrAttachments() As String
    Dim astrRecipients() As String
    Dim strRecipients As String
    Dim lngCtr As Long
    Dim lngReturnVal As Long
    Dim blnBlatError As Boolean
    
    blnBodied = False
    
    If strFileName <> "" And strTempPath = "" Then
    
        astrRecipients = Split(m_strRecipients, ";")
        strRecipients = ""
        For lngCtr = 0 To UBound(astrRecipients)
            If Trim(astrRecipients(lngCtr)) <> "" Then
                strRecipients = IIf(strRecipients = "", "", strRecipients & ",") & astrRecipients(lngCtr)
            End If
        Next
        
        'CSCLP-392
        'strCmdLine = """ -s """ & IIf(Trim(m_strSubject) = "", "<None>", Trim(m_strSubject)) & _
            """ -t " & strRecipients & " -f " & m_strUserName & _
            " -server " & m_strSMTPServer & _
            IIf(m_blnAuthentication, " -u " & m_strUserName & " -pw """ & m_strPassWord & """", "") & _
            " -body """ & IIf(Trim(m_strBody) = "", "<None>", Trim(m_strBody)) & _
            """"
        strCmdLine = """ -s """ & IIf(Trim(m_strSubject) = "", "<None>", Trim(m_strSubject)) & _
            """ -t " & strRecipients & " -f " & m_strEmail & _
            " -server " & m_strSMTPServer & _
            IIf(m_blnAuthentication, " -u " & m_strUserName & " -pw """ & m_strPassWord & """", "") & _
            " -body """ & IIf(Trim(m_strBody) = "", "<None>", Trim(m_strBody)) & _
            """"
            
        astrAttachments = Split(strFileName, "***")
        
        For lngCtr = 0 To UBound(astrAttachments)
            If Trim(astrAttachments(lngCtr)) <> "" Then
                If blnBodied = False Then
                    strCmdLine = """" & astrAttachments(lngCtr) & strCmdLine
                    blnBodied = True
                End If
                strCmdLine = strCmdLine & " -attach """ & astrAttachments(lngCtr) & """"
            End If
        Next
        
        strCmdLine = strCmdLine & " -port " & m_strSMTPPort
        
        On Error GoTo BlatErrorHandler
        lngReturnVal = SendBlat(strCmdLine)
        On Error GoTo 0
            
        For lngCtr = 0 To UBound(astrAttachments)
            If (Len(Trim(astrAttachments(lngCtr))) > 0) Then
                On Error Resume Next
                Kill astrAttachments(lngCtr)
                On Error GoTo 0
            End If
        Next
        
    Else
        
        On Error Resume Next
        Kill strTempPath & "\automail.ian"
        On Error GoTo 0
        
        lngCtr = FreeFile
           ' number.
        Open strTempPath & "\automail.ian" For Output As #lngCtr
        Write #lngCtr, m_strBody
        Close #lngCtr
        
        'CSCLP-392
        'strCmdLine = """" & strTempPath & "\automail.ian"" -s """ & IIf(Trim(m_strSubject) = "", "<None>", Trim(m_strSubject)) & _
            """ -t " & Replace(m_strRecipients, ";", ",") & " -f " & m_strUserName & _
            " -server " & m_strSMTPServer & _
            IIf(m_blnAuthentication, " -u " & m_strUserName & " -pw """ & m_strPassWord & """", "")
        strCmdLine = """" & strTempPath & "\automail.ian"" -s """ & IIf(Trim(m_strSubject) = "", "<None>", Trim(m_strSubject)) & _
            """ -t " & Replace(m_strRecipients, ";", ",") & " -f " & m_strEmail & _
            " -server " & m_strSMTPServer & _
            IIf(m_blnAuthentication, " -u " & m_strUserName & " -pw """ & m_strPassWord & """", "")
                  
        strCmdLine = strCmdLine & " -port " & m_strSMTPPort
                  
        On Error GoTo BlatErrorHandler
        lngReturnVal = SendBlat(strCmdLine)
        On Error GoTo 0
        
        Kill strTempPath & "\automail.ian"
        
    End If
    
    If blnBlatError = False Then
        Select Case lngReturnVal
        
            Case 0
                ReportAutoMail = True
            Case 1
                ReportAutoMail = False
                strErrDesc = "1: Mail server error accepting message data"
            Case 2
                ReportAutoMail = False
                strErrDesc = "2: The server actively denied our connection"
            Case 3
                ReportAutoMail = False
                strErrDesc = "3: Error reading the file (message text) or attached file"
            Case 4
                ReportAutoMail = False
                strErrDesc = "4: File (message text) not of type FILE_TYPE_DISK"
            Case 5
                ReportAutoMail = False
                strErrDesc = "5: Error Reading File (message text)"
            Case 12
                ReportAutoMail = False
                strErrDesc = "12: SMTP Server or Username not specified"
            Case 13
                ReportAutoMail = False
                strErrDesc = "13: Error opening temporary file in temp directory"
            Case 4001
                ReportAutoMail = False
                strErrDesc = "4001: Malloc failed (possibly out of memory)"
            Case 4002
                ReportAutoMail = False
                strErrDesc = "4002: Error sending data"
            Case 4003
                ReportAutoMail = False
                strErrDesc = "4003: Error initializing gensock.dll"
            Case 4004
                ReportAutoMail = False
                strErrDesc = "4004: Version not supported"
            Case 4005
                ReportAutoMail = False
                strErrDesc = "4005: The winsock version specified by gensock is not supported by this winsock.dll"
            Case 4006
                ReportAutoMail = False
                strErrDesc = "4006: Network not ready"
            Case 4007
                ReportAutoMail = False
                strErrDesc = "4007: Can't resolve (mailserver) hostname"
            Case 4008
                ReportAutoMail = False
                strErrDesc = "4008: Can't create a socket (too many simultaneous links?)"
            Case 4009
                ReportAutoMail = False
                strErrDesc = "4009: Error reading socket"
            Case 4010
                ReportAutoMail = False
                strErrDesc = "4010: Not a socket"
            Case 4011
                ReportAutoMail = False
                strErrDesc = "4011: Busy"
            Case 4012
                ReportAutoMail = False
                strErrDesc = "4012: Error closing socket"
            Case 4013
                ReportAutoMail = False
                strErrDesc = "4013: Wait a bit (possible timeout)"
            Case 4014
                ReportAutoMail = False
                strErrDesc = "4014: Can't resolve service"
            Case 4015
                ReportAutoMail = False
                strErrDesc = "4015: Can't connect to mailserver (timed out if winsock.dll error 10060)"
            Case 4016
                ReportAutoMail = False
                strErrDesc = "4016: Connection to mailserver was dropped"
            Case 4017
                ReportAutoMail = False
                strErrDesc = "4017: Mail server refused connection"
            Case Else
                ReportAutoMail = False
                strErrDesc = lngReturnVal & ": Unspecefied reason"
        End Select
    End If
    
'    Shell (App.Path & "\blat.exe """ & strFileName & """ -s " & m_strMail_Subject & _
        " -t " & Replace(m_strMail_Recipients, ";", ",") & " -f " & strUserName & _
        " -server " & strSMTPServer & " -u " & strUserName & _
        " -pw " & strPassword & " -body """ & m_strMail_Body & _
        """ -attach """ & strFileName & ""), vbHide

    Exit Function
    
BlatErrorHandler:

    strErrDesc = "(BLAT INTERNAL ERROR) " & Err.Number & " " & Err.Source & ": " & Err.Description
    ReportAutoMail = False
    blnBlatError = True
    Err.Clear
    Resume Next
    
End Function


