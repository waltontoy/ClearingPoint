VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   3780
   ClientTop       =   2475
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1875
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1260
      Width           =   2115
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      Left            =   1875
      TabIndex        =   0
      Top             =   840
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4185
      TabIndex        =   3
      Top             =   1238
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4185
      TabIndex        =   2
      Top             =   825
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   0
      X2              =   5505
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please type in your user name and password to log on."
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Tag             =   "196"
      Top             =   420
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User name:"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Tag             =   "197"
      Top             =   885
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Tag             =   "198"
      Top             =   1305
      Width           =   735
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "  Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   2865
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   -15
      Width           =   5550
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1590
      Left            =   0
      Top             =   300
      Width           =   5550
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsMain As CMainControls
Private m_conLogin As ADODB.Connection
Private m_objApplication As Object

Private m_strUser As String
Private m_lngLoginCtr As Long
Private m_blnSecurityOn As Boolean

Private m_strApplicationName As String
Private m_strMDBPath As String

Private m_enuResult As QueryResultConstants

Private Sub cmdCancel_Click()
    m_enuResult = QueryResultNoRecord
    Unload Me
End Sub

Public Function Login(ByRef MainProps As CMainControls, _
                        ByVal Application As Object, _
                        ByRef OwnerForm As Object, _
                        ByRef ADOConnection As ADODB.Connection, _
                        ByVal SecurityOn As Boolean, _
                        ByRef UserName As String) _
                        As QueryResultConstants
    
    Dim clsRegistry As CRegistry
    Dim strPassword As String
    
    Dim lngFileNum As Long
    Dim blnUsernamePassword_AutoLog As Boolean
        
    
    On Error GoTo Error_Handler
10200   Set m_clsMain = MainProps
10300   Set m_conLogin = ADOConnection
10400   Set m_objApplication = Application
        
10500   m_blnSecurityOn = SecurityOn
10600   m_strApplicationName = m_objApplication.ProductName
10800   m_strUser = UserName_ToUse(UserName)
        
        
        '--->Retrieve MdbPath from registry
11000   Set clsRegistry = New CRegistry
11100   If (clsRegistry.GetRegistry(cpiLocalMachine, m_strApplicationName, "Settings", "MDBPath") = True) Then
11200       m_strMDBPath = clsRegistry.RegistryValue
        End If
11400   Set clsRegistry = Nothing
        
        
        '--->Retrieve password saved in registry, if there is any
11600   Set clsRegistry = New CRegistry
11700   If (clsRegistry.GetRegistry(cpiCurrentUser, m_strApplicationName, "Settings", "Password") = True) Then
11800       strPassword = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
        End If
        
11900   If (Len(Trim(strPassword)) > 0) Then
            '--->Delete saved password in registry
12100       clsRegistry.SaveRegistry cpiCurrentUser, m_strApplicationName, "Settings", "Password", ""
        End If
12200   Set clsRegistry = Nothing
        
        
        '--->Check if user exists in the database and password is correct
12400   blnUsernamePassword_AutoLog = IsPassword_Exist(m_strUser, strPassword)
12500   lblWelcome.Caption = "Welcome to " & Application.FileDescription
        
        
12600   If (m_blnSecurityOn = False) Or (blnUsernamePassword_AutoLog = True) Then
            
            '--->Security is OFF or Username exists and Password is correct - No need to display login form
12800       If (MainProps.UserID > 0) Then
12900           lngFileNum = FreeFile
                
                On Error Resume Next
13100           Open m_strMDBPath & "\" & IIf(Len(CStr(MainProps.UserID)) = 1, "0" & MainProps.UserID, MainProps.UserID) & m_strApplicationName & ".flk" _
                    For Output Lock Read Write As #lngFileNum
            End If
            
            If (Err.Number <> 0) Then
                Err.Clear
                On Error GoTo Error_Handler
                
13300           MsgBox "The account '" & m_strUser & "' is in use. Please use another user account.", vbInformation, m_strApplicationName & " Login"
13400           m_strUser = ""
                
                Me.Show vbModal
                
13500       ElseIf (GetUser(m_strUser, Trim(strPassword)) = False) Then
                Me.Show vbModal
                
            End If
            
        Else
            Me.Show vbModal
            
        End If
        
        
        '--->Return values
13900   UserName = m_strUser
14000   Set MainProps = m_clsMain
14100   Login = m_enuResult
14200   Set ADOConnection = m_conLogin
14300   Set m_objApplication = Nothing
        
        Exit Function
        
Error_Handler:
    Err.Clear
    
End Function

Private Function IsPassword_Exist(ByVal LastUsedUserName As String, ByVal LastUsedPassword As String) As Boolean
    Dim rstUsers As ADODB.Recordset
    'Dim clsRecordset As CRecordset
    
    
    On Error GoTo Error_Handler
        '--->Open recordset for user
'10200   Set clsRecordset = New CRecordset
        
10300   ADORecordsetOpen "SELECT User_Password FROM Users WHERE User_Password='" & LastUsedPassword & "'" & " AND User_Name='" & LastUsedUserName & "'", m_conLogin, rstUsers, adOpenKeyset, adLockOptimistic

'10300   clsRecordset.cpiOpen "SELECT User_Password FROM Users WHERE User_Password='" & LastUsedPassword & "'" & " AND User_Name='" & LastUsedUserName & "'", m_conLogin, rstUsers, adOpenKeyset, adLockOptimistic, , True
        
        
        '--->Check if password is the same
10500   If (rstUsers.RecordCount > 0) Then
10700       IsPassword_Exist = True
        Else
10900       IsPassword_Exist = False
            
        End If
        
11000   ADORecordsetClose rstUsers

'11000   clsRecordset.cpiClose rstUsers
'11100   Set clsRecordset = Nothing
        
        Exit Function
        
Error_Handler:
    Err.Clear
    
End Function

Private Function UserName_ToUse(ByVal LastUsedUserName As String) As String
    Dim rstUsers As ADODB.Recordset
    'Dim clsRecordset As CRecordset
    
    
    On Error GoTo Error_Handler
        '--->Open recordset for user
        
        ADORecordsetOpen "SELECT User_Name FROM Users WHERE User_Name='" & LastUsedUserName & "'", m_conLogin, rstUsers, adOpenKeyset, adLockOptimistic
'10200   Set clsRecordset = New CRecordset
'10300   clsRecordset.cpiOpen "SELECT User_Name FROM Users WHERE User_Name='" & LastUsedUserName & "'", m_conLogin, rstUsers, adOpenKeyset, adLockOptimistic, , True
        
        
        '--->Check if user exists in the database
10500   If (rstUsers.RecordCount > 0) Then
            
10600       UserName_ToUse = LastUsedUserName
        Else
10800       UserName_ToUse = ""
        End If
        
        
11000   ADORecordsetClose rstUsers
'11100   Set clsRecordset = Nothing
        
        Exit Function
        
        
Error_Handler:
    Err.Clear
    
End Function
Private Sub cmdOK_Click()
    
    Call GetUser(txtUser.Text, txtPassword.Text)
    
End Sub

Private Function GetUser(ByRef UserName As String, ByVal Password As String) As Boolean
    Dim clsRegistryLastUser As CRegistry
    
    Dim rstUser As ADODB.Recordset
    Dim strCommandText As String
    Dim lngFileNum As Long
    Dim lngErrNumber As Long
    
    
    On Error GoTo Error_Handler
        '--->Increment login counter
10100   m_lngLoginCtr = m_lngLoginCtr + 1
        
        '--->Open recordset for user
10400   strCommandText = vbNullString
10500   strCommandText = strCommandText & "SELECT "
10600   strCommandText = strCommandText & "Users.User_ID AS User_ID, "
10601   strCommandText = strCommandText & "Users.User_Name AS User_Name, "
10602   strCommandText = strCommandText & "Users.User_Password AS User_Password "
10700   strCommandText = strCommandText & "FROM "
10800   strCommandText = strCommandText & "Users "
10900   strCommandText = strCommandText & "WHERE "
11000   strCommandText = strCommandText & "Ucase(User_Name) = '" & ProcessQuotes(Trim(UCase(UserName))) & "' "
        
11100   ADORecordsetOpen strCommandText, m_conLogin, rstUser, adOpenKeyset, adLockOptimistic
'11100   Set rstUser = m_conLogin.Execute(strCommandText)
        
        
11200   If (rstUser.EOF And rstUser.BOF) Then
11400       MsgBox "The user name you have specified is invalid.", vbInformation, m_objApplication.FileDescription
            
11500       GetUser = False
'<<< dandan 9/4/06
'Note: This line produces RTE 5: Invalid call or procedure. Purpose of this isto assign the focus of the cursor to
'txtUser textbox, however, even if not assigned, the focus of the cursor will be in txtUser textbox since it is the first
'index control of the form.
'11600       txtUser.SetFocus
11700       SendKeysEx "{HOME}+{END}"
            
        Else
            
            rstUser.MoveFirst
            
11800       If ((rstUser!User_Password <> Password) And (m_blnSecurityOn = True)) Then
12000           MsgBox "The password you have typed is invalid.", vbInformation, m_objApplication.FileDescription
                
12100           GetUser = False
12200           txtPassword.SetFocus
12300           SendKeysEx "{HOME}+{END}"
                
            Else
12600           lngFileNum = FreeFile
                On Error Resume Next
12700           Open m_strMDBPath & "\" & IIf(Len(CStr(rstUser!User_ID)) = 1, "0" & rstUser!User_ID, rstUser!User_ID) & m_strApplicationName & ".flk" _
                    For Output Lock Read Write As #lngFileNum
                
12800           lngErrNumber = Err.Number
                On Error GoTo 0
                
                
                If (lngErrNumber <> 0) Then
                    Err.Clear
13000               MsgBox "The account '" & rstUser!User_Name & "' is in use. Please use another user account.", vbInformation, m_strApplicationName & " Login"
                    
13100               GetUser = False
                    If frmLogin.Visible = True Then
13200                   txtUser.SetFocus
13300                   SendKeysEx "{HOME}+{END}"
                        GoTo Exit_Function
                    End If
                    
                Else
13500               m_strUser = UserName
13600               m_enuResult = QueryResultSuccessful
13700               GetUser = True
                    
13800               Set clsRegistryLastUser = New CRegistry
                    
13900               m_clsMain.UserID = rstUser!User_ID
14000               m_clsMain.User_Password = rstUser!User_Password
14100               m_clsMain.User_Name = rstUser!User_Name
                    
                    '--->Save UserName into the registry
14300               clsRegistryLastUser.SaveRegistry cpiCurrentUser, m_strApplicationName, "Settings", "UserName", txtUser.Text
14400               clsRegistryLastUser.SaveRegistry cpiCurrentUser, m_strApplicationName, "Temp", "Temp", "Temp"
                    
14500               Set clsRegistryLastUser = Nothing
14600               Unload Me
                    
                    GoTo Exit_Function
                    
                End If
                
            End If
            
        End If
        
        
        '--->Check number of tries logging in
14800   If (m_lngLoginCtr = 3) Then
15000       MsgBox "Access Denied. Please contact your Password Administrator.", vbInformation, m_objApplication.FileDescription
            
15100       m_enuResult = QueryResultNoRecord
15200       GetUser = False
15300       Unload Me
            
        End If
        
        
Exit_Function:
        ' hobbes 10/18/2005
        Call ADORecordsetClose(rstUser)
        
        Exit Function
        
Error_Handler:
    Err.Clear
    
End Function

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    Me.Refresh
    If (Len(Trim(txtUser.Text)) <> 0) Then
        txtPassword.SetFocus
    Else
        txtUser.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    txtUser.Text = m_strUser
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim clsRegistry As CRegistry
    
        If (m_enuResult = QueryResultSuccessful) Then
            Set clsRegistry = New CRegistry
10200           clsRegistry.SaveRegistry cpiCurrentUser, m_objApplication.ProductName, "Settings", "Security", Encrypt(IIf(m_blnSecurityOn, "On", "Off"), KEY_ENCRYPT)
10300           clsRegistry.SaveRegistry cpiCurrentUser, m_objApplication.ProductName, "Settings", "UserName", Encrypt(m_strUser, KEY_ENCRYPT)
10400           clsRegistry.SaveRegistry cpiLocalMachine, m_objApplication.ProductName, "Settings", "AppPath", Encrypt(m_objApplication.Path, KEY_ENCRYPT)
            Set clsRegistry = Nothing
        End If
End Sub

Private Sub txtPassword_GotFocus()
    
    SendKeysEx "{Home}+{End}"
    
End Sub

Private Sub txtUser_GotFocus()
    
    SendKeysEx "{Home}+{End}"
    
End Sub


