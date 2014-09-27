VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3480
   ClientLeft      =   3375
   ClientTop       =   4500
   ClientWidth     =   4785
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   3000
      Width           =   1355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3000
      Width           =   1355
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Security"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraOpLock"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Locations"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdModify"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvwFileLocations"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Refresh "
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmRefresh"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&E-mail"
      TabPicture(3)   =   "frmOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdClear"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtDefaultMessage"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame fraOpLock 
         Caption         =   "Opportunistic Locking Checking Upon Startup"
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   4335
         Begin VB.OptionButton optOpLock 
            Caption         =   "Disabled"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optOpLock 
            Caption         =   "Enabled"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwFileLocations 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "  "
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.TextBox txtDefaultMessage 
         Height          =   1575
         Left            =   -74880
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   -71880
         TabIndex        =   11
         Top             =   2280
         Width           =   1355
      End
      Begin VB.Frame frmRefresh 
         Caption         =   "Rate"
         Height          =   975
         Left            =   -74880
         TabIndex        =   18
         Top             =   1380
         Width           =   4335
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtRefreshRate"
            BuddyDispid     =   196614
            OrigLeft        =   1320
            OrigTop         =   360
            OrigRight       =   1515
            OrigBottom      =   645
            Max             =   60
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtRefreshRate 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "seconds"
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   375
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Every"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   375
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mode"
         Height          =   885
         Left            =   -74880
         TabIndex        =   17
         Top             =   420
         Width           =   4335
         Begin VB.OptionButton optRefreshMode 
            Caption         =   "Manual"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optRefreshMode 
            Caption         =   "Automatic"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify..."
         Height          =   375
         Left            =   -71880
         TabIndex        =   5
         Top             =   2325
         Width           =   1355
      End
      Begin VB.Frame Frame5 
         Caption         =   "Password Security"
         Height          =   645
         Left            =   120
         TabIndex        =   15
         Tag             =   "174"
         Top             =   420
         Width           =   4335
         Begin VB.OptionButton optSecurityState 
            Caption         =   "Enabled"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Tag             =   "165"
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optSecurityState 
            Caption         =   "Disabled"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   2
            Tag             =   "166"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Tag             =   "175"
         Top             =   1860
         Width           =   4335
         Begin VB.PictureBox picFixedUserName 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   3975
            TabIndex        =   22
            Top             =   240
            Width           =   3975
            Begin VB.TextBox txtFixedUserName 
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   3015
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Fixed User Name"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Default Message :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CONST_MAXIMUM_PATH_LENGTH = 100
            
Private strListItems(CONST_MAXIMUM_PATH_LENGTH) As String
Private mvarApplication As Object
Private mvarApplicationName As String
Private strRefreshValueSnapShot As String
Private mvarOWnerForm As Object

Private clsMachinePaths As CMachinePaths
Private clsRegistry As CRegistry
Private clsFilePath As CBrowse
Private clsValidateNumber As CNumbers
Private mvarEndApplication As String

Private strUserPassword As String

Private m_clsMainSettings As CMainControls



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtDefaultMessage.Text = ""
End Sub

Private Sub cmdModify_Click()
    Dim clsBrowse As CBrowse
                    
    Dim strDefaultPath As String
    Dim strPathValue As String
    Dim strRebuildTag As String
    Dim lngRebuildTagCtr As Long
    Dim arrTagDetails
        
    If Not lvwFileLocations.SelectedItem Is Nothing Then
        Set clsBrowse = New CBrowse
                    
        arrTagDetails = Split(CStr(lvwFileLocations.SelectedItem.Tag), "|")
                                                    
        If Trim(CStr(arrTagDetails(UBound(arrTagDetails)))) = "" Then
            clsRegistry.GetRegistry cpiLocalMachine, mvarApplicationName, clsMachinePaths.Item(lvwFileLocations.SelectedItem.Key).RegistryKey, clsMachinePaths.Item(lvwFileLocations.SelectedItem.Key).RegistrySetting, , False
                        
            strPathValue = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
        
            If clsMachinePaths.Item(lvwFileLocations.SelectedItem.Key).PathType <> "File" Then
                strDefaultPath = strPathValue
            Else
                strDefaultPath = Left(strPathValue, InStrRev(strPathValue, "\") - 1)
            End If
        Else
            strDefaultPath = CStr(arrTagDetails(UBound(arrTagDetails)))
        End If
        
        With clsMachinePaths.Item(lvwFileLocations.SelectedItem.Key)
            Select Case .PathType
                Case "File"
                    If Not clsBrowse.BrowseFile(Me, mvarApplication, strDefaultPath, CStr(arrTagDetails(4))) Then
                        Set clsBrowse = Nothing
                        Exit Sub
                    End If
                Case "Folder"
                    If Not clsBrowse.BrowseFolder(Me, "Browse for folder to map with " & .DisplayName, strDefaultPath) Then
                        Set clsBrowse = Nothing
                        Exit Sub
                    End If
                Case "Database"
                    If Not clsBrowse.BrowseFile(Me, mvarApplication, strDefaultPath, CStr(arrTagDetails(3))) Then
                        Set clsBrowse = Nothing
                        Exit Sub
                    End If
            End Select
            
            strRebuildTag = ""
            For lngRebuildTagCtr = 0 To UBound(arrTagDetails)
                If lngRebuildTagCtr = UBound(arrTagDetails) Then
                    Select Case .PathType
                        Case "File"
                            strRebuildTag = strRebuildTag & "|" & clsBrowse.Path & "\" & clsBrowse.FileName
                        Case "Folder", "Database"
                            strRebuildTag = strRebuildTag & "|" & clsBrowse.Path
                    End Select
                Else
                    strRebuildTag = strRebuildTag & "|" & CStr(arrTagDetails(lngRebuildTagCtr))
                End If
            Next lngRebuildTagCtr
            
            strRebuildTag = Mid(strRebuildTag, 2)
                                
            lvwFileLocations.SelectedItem.Tag = strRebuildTag
        End With
            
        Set clsBrowse = Nothing
    End If
End Sub

Private Sub cmdOK_Click()

    Call SaveOptionSettings
    
End Sub

Private Sub Form_Load()
    Dim varListItem As Variant
    
    Dim strSecurityStatus As String
    Dim strRefreshValue As String
    Dim strPathDisplay As String
    Dim strPathValue As String
    Dim lngListItemCtr As Long
    
    Dim strOpLockSetting As String
    
    Set clsRegistry = New CRegistry
    Set clsValidateNumber = New CNumbers
            
    Screen.MousePointer = vbHourglass
    
    ' Set Security Status
    clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Security", cpiStandard, Encrypt("On", KEY_ENCRYPT)
    strSecurityStatus = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
    
    optSecurityState(0).Value = IIf(strSecurityStatus = "On", True, False)
    optSecurityState(1).Value = IIf(strSecurityStatus = "On", False, True)
    
    ' Set Opportunistic Locking Checking
    clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "OpLockSetting", cpiStandard, Encrypt("On", KEY_ENCRYPT)
    strOpLockSetting = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
    
    optOpLock(0).Value = IIf(strOpLockSetting = "On", True, False)
    optOpLock(1).Value = IIf(strOpLockSetting = "On", False, True)
    
    ' Get Currently Logged User
    clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", Encrypt("UserName", KEY_ENCRYPT)
    txtFixedUserName.Text = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
    
    ' Set Refresh Settings
    clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshValue", cpiStandard, Encrypt("10", KEY_ENCRYPT)
    strRefreshValue = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
    txtRefreshRate.Text = strRefreshValue
        
    clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshManual", cpiStandard, Encrypt("False", KEY_ENCRYPT)
    optRefreshMode(0).Value = IIf(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) = "False", 1, 0)
    optRefreshMode(1).Value = IIf(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) = "False", 0, 1)
    
    ' Set Mail Message Default
    clsRegistry.GetRegistry cpiCurrentUser, _
                            mvarApplicationName, _
                            "Settings", _
                            "MailMsg", _
                            cpiStandard, _
                            Encrypt("<recipient>" _
                            & vbCrLf & _
                            "<address> " & vbCrLf & _
                            "<zip code> <city>" & vbCrLf & _
                            "<country>" & vbCrLf & _
                            " " & vbCrLf & _
                            "Dear Sir / Madam:" & vbCrLf & _
                            " " & vbCrLf & _
                            "Please see the attached <report>." & vbCrLf & _
                            " " & vbCrLf & _
                            "We appreciate your business." & vbCrLf & _
                            "Thank you very much." & _
                            " " & _
                            "<user>", KEY_ENCRYPT)
                                                                            
    txtDefaultMessage.Text = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
                        
    For lngListItemCtr = 1 To clsMachinePaths.Count
        If Not clsRegistry.GetRegistry(cpiLocalMachine, mvarApplicationName, clsMachinePaths.Item(lngListItemCtr).RegistryKey, clsMachinePaths.Item(lngListItemCtr).RegistrySetting, , False) Then
            ' Disable All Settings
        Else
            strPathValue = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
            
            strPathDisplay = clsMachinePaths.Item(lngListItemCtr).PathType
            
            Set varListItem = lvwFileLocations.ListItems.Add(, clsMachinePaths.Item(lngListItemCtr).Key, strPathDisplay)
            varListItem.SubItems(1) = clsMachinePaths.Item(lngListItemCtr).DisplayName
                                                                        
            varListItem.Tag = clsMachinePaths.Item(lngListItemCtr).Key & "|" & " "
        End If
    Next lngListItemCtr
                          
    Set clsValidateNumber = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsRegistry = Nothing
    Set clsValidateNumber = Nothing
    Set clsMachinePaths = Nothing
End Sub

Private Sub lvwFileLocations_GotFocus()
    SSTab1.Tab = 1
End Sub

Private Sub lvwFileLocations_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim varHitListItem As Variant
    
    Set varHitListItem = lvwFileLocations.HitTest(x, y)
    
    If Not varHitListItem Is Nothing Then
        lvwFileLocations.ToolTipText = CStr(varHitListItem.SubItems(1))
    End If
End Sub

Private Sub optRefreshMode_Click(Index As Integer)
        
    ToggleRefreshControls
        
End Sub

Private Sub optRefreshMode_GotFocus(Index As Integer)
    SSTab1.Tab = 2
End Sub

Private Sub optSecurityState_GotFocus(Index As Integer)
    SSTab1.Tab = 0
End Sub

Private Sub txtDefaultMessage_GotFocus()
    SSTab1.Tab = 3
End Sub

Private Sub txtRefreshRate_Change()
    Set clsValidateNumber = New CNumbers
    
    If clsValidateNumber.IsValidValue(txtRefreshRate.Text, "5", True, "60", True, False, False, False, False) = True Then
        strRefreshValueSnapShot = txtRefreshRate.Text
    End If
    
    Set clsValidateNumber = Nothing
End Sub

Private Sub txtRefreshRate_GotFocus()
    strRefreshValueSnapShot = txtRefreshRate.Text
End Sub

Private Sub txtRefreshRate_Validate(Cancel As Boolean)
    Set clsValidateNumber = New CNumbers
    
    If clsValidateNumber.IsValidValue(txtRefreshRate.Text, "5", True, "60", True, False, False, False, False) = False Then
        txtRefreshRate.Text = strRefreshValueSnapShot
    End If
    
    Set clsValidateNumber = Nothing
End Sub

Public Function ShowOptions(ByRef OwnerForm As Object, ByVal Application As Object, ByRef MachinePaths As CMachinePaths, ByRef MainSettings As CMainControls) As String
    mvarEndApplication = ""
    
    Set m_clsMainSettings = MainSettings
    
    strUserPassword = m_clsMainSettings.User_Password
    
    Set mvarOWnerForm = OwnerForm
    Set clsMachinePaths = MachinePaths
    
    Set mvarApplication = Application
    mvarApplicationName = Application.ProductName
    
    
    Set Me.Icon = OwnerForm.Icon
    
    Me.Show vbModal
    
    Set MainSettings = m_clsMainSettings
    
    ShowOptions = mvarEndApplication
End Function

Public Sub ToggleRefreshControls()
    Dim blnAutoRefresh As Boolean
    
    blnAutoRefresh = IIf(optRefreshMode(0).Value = True, True, False)
    
    frmRefresh.Enabled = blnAutoRefresh
    Label4.Enabled = blnAutoRefresh
    Label5.Enabled = blnAutoRefresh
    txtRefreshRate.Enabled = blnAutoRefresh
    UpDown1.Enabled = blnAutoRefresh
End Sub

Public Sub SaveOptionSettings()
    
    Dim lngListCtr As Long
    Dim arrList
    Dim blnDatabaseChanged As Boolean
    Dim blnCancel As Boolean
    Dim strOldDatabasePath As String
    Dim varReturnCode As Variant
    Dim strUsernameTemp As String
    Dim strPasswordTemp As String
    Dim strSecurityTemp As String
    Dim strRefreshManual As String
    Dim strRefreshValue As String
    Dim strNewDBPath As String
    Dim strOpLockSetting As String
            
    
    ' Save Security Settings
    clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Security", Encrypt(IIf(optSecurityState(0).Value = True, "On", "Off"), KEY_ENCRYPT)
                    
    ' Save Opportunistic Locking Checking - 'Edwin Dec3
    clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "OpLockSetting", Encrypt(IIf(optOpLock(0).Value = True, "On", "Off"), KEY_ENCRYPT)
                        
    ' Set Refresh Settings
    clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshValue", Encrypt(Trim(txtRefreshRate.Text), KEY_ENCRYPT)
    clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshManual", Encrypt(IIf(optRefreshMode(0).Value = True, "False", "True"), KEY_ENCRYPT)
    
    
    ' Set Mail Message Default
    clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "MailMsg", Encrypt(IIf(Trim(txtDefaultMessage.Text) <> "", txtDefaultMessage.Text, "<recipient>" & vbCrLf & _
                                                                                                                                                         "<address> " & vbCrLf & _
                                                                                                                                                         "<user>"), KEY_ENCRYPT)

    ' Set Main Work Area Timer Settings If Needed
    mvarOWnerForm.tmrRefresh.Enabled = False
    If optRefreshMode(0).Value = True Then
        mvarOWnerForm.tmrRefresh.Interval = CLng(txtRefreshRate.Text) * 1000
        mvarOWnerForm.tmrRefresh.Enabled = True
    End If
    
    blnDatabaseChanged = False
    
    For lngListCtr = 1 To lvwFileLocations.ListItems.Count
        arrList = Split(CStr(lvwFileLocations.ListItems(lngListCtr).Tag), "|")
        strOldDatabasePath = ""
        
        If Len(Trim(CStr(arrList(UBound(arrList))))) <> 0 Then
            If UCase(CStr(arrList(0))) = "D" Then
                If clsRegistry.GetRegistry(cpiLocalMachine, mvarApplicationName, clsMachinePaths.Item(lvwFileLocations.ListItems(lngListCtr).Key).RegistryKey, clsMachinePaths.Item(lvwFileLocations.ListItems(lngListCtr).Key).RegistrySetting) Then
                    strOldDatabasePath = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
                End If
            End If
            
            If UCase(CStr(arrList(0))) = "D" And strOldDatabasePath <> CStr(arrList(UBound(arrList))) Then
                strNewDBPath = CStr(arrList(UBound(arrList)))
                
                Call clsMachinePaths.RestartEvents(strNewDBPath, blnCancel)
                
                If blnCancel = True Then Exit Sub
                
                blnDatabaseChanged = True
            End If
            
            clsRegistry.SaveRegistry cpiLocalMachine, mvarApplicationName, clsMachinePaths.Item(lvwFileLocations.ListItems(lngListCtr).Key).RegistryKey, clsMachinePaths.Item(lvwFileLocations.ListItems(lngListCtr).Key).RegistrySetting, Encrypt(CStr(arrList(UBound(arrList))), KEY_ENCRYPT)
            
        End If
    Next lngListCtr
        
    ' Check if the user has to restart
    If blnDatabaseChanged Then
        Call clsMachinePaths.TriggerRegistryChanged(strNewDBPath)
        
        '--->Retrieve values to save to registry
        clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Username"
        strUsernameTemp = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
        
        strPasswordTemp = strUserPassword
        
        clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Security"
        strSecurityTemp = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
        
        clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "OpLockSetting" 'Edwin Dec3
        strOpLockSetting = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) 'Edwin Dec3
                
        clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshManual"
        strRefreshManual = IIf(Len(Trim(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT))) = 0, "True", Trim(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)))
        
        clsRegistry.GetRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshValue"
        strRefreshValue = IIf(Len(Trim(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT))) = 0, "10", Trim(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)))
        
        '--->Delete registry setting "Settings"
        clsRegistry.DeleteRegistryKey cpiCurrentUser, mvarApplication.ProductName, "Settings"
        
        '--->Save registry keys
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Username", Encrypt(strUsernameTemp, KEY_ENCRYPT)
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Password", Encrypt(strPasswordTemp, KEY_ENCRYPT)
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "Security", Encrypt(strSecurityTemp, KEY_ENCRYPT)
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "OpLockSetting", Encrypt(strOpLockSetting, KEY_ENCRYPT) 'Edwin Dec3
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshManual", Encrypt(strRefreshManual, KEY_ENCRYPT)
        clsRegistry.SaveRegistry cpiCurrentUser, mvarApplicationName, "Settings", "RefreshValue", Encrypt(strRefreshValue, KEY_ENCRYPT)
        
        
        If MsgBox("Would you like to end so you can restart now?", vbInformation + vbYesNo, mvarApplicationName) = vbYes Then
            mvarEndApplication = "End"
            Unload Me
            UnloadControls mvarOWnerForm
            Unload mvarOWnerForm
            varReturnCode = ShellExecute(0, "Open", mvarApplication.EXEName, "?????", mvarApplication.Path, 1)
        Else
            Unload Me
        End If
        
    Else
        Unload Me
    End If
    
End Sub

Sub UnloadControls(ByVal frmBeingUnloaded As Form, Optional ByVal blnTag As Boolean)
    Dim ctlToUnload As Control
    
    On Error Resume Next
    
    For Each ctlToUnload In frmBeingUnloaded.Controls
        Set ctlToUnload = Nothing
    Next
End Sub
