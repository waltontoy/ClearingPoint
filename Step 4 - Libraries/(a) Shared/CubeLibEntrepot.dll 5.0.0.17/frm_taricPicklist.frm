VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_taricpicklist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TARIC - Picklist"
   ClientHeight    =   7425
   ClientLeft      =   2415
   ClientTop       =   1170
   ClientWidth     =   9390
   Icon            =   "frm_taricPicklist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "859"
   Begin MSComctlLib.ListView lvwPicklistFilter 
      Height          =   1005
      Left            =   7680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwPicklist 
      Height          =   5055
      Left            =   105
      TabIndex        =   4
      Top             =   1800
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "0"
         Text            =   "TARIC Code"
         Object.Width           =   1917
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "0"
         Text            =   "Keyword"
         Object.Width           =   3387
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "0"
         Text            =   "Description"
         Object.Width           =   7752
      EndProperty
   End
   Begin VB.CommandButton cmdSelectCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   14
      Tag             =   "179"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelectCancel 
      Caption         =   "&Select"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   13
      Tag             =   "426"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPicklist 
      Caption         =   "&Modify..."
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   8
      Tag             =   "834"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPicklist 
      Caption         =   "&Add..."
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   7
      Tag             =   "811"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdPicklist 
      Caption         =   "&Kluwer..."
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   6
      Tag             =   "830"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdPicklist 
      Caption         =   "&CN Codes..."
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   5
      Tag             =   "819"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPicklist 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   3
      Tag             =   "827"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   3120
      MaxLength       =   78
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox txtKeyword 
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtTaricCode 
      Height          =   375
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraShowOnly 
      Caption         =   "Show only TARIC codes associated with"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Tag             =   "850"
      Top             =   120
      Width           =   9135
      Begin VB.CheckBox chkShowOnly 
         Caption         =   "Country <country code> - <country name>"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Value           =   2  'Grayed
         Width           =   5160
      End
      Begin VB.CheckBox chkShowOnly 
         Caption         =   "Client <VAT number> - <company name>"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Value           =   2  'Grayed
         Width           =   5160
      End
      Begin VB.CheckBox chkShowOnly 
         Caption         =   "<document type>"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   2  'Grayed
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_taricpicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public m_conSADBEL As ADODB.Connection    'Public datSADBEL As DAO.Database
Public m_conTemplate As ADODB.Connection  'Public datTemplate As DAO.Database
Public m_conTaric As ADODB.Connection     'Private datTaric As DAO.Database

Private intArrayCtr As Integer
Private blnPassBoxValues As Boolean
Public strDocType As String
Public strBoxValue As String

Private strDocName As String
Private strFolderID As String

Public strLangOfDesc As String
Private strVATNumber As String
Private strClientName As String
Private strRecipientName As String
Private strVATNumOrName As String
Public strCtryCode As String
Public strCtryName As String
Public blnOKWasPressed As Boolean

Private blnWasInvokedFromCode As Boolean
Private blnFindWasClicked As Boolean
Private mblnClientLowerLeftChanged As Boolean
Private mintClientLowerLeft As Integer
Private msngColumnHeaderWidths() As Single
Private msngMousePointerX As Single
Private msngMousePointerY As Single

Private Enum CommandButtonIndexConstants
    sbpFind = 0
    sbpCNCodes
    sbpKluwer
    sbpAdd
    sbpModify
End Enum

'-----> For third party database use
Private intThirdPartyDatabase As Integer

Private Const SW_SHOWNORMAL = 1    ' Restores window if minimized or maximized
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Private strClient As String     ' "Client"
Private strCountry As String    ' "Country"

Private Sub chkShowOnly_Click(Index As Integer)
    Dim itmListItem As MSComctlLib.ListItem
    
    Dim strSQL As String
    Dim strCommImpExpSQL As String
    Dim strUnionSQL As String
    Dim strTableType As String

    Dim rstTaricCodes As ADODB.Recordset
    
    If Not blnWasInvokedFromCode Then
        Screen.MousePointer = vbHourglass
        
        strSQL = ""
        strUnionSQL = ""
        strTableType = IIf(strDocType = "Import", "Import", "Export")
        
        lvwPicklist.ListItems.Clear
        
        If chkShowOnly(0).Value = vbChecked Then
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "DISTINCT "
            strSQL = strSQL & "COMMON.[Taric Code], "
            strSQL = strSQL & "[Key " & strLangOfDesc & "], "
            strSQL = strSQL & "[Desc " & strLangOfDesc & "] "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "COMMON, "
            strSQL = strSQL & strTableType & " "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "COMMON.[Taric Code] = " & strTableType & ".[Taric Code] "
        End If
        
        If strVATNumOrName <> "" And strVATNumOrName <> "0" Then
            If chkShowOnly(1).Value = vbChecked Then
                If Len(strSQL) Then
                    'allanSQL
                    strSQL = strSQL & "AND "
                    strSQL = strSQL & "COMMON.[Taric Code] "
                    strSQL = strSQL & "IN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "SELECT "
                        strSQL = strSQL & "COMMON.[Taric Code] "
                Else
                    'allanSQL
                    strSQL = vbNullString
                    strSQL = strSQL & "SELECT "
                    strSQL = strSQL & "COMMON.[Taric Code], "
                    strSQL = strSQL & "[Key " & strLangOfDesc & "], "
                    strSQL = strSQL & "[Desc " & strLangOfDesc & "] "
                End If
                
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "COMMON, "
                strSQL = strSQL & "CLIENTS "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "COMMON.[Taric Code] = CLIENTS.[Taric Code] "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "CLIENTS.[VAT Num Or Name] = " & Chr(39) & ProcessQuotes(GetVAT(strVATNumOrName)) & Chr(39) & " "
                
                If InStr(1, strSQL, "IN(") Then
                    strSQL = strSQL & ") "
                End If
            End If
        End If
        
        If strCtryCode <> "" And strCtryCode <> "0" Then
            If chkShowOnly(2).Value = vbChecked Then
                If Len(strSQL) Then
                    strSQL = strSQL & " AND COMMON.[Taric Code] IN(SELECT * FROM qdfCommImpExp"
                    strCommImpExpSQL = "SELECT COMMON.[Taric Code] "
                    strUnionSQL = "UNION SELECT COMMON.[Taric Code] "
                Else
                    strSQL = "SELECT * FROM qdfCommImpExp"
                    strCommImpExpSQL = "SELECT COMMON.[Taric Code], [Key " & strLangOfDesc & "], [Desc " & strLangOfDesc & "] "
                    strUnionSQL = "UNION SELECT COMMON.[Taric Code], [Key " & strLangOfDesc & "], [Desc " & strLangOfDesc & "] "
                End If
                
                If chkShowOnly(0).Value = vbChecked Then
                    strCommImpExpSQL = strCommImpExpSQL & "FROM COMMON, " & strTableType & " " & _
                             "WHERE COMMON.[Taric Code] = " & strTableType & ".[Taric Code] " & _
                             "AND " & strTableType & ".[Ctry Code] = " & Chr(39) & ProcessQuotes(strCtryCode) & Chr(39)
                Else
                    strCommImpExpSQL = strCommImpExpSQL & "FROM COMMON, IMPORT " & _
                             "WHERE COMMON.[Taric Code] = IMPORT.[Taric Code] " & _
                             "AND IMPORT.[Ctry Code] = '" & strCtryCode & "' " & _
                             strUnionSQL & "FROM COMMON, EXPORT " & _
                             "WHERE COMMON.[Taric Code] = EXPORT.[Taric Code] " & _
                             "AND EXPORT.[Ctry Code] = " & Chr(39) & ProcessQuotes(strCtryCode) & Chr(39)
                End If
                
                ExecuteNonQuery m_conTaric, "CREATE VIEW [qdfCommImpExp] AS " & strCommImpExpSQL
                'Set qdfCommImpExp = m_conTaric.CreateQueryDef("qdfCommImpExp", strCommImpExpSQL)
                
                If InStr(1, strSQL, "IN(") Then
                    strSQL = strSQL & ") "
                End If
            End If
        End If
        
        If Len(strSQL) = 0 Then
            strSQL = "SELECT [TARIC Code], [Key " & strLangOfDesc & "], [Desc " & strLangOfDesc & "] " & _
                     "FROM COMMON "
        End If
        
        strSQL = strSQL & " ORDER BY COMMON.[Taric Code]"
                
        ADORecordsetOpen strSQL, m_conTaric, rstTaricCodes, adOpenKeyset, adLockOptimistic
        'Set rstTaricCodes = m_conTaric.OpenRecordset(strSQL, dbOpenForwardOnly)
        With rstTaricCodes
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do Until .EOF
                    Set itmListItem = lvwPicklist.ListItems.Add(, , .Fields("Taric Code").Value)
                    itmListItem.ListSubItems.Add , , IIf(IsNull(.Fields("Key " & strLangOfDesc).Value), "", UCase(.Fields("Key " & strLangOfDesc).Value))
                    itmListItem.ListSubItems.Add , , IIf(IsNull(.Fields("Desc " & strLangOfDesc).Value), "", .Fields("Desc " & strLangOfDesc).Value)
                    
                    .MoveNext
                Loop
            End If
        End With
        
'        MsgBox lvwPicklist.ListItems.Count    ' For debugging purposes!
        
        mblnClientLowerLeftChanged = True
        
        If lvwPicklist.ListItems.Count Then
            cmdSelectCancel(0).Enabled = True
            cmdPicklist(sbpModify).Enabled = True
        Else
            cmdSelectCancel(0).Enabled = False
            cmdPicklist(sbpModify).Enabled = False
        End If
        
        On Error Resume Next
        
        ExecuteNonQuery m_conTaric, "DROP VIEW [qdfCommImpExp] "
        'm_conTaric.QueryDefs.Delete "qdfCommImpExp"
        
        Set itmListItem = Nothing
        
        ADORecordsetClose rstTaricCodes
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdPicklist_Click(Index As Integer)
' ********** 10/17/02 **********
' Modified various parts of module to
' accomodate possible new version of
' KLUWER. Config.sgm will be replaced
' by Config.xml, the Country name shall
' no longer be used but instead the country
' code will be supplied. The items in the ECO
' return shall not be found in parenthesis anymore.
' **********   END.   **********

    Dim itmFound As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intLastItemIndex As Integer
    
    Dim strPattern As String
    
    '-----> For TARBEL: Table of Contents HTML page; for Kluwer: config.sgm text file.
    Dim strFileName As String
    '-----> For TARBEL: HTML browser; for Kluwer: Kluwer database application.
    Dim strDefaultBrowser As String
    '-----> For TARBEL: dummy variable; for Kluwer: current directory.
    Dim strDummyDirectory As String
    '-----> For TARBEL: HTML file was opened successfully; for Kluwer: backslash before the Kluwer database application file name was found.
    Dim lngRetVal As Long
    
    '-----> For Kluwer
    Dim intFreeFile As Integer
    Dim intSubscript As Integer
    Dim strTempText() As String
    Dim strClipboard As String
    Dim strPathName As String
    
    Dim udtBrowseInfo As BROWSEINFO
    Dim pIDList As Long
    
    blnOKWasPressed = False    ' Reinitialize before calling frm_taricmain
    
    Select Case Index
        Case sbpFind
            With lvwPicklist.ListItems
                intLastItemIndex = .Count
                lvwPicklistFilter.ListItems.Clear
                
                If intLastItemIndex Then
                    Screen.MousePointer = vbHourglass
                    
                    For intItemIndex = 1 To intLastItemIndex
                        Set itmFound = lvwPicklistFilter.ListItems.Add(, , .Item(intItemIndex).Text)
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                    Next
                    
                    .Clear
                    
                    With lvwPicklistFilter.ListItems
                        For intItemIndex = 1 To intLastItemIndex
                            strPattern = "*" & Trim(txtDescription.Text) & "*"
                            If .Item(intItemIndex).ListSubItems(2) Like strPattern Then
                                Set itmFound = lvwPicklist.ListItems.Add(, , .Item(intItemIndex).Text)
                                itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                                itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                            End If
                        Next
                    End With
                    
                    mblnClientLowerLeftChanged = True
                    
                    If .Count Then
                        .Item(1).Selected = True
                        cmdSelectCancel(0).Enabled = True
                        cmdPicklist(sbpModify).Enabled = True
                    Else
                        cmdSelectCancel(0).Enabled = False
                        cmdPicklist(sbpModify).Enabled = False
                    End If
                    
                    cmdPicklist(sbpFind).Enabled = False
                    
                    Set itmFound = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End With
            
            blnFindWasClicked = True
        Case sbpCNCodes
            Me.MousePointer = vbHourglass
            gstrTaricCNCallType = Me.Name
            frm_tariccn.Show vbModal, Me
            Me.MousePointer = vbDefault
        Case sbpKluwer
            Select Case intThirdPartyDatabase
                Case 1    '-----> TARBEL
                    strFileName = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchHTML"), GetSetting(App.Title, "Third Party Database", "FrenchHTML"))
                    strDefaultBrowser = Space$(255)
                    
                    '-----> Find the application associated with HTML files.
                    lngRetVal = FindExecutable(strFileName, strDummyDirectory, strDefaultBrowser)
                    strDefaultBrowser = Trim$(Replace(strDefaultBrowser, vbNullChar, " "))
                    
                    '-----> If an application is found, launch it!
                    If lngRetVal <= 32 Or Len(strDefaultBrowser) = 0 Then
                        MsgBox Translate(1040), vbExclamation   '"Could not find associated browser."
                    Else
                        lngRetVal = ShellExecute(Me.hwnd, "open", strFileName, 0&, strDummyDirectory, SW_SHOWNORMAL)
                        
                        If lngRetVal <= 32 Then
                            MsgBox Translate(1041), vbExclamation   '"HTML file cannot be opened."
                        End If
                    End If
                Case 2    '-----> Kluwer
                    strFileName = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchFile"), GetSetting(App.Title, "Third Party Database", "FrenchFile"))
                    
                    'On Error GoTo FileErrHandler
                    
                    ReDim strTempText(12)
                    
                    ' These default values will be used if config.sgm doesn't exist or
                    ' if it does indeed exist but has missing lines or no lines at all.
                    strTempText(1) = "<Config>"
                    strTempText(4) = "<ThirdWindow>0</ThirdWindow>"
                    
                    If strLangOfDesc = "French" Then
                        strTempText(5) = "<Vertical>0</Vertical>"
                    Else
                        strTempText(5) = "<Verticaal>0</Verticaal>"
                    End If
                    
                    strTempText(6) = "<InstellingenTar>0</InstellingenTar>"
                    strTempText(7) = "<IncludeTar>0000000000</IncludeTar>"
                    strTempText(8) = "<StartCX>640</StartCX>"
                    strTempText(9) = "<StartCY>480</StartCY>"
                    strTempText(10) = "<PiramidReport>0</PiramidReport>"
                    strTempText(11) = "<NoStartMessage>0</NoStartMessage>"
                    strTempText(12) = "</Config>"
                    
                    If Len(Dir(strFileName)) Then
                        intFreeFile = FreeFile()
                        
                        Open strFileName For Input As #intFreeFile
                        
                        Do Until EOF(intFreeFile)
                            intSubscript = intSubscript + 1
                            
                            If intSubscript > 12 Then
                                ReDim Preserve strTempText(intSubscript)
                            End If
                            
                            Line Input #intFreeFile, strTempText(intSubscript)
                        Loop
                        
                        Close #intFreeFile
                    End If
                    
                    intFreeFile = FreeFile()
                    
                    Open strFileName For Output As #intFreeFile
                    
                    For intSubscript = 1 To 12
                        If intSubscript = 2 Then
                            If lvwPicklist.ListItems.Count > 0 Then
                                strClipboard = "<SelectCode>" & lvwPicklist.SelectedItem & "</SelectCode>"
                            Else
                                strClipboard = "<SelectCode></SelectCode>"
                            End If
                        ElseIf intSubscript = 3 Then
                            If strLangOfDesc = "French" Then
                                ' ***** 10/18/02 *****
                                ' If config file is .sgm then old version, therefore use country name
                                ' If .xml then newer version, therefore use country code.
                                ' *****   end.   *****
                                'strClipboard = "<SelectPays>" & strCtryName & "</SelectPays>"
                                If Right(Trim(strFileName), 3) = "sgm" Then
                                    strClipboard = "<SelectPays>" & strCtryName & "</SelectPays>"
                                Else
                                    strClipboard = "<SelectPays>" & strCtryCode & "</SelectPays>"
                                End If
                            Else
                                ' ***** 10/18/02 *****
                                ' If config file is .sgm then old version, therefore use country name
                                ' If .xml then newer version, therefore use country code.
                                ' *****   end.   *****
                                'strClipboard = "<SelectLand>" & strCtryName & "</SelectLand>"
                                If Right(Trim(strFileName), 3) = "sgm" Then
                                    strClipboard = "<SelectLand>" & strCtryName & "</SelectLand>"
                                Else
                                    strClipboard = "<SelectLand>" & strCtryCode & "</SelectLand>"
                                End If
                            End If
                            
                            ' Replace special characters with characters as they appear in config.sgm.
                            strClipboard = Replace(strClipboard, "&", "&amp;")
                            strClipboard = Replace(strClipboard, "'", "&rquo;")
                            strClipboard = Replace(strClipboard, "-", "&ndash;")
                            strClipboard = Replace(strClipboard, "é", "&eacute;")
                            strClipboard = Replace(strClipboard, "ç", "‡")
                            strClipboard = Replace(strClipboard, "ê", "ˆ")
                            strClipboard = Replace(strClipboard, "ë", "‰")
                            strClipboard = Replace(strClipboard, "è", "Š")
                            strClipboard = Replace(strClipboard, "ï", "‹")
                            strClipboard = Replace(strClipboard, "î", "Œ")
                            strClipboard = Replace(strClipboard, "ô", "“")
                            strClipboard = Replace(strClipboard, "ö", "”")
                        Else
                            strClipboard = strTempText(intSubscript)
                        End If
                        
                        Print #intFreeFile, strClipboard
                    Next
                    
                    Close #intFreeFile
                    
                    '-----> Launch Kluwer program
                    strDefaultBrowser = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchCmd"), GetSetting(App.Title, "Third Party Database", "FrenchCmd"))
                    
                    strTempText() = Split(strDefaultBrowser, " -")
                    strTempText(0) = Replace(strTempText(0), Chr(34), "")
                    
                    lngRetVal = InStrRev(strFileName, "\")
                    
                    If lngRetVal Then
                        strPathName = Left(strFileName, lngRetVal - 1)
                    Else
                        strPathName = "C:\Program Files\DBTARN"    ' Default directory.
                    End If
                    
                    strFileName = ""
                    strDummyDirectory = CurDir()
                    
                    ChDir strPathName          ' Change current directory.
                    Shell strDefaultBrowser, vbNormalFocus
                    ChDir strDummyDirectory    ' Revert to previous current directory.
            End Select
        Case sbpAdd
            If blnPassBoxValues Then
                gstrTaricMainCallType = "AddFull" & intArrayCtr & "/" & Me.Name
            Else
                gstrTaricMainCallType = "AddBlank/" & Me.Name
            End If
            frm_taricmain.Show vbModal, Me
            
            If blnOKWasPressed Then
                ' Refresh lvwPicklist! (already done in frm_taricmain)
                
                blnWasInvokedFromCode = True
                With lvwPicklist.SelectedItem
                    txtTaricCode.Text = .Text
                    txtKeyword.Text = .ListSubItems(1).Text
                    txtDescription.Text = .ListSubItems(2).Text
                    .Tag = "Added"
                End With
                blnWasInvokedFromCode = False
                
                cmdSelectCancel(0).Enabled = True
                cmdPicklist(sbpModify).Enabled = True
            End If
        Case sbpModify
            gstrTaricMainCallType = cmdPicklist(sbpModify).Tag & Me.Name
            frm_taricmain.Show vbModal, Me
            
            If blnOKWasPressed Then
                lvwPicklist.SelectedItem.Tag = "Modified"
            End If
    End Select
    
    Exit Sub
    
FileErrHandler:
    
    Select Case Err.Number
        Case 53    ' File not found
            If MsgBox(Err.Description & ":" & vbCrLf & strTempText(0), vbRetryCancel + vbExclamation, Err.Source & " (" & Err.Number & ")") = vbRetry Then
                Resume
            End If
        Case 76    ' Path not found
            If Len(strFileName) Then
                strPathName = strFileName
                lngRetVal = InStrRev(strPathName, "\")
                strFileName = Mid(strPathName, lngRetVal)
            End If
            
            With udtBrowseInfo
                .hWndOwner = Me.hwnd
                .lpszTitle = "The path " & strPathName & " could not be found.  Please select a folder below, then click OK."
                .ulFlags = BIF_RETURNONLYFSDIRS
                .pIDListRoot = 0
            End With
            
            strPathName = String(255, vbNullChar)
            pIDList = SHBrowseForFolder(udtBrowseInfo)
            
            If SHGetPathFromIDList(pIDList, strPathName) Then
                
                If Mid(strPathName, InStr(1, strPathName, vbNullChar) - 1, 1) = "\" Then
                    strPathName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 2)
                Else
                    strPathName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 1)
                End If
                
                If Len(strFileName) Then
                    'strFileName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 1) & strFileName
                    strFileName = strPathName & strFileName
                Else
                    'strPathName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 1)
                    strPathName = strPathName
                End If
                
                Resume
            End If
        Case Else
            If MsgBox(Err.Description, vbRetryCancel + vbExclamation, Err.Source & " (" & Err.Number & ")") = vbRetry Then
                Resume
            End If
    End Select
End Sub

Private Sub cmdSelectCancel_Click(Index As Integer)
    If Index = 0 Then
        With lvwPicklist
            If Not .SelectedItem Is Nothing Then
                Call SelectTaricCode(.SelectedItem.Text, .SelectedItem.ListSubItems(2).Text)
'                Call SaveFilterPreferences
                
                If Len(.SelectedItem.Tag) Then
                    Select Case strDocType
                        Case "Import"
                            newform(intArrayCtr).blnItemWasModified = True
                        Case "Export"
                            newformE(intArrayCtr).blnItemWasModified = True
                        Case "Transit"
                            newformT(intArrayCtr).blnItemWasModified = True
                    End Select
                End If
            End If
        End With
    Else
        Call SelectTaricCode(strBoxValue)
    End If
    
    Unload Me
End Sub

Private Sub Form_Initialize()
    Dim strBoxPropFull As String
    Dim strBoxProps() As String
    
    Dim strTabCaption As String
    
    Dim strSQL As String
    Dim strBoxCode As String
    Dim rstPicklist As ADODB.Recordset
    
    Const L1_BOX_INDEX = 0
    
    Screen.MousePointer = vbHourglass
    
    If Not OpenMDB(g_objDataSourceProperties, Me, DBInstanceType_DATABASE_SADBEL) Then
        MsgBox Translate(1042), vbRetryCancel, Translate(290) '"Database cannot be opened!"
    End If
    If Not OpenMDB(g_objDataSourceProperties, Me, DBInstanceType_DATABASE_TEMPLATE) Then
        MsgBox Translate(1042), vbRetryCancel, Translate(290) '"Database cannot be opened!"
    End If
    
    strBoxPropFull = GetSetting(App.Title, "Settings", "BoxProperty", "")
    
    If Len(Trim(strBoxPropFull)) Then
        strBoxProps() = Split(strBoxPropFull, "#*#")
        
        strTabCaption = strBoxProps(0)
        intArrayCtr = CInt(strBoxProps(1))
        blnPassBoxValues = CBool(strBoxProps(2))
        ' strBoxProps(3) is reserved
        strDocType = strBoxProps(4)
        strDocName = strBoxProps(5)
        strFolderID = strBoxProps(6)
        strBoxValue = strBoxProps(7)
        strLangOfDesc = strBoxProps(8)
        strCtryCode = strBoxProps(9)
        strVATNumber = strBoxProps(10)
        strRecipientName = strBoxProps(11)
        
        If strCtryCode <> "" And strCtryCode <> "0" Then
            ' SELECT STATEMENT for strCtryName
            strBoxCode = IIf(strDocType = "Import", "C1", "C2")
            strSQL = "SELECT [Description " & strLangOfDesc & "] AS Description " & _
                     "FROM [PICKLIST DEFINITION] AS PickDef " & _
                     "INNER JOIN [PICKLIST MAINTENANCE " & strLangOfDesc & "] AS PickMaint " & _
                     "ON PickDef.[Internal Code] = PickMaint.[Internal Code] " & _
                     "WHERE PickDef.[Box Code] = " & Chr(39) & ProcessQuotes(strBoxCode) & Chr(39) & " AND PickDef.[DOCUMENT] = " & Chr(39) & ProcessQuotes(IIf(strDocType = "Import", "Import", "Export/Transit")) & Chr(39) & " " & _
                     "AND PickMaint.[Code] = " & Chr(39) & ProcessQuotes(strCtryCode) & Chr(39)
            ADORecordsetOpen strSQL, m_conSADBEL, rstPicklist, adOpenKeyset, adLockOptimistic
            'Set rstPicklist = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
            If Not (rstPicklist.EOF And rstPicklist.BOF) Then
                rstPicklist.MoveFirst
                
                strCtryName = rstPicklist![Description]
            End If
        End If
        
        If strVATNumber = "" Or strVATNumber = "0" Or strVATNumber = "000000000" Or strVATNumber Like "796??????" Then
            strClientName = strRecipientName
        Else
            ' SELECT STATEMENT for strClientName
            strBoxCode = IIf(strDocType = "Import", "D8", "D6")
            strSQL = "SELECT [Description " & strLangOfDesc & "] AS Description " & _
                     "FROM [PICKLIST DEFINITION] AS PickDef " & _
                     "INNER JOIN [PICKLIST MAINTENANCE " & strLangOfDesc & "] AS PickMaint " & _
                     "ON PickDef.[Internal Code] = PickMaint.[Internal Code] " & _
                     "WHERE PickDef.[Box Code] = " & Chr(39) & ProcessQuotes(strBoxCode) & Chr(39) & " AND PickDef.[DOCUMENT] = " & Chr(39) & ProcessQuotes(IIf(strDocType = "Import", "Import", "Export/Transit")) & Chr(39) & " " & _
                     "AND PickMaint.[Code] = " & Chr(39) & ProcessQuotes(strVATNumber) & Chr(39)
            
            ADORecordsetOpen strSQL, m_conSADBEL, rstPicklist, adOpenKeyset, adLockOptimistic
            'Set rstPicklist = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
            If Not (rstPicklist.EOF And rstPicklist.BOF) Then
                rstPicklist.MoveFirst
                strClientName = rstPicklist![Description]
            End If
        End If
        
        If Left$(strTabCaption, 1) = csDetl Then
            Select Case strDocType
                Case "Import"
                    newform(intArrayCtr).Text2_GotFocus L1_BOX_INDEX
                Case "Export"
                    newformE(intArrayCtr).Text2_GotFocus L1_BOX_INDEX
                Case "Transit"
                    newformT(intArrayCtr).Text2_GotFocus L1_BOX_INDEX
            End Select
        End If
        
        ADORecordsetClose rstPicklist
    End If
End Sub

Private Sub Form_Load()
    Dim blnSelectedItemExists As Boolean
    
    Dim intBoxValueLength As Integer
    Dim intCharLength As Integer
    
    Dim strTaricCode As String
    Dim strKeyword As String
    Dim strDescription As String
    
    Dim strSQL As String
    Dim rstUsers As ADODB.Recordset
    
    '<<< dandan 112306
    '<<< Update with database password
    'Set m_conTaric = OpenDatabase(cAppPath & "\mdb_taric.mdb")
    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTaric, cAppPath, "mdb_taric.mdb"
                        
    Call LoadResStrings(Me, True)
    
' ********** Initialize chkShowOnly(0) **********
    chkShowOnly(0).Caption = IIf(strDocType = "Import", "&Import", "&Export/Transit")
    
    strClient = Translate(816) & " "
    strCountry = Translate(822) & " "
    
' ********** Initialize chkShowOnly(1) **********
    If Len(strClientName) Then
        If strClientName <> strRecipientName Then
            'EDITED BY ALG May 26, 2003
            chkShowOnly(1).Caption = strClient '& strVATNumber & " - " & Replace(strClientName, "&", "&&")
            strVATNumOrName = strVATNumber
        Else
            chkShowOnly(1).Caption = strClient '& Replace(strClientName, "&", "&&")
            strVATNumOrName = strClientName
            
'            If strVATNumOrName = "0" Or strVATNumOrName = "" Then
'                chkShowOnly(1).Enabled = False
'            End If
        End If
    Else
        chkShowOnly(1).Caption = strClient '& strVATNumber
        strVATNumOrName = strVATNumber
    End If
    
    If strVATNumOrName = "" Or strVATNumOrName = "0" Then
        chkShowOnly(1).Enabled = False
    End If
    
' ********** Initialize chkShowOnly(2) **********
    If strCtryCode <> "0" And Len(strCtryName) Then
        chkShowOnly(2).Caption = strCountry '& strCtryCode & " - " & Replace(strCtryName, "&", "&&")
    Else
        chkShowOnly(2).Caption = strCountry '& strCtryCode
    End If
    
    If strCtryCode = "" Or strCtryCode = "0" Then
        chkShowOnly(2).Enabled = False
    End If
    
    'strSQL = "SELECT [Show Only DocType], [Show Only VATNum], [Show Only CtryCode] " & _
             "FROM USERS WHERE [User No] = '" & cUserNo & "'" 'usernoedit by alg
    strSQL = "SELECT [Show Only DocType], [Show Only VATNum], [Show Only CtryCode] " & _
             "FROM USERS WHERE [User_ID] = " & CStr(lngUserNo)
    
    ADORecordsetOpen strSQL, m_conTemplate, rstUsers, adOpenKeyset, adLockOptimistic
    'Set rstUsers = m_conTemplate.OpenRecordset(strSQL, dbOpenForwardOnly)
    With rstUsers
        blnWasInvokedFromCode = True
        
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            chkShowOnly(0).Value = IIf(![Show Only DocType], vbChecked, vbUnchecked)
            chkShowOnly(1).Value = IIf(![Show Only VATNum], vbChecked, vbUnchecked)
            
            blnWasInvokedFromCode = False
            
            chkShowOnly(2).Value = IIf(![Show Only CtryCode], vbChecked, vbUnchecked)
        Else
            chkShowOnly(0).Value = vbUnchecked
            chkShowOnly(1).Value = vbUnchecked
            
            blnWasInvokedFromCode = False
            
            chkShowOnly(2).Value = vbUnchecked
        End If
    End With
    
    '-----> Initialize Kluwer button
    intThirdPartyDatabase = IIf(GetSetting(App.Title, "Third Party Database", "Usage") = "", 0, GetSetting(App.Title, "Third Party Database", "Usage"))
    
    Select Case intThirdPartyDatabase
        Case 0
            cmdPicklist(2).Enabled = False
        Case 1
            cmdPicklist(2).Enabled = True
            cmdPicklist(2).Caption = "TARBEL..."
        Case 2
            cmdPicklist(2).Enabled = True
            cmdPicklist(2).Caption = "Kluwer..."
    End Select
    
' ********** Initialize lvwPicklist **********
    With lvwPicklist.ColumnHeaders
        .Item(1).Text = Translate(861)    ' TARIC Code
        .Item(2).Text = Translate(829)    ' Keyword
        .Item(3).Text = Translate(292)    ' Description
        
        ReDim msngColumnHeaderWidths(.Count)
        
        msngColumnHeaderWidths(1) = .Item(1).Width
        msngColumnHeaderWidths(2) = .Item(2).Width
        msngColumnHeaderWidths(3) = .Item(3).Width
    End With
    
    lvwPicklist.Sorted = True
    mblnClientLowerLeftChanged = True
    blnSelectedItemExists = Not lvwPicklist.SelectedItem Is Nothing
    
' ********** Initialize txtTaricCode, txtKeyword, txtDescription **********
    intBoxValueLength = Len(strBoxValue)
    If intBoxValueLength Then
        For intCharLength = 1 To intBoxValueLength
            txtTaricCode.Text = Left$(strBoxValue, intCharLength)
        Next
        
        txtTaricCode.SelStart = intBoxValueLength
        
        If blnSelectedItemExists Then
            With lvwPicklist.SelectedItem
                strTaricCode = .Text
                strKeyword = .ListSubItems(1).Text
                strDescription = .ListSubItems(2).Text
            End With
            
            If strTaricCode = strBoxValue Then
                blnWasInvokedFromCode = True
                txtKeyword.Text = strKeyword
                txtDescription.Text = strDescription
                blnWasInvokedFromCode = False
                
                lvwPicklist.TabIndex = 0
            End If
        End If
    Else
        If blnSelectedItemExists Then
            With lvwPicklist
                .SortKey = 1
                .SortOrder = lvwAscending
                .ListItems(1).Selected = True
            End With
            
            txtKeyword.TabIndex = 0
            txtTaricCode.TabIndex = Controls.Count - 1
        End If
    End If
    
' ********** Initialize cmdPicklist, cmdSelectCancel **********
    If blnSelectedItemExists Then
        cmdSelectCancel(0).Enabled = True
        cmdPicklist(sbpModify).Enabled = True
'        cmdPicklist(sbpModify).Tag = IIf(lMaintainTables, "Modify/", "View/")
    End If
    
    cmdPicklist(sbpModify).Tag = IIf(lMaintainTables, "Modify/", "View/")
    
    If lMaintainTables Then
        cmdPicklist(sbpAdd).Enabled = True
    End If
    
    Set rstUsers = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Call SelectTaricCode(strBoxValue)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFilterPreferences
    Call UnloadControls(Me, True)
    
    ADODisconnectDB m_conTaric
    ADODisconnectDB m_conSADBEL
    
    Set frm_taricpicklist = Nothing
End Sub

Private Sub lvwPicklist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwPicklist
        If Not .SelectedItem Is Nothing Then
            ColumnHeader.Tag = ColumnHeader.Tag Xor lvwDescending    ' Reverses current .SortOrder stored in .Tag
            
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = ColumnHeader.Tag
            .SelectedItem.EnsureVisible
        End If
    End With
End Sub

Private Sub lvwPicklist_DblClick()
    With lvwPicklist
        If Not .HitTest(msngMousePointerX, msngMousePointerY) Is Nothing Then
'        If Not .SelectedItem Is Nothing Then
            Call SelectTaricCode(.SelectedItem.Text, .SelectedItem.ListSubItems(2).Text)
'            Call SaveFilterPreferences
            
            If Len(.SelectedItem.Tag) Then
                Select Case strDocType
                    Case "Import"
                        newform(intArrayCtr).blnItemWasModified = True
                    Case "Export"
                        newformE(intArrayCtr).blnItemWasModified = True
                    Case "Transit"
                        newformT(intArrayCtr).blnItemWasModified = True
                End Select
            End If
            
            Unload Me
        End If
    End With
End Sub

Private Sub lvwPicklist_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnWasInvokedFromCode = True
    With Item
        txtTaricCode.Text = .Text
        txtKeyword.Text = .ListSubItems(1).Text
        txtDescription.Text = .ListSubItems(2).Text
    End With
    blnWasInvokedFromCode = False
    
    ' Enable CommandButtons here!
End Sub

Private Sub lvwPicklist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngMousePointerX = x
    msngMousePointerY = Y
End Sub

Private Sub lvwPicklist_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not lvwPicklist.SelectedItem Is Nothing Then
        lvwPicklist.SelectedItem.Selected = True
    End If
End Sub

Private Sub txtDescription_Change()
    Dim itmFound As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intLastItemIndex As Integer
    
    If Not blnWasInvokedFromCode Then
        If Len(Trim(txtDescription.Text)) Then
            If Not blnFindWasClicked And lvwPicklist.ListItems.Count Then
                cmdPicklist(sbpFind).Enabled = True
            End If
        Else
            cmdPicklist(sbpFind).Enabled = False
            
            With lvwPicklistFilter.ListItems
                intLastItemIndex = .Count
                
                If intLastItemIndex And blnFindWasClicked Then
                    Screen.MousePointer = vbHourglass
                    lvwPicklist.ListItems.Clear
                    
                    For intItemIndex = 1 To intLastItemIndex
                        Set itmFound = lvwPicklist.ListItems.Add(, , .Item(intItemIndex).Text)
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                    Next
                    
                    .Item(1).Selected = True
                    cmdSelectCancel(0).Enabled = True
                    cmdPicklist(sbpModify).Enabled = True
                    
                    blnFindWasClicked = False
                    Set itmFound = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End With
        End If
    End If
End Sub

Private Sub txtKeyword_Change()
    Dim itmFound As MSComctlLib.ListItem
    Dim itmFirstVisible As MSComctlLib.ListItem
    Dim itmLastVisible As MSComctlLib.ListItem
    Dim itmLastIntegral As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intFoundIndex As Integer
    Dim intFirstVisibleIndex As Integer
    Dim intLastIntegralIndex As Integer
    Dim intLastItemIndex As Integer
    
    Dim strPattern As String
    
    Dim intScrollValue As Integer
    Dim intPixelDiff As Integer
    
    If Not blnWasInvokedFromCode Then
        With lvwPicklist
            If Len(Trim(txtKeyword.Text)) Then
                .SortKey = 1
                .SortOrder = lvwAscending
            Else
                .SortKey = 0
                .SortOrder = lvwAscending
            End If
            
            strPattern = Trim(txtKeyword.Text) & "*"
            intLastItemIndex = .ListItems.Count
            
            For intItemIndex = 1 To intLastItemIndex
                If .ListItems(intItemIndex).ListSubItems(1) Like strPattern Then
                    Set itmFound = .ListItems(intItemIndex)
                    itmFound.Selected = True
                    itmFound.EnsureVisible
                    intFoundIndex = itmFound.Index
                    
                    Set itmFirstVisible = .GetFirstVisible
                    intFirstVisibleIndex = itmFirstVisible.Index
                    
                    intScrollValue = intFoundIndex - intFirstVisibleIndex
                    
                    If .ColumnHeaders(1).Width <> msngColumnHeaderWidths(1) Or _
                       .ColumnHeaders(2).Width <> msngColumnHeaderWidths(2) Or _
                       .ColumnHeaders(3).Width <> msngColumnHeaderWidths(3) Then
                        msngColumnHeaderWidths(1) = .ColumnHeaders(1).Width
                        msngColumnHeaderWidths(2) = .ColumnHeaders(2).Width
                        msngColumnHeaderWidths(3) = .ColumnHeaders(3).Width
                        
                        mblnClientLowerLeftChanged = True
                    End If
                    
                    If mblnClientLowerLeftChanged Then
                        Do
                            intPixelDiff = intPixelDiff + 1
                            mintClientLowerLeft = .Height - intPixelDiff
                            Set itmLastVisible = .HitTest(.Left, mintClientLowerLeft)
                        Loop While itmLastVisible Is Nothing
                        
                        mblnClientLowerLeftChanged = False
                    End If
                    
                    Set itmLastIntegral = .HitTest(.Left, mintClientLowerLeft - itmFound.Height + 1)
                    intLastIntegralIndex = itmLastIntegral.Index
                    
                    If intLastItemIndex >= intLastIntegralIndex + intScrollValue Then
                        .ListItems(intLastIntegralIndex + intScrollValue).EnsureVisible
                    Else
                        .ListItems(intLastItemIndex).EnsureVisible
                    End If
                    
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub txtTaricCode_Change()
    Dim itmFound As MSComctlLib.ListItem
    Dim itmFirstVisible As MSComctlLib.ListItem
    Dim itmLastVisible As MSComctlLib.ListItem
    Dim itmLastIntegral As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intFoundIndex As Integer
    Dim intFirstVisibleIndex As Integer
    Dim intLastIntegralIndex As Integer
    Dim intLastItemIndex As Integer
    
    Dim strPattern As String
    
    Dim intScrollValue As Integer
    Dim intPixelDiff As Integer
    
    If Not blnWasInvokedFromCode Then
        With lvwPicklist
            .SortKey = 0
            .SortOrder = lvwAscending
            
            strPattern = Trim(txtTaricCode.Text) & "*"
            intLastItemIndex = .ListItems.Count
            
            For intItemIndex = 1 To intLastItemIndex
                If .ListItems(intItemIndex).Text Like strPattern Then
                    Set itmFound = .ListItems(intItemIndex)
                    itmFound.Selected = True
                    itmFound.EnsureVisible
                    intFoundIndex = itmFound.Index
                    
                    Set itmFirstVisible = .GetFirstVisible
                    intFirstVisibleIndex = itmFirstVisible.Index
                    
                    intScrollValue = intFoundIndex - intFirstVisibleIndex
                    
                    If .ColumnHeaders(1).Width <> msngColumnHeaderWidths(1) Or _
                       .ColumnHeaders(2).Width <> msngColumnHeaderWidths(2) Or _
                       .ColumnHeaders(3).Width <> msngColumnHeaderWidths(3) Then
                        msngColumnHeaderWidths(1) = .ColumnHeaders(1).Width
                        msngColumnHeaderWidths(2) = .ColumnHeaders(2).Width
                        msngColumnHeaderWidths(3) = .ColumnHeaders(3).Width
                        
                        mblnClientLowerLeftChanged = True
                    End If
                    
                    If mblnClientLowerLeftChanged Then
                        Do
                            intPixelDiff = intPixelDiff + 1
                            mintClientLowerLeft = .Height - intPixelDiff
                            Set itmLastVisible = .HitTest(.Left, mintClientLowerLeft)
                        Loop While itmLastVisible Is Nothing
                        
                        mblnClientLowerLeftChanged = False
                    End If
                    
                    Set itmLastIntegral = .HitTest(.Left, mintClientLowerLeft - itmFound.Height + 1)
                    intLastIntegralIndex = itmLastIntegral.Index
                    
                    If intLastItemIndex >= intLastIntegralIndex + intScrollValue Then
                        .ListItems(intLastIntegralIndex + intScrollValue).EnsureVisible
                    Else
                        .ListItems(intLastItemIndex).EnsureVisible
                    End If
                    
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub SaveFilterPreferences()
    Dim rstUsers As ADODB.Recordset
    
    'Set rstUsers = m_conSADBEL.OpenRecordset("USERS", dbOpenTable)
    ADORecordsetOpen GetSQLCommandFromTableName("USERS"), m_conTemplate, rstUsers, adOpenKeyset, adLockOptimistic
    'Set rstUsers = m_conTemplate.OpenRecordset("USERS", dbOpenTable)
    With rstUsers
        If Not (.EOF And .BOF) Then
            .MoveFirst
            .Find "[USER_ID] = " & lngUserNo & " ", , adSearchForward
            
            If Not .EOF Then
            '.Index = "USER_ID"
            '.Seek "=", lngUserNo
            '
            'If Not .NoMatch Then
                '.Edit
                    ![Show Only DocType] = chkShowOnly(0).Value
                    ![Show Only VATNum] = chkShowOnly(1).Value
                    ![Show Only CtryCode] = chkShowOnly(2).Value
                .Update
                
                UpdateRecordset m_conTemplate, rstUsers, "USERS"
            End If
        End If
    End With
    
    ADORecordsetClose rstUsers
End Sub

Private Sub SelectTaricCode(ByVal strTaricCode As String, Optional ByVal strDescription As String)
    Dim strRegSection As String
    
    Select Case strDocType
        Case "Import"
            strRegSection = "CodiSheet"
        Case "Export"
            strRegSection = "ExSheet"
        Case "Transit"
            strRegSection = "TrSheet"
        
        'added by alg - april 8, 2003
        Case "Transit NCTS"
            strRegSection = G_CONST_NCTS1_SHEET
        Case "Combined NCTS"
            strRegSection = G_CONST_NCTS2_SHEET
            
        Case G_CONST_EDINCTS1_TYPE
            strRegSection = G_CONST_EDINCTS1_SHEET
    End Select
    
' ********** Modified September 5, 2001 **********
' ********** The close and open brace combination serves as an escape character for descriptions
' ********** containing a percent sign.
    SaveSetting App.Title, strRegSection, "Pick_" & strDocName & "_" & strFolderID, _
                strTaricCode & "@@" & Replace(strDescription, "%", "}{") & "%" & Len(strTaricCode)
' ********** End Modify **************************
End Sub
