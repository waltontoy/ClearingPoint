VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEntrepot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrepot"
   ClientHeight    =   6210
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   4965
   Icon            =   "frmEntrepot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntrepotCountry 
      Height          =   315
      Left            =   4035
      MaxLength       =   2
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdEntrepotCountry 
      Caption         =   "..."
      Height          =   315
      Left            =   4515
      TabIndex        =   4
      Top             =   600
      Width           =   315
   End
   Begin VB.Frame fraArchiving 
      Caption         =   "Stock Card Archiving"
      Height          =   1245
      Left            =   75
      TabIndex        =   22
      Tag             =   "2208"
      Top             =   4260
      Width           =   4815
      Begin VB.OptionButton optArchiveManual 
         Caption         =   "Manual"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Tag             =   "168"
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optArchiveZero 
         Caption         =   "Upon zero balance"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Tag             =   "2209"
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmEntrepot.frx":08CA
      Left            =   1800
      List            =   "frmEntrepot.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdAuthParty 
      Caption         =   "..."
      Height          =   315
      Left            =   4515
      TabIndex        =   6
      Top             =   960
      Width           =   315
   End
   Begin VB.TextBox txtAuthParty 
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   5
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Settings for Stock Card numbering"
      Height          =   1605
      Left            =   75
      TabIndex        =   20
      Top             =   2610
      Width           =   4815
      Begin VB.TextBox txtSetStartNum 
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1485
      End
      Begin VB.OptionButton optSetUserDef 
         Caption         =   "User-defined"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Tag             =   "2207"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optSetSequential 
         Caption         =   "Use sequential numbering"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Tag             =   "2205"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblSetStartNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Number:"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Tag             =   "2206"
         Top             =   780
         Width           =   1575
      End
   End
   Begin VB.Frame fraValidity 
      Caption         =   "Validity Period"
      Height          =   1245
      Left            =   75
      TabIndex        =   18
      Top             =   1320
      Width           =   4815
      Begin VB.CheckBox chkDateEnd 
         Caption         =   "End Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Tag             =   "2203"
         Top             =   780
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpDateStart 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58523649
         CurrentDate     =   38132
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58523649
         CurrentDate     =   38132
      End
      Begin VB.Label lblDateStart 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
         Height          =   195
         Left            =   390
         TabIndex        =   19
         Tag             =   "2202"
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Tag             =   "179"
      Top             =   5730
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Tag             =   "178"
      Top             =   5730
      Width           =   1215
   End
   Begin VB.TextBox txtAuthNum 
      Height          =   315
      Left            =   1800
      MaxLength       =   17
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblAuthParty 
      BackStyle       =   0  'Transparent
      Caption         =   "Authorised Party :"
      Height          =   195
      Left            =   75
      TabIndex        =   17
      Top             =   1020
      Width           =   1695
   End
   Begin VB.Label lblAuthNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrepot Number :"
      Height          =   195
      Left            =   75
      TabIndex        =   16
      Tag             =   "2273"
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Tag             =   "438"
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "frmEntrepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is called from the Entrepot picklist
Option Explicit

Private cnnEntrepot As ADODB.Connection
Private rstEntrepot As ADODB.Recordset
Private rstExistingEntrepots As ADODB.Recordset

Private pckAuthorized As PCubeLibEntrepot.cAuthorizedParty
Private Lang_Entrepot As String
Private strAuthParty As String
Private mblnCancel As Boolean
Private lngButton As Long
Private bytAuthPartyDeclined As Byte
Public lngAuthID As Long

Private Sub chkDateEnd_Click()
    Select Case chkDateEnd.Value
        Case 0
            dtpDateEnd.Enabled = False
        Case 1
            dtpDateEnd.Enabled = True
            dtpDateEnd.Value = dtpDateStart.Value
    End Select
End Sub

Private Sub cmdAuthParty_Click()
    'Call Vince's Authorized Party procedure here
    Set pckAuthorized = New PCubeLibEntrepot.cAuthorizedParty
    pckAuthorized.ShowAuthorizedParty 1, Me, cnnEntrepot, Lang_Entrepot, ResourceHandler, "txtAuthParty", lngAuthID
    Set pckAuthorized = Nothing
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdEntrepotCountry_Click()
    Dim pckCountry As PCubeLibPick.CPicklist
    Dim gsdCountry As PCubeLibPick.CGridSeed
    Dim strCountrySQL As String
    
    Set pckCountry = New CPicklist
    Set gsdCountry = New CGridSeed
    
    Set gsdCountry = pckCountry.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
    
    ' The primary key is mentioned twice to conform to the design of the picklist class.
    strCountrySQL = ""
    strCountrySQL = strCountrySQL & "SELECT "
    strCountrySQL = strCountrySQL & "Code AS [Key Code], "
    strCountrySQL = strCountrySQL & "Code as [CODE], "
    strCountrySQL = strCountrySQL & "[Description " & IIf(UCase(Lang_Entrepot) = "ENGLISH", "ENGLISH", IIf(UCase(Lang_Entrepot) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Key Description] "
    strCountrySQL = strCountrySQL & "FROM "
    strCountrySQL = strCountrySQL & "[PICKLIST MAINTENANCE " & IIf(UCase(Lang_Entrepot) = "ENGLISH", "ENGLISH", IIf(UCase(Lang_Entrepot) = "FRENCH", "FRENCH", "DUTCH")) & "] "
    strCountrySQL = strCountrySQL & "INNER JOIN "
    strCountrySQL = strCountrySQL & "[PICKLIST DEFINITION] "
    strCountrySQL = strCountrySQL & "ON "
    strCountrySQL = strCountrySQL & "[PICKLIST MAINTENANCE " & IIf(UCase(Lang_Entrepot) = "ENGLISH", "ENGLISH", IIf(UCase(Lang_Entrepot) = "FRENCH", "FRENCH", "DUTCH")) & "].[INTERNAL CODE] = [PICKLIST DEFINITION].[INTERNAL CODE] "
    strCountrySQL = strCountrySQL & "WHERE "
    strCountrySQL = strCountrySQL & "Document = 'PLDA Import' "
    strCountrySQL = strCountrySQL & "AND "
    strCountrySQL = strCountrySQL & "[BOX CODE] = 'M5' "
    With pckCountry
        .Search True, "Key Code", Trim(txtEntrepotCountry.Text)
        
        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick Me, cpiSimplePicklist, cnnEntrepot, strCountrySQL, "Key Code", "Countries", vbModal, gsdCountry, , , True, cpiKeyF2
        
        If Not .SelectedRecord Is Nothing Then
            txtEntrepotCountry.Text = .SelectedRecord.RecordSource.Fields("Key Code").Value
        End If
    End With
    
    Set gsdCountry = Nothing
    Set pckCountry = Nothing
End Sub



Private Sub cmdOK_Click()
'    Dim rstEntrepotVerify As ADODB.Recordset
    Dim strFilter As String
    
    'Checks to make sure all required fields have values.
    If Validation = False Then Exit Sub
    
    'Commit changes.
    MousePointer = vbHourglass
    
    If Not (lngButton = 1) Then
                                       
        strFilter = rstExistingEntrepots.Filter
        rstExistingEntrepots.Filter = 0
        rstExistingEntrepots.Filter = "[Entrepot Type] = '" & cboType.Text & "' AND [Entrepot Number] = '" & txtAuthNum.Text & "'"
        If rstExistingEntrepots.RecordCount > 0 Then
            MousePointer = vbDefault
            'MsgBox "The Entrepot Type and Entrepot Number combination entered already exists." & vbCrLf, _
                   vbOKOnly + vbInformation, "Entrepot"
            MsgBox Translate(2184) & vbCrLf, _
                   vbOKOnly + vbInformation, "Entrepot"
                   
            txtAuthNum.SetFocus
            Exit Sub
        End If
        rstExistingEntrepots.Filter = 0
        rstExistingEntrepots.Filter = IIf(strFilter = "0", 0, strFilter)
'        rstEntrepotVerify.Close
    End If
    
    mblnCancel = False
    
    'Writes values to recordset picklist recordset.
    'Not permanent as clicking cancel on picklist will discard these changes.
    With rstEntrepot
        rstEntrepot("Entrepot Type") = cboType.Text
        rstEntrepot("Entrepot Number") = txtAuthNum.Text
        rstEntrepot("Entrepot Country") = txtEntrepotCountry.Text
        rstEntrepot("Auth_ID") = lngAuthID
        rstEntrepot("Entrepot_StartDate") = dtpDateStart.Value
        If chkDateEnd.Value = 0 Then
            rstEntrepot("Entrepot_EndDate") = "12/30/2999"
        ElseIf chkDateEnd.Value = 1 Then
            rstEntrepot("Entrepot_EndDate") = dtpDateEnd.Value
        End If
            
        If IsNull(txtSetStartNum.Text) Then
            txtSetStartNum.SetFocus
            Exit Sub
        Else
            If IsNumeric(txtSetStartNum.Text) Then
                rstEntrepot("Starting Num") = txtSetStartNum.Text
            Else
                If txtSetStartNum.Enabled = True Then 'Allan Oct22
                    txtSetStartNum.SetFocus
                    Exit Sub
                Else
                    Debug.Assert False
                End If
            End If
        End If
        
        
        If optSetSequential.Value = True Then
            rstEntrepot("Entrepot_StockCard_Numbering") = 0
        ElseIf optSetUserDef.Value = True Then
            rstEntrepot("Entrepot_StockCard_Numbering") = 1
        End If
        If optArchiveZero.Value = True Then
            rstEntrepot("Entrepot_StockCard_Archiving") = 0
        ElseIf optArchiveManual.Value = True Then
            rstEntrepot("Entrepot_StockCard_Archiving") = 1
        End If
    End With
    
    If lngButton = ButtonType.cpiCopy Or lngButton = ButtonType.cpiAdd Then
        If Not IsNull(txtSetStartNum.Text) Then
            rstEntrepot.Fields("Entrepot_LastSeqNum") = txtSetStartNum.Text
        End If
    End If
    
    rstEntrepot.Update
    
    ' TO DO FOR CP.NET
       
    Me.MousePointer = vbHourglass
    Me.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub dtpDateStart_Change()
    'Ensure validity start date earlier than end date.
    If chkDateEnd.Value = 1 Then
        If dtpDateStart.Value > dtpDateEnd.Value Then dtpDateStart.Value = dtpDateEnd.Value
    End If
End Sub

Private Sub dtpDateEnd_Change()
    'Ensure validity start date earlier than end date.
    If dtpDateEnd.Value < dtpDateStart.Value Then dtpDateEnd.Value = dtpDateStart.Value
End Sub

Public Sub Pre_Load(ByVal conn As ADODB.Connection, _
                    ByVal rst As ADODB.Recordset, _
                          Language As String, _
                          Button As PCubeLibPick.ButtonType, _
                          Cancel As Boolean, ByVal MyResourceHandler As Long, ByVal PickRecords As ADODB.Recordset)
                          
    ResourceHandler = MyResourceHandler
    modGlobals.LoadResStrings Me, True
    
    Set cnnEntrepot = conn
    Set rstEntrepot = rst
    Set rstExistingEntrepots = PickRecords
    Lang_Entrepot = Language
    lngButton = Button
    
    
    If Button = cpiModify Or Button = cpiCopy Then
        'Passes values from Entrepot pick to maintainance form.
        LoadValues
    ElseIf Button = cpiAdd Then
        'Load default values.
        optSetSequential.Value = True
        optArchiveZero.Value = True
        dtpDateStart.Value = Format(Now)
        dtpDateEnd.Value = Format(Now)
    End If
    
    'Load maintaince form.
    Me.Show vbModal
    Cancel = mblnCancel
End Sub

Private Sub Form_Load()

    'Ini premade Entrepot type values.
    cboType.AddItem "C"
    cboType.AddItem "D"
    cboType.AddItem "E"
    'Selects first entry.
    cboType.ListIndex = 0
    
    'Initialize Auth_Party flags.
    lngAuthID = 0
    strAuthParty = Empty
    bytAuthPartyDeclined = 0
    
    chkDateEnd_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mblnCancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstExistingEntrepots = Nothing
End Sub

Private Sub optSetSequential_Click()
    'Enables sequential starting number textbox when opt is not selected.
    txtSetStartNum.Enabled = True
    txtSetStartNum.BackColor = txtAuthNum.BackColor
End Sub

Private Sub optSetUserDef_Click()
    'Disables sequential starting number textbox when opt is not selected.
    txtSetStartNum.Enabled = False
    txtSetStartNum.BackColor = Me.BackColor
End Sub

Private Sub LoadValues()
    'Temporary recordset for Auth_ID name resolution.
    'Will be destroyed after passing name to textbox.
    Dim rstAuth As ADODB.Recordset

    ADORecordsetOpen "SELECT Auth_ID, Auth_Name FROM AuthorizedParties WHERE Auth_ID = " & rstEntrepot!Auth_ID, cnnEntrepot, rstAuth, adOpenKeyset, adLockOptimistic
    'rstAuth.Open "SELECT Auth_ID, Auth_Name FROM AuthorizedParties WHERE Auth_ID = " & rstEntrepot!Auth_ID, cnnEntrepot, adOpenKeyset, adLockOptimistic
    
    'Loads selected picklist values to maintanance form.
    If Not (rstAuth.EOF Or rstAuth.BOF) Then
        rstAuth.MoveFirst
        
        txtAuthParty.Text = IIf(IsNull(rstAuth!Auth_Name), "", rstAuth!Auth_Name)
    End If
    ADORecordsetClose rstAuth

    lngAuthID = rstEntrepot!Auth_ID
    
    cboType.Text = IIf(IsNull(rstEntrepot![Entrepot Type]), cboType.Text, rstEntrepot![Entrepot Type])
    txtAuthNum.Text = IIf(IsNull(rstEntrepot![Entrepot Number]), "", rstEntrepot![Entrepot Number])
    txtAuthNum.Tag = IIf(IsNull(rstEntrepot![Entrepot Number]), "", rstEntrepot![Entrepot Number])
    txtEntrepotCountry.Text = IIf(IsNull(rstEntrepot![Entrepot Country]), "", rstEntrepot![Entrepot Country])
    
    dtpDateStart.Value = IIf(IsNull(rstEntrepot!Entrepot_StartDate), Format(Now), rstEntrepot!Entrepot_StartDate)
    If Not (IsNull(rstEntrepot!Entrepot_EndDate)) Then
        If rstEntrepot!Entrepot_EndDate = "12/30/2999" Then
            chkDateEnd.Value = 0
            dtpDateEnd.Value = dtpDateStart.Value
            dtpDateEnd.Enabled = False
        Else
            chkDateEnd.Value = 1
            dtpDateEnd.Value = IIf(IsNull(rstEntrepot!Entrepot_EndDate), Format(Now), rstEntrepot!Entrepot_EndDate)
            dtpDateEnd.Enabled = True
        End If
    End If
    
    If lngButton = 2 Then 'copy button
        txtSetStartNum.Text = 0
    Else
        txtSetStartNum.Text = IIf(IsNull(rstEntrepot.Fields("Starting Num").Value), "", rstEntrepot.Fields("Starting Num").Value)
    End If
    If Not (IsNull(rstEntrepot!Entrepot_StockCard_Numbering)) Then
        Select Case rstEntrepot!Entrepot_StockCard_Numbering
            Case 0
                optSetSequential.Value = True
            Case 1
                optSetUserDef.Value = True
        End Select
    End If
    If Not (IsNull(rstEntrepot!Entrepot_StockCard_Archiving)) Then
        Select Case rstEntrepot!Entrepot_StockCard_Archiving
            Case 0
                optArchiveZero.Value = True
            Case 1
                optArchiveManual.Value = True
        End Select
    End If
End Sub

Private Sub txtAuthNum_LostFocus()
    'Removes unwanted spaces.
    If Len(txtAuthNum.Text) <> 0 Then txtAuthNum.Text = Trim$(txtAuthNum.Text)
End Sub

Private Sub txtAuthParty_GotFocus()
    If bytAuthPartyDeclined = 0 Then strAuthParty = Trim$(UCase(txtAuthParty.Text))
End Sub

Private Sub txtAuthParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdAuthParty_Click
    End If
End Sub

Private Sub txtAuthParty_LostFocus()
    'Only creates a recordset and prompt for quick registration if value has changed.
    If txtAuthParty.Text <> "" Then
        If Not (strAuthParty = Trim$(UCase(txtAuthParty.Text))) Or bytAuthPartyDeclined = 0 Then
            
            Dim rstAuthPartyVerify As ADODB.Recordset
            
            'Query and search database for entered Authorized Party.
            ADORecordsetOpen "SELECT Auth_ID AS [Auth ID], Auth_Name AS [Auth Name] " & _
                                    "FROM AuthorizedParties " & _
                                    "WHERE Auth_Name = '" & txtAuthParty.Text & "'", _
                                    cnnEntrepot, rstAuthPartyVerify, adOpenKeyset, adLockOptimistic
            'rstAuthPartyVerify.Open "SELECT Auth_ID AS [Auth ID], Auth_Name AS [Auth Name] " & _
                                    "FROM AuthorizedParties " & _
                                    "WHERE Auth_Name = '" & txtAuthParty.Text & "'", _
                                    cnnEntrepot, adOpenKeyset, adLockOptimistic
                                    
            If (rstAuthPartyVerify.BOF Or rstAuthPartyVerify.EOF) Then
                'Prompts user to quickly register Authorized Party if not found in database.
                'If MsgBox("The entered Authorized Party is not recognized." & vbCrLf & _
                          "Would you like to add " & UCase(txtAuthParty.Text) & " to the database?", _
                          vbYesNo + vbQuestion, "Entrepot") = vbYes Then
                If MsgBox(Translate(2185) & vbCrLf & _
                          Translate(2186) & Space(1) & UCase(txtAuthParty.Text) & Space(1) & Translate(2187), _
                          vbYesNo + vbQuestion, "Entrepot") = vbYes Then
                          
                    rstAuthPartyVerify.AddNew
                    rstAuthPartyVerify.Fields("Auth Name").Value = txtAuthParty.Text
                    rstAuthPartyVerify.Update
                    
                    'Get Auth_ID after adding Authorized Party to database.
                    'lngAuthID = rstAuthPartyVerify.Fields("Auth ID").Value
                    lngAuthID = InsertRecordset(cnnEntrepot, rstAuthPartyVerify, "AuthorizedParties")
                    
                    bytAuthPartyDeclined = 1
                Else
                    'Sets Auth_ID to 0 so that user cannot add this Entrepot to the picklist.
                    lngAuthID = 0
                End If
                
                ADORecordsetClose rstAuthPartyVerify
                
            Else
                rstAuthPartyVerify.MoveFirst
                
                lngAuthID = rstAuthPartyVerify.Fields("Auth ID").Value
                
                ADORecordsetClose rstAuthPartyVerify
                
            End If
        End If
    Else
        lngAuthID = 0
    End If
End Sub

Private Sub txtEntrepotCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdEntrepotCountry_Click
    End If
End Sub


Private Sub txtSetStartNum_KeyPress(KeyAscii As Integer)
    'Prevents user from entering non-numeric characters.
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End Sub

Private Function Validation() As Boolean
    Dim bytProblemsFlag As Byte
    Dim strProblems As String
    bytProblemsFlag = 0
    Validation = True
    If Len(cboType.Text) = 0 Then
        Validation = False
        'strProblems = strProblems & Space(5) & "* Entrepot Type - Missing" & vbCrLf
        strProblems = strProblems & Space(5) & Translate(2188) & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 1
    End If
    If Len(txtAuthNum.Text) = 0 Then
        Validation = False
        'strProblems = strProblems & Space(5) & "* Entrepot Number - Missing" & vbCrLf
        strProblems = strProblems & Space(5) & Translate(2189) & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 2
    End If
    If Len(txtAuthParty.Text) = 0 Then
        Validation = False
        'strProblems = strProblems & Space(5) & "* Authorized Party - Missing" & vbCrLf
        strProblems = strProblems & Space(5) & Translate(2190) & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 4
    ElseIf lngAuthID = 0 Then
        Validation = False
        'strProblems = strProblems & Space(5) & "* Authorized Party - Not in database" & vbCrLf
        strProblems = strProblems & Space(5) & Translate(2191) & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 4
    End If
    If optSetSequential.Value = True And Len(txtSetStartNum.Text) = 0 Then
        Validation = False
        'strProblems = strProblems & Space(5) & "* Sequential Number - Missing" & vbCrLf
        strProblems = strProblems & Space(5) & Translate(2192) & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 8
    End If
    If Validation = False Then
        'MsgBox "In order to create a valid Entrepot, kindly correct the following item(s):" & vbCrLf & vbCrLf & strProblems, vbOKOnly + vbInformation, "Entrepot"
        MsgBox Translate(2193) & vbCrLf & vbCrLf & strProblems, vbOKOnly + vbInformation, "Entrepot"
    End If
    
    'Checks flags and sets focus to top-most affected control in the list.
    Select Case bytProblemsFlag
        Case 1, 3, 5, 7, 9, 15
            cboType.SetFocus
        Case 2, 6, 10, 14
            txtAuthNum.SetFocus
        Case 4, 12
            txtAuthParty.SetFocus
        Case 8
            txtSetStartNum.SetFocus
    End Select
End Function
