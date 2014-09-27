VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_taricclients 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TARIC - Clients"
   ClientHeight    =   5760
   ClientLeft      =   4725
   ClientTop       =   2475
   ClientWidth     =   6060
   Icon            =   "frm_taricClients.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Tag             =   "856"
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   600
      Width           =   5535
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "865"
            Text            =   "VAT Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "292"
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lblCode 
         Caption         =   "The List below shows all clients that use TARIC codes"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Tag             =   "862"
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Tag             =   "179"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Tag             =   "260"
      Top             =   5280
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Tag             =   "817"
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clients"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_taricclients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----> for testing
Public strLangOfDesc As String
'Public CAppPath As String
'-----> End testing

Dim m_rstClient As ADODB.Recordset
Dim m_rstVat As ADODB.Recordset
Dim m_rstNext As ADODB.Recordset

Dim blnLvw As Boolean
Dim blnsortA As Boolean

Private m_conSADBEL As ADODB.Connection      'Private datVat As DAO.Database
Private m_conTaric As ADODB.Connection      'Private datClient As DAO.Database

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    '-----> Deletes selcted record
    frm_taricclients.MousePointer = 11
    
    'BeginTrans
    With m_rstClient
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                If ListView1.SelectedItem.Text = ![VAT NUM OR NAME] And _
                    frm_taricmain.txtCode.Text = ![TARIC CODE] Then
                    
                    .Delete
                    
                    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                    
                    frm_taricclients.MousePointer = 0
                    'CommitTrans

                    If ListView1.ListItems.Count = 0 Then
                        cmdDelete.Enabled = False
                    End If
                    
                    Exit Sub
                End If
                .MoveNext
            Loop
            
            ExecuteNonQuery m_conTaric, "DELETE * FROM CLIENTS WHERE [VAT NUM OR NAME] = '" & ListView1.SelectedItem.Text & "' AND [TARIC CODE] = '" & frm_taricmain.txtCode.Text & "' "
        End If
    End With
    
    frm_taricclients.MousePointer = 0
End Sub



Private Sub Form_Load()

    '-----> To convert captions to default language
    Call LoadResStrings(Me, True)
    
    strLangOfDesc = frm_taricmain.strLangOfDesc
    
    '-----> add Client Taric code to label
    Dim strLeft As String
    Dim strRight As String
    
    strLeft = Left(lblCode.Caption, InStr(lblCode.Caption, "<") - 1)
    strRight = Right(lblCode.Caption, Len(lblCode.Caption) - InStr(lblCode.Caption, ">"))
    lblCode.Caption = strLeft & " " & frm_taricmain.txtCode.Text & " " & strRight
    
    ListView1.ColumnHeaders(1).Text = Translate(865)
    ListView1.ColumnHeaders(2).Text = Translate(292)
    
    '<<< dandan 112306
    '<<< Update with database password
    'Set datVat = OpenDatabase(cAppPath & "\mdb_sadbel.mdb")
    ADOConnectDB m_conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    'OpenDAODatabase datVat, cAppPath, "mdb_sadbel.mdb"
                            
    'Set m_conTARIC = OpenDatabase(cAppPath & "\mdb_taric.mdb")
    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTARIC, cAppPath, "mdb_taric.mdb"
    
    '-----> Set to mdb_taric to clients table
    ADORecordsetOpen "select * from CLIENTS", m_conTaric, m_rstClient, adOpenKeyset, adLockOptimistic
    'Set m_rstClient = m_conTARIC.OpenRecordset("select * from CLIENTS")
    
    '-----> Set recordset to a table depending on the used language. to mdb_sadbel, picklist maintenance
    If strLangOfDesc = "Dutch" Then
        ADORecordsetOpen "select * from [PICKLIST MAINTENANCE DUTCH]", m_conSADBEL, m_rstVat, adOpenKeyset, adLockOptimistic
        'Set m_rstVat = m_conSADBEL.OpenRecordset("select * from [PICKLIST MAINTENANCE DUTCH]")
    ElseIf strLangOfDesc = "French" Then
        ADORecordsetOpen "select * from [PICKLIST MAINTENANCE FRENCH]", m_conSADBEL, m_rstVat, adOpenKeyset, adLockOptimistic
        'Set m_rstVat = m_conSADBEL.OpenRecordset("select * from [PICKLIST MAINTENANCE FRENCH]")
    End If
    
    
    '----->Load listview1 with data
    LoadLvw
    
    frm_taricmain.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ADORecordsetClose m_rstVat
    ADORecordsetClose m_rstClient
    ADORecordsetClose m_rstNext
    
    ADODisconnectDB m_conSADBEL
    ADODisconnectDB m_conTaric
    
    UnloadControls Me

End Sub

Private Sub ListView1_Click()

    If blnLvw Then
        blnLvw = False
    Else:
        cmdDelete.Enabled = False
    End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo NoItem:
    ListView1.SortKey = ColumnHeader.Index - 1
        
    If blnsortA Then
        ListView1.SortOrder = lvwDescending
        blnsortA = False
    Else
        ListView1.SortOrder = lvwAscending
        blnsortA = True
    End If
    
    ListView1.Sorted = True
    ListView1.ListItems(1).Selected = True
    
NoItem:
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Item.Index <> 0 Then
        cmdDelete.Enabled = True
        blnLvw = True
    End If

End Sub

Private Sub LoadLvw()
    Dim strSQL As String
    '-----> Vat number from mdb_taric Clients
    '-----> Description from mdb_sadbel Picklist 'Language' defenition
    
    With m_rstClient
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                If frm_taricmain.txtCode.Text = ![TARIC CODE] Then
                    m_rstVat.MoveFirst
                    Do While Not m_rstVat.EOF
                       If m_rstVat![Internal Code] = 7.60856091976166E+19 And ![VAT NUM OR NAME] = m_rstVat![code] Then
                            ListView1.ListItems.Add , , ![VAT NUM OR NAME]
                            If strLangOfDesc = "Dutch" Then
                                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , m_rstVat![DESCRIPTION DUTCH]
                                
                            ElseIf strLangOfDesc = "French" Then
                                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , m_rstVat![DESCRIPTION FRENCH]
                            End If
                            GoTo Labas:
                       End If
                       m_rstVat.MoveNext
                    Loop
                    
                    If strLangOfDesc = "Dutch" Then
                        ADORecordsetOpen "select * from [PICKLIST MAINTENANCE FRENCH]", m_conSADBEL, m_rstNext, adOpenKeyset, adLockOptimistic
                        'Set m_rstNext = m_conSADBEL.OpenRecordset("select * from [PICKLIST MAINTENANCE FRENCH]")
                    ElseIf strLangOfDesc = "French" Then
                        ADORecordsetOpen "select * from [PICKLIST MAINTENANCE DUTCH]", m_conSADBEL, m_rstNext, adOpenKeyset, adLockOptimistic
                        'Set m_rstNext = m_conSADBEL.OpenRecordset("select * from [PICKLIST MAINTENANCE DUTCH]")
                    End If
                    
                    If Not (m_rstNext.EOF And m_rstNext.BOF) Then
                        m_rstNext.MoveFirst
                    
                        Do While Not m_rstNext.EOF
                           If m_rstNext![Internal Code] = 7.60856091976166E+19 And ![VAT NUM OR NAME] = m_rstNext![code] Then
                                ListView1.ListItems.Add , , ![VAT NUM OR NAME]
                                If strLangOfDesc = "Dutch" Then
                                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , m_rstNext![DESCRIPTION DUTCH]
                                ElseIf strLangOfDesc = "French" Then
                                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , m_rstNext![DESCRIPTION FRENCH]
                                End If
                                GoTo Labas:
                           End If
                           m_rstNext.MoveNext
                        Loop
                    End If
                    
                    ADORecordsetClose m_rstNext
                    ADORecordsetOpen "select * from [PICKLIST MAINTENANCE ENGLISH]", m_conSADBEL, m_rstNext, adOpenKeyset, adLockOptimistic
                    'Set m_rstNext = m_conSADBEL.OpenRecordset("select * from [PICKLIST MAINTENANCE ENGLISH]")
                    
                    If Not (m_rstNext.EOF And m_rstNext.BOF) Then
                        m_rstNext.MoveFirst
                    
                        Do While Not m_rstNext.EOF
                           If m_rstNext![Internal Code] = 7.60856091976166E+19 And ![VAT NUM OR NAME] = m_rstNext![code] Then
                                ListView1.ListItems.Add , , ![VAT NUM OR NAME]
                                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , m_rstNext![DESCRIPTION ENGLISH]
                                GoTo Labas:
                           End If
                           m_rstNext.MoveNext
                        Loop
                    End If
                    
                    ListView1.ListItems.Add , , ![VAT NUM OR NAME]
Labas:
               
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub


