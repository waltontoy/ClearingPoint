VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_tariccn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CN - Picklist"
   ClientHeight    =   5790
   ClientLeft      =   2835
   ClientTop       =   2715
   ClientWidth     =   8070
   Icon            =   "frm_taricCN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Tag             =   "868"
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   100
      TabIndex        =   5
      Top             =   410
      Width           =   7840
      _ExtentX        =   13811
      _ExtentY        =   8493
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
         Object.Tag             =   "820"
         Text            =   "Code"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "292"
         Text            =   "Description"
         Object.Width           =   11113
      EndProperty
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Tag             =   "426"
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Tag             =   "179"
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Tag             =   "827"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5595
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_tariccn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----> For Testing
Dim strLangOfDesc As String
'-----> end testing

Private m_conTaric As ADODB.Connection   'Private datCN As DAO.Database

Dim m_rstCN As ADODB.Recordset

Dim blnClick As Boolean

'-----> For Command Find
Dim intFind As Integer
'Dim blnFind As Boolean

'-----> For sorting (ascend/descend)
Dim blnsortA As Boolean

Private Sub cmdCancel_Click()

    frm_tariccn.MousePointer = 11
    Unload Me

End Sub

Private Sub cmdFind_Click()

    Dim intDescLen As Integer
    Dim Counter As Integer
    Dim x As Integer
    Dim intDel As Integer
    
    frm_tariccn.MousePointer = 11
    
    For Counter = intFind To ListView1.ListItems.Count
        If InStr(LCase(ListView1.ListItems(Counter).ListSubItems(1).Text), LCase(txtDescription.Text)) <> 0 Then
            For x = 20 To 0 Step -1
                If ListView1.ListItems.Count < Counter + x Then GoTo Out:
                ListView1.ListItems(Counter + x).EnsureVisible
            Next x
Out:
            ListView1.ListItems(Counter).Selected = True
            txtDescription.SetFocus
            frm_tariccn.MousePointer = 0
            intFind = Counter + 1
            Exit Sub
        Else
        End If
    Next Counter
    cmdFind.Enabled = False
    frm_tariccn.MousePointer = 0

    frm_tariccn.MousePointer = 0
End Sub


Private Sub cmdSelect_Click()

    frm_tariccn.MousePointer = 11
    
    If gstrTaricCNCallType = "frm_taricpicklist" Then
        frm_taricpicklist.txtTaricCode.Text = ListView1.SelectedItem.Text
    ElseIf gstrTaricCNCallType = "frm_taricmain" Then
        frm_taricmain.txtCode.Text = ListView1.SelectedItem.Text
    Else
        frm_taricmaintenance.txtTaricCode.Text = ListView1.SelectedItem.Text
    End If
    Unload Me

End Sub

Private Sub Form_Load()

    '-----> To convert captions to default language
    Call LoadResStrings(Me, True)
    
    ListView1.ColumnHeaders(2).Text = Translate(292)
    
    '-----> get what language is to be used
    If gstrTaricCNCallType = "frm_taricpicklist" Then
        strLangOfDesc = frm_taricpicklist.strLangOfDesc
    ElseIf gstrTaricCNCallType = "frm_taricmain" Then
        strLangOfDesc = frm_taricmain.strLangOfDesc
    Else
        strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
    End If
       

    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTaric, cAppPath, "mdb_taric.mdb"
                           
    ADORecordsetOpen "select * from [CN]", m_conTaric, m_rstCN, adOpenKeyset, adLockOptimistic
    'Set m_rstCN = m_conTaric.OpenRecordset("select * from [CN]")
    
    '-----> Data from Mdb_TARIC.mdb to listview1
    With m_rstCN
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                ListView1.ListItems.Add , , ![CN code]
                If strLangOfDesc = "Dutch" Then
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ![DESC DUTCH]
                ElseIf strLangOfDesc = "French" Then
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ![DESC FRENCH]
                End If
                .MoveNext
            Loop
        End If
    End With
    
    ADORecordsetClose m_rstCN
    
    ADODisconnectDB m_conTaric
    
    '-----> Input code to text box
    If gstrTaricCNCallType = "frm_taricpicklist" Then
        txtCode.Text = Left(frm_taricpicklist.txtTaricCode.Text, 8)
    ElseIf gstrTaricCNCallType = "frm_taricmain" Then
        txtCode.Text = Left(frm_taricmain.txtCode.Text, 8)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnloadControls Me
    Set frm_tariccn = Nothing
End Sub

Private Sub ListView1_Click()

    If gstrTaricCNCallType = "frm_taricmain" Then
        cmdSelect.Enabled = True
    End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    frm_tariccn.MousePointer = 11
    
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
    frm_tariccn.MousePointer = 0

End Sub

Private Sub ListView1_DblClick()

    If gstrTaricCNCallType = "frm_taricmain" Then

        frm_tariccn.MousePointer = 11
        
        If gstrTaricCNCallType = "frm_taricpicklist" Then
            frm_taricpicklist.txtTaricCode.Text = ListView1.SelectedItem.Text
        ElseIf gstrTaricCNCallType = "frm_taricmain" Then
            frm_taricmain.txtCode.Text = ListView1.SelectedItem.Text
        Else
            frm_taricmaintenance.txtTaricCode.Text = ListView1.SelectedItem.Text
        End If
        
        Unload Me
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    blnClick = True
    '-----> Display to Text Boxes the clicked row on the list view
    txtCode.Text = Item
    txtDescription.Text = ListView1.ListItems(Item.Index).SubItems(1)

End Sub

Private Sub txtCode_Change()

    If Not blnClick Then
        frm_tariccn.MousePointer = 11
        
        '----->Sort Column Code Ascending before proceeding
        ListView1.SortKey = 0
        ListView1.SortOrder = lvwAscending
        ListView1.Sorted = True
        
        Dim x As Integer
        Dim itmCode
        
        Set itmCode = ListView1.FindItem(txtCode.Text, , , lvwPartial)
        If itmCode Is Nothing Then
            frm_tariccn.MousePointer = 0
            Exit Sub
        Else
            itmCode.Selected = True
            For x = 20 To 0 Step -1 '-----> 20 items made visible to ensure that the selected item would be on top
                If ListView1.ListItems.Count < itmCode.Index + x Then
                    GoTo Out:
                End If
                
                ListView1.ListItems(itmCode.Index + x).EnsureVisible
            Next x
        End If
Out:
        
        frm_tariccn.MousePointer = 0
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    blnClick = False
End Sub

Private Sub txtDescription_Change()

    '-----> Set for command find
    intFind = 1
    If Not Trim(txtDescription.Text) = "" Then
        cmdFind.Enabled = True
    End If

End Sub



