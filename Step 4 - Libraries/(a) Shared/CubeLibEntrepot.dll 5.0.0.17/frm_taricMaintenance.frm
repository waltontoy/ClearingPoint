VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_taricmaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TARIC - Maintenance"
   ClientHeight    =   7185
   ClientLeft      =   2685
   ClientTop       =   1545
   ClientWidth     =   9990
   Icon            =   "frm_taricMaintenance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "858"
   Begin MSComctlLib.ListView lvwMaintenanceFilter 
      Height          =   975
      Left            =   8160
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
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
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   10
      Tag             =   "179"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "Select"
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   9
      Tag             =   "178"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox picMaintenance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   240
      ScaleHeight     =   6015
      ScaleWidth      =   9495
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   9495
      Begin MSComctlLib.ListView lvwMaintenance 
         Height          =   5055
         Left            =   105
         TabIndex        =   0
         Top             =   840
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
      Begin VB.CommandButton cmdMaintenance 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   7920
         TabIndex        =   8
         Tag             =   "356"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdMaintenance 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   7920
         TabIndex        =   7
         Tag             =   "355"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMaintenance 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   7920
         TabIndex        =   6
         Tag             =   "260"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdMaintenance 
         Caption         =   "&Add"
         Height          =   375
         Index           =   1
         Left            =   7920
         TabIndex        =   5
         Tag             =   "836"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtTaricCode 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtKeyword 
         Height          =   375
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   3120
         MaxLength       =   78
         TabIndex        =   3
         Top             =   480
         Width           =   4695
      End
      Begin VB.CommandButton cmdMaintenance 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   7920
         TabIndex        =   4
         Tag             =   "827"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblTaricCodeList 
         Caption         =   "The list below shows all TARIC codes currently available in SADBELplus."
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "863"
         Top             =   120
         Width           =   9255
      End
   End
   Begin MSComctlLib.TabStrip tabTaricCodes 
      Height          =   6495
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "861"
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11456
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TARIC Codes"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_taricmaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private m_conTaric As ADODB.Connection    'Private datTaric As DAO.Database

Private blnWasInvokedFromCode As Boolean
Private blnFindWasClicked As Boolean
Private mblnClientLowerLeftChanged As Boolean
Private mintClientLowerLeft As Integer
Private msngColumnHeaderWidths() As Single
Private msngMousePointerX As Single
Private msngMousePointerY As Single
Private aSelected(1 To 2) As String

Private Enum CommandButtonIndexConstants
    sbpFind = 0
    sbpAdd
    sbpDelete
    sbpCopy
    sbpModify
End Enum

Private Sub cmdMaintenance_Click(Index As Integer)
    Dim itmFound As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intLastItemIndex As Integer
    
    Dim strPattern As String
    
    Dim strSelectedCode As String
    Dim intResponse As Integer
    
    Select Case Index
        Case sbpFind
            With lvwMaintenance.ListItems
                intLastItemIndex = .Count
                lvwMaintenanceFilter.ListItems.Clear
                
                If intLastItemIndex Then
                    Screen.MousePointer = vbHourglass
                    
                    For intItemIndex = 1 To intLastItemIndex
                        Set itmFound = lvwMaintenanceFilter.ListItems.Add(, , .Item(intItemIndex).Text)
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                    Next
                    
                    .Clear
                    
                    With lvwMaintenanceFilter.ListItems
                        For intItemIndex = 1 To intLastItemIndex
                            strPattern = "*" & Trim(txtDescription.Text) & "*"
                            If .Item(intItemIndex).ListSubItems(2) Like strPattern Then
                                Set itmFound = lvwMaintenance.ListItems.Add(, , .Item(intItemIndex).Text)
                                itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                                itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                            End If
                        Next
                    End With
                    
                    mblnClientLowerLeftChanged = True
                    
                    If .Count Then
                        .Item(1).Selected = True
                        cmdMaintenance(sbpDelete).Enabled = True
                        cmdMaintenance(sbpCopy).Enabled = True
                        cmdMaintenance(sbpModify).Enabled = True
                    Else
                        cmdMaintenance(sbpDelete).Enabled = False
                        cmdMaintenance(sbpCopy).Enabled = False
                        cmdMaintenance(sbpModify).Enabled = False
                    End If
                    
                    cmdMaintenance(sbpFind).Enabled = False
                    
                    Set itmFound = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End With
            
            blnFindWasClicked = True
        Case sbpAdd       ' Blank settings
            gstrTaricMainCallType = "AddBlank/" & Me.Name
            frm_taricmain.Show vbModal, Me
            
            ' If blnOKWasPressed Then
            blnWasInvokedFromCode = True
            With lvwMaintenance
                If Not .SelectedItem Is Nothing Then
                    txtTaricCode.Text = .SelectedItem.Text
                    txtKeyword.Text = .SelectedItem.ListSubItems(1).Text
                    txtDescription.Text = .SelectedItem.ListSubItems(2).Text
                End If
            End With
            blnWasInvokedFromCode = False
            
            cmdMaintenance(sbpDelete).Enabled = True
            cmdMaintenance(sbpCopy).Enabled = True
            cmdMaintenance(sbpModify).Enabled = True
            ' End If
        Case sbpDelete
            With lvwMaintenance
                strSelectedCode = .SelectedItem.Text
                intItemIndex = .SelectedItem.Index
                intResponse = MsgBox(GetNewStr(358, Array(strSelectedCode, strSelectedCode)), vbYesNo + vbCritical + vbApplicationModal, Me.Caption)
                
                If intResponse = vbYes Then
                    .ListItems.Remove intItemIndex
                    
                    If .ListItems.Count Then
                        .SelectedItem.Selected = True    ' Highlights preceding listitem.
                    Else
                        cmdMaintenance(sbpDelete).Enabled = False
                        cmdMaintenance(sbpCopy).Enabled = False
                        cmdMaintenance(sbpModify).Enabled = False
                    End If
                    
                    With m_conTaric
                        ExecuteNonQuery m_conTaric, "DELETE * FROM [COMMON] WHERE [TARIC CODE] = " & Chr(39) & ProcessQuotes(strSelectedCode) & Chr(39)
                        ExecuteNonQuery m_conTaric, "DELETE * FROM [IMPORT] WHERE [TARIC CODE] = " & Chr(39) & ProcessQuotes(strSelectedCode) & Chr(39)
                        ExecuteNonQuery m_conTaric, "DELETE * FROM [EXPORT] WHERE [TARIC CODE] = " & Chr(39) & ProcessQuotes(strSelectedCode) & Chr(39)
                        ExecuteNonQuery m_conTaric, "DELETE * FROM [CLIENTS] WHERE [TARIC CODE] = " & Chr(39) & ProcessQuotes(strSelectedCode) & Chr(39)
                        
                        '.Execute "DELETE * FROM COMMON WHERE [TARIC CODE] = " & Chr(39) & processquotes(strSelectedCode) & Chr(39)
                        '.Execute "DELETE * FROM IMPORT WHERE [TARIC CODE] = " & Chr(39) & processquotes(strSelectedCode) & Chr(39)
                        '.Execute "DELETE * FROM EXPORT WHERE [TARIC CODE] = " & Chr(39) & processquotes(strSelectedCode) & Chr(39)
                        '.Execute "DELETE * FROM CLIENTS WHERE [TARIC CODE] = " & Chr(39) & processquotes(strSelectedCode) & Chr(39)
                    End With
                End If
            End With
        Case sbpCopy      ' Same settings, blank code
            gstrTaricMainCallType = "Copy/" & Me.Name
            frm_taricmain.Show vbModal, Me
        Case sbpModify    ' Same settings
            gstrTaricMainCallType = "Modify/" & Me.Name
            frm_taricmain.Show vbModal, Me
    End Select
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)
    Select Case Index
        Case 0
            gblnFormWasCanceled = False
            'Added by BCo 2006-05-08
            'Prevents crashing when listview is empty
            If Not lvwMaintenance.SelectedItem Is Nothing Then
                aSelected(1) = lvwMaintenance.SelectedItem
                aSelected(2) = lvwMaintenance.SelectedItem.SubItems(2)
            End If
        Case 1
            gblnFormWasCanceled = True
    End Select
    
    Unload Me
End Sub

Public Sub My_Load(CallingForm As Object, Optional Taric_Code As String)
    If Len(Taric_Code) > 0 Then
        'Added by BCo 2006-05-08
        'Prevents crashing when listview is empty
        If Not lvwMaintenance.SelectedItem Is Nothing Then
            txtTaricCode.Text = Taric_Code
            txtKeyword = lvwMaintenance.SelectedItem.SubItems(1)
            txtDescription = lvwMaintenance.SelectedItem.SubItems(2)
        End If
    End If
    Me.Show vbModal, CallingForm
    If gblnFormWasCanceled = False Then
        CallingForm.txtTaricCode.Text = aSelected(1)
        CallingForm.txtDescription.Text = aSelected(2)
    End If
End Sub

Private Sub Form_Load()
    Dim itmListItem As MSComctlLib.ListItem
    
    Dim strLangOfDesc As String
    Dim blnSelectedItemExists As Boolean
    
    Dim strSQL As String
    Dim rstTaricCodes As ADODB.Recordset
    
    'cAppPath = GetSetting("ClearingPoint", "Settings", "MdbPath")
    
    Screen.MousePointer = vbHourglass
    
    '<<< dandan 112306
    '<<< Update with database password
    'Set m_conTaric = OpenDatabase(cAppPath & "\mdb_taric.mdb")
    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTaric, cAppPath, "mdb_taric.mdb"
'    BeginTrans
    
    Call LoadResStrings(Me, True)
    
    strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
    
' ********** Initialize lvwMaintenance **********
    With lvwMaintenance.ColumnHeaders
        .Item(1).Text = Translate(861)    ' TARIC Code
        .Item(2).Text = Translate(829)    ' Keyword
        .Item(3).Text = Translate(292)    ' Description
        
        ReDim msngColumnHeaderWidths(.Count)
        
        msngColumnHeaderWidths(1) = .Item(1).Width
        msngColumnHeaderWidths(2) = .Item(2).Width
        msngColumnHeaderWidths(3) = .Item(3).Width
    End With
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "[TARIC CODE], "
        strSQL = strSQL & "[KEY " & strLangOfDesc & "] AS KEYWORD, "
        strSQL = strSQL & "[DESC " & strLangOfDesc & "] AS DESCRIPTION "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "COMMON "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "[TARIC CODE] ASC "
    ADORecordsetOpen strSQL, m_conTaric, rstTaricCodes, adOpenKeyset, adLockOptimistic
    'Set rstTaricCodes = m_conTaric.OpenRecordset(strSQL, dbOpenForwardOnly)
    With rstTaricCodes
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set itmListItem = lvwMaintenance.ListItems.Add(, , ![TARIC CODE])
                itmListItem.ListSubItems.Add , , IIf(IsNull(![Keyword]), "", UCase(![Keyword]))
                itmListItem.ListSubItems.Add , , IIf(IsNull(![Description]), "", ![Description])
                
                .MoveNext
            Loop
        End If
    End With
    
    lvwMaintenance.Sorted = True
    mblnClientLowerLeftChanged = True
    blnSelectedItemExists = Not lvwMaintenance.SelectedItem Is Nothing
    
' ********** Initialize cmdMaintenance **********
    If blnSelectedItemExists Then
        cmdMaintenance(sbpDelete).Enabled = True
        cmdMaintenance(sbpCopy).Enabled = True
        cmdMaintenance(sbpModify).Enabled = True
    End If
    
    Set itmListItem = Nothing
    Set rstTaricCodes = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadControls(Me, True)
    
    ADODisconnectDB m_conTaric
    
    Set frm_taricmaintenance = Nothing
End Sub

Private Sub lvwMaintenance_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwMaintenance
        If Not .SelectedItem Is Nothing Then
            ColumnHeader.Tag = ColumnHeader.Tag Xor lvwDescending    ' Reverses current .SortOrder stored in .Tag
            
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = ColumnHeader.Tag
            .SelectedItem.EnsureVisible
        End If
    End With
End Sub

Private Sub lvwMaintenance_DblClick()
    With lvwMaintenance
        If Not .HitTest(msngMousePointerX, msngMousePointerY) Is Nothing Then
'        If Not .SelectedItem Is Nothing Then
            gstrTaricMainCallType = "Modify/" & Me.Name
            frm_taricmain.Show vbModal, Me
        End If
    End With
End Sub

Private Sub lvwMaintenance_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnWasInvokedFromCode = True
    With Item
        txtTaricCode.Text = .Text
        txtKeyword.Text = .ListSubItems(1).Text
        txtDescription.Text = .ListSubItems(2).Text
    End With
    blnWasInvokedFromCode = False
    
    ' Enable CommandButtons here!
End Sub

Private Sub lvwMaintenance_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngMousePointerX = x
    msngMousePointerY = Y
End Sub

Private Sub lvwMaintenance_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not lvwMaintenance.SelectedItem Is Nothing Then
        lvwMaintenance.SelectedItem.Selected = True
    End If
End Sub

Private Sub txtDescription_Change()
    Dim itmFound As MSComctlLib.ListItem
    
    Dim intItemIndex As Integer
    Dim intLastItemIndex As Integer
    
    If Not blnWasInvokedFromCode Then
        If Len(Trim(txtDescription.Text)) Then
            If Not blnFindWasClicked And lvwMaintenance.ListItems.Count Then
                cmdMaintenance(sbpFind).Enabled = True
            End If
        Else
            cmdMaintenance(sbpFind).Enabled = False
            
            With lvwMaintenanceFilter.ListItems
                intLastItemIndex = .Count
                
                If intLastItemIndex And blnFindWasClicked Then
                    Screen.MousePointer = vbHourglass
                    lvwMaintenance.ListItems.Clear
                    
                    For intItemIndex = 1 To intLastItemIndex
                        Set itmFound = lvwMaintenance.ListItems.Add(, , .Item(intItemIndex).Text)
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(1).Text
                        itmFound.ListSubItems.Add , , .Item(intItemIndex).ListSubItems(2).Text
                    Next
                    
                    .Item(1).Selected = True
                    cmdMaintenance(sbpDelete).Enabled = True
                    cmdMaintenance(sbpCopy).Enabled = True
                    cmdMaintenance(sbpModify).Enabled = True
                    
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
        With lvwMaintenance
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
        With lvwMaintenance
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
