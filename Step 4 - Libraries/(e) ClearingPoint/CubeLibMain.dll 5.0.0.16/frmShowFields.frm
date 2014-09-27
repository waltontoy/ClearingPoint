VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Fields"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdUp 
         Caption         =   "Move &Up"
         Height          =   350
         Left            =   4440
         TabIndex        =   4
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Move &Down"
         Height          =   350
         Left            =   5640
         TabIndex        =   5
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add ->"
         Height          =   350
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<- &Remove"
         Height          =   350
         Left            =   3000
         TabIndex        =   2
         Top             =   915
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwNotShown 
         Height          =   3375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
      Begin MSComctlLib.ListView lvwShown 
         Height          =   3375
         Left            =   4200
         TabIndex        =   3
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
      Begin VB.Label lblcaption 
         Caption         =   "Available Fields :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblcaption 
         Caption         =   "Show these fields in this order:"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   7320
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   7320
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmShowFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conShowFields As ADODB.Connection
Private m_rstDescriptions As ADODB.Recordset
Private m_clsColumns As CGrid

Private arrColumns
Private arrAlignments
Private arrWidths

Private Sub cmdAdd_Click()

    '>> add selected field to the list of available fields
    If Not lvwNotShown.SelectedItem Is Nothing Then
        lvwShown.ListItems.Add lvwShown.ListItems.Count + 1, lvwNotShown.SelectedItem.Key, lvwNotShown.SelectedItem.Text
        lvwNotShown.ListItems.Remove lvwNotShown.SelectedItem.Index
    End If
    
    cmdRemove.Enabled = True
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    
    If lvwNotShown.ListItems.Count = 0 Then
        cmdAdd.Enabled = False
    Else
        lvwNotShown.SetFocus
        lvwNotShown.DropHighlight = lvwNotShown.SelectedItem
        
        If lvwShown.SelectedItem.Tag = "2" Then
            cmdRemove.Enabled = False
        Else
            cmdRemove.Enabled = True
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdDown_Click()

    Dim strColumn As String
    Dim strKey As String
    Dim strKey2 As String
    
    '>> move selected column to the next column position
    If lvwShown.SelectedItem.Index < lvwShown.ListItems.Count Then
        strColumn = lvwShown.SelectedItem.Text
        strKey = lvwShown.SelectedItem.Key
    
    
        lvwShown.ListItems(lvwShown.SelectedItem.Index).Text = lvwShown.ListItems(lvwShown.SelectedItem.Index + 1).Text
        strKey2 = lvwShown.ListItems(lvwShown.SelectedItem.Index + 1).Key
        lvwShown.ListItems(lvwShown.SelectedItem.Index + 1).Key = ""
        lvwShown.ListItems(lvwShown.SelectedItem.Index).Key = strKey2
        
        lvwShown.ListItems(lvwShown.SelectedItem.Index + 1).Text = strColumn
        lvwShown.ListItems(lvwShown.SelectedItem.Index + 1).Key = strKey
        
        lvwShown.ListItems(strKey).Selected = True
    End If
    
    lvwShown.SetFocus

    If lvwShown.ListItems.Count > 1 Then
        cmdUp.Enabled = True
    End If
    
    If lvwShown.SelectedItem.Index = lvwShown.ListItems.Count Then
        cmdDown.Enabled = False
    End If

End Sub

Private Sub cmdOK_Click()

    Dim lngCtr As Long
    Dim lngGroupCtr As Long
    Dim lngSortCtr As Long
    Dim lngDVCID As Long
    Dim lngIDCtr As Long
    Dim lngPos As Long
    Dim strGrouping() As String
    Dim strSorting() As String
    
    Dim strField As String
    
    Dim rstDVCFieldAlias As ADODB.Recordset
    
    
    If Trim$(m_clsColumns.Sort) <> "" Or Trim$(m_clsColumns.GroupHeaders) <> "" Then
    
        strSorting = Split(m_clsColumns.Sort, "*****")
        strGrouping = Split(m_clsColumns.GroupHeaders, "*****")
                    
        For lngCtr = 1 To lvwNotShown.ListItems.Count
            lngDVCID = Val(Trim(Mid(lvwNotShown.ListItems(lngCtr).Key, 3)))
            
            ADORecordsetOpen "SELECT DVC_FieldAlias AS FieldAlias FROM DefaultViewColumns WHERE DVC_ID = " & lngDVCID, _
                                conShowFields, rstDVCFieldAlias, adOpenKeyset, adLockOptimistic
            'Call RstOpen("SELECT DVC_FieldAlias AS FieldAlias FROM DefaultViewColumns WHERE DVC_ID = " & lngDVCID, _
                conShowFields, rstDVCFieldAlias, adOpenKeyset, adLockReadOnly, , True)
            
            If rstDVCFieldAlias.RecordCount > 0 Then
                strField = rstDVCFieldAlias!FieldAlias
            End If
                    
            For lngSortCtr = 0 To UBound(strSorting) Step 2
                If UCase$(Trim$(strField)) = UCase$(Trim$(strSorting(lngSortCtr))) Then
                    MsgBox "Cannot remove '" & Trim(lvwNotShown.ListItems(lngCtr).Text) & "'. The records are still sorted on this field.", vbOKOnly + vbInformation, Me.Caption
                    Exit Sub
                End If
            Next lngSortCtr
            
            For lngGroupCtr = 0 To UBound(strGrouping) Step 3
                If strGrouping(lngGroupCtr) = lngDVCID Then
                    MsgBox "Cannot remove '" & Trim(lvwNotShown.ListItems(lngCtr).Text) & "'. The records are still grouped on this field.", vbOKOnly + vbInformation, Me.Caption
                    Exit Sub
                End If
            Next lngGroupCtr
            
        Next
    
    End If
    
    m_clsColumns.DVCIDs = ""
    m_clsColumns.Alignments = ""
    m_clsColumns.Widths = ""
    
    
    '>> save available fields
    For lngCtr = 1 To lvwShown.ListItems.Count
        lngDVCID = Val(Trim(Mid(lvwShown.ListItems(lngCtr).Key, 3)))
        m_clsColumns.DVCIDs = m_clsColumns.DVCIDs & lngDVCID & "*****"
        For lngIDCtr = 0 To UBound(arrColumns)
            If Val(arrColumns(lngIDCtr)) = lngDVCID Then
                m_clsColumns.Alignments = m_clsColumns.Alignments & arrAlignments(lngIDCtr) & "*****"
                m_clsColumns.Widths = m_clsColumns.Widths & arrWidths(lngIDCtr) & "*****"
                Exit For
            End If
        Next
        
        If lngIDCtr > UBound(arrColumns) Then
            m_clsColumns.Alignments = m_clsColumns.Alignments & "LEFT*****"
            m_clsColumns.Widths = m_clsColumns.Widths & "1200*****"
        End If
    Next
    
    m_clsColumns.DVCIDs = Mid(m_clsColumns.DVCIDs, 1, Len(m_clsColumns.DVCIDs) - 5)
    m_clsColumns.Alignments = Mid(m_clsColumns.Alignments, 1, Len(m_clsColumns.Alignments) - 5)
    m_clsColumns.Widths = Mid(m_clsColumns.Widths, 1, Len(m_clsColumns.Widths) - 5)
    
    
    '>> Remove dependencies of columns removed from the list of visible columns
'    For lngCtr = 1 To lvwNotShown.ListItems.Count
'
'        lngDVCID = Val(Trim(Mid(lvwNotShown.ListItems(lngCtr).Key, 3)))
'
'        lngPos = InStr(1, m_clsColumns.Sort, lvwNotShown.ListItems(lngCtr).Text & "*****")
'
'        If lngPos > 0 Then
'            If lngPos = 1 Then
'                m_clsColumns.Sort = Replace(m_clsColumns.Sort, Trim(lvwNotShown.ListItems(lngCtr).Text) & "*****1", "")
'                m_clsColumns.Sort = Replace(m_clsColumns.Sort, Trim(lvwNotShown.ListItems(lngCtr).Text) & "*****-1", "")
'                If right(m_clsColumns.Sort, 5) = "*****" Then
'                    m_clsColumns.Sort = Mid(m_clsColumns.Sort, 1, Len(m_clsColumns.Sort) - 5)
'                End If
'            ElseIf Mid(m_clsColumns.Sort, lngPos - 1, 1) = "*" Then
'                m_clsColumns.Sort = Replace(m_clsColumns.Sort, Trim(lvwNotShown.ListItems(lngCtr).Text) & "*****1", "")
'                m_clsColumns.Sort = Replace(m_clsColumns.Sort, Trim(lvwNotShown.ListItems(lngCtr).Text) & "*****-1", "")
'                If left(m_clsColumns.Sort, 5) = "*****" Then
'                    m_clsColumns.Sort = Mid(m_clsColumns.Sort, 6)
'                End If
'            End If
'        End If
'
'        lngPos = InStr(1, m_clsColumns.GroupHeaders, lngDVCID & "*****")
'
'        If lngPos > 0 Then
'            If lngPos = 1 Then
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****1*****1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****1*****-1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****-1*****1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****-1*****-1", "")
'                If right(m_clsColumns.GroupHeaders, 5) = "*****" Then
'                    m_clsColumns.GroupHeaders = Mid(m_clsColumns.GroupHeaders, 1, Len(m_clsColumns.GroupHeaders) - 5)
'                End If
'            ElseIf Mid(m_clsColumns.GroupHeaders, lngPos - 1, 1) = "*" Then
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****1*****1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****1*****-1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****-1*****1", "")
'                m_clsColumns.GroupHeaders = Replace(m_clsColumns.GroupHeaders, lngDVCID & "*****-1*****-1", "")
'                If left(m_clsColumns.GroupHeaders, 5) = "*****" Then
'                    m_clsColumns.GroupHeaders = Mid(m_clsColumns.GroupHeaders, 6)
'                End If
'            End If
'        End If
'
'    Next
    
    If frmEditView.Visible = False Then
        m_clsColumns.UpdateGridSetting conShowFields
        m_clsColumns.DataChanged = True
    End If
    
    Call ADORecordsetClose(rstDVCFieldAlias)
    
    Unload Me
    
End Sub


Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set m_clsColumns = GridProps
    Set conShowFields = ADOConnection
    
    Set Me.Icon = Window.Icon
    
    Me.Show vbModal
    
    Set GridProps = m_clsColumns
    Set ADOConnection = conShowFields
    
    Set conShowFields = Nothing
    Set m_clsColumns = Nothing
    
End Sub

Private Sub cmdRemove_Click()

    '>> remove selected column from the list of available fields
    If Not lvwShown.SelectedItem Is Nothing Then
        If lvwShown.ListItems.Count > 1 Then
            If Val(lvwShown.SelectedItem.Tag) = 0 Then
                lvwNotShown.ListItems.Add lvwNotShown.ListItems.Count + 1, lvwShown.SelectedItem.Key, lvwShown.SelectedItem.Text
                lvwShown.ListItems.Remove lvwShown.SelectedItem.Index
            End If
        End If
    End If
    
    cmdAdd.Enabled = True
    
    If lvwShown.ListItems.Count = 0 Then
        cmdRemove.Enabled = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    Else
        lvwShown.SetFocus
        lvwShown.DropHighlight = lvwShown.SelectedItem
        
        If lvwShown.SelectedItem.Tag = "2" Then
            cmdRemove.Enabled = False
        Else
            cmdRemove.Enabled = True
        End If
    End If
    
    
End Sub

Private Sub cmdUp_Click()

    Dim strColumn As String
    Dim strKey As String
    Dim strKey2 As String
    
    '>> move selected column to the previous column position
    If lvwShown.SelectedItem.Index > 1 Then
        strColumn = lvwShown.SelectedItem.Text
        strKey = lvwShown.SelectedItem.Key
    
    
        lvwShown.ListItems(lvwShown.SelectedItem.Index).Text = lvwShown.ListItems(lvwShown.SelectedItem.Index - 1).Text
        strKey2 = lvwShown.ListItems(lvwShown.SelectedItem.Index - 1).Key
        lvwShown.ListItems(lvwShown.SelectedItem.Index - 1).Key = ""
        lvwShown.ListItems(lvwShown.SelectedItem.Index).Key = strKey2
        
        lvwShown.ListItems(lvwShown.SelectedItem.Index - 1).Text = strColumn
        lvwShown.ListItems(lvwShown.SelectedItem.Index - 1).Key = strKey
        
        lvwShown.ListItems(strKey).Selected = True
    End If
    
    lvwShown.SetFocus
    
    If lvwShown.ListItems.Count > 1 Then
        cmdDown.Enabled = True
    End If
    If lvwShown.SelectedItem.Index = 1 Then
        cmdUp.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim rstDVC As ADODB.Recordset
    Dim lngColumnCtr As Long
    Dim blnShowField As Boolean
    Dim strCommandText As String
    Dim strDesc As String
    
    
    Set m_rstDescriptions = New ADODB.Recordset
    
    Call m_clsColumns.TriggerBeforeShowFields(m_rstDescriptions)

    arrColumns = Split(m_clsColumns.DVCIDs, "*****")
    arrAlignments = Split(m_clsColumns.Alignments, "*****")
    arrWidths = Split(m_clsColumns.Widths, "*****")
    
    '>> get all columns of the selected view
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_ID, "
    strCommandText = strCommandText & "DVC_FieldSource, "
    strCommandText = strCommandText & "DVC_FieldAlias, "
    strCommandText = strCommandText & "DVC_Requirement, "
    strCommandText = strCommandText & "DVC_Position "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "TView_ID = " & m_clsColumns.TView_ID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "DVC_Requirement <> 1 "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "DVC_Position "
    
    ADORecordsetOpen strCommandText, conShowFields, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = conShowFields.Execute(strCommandText)
    
    For lngColumnCtr = 0 To UBound(arrColumns)
        lvwShown.ListItems.Add , , ""
    Next
    
    '>> add column to listview
    If Not (rstDVC.EOF And rstDVC.BOF) Then
        rstDVC.MoveFirst
        
        Do While Not rstDVC.EOF
            '>> check if column is shown in the grid
            blnShowField = False
            For lngColumnCtr = 0 To UBound(arrColumns)
                If rstDVC!DVC_ID = Val(arrColumns(lngColumnCtr)) Then
                    blnShowField = True
                    Exit For
                End If
            Next
            
            strDesc = GetFieldDescription(FNullField(rstDVC!DVC_FieldAlias))
            
            If blnShowField = True Then
                '>> add column to the list of available fields
                lvwShown.ListItems(lngColumnCtr + 1).Key = "ID" & rstDVC!DVC_ID
                lvwShown.ListItems(lngColumnCtr + 1).Text = FNullField(rstDVC!DVC_FieldAlias) & _
                                                            IIf(Trim(strDesc) = "", "", " - " & LCase(strDesc))
                lvwShown.ListItems(lngColumnCtr + 1).Tag = FNullField(rstDVC!DVC_Requirement)
            Else
                '>> add column to the list of fields that are not shown
                lvwNotShown.ListItems.Add lvwNotShown.ListItems.Count + 1, "ID" & rstDVC!DVC_ID, FNullField(rstDVC!DVC_FieldAlias) & _
                                                                                                IIf(Trim(strDesc) = "", "", " - " & LCase(strDesc))
            End If
            
            rstDVC.MoveNext
        Loop
    End If
    
    If lvwNotShown.ListItems.Count = 0 Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
    
    If lvwShown.ListItems.Count = 0 Then
        cmdRemove.Enabled = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    Else
        cmdRemove.Enabled = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call ADORecordsetClose(m_rstDescriptions)
    
End Sub

Private Sub lvwNotShown_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Set lvwNotShown.DropHighlight = Nothing
    
End Sub

Private Sub lvwShown_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Set lvwShown.DropHighlight = Nothing
    If lvwShown.SelectedItem.Tag = "2" Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
    
    If lvwShown.SelectedItem.Index = 1 Then
        cmdUp.Enabled = False
    Else
        cmdUp.Enabled = True
    End If
    
    If lvwShown.SelectedItem.Index = lvwShown.ListItems.Count Then
        cmdDown.Enabled = False
    Else
        cmdDown.Enabled = True
    End If

End Sub

Private Function GetFieldDescription(FieldAlias As String) As String
    
    GetFieldDescription = ""
    
    If m_rstDescriptions.State = adStateOpen Then
        If m_rstDescriptions.RecordCount > 0 Then
            
            m_rstDescriptions.MoveFirst
            m_rstDescriptions.Find "Code='" & FieldAlias & "'"
            
            If Not m_rstDescriptions.EOF Then
                GetFieldDescription = FNullField(m_rstDescriptions.Fields("Description").Value)
            End If
            
        End If
    End If
    
End Function
