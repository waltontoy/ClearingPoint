VERSION 5.00
Begin VB.Form frmGroupBy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group By"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboExpandCollapse 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5520
      Width           =   4575
   End
   Begin VB.CheckBox chkAutoGroup 
      Caption         =   "Automatically group according to arrangement"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame fraGrouping 
      Caption         =   "Group items by"
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      Begin VB.CheckBox chkShowField 
         Caption         =   "Show field in view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   650
         Width           =   1575
      End
      Begin VB.ComboBox cboFields 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   290
         Width           =   2415
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraGrouping 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
      Begin VB.CheckBox chkShowField 
         Caption         =   "Show field in view"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   650
         Width           =   1575
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   290
         Width           =   2415
      End
   End
   Begin VB.Frame fraGrouping 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   4215
      Begin VB.CheckBox chkShowField 
         Caption         =   "Show field in view"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   650
         Width           =   1575
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   290
         Width           =   2415
      End
   End
   Begin VB.Frame fraGrouping 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   4080
      Width           =   4215
      Begin VB.CheckBox chkShowField 
         Caption         =   "Show field in view"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   650
         Width           =   1575
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   290
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   4920
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4920
      TabIndex        =   23
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   4920
      TabIndex        =   24
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblExpandCollapse 
      Caption         =   "Expand/collapse defaults:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   2055
   End
End
Attribute VB_Name = "frmGroupBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private conGroups As ADODB.Connection
Private clsGroupings As CGrid
Private arrColumns
Private arrGroupings


Private Sub cboFields_Click(Index As Integer)
    
    Dim lngIndex As Long
    
    '>> enable/disable controls
    If cboFields(Index).ListIndex > 0 Then
        chkShowField(Index).Enabled = True
        optAscending(Index).Enabled = True
        optDescending(Index).Enabled = True
        
        If Index < 4 Then
            fraGrouping(Index + 1).Enabled = True
            cboFields(Index + 1).Enabled = True
'            chkShowField(Index + 1).Enabled = True
'            optAscending(Index + 1).Enabled = True
'            optDescending(Index + 1).Enabled = True
        End If
    Else
        chkShowField(Index).Enabled = False
        optAscending(Index).Enabled = False
        optDescending(Index).Enabled = False
        
        lngIndex = Index
        Do While lngIndex < 4
            fraGrouping(lngIndex + 1).Enabled = False
            cboFields(lngIndex + 1).Enabled = False
            chkShowField(lngIndex + 1).Enabled = False
            optAscending(lngIndex + 1).Enabled = False
            optDescending(lngIndex + 1).Enabled = False
            lngIndex = lngIndex + 1
        Loop

    End If
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdClear_Click()

    Dim lngCtr As Long
    
    '>> clear all groupings
    For lngCtr = 1 To 4
        cboFields(lngCtr).ListIndex = 0
    Next
    
    chkAutoGroup.Value = 0
    cboExpandCollapse.ListIndex = 0
    
End Sub

Private Sub cmdOK_Click()
    
    '>> save all grouping to db
    If EntriesAreValid = True Then
        If chkAutoGroup.Value = 0 Then
            Call SaveGroupings
        Else
            clsGroupings.GroupHeaders = ""
        End If
        clsGroupings.AutoGroup = chkAutoGroup.Value
        clsGroupings.ExpandCollapseDefault = cboExpandCollapse.ListIndex

        If frmEditView.Visible = False Then
            clsGroupings.UpdateGridSetting conGroups
            clsGroupings.DataChanged = True
        End If

        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    '>> load all combobox fields
    Call PopCombo
    
End Sub

Private Sub chkAutoGroup_Click()
    
    Dim lngCtr As Long
    
    If chkAutoGroup.Value = 1 Then
        '>> disable grouping controls if autogroup was enabled
        For lngCtr = 1 To 4
            fraGrouping(lngCtr).Enabled = False
            cboFields(lngCtr).Enabled = False
            chkShowField(lngCtr).Enabled = False
            optAscending(lngCtr).Enabled = False
            optDescending(lngCtr).Enabled = False
        Next
    Else
        '>> enable first grouping if autogroup was duisabled
        For lngCtr = 1 To 4
            cboFields(lngCtr).ListIndex = 0
            chkShowField(lngCtr).Value = False
            optAscending(lngCtr).Value = True
            optDescending(lngCtr).Value = False
        Next
        
        fraGrouping(1).Enabled = True
        cboFields(1).Enabled = True
    End If
    
    
End Sub

Private Sub PopCombo()

    Dim rstFields As ADODB.Recordset
    Dim lngCtr As Long
    Dim lngArrayCtr As Long
    Dim strWhere As String
    Dim strCommandText As String
    
    '>> load current groupings defined for the selected view
    
    arrColumns = Split(clsGroupings.DVCIDs, "*****")
    arrGroupings = Split(clsGroupings.GroupHeaders, "*****")
    
    For lngCtr = 0 To UBound(arrColumns)
        strWhere = strWhere & "DVC_ID = " & arrColumns(lngCtr) & " OR "
    Next
    
    strWhere = Mid(strWhere, 1, Len(strWhere) - 3)
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_ID, "
    strCommandText = strCommandText & "DVC_FieldAlias "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE " & strWhere
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & "DVC_FieldAlias "
    
    Set rstFields = conGroups.Execute(strCommandText)
    
    lngCtr = 0
    
    cboFields(1).AddItem "(none)"
    cboFields(2).AddItem "(none)"
    cboFields(3).AddItem "(none)"
    cboFields(4).AddItem "(none)"

    Do While Not rstFields.EOF
        lngCtr = lngCtr + 1
        cboFields(1).AddItem rstFields!DVC_FieldAlias
        cboFields(2).AddItem rstFields!DVC_FieldAlias
        cboFields(3).AddItem rstFields!DVC_FieldAlias
        cboFields(4).AddItem rstFields!DVC_FieldAlias
        
        rstFields.MoveNext
    Loop
    
    cboFields(1).ListIndex = 0
    cboFields(2).ListIndex = 0
    cboFields(3).ListIndex = 0
    cboFields(4).ListIndex = 0
    
    lngArrayCtr = 0
    
    For lngCtr = 1 To 4
        rstFields.MoveFirst
        If lngArrayCtr < UBound(arrGroupings) Then
            Do While Not rstFields.EOF
                If rstFields!DVC_ID = Val(arrGroupings(lngArrayCtr)) Then
                    cboFields(lngCtr).Text = rstFields!DVC_FieldAlias
                    If Val(arrGroupings(lngArrayCtr + 1)) = 1 Then
                        optAscending(lngCtr).Value = True
                        optDescending(lngCtr).Value = False
                    Else
                        optAscending(lngCtr).Value = False
                        optDescending(lngCtr).Value = True
                    End If
                    If Val(arrGroupings(lngArrayCtr + 2)) = 1 Then
                        chkShowField(lngCtr).Value = 1
                    Else
                        chkShowField(lngCtr).Value = 0
                    End If
                    Exit Do
                End If
                rstFields.MoveNext
            Loop
            lngArrayCtr = lngArrayCtr + 3
        Else
            Exit For
        End If
    Next
    
    'Set rstFields = Nothing
    ' by hobbes 10/18/2005
    Call ADORecordsetClose(rstFields)
    
    chkAutoGroup.Value = IIf(clsGroupings.AutoGroup = True, 1, 0)
    
    cboExpandCollapse.AddItem "As last viewed"
    cboExpandCollapse.AddItem "All expanded"
    cboExpandCollapse.AddItem "All collapsed"
    
    cboExpandCollapse.ListIndex = clsGroupings.ExpandCollapseDefault
    
End Sub

Private Function EntriesAreValid() As Boolean
    
    Dim lngCtr As Long
    Dim lngCtr2 As Long
    
    '>> check if groupings are valid
    For lngCtr = 1 To 4
        For lngCtr2 = (lngCtr + 1) To 4
            If cboFields(lngCtr).Text = cboFields(lngCtr2) And cboFields(lngCtr).ListIndex > 0 Then
                MsgBox "You cannot group items by the field '" & cboFields(lngCtr).Text & _
                        "' more than once.", vbInformation + vbOKOnly, "Cubepoint Library"
                cboFields(lngCtr2).SetFocus
                EntriesAreValid = False
                Exit Function
            End If
        Next
    Next
    
    EntriesAreValid = True
    
End Function

Private Sub SaveGroupings()
    
    Dim lngCtr As Long
    Dim lngDVCID As Long
    Dim lngIDPos As Long
    Dim lngAliasLength As Long
    Dim lngStart As Long
    Dim lngSortOrder As Long
    Dim blnSortWasFound As Boolean
    
    '>> save groupings defined
    
    clsGroupings.GroupHeaders = ""
    
    For lngCtr = 1 To 4
        If cboFields(lngCtr).ListIndex > 0 Then
            clsGroupings.GetDVCID lngDVCID, cboFields(lngCtr).Text, conGroups
            
            clsGroupings.GroupHeaders = clsGroupings.GroupHeaders & lngDVCID & _
                                        "*****" & IIf(optAscending(lngCtr).Value = True, "1", "-1") & _
                                        "*****" & IIf(chkShowField(lngCtr).Value = 1, "1", "-1") & _
                                        "*****"
                                        
            lngAliasLength = Len(cboFields(lngCtr).Text)
            lngIDPos = InStr(1, clsGroupings.Sort, cboFields(lngCtr).Text & "*****")
            
                           
            If lngIDPos > 0 Then
                lngSortOrder = Val(Mid(clsGroupings.Sort, lngIDPos + lngAliasLength + 5, 1))
                If lngSortOrder = jgexSortAscending Then
                    clsGroupings.Sort = Mid(clsGroupings.Sort, 1, lngIDPos + lngAliasLength + 4) & _
                                        IIf(optAscending(lngCtr).Value = True, "1", "-1") & _
                                        Mid(clsGroupings.Sort, lngIDPos + lngAliasLength + 6)
                Else
                    clsGroupings.Sort = Mid(clsGroupings.Sort, 1, lngIDPos + lngAliasLength + 4) & _
                                        IIf(optAscending(lngCtr).Value = True, "1", "-1") & _
                                        Mid(clsGroupings.Sort, lngIDPos + lngAliasLength + 7)
                End If
            End If
        Else
            Exit For
        End If
    Next
    
    If Trim(clsGroupings.GroupHeaders) <> "" Then
        clsGroupings.GroupHeaders = Mid(clsGroupings.GroupHeaders, 1, Len(clsGroupings.GroupHeaders) - 5)
    End If
    
End Sub


Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set clsGroupings = GridProps
    Set conGroups = ADOConnection
    
    Set Me.Icon = Window.Icon

    Me.Show vbModal
    
    Set GridProps = clsGroupings
    Set ADOConnection = conGroups
    
    Set conGroups = Nothing
    Set clsGroupings = Nothing
    
End Sub


