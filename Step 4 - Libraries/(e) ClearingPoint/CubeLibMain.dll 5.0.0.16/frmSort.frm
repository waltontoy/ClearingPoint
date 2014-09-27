VERSION 5.00
Begin VB.Form frmSort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   4560
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4560
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   4560
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraSort 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   4215
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   4215
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Then By"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   4215
      Begin VB.ComboBox cboFields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort items by"
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton optDescending 
         Caption         =   "Descending"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cboFields 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conSort As ADODB.Connection
Private clsSort As CGrid
Private arrColumns
Private arrSort

Private Sub cboFields_Click(Index As Integer)
    
    If Index < 4 Then
        fraSort(Index + 1).Enabled = True
        cboFields(Index + 1).Enabled = True
        optAscending(Index + 1).Enabled = True
        optDescending(Index + 1).Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdClear_Click()
    
    Dim lngCtr As Long
    
    '>> clear all sorting
    For lngCtr = 1 To 4
        cboFields(lngCtr).ListIndex = 0
        If lngCtr > 1 Then
            fraSort(lngCtr).Enabled = False
            cboFields(lngCtr).Enabled = False
            optAscending(lngCtr).Value = True
            optAscending(lngCtr).Enabled = False
            optDescending(lngCtr).Value = False
            optDescending(lngCtr).Enabled = False
        End If
    Next
    
    clsSort.Sort = ""
    
End Sub

Private Sub cmdOK_Click()
    
    Dim lngCtr As Long
    Dim lngDVCID As Long
    
    clsSort.Sort = ""
    
    '>> save all sorting defined
    
    For lngCtr = 1 To 4
        If fraSort(lngCtr).Enabled = True And cboFields(lngCtr).ListIndex > 0 Then
            clsSort.Sort = clsSort.Sort & cboFields(lngCtr).Text & "*****"
            clsSort.GetDVCID lngDVCID, cboFields(lngCtr).Text, conSort
            
            If optAscending(lngCtr).Value = True Then
                clsSort.Sort = clsSort.Sort & "1*****"
                clsSort.GroupHeaders = Replace(clsSort.GroupHeaders, lngDVCID & "*****-1", lngDVCID & "*****1")
            Else
                clsSort.Sort = clsSort.Sort & "-1*****"
                clsSort.GroupHeaders = Replace(clsSort.GroupHeaders, lngDVCID & "*****1", lngDVCID & "*****-1")
            End If
        End If
    Next
    
    If Trim(clsSort.Sort) <> "" Then
        clsSort.Sort = Mid(clsSort.Sort, 1, Len(clsSort.Sort) - 5)
    End If

    If frmEditView.Visible = False Then
        clsSort.UpdateGridSetting conSort
        clsSort.DataChanged = True
    End If

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim rstFields As ADODB.Recordset
    Dim lngCtr As Long
    Dim strWhere As String
    Dim strCommandText As String
    
    '>> load defined sorting for the selected view
    arrColumns = Split(clsSort.DVCIDs, "*****")
    arrSort = Split(clsSort.Sort, "*****")
    
    For lngCtr = 0 To UBound(arrColumns)
        strWhere = strWhere & "DVC_ID = " & arrColumns(lngCtr) & " OR "
    Next
    
    strWhere = Mid(strWhere, 1, Len(strWhere) - 3)
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_FieldAlias "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE " & strWhere
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & "DVC_FieldAlias "
    
    ADORecordsetOpen strCommandText, conSort, rstFields, adOpenKeyset, adLockOptimistic
    'Set rstFields = conSort.Execute(strCommandText)
    
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
    
    If UBound(arrSort) >= 1 Then
        cboFields(1).Text = arrSort(0)
        If arrSort(1) = "1" Then
            optAscending(1).Value = True
            optDescending(1).Value = False
        Else
            optAscending(1).Value = False
            optDescending(1).Value = True
        End If
        fraSort(2).Enabled = True
        cboFields(2).Enabled = True
        optAscending(2).Enabled = True
        optDescending(2).Enabled = True
        cboFields(2).ListIndex = 0
    End If
    
    If UBound(arrSort) >= 3 Then
        cboFields(2).Text = arrSort(2)
        If arrSort(3) = "1" Then
            optAscending(2).Value = True
            optDescending(2).Value = False
        Else
            optAscending(2).Value = False
            optDescending(2).Value = True
        End If
        fraSort(3).Enabled = True
        cboFields(3).Enabled = True
        optAscending(3).Enabled = True
        optDescending(3).Enabled = True
        cboFields(3).ListIndex = 0
    End If
    
    If UBound(arrSort) >= 5 Then
        cboFields(3).Text = arrSort(4)
        If arrSort(5) = "1" Then
            optAscending(3).Value = True
            optDescending(3).Value = False
        Else
            optAscending(3).Value = False
            optDescending(1).Value = True
        End If
        fraSort(4).Enabled = True
        cboFields(4).Enabled = True
        optAscending(4).Enabled = True
        optDescending(4).Enabled = True
        cboFields(4).ListIndex = 0
    End If
    
    If UBound(arrSort) >= 7 Then
        cboFields(4).Text = arrSort(6)
        If arrSort(7) = "1" Then
            optAscending(4).Value = True
            optDescending(4).Value = False
        Else
            optAscending(4).Value = False
            optDescending(4).Value = True
        End If
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstFields)
    
End Sub



Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set clsSort = GridProps
    Set conSort = ADOConnection
    
    Set Me.Icon = Window.Icon
    
    Me.Show vbModal
    
    Set GridProps = clsSort
    Set ADOConnection = conSort
    
    Set conSort = Nothing
    Set clsColumns = Nothing
    
End Sub


