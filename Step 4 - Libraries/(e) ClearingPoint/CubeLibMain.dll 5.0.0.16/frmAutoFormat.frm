VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAutoFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Formatting"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   3360
      TabIndex        =   18
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4680
      TabIndex        =   19
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties of selected rule"
      Height          =   4455
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   5775
      Begin VB.Frame fraFont 
         Caption         =   "Font:"
         Height          =   2535
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2655
         Begin VB.TextBox txtPreview 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "frmAutoFormat.frx":0000
            Top             =   1800
            Width           =   2265
         End
         Begin VB.TextBox txtColor 
            BackColor       =   &H80000012&
            Enabled         =   0   'False
            Height          =   350
            Left            =   1560
            TabIndex        =   28
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "&Color"
            Height          =   350
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkStrikethru 
            Caption         =   "Strikethru"
            Height          =   255
            Left            =   1560
            TabIndex        =   9
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkItalic 
            Caption         =   "Italic"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkUnderline 
            Caption         =   "Underline"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Preview:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Column:"
         Height          =   2535
         Left            =   2880
         TabIndex        =   32
         Top             =   720
         Width           =   2775
         Begin VB.TextBox txtColumnText 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            Height          =   350
            Left            =   1560
            TabIndex        =   13
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   2415
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   840
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Icon Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Text Value: "
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Condition:"
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   5535
         Begin VB.TextBox txtValue2 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4440
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboOperator 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtValue 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboField 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblValue2 
            Caption         =   "Value 2:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4440
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblField 
            Caption         =   "Field:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblOperator 
            Caption         =   "Operator:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblValue 
            Caption         =   "Value 1:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3360
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1670
         TabIndex        =   5
         Top             =   360
         Width           =   3850
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Do&wn"
      Height          =   350
      Left            =   4800
      TabIndex        =   4
      Top             =   1635
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move &Up"
      Height          =   350
      Left            =   4800
      TabIndex        =   3
      Top             =   1215
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   350
      Left            =   4800
      TabIndex        =   2
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwFormat 
      Height          =   1770
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3122
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgFormat 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Rules for this view:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmAutoFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAutoFormat As CGrid
Private m_conAutoFormat As ADODB.Connection
Private m_objMain As Object
Private m_rstFormatOffline As ADODB.Recordset
Private m_lngItemCtr As Long
Private m_blnLoadOnly As Boolean
Private m_arrFields()

Private Sub cboField_Click()
    
    '>> enable operator combo box
    cboOperator.Enabled = True
    cboOperator.BackColor = vbWhite
    lblOperator.Enabled = True
    
    If (UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "DOUBLE" Or _
        UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "LONG" Or _
        UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "INTEGER" Or _
        UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "NUMBER" Or _
        UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "BOOLEAN" Or _
        UCase(Trim(m_arrFields((cboField.ListIndex * 4) + 3))) = "DATE") Then
            If cboOperator.ListCount > 8 Then
                cboOperator.RemoveItem 9 '>> contains
                cboOperator.RemoveItem 8 '>> not contains
            End If
    ElseIf cboOperator.ListCount < 9 Then
        cboOperator.AddItem "contains"
        cboOperator.AddItem "not contains"
    End If

    If Not m_blnLoadOnly Then
        '>> Update offline record if new field was selected
        Call UpdateOfflineRecord
    End If
    
End Sub

Private Sub cboOperator_Click()
    
    '>> enable 'value' text box
    txtValue.Enabled = True
    txtValue.BackColor = vbWhite
    lblValue.Enabled = True
    
    
    If (cboOperator.ListIndex = 6 Or cboOperator.ListIndex = 7) Then
        '>> if selected operator is equal to 'between' or 'not between', enable value2 text box
        txtValue2.Enabled = True
        txtValue2.BackColor = vbWhite
        lblValue2.Enabled = True
    Else
        '>> disable value 2 text box
        txtValue2.Enabled = False
        txtValue2.BackColor = vbButtonFace
        lblValue2.Enabled = False
    End If
    
    If (cboOperator.ListIndex <> 0) Then
        cboType.ListIndex = 0
    End If
    
    If (m_blnLoadOnly = False) Then
        '>> update offline record if new operator was selected
        Call UpdateOfflineRecord
    End If
    
End Sub

Private Sub cboType_Click()

    Select Case cboType.ListIndex
        Case 0
            txtColumnText.Enabled = False
            txtColumnText.BackColor = vbButtonFace
            cmdBrowse.Enabled = False
            chkBold.Enabled = True
            chkItalic.Enabled = True
            chkStrikethru.Enabled = True
            chkUnderline.Enabled = True
            cmdColor.Enabled = True
            txtColor.Enabled = True
            txtPreview.Enabled = True
            cboOperator.Enabled = True
            
        Case 1
            txtColumnText.Enabled = True
            txtColumnText.BackColor = vbWhite
            cmdBrowse.Enabled = False
            cboOperator.ListIndex = 0
            chkBold.Enabled = False
            chkItalic.Enabled = False
            chkStrikethru.Enabled = False
            chkUnderline.Enabled = False
            cmdColor.Enabled = False
            txtColor.Enabled = False
            txtPreview.Enabled = False
            cboOperator.Enabled = False
            
        Case 2
            txtColumnText.Enabled = False
            txtColumnText.BackColor = vbButtonFace
            cmdBrowse.Enabled = True
            cboOperator.ListIndex = 0
            chkBold.Enabled = False
            chkItalic.Enabled = False
            chkStrikethru.Enabled = False
            chkUnderline.Enabled = False
            cmdColor.Enabled = False
            txtColor.Enabled = False
            txtPreview.Enabled = False
            cboOperator.Enabled = False
            
        Case 3
            txtColumnText.Enabled = True
            txtColumnText.BackColor = vbWhite
            cmdBrowse.Enabled = True
            cboOperator.ListIndex = 0
            chkBold.Enabled = False
            chkItalic.Enabled = False
            chkStrikethru.Enabled = False
            chkUnderline.Enabled = False
            cmdColor.Enabled = False
            txtColor.Enabled = False
            txtPreview.Enabled = False
            cboOperator.Enabled = False
    End Select
    
End Sub

Private Sub cboType_Validate(Cancel As Boolean)
        
    Call UpdateOfflineRecord
    
End Sub

Private Sub chkBold_Click()

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
    txtPreview.FontBold = CBool(chkBold.Value)
    
End Sub

Private Sub chkItalic_Click()

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
    txtPreview.FontItalic = CBool(chkItalic.Value)
    
End Sub

Private Sub chkStrikethru_Click()

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
    txtPreview.FontStrikethru = CBool(chkStrikethru.Value)
    
End Sub

Private Sub chkUnderline_Click()

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
    txtPreview.FontUnderline = CBool(chkUnderline.Value)
    
End Sub

Private Sub cmdAdd_Click()

    m_lngItemCtr = m_lngItemCtr + 1
    
    '>> load default settings
    
    m_rstFormatOffline.AddNew
    m_rstFormatOffline!ID = "Temp" & m_lngItemCtr
    m_rstFormatOffline!Name = "Untitled"
    m_rstFormatOffline!Field = ""
    m_rstFormatOffline!Operator = 0
    m_rstFormatOffline!Value1 = ""
    m_rstFormatOffline!Value2 = ""
    m_rstFormatOffline!Bold = False
    m_rstFormatOffline!Italic = False
    m_rstFormatOffline!Underline = False
    m_rstFormatOffline!Strikethru = False
    m_rstFormatOffline!ForeColor = vbBlack
    m_rstFormatOffline!ColumnType = 0
    m_rstFormatOffline!ColumnText = ""
    m_rstFormatOffline!Selected = True
    m_rstFormatOffline.Update
    
    lvwFormat.ListItems.Add lvwFormat.ListItems.Count + 1, "Temp" & m_lngItemCtr, "Untitled"
    lvwFormat.ListItems(lvwFormat.ListItems.Count).Selected = True
    lvwFormat.SelectedItem.Checked = True
    lvwFormat.SelectedItem.EnsureVisible
    lvwFormat.Refresh
    
    cmdDelete.Enabled = True
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    txtName.Enabled = True
    chkBold.Enabled = True
    chkUnderline.Enabled = True
    chkItalic.Enabled = True
    chkStrikethru.Enabled = True
    txtColor.Enabled = True
    cmdColor.Enabled = True
    txtPreview.Enabled = True
    cboType.Enabled = True
    txtColumnText.Enabled = True
    cmdBrowse.Enabled = False
    cboField.Enabled = True
    cboOperator.Enabled = True
    txtValue.Enabled = True
    txtValue2.Enabled = True
    
    Call LoadItemSettings
    
    txtName.SetFocus
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdColor_Click()

    '>> load color dialog box
    cdgFormat.CancelError = True
    
    On Error Resume Next
    cdgFormat.ShowColor
    
    If Err.Number <> cdlCancel Then
        txtColor.BackColor = cdgFormat.Color
        txtPreview.ForeColor = cdgFormat.Color
    End If
    cdgFormat.CancelError = False
    Err.Clear
    
    Call UpdateOfflineRecord
    

End Sub

Private Sub cmdDelete_Click()

    '>> delete selected auto format
    If Not lvwFormat.SelectedItem Is Nothing Then
        If m_rstFormatOffline.RecordCount > 0 Then
            m_rstFormatOffline.MoveFirst
            m_rstFormatOffline.Find "ID = '" & lvwFormat.SelectedItem.Key & "'", , adSearchForward, 0
            If Not m_rstFormatOffline.EOF Then
                If m_rstFormatOffline!Default = True Then
                    MsgBox "Selected record is a default format. Uncheck the record if you want to disable this condition.", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
                If InStr(1, m_rstFormatOffline!ID, "Temp") > 0 Then
                    '>> delete record in the offline recordset
                    m_rstFormatOffline.Delete
                    m_rstFormatOffline.Update
                Else
                    '>> do not delete yet, just put a tag in the ID field
                    m_rstFormatOffline!ID = Trim(Replace(m_rstFormatOffline!ID, "ID", "DEL"))
                End If
                m_rstFormatOffline.Update
            End If
        End If
        
        lvwFormat.ListItems.Remove lvwFormat.SelectedItem.Index
        
    End If
    
    If lvwFormat.ListItems.Count = 0 Then
        cmdDelete.Enabled = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        txtName.Enabled = False
        chkBold.Enabled = False
        chkUnderline.Enabled = False
        chkItalic.Enabled = False
        chkStrikethru.Enabled = False
        txtColor.Enabled = False
        cmdColor.Enabled = False
        txtPreview.Enabled = False
        cboType.Enabled = False
        txtColumnText.Enabled = False
        cmdBrowse.Enabled = False
        cboField.Enabled = False
        cboOperator.Enabled = False
        txtValue.Enabled = False
        txtValue2.Enabled = False
    Else
        lvwFormat.SelectedItem = lvwFormat.ListItems(1)
        Call lvwFormat_ItemCheck(lvwFormat.SelectedItem)
    End If
    
End Sub

Private Sub cmdDown_Click()

    Dim strName As String
    Dim strKey As String
    Dim strKey2 As String
    
    '>> lower the priority level of the selected item
    If lvwFormat.SelectedItem.Index < lvwFormat.ListItems.Count Then
        strName = lvwFormat.SelectedItem.Text
        strKey = lvwFormat.SelectedItem.Key
        strKey2 = lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Key
        
        lvwFormat.SelectedItem.Key = ""
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Key = ""
        
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index).Text = lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Text
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index).Key = strKey2
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Text = strName
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Key = strKey
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index + 1).Selected = True
    End If

End Sub

Private Sub cmdbrowse_Click()
    
    cdgFormat.Filter = "Picture Files(*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
    
    cdgFormat.CancelError = True
    On Error Resume Next
    cdgFormat.ShowOpen
    
    If Err.Number <> cdlCancel Then
        Err.Clear
        If Trim(cdgFormat.FileName) <> "" Then
            imgIcon.Picture = LoadPicture(cdgFormat.FileName, , vbLPColor)
        Else
            imgIcon.Picture = Nothing
        End If
    
        Call UpdateOfflineRecord
    Else
        cdgFormat.CancelError = False
        Err.Clear
    End If
    
End Sub

Private Sub cmdOK_Click()
    
    Dim lngCtr As Long
    
    '>> check if items are valid
    If m_rstFormatOffline.RecordCount > 0 Then
        m_rstFormatOffline.MoveFirst
        Do While Not m_rstFormatOffline.EOF
            If Trim(m_rstFormatOffline!Field) = "" And InStr(1, m_rstFormatOffline!ID, "DEL") = 0 Then
                MsgBox "Please enter a valid condition.", vbInformation, "Cubepoint Library"
                lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                cboField.SetFocus
                Exit Sub
            End If
            If m_rstFormatOffline!Operator < 0 And InStr(1, m_rstFormatOffline!ID, "DEL") = 0 Then
                MsgBox "Please enter a valid condition.", vbInformation, "Cubepoint Library"
                lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                cboOperator.SetFocus
                Exit Sub
            End If
            If Trim(m_rstFormatOffline!Value1) = "" And InStr(1, m_rstFormatOffline!ID, "DEL") = 0 Then
                MsgBox "Please enter a valid condition.", vbInformation, "Cubepoint Library"
                lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                txtValue.SetFocus
                Exit Sub
            End If
            If (m_rstFormatOffline!Operator = 6 Or m_rstFormatOffline!Operator = 7) And Trim(m_rstFormatOffline!Value2) = "" And InStr(1, m_rstFormatOffline!ID, "DEL") = 0 Then
                MsgBox "Please enter a valid condition.", vbInformation, "Cubepoint Library"
                lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                If txtValue2.Enabled = True Then
                    txtValue2.SetFocus
                End If
                Exit Sub
            End If
            
            For lngCtr = 1 To lvwFormat.ListItems.Count
                If UCase(Trim(m_rstFormatOffline!ID)) <> UCase(lvwFormat.ListItems(lngCtr).Key) Then
                    If UCase(Trim(m_rstFormatOffline!Name)) = UCase(lvwFormat.ListItems(lngCtr).Text) Then
                        MsgBox "Duplicate name was found. Please enter a unique 'Rule' name.", vbInformation, "Cubepoint Library"
                        lvwFormat.ListItems(lngCtr).Selected = True
                        txtName.SetFocus
                        Exit Sub
                    End If
                End If
            Next
            
            Select Case m_rstFormatOffline!ColumnType
                Case 0
                    
                Case 1
                    If Trim(m_rstFormatOffline!ColumnText) = "" Then
                        MsgBox "Please enter a valid text value.", vbInformation, "Cubepoint Library"
                        lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                        Call lvwFormat_ItemClick(lvwFormat.SelectedItem)
                        If txtColumnText.Enabled = True Then
                            txtColumnText.SetFocus
                        End If
                        Exit Sub
                    End If
                            
                Case 2
                    If Trim(m_rstFormatOffline!Icon) = "" And InStr(1, m_rstFormatOffline!ID, "DEL") = 0 Then
                        MsgBox "Please select an icon for rule '" & Trim(m_rstFormatOffline!Name) & "'.", vbInformation, "Cubepoint Library"
                        lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                        Call lvwFormat_ItemClick(lvwFormat.SelectedItem)
                        If cmdBrowse.Enabled = True Then
                            cmdBrowse.SetFocus
                        End If
                        Exit Sub
                    End If
                Case 3
                    If Trim(m_rstFormatOffline!ColumnText) = "" Then
                        MsgBox "Please enter a valid text value.", vbInformation, "Cubepoint Library"
                        lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                        Call lvwFormat_ItemClick(lvwFormat.SelectedItem)
                        If txtColumnText.Enabled = True Then
                            txtColumnText.SetFocus
                        End If
                        Exit Sub
                    End If
                    If Trim(m_rstFormatOffline!Icon) = "" Then
                        MsgBox "Please select an icon for rule '" & Trim(m_rstFormatOffline!Name) & "'.", vbInformation, "Cubepoint Library"
                        lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Selected = True
                        Call lvwFormat_ItemClick(lvwFormat.SelectedItem)
                        If cmdBrowse.Enabled = True Then
                            cmdBrowse.SetFocus
                        End If
                        Exit Sub
                    End If
            End Select
            
            m_rstFormatOffline.MoveNext
        Loop
    End If
    
    '>> save autoformat items
    Call SaveFormats

    If frmEditView.Visible = False Then
        m_clsAutoFormat.DataChanged = True
    End If

    Unload Me
    
End Sub

Private Sub cmdUp_Click()

    Dim strName As String
    Dim strKey As String
    Dim strKey2 As String
    
    '>> increase priority level of selected autoformat item
    If lvwFormat.SelectedItem.Index > 1 Then
        strName = lvwFormat.SelectedItem.Text
        strKey = lvwFormat.SelectedItem.Key
        strKey2 = lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Key
        
        lvwFormat.SelectedItem.Key = ""
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Key = ""
        
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index).Text = lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Text
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index).Key = strKey2
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Text = strName
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Key = strKey
        lvwFormat.ListItems(lvwFormat.SelectedItem.Index - 1).Selected = True
        
    End If
    

End Sub

Private Sub Form_Load()
    
    m_lngItemCtr = 0
    
    '>> create offline recordset for autoformat items
    CreateOffline
    
    '>> load settings to offline recordset
    LoadFormats
    
    '>> load combobox fields
    LoadFields
    
    If (lvwFormat.ListItems.Count > 0) Then
        
        lvwFormat.ListItems(1).Selected = True
        
        '>> load autoformat settings of the first item in the
        LoadItemSettings
        
    Else
        
        cmdDelete.Enabled = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        txtName.Enabled = False
        chkBold.Enabled = False
        chkUnderline.Enabled = False
        chkItalic.Enabled = False
        chkStrikethru.Enabled = False
        txtColor.Enabled = False
        cmdColor.Enabled = False
        txtPreview.Enabled = False
        cboType.Enabled = False
        txtColumnText.Enabled = False
        cmdBrowse.Enabled = False
        cboField.Enabled = False
        cboOperator.Enabled = False
        txtValue.Enabled = False
        txtValue2.Enabled = False
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If (m_rstFormatOffline Is Nothing = False) Then
        If (m_rstFormatOffline.RecordCount > 0) Then
            m_rstFormatOffline.MoveFirst
            
            Do While m_rstFormatOffline.EOF = False
                If (Len(Dir(Trim(FNullField(m_rstFormatOffline![Icon])))) > 0) Then
                    On Error Resume Next
                    Kill Trim(m_rstFormatOffline!Icon)
                    On Error GoTo 0
                End If
                
                m_rstFormatOffline.MoveNext
            Loop
        End If
    End If
    
    Set m_rstFormatOffline = Nothing
    
End Sub

Private Sub lvwFormat_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    '>> select checked item
    lvwFormat.SelectedItem = Item
    '>> load autoformat settings of selected item
    Call LoadItemSettings
    '>> update autoformat offline recordset
    Call UpdateOfflineRecord
    
End Sub

Private Sub lvwFormat_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    '>> load autoformat settings of selected item
    Call LoadItemSettings
    
End Sub

Private Sub txtColumnText_Validate(Cancel As Boolean)

    Call UpdateOfflineRecord
    
End Sub

Private Sub txtName_GotFocus()

    SendKeysEx "{Home}+{End}"
    
End Sub

Private Sub CreateOffline()

    '>> create offline recordset for autoformat records
    
    Set m_rstFormatOffline = New ADODB.Recordset
    m_rstFormatOffline.CursorLocation = adUseClient

    m_rstFormatOffline.Fields.Append "Name", adChar, 50, adFldIsNullable
    m_rstFormatOffline.Fields.Append "ID", adChar, 10, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Field", adChar, 50, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Operator", adDouble
    m_rstFormatOffline.Fields.Append "Value1", adChar, 50, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Value2", adChar, 50, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Bold", adBoolean
    m_rstFormatOffline.Fields.Append "Italic", adBoolean
    m_rstFormatOffline.Fields.Append "Strikethru", adBoolean
    m_rstFormatOffline.Fields.Append "Underline", adBoolean
    m_rstFormatOffline.Fields.Append "ForeColor", adDouble
    m_rstFormatOffline.Fields.Append "ColumnType", adDouble
    m_rstFormatOffline.Fields.Append "ColumnText", adChar, 50, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Icon", adChar, 255, adFldIsNullable
    m_rstFormatOffline.Fields.Append "Selected", adBoolean
    m_rstFormatOffline.Fields.Append "Default", adBoolean
    
    m_rstFormatOffline.Open
    
End Sub

Private Sub LoadFormats()
    Dim rstFormat As ADODB.Recordset
    Dim strCommandText As String
    Dim lngCtr As Long
    
    Dim strLicLogo As String
    Dim lngImgHandle As Long
    Dim bytImage() As Byte
    
    
    '>> load autoformat settings of the current view
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "UVCFormatCondition "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "UVC_ID = " & m_clsAutoFormat.UVC_ID & " "
        strCommandText = strCommandText & "AND "
        strCommandText = strCommandText & "Node_ID = " & m_clsAutoFormat.NodeID & " "
        strCommandText = strCommandText & "ORDER BY "
        strCommandText = strCommandText & "FC_Priority, FC_ID "
    
    ADORecordsetOpen strCommandText, m_conAutoFormat, rstFormat, adOpenKeyset, adLockOptimistic
    'RstOpen strCommandText, m_conAutoFormat, rstFormat, adOpenKeyset, adLockOptimistic, , True
    
    If (rstFormat.RecordCount > 0) Then
        rstFormat.MoveFirst
        
        For lngCtr = 1 To rstFormat.RecordCount
            With m_rstFormatOffline
                .AddNew
                
                ![ID] = "ID" & rstFormat![FC_ID]
                ![Name] = FNullField(rstFormat![FC_Name])
                ![Field] = FNullField(rstFormat![FC_Field])
                ![Operator] = rstFormat![FC_Operator]
                ![Value1] = FNullField(rstFormat![FC_Value1])
                ![Value2] = FNullField(rstFormat![FC_Value2])
                ![Bold] = rstFormat![FC_FontBold]
                ![Italic] = rstFormat![FC_FontItalic]
                ![Underline] = rstFormat![FC_FontUnderline]
                ![Strikethru] = rstFormat![FC_FontStrikeThru]
                ![ForeColor] = rstFormat![FC_ForeColor]
                ![Selected] = rstFormat![FC_Selected]
                ![Default] = rstFormat![FC_Default]
                ![ColumnType] = rstFormat![FC_ColumnType]
                ![ColumnText] = FNullField(rstFormat![FC_ColumnText])
                
                If (![ColumnType] > 1) Then
                    ![Icon] = PicPath(rstFormat![FC_ID], rstFormat)
                End If
                
                .Update
            End With
            
            lvwFormat.ListItems.Add lvwFormat.ListItems.Count + 1, "ID" & CStr(rstFormat![FC_ID]), rstFormat![FC_Name]
            lvwFormat.ListItems(lvwFormat.ListItems.Count).Checked = CStr(rstFormat![FC_Selected])
            
            rstFormat.MoveNext
        Next lngCtr
    End If
    
    ADORecordsetClose rstFormat
    
End Sub

Private Sub LoadFields()
    Dim rstFields As ADODB.Recordset
    Dim lngCtr As Long
    Dim strCommandText As String
    
    
    '>> load combo box items
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "DVC_ID, "
        strCommandText = strCommandText & "DVC_FieldSource, "
        strCommandText = strCommandText & "DVC_FieldAlias, "
        strCommandText = strCommandText & "DVC_DataType "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "DefaultViewColumns "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "Tview_ID = " & m_clsAutoFormat.TView_ID & " "
        strCommandText = strCommandText & "ORDER BY "
        strCommandText = strCommandText & "DVC_FieldAlias "
    
    ADORecordsetOpen strCommandText, m_conAutoFormat, rstFields, adOpenKeyset, adLockOptimistic
    'Set rstFields = m_conAutoFormat.Execute(strCommandText)
    
    lngCtr = 0
    
    Do While Not rstFields.EOF
        ReDim Preserve m_arrFields(lngCtr + 3)
        
        m_arrFields(lngCtr) = rstFields![DVC_ID]
        m_arrFields(lngCtr + 1) = rstFields![DVC_FieldSource]
        m_arrFields(lngCtr + 2) = rstFields![DVC_FieldAlias]
        m_arrFields(lngCtr + 3) = rstFields![DVC_DataType]
        
        lngCtr = lngCtr + 4
        
        cboField.AddItem FNullField(rstFields![DVC_FieldAlias])
    
        rstFields.MoveNext
    Loop
    
    
    cboOperator.AddItem "equal"
    cboOperator.AddItem "not equal"
    cboOperator.AddItem "greater than"
    cboOperator.AddItem "less than"
    cboOperator.AddItem "greater than or equal to"
    cboOperator.AddItem "less than or equal to"
    cboOperator.AddItem "between"
    cboOperator.AddItem "not between"
    cboOperator.AddItem "contains"
    cboOperator.AddItem "not contains"
    
    cboType.AddItem "Default"
    cboType.AddItem "Text"
    cboType.AddItem "Icon"
    cboType.AddItem "Text and Icon"
    
    cboType.ListIndex = 0
    
End Sub

Private Sub LoadItemSettings()

    '>> load settings of the selected autoformat
    If (lvwFormat.SelectedItem Is Nothing = False) Then
        
        If (m_rstFormatOffline.RecordCount > 0) Then
            
            m_rstFormatOffline.MoveFirst
            m_rstFormatOffline.Find "ID = '" & lvwFormat.SelectedItem.Key & "'", , adSearchForward, 0
            
            If (m_rstFormatOffline.EOF = False) Then
                txtName.Text = Trim(m_rstFormatOffline![Name])
                
                m_blnLoadOnly = True
                chkBold.Value = IIf(m_rstFormatOffline![Bold], vbChecked, vbUnchecked)
                chkUnderline.Value = IIf(m_rstFormatOffline![Underline], vbChecked, vbUnchecked)
                chkItalic.Value = IIf(m_rstFormatOffline![Italic], vbChecked, vbUnchecked)
                chkStrikethru.Value = IIf(m_rstFormatOffline![Strikethru], vbChecked, vbUnchecked)
                m_blnLoadOnly = False
                
                txtPreview.FontBold = m_rstFormatOffline![Bold]
                txtPreview.FontItalic = m_rstFormatOffline![Italic]
                txtPreview.FontUnderline = m_rstFormatOffline![Underline]
                txtPreview.FontStrikethru = m_rstFormatOffline![Strikethru]
                txtPreview.ForeColor = m_rstFormatOffline![ForeColor]
                txtColor.BackColor = m_rstFormatOffline![ForeColor]
                
                m_blnLoadOnly = True
                cboType.ListIndex = m_rstFormatOffline![ColumnType]
                m_blnLoadOnly = False
                
                
                Select Case m_rstFormatOffline!ColumnType
                    Case 1
                        cboType.ListIndex = 1
                        txtColumnText.Text = Trim(m_rstFormatOffline!ColumnText)
                        imgIcon.Picture = Nothing
                        
                    Case 2
                        cboType.ListIndex = 2
                        txtColumnText.Text = ""
                        imgIcon.Picture = LoadPicture(Trim(m_rstFormatOffline!Icon), , vbLPColor)
                        cdgFormat.FileName = Trim(m_rstFormatOffline!Icon)
                        
                    Case 3
                        cboType.ListIndex = 3
                        txtColumnText.Text = Trim(m_rstFormatOffline!ColumnText)
                        imgIcon.Picture = LoadPicture(Trim(m_rstFormatOffline!Icon), , vbLPColor)
                        cdgFormat.FileName = Trim(m_rstFormatOffline!Icon)
                        
                    Case Else
                        txtColumnText.Text = ""
                        imgIcon.Picture = Nothing
                        
                End Select
                
                
                m_blnLoadOnly = True
                If Trim(m_rstFormatOffline!Field) = "" Then
                    cboField.ListIndex = 0
                Else
                    cboField.Text = Trim(m_rstFormatOffline!Field)
                End If
                cboOperator.ListIndex = m_rstFormatOffline!Operator
                m_blnLoadOnly = False
                
                txtValue.Text = Trim(m_rstFormatOffline!Value1)
                txtValue2.Text = Trim(m_rstFormatOffline!Value2)
            End If
        End If
    End If

End Sub
Private Sub UpdateOfflineRecord()

    Dim lngImgHandle As Long
    Dim lngPos As Long
    Dim strTempFile As String
    Dim bytImage() As Byte
    
    '>> update offline recordset to save changes made in the settings
    If m_rstFormatOffline.RecordCount > 0 And Not (lvwFormat.SelectedItem Is Nothing) Then
        m_rstFormatOffline.MoveFirst
        m_rstFormatOffline.Find "ID = '" & lvwFormat.SelectedItem.Key & "'", , adSearchForward, 0
        
        If Not m_rstFormatOffline.EOF Then
            m_rstFormatOffline!Name = txtName.Text
            m_rstFormatOffline!Field = cboField.Text
            m_rstFormatOffline!Operator = cboOperator.ListIndex
            m_rstFormatOffline!Value1 = Trim(txtValue.Text)
            m_rstFormatOffline!Value2 = Trim(txtValue2.Text)
            m_rstFormatOffline!Bold = CBool(chkBold.Value)
            m_rstFormatOffline!Italic = CBool(chkItalic.Value)
            m_rstFormatOffline!Underline = CBool(chkUnderline.Value)
            m_rstFormatOffline!Strikethru = CBool(chkStrikethru.Value)
            m_rstFormatOffline!ForeColor = txtColor.BackColor
            m_rstFormatOffline!Selected = lvwFormat.SelectedItem.Checked
            m_rstFormatOffline!ColumnType = cboType.ListIndex
            Select Case m_rstFormatOffline!ColumnType
                Case 1
                    m_rstFormatOffline!ColumnText = txtColumnText.Text
                    m_rstFormatOffline!Icon = ""
                Case 2
                    m_rstFormatOffline!ColumnText = ""
                    m_rstFormatOffline!Icon = cdgFormat.FileName
                Case 3
                    m_rstFormatOffline!ColumnText = txtColumnText.Text
                    m_rstFormatOffline!Icon = cdgFormat.FileName
                Case Else
                    m_rstFormatOffline!ColumnText = ""
                    m_rstFormatOffline!Icon = ""
            End Select
            m_rstFormatOffline.Update
        End If
    End If
    
End Sub

Private Sub SaveFormats()
    
    '>> save new rules and update db for modifications
    If m_rstFormatOffline.RecordCount > 0 Then
        m_rstFormatOffline.MoveFirst
        Do While Not m_rstFormatOffline.EOF
            If InStr(1, m_rstFormatOffline!ID, "Temp") > 0 Then
                '>>  add new rule
                Call AppendRecord(m_rstFormatOffline)
                If m_rstFormatOffline!ColumnType > 1 Then
                    Call SaveIcon(m_rstFormatOffline)
                End If
            ElseIf InStr(1, m_rstFormatOffline!ID, "ID") > 0 Then
                '>> update rule
                Call UpdateRecord(m_rstFormatOffline)
                If m_rstFormatOffline!ColumnType > 1 Then
                    Call SaveIcon(m_rstFormatOffline)
                End If
            ElseIf InStr(1, m_rstFormatOffline!ID, "DEL") > 0 Then
                '>> delete rule
                Call DeleteRecord(m_rstFormatOffline)
            End If
            
            m_rstFormatOffline.MoveNext
        Loop
    End If
    
End Sub

Private Sub AppendRecord(ByRef FormatRecord As ADODB.Recordset)

    Dim strCommandText As String
    
    '>> add new rule
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & "UVC_ID, "
    strCommandText = strCommandText & "FC_Name, "
    strCommandText = strCommandText & "FC_Field, "
    strCommandText = strCommandText & "FC_Operator, "
    strCommandText = strCommandText & "FC_Value1, "
    strCommandText = strCommandText & "FC_Value2, "
    strCommandText = strCommandText & "FC_FontBold, "
    strCommandText = strCommandText & "FC_FontItalic, "
    strCommandText = strCommandText & "FC_FontStrikeThru, "
    strCommandText = strCommandText & "FC_FontUnderline, "
    strCommandText = strCommandText & "FC_ForeColor, "
    strCommandText = strCommandText & "FC_ColumnType, "
    strCommandText = strCommandText & "FC_ColumnText, "
    strCommandText = strCommandText & "FC_Selected, "
    strCommandText = strCommandText & "Node_ID, "
    strCommandText = strCommandText & "FC_Default, "
    strCommandText = strCommandText & "FC_Priority "
    strCommandText = strCommandText & ") "
    strCommandText = strCommandText & "VALUES "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & m_clsAutoFormat.UVC_ID & ", "
    strCommandText = strCommandText & "'" & Trim(FormatRecord!Name) & "', "
    strCommandText = strCommandText & "'" & Trim(FormatRecord!Field) & "', "
    strCommandText = strCommandText & FormatRecord!Operator & ", "
    strCommandText = strCommandText & "'" & Trim(FormatRecord!Value1) & "', "
    strCommandText = strCommandText & "'" & Trim(FormatRecord!Value2) & "', "
    strCommandText = strCommandText & IIf(FormatRecord!Bold, "True", "False") & ", "
    strCommandText = strCommandText & IIf(FormatRecord!Italic, "True", "False") & ", "
    strCommandText = strCommandText & IIf(FormatRecord!Strikethru, "True", "False") & ", "
    strCommandText = strCommandText & IIf(FormatRecord!Underline, "True", "False") & ", "
    strCommandText = strCommandText & FormatRecord!ForeColor & ", "
    strCommandText = strCommandText & FormatRecord!ColumnType & ", "
    strCommandText = strCommandText & "'" & Trim(FormatRecord!ColumnText) & "', "
    strCommandText = strCommandText & IIf(FormatRecord!Selected, "True", "False") & ", "
    strCommandText = strCommandText & m_clsAutoFormat.NodeID & ", "
    strCommandText = strCommandText & 0 & ", "
    strCommandText = strCommandText & lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Index
    strCommandText = strCommandText & ") "
    
    ExecuteNonQuery m_conAutoFormat, strCommandText
    'm_conAutoFormat.Execute strCommandText
    
End Sub

Private Sub UpdateRecord(ByRef FormatRecord As ADODB.Recordset)

    Dim strCommandText As String
    
    '>> update rule
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "FC_Name = '" & Trim(FormatRecord!Name) & "', "
    strCommandText = strCommandText & "FC_Field = '" & Trim(FormatRecord!Field) & "', "
    strCommandText = strCommandText & "FC_Operator = " & FormatRecord!Operator & ", "
    strCommandText = strCommandText & "FC_Value1 = '" & Trim(FormatRecord!Value1) & "', "
    strCommandText = strCommandText & "FC_Value2 = '" & Trim(FormatRecord!Value2) & "', "
    strCommandText = strCommandText & "FC_FontBold = " & IIf(FormatRecord!Bold, "True", "False") & ", "
    strCommandText = strCommandText & "FC_FontItalic = " & IIf(FormatRecord!Italic, "True", "False") & ", "
    strCommandText = strCommandText & "FC_FontStrikeThru = " & IIf(FormatRecord!Strikethru, "True", "False") & ", "
    strCommandText = strCommandText & "FC_FontUnderline = " & IIf(FormatRecord!Underline, "True", "False") & ", "
    strCommandText = strCommandText & "FC_ForeColor = " & FormatRecord!ForeColor & ", "
    strCommandText = strCommandText & "FC_ColumnType = " & FormatRecord!ColumnType & ", "
    strCommandText = strCommandText & "FC_ColumnText = '" & Trim(FormatRecord!ColumnText) & "', "
    strCommandText = strCommandText & "FC_Selected = " & IIf(FormatRecord!Selected, "True", "False") & ", "
    strCommandText = strCommandText & "FC_Priority = " & lvwFormat.ListItems(Trim(m_rstFormatOffline!ID)).Index & " "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "FC_ID = " & Val(Replace(FormatRecord!ID, "ID", ""))
    
    ExecuteNonQuery m_conAutoFormat, strCommandText
    'm_conAutoFormat.Execute strCommandText
    
End Sub

Private Sub DeleteRecord(ByRef FormatRecord As ADODB.Recordset)
    
    Dim strCommandText As String
    
    '>> delete rule
    strCommandText = vbNullString
    strCommandText = strCommandText & "DELETE "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "FC_ID = " & Val(Mid(FormatRecord!ID, 4))
    
    ExecuteNonQuery m_conAutoFormat, strCommandText
    'm_conAutoFormat.Execute strCommandText
    
End Sub

Private Sub SaveIcon(OfflineRecordset As ADODB.Recordset)

    Dim rstIcons As ADODB.Recordset
    Dim lngImgHandle As Long
    Dim strCommandText As String
    Dim strTempFile As String
    Dim bytImage() As Byte
    
    '>> save the icon of the current rule
    If OfflineRecordset.EOF = False Then
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "FC_Icon "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "UVCFormatCondition "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "UVC_ID = " & m_clsAutoFormat.UVC_ID & " "
        strCommandText = strCommandText & "AND "
        strCommandText = strCommandText & "FC_Name = '" & Trim(OfflineRecordset!Name) & "'"
           
        ADORecordsetOpen strCommandText, m_conAutoFormat, rstIcons, adOpenKeyset, adLockOptimistic
        'Call RstOpen(strCommandText, m_conAutoFormat, rstIcons, adOpenKeyset, adLockOptimistic)
        
        If rstIcons.RecordCount > 0 Then
        
            
            strTempFile = WindowsTempPath & "Icon.img"
            
            If Len(Dir$(strTempFile)) Then
                Kill strTempFile
            End If
            
            SavePicture LoadPicture(Trim(OfflineRecordset!Icon), , vbLPColor), strTempFile
                            
            lngImgHandle = FreeFile()
            Open strTempFile For Binary Access Read As #lngImgHandle
            ReDim bytImage(LOF(lngImgHandle))
            Get #lngImgHandle, , bytImage()
            Close #lngImgHandle
            
            rstIcons.Fields("FC_Icon").AppendChunk bytImage()
            
            On Error Resume Next
            Kill Trim(strTempFile)
            On Error GoTo 0
            
            rstIcons.Update
            
            UpdateRecordset m_conAutoFormat, rstIcons, "UVCFormatCondition"
        End If
            
    End If
    
    
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    
    '>> update listview after modifying the name of the rule
    If Not lvwFormat.SelectedItem Is Nothing Then
        lvwFormat.SelectedItem.Text = txtName.Text
    End If

    Call UpdateOfflineRecord
    
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
End Sub

Private Sub txtValue2_Validate(Cancel As Boolean)

    If Not m_blnLoadOnly Then
        Call UpdateOfflineRecord
    End If
    
End Sub

Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set m_clsAutoFormat = GridProps
    Set m_conAutoFormat = ADOConnection
    Set m_objMain = Window
    
    Set Me.Icon = Window.Icon
    
    Me.Show vbModal
    
    Set GridProps = m_clsAutoFormat
    Set ADOConnection = m_conAutoFormat
    Set Window = m_objMain
    
    Set m_objMain = Nothing
    Set m_clsAutoFormat = Nothing
    Set m_conAutoFormat = Nothing
    
    Unload Me
    
End Sub
