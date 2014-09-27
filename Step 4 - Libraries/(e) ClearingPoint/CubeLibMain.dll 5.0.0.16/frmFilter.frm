VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3720
      TabIndex        =   3
      Top             =   3960
      Width           =   1200
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   5040
      TabIndex        =   4
      Top             =   3960
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
      Width           =   1200
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   3765
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6641
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Contents"
      TabPicture(0)   =   "frmFilter.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Advanced"
      TabPicture(1)   =   "frmFilter.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cmdRemove"
      Tab(1).Control(2)=   "lvwConditions"
      Tab(1).Control(3)=   "Label3"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Define more criteria:"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   16
         Top             =   2160
         Width           =   5655
         Begin VB.ComboBox cboField 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add to List"
            Height          =   315
            Left            =   4320
            TabIndex        =   11
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtValue 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            MaxLength       =   50
            TabIndex        =   10
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox cboOperator 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblValue 
            Caption         =   "Value:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3720
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblOperator 
            Caption         =   "Operator:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblField 
            Caption         =   "Field:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3015
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   5655
         Begin VB.TextBox txtSearchKeys 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox cboFields 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1200
            Width           =   5175
         End
         Begin VB.Label Label1 
            Caption         =   "Search for the word(s):"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "In:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70200
         TabIndex        =   7
         Top             =   1785
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwConditions 
         Height          =   1095
         Left            =   -74760
         TabIndex        =   6
         Top             =   660
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Find items that match these criteria:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   420
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsRegistry As CRegistry

Private conFilter As ADODB.Connection
Private clsFilter As CGrid
Private arrFields()

Private Sub cboField_Click()
    
    Dim lngCtr As Long
    Dim strDataType As String
    
    If cboField.Text <> "" Then
        '>> enable operator combo box
        cboOperator.Enabled = True
        cboOperator.BackColor = &H80000005
        lblOperator.Enabled = True
        lblValue.Enabled = True
        txtValue.Enabled = True
        txtValue.BackColor = &H80000005
    Else
        '>> disable operator combo box
        lblOperator.Enabled = False
        cboOperator.Enabled = False
        cboOperator.BackColor = &H8000000F
        lblValue.Enabled = False
        txtValue.Enabled = False
        txtValue.BackColor = &H8000000F
    End If
    
    '>> get field's data type
    For lngCtr = 0 To UBound(arrFields) Step 4
        If arrFields(lngCtr + 2) = cboField.Text Then
            strDataType = arrFields(lngCtr + 3)
            Exit For
        End If
    Next
    
    cboOperator.Clear
    
    '>> load opertator combo box based on the data type of the selected field
    Select Case Trim(UCase(strDataType))
        Case "TEXT"
            cboOperator.AddItem "contains"
            cboOperator.AddItem "is (exactly)"
            cboOperator.AddItem "doesn't contain"
            cboOperator.AddItem "is empty"
            cboOperator.AddItem "is not empty"
        
        Case "NUMBER", "INTEGER", "LONG", "DOUBLE"
            cboOperator.AddItem "contains"
            cboOperator.AddItem "is equal"
            cboOperator.AddItem "less than or equal"
            cboOperator.AddItem "greater than or equal"
            cboOperator.AddItem "between"
        
        Case "DATE"
            'cboOperator.AddItem "contains"
            cboOperator.AddItem "on"
            cboOperator.AddItem "on or before"
            cboOperator.AddItem "on or after"
            cboOperator.AddItem "between"
        
        Case "BOOLEAN"
            cboOperator.AddItem "is true"
            cboOperator.AddItem "is false"
            
    End Select
    
    On Error Resume Next
    cboOperator.ListIndex = 0
    On Error GoTo 0
    
End Sub

Private Sub cboOperator_Click()

    If cboOperator.Text <> "is empty" And cboOperator.Text <> "is not empty" Then
        '>> enable value text box
        lblValue.Enabled = True
        txtValue.Enabled = True
        txtValue.BackColor = &H80000005
    Else
        '>> disable value text box
        lblValue.Enabled = False
        txtValue.Enabled = False
        txtValue.BackColor = &H8000000F
    End If
    
End Sub

Private Sub cmdAdd_Click()
    
    Dim strValue1 As String
    Dim strValue2 As String
    
    '>> add new filter string to listview
    If IsValidCondition(strValue1, strValue2) Then
        lvwConditions.ListItems.Add lvwConditions.ListItems.Count + 1, , Trim(cboField.Text)
        lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID" & CStr(cboOperator.ListIndex + 1), Trim(cboOperator.Text)
        
        If Trim(strValue1) <> "" And Trim(strValue2) <> "" Then
            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 2, , "<" & Trim(strValue1) & "> and <" & Trim(strValue2) & ">"
        Else
            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 2, , Trim(txtValue.Text)
        End If
        
        lvwConditions.ListItems(lvwConditions.ListItems.Count).Selected = True
        lvwConditions.ListItems(lvwConditions.ListItems.Count).EnsureVisible
        lvwConditions.Refresh
        cmdRemove.Enabled = True
    End If
End Sub

Private Function IsValidCondition(ByRef Value1 As String, ByRef Value2 As String) As Boolean

    Dim lngCtr As Long
    Dim arrValues
    
    '>> check if new filter string is valid
    If cboField.ListIndex = -1 Then
        MsgBox "Please enter a valid field.", vbInformation, "Cubepoint Library"
        cboField.SetFocus
        IsValidCondition = False
        Exit Function
    End If
    
    For lngCtr = 0 To UBound(arrFields) Step 4
        If cboField.Text = arrFields(lngCtr + 2) Then
            Select Case Trim(UCase(arrFields(lngCtr + 3)))
                Case "TEXT"
                    If InStr(1, cboOperator.Text, "empty") = 0 And Trim(txtValue.Text) = "" Then
                        MsgBox "Please enter a valid filter value.", vbInformation, "Cubepoint Library"
                        txtValue.SetFocus
                        IsValidCondition = False
                        Exit Function
                    End If
                    
                Case "NUMBER", "INTEGER", "LONG", "DOUBLE"
                    If cboOperator.Text = "between" Then
                        arrValues = Split(UCase(txtValue.Text), " AND ")
                        If UBound(arrValues) <= 0 Then
                            MsgBox "The value you enter must be in this format: '<Value 1> and <Value 2>'", _
                                    vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                        
                        If InStr(1, arrValues(0), ">") > 0 And InStr(1, arrValues(0), "<") > 0 Then
                            arrValues(0) = Replace(Replace(arrValues(0), "<", ""), ">", "")
                            arrValues(1) = Replace(Replace(arrValues(1), "<", ""), ">", "")
                            If Not IsNumeric(arrValues(0)) Or Not IsNumeric(arrValues(1)) Then
                                MsgBox "Please enter a valid filter value.", vbInformation, "Cubepoint Library"
                                txtValue.SetFocus
                                IsValidCondition = False
                                Exit Function
                            End If
                        Else
                            MsgBox "The value you enter must be in this format: '<Value 1> and <Value 2>'", _
                                    vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                        
                    Else
                        If Not IsNumeric(txtValue.Text) Then
                            MsgBox "Please enter a valid filter value.", vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                    End If
                    
                Case "DATE"
                    If cboOperator.Text = "between" Then
                        arrValues = Split(UCase(txtValue.Text), " AND ")
                        If UBound(arrValues) = 0 Then
                            MsgBox "The value you enter must be in this format: '<Value 1> and <Value 2>'", _
                                    vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                        
                        arrValues(0) = Trim$(arrValues(0))
                        arrValues(1) = Trim$(arrValues(1))
                        
                        If Left$(arrValues(0), 1) = "<" And Right$(arrValues(0), 1) = ">" And _
                            Left$(arrValues(1), 1) = "<" And Right$(arrValues(1), 1) = ">" Then

                            arrValues(0) = Replace(Replace(arrValues(0), "<", ""), ">", "")
                            arrValues(1) = Replace(Replace(arrValues(1), "<", ""), ">", "")
                                                                                    
                            If Not IsDate(arrValues(0)) Then
                                MsgBox arrValues(0) & " is an invalid filter value.", vbInformation, "Cubepoint Library"
                                txtValue.SetFocus
                                IsValidCondition = False
                                Exit Function
                                
                            ElseIf Not IsDate(arrValues(1)) Then
                                MsgBox arrValues(1) & " is an invalid filter value.", vbInformation, "Cubepoint Library"
                                                                
                                txtValue.SetFocus
                                IsValidCondition = False
                                Exit Function
                            Else
                                ' Must pass back because we removed the '<' and '>' characters as well as the
                                ' 'AND' keyword
                                Value1 = arrValues(0)
                                Value2 = arrValues(1)
                            End If
                        Else
                            MsgBox "The value you enter must be in this format: '<Value 1> and <Value 2>'", _
                                    vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                    Else
                        If Not IsDate(txtValue.Text) Then
                            MsgBox "Please enter a valid filter value.", vbInformation, "Cubepoint Library"
                            txtValue.SetFocus
                            IsValidCondition = False
                            Exit Function
                        End If
                    End If
                    
                Case "BOOLEAN"
                    If UCase(Trim(txtValue.Text)) <> "TRUE" And UCase(Trim(txtValue.Text)) <> "FALSE" Then
                        MsgBox "Please enter either 'True' or 'False' as value.", vbInformation, "Cubepoint Library"
                        txtValue.SetFocus
                        IsValidCondition = False
                        Exit Function
                    End If
                    
            End Select
            Exit For
        End If
    Next

    IsValidCondition = True
    
End Function


Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdClear_Click()

    '>> clear all filters
    txtSearchKeys.Text = ""
    cboFields.ListIndex = 0
    cmdRemove.Enabled = False
    
    lvwConditions.ListItems.Clear
    cboField.ListIndex = -1
    cboOperator.ListIndex = -1
    txtValue.Text = ""
    
End Sub

Private Sub cmdClear_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyTab Then
        sstFilter.SetFocus
        KeyCode = 0
    End If
    
End Sub

Private Sub cmdOK_Click()
    
        
    '>> save all filters
    If SaveFilters = True Then
        Unload Me
    End If
    
End Sub

Private Sub cmdRemove_Click()
    
    '>> remove filter string from the listview
    lvwConditions.ListItems.Remove lvwConditions.SelectedItem.Index
    
    If lvwConditions.ListItems.Count = 0 Then
        cmdRemove.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
        
    
    Dim strFilterTab As String
    
    Set m_clsRegistry = New CRegistry
    
    m_clsRegistry.GetRegistry cpiCurrentUser, g_typInterface.IApplication.ProductName, "Library2003", "FilterTab"
    
    strFilterTab = m_clsRegistry.RegistryValue
    
    If strFilterTab = "0" Then
        sstFilter.Tab = 0
    Else
        sstFilter.Tab = 1
    End If
    
    '>> load combobox items and filters
    Call LoadFields
    Call LoadFilters
    
End Sub

Private Sub LoadFields()

    Dim rstFields As ADODB.Recordset
    Dim lngCtr As Long
    Dim strCommandText As String
    
    '>> load available fields to combo box
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_ID, "
    strCommandText = strCommandText & "DVC_FieldSource, "
    strCommandText = strCommandText & "DVC_FieldAlias, "
    strCommandText = strCommandText & "DVC_DataType "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "DVC_ID IN (" & Replace(clsFilter.DVCIDs, "*****", ",") & ") "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "DVC_FieldAlias "
    
    ADORecordsetOpen strCommandText, conFilter, rstFields, adOpenKeyset, adLockOptimistic
    'Set rstFields = conFilter.Execute(strCommandText)
    
    lngCtr = 0
    
    Do While Not rstFields.EOF
    
        ReDim Preserve arrFields(lngCtr + 3)
        arrFields(lngCtr) = rstFields!DVC_ID
        arrFields(lngCtr + 1) = rstFields!DVC_FieldSource
        arrFields(lngCtr + 2) = rstFields!DVC_FieldAlias
        arrFields(lngCtr + 3) = rstFields!DVC_DataType
        
        lngCtr = lngCtr + 4
        
        cboFields.AddItem FNullField(rstFields!DVC_FieldAlias)
        cboField.AddItem FNullField(rstFields!DVC_FieldAlias)
    
        rstFields.MoveNext
    Loop
    
    cboFields.ListIndex = 0
    
    ' by hobbes 10/18/2005
    Call ADORecordsetClose(rstFields)
        
End Sub

Private Sub LoadFilters()

    Dim rstFilters As ADODB.Recordset
    Dim lngCtr As Long
    Dim strCommandText As String
    Dim strField As String
    
    '>> load all filters of the selected view
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Filter "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & clsFilter.UVC_ID
    
    ADORecordsetOpen strCommandText, conFilter, rstFilters, adOpenKeyset, adLockOptimistic
    'Set rstFilters = conFilter.Execute(strCommandText)
        
    Do While Not rstFilters.EOF
        If rstFilters!Filter_Type = 1 Then
            txtSearchKeys.Text = txtSearchKeys.Text & FNullField(rstFilters!Filter_Value) & ", "
            
            For lngCtr = 0 To UBound(arrFields)
                If arrFields(lngCtr + 1) = rstFilters!Filter_Field Then
                    strField = arrFields(lngCtr + 2)
                    Exit For
                End If
            Next
            cboFields.Text = strField
        Else
            For lngCtr = 0 To UBound(arrFields) Step 4
                If rstFilters!Filter_Field = arrFields(lngCtr + 1) Then
                    strField = arrFields(lngCtr + 2)
                    Exit For
                End If
            Next
            
            lvwConditions.ListItems.Add lvwConditions.ListItems.Count + 1, , strField
            Select Case rstFilters!Filter_Operator
                Case 1
                    Select Case rstFilters!Filter_DataType
                        Case 3
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID1", "on"
                        Case 4
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID1", "is true"
                        Case Else
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID1", "contains"
                    End Select
                Case 2
                    Select Case rstFilters!Filter_DataType
                        Case 1
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID2", "is(exactly)"
                        Case 2
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID2", "is equal"
                        Case 3
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID2", "on or before"
                        Case 4
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID1", "is false"
                    End Select
                Case 3
                    Select Case rstFilters!Filter_DataType
                        Case 1
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID3", "doesn't contain"
                        Case 2
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID3", "less than or equal"
                        Case 3
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID3", "on or after"
                        Case 4
                        
                    End Select
                    
                Case 4
                    Select Case rstFilters!Filter_DataType
                        Case 1
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID4", "is empty"
                        Case 2
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID4", "greater than or equal"
                        Case 3
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID4", "between"
                        Case 4
                        
                    End Select
                Case 5
                    Select Case rstFilters!Filter_DataType
                        Case 1
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID5", "is not empty"
                        Case 2
                            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 1, "ID5", "between"
                        Case 3
                            Debug.Assert False
                        Case 4
                        
                    End Select
                    
            End Select
            
            lvwConditions.ListItems(lvwConditions.ListItems.Count).ListSubItems.Add 2, , FNullField(rstFilters!Filter_Value)
                
            cmdRemove.Enabled = True
            
        End If
        
        rstFilters.MoveNext
    Loop
        
    On Error Resume Next
    txtSearchKeys.Text = Mid(txtSearchKeys.Text, 1, Len(txtSearchKeys.Text) - 3) & Trim(Replace(Mid(txtSearchKeys.Text, Len(txtSearchKeys.Text) - 2), ",", ""))
    On Error GoTo 0
    
    ' by hobbes 10/18/2005
    Call ADORecordsetClose(rstFilters)
    
End Sub

Private Function SaveFilters() As Boolean
    
    Dim lngFilterCtr As Long
    Dim lngCtr As Long
    Dim lngDataType As Long
    
    Dim strCommandText As String
    Dim strField As String
    Dim strDataType As String
    Dim arrFilters
    
    '>> delete current filters
    strCommandText = vbNullString
    strCommandText = strCommandText & "DELETE "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Filter "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & clsFilter.UVC_ID
    
    ExecuteNonQuery conFilter, strCommandText
    'conFilter.Execute strCommandText
    
    '>> save all filters
    If Trim(txtSearchKeys.Text) <> "" Then
        If Mid(Trim(txtSearchKeys.Text), Len(Trim(txtSearchKeys.Text)), 1) = "," Then
            txtSearchKeys.Text = Mid(Trim(txtSearchKeys.Text), 1, Len(Trim(txtSearchKeys.Text)) - 1)
        End If
        
        arrFilters = Split(Trim(txtSearchKeys.Text), ",")
        
        For lngFilterCtr = 0 To UBound(arrFilters)
            For lngCtr = 0 To UBound(arrFields) Step 4
                If cboFields.Text = arrFields(lngCtr + 2) Then
                    strField = arrFields(lngCtr + 1)
                    strDataType = arrFields(lngCtr + 3)
                    Exit For
                End If
            Next
            
            Select Case Trim(UCase(strDataType))
                Case "TEXT"
                    lngDataType = 1
                    
                Case "NUMBER", "INTEGER", "LONG", "DOUBLE"
                    lngDataType = 2
                    If Not IsNumeric(Trim$(arrFilters(lngFilterCtr))) Then
                        MsgBox "Please enter valid Filter keys.", vbInformation + vbOKOnly, "Cubepoint Library"
                        txtSearchKeys.SetFocus
                        SaveFilters = False
                        Exit Function
                    End If
                    
                Case "DATE"
                    lngDataType = 3
                    If Not IsDate(Trim$(arrFilters(lngFilterCtr))) And Trim$(arrFilters(lngFilterCtr)) <> "" Then
                        MsgBox "Please enter valid Filter keys.", vbInformation + vbOKOnly, "Cubepoint Library"
                        txtSearchKeys.SetFocus
                        SaveFilters = False
                        Exit Function
                    End If
                    
                Case "BOOLEAN"
                    lngDataType = 4
                    If (UCase(Trim$(arrFilters(lngFilterCtr))) <> Trim(UCase(CBool("True"))) Or UCase(Trim$(arrFilters(lngFilterCtr))) <> Trim(UCase(CBool("False")))) And Trim$(arrFilters(lngFilterCtr)) <> "" Then
                        MsgBox "Please enter valid Filter keys.", vbInformation + vbOKOnly, "Cubepoint Library"
                        txtSearchKeys.SetFocus
                        SaveFilters = False
                        Exit Function
                    End If
                    
            End Select
            
            '>> add basic filter defined
            strCommandText = vbNullString
            strCommandText = strCommandText & "INSERT INTO "
            strCommandText = strCommandText & "Filter "
            strCommandText = strCommandText & "("
            strCommandText = strCommandText & "UVC_ID, "
            strCommandText = strCommandText & "Filter_Field, "
            strCommandText = strCommandText & "Filter_Operator, "
            strCommandText = strCommandText & "Filter_Value, "
            strCommandText = strCommandText & "Filter_Type, "
            strCommandText = strCommandText & "Filter_DataType "
            strCommandText = strCommandText & ") "
            strCommandText = strCommandText & "VALUES "
            strCommandText = strCommandText & "("
            strCommandText = strCommandText & clsFilter.UVC_ID & ", "
            strCommandText = strCommandText & "'" & strField & "', "
            If lngDataType = 1 Then
                strCommandText = strCommandText & "1, "
                strCommandText = strCommandText & "'" & AQ(Trim(arrFilters(lngFilterCtr))) & "', "
            Else
                If lngDataType = 2 Then
                    strCommandText = strCommandText & "2, "
                    strCommandText = strCommandText & AQ(Trim(arrFilters(lngFilterCtr))) & ", "
                ElseIf lngDataType = 3 Then
                    strCommandText = strCommandText & "1, " ' [Field Name] ON [Date Value]
                    strCommandText = strCommandText & "#" & AQ(Trim(arrFilters(lngFilterCtr))) & "#, "
                Else
                    strCommandText = strCommandText & "2, "
                    strCommandText = strCommandText & IIf(AQ(Trim(arrFilters(lngFilterCtr))), "True", "False") & ", "
                End If
            End If
            strCommandText = strCommandText & 1 & ", "
            strCommandText = strCommandText & lngDataType
            strCommandText = strCommandText & ") "
                   
            ExecuteNonQuery conFilter, strCommandText
            'conFilter.Execute strCommandText
        Next
    End If
    
    For lngFilterCtr = 1 To lvwConditions.ListItems.Count
        For lngCtr = 0 To UBound(arrFields) Step 4
            If lvwConditions.ListItems(lngFilterCtr).Text = arrFields(lngCtr + 2) Then
                strField = arrFields(lngCtr + 1)
                
                Select Case Trim(UCase(arrFields(lngCtr + 3)))
                    Case "TEXT"
                        lngDataType = 1
                        
                    Case "NUMBER", "INTEGER", "LONG", "DOUBLE"
                        lngDataType = 2
                        
                    Case "DATE"
                        lngDataType = 3
                        
                    Case "BOOLEAN"
                        lngDataType = 4
                        
                End Select
                Exit For
            End If
        Next
        
        
        
        '>> add advanced filters defined
        strCommandText = vbNullString
        strCommandText = strCommandText & "INSERT INTO "
        strCommandText = strCommandText & "Filter "
        strCommandText = strCommandText & "("
        strCommandText = strCommandText & "UVC_ID, "
        strCommandText = strCommandText & "Filter_Field, "
        strCommandText = strCommandText & "Filter_Operator, "
        strCommandText = strCommandText & "Filter_Value, "
        strCommandText = strCommandText & "Filter_Type, "
        strCommandText = strCommandText & "Filter_DataType "
        strCommandText = strCommandText & ") "
        strCommandText = strCommandText & "VALUES "
        strCommandText = strCommandText & "("
        strCommandText = strCommandText & clsFilter.UVC_ID & ", "
        strCommandText = strCommandText & "'" & strField & "', "
        strCommandText = strCommandText & Val(Mid(lvwConditions.ListItems(lngFilterCtr).ListSubItems(1).Key, 3)) & ", "
        If lngDataType = 2 Then
            strCommandText = strCommandText & CDbl(lvwConditions.ListItems(lngFilterCtr).ListSubItems(2).Text) & ", "
        Else
            strCommandText = strCommandText & "'" & AQ(lvwConditions.ListItems(lngFilterCtr).ListSubItems(2).Text, Apostrophe) & "', "
        End If
        strCommandText = strCommandText & 2 & ", "
        strCommandText = strCommandText & lngDataType
        strCommandText = strCommandText & ") "
        
        ExecuteNonQuery conFilter, strCommandText
        'conFilter.Execute strCommandText
    Next

    If frmEditView.Visible = False Then
        clsFilter.DataChanged = True
    End If

    SaveFilters = True
    
End Function

Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set conFilter = ADOConnection
    Set clsFilter = GridProps
    
    Set Me.Icon = Window.Icon

    Me.Show vbModal
    
    Set GridProps = clsFilter
    Set ADOConnection = conFilter
    
    Set clsFilter = Nothing
    Set conFilter = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call m_clsRegistry.SaveRegistry(cpiCurrentUser, g_typInterface.IApplication.ProductName, "Library2003", "FilterTab", sstFilter.Tab)
    
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyTab Then
        If sstFilter.Tab = 0 Then
            txtSearchKeys.SetFocus
        Else
            If lvwConditions.ListItems.Count > 0 Then
                lvwConditions.SetFocus
            Else
                cboField.SetFocus
            End If
        End If
        KeyCode = 0
    End If

End Sub

Private Sub txtSearchKeys_GotFocus()

    If sstFilter.Tab = 1 Then
        cmdOk.SetFocus
    Else
        SendKeysEx "{HOME}+{END}"
    End If
    
End Sub



Private Sub txtValue_GotFocus()

    SendKeysEx "{HOME}+{END}"

End Sub
