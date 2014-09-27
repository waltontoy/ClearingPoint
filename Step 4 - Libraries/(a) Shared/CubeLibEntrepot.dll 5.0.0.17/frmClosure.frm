VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmClosure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Closure / Re-Opening"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmClosure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      Pattern         =   "MDB_HISTORY*.MDB"
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraClosure 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmdPicklist 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   0
         Top             =   240
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy h:nn"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5132
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   38616
      End
      Begin VB.TextBox txtClosure 
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Index           =   1
         Left            =   4320
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   38616
      End
      Begin VB.TextBox txtClosure 
         Height          =   330
         Index           =   1
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblClosure 
         Alignment       =   2  'Center
         Caption         =   "dated "
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   14
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label lblClosure 
         Caption         =   "Re-Open with IM7 :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   998
         Width           =   1455
      End
      Begin VB.Label lblClosure 
         Caption         =   "Closure Date :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   638
         Width           =   1455
      End
      Begin VB.Label lblClosure 
         Caption         =   "Entrepot Number :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Frame fraCancel 
      Height          =   3735
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   3240
         Width           =   855
      End
      Begin GridEX16.GridEX GridEX1 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5106
         TabKeyBehavior  =   1
         CursorLocation  =   3
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowDelete     =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnCount     =   4
         CardCaption1    =   -1  'True
         ColCaption1     =   "Entrepot Number"
         ColWidth1       =   1395
         ColCaption2     =   "Closure Date"
         ColWidth2       =   1200
         ColCaption3     =   "IM7 for Re-opening"
         ColWidth3       =   1605
         ColCaption4     =   "Re-opening Date"
         ColWidth4       =   1395
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Frame fraProgress 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Processing ..."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Closure"
            Key             =   "Closure"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Closure Cancellation"
            Key             =   "XClosure"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClosure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_conSADBEL As ADODB.Connection
Private m_conHistory As ADODB.Connection

Private m_strLanguage As String
Private m_ResourceHandler As Long
Private blnCancelProcessing As Boolean
Private blnSaveFirstBeforeExiting As Boolean
Private rstOfflineRecordset As ADODB.Recordset
Private strDocNum As String
Private strDocDate As String
Private strEntNum As String
Private mrstStockCard As ADODB.Recordset

Private m_lngUserID As Long

Public Sub MyLoad(ByRef SADBELDB As ADODB.Connection, _
                    ByVal strLanguage As String, _
                    ByVal ResourceHandler As Long, _
                    ByVal UserID As Long)
    Set m_conSADBEL = SADBELDB
    m_lngUserID = UserID
    
    m_strLanguage = strLanguage
    m_ResourceHandler = ResourceHandler
    
    cmdProcess(0).Enabled = False
    Me.Show vbModal
End Sub

Private Sub cmdDelete_Click()
    Dim strClosureDocumentYear As String
    Dim strHistoryPathForChecking As String
    
        With rstOfflineRecordset
            If Not (.BOF And .EOF) Then
                .MoveFirst
            Else
                Exit Sub
            End If
            
            Do While Not .EOF
                If GridEX1.Value(GridEX1.Columns("Entrepot Number").Index) = .Fields("Entrepot Number").Value Then
                    If DateValue(GridEX1.Value(GridEX1.Columns("Closure Date").Index)) < DateValue(.Fields("Closure Date").Value) Then
                        MsgBox "Please delete first the closure document that was made after this one.", vbInformation + vbOKOnly, "Closure Cancellation"
                        Exit Sub
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        strClosureDocumentYear = Year(GridEX1.Value(GridEX1.Columns("Closure Date").Index))
        strClosureDocumentYear = Right$(strClosureDocumentYear, 2)
        
        strHistoryPathForChecking = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & "mdb_history" & strClosureDocumentYear & ".mdb"
        If Len(Trim(Dir(strHistoryPathForChecking))) = 0 Then
            MsgBox "The Closure cannot be deleted/cancelled because the database " & Trim(strHistoryPathForChecking) & " is missing.", vbInformation, "ClearingPoint"
            Exit Sub
        End If
        
        cmdProcess(0).Enabled = True
        strDocNum = strDocNum & "*" & GridEX1.Value(GridEX1.Columns("IM7 for Re-opening").Index)
        strDocDate = strDocDate & "*" & GridEX1.Value(GridEX1.Columns("Closure Date").Index)
        strEntNum = strEntNum & "*" & GridEX1.Value(GridEX1.Columns("Entrepot Number").Index)
                
        
        
        GridEX1.Delete
        
        If GridEX1.RowCount = 0 Then cmdDelete.Enabled = False
End Sub

Private Sub cmdPicklist_Click(Index As Integer)
    Dim clsEntrepot As cEntrepot
    
    Set clsEntrepot = New cEntrepot
    
    clsEntrepot.ShowEntrepot Me, m_conSADBEL, True, m_strLanguage, m_ResourceHandler
    
    If clsEntrepot.Cancelled = False Then
        txtClosure(0).Text = clsEntrepot.SelectedEntrepot
        txtClosure(0).Tag = clsEntrepot.Entrepot_ID
    End If
    
    Set clsEntrepot = Nothing
End Sub

Private Sub cmdProcess_Click(Index As Integer)
    Dim blnContinueProcess As Boolean
    Dim strDocument() As String
    Dim strIndate() As String
    Dim strEnt() As String
    Dim lngCounter As Long
    
    Select Case Index
        Case 0
            If fraClosure.Visible = True Then
                'Do the processing
                Me.MousePointer = vbHourglass
                ClosetheStocks
                Me.MousePointer = vbDefault
            ElseIf fraCancel.Visible = True Then
                'DELETE!!! -> UPDATE SADBEL
                strDocument = Split(strDocNum, "*")
                strIndate = Split(strDocDate, "*")
                strEnt = Split(strEntNum, "*")
                For lngCounter = 1 To UBound(strDocument)
                    If strDocument(lngCounter) <> "" And strIndate(lngCounter) <> "" Then
                        DeleteDoc strDocument(lngCounter), strIndate(lngCounter), strEnt(lngCounter)
                    End If
                Next lngCounter
                PopulateGrid
                cmdProcess(0).Enabled = False
            End If
            
        Case 1
            If fraProgress.Visible = True Then
                'Cancel closure process
                Select Case MsgBox("Do you want to revert changes?", vbInformation + vbYesNoCancel, "Closure")
                    Case vbCancel
                        'Continue processing
                        blnContinueProcess = True
                        blnCancelProcessing = False
                    Case vbYes
                        'Rollback then End processing
                        blnContinueProcess = False
                        blnCancelProcessing = True
                        blnSaveFirstBeforeExiting = False
                    Case vbNo
                        'Save then stop
                        blnContinueProcess = False
                        blnCancelProcessing = True
                        blnSaveFirstBeforeExiting = True
                End Select

                If blnContinueProcess = False Then
                    cmdProcess(1).Caption = "&Close"
                    fraClosure.Visible = True
                    fraProgress.Visible = False
                    ProgressBar1.Visible = False
                    Me.Height = 5415
                    cmdProcess(0).Visible = True
                    cmdProcess(1).Visible = True
                    cmdProcess(1).Top = 4440
                    TabStrip1.Visible = True
                End If
            Else
                Unload Me
            End If
            
    End Select
End Sub

Private Sub DTPicker1_Change(Index As Integer)
    If Index = 1 Then
        If DateValue(DTPicker1(Index)) < DateValue(DTPicker1(Index - 1)) Then
            DTPicker1(Index) = (DTPicker1(Index - 1))
        ElseIf DateValue(DTPicker1(Index)) > DateValue(DTPicker1(Index - 1) + 1) And DateValue(DTPicker1(Index - 1) + 1) <> DateValue(Now + 1) Then
            DTPicker1(Index) = (DTPicker1(Index - 1) + 1)
        End If
    Else
        If DateValue(DTPicker1(Index + 1)) > DateValue(DTPicker1(Index) + 1) Then
            DTPicker1(Index + 1) = (DTPicker1(Index))
        ElseIf DateValue(DTPicker1(Index)) > DateValue(DTPicker1(Index + 1)) Then
            DTPicker1(Index + 1) = (DTPicker1(Index))
        End If
    End If
End Sub

Private Sub DTPicker1_Click(Index As Integer)
    If Index = 1 Then
        If DateValue(DTPicker1(Index)) < DateValue(DTPicker1(Index - 1)) Then
            DTPicker1(Index) = (DTPicker1(Index - 1))
        ElseIf DateValue(DTPicker1(Index)) > DateValue(DTPicker1(Index - 1) + 1) And DateValue(DTPicker1(Index - 1) + 1) <> DateValue(Now + 1) Then
            DTPicker1(Index) = (DTPicker1(Index - 1) + 1)
        End If
    Else
        If DateValue(DTPicker1(Index + 1)) > DateValue(DTPicker1(Index) + 1) Then
            DTPicker1(Index + 1) = (DTPicker1(Index))
        ElseIf DateValue(DTPicker1(Index)) > DateValue(DTPicker1(Index + 1)) Then
            DTPicker1(Index + 1) = (DTPicker1(Index))
        End If
    End If
End Sub

Private Sub Form_Load()
    DTPicker1(0).MaxDate = Now - 1
    DTPicker1(1).MaxDate = Now
    DTPicker1(0).Value = DateValue(Now - 1)
    DTPicker1(1).Value = DateValue(Now)
    
    fraCancel.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rstOfflineRecordset Is Nothing Then
        If rstOfflineRecordset.State = adStateOpen Then
            rstOfflineRecordset.Close
        End If
        Set rstOfflineRecordset = Nothing
    End If
End Sub

Private Sub GridEX1_AfterDelete()
        If GridEX1.RowCount = 0 Then cmdDelete.Enabled = False
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    Dim lngSortOrder As Long
    lngSortOrder = 0
    
    If GridEX1.SortKeys.Count = 1 Then
        If GridEX1.SortKeys.Item(1).ColIndex = Column.Index Then
            lngSortOrder = GridEX1.SortKeys.Item(1).SortOrder
        End If
    End If
    
    GridEX1.SortKeys.Clear
    GridEX1.SortKeys.Add Column.Index, IIf(lngSortOrder = 1, jgexSortDescending, jgexSortAscending)
    GridEX1.RefreshSort
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        With rstOfflineRecordset
            If Not (.BOF And .EOF) Then
                .MoveFirst
            Else
                Exit Sub
            End If

            Do While Not .EOF
                If DateValue(GridEX1.Value(GridEX1.Columns("Closure Date").Index)) < DateValue(.Fields("Closure Date").Value) Then
                    MsgBox "Please delete first the closure document that was made after this one.", vbInformation + vbOKOnly, "Closure Cancellation"
                    KeyCode = 0
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
        If KeyCode = 46 Then
            cmdProcess(0).Enabled = True
            strDocNum = strDocNum & "*" & GridEX1.Value(GridEX1.Columns("IM7 for Re-opening").Index)
            strDocDate = strDocDate & "*" & GridEX1.Value(GridEX1.Columns("Closure Date").Index)
            strEntNum = strEntNum & "*" & GridEX1.Value(GridEX1.Columns("Entrepot Number").Index)
        End If
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Key = "Closure" Then
        fraClosure.Visible = True
        fraCancel.Visible = False
        If Not rstOfflineRecordset Is Nothing Then
            If rstOfflineRecordset.State = adStateOpen Then
                rstOfflineRecordset.Close
                Set rstOfflineRecordset = Nothing
            End If
        End If
        If txtClosure(0).Text <> "" And txtClosure(1) <> "" Then
            cmdProcess(0).Enabled = True
        Else
            cmdProcess(0).Enabled = False
        End If
        cmdProcess(0).Caption = "&Process"
    Else
        fraCancel.Visible = True
        fraClosure.Visible = False
        strDocNum = ""
        cmdProcess(0).Enabled = False
        cmdProcess(0).Caption = "&Update"
        PopulateGrid
    End If
End Sub

Private Sub txtClosure_Change(Index As Integer)
    If Len(txtClosure(1).Text) = 7 And txtClosure(0).Text <> "" Then
        cmdProcess(0).Enabled = True
    Else
        cmdProcess(0).Enabled = False
    End If
End Sub

Private Sub txtClosure_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Chr(KeyAscii) = "'" Then
            KeyAscii = 0
        End If

        If Len(txtClosure(1).Text) = 7 And txtClosure(0).Text <> "" Then
            cmdProcess(0).Enabled = True
        Else
            cmdProcess(0).Enabled = False
        End If
    ElseIf Index = 0 Then
        If Len(txtClosure(1).Text) = 7 And txtClosure(0).Text <> "" Then
            cmdProcess(0).Enabled = True
        Else
            cmdProcess(0).Enabled = False
        End If
    End If
End Sub

Private Sub txtClosure_LostFocus(Index As Integer)
    Dim lngZeroCount As Long
    Dim lngCounter As Long
    Dim strTextValue As String
    
    If Len(txtClosure(1).Text) < 7 And Len(txtClosure(1).Text) <> 0 And Index = 1 Then
        lngZeroCount = 7 - Len(txtClosure(1).Text)
        strTextValue = txtClosure(1).Text
        txtClosure(1).Text = ""
        For lngCounter = 1 To lngZeroCount
            txtClosure(1).Text = txtClosure(1).Text & "0"
        Next lngCounter
        txtClosure(1).Text = txtClosure(1).Text & strTextValue
    End If
    
    If txtClosure(0).Text <> "" And txtClosure(1).Text <> "" Then
        cmdProcess(0).Enabled = True
    Else
        cmdProcess(0).Enabled = False
    End If
End Sub

Private Sub ClosetheStocks()
    Dim rstInbounds As ADODB.Recordset
    Dim rstInboundDocs As ADODB.Recordset
    Dim rstOutbounds As ADODB.Recordset
    Dim rstOutboundDocs As ADODB.Recordset
    
    Dim strSQL As String
    Dim strInDoc_ID As String
    Dim strIn_ID As String
    Dim strOut_ID As String
    Dim strOutDoc_ID As String
    Dim blnSave As Boolean
    Dim rstEntrepot As ADODB.Recordset
    Dim strSEQ As String
    Dim blnInDocExisting As Boolean
    Dim blnOutDocExisting As Boolean
    Dim blnDateExisting As Boolean
    
    Dim rstStockcards As ADODB.Recordset
    Dim lngCtr As Long
    Dim rstChecker As ADODB.Recordset
    Dim lngCtrCtr As Long
    
    Dim conHist As ADODB.Connection
    Dim conHistory As ADODB.Connection  'Dim datHistory As DAO.Database
    
    Dim strCommand As String
    Dim strHistoryPathForChecking As String
    Dim strHistoryYear As String
    
    strHistoryYear = Right(DTPicker1(0).Year, 2)
    
    '<<< dandan 112107
    'Added checking for missing databases
    strHistoryPathForChecking = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & "mdb_history" & strHistoryYear & ".mdb"
    
    If Len(Trim(Dir(strHistoryPathForChecking))) = 0 Then
        MsgBox "Closure of stock is not available because the database " & Trim(strHistoryPathForChecking) & " is missing.", vbInformation, "ClearingPoint"
        Exit Sub
    Else
        ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, strHistoryYear
        'OpenADODatabase conHistory, m_strPath, "mdb_history" & strHistoryYear & ".mdb"
    End If
    
    blnCancelProcessing = False
    blnSaveFirstBeforeExiting = True
    blnSave = True
        
    ' Recreate temporary tables
    '       A. TempInboundDocs
    '       B. TempInbounds
    '       C. TempOutbounds
    '       D. TempOutboundDocs
    ' Reims:    Why should the Primary Keys be NULL?
    On Error Resume Next
        strCommand = vbNullString
        strCommand = strCommand & "DROP TABLE "
        strCommand = strCommand & "TempInboundDocs" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conSADBEL, strCommand
    On Error GoTo 0
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "INTO "
        strCommand = strCommand & "TempInboundDocs" & "_" & Format(m_lngUserID, "00") & " "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "InboundDocs "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "InDoc_ID IS NULL "
    ExecuteNonQuery m_conSADBEL, strCommand
    
    On Error Resume Next
        strCommand = vbNullString
        strCommand = strCommand & "DROP TABLE "
        strCommand = strCommand & "TempInbounds" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conSADBEL, strCommand
    On Error GoTo 0
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "INTO "
        strCommand = strCommand & "TempInbounds" & "_" & Format(m_lngUserID, "00") & " "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "Inbounds "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "In_ID IS NULL "
    ExecuteNonQuery m_conSADBEL, strCommand
    
    On Error Resume Next
        strCommand = vbNullString
        strCommand = strCommand & "DROP TABLE "
        strCommand = strCommand & "TempOutbounds" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conSADBEL, strCommand
    On Error GoTo 0
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "INTO "
        strCommand = strCommand & "TempOutbounds" & "_" & Format(m_lngUserID, "00") & " "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "Outbounds "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "Out_ID IS NULL "
    ExecuteNonQuery m_conSADBEL, strCommand
    
    On Error Resume Next
        strCommand = vbNullString
        strCommand = strCommand & "DROP TABLE "
        strCommand = strCommand & "TempOutboundDocs" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conSADBEL, strCommand
    On Error GoTo 0
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "INTO "
        strCommand = strCommand & "TempOutboundDocs" & "_" & Format(m_lngUserID, "00") & " "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "OutboundDocs "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "OutboundDocs.OUTDOC_ID IS NULL "
    ExecuteNonQuery m_conSADBEL, strCommand
    
    
    If (blnSave = True) Then
        '<<< dandan 112806
        '<<< Create linked tables into the history db
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "TempInboundDocs" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "TempInboundDocs" & "_" & Format(m_lngUserID, "00"), strHistoryYear
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "TempInbounds" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "TempInbounds" & "_" & Format(m_lngUserID, "00"), strHistoryYear
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "TempOutbounds" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "TempOutbounds" & "_" & Format(m_lngUserID, "00"), strHistoryYear
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "TempOutboundDocs" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "TempOutboundDocs" & "_" & Format(m_lngUserID, "00"), strHistoryYear
        
        'AddLinkedTableEx "TempInboundDocs" & "_" & Format(m_lngUserID, "00"), conHistory.Properties("Data Source Name").Value, G_Main_Password, "TempInboundDocs" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_sadbel.mdb", G_Main_Password
        'AddLinkedTableEx "TempInbounds" & "_" & Format(m_lngUserID, "00"), conHistory.Properties("Data Source Name").Value, G_Main_Password, "TempInbounds" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_sadbel.mdb", G_Main_Password
        'AddLinkedTableEx "TempOutbounds" & "_" & Format(m_lngUserID, "00"), conHistory.Properties("Data Source Name").Value, G_Main_Password, "TempOutbounds" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_sadbel.mdb", G_Main_Password
        'AddLinkedTableEx "TempOutboundDocs" & "_" & Format(m_lngUserID, "00"), conHistory.Properties("Data Source Name").Value, G_Main_Password, "TempOutboundDocs" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_sadbel.mdb", G_Main_Password
    End If
    
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "STOCK_ID AS [ID], "
        strSQL = strSQL & "Choose(PROD_HANDLING + 1, 'In_Orig_Packages_Qty', 'In_Orig_Gross_Weight', 'In_Orig_Net_Weight') AS [HANDLING] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "STOCKCARDS "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "PRODUCTS "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "ENTREPOTS "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "ENTREPOTS.Entrepot_ID = " & txtClosure(0).Tag & " "
    ADORecordsetOpen strSQL, m_conSADBEL, rstStockcards, adOpenKeyset, adLockOptimistic
    
    
    blnInDocExisting = CheckIfExisting(m_conSADBEL, "InboundDocs", txtClosure(1).Text, Val(txtClosure(0).Tag))
    blnOutDocExisting = CheckIfExisting(m_conSADBEL, "OutboundDocs", txtClosure(1).Text, Val(txtClosure(0).Tag))
    blnDateExisting = CheckIfDateExisting(m_conSADBEL, DTPicker1(0).Value, txtClosure(0).Tag)
    
    strInDoc_ID = 0
    strOutDoc_ID = 0
    strSEQ = ""
    lngCtr = 0
    
    
    If (blnInDocExisting = False And blnOutDocExisting = False And blnDateExisting = False) Then
        fraClosure.Visible = False
        fraProgress.Visible = True
        fraProgress.Top = 100
        ProgressBar1.Visible = True
        Me.Height = fraProgress.Height + cmdProcess(1).Height + 810
        TabStrip1.Visible = False
        cmdProcess(0).Visible = False
        cmdProcess(1).Visible = True
        cmdProcess(1).Top = 1080
        cmdProcess(1).Caption = "Cancel"
        Me.Refresh
        
        DoEvents
        
        Me.MousePointer = vbHourglass
        If (blnCancelProcessing = True) Then GoTo EarlyExit
        
        Me.MousePointer = vbHourglass
        If (rstStockcards.BOF And rstStockcards.EOF) Then
            GoTo EarlyExit
            Exit Sub
        End If
        
        
            strSQL = "SELECT ENTREPOT_LASTSEQNUM AS [SEQ] FROM ENTREPOTS WHERE ENTREPOT_ID = " & txtClosure(0).Tag
        ADORecordsetOpen strSQL, m_conSADBEL, rstEntrepot, adOpenKeyset, adLockOptimistic
        If Not (rstEntrepot.EOF And rstEntrepot.BOF) Then
            rstEntrepot.MoveFirst
            
            strSEQ = IIf(IsNull(rstEntrepot.Fields("SEQ").Value), 1, rstEntrepot.Fields("SEQ").Value + 1)
            rstEntrepot.Fields("SEQ").Value = Val(strSEQ)
            rstEntrepot.Update
            
            UpdateRecordset m_conSADBEL, rstEntrepot, "ENTREPOTS"
        End If
        ADORecordsetClose rstEntrepot
        
        
            strSQL = "SELECT * FROM TempInboundDocs" & "_" & Format(m_lngUserID, "00")
        ADORecordsetOpen strSQL, m_conSADBEL, rstInboundDocs, adOpenKeyset, adLockOptimistic
        
        rstInboundDocs.AddNew
            strInDoc_ID = GenerateID
            Do Until IsIDUnique(conHistory, "InDoc_ID", "InboundDocs", Val(strInDoc_ID))
               strInDoc_ID = GenerateID
            Loop
            
            rstInboundDocs.Fields("InDoc_ID").Value = strInDoc_ID
            rstInboundDocs.Fields("InDoc_Type").Value = "IM7"
            rstInboundDocs.Fields("InDoc_Num").Value = txtClosure(1).Text
            
            If (DateValue(DTPicker1(1).Value) = DateValue(DTPicker1(0).Value)) Then
                DTPicker1(1).Hour = 23
                DTPicker1(1).Minute = 59
                DTPicker1(1).Second = 58
                DTPicker1(1).Refresh
                If (DTPicker1(1).Hour = 0) Then
                    rstInboundDocs.Fields("InDoc_Date").Value = Replace(DTPicker1(1).Value, "0:59", "23:59")
                    If (UCase(Right(rstInboundDocs.Fields("InDoc_Date").Value, 2)) = "AM") Then
                        rstInboundDocs.Fields("InDoc_Date").Value = Left(rstInboundDocs.Fields("InDoc_Date").Value, Len(rstInboundDocs.Fields("InDoc_Date").Value) - 2) & "PM"
                        rstInboundDocs.Fields("InDoc_Date").Value = Replace(rstInboundDocs.Fields("InDoc_Date").Value, "12:59", "11:59")
                    End If
                Else
                    rstInboundDocs.Fields("InDoc_Date").Value = DTPicker1(1).Value
                End If
            Else
                rstInboundDocs.Fields("InDoc_Date").Value = DateValue(DTPicker1(1).Value) & " 00:00:01"
            End If
            
            rstInboundDocs.Fields("InDoc_Office").Value = "0"
            rstInboundDocs.Fields("InDoc_SeqNum").Value = strSEQ
            rstInboundDocs.Fields("InDoc_Global").Value = 0
        rstInboundDocs.Update
        
        InsertRecordset m_conSADBEL, rstInboundDocs, "TempInboundDocs"
        
        ADORecordsetClose rstInboundDocs
        
        
            strSQL = "SELECT * FROM TempOutboundDocs" & "_" & Format(m_lngUserID, "00")
        ADORecordsetOpen strSQL, m_conSADBEL, rstOutboundDocs, adOpenKeyset, adLockOptimistic
        
        rstOutboundDocs.AddNew
            strOutDoc_ID = GenerateID
            Do Until IsIDUnique(conHistory, "OutDoc_ID", "OutboundDocs", Val(strOutDoc_ID))
               strOutDoc_ID = GenerateID
            Loop
            
            rstOutboundDocs.Fields("OutDoc_ID").Value = strOutDoc_ID
            rstOutboundDocs.Fields("OutDoc_Type").Value = "XXX"
            rstOutboundDocs.Fields("OutDoc_Num").Value = txtClosure(1).Text
            
            DTPicker1(0).Hour = 23
            DTPicker1(0).Minute = 59
            DTPicker1(0).Second = 57
            DTPicker1(0).Refresh
            
            If (DTPicker1(0).Hour = 0) Then
                rstOutboundDocs.Fields("OutDoc_Date").Value = Replace(DTPicker1(0).Value, "0:59", "23:59")
                If (UCase(Right(rstOutboundDocs.Fields("OutDoc_Date").Value, 2)) = "AM") Then
                    rstOutboundDocs.Fields("OutDoc_Date").Value = Left(rstOutboundDocs.Fields("OutDoc_Date").Value, Len(rstOutboundDocs.Fields("OutDoc_Date").Value) - 2) & "PM"
                    rstOutboundDocs.Fields("OutDoc_Date").Value = Replace(rstOutboundDocs.Fields("OutDoc_Date").Value, "12:59", "11:59")
                End If
            Else
                rstOutboundDocs.Fields("OutDoc_Date").Value = DTPicker1(0).Value
            End If
            
            rstOutboundDocs.Fields("OutDoc_Comm_Settlement").Value = "4071"
            rstOutboundDocs.Fields("OutDoc_Global").Value = 0
        rstOutboundDocs.Update
        
        InsertRecordset m_conSADBEL, rstOutboundDocs, "TempOutboundDocs"
        
        ADORecordsetClose rstOutboundDocs
        
        
        rstStockcards.MoveFirst
        
        File1.Pattern = "mdb_History**.mdb"
        File1.Path = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
        
        Do While Not rstStockcards.EOF
            lngCtr = lngCtr + 1
            ProgressBar1.Value = (Val(lngCtr) / Val(rstStockcards.RecordCount)) * 100
            
                strSQL = "SELECT TOP 1 * FROM Inbounds WHERE STOCK_ID = " & rstStockcards.Fields("ID").Value
            ADORecordsetOpen strSQL, m_conSADBEL, rstChecker, adOpenKeyset, adLockOptimistic
            
            
            If (rstChecker.BOF And rstChecker.EOF) Then
                For lngCtrCtr = 0 To File1.ListCount - 1
                    
                    ADOConnectDB conHist, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(File1.List(lngCtrCtr))
                    'OpenADODatabase conHist, m_strPath, File1.List(lngCtrCtr)
                    
                    ADORecordsetOpen strSQL, conHist, rstChecker, adOpenKeyset, adLockOptimistic
                    'rstChecker.Open strSQL, conHist, adOpenKeyset, adLockOptimistic
                    If (rstChecker.BOF And rstChecker.EOF) Then
                        If lngCtrCtr = File1.ListCount - 1 Then
                        
                            ADORecordsetClose rstChecker
                            ADODisconnectDB conHist
                            
                            GoTo Continue
                        Else
                            ADORecordsetClose rstChecker
                            ADODisconnectDB conHist
                        End If
                    Else
                        ADODisconnectDB conHist

                        Exit For
                    End If
                Next lngCtrCtr
            End If
            
            ADORecordsetClose rstChecker
            
            Label1.Caption = "Closing and Opening Stocks..."
            Label1.Refresh
            
                strSQL = "SELECT * FROM TempInbounds" & "_" & Format(m_lngUserID, "00") & " "
            ADORecordsetOpen strSQL, m_conSADBEL, rstInbounds, adOpenKeyset, adLockOptimistic
            With rstInbounds
                .AddNew
                strIn_ID = GenerateID
                Do Until IsIDUnique(conHistory, "In_ID", "Inbounds", Val(strIn_ID))
                   strIn_ID = GenerateID
                Loop
                
                .Fields("In_ID").Value = strIn_ID
                .Fields("In_Code").Value = "<<Closure>>"
                .Fields("In_Orig_Packages_Qty").Value = 1
                .Fields("In_Orig_Gross_Weight").Value = 1
                .Fields("In_Orig_Net_Weight").Value = 1
                .Fields("In_Batch_Num").Value = "Re-opening"
                .Fields("In_Job_Num").Value = "Re-opening"
                .Fields("In_Orig_Packages_Type").Value = "**"
                .Fields("In_Header").Value = 0
                .Fields("In_Detail").Value = 0
                .Fields("In_Avl_Qty_Wgt").Value = 1
                .Fields("Stock_ID").Value = IIf(IsNull(rstStockcards.Fields("ID").Value), 0, rstStockcards.Fields("ID").Value)
                .Fields("InDoc_ID").Value = IIf(strInDoc_ID = "", 0, strInDoc_ID)
                .Update
            End With
            
            InsertRecordset m_conSADBEL, rstInbounds, "TempInbounds" & "_" & Format(m_lngUserID, "00")
            
            ADORecordsetClose rstInbounds
            
            
                strSQL = "SELECT * FROM TempOutbounds" & "_" & Format(m_lngUserID, "00") & " "
            ADORecordsetOpen strSQL, m_conSADBEL, rstOutbounds, adOpenKeyset, adLockOptimistic
            With rstOutbounds
                .AddNew
                strOut_ID = GenerateID
                Do Until IsIDUnique(conHistory, "Out_ID", "Outbounds", Val(strOut_ID))
                    strOut_ID = GenerateID
                Loop
                
                .Fields("Out_ID").Value = strOut_ID
                .Fields("Out_Code").Value = "<<Closure>>"
                .Fields("Out_Header").Value = 0
                .Fields("Out_Detail").Value = 0
                .Fields("In_ID").Value = strIn_ID
                .Fields("Out_Batch_Num").Value = "Closure"
                .Fields("Out_Job_Num").Value = "Closure"
                .Fields("Out_Packages_Qty_Wgt").Value = 1
                .Fields("OutDoc_ID").Value = strOutDoc_ID
                .Update
            End With
            
            InsertRecordset m_conSADBEL, rstOutbounds, "TempOutbounds" & "_" & Format(m_lngUserID, "00")
            
            ADORecordsetClose rstOutbounds
            
            
Continue:
            rstStockcards.MoveNext
            
            DoEvents
            If (blnCancelProcessing = True) Then
                'kapag isasave muna tapos eexit
                If (blnSaveFirstBeforeExiting = True) Then
                    'save here!
                    blnSave = True
                Else
                    blnSave = False
                End If
                Exit Do
            End If
        Loop
        
    Else
        
        If (blnInDocExisting = True Or blnOutDocExisting = True) Then
            MsgBox "A document already exists with the declared document number." & vbCrLf & "Please create another one.", vbInformation + vbOKOnly, "Closure"
        ElseIf blnDateExisting = True Then
            MsgBox "A closure for this date has already been done.", vbInformation + vbOKOnly, "Closure"
        End If
        
    End If
    
    
    'Save!!!
    If (blnSave = True) Then
            strSQL = "INSERT INTO InboundDocs SELECT * FROM TempInboundDocs" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery m_conSADBEL, strSQL
            strSQL = "INSERT INTO Inbounds SELECT * FROM TempInbounds" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery m_conSADBEL, strSQL
            strSQL = "INSERT INTO Outbounds SELECT * FROM TempOutbounds" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery m_conSADBEL, strSQL
            strSQL = "INSERT INTO OutboundDocs SELECT * FROM TempOutboundDocs" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery m_conSADBEL, strSQL
        
        
        ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, strHistoryYear
        'OpenDAODatabase conHistory, m_strPath, "mdb_history" & strHistoryYear & ".mdb"
        
        
        On Error GoTo LinkedTableError
        strSQL = "INSERT INTO InboundDocs SELECT * FROM TempInboundDocs" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery conHistory, strSQL

        strSQL = "INSERT INTO Inbounds SELECT * FROM TempInbounds" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery conHistory, strSQL
        
        strSQL = "INSERT INTO Outbounds SELECT * FROM TempOutbounds" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery conHistory, strSQL

        strSQL = "INSERT INTO OutboundDocs SELECT * FROM TempOutboundDocs" & "_" & Format(m_lngUserID, "00")
        ExecuteNonQuery conHistory, strSQL
        On Error GoTo 0
        
        ADODisconnectDB conHistory

    End If
    
    
EarlyExit:
    
    On Error Resume Next
    ExecuteNonQuery conHistory, "DROP TABLE TempInboundDocs" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery conHistory, "DROP TABLE TempInbounds" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery conHistory, "DROP TABLE TempOutbounds" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery conHistory, "DROP TABLE TempOutboundDocs" & "_" & Format(m_lngUserID, "00")
    
    ExecuteNonQuery m_conSADBEL, "DROP TABLE TempInboundDocs" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE TempInbounds" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE TempOutboundDocs" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE TempOutbounds" & "_" & Format(m_lngUserID, "00")
    On Error GoTo 0
    
    
    If (ProgressBar1.Value = 100 Or ProgressBar1.Value = 0) Then
        fraClosure.Visible = True
        fraProgress.Visible = False
        ProgressBar1.Visible = False
        Me.Height = 5415
        
        cmdProcess(0).Visible = True
        cmdProcess(0).Enabled = False
        
        cmdProcess(1).Visible = True
        cmdProcess(1).Top = 4440
        
        cmdProcess(1).Caption = "&Close"
        TabStrip1.Visible = True
    End If
    
    
    ADORecordsetClose rstStockcards
    
    ADORecordsetClose rstInbounds
    ADORecordsetClose rstInboundDocs
    ADORecordsetClose rstOutbounds
    ADORecordsetClose rstOutboundDocs
    
    ADODisconnectDB conHistory
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
LinkedTableError:
    Dim strTableName As String
    
    If (Err.Number = 3078) Then
        
        'The Microsoft Jet database engine cannot find the input table or query '???'.  Make sure it exists and that its name is spelled correctly.
        If (InStr(1, UCase(Err.Description), "INBOUNDS") > 0) Then
            strTableName = "Inbounds"
        ElseIf (InStr(1, UCase(Err.Description), "INBOUNDDOCS") > 0) Then
            strTableName = "InboundDocs"
        ElseIf (InStr(1, UCase(Err.Description), "OUTBOUNDS") > 0) Then
            strTableName = "Outbounds"
        ElseIf (InStr(1, UCase(Err.Description), "OUTBOUNDDOCS") > 0) Then
            strTableName = "OutboundDocs"
        End If
        
        Err.Clear
        
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "Temp" & strTableName & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, strTableName, strHistoryYear
        'AddLinkedTableEx "Temp" & strTableName & "_" & Format(m_lngUserID, "00") & " ", conHistory.Properties("Data Source Name").Value, G_Main_Password, strTableName, m_conSADBEL.Properties("Data Source Name").Value, G_Main_Password
        
        Resume
        
    Else
        
        MsgBox "An error has occurred." & vbCrLf & "Error (" & Err.Number & "): " & Err.Description, vbInformation, "Closure"
        Err.Clear
        
        GoTo EarlyExit
        
    End If
    
    
End Sub

' Document Number: txtClosure(1).Text
' Entrepot ID:  txtClosure(0).Tag
Private Function CheckIfExisting(ByRef SADBELDB As ADODB.Connection, _
                                ByVal DocumentType As String, _
                                ByVal DocumentNumber As String, _
                                ByVal EntrepotID As Long) _
                                As Boolean
    Dim strSQL As String
    Dim rstChecker As ADODB.Recordset
    
    
    Select Case UCase(DocumentType)
        Case "INBOUNDDOCS"
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "InboundDocs "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "Inbounds "
                strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Stockcards "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "Products "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "Entrepots "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "Products.Entrepot_ID = Entrepots.Entrepot_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Stockcards.Prod_ID = Products.Prod_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Inbounds.Stock_ID = Stockcards.Stock_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "InboundDocs.InDoc_ID = Inbounds.In_ID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "InDoc_Num = '" & DocumentNumber & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Entrepots.Entrepot_ID = " & EntrepotID & " "
                
        Case "OUTBOUNDDOCS"
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "OutboundDocs "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "Outbounds "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "Inbounds "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Stockcards "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "Products "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "Entrepots "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "Products.Entrepot_ID = Entrepots.Entrepot_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Stockcards.Prod_ID = Products.Prod_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Inbounds.Stock_ID = Stockcards.Stock_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Outbounds.In_ID = Inbounds.In_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Outbounds.OutDoc_ID = Outbounddocs.OutDoc_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OutDoc_Num = '" & DocumentNumber & "' "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "Entrepots.Entrepot_ID = " & EntrepotID & " "
    End Select
    
    ADORecordsetOpen strSQL, SADBELDB, rstChecker, adOpenKeyset, adLockOptimistic
    If (rstChecker.BOF And rstChecker.EOF) Then
        CheckIfExisting = False
    Else
        CheckIfExisting = True
    End If
    ADORecordsetClose rstChecker
    
End Function

' Entrepot ID:  txtClosure(0).Tag
' Document Date:    DTPicker1(0).Value
Private Function CheckIfDateExisting(ByRef SADBELDB As ADODB.Connection, _
                                    ByVal DocumentDate As Date, _
                                     ByVal EntrepotID As Long) _
                                     As Boolean
    Dim strSQL As String
    Dim rstChecker As ADODB.Recordset
    
    '    strSQL = "SELECT * FROM OUTBOUNDS INNER JOIN OUTBOUNDDOCS ON OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID WHERE DATEVALUE(OUTDOC_DATE) = DATEVALUE('" & DTPicker1(0).Value & "') AND UCASE(RIGHT(OUT_CODE,11)) = '<<CLOSURE>>'"
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "OUTBOUNDDOCS "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "OUTBOUNDS "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "INBOUNDS "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "STOCKCARDS "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "PRODUCTS "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "ENTREPOTS "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "OUTBOUNDS.IN_ID = INBOUNDS.IN_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "DATEVALUE(OUTDOC_DATE) = DATEVALUE('" & DocumentDate & "') "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "UCASE(RIGHT(OUT_CODE,11)) = '<<CLOSURE>>' "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "ENTREPOTS.ENTREPOT_ID = " & EntrepotID & " "
    ADORecordsetOpen strSQL, m_conSADBEL, rstChecker, adOpenKeyset, adLockOptimistic
    'rstChecker.Open strSQL, m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If rstChecker.BOF And rstChecker.EOF Then
        CheckIfDateExisting = False
    Else
        CheckIfDateExisting = True
    End If
    
    ADORecordsetClose rstChecker
End Function

Private Function IsIDUnique(ByRef ADOConnection As ADODB.Connection, FieldName As String, TableName As String, ID As Long) As Boolean
    Dim rstIDInHistory As ADODB.Recordset
    
    ADORecordsetOpen "SELECT " & FieldName & " FROM " & TableName & " WHERE " & FieldName & " = " & ID, ADOConnection, rstIDInHistory, adOpenKeyset, adLockOptimistic
    'rstIDInHistory.Open "SELECT " & FieldName & " FROM " & TableName & " WHERE " & FieldName & " = " & ID, ADOConnection, adOpenKeyset, adLockOptimistic
    
    If rstIDInHistory.BOF And rstIDInHistory.EOF And ID <> 0 Then
        IsIDUnique = True
    Else
        IsIDUnique = False
    End If
    
    ADORecordsetClose rstIDInHistory
End Function

Private Function GenerateID()
    Randomize
    GenerateID = Round(Rnd * 1000000000, 0)
End Function

Private Sub PopulateGrid()
    Dim rstGrid As ADODB.Recordset
    Dim strSQL As String

    Set rstOfflineRecordset = New ADODB.Recordset
    
    rstOfflineRecordset.CursorLocation = adUseClient
        
    rstOfflineRecordset.Fields.Append "Entrepot Number", adVarChar, 25
    rstOfflineRecordset.Fields.Append "Closure Date", adVarChar, 25
    rstOfflineRecordset.Fields.Append "IM7 for Re-opening", adVarChar, 50 'CSCLP-300 02032009
    rstOfflineRecordset.Fields.Append "Re-Opening Date", adVarChar, 25
    rstOfflineRecordset.Open
        
    strSQL = "SELECT DISTINCT ENTREPOT_TYPE & '-' & ENTREPOT_NUM AS [Entrepot Number], OUTBOUNDDOCS.OUTDOC_DATE AS [Closure Date], INBOUNDDOCS.INDOC_NUM AS [IM7 for Re-opening], INBOUNDDOCS.INDOC_DATE AS [Re-opening Date] FROM ((INBOUNDS INNER JOIN (STOCKCARDS INNER JOIN (PRODUCTS INNER JOIN ENTREPOTS ON PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID) ON STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID)ON INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID) INNER JOIN (OUTBOUNDS INNER JOIN OUTBOUNDDOCS ON OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID) ON INBOUNDS.IN_ID = OUTBOUNDS.IN_ID) INNER JOIN INBOUNDDOCS ON INBOUNDS.INDOC_ID = INBOUNDDOCS.INDOC_ID WHERE UCASE(RIGHT(INBOUNDS.IN_CODE,11)) = '<<CLOSURE>>' ORDER BY OUTDOC_DATE ASC"
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstGrid, adOpenKeyset, adLockOptimistic
    With rstGrid
        
        '.Open strSQL, m_conSADBEL, adOpenKeyset, adLockReadOnly
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
        
            Do While Not .EOF
                rstOfflineRecordset.AddNew
                rstOfflineRecordset.Fields("Entrepot Number").Value = .Fields("Entrepot Number").Value
                rstOfflineRecordset.Fields("Closure Date").Value = CStr(DateValue(.Fields("Closure Date").Value))
                rstOfflineRecordset.Fields("IM7 for Re-opening").Value = .Fields("IM7 for Re-opening").Value
                rstOfflineRecordset.Fields("Re-Opening Date").Value = CStr(DateValue(.Fields("Re-Opening Date").Value))
                rstOfflineRecordset.Update
            
                .MoveNext
            Loop
        End If
        
    End With
    
    ADORecordsetClose rstGrid
    
    If Not (rstOfflineRecordset.BOF And rstOfflineRecordset.EOF) Then
        Set GridEX1.ADORecordset = Nothing
        Set GridEX1.ADORecordset = rstOfflineRecordset
        GridEX1.Columns(1).Width = 1400
        GridEX1.Columns(2).Width = 1200
        GridEX1.Columns(3).Width = 1600
        GridEX1.Columns(4).Width = 1400
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If

End Sub

Private Sub DeleteDoc(ByVal strDocuNum As String, ByVal strInDocDate As String, ByVal strEntrepotNumber As String)
    Dim rstDelete As ADODB.Recordset
    Dim strSQL As String
    Dim strDocID As String
        
    ADOConnectDB m_conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(strInDocDate, 2)
    'OpenADODatabase m_conHistory, m_strPath, "mdb_history" & Right(strInDocDate, 2) & ".mdb"
    
    Me.MousePointer = vbHourglass
    
    strDocID = 0
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT TOP 1 "
        strSQL = strSQL & "OUTBOUNDDOCS.OUTDOC_ID "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
            strSQL = strSQL & "OUTBOUNDS "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "INBOUNDS "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "STOCKCARDS "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "PRODUCTS "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "ENTREPOTS "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "OUTBOUNDS.IN_ID = INBOUNDS.IN_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "OUTBOUNDDOCS "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "OUTBOUNDDOCS.OUTDOC_NUM = '" & strDocuNum & "' "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "RIGHT(OUTBOUNDS.OUT_CODE,11) = '<<CLOSURE>>' "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM = '" & strEntrepotNumber & "' "
    ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
    'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    If Not (rstDelete.BOF And rstDelete.EOF) Then
        rstDelete.MoveFirst
        
        strDocID = rstDelete.Fields("OUTDOC_ID").Value
    End If
    
    ADORecordsetClose rstDelete
    
    'OUBOUNDS-----------------------------------------------------------------
    'SADBEL-------------------------------------------------------------------
    If strDocID <> 0 Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "OUT_ID "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "( "
                strSQL = strSQL & "OUTBOUNDS "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "INBOUNDS "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "STOCKCARDS "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "PRODUCTS "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "ENTREPOTS "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID "
                strSQL = strSQL & ")"
                strSQL = strSQL & "ON "
                strSQL = strSQL & "INBOUNDS.IN_ID = OUTBOUNDS.IN_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = " & Val(strDocID) & " "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM = '" & strEntrepotNumber & "' "
        
        ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
        If Not (rstDelete.BOF And rstDelete.EOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conSADBEL, GetDeleteCommandFromSelect(strSQL, "OUTBOUNDS")
        End If
        
        ADORecordsetClose rstDelete
    End If
    '-------------------------------------------------------------------------
    
    'HISTORY------------------------------------------------------------------
    '<<< Inilabas ang adding ng link table pra may buffer bago gamitin ung created link table
    If (strDocID <> 0) Then
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "STOCKCARDS", Right(strInDocDate, 2)
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "PRODUCTS", Right(strInDocDate, 2)
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, "ENTREPOTS", Right(strInDocDate, 2)
        
        'AddLinkedTableEx "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_history" & Right(strInDocDate, 2) & ".mdb", G_Main_Password, _
                        "STOCKCARDS", m_strPath & "\MDB_SADBEL.mdb", G_Main_Password
        'AddLinkedTableEx "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_history" & Right(strInDocDate, 2) & ".mdb", G_Main_Password, _
                        "PRODUCTS", m_strPath & "\MDB_SADBEL.mdb", G_Main_Password
        'AddLinkedTableEx "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_history" & Right(strInDocDate, 2) & ".mdb", G_Main_Password, _
                        "ENTREPOTS", m_strPath & "\MDB_SADBEL.mdb", G_Main_Password
    End If
    
    If (strDocID <> 0) Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "OUT_ID "
            strSQL = strSQL & "FROM OUTBOUNDS INNER JOIN ("
            strSQL = strSQL & "INBOUNDS INNER JOIN ("
            strSQL = strSQL & "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & " INNER JOIN ("
            strSQL = strSQL & "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & " INNER JOIN MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & " "
            strSQL = strSQL & "ON MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID = MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID) "
            strSQL = strSQL & "ON MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID = MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID) "
            strSQL = strSQL & "ON INBOUNDS.STOCK_ID = MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".STOCK_ID)"
            strSQL = strSQL & "ON INBOUNDS.IN_ID = OUTBOUNDS.IN_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = " & Val(strDocID) & " "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_TYPE & '-' & MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_NUM = '" & strEntrepotNumber & "'"
        
        On Error GoTo LinkedTableError
        ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If Not (rstDelete.BOF And rstDelete.EOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conSADBEL, GetDeleteCommandFromSelect(strSQL, "OUTBOUNDS")
        End If
        
        ADORecordsetClose rstDelete
    End If
    
    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------
    
    'OUBOUNDDOCS--------------------------------------------------------------
    'SADBEL-------------------------------------------------------------------
    If strDocID <> 0 Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "( "
                strSQL = strSQL & "OUTBOUNDS "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "INBOUNDS "
                    strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "STOCKCARDS "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "( "
                                strSQL = strSQL & "PRODUCTS "
                                strSQL = strSQL & "INNER JOIN "
                                strSQL = strSQL & "ENTREPOTS "
                                strSQL = strSQL & "ON "
                                strSQL = strSQL & "PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID "
                            strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "INBOUNDS.IN_ID = OUTBOUNDS.IN_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "OUTBOUNDDOCS "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OUTBOUNDDOCS.OUTDOC_ID = " & Val(strDocID) & " "
            strSQL = strSQL & "AND ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM <> '" & strEntrepotNumber & "' "
        ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
        If rstDelete.BOF And rstDelete.EOF Then
            ADORecordsetClose rstDelete

            strSQL = "SELECT * FROM OUTBOUNDDOCS WHERE OUTDOC_ID = " & Val(strDocID)
            ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
            'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
            If Not (rstDelete.EOF And rstDelete.BOF) Then
                rstDelete.MoveFirst
                Do While Not rstDelete.EOF
                    rstDelete.Delete
                    rstDelete.Update
                    rstDelete.MoveNext
                Loop
                
                ' TO DO FOR CP.NET
                ExecuteNonQuery m_conSADBEL, GetDeleteCommandFromSelect(strSQL, "OUTBOUNDS")
            End If
        End If
        
        ADORecordsetClose rstDelete
    End If
    '-------------------------------------------------------------------------
    
    'HISTORY------------------------------------------------------------------
    If (strDocID <> 0) Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "( "
                strSQL = strSQL & "OUTBOUNDS "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "INBOUNDS "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & " "
                        strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "( "
                                strSQL = strSQL & "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & " "
                                strSQL = strSQL & "INNER JOIN "
                                strSQL = strSQL & "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & " "
                                strSQL = strSQL & "ON "
                                strSQL = strSQL & "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID = MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID "
                            strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID = MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "INBOUNDS.STOCK_ID = MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".STOCK_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "INBOUNDS.IN_ID = OUTBOUNDS.IN_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "OUTBOUNDDOCS "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OUTBOUNDDOCS.OUTDOC_ID = " & Val(strDocID) & " "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_TYPE & '-' & MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_NUM <> '" & strEntrepotNumber & "' "
        
        On Error GoTo LinkedTableError
        ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        If rstDelete.BOF And rstDelete.EOF Then
            ADORecordsetClose rstDelete
            
            strSQL = "SELECT * FROM OUTBOUNDDOCS WHERE OUTDOC_ID = " & Val(strDocID)
            ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
            'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
            If Not (rstDelete.EOF And rstDelete.BOF) Then
                rstDelete.MoveFirst
                
                Do While Not rstDelete.EOF
                    rstDelete.Delete
                    rstDelete.Update
                    rstDelete.MoveNext
                Loop
                
                ' TO DO FOR CP.NET
                ExecuteNonQuery m_conHistory, GetDeleteCommandFromSelect(strSQL, "OUTBOUNDS")
            End If
        End If
        
        ADORecordsetClose rstDelete
    End If
    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------

    
    strDocID = 0
    
    strSQL = "SELECT TOP 1 INBOUNDDOCS.INDOC_ID FROM (INBOUNDS INNER JOIN (STOCKCARDS INNER JOIN (PRODUCTS INNER JOIN ENTREPOTS ON PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID) ON STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID) ON INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID) INNER JOIN INBOUNDDOCS ON INBOUNDS.INDOC_ID = INBOUNDDOCS.INDOC_ID WHERE INDOC_NUM = '" & strDocuNum & "' AND RIGHT(INBOUNDS.IN_CODE,11) = '<<CLOSURE>>' AND ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM ='" & strEntrepotNumber & "'"
    ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
    'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    If Not (rstDelete.BOF And rstDelete.EOF) Then
        rstDelete.MoveFirst
        
        strDocID = rstDelete.Fields("INDOC_ID").Value
    End If
    ADORecordsetClose rstDelete
    
    'INBOUNDS-----------------------------------------------------------------
    'SADBEL-------------------------------------------------------------------
    If strDocID <> 0 Then
        strSQL = "SELECT IN_ID FROM (INBOUNDS INNER JOIN (STOCKCARDS INNER JOIN (PRODUCTS INNER JOIN ENTREPOTS ON PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID) ON STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID) ON INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID) WHERE INBOUNDS.INDOC_ID = " & Val(strDocID) & " AND ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM = '" & strEntrepotNumber & "'"
        ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
        If Not (rstDelete.BOF And rstDelete.EOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conSADBEL, GetDeleteCommandFromSelect(strSQL, "INBOUNDS")
        End If
        
        ADORecordsetClose rstDelete
    End If
    '-------------------------------------------------------------------------
    
    'HISTORY------------------------------------------------------------------
    If (strDocID <> 0) Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "IN_ID "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "INBOUNDS "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & " "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & " "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID = MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID = MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & ".PROD_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "INBOUNDS.STOCK_ID = MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & ".STOCK_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "INBOUNDS.INDOC_ID = " & Val(strDocID) & " "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_TYPE & '-' & MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & ".ENTREPOT_NUM = '" & strEntrepotNumber & "' "
        
        On Error GoTo LinkedTableError
        ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If Not (rstDelete.BOF And rstDelete.EOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conHistory, GetDeleteCommandFromSelect(strSQL, "INBOUNDS")
        End If
        ADORecordsetClose rstDelete
    End If
    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------
    
    '<<< dandan 112806
    '<<< Drop link tables
    On Error Resume Next
    ExecuteNonQuery m_conHistory, "DROP TABLE MDB_STOCKCARDS" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conHistory, "DROP TABLE MDB_PRODUCTS" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery m_conHistory, "DROP TABLE MDB_ENTREPOTS" & "_" & Format(m_lngUserID, "00") & " "
    On Error GoTo 0
    
    
    'INBOUNDDOCS--------------------------------------------------------------
    'SADBEL-------------------------------------------------------------------
    strSQL = vbNullString
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM INBOUNDS "
    strSQL = strSQL & "INNER JOIN INBOUNDDOCS "
    strSQL = strSQL & "ON INBOUNDS.INDOC_ID = INBOUNDDOCS.INDOC_ID "
    strSQL = strSQL & "WHERE INBOUNDDOCS.INDOC_ID = " & Val(strDocID)
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
    'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    If rstDelete.BOF And rstDelete.EOF Then
        
        ADORecordsetClose rstDelete
        
        strSQL = "SELECT * FROM INBOUNDDOCS WHERE INBOUNDDOCS.INDOC_ID = " & Val(strDocID)
        ADORecordsetOpen strSQL, m_conSADBEL, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
        If Not (rstDelete.EOF And rstDelete.BOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conSADBEL, GetDeleteCommandFromSelect(strSQL, "INBOUNDDOCS")
        End If
    End If
    
    ADORecordsetClose rstDelete
    '-------------------------------------------------------------------------
    'HISTORY------------------------------------------------------------------
    strSQL = vbNullString
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM INBOUNDS "
    strSQL = strSQL & "INNER JOIN INBOUNDDOCS "
    strSQL = strSQL & "ON INBOUNDS.INDOC_ID = INBOUNDDOCS.INDOC_ID "
    strSQL = strSQL & "WHERE INBOUNDDOCS.INDOC_ID = " & Val(strDocID)
    
    ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
    'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
    If rstDelete.BOF And rstDelete.EOF Then
    
        ADORecordsetClose rstDelete
        strSQL = "SELECT * FROM INBOUNDDOCS WHERE INBOUNDDOCS.INDOC_ID = " & Val(strDocID)
        
        ADORecordsetOpen strSQL, m_conHistory, rstDelete, adOpenKeyset, adLockOptimistic
        'rstDelete.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
        If Not (rstDelete.EOF And rstDelete.BOF) Then
            rstDelete.MoveFirst
            Do While Not rstDelete.EOF
                rstDelete.Delete
                rstDelete.Update
                rstDelete.MoveNext
            Loop
            
            ' TO DO FOR CP.NET
            ExecuteNonQuery m_conHistory, GetDeleteCommandFromSelect(strSQL, "INBOUNDDOCS")
        End If
    End If
    ADORecordsetClose rstDelete

    ADODisconnectDB m_conHistory
    
    Me.MousePointer = vbDefault

    
    Exit Sub
    
LinkedTableError:
    Dim strTableName As String
    
    If (Err.Number = 3078) Or (Err.Number = -2147217865) Then
        
        'The Microsoft Jet database engine cannot find the input table or query '???'.  Make sure it exists and that its name is spelled correctly.
        If (InStr(1, UCase(Err.Description), "PRODUCTS") > 0) Then
            strTableName = "PRODUCTS"
        ElseIf (InStr(1, UCase(Err.Description), "ENTREPOTS") > 0) Then
            strTableName = "ENTREPOTS"
        ElseIf (InStr(1, UCase(Err.Description), "STOCKCARDS") > 0) Then
            strTableName = "STOCKCARDS"
        End If
        
        Err.Clear
            
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, "MDB_" & strTableName & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_SADBEL, strTableName, Right(strInDocDate, 2)
        'AddLinkedTableEx "MDB_" & strTableName & "_" & Format(m_lngUserID, "00"), m_strPath & "\mdb_history" & Right(strInDocDate, 2) & ".mdb", G_Main_Password, strTableName, m_strPath & "\MDB_SADBEL.mdb", G_Main_Password
        
        Resume
        
    Else
        
        MsgBox "An error has occurred." & vbCrLf & "Error (" & Err.Number & "): " & Err.Description, vbInformation, "Closure"
        Err.Clear
        
    End If
    
End Sub
