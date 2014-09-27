VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmManualOutbound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Outbound"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmManualOutbound.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   120
      Pattern         =   "mdb_history*.mdb"
      System          =   -1  'True
      TabIndex        =   36
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame FraOutbounds 
      Caption         =   "Outbounds"
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   11415
      Begin GridEX16.GridEX jgxOutbounds 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3836
         MethodHoldFields=   -1  'True
         Options         =   -1
         RecordsetType   =   1
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         ColumnCount     =   9
         CardCaption1    =   -1  'True
         ColCaption1     =   "Entrepot No"
         ColKey1         =   "Entrepot No"
         ColWidth1       =   1545
         ColSelectable2  =   0   'False
         ColCaption2     =   "Product No"
         ColKey2         =   "Product No"
         ColWidth2       =   1545
         ColCaption3     =   "Stock Card No"
         ColKey3         =   "Stock Card No"
         ColWidth3       =   1545
         ColSortType3    =   2
         ColCaption4     =   "Document No"
         ColKey4         =   "Document No"
         ColVisible4     =   0   'False
         ColWidth4       =   1395
         ColCaption5     =   "Quantity/Weight"
         ColKey5         =   "Quantity/Weight"
         ColVisible5     =   0   'False
         ColWidth5       =   1395
         ColSortType5    =   2
         ColCaption6     =   "Outbound Qty"
         ColKey6         =   "Outbound Qty"
         ColWidth6       =   1545
         ColCaption7     =   "Package Type"
         ColKey7         =   "Package Type"
         ColWidth7       =   1545
         ColCaption8     =   "Job No"
         ColKey8         =   "Job No"
         ColWidth8       =   1545
         ColCaption9     =   "Batch No"
         ColKey9         =   "Batch No"
         ColWidth9       =   1545
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Frame fraDocument 
      Caption         =   "Document"
      Height          =   2175
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboDocType 
         Height          =   315
         ItemData        =   "frmManualOutbound.frx":08CA
         Left            =   1680
         List            =   "frmManualOutbound.frx":08FB
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboCodiType 
         Height          =   315
         ItemData        =   "frmManualOutbound.frx":094A
         Left            =   1680
         List            =   "frmManualOutbound.frx":0966
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58720257
         CurrentDate     =   38350
      End
      Begin VB.TextBox txtDocNum 
         Height          =   315
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtMRN 
         Height          =   315
         Left            =   1680
         MaxLength       =   18
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtCommunalSettlement 
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblCodiType 
         Caption         =   "Codisheet Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblDocType 
         Caption         =   "Doc Type && Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   675
         Width           =   1815
      End
      Begin VB.Label lblMRN 
         Caption         =   "MRN:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label lblDocumentDate 
         Caption         =   "Document Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1350
         Width           =   1815
      End
      Begin VB.Label lblCommunalSettlement 
         Caption         =   "Communal Settlement:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1710
         Width           =   1815
      End
   End
   Begin VB.Frame fraProduct 
      Caption         =   "Product Information"
      Height          =   2175
      Left            =   4080
      TabIndex        =   16
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtProductNum 
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   240
         Width           =   1605
      End
      Begin VB.CommandButton cmdProductPicklist 
         Caption         =   "..."
         Height          =   315
         Left            =   3285
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblProdDesc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   1680
         TabIndex        =   32
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label lblTARICCode 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblCtryExportDesc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5760
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCtryOriginDesc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5760
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCtryExportCode 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblCtryOriginCode 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5280
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblProduct 
         Caption         =   "Product Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label lblTaric 
         Caption         =   "Taric Code:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblCtryOrigin 
         Caption         =   "Country of Origin:"
         Height          =   195
         Left            =   3720
         TabIndex        =   18
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label lblCtryExport 
         Caption         =   "Country of Export:"
         Height          =   195
         Left            =   3720
         TabIndex        =   17
         Top             =   630
         Width           =   1575
      End
   End
   Begin VB.Frame fraAvailableStock 
      Caption         =   "Available Stock"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   11415
      Begin VB.CheckBox chkShowZero 
         Caption         =   "Display &zero balance stocks"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddToOutbounds 
         Caption         =   "A&dd To Outbounds"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   10
         Top             =   2760
         Width           =   1575
      End
      Begin GridEX16.GridEX jgxAvailableStock 
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3836
         MethodHoldFields=   -1  'True
         Options         =   -1
         RecordsetType   =   1
         GroupByBoxVisible=   0   'False
         ColumnCount     =   9
         CardCaption1    =   -1  'True
         ColCaption1     =   "Entrepot No"
         ColKey1         =   "Entrepot No"
         ColWidth1       =   1200
         ColSelectable2  =   0   'False
         ColCaption2     =   "Product No"
         ColKey2         =   "Product No"
         ColWidth2       =   1200
         ColCaption3     =   "Stock Card No"
         ColKey3         =   "Stock Card No"
         ColWidth3       =   1200
         ColSortType3    =   2
         ColCaption4     =   "Document No"
         ColKey4         =   "Document No"
         ColWidth4       =   1200
         ColCaption5     =   "Quantity/Weight"
         ColKey5         =   "Quantity/Weight"
         ColWidth5       =   1305
         ColSortType5    =   2
         ColCaption6     =   "Package Type"
         ColKey6         =   "Package Type"
         ColWidth6       =   1200
         ColCaption7     =   "Batch No"
         ColKey7         =   "Batch No"
         ColWidth7       =   1245
         ColCaption8     =   "Qty for Outbound"
         ColKey8         =   "Qty for Outbound"
         ColWidth8       =   1350
         ColSortType8    =   2
         ColCaption9     =   "Job No"
         ColKey9         =   "Job No"
         ColWidth9       =   1200
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Press Enter or click Add To Outbounds to add stock of the currently selected row to this manual outbound movement."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      Top             =   8400
      Width           =   1215
   End
End
Attribute VB_Name = "frmManualOutbound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private m_conSADBEL As ADODB.Connection
Private m_conHistory As ADODB.Connection
Private m_conTaric As ADODB.Connection

Private m_rstAvailableOff As ADODB.Recordset
Private m_rstOutboundsOff As ADODB.Recordset
Private m_rstSadbel As ADODB.Recordset
Private m_rstHistory As ADODB.Recordset
    
    
Private m_lngUserID As Long
Private m_dteLastDate As Date

Private strProdID As String
Private lngOutDocID As Long
Private lngHandling As Long

Private strYear As String
Private strLanguage As String
Private intTaricProperties As Integer
Private lngResourceHandler As Long

Private blnSystemChanged As Boolean
Private blnFormLoaded As Boolean
Private blnOutboundsChanged As Boolean
Private blnAvailableIsLastFocus As Boolean
Private blnAvailableHasFocus As Boolean

Private alngDeleted() As Long

Private strLastCodiType As String
Private strLastDocType As String
Private strLastDocNum As String
Private strLastMRN As String
Private strLastProdNum As String

Private strCorrectionMode As String

Private m_astrHistoryDBs() As String

Private Sub cboCodiType_Click()
    
    cboDocType.Clear
    Select Case cboCodiType.ListIndex
        Case 0, 1, 2, 3, 4, 5
            cboDocType.AddItem "IM0"
            cboDocType.AddItem "IM4"
            cboDocType.AddItem "IM5"
            cboDocType.AddItem "IM7"
            cboDocType.AddItem "IM9"
            cboDocType.AddItem "EU0"
            cboDocType.AddItem "EU4"
            cboDocType.AddItem "EU5"
            cboDocType.AddItem "E30"
            cboDocType.AddItem "E31"
            cboDocType.AddItem "E32"
            cboDocType.AddItem "U30"
            cboDocType.AddItem "U31"
            cboDocType.AddItem "U32"
            cboDocType.AddItem "001"
            
            txtDocNum.MaxLength = 7
            
        Case 6
            cboDocType.AddItem "IMH"
            cboDocType.AddItem "IMI"
            cboDocType.AddItem "IMJ"
            cboDocType.AddItem "IMK"
            
            txtDocNum.MaxLength = 25
            
        Case 7
            cboDocType.AddItem "EXA"
            cboDocType.AddItem "EXB"
            cboDocType.AddItem "EXC"
            cboDocType.AddItem "EXD"
            cboDocType.AddItem "EXE"
            cboDocType.AddItem "EXF"
            cboDocType.AddItem "EXG"
            
            txtDocNum.MaxLength = 25
    End Select
    
    cboDocType.ListIndex = -1
    
    EnableDisableOutDocFields (cboCodiType.ListIndex)
    
    If strLastCodiType <> cboCodiType.Text Then
        Call CheckOutDoc
        strLastCodiType = cboCodiType.Text
    End If
End Sub

Private Sub EnableDisableOutDocFields(Index As Long)
    Dim lngCtr As Long
    
    'Locks doc type as IM7 when entering as Inbound Correction
    If strCorrectionMode = "I" Then
        'Only document type IM7 will be available so no need to lock control anymore
        cboDocType.Clear
        Select Case Index
            Case 6, 7
                cboDocType.AddItem "IMJ"
                cboDocType.AddItem "IMK"

            Case Else
                cboDocType.AddItem "IM7"
        End Select
        
        cboDocType.ListIndex = 0
        'Can't have zero valued initial stock anyway so this control is useless when in Inbounds Correction mode
        chkShowZero.Enabled = False
        chkShowZero.Value = vbChecked
    Else
    
Check_Again:
        'Removes IM7 document type since it's Inbound exclusive
        For lngCtr = 0 To cboDocType.ListCount - 1
            Select Case Index
                Case 6
                    If cboDocType.List(lngCtr) = "IMJ" Or _
                        cboDocType.List(lngCtr) = "IMK" Then
                        cboDocType.RemoveItem lngCtr
                        
                        GoTo Check_Again
                    End If
    
                Case Else
                    If cboDocType.List(lngCtr) = "IM7" Then
                        cboDocType.RemoveItem lngCtr
                        Exit For
                    End If
            End Select
        Next
        
        If strCorrectionMode = "O" Then
            chkShowZero.Visible = True
            chkShowZero.Value = vbUnchecked
        Else
            'Show zero stocks checkbox is hidden in manual outbound mode, cuz zero display is for corrections only
            chkShowZero.Visible = False
            chkShowZero.Value = vbUnchecked
        End If
    End If
    
    Select Case Index
        Case 0, 1, 2
            txtMRN.Text = ""
            txtMRN.Enabled = False
            txtDocNum.Enabled = True
            cboDocType.Enabled = True
        Case 3, 5
            txtDocNum.Text = ""
            txtDocNum.Enabled = False
            cboDocType.ListIndex = -1
            cboDocType.Enabled = False
            txtMRN.Enabled = True
        Case 4
            txtDocNum.Enabled = True
            cboDocType.Enabled = True
            txtMRN.Enabled = True
            
        Case 6, 7   ' PLDA Import, PLDA Export
            txtDocNum.Enabled = True
            cboDocType.Enabled = True
            txtMRN.Text = ""
            txtMRN.Enabled = False
    End Select
End Sub

Private Sub chkShowZero_Click()
    'Allows user to display stocks that have been zeroed out
    If Len(Trim$(strProdID)) > 0 Then
        If chkShowZero.Value = Unchecked Then
            m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR ([Qty for Outbound] <> 0 AND [Qty for Outbound] <> '') OR [Outbound Edits] <> 0 "
        Else
            m_rstAvailableOff.Filter = adFilterNone
        End If
        
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        Call FormatAvailable
    End If
End Sub

Private Sub cmdAddToOutbounds_Click()
    
    If jgxAvailableStock.Row > 0 Then
        jgxAvailableStock.Update
        Call AddRecordToOutbound
    End If
    
End Sub

Private Sub cmdApply_Click()
    
    If Not ValidateOutbounds Then
        Exit Sub
    End If
    
    'RACHELLE 092705: add checking for closure date
    If IsStockClosed = True Then
        MsgBox "A closure has already been done that includes this document.", vbInformation + vbOKOnly, "Initial Stock"
        Exit Sub
    End If
    
    Call ApplyChanges
    
    cmdApply.Enabled = False
    blnOutboundsChanged = False
    
End Sub

Private Sub GetOutDocID()

    Dim rstTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strFilter As String
    Dim bytCounter As Byte
    Dim bytNumberofZeros As Byte
    Dim strZeros As String
    
    If Len(Trim(txtDocNum.Text)) < 7 And Len(Trim(txtDocNum.Text)) <> 0 Then
        bytNumberofZeros = 7 - Len(txtDocNum.Text)
        For bytCounter = 1 To bytNumberofZeros
            strZeros = "0" & strZeros
        Next bytCounter
    End If
    
        strSQL = ""
        strSQL = "SELECT OutDoc_ID, OutDoc_Type, OutDoc_Num, OutDoc_MRN, OutDoc_Comm_Settlement, OutDoc_Date FROM OutboundDocs WHERE "
        strSQL = strSQL & IIf(cboDocType.Text <> "", "OutDoc_Type = '" & cboDocType.Text & "'", "(ISNULL(OutDoc_Type) OR OutDoc_Type='')") & " AND "
        strSQL = strSQL & IIf(Trim$(txtDocNum.Text) <> "", "OutDoc_Num = '" & strZeros & txtDocNum.Text & "'", "(ISNULL(OutDoc_Num) OR OutDoc_Num='')") & " AND "
        strSQL = strSQL & IIf(txtMRN.Text <> "", "OutDoc_MRN = '" & txtMRN.Text & "'", "(ISNULL(OutDoc_MRN) OR OutDoc_MRN='')") & " AND "
        strSQL = strSQL & "FORMAT(OutDoc_Date, 'Short Date') = #" & Format(dtpDate.Value, "Short Date") & "# AND "
        strSQL = strSQL & "OutDoc_Global=1"
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
                
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.MoveFirst
        
        lngOutDocID = rstTemp!OutDoc_ID
        txtCommunalSettlement.Text = IIf(IsNull(rstTemp!OutDoc_Comm_Settlement), "", rstTemp!OutDoc_Comm_Settlement)
        strYear = Year(rstTemp!OutDoc_Date)
        dtpDate.Value = rstTemp!OutDoc_Date
        cmdApply.Enabled = False
    Else
        lngOutDocID = 0
    End If
    
    ADORecordsetClose rstTemp
End Sub

Private Sub CheckOutDoc()
                
    Dim bytCounter As Byte
    Dim bytNumberofZeros As Byte
    Dim strZeros As String
    
    If Len(Trim(txtDocNum.Text)) < 7 And Len(Trim(txtDocNum.Text)) <> 0 Then
        bytNumberofZeros = 7 - Len(txtDocNum.Text)
        For bytCounter = 1 To bytNumberofZeros
            strZeros = "0" & strZeros
        Next bytCounter
    End If
    
    strLastDocNum = strZeros & txtDocNum.Text
    strLastDocType = cboDocType.Text
    strLastMRN = txtMRN.Text
    m_dteLastDate = dtpDate.Value
    
    Select Case cboCodiType.ListIndex
        Case 0, 1, 2
            If Len(Trim$(txtDocNum.Text)) > 0 And Len(Trim$(cboDocType.Text)) > 0 Then
                Call GetOutDocID
            Else
                lngOutDocID = 0
            End If
        Case 3, 5
            If Len(Trim$(txtMRN.Text)) > 0 Then
                Call GetOutDocID
            Else
                lngOutDocID = 0
            End If
        Case 4
            If Len(Trim$(txtDocNum.Text)) > 0 And Len(Trim$(cboDocType.Text)) > 0 And Len(Trim$(txtMRN.Text)) > 0 Then
                Call GetOutDocID
            Else
                lngOutDocID = 0
            End If
        
        Case 6, 7
            If Len(Trim$(txtDocNum.Text)) > 0 And Len(Trim$(cboDocType.Text)) > 0 Then
                Call GetOutDocID
            Else
                lngOutDocID = 0
            End If
    End Select
        
    If lngOutDocID <> 0 Then
        
        blnSystemChanged = True
        txtProductNum.Text = ""
        Call ResetProduct
        blnSystemChanged = False
        
        Call CreateOutboundsOffline
        Call CreateAvailableOffline
    
        Call PopulateOutbounds
        'Checks if user wanted to display zero stocks
        Call PopulateAvailable
        
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not ValidateOutbounds Then
        Exit Sub
    End If
    
    'RACHELLE 092705: add checking for closure date
    If IsStockClosed = True Then
        MsgBox "A closure has already been done that includes this document.", vbInformation + vbOKOnly, "Initial Stock"
        Exit Sub
    End If
    
   If cmdApply.Enabled = True Then Call ApplyChanges
    
    Unload Me
    
End Sub

Private Sub cmdProductPicklist_Click()
    Dim clsProducts As cProducts
    Set clsProducts = New cProducts
    
    With clsProducts
        .Product_Num = txtProductNum.Text
        .ShowProducts 1, Me, m_conSADBEL, m_conTaric, strLanguage, intTaricProperties, lngResourceHandler, txtProductNum.Text
        
        'Prevents updating of controls when Product selection has been cancelled.
        If clsProducts.Cancelled = False Then
            
            If cmdApply.Enabled Then
                
                m_rstAvailableOff.Filter = adFilterNone
                m_rstAvailableOff.Filter = "[Qty for Outbound] > 0 OR [Outbound Edits] <> 0"
                m_rstOutboundsOff.Filter = "Out_ID = 0"
                If m_rstAvailableOff.RecordCount > 0 Or m_rstOutboundsOff.RecordCount > 0 Then
                    If MsgBox("Save pending stock for outbound?", vbYesNo + vbQuestion, "Manual Outbound") = vbYes Then
                        Call cmdApply_Click
                        If cmdApply.Enabled Then
                            Set clsProducts = Nothing
                            Exit Sub
                        End If
                    End If
                    ReDim alngDeleted(0)
                End If
                m_rstOutboundsOff.Filter = adFilterNone
                m_rstAvailableOff.Filter = adFilterNone
                
            End If
                                    
            strProdID = IIf(CStr(.Product_ID) = "", "0", .Product_ID)
            
            If lngOutDocID <> 0 Then
                Call CreateOutboundsOffline
                Call PopulateOutbounds(strProdID)
            End If
            
            Call CreateAvailableOffline
            'Checks if user wanted to display zero stocks
            Call PopulateAvailable
                
            blnSystemChanged = True
            txtProductNum.Text = .Product_Num
            strLastProdNum = txtProductNum.Text
            blnSystemChanged = False
            
'            txtProductNum.Tag = .Product_ID
            lblProdDesc.Caption = .Prod_Desc
            lblTARICCode.Caption = .Taric_Code
            lblCtryOriginCode.Caption = .Ctry_Origin
            lblCtryExportCode.Caption = .Ctry_Export
            lblCtryOriginDesc.Caption = .Origin_Desc
            lblCtryExportDesc.Caption = .Export_Desc
            lngHandling = .Prod_Handling
            
            If chkShowZero.Value = Unchecked Then
                m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Outbound Edits] <> 0"
            Else
                m_rstAvailableOff.Filter = adFilterNone
            End If

            If m_rstAvailableOff.RecordCount > 0 Or jgxAvailableStock.RowCount > 0 Then
                Set jgxAvailableStock.ADORecordset = Nothing
                Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
                Call FormatAvailable
            End If
            
            If m_rstOutboundsOff.RecordCount > 0 Or jgxOutbounds.RowCount > 0 Then
                Set jgxOutbounds.ADORecordset = Nothing
                Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
                Call FormatOutbounds
            End If
            
        End If
    End With

    Set clsProducts = Nothing
End Sub

Private Sub dtpDate_Change()
    If Format(dtpDate.Value, "Short Date") <> Format(m_dteLastDate, "Short Date") Then
        If lngOutDocID <> 0 Then
            If Not SavedIfChanged Then
                txtDocNum.Text = strLastDocNum
                Exit Sub
            End If
            
            blnSystemChanged = True
            txtProductNum.Text = ""
            Call ResetProduct
            blnSystemChanged = False
                        
            strLastDocNum = ""
            strLastDocType = ""
            strLastMRN = ""
            m_dteLastDate = 0
    
            Set jgxAvailableStock.ADORecordset = Nothing
            Call CreateAvailableOffline
            Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
            Call FormatAvailable
    
            Set jgxOutbounds.ADORecordset = Nothing
            Call CreateOutboundsOffline
            Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
            Call FormatOutbounds
        
        End If
        
        Call CheckOutDoc
        
    End If
    
    CheckYear
    If jgxAvailableStock.RowCount > 0 Or jgxOutbounds.RowCount > 0 Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub dtpDate_DropDown()
    CheckYear
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckYear
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    CheckYear
End Sub

Private Sub Form_Activate()
    blnFormLoaded = True
    strLastCodiType = "Import"
    cboCodiType.Text = "Import"
End Sub

Private Sub Form_Load()
    
    CreateLinkedTables
    
    ADOCloseOpenDB m_conSADBEL
    
    ReDim alngDeleted(0)
    dtpDate.Value = Date
    
    Call CreateAvailableOffline
    Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
    Call FormatAvailable
    
    Call CreateOutboundsOffline
    Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
    Call FormatOutbounds
    
    Call CreateInboundHistory
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    DeleteLinkedTempTables
    
    
    Set jgxAvailableStock.ADORecordset = Nothing
    Set jgxOutbounds.ADORecordset = Nothing
    
    ADORecordsetClose m_rstAvailableOff
    ADORecordsetClose m_rstOutboundsOff
    
    Erase alngDeleted
    
End Sub

Private Sub jgxAvailableStock_AfterColUpdate(ByVal ColIndex As Integer)
    Dim lngrow As Long
    Dim lngIndex As Long
    Dim lngSort As Long
    
    If ColIndex = jgxAvailableStock.Columns("Qty for Outbound").Index Then
        lngrow = jgxAvailableStock.Row
        
        If chkShowZero.Value = vbUnchecked Then
'            m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Qty for Outbound] > 0"
            m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR ([Qty for Outbound] <> 0 AND [Qty for Outbound] <> '') OR [Outbound Edits] <> 0"
        End If
        
        If jgxAvailableStock.SortKeys.Count > 0 Then
            lngIndex = jgxAvailableStock.SortKeys(1).ColIndex
            lngSort = jgxAvailableStock.SortKeys(1).SortOrder
        End If
        
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        
        Call FormatAvailable
        
        If lngIndex > 0 Then
            jgxAvailableStock.SortKeys.Add lngIndex, lngSort
            jgxAvailableStock.RefreshSort
            lngIndex = 0
        End If
        
        jgxAvailableStock.Row = lngrow
        jgxAvailableStock.Col = ColIndex
    End If
End Sub

Private Sub jgxAvailableStock_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX16.JSRetBoolean)
    Dim strSQL As String
    Dim lngCtr As Long
    
    Dim strHistoryDBYear As String
    
    
    
    'If chkShowZero.Value = vbChecked Or strCorrectionMode = "I" Or jgxAvailableStock.Value(jgxAvailableStock.Columns("Quantity/Weight").Index) < 0 then
    '-------------------------------------------------------
    'CP v3.53
    'October 20, 2005
    If chkShowZero.Value = vbChecked Or strCorrectionMode = "I" Or (jgxAvailableStock.Value(jgxAvailableStock.Columns("Quantity/Weight").Index) + jgxAvailableStock.Value(jgxAvailableStock.Columns("Outbound Edits").Index)) < 0 Then
    '-------------------------------------------------------
        'Connects to mdb_history## instead when showzerostocks is checked or Inbounds Correction mode
        For lngCtr = 0 To File1.ListCount - 1
            strHistoryDBYear = Replace(File1.List(lngCtr), "mdb_history", vbNullString)
            strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
            
            If (UCase(Trim(strCorrectionMode)) = "I") Then
                'Inbounds Correction validation
                strSQL = "SELECT " & IIf((strCorrectionMode = "I" And InStr(jgxAvailableStock.Value(jgxAvailableStock.Columns("Qty for Outbound").Index), "-") > 0), _
                              "CHOOSE(Products!Prod_Handling + 1, Inbounds!In_Orig_Packages_Qty, Inbounds!In_Orig_Gross_Weight, Inbounds!In_Orig_Net_Weight)", _
                              "Inbounds!In_Avl_Qty_Wgt") & " AS [Quantity/Weight] " & _
                              "FROM Inbounds " & _
                              "INNER JOIN (StockCards INNER JOIN Products " & _
                              "ON Stockcards.Prod_ID = Products.Prod_ID) " & _
                              "ON Inbounds.Stock_ID = Stockcards.Stock_ID " & _
                              "WHERE Inbounds!In_ID = " & m_rstAvailableOff!In_ID
            Else
                strSQL = "SELECT In_TotalOut_Qty_Wgt, In_Avl_Qty_Wgt FROM Inbounds WHERE Inbounds!In_ID = " & m_rstAvailableOff!In_ID
            End If
            
            strSQL = Replace(strSQL, "Inbounds", "InboundHistory" & strHistoryDBYear & "_" & Format(m_lngUserID, "00"))
            
            
            ADORecordsetOpen strSQL, m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
            If Not (m_rstSadbel.BOF And m_rstSadbel.EOF) Then
                Exit For
            End If
        Next lngCtr
    End If
                
    With jgxAvailableStock
        If ColIndex = .Columns("Qty for Outbound").Index Then
            If InStr(.Value(.Columns("Qty for Outbound").Index), "-") > 0 Then
                'If user entered a hyphen, corrects placement of negative sign
                If InStr(.Value(.Columns("Qty for Outbound").Index), "-") > 1 Then
                    .Value(.Columns("Qty for Outbound").Index) = "-" & Replace$(.Value(.Columns("Qty for Outbound").Index), "-", Empty)
                End If
                
                If strCorrectionMode = "I" Then

                    .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                    If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                        .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                    End If
                Else
                    
                    ADORecordsetOpen "SELECT In_TotalOut_Qty_Wgt, In_Avl_Qty_Wgt FROM Inbounds WHERE Inbounds!In_ID = " & m_rstAvailableOff!In_ID, m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
                    
                    If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                        m_rstSadbel.MoveFirst
                        
                        If m_rstSadbel.Fields("In_TotalOut_Qty_Wgt").Value = 0 Then
                            MsgBox "There are no remaining units left to correct.", vbInformation + vbOKOnly, Me.Caption
                            Cancel = True
                            
                        '-------------------------------------------------------
                        'CP v3.53
                        'October 20, 2005
                        ElseIf m_rstSadbel.Fields("In_Avl_Qty_Wgt").Value < 0 Then
                            If Val(.Value(ColIndex)) > 0 Then
                                MsgBox "There are no remaining units left to correct.", vbInformation + vbOKOnly, Me.Caption
                                Cancel = True
                            ElseIf (m_rstSadbel.Fields("In_TotalOut_Qty_Wgt").Value + Val(.Value(ColIndex)) + .Value(.Columns("Outbound Edits").Index)) < 0 Then
                                MsgBox "The correction value entered is greater than the selected item's available stocks." & vbCrLf & "(Available stocks can be increased by making outbound stock corrections.)", vbInformation + vbOKOnly, Me.Caption
                                Cancel = True
                            Else
                                .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                                If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                                    .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                                End If
                            End If
                        '-------------------------------------------------------
                        'ElseIf Val(.Value(.Columns("Qty for Outbound").Index) - m_rstSadbel.Fields("In_TotalOut_Qty_Wgt").Value) + (m_rstSadbel.Fields("In_Avl_Qty_Wgt").Value - .Value(.Columns("Outbound Edits").Index)) < 0 Then
                        'ElseIf (.Value(.Columns("Outbound Edits").Index) + Val(.Value(.Columns("Qty for Outbound").Index))) - (m_rstSadbel.Fields("In_Avl_Qty_Wgt").Value - m_rstSadbel.Fields("In_TotalOut_Qty_Wgt").Value) < 0 Then
                        'Glenn 3/30/2006 - corrected computation.
                        ElseIf (m_rstSadbel.Fields("In_TotalOut_Qty_Wgt").Value) < Abs((.Value(.Columns("Outbound Edits").Index) + Val(.Value(.Columns("Qty for Outbound").Index)))) Then
                            MsgBox "The correction value entered exceeds the total outbound quantity specified.", vbInformation + vbOKOnly, Me.Caption
                            Cancel = True
                        Else
                            .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                            If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                                .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                            End If
                        End If
                    Else
                        .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                        If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                            .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                        End If
                    End If
                End If
                
                ADORecordsetClose m_rstSadbel
            Else
                If strCorrectionMode = "I" Then
                    If Val(.Value(.Columns("Qty for Outbound").Index)) > m_rstSadbel.Fields("Quantity/Weight").Value Then
                        MsgBox "The correction value entered is greater than the selected item's available stocks." & vbCrLf & "(Available stocks can be increased by making outbound stock corrections.)", vbInformation + vbOKOnly, Me.Caption
                        Cancel = True
                    Else
                        .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                        If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                            .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                        End If
                    End If
                Else
                    '-------------------------------------------------------
                    'CP v3.53
                    'October 20, 2005
                    If chkShowZero.Value = vbChecked Then
                        Debug.Print Val(.Value(.Columns("Qty for Outbound").Index)) & " " & (Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue))
                        If .Value(.Columns("Quantity/Weight").Index) <= 0 Then
                            MsgBox "There are no remaining units left to correct.", vbInformation + vbOKOnly, Me.Caption
                            Cancel = True
                        Else
                            .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                            If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                                .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                            End If
                        End If
                    Else
                    '-------------------------------------------------------
                        'Original code
                        '-------------
                        If Val(.Value(.Columns("Qty for Outbound").Index)) > (Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue)) Then
                                MsgBox "There are no units left for outbound.", vbInformation + vbOKOnly, "Manual Outbound"
                                Cancel = True
                        Else
                            .Value(.Columns("Quantity/Weight").Index) = Replace(CStr(Round(Val(.Value(.Columns("Quantity/Weight").Index)) + Val(OldValue) - Val(.Value(ColIndex)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                            If InStr(.Value(.Columns("Qty for Outbound").Index), ".") Then
                                .Value(.Columns("Qty for Outbound").Index) = Replace(Format(Val(.Value(.Columns("Qty for Outbound").Index)), "0.###"), ",", ".")
                            End If
                        End If
                        '-------------
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub jgxAvailableStock_Change()
    If Not blnSystemChanged Then
        cmdApply.Enabled = True
    End If
    
    'Ensures quantity entered does not exceed stock quantity.
    With jgxAvailableStock
        Select Case .Col
            Case .Columns("Qty for Outbound").Index
                'Decimal places allowed
                If lngHandling = 1 Then         'up to thousandth for Gross Wgt.
                    If InStr(.Value(.Col), ".") Then
                        If Len(Mid(.Value(.Col), InStr(.Value(.Col), ".") + 1)) > 2 Then
                            .Value(.Col) = Mid(.Value(.Col), 1, Len(.Value(.Col)) - 1)
                        End If
                    End If
                ElseIf lngHandling = 2 Then     'up to ten thousandth for Net Wgt.
                    If InStr(.Value(.Col), ".") Then
                        If Len(Mid(.Value(.Col), InStr(.Value(.Col), ".") + 1)) > 3 Then
                            .Value(.Col) = Mid(.Value(.Col), 1, Len(.Value(.Col)) - 1)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub jgxAvailableStock_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    jgxAvailableStock.Update
    If jgxAvailableStock.SortKeys.Count > 0 Then
        If jgxAvailableStock.SortKeys.Item(1).ColIndex = Column.Index Then
            jgxAvailableStock.SortKeys.Item(1).SortOrder = IIf(jgxAvailableStock.SortKeys.Item(1).SortOrder = jgexSortAscending, jgexSortDescending, jgexSortAscending)
        Else
            jgxAvailableStock.SortKeys.Clear
            jgxAvailableStock.SortKeys.Add Column.Index, jgexSortAscending
        End If
    Else
        jgxAvailableStock.SortKeys.Add Column.Index, jgexSortAscending
    End If
    
    jgxAvailableStock.RefreshSort
    
End Sub

Private Sub jgxAvailableStock_GotFocus()
    If jgxAvailableStock.RowCount > 0 Then
        cmdAddToOutbounds.Enabled = True
    End If
    blnAvailableHasFocus = True
End Sub

Private Sub jgxAvailableStock_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = vbKeyReturn And jgxAvailableStock.Row > 0 Then
    
        jgxAvailableStock.Update
        Call AddRecordToOutbound
                
    End If
    
End Sub

Private Sub jgxAvailableStock_KeyPress(KeyAscii As Integer)
    If jgxAvailableStock.Col > 0 Then
        Select Case UCase(jgxAvailableStock.Columns(jgxAvailableStock.Col).Key)
            Case "QTY FOR OUTBOUND"
                If Chr(KeyAscii) = "." Then
                    If lngHandling = 0 Then
                        KeyAscii = 0
                    Else
                        If InStr(CStr(IIf(IsNull(jgxAvailableStock.Value(jgxAvailableStock.Col)), "", jgxAvailableStock.Value(jgxAvailableStock.Col))), ".") Then
                            KeyAscii = 0
                        End If
                    End If
                ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
                    If Len(strCorrectionMode) > 0 And KeyAscii = 45 Then
                        If InStr(CStr(IIf(IsNull(jgxAvailableStock.Value(jgxAvailableStock.Col)), "", jgxAvailableStock.Value(jgxAvailableStock.Col))), "-") Then
                            KeyAscii = 0
                        End If
                    Else
                        KeyAscii = 0
                    End If
                ElseIf UCase(jgxAvailableStock.Columns(jgxAvailableStock.Col).Key) = "QTY FOR OUTBOUND" And UCase(jgxAvailableStock.Value(jgxAvailableStock.Col)) = "<NEW>" Then
                    jgxAvailableStock.SelStart = 0
                    jgxAvailableStock.SelLength = Len(jgxAvailableStock.Value(jgxAvailableStock.Col))
                End If
            Case "JOB NO"
                'Only allows numberic + upper/lower case alpha.
                If Not (KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

Private Sub jgxAvailableStock_LostFocus()
    blnAvailableHasFocus = False
    If UCase(Me.ActiveControl.Name) <> "CMDADDTOOUTBOUNDS" Then
        cmdAddToOutbounds.Enabled = False
        
        If UCase(Me.ActiveControl.Name) = "CMDAPPLY" Then
            blnAvailableIsLastFocus = True
        End If
    End If
End Sub

Private Sub jgxAvailableStock_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    
    If Not jgxAvailableStock.ADORecordset Is Nothing And jgxAvailableStock.Row > 0 Then
        lngHandling = jgxAvailableStock.Value(jgxAvailableStock.Columns("Handling").Index)
    End If
    
End Sub

Private Sub jgxAvailableStock_Validate(Cancel As Boolean)
    jgxAvailableStock.Update
End Sub

Private Sub jgxOutbounds_AfterColUpdate(ByVal ColIndex As Integer)
    
    Dim strAvailable As String
    Dim lngrow As Long
    Dim lngCol As Long
    Dim lngStart As Long
    Dim lngLen As Long
    
    If ColIndex = jgxOutbounds.Columns("Outbound Qty").Index Then
           
        strAvailable = jgxOutbounds.Value(jgxOutbounds.Columns("Quantity/Weight").Index)
        lngrow = jgxOutbounds.Row
        lngCol = jgxOutbounds.Col
        lngStart = jgxOutbounds.SelStart
        lngLen = jgxOutbounds.SelLength
        
        m_rstOutboundsOff.Filter = "In_ID = " & m_rstOutboundsOff!In_ID
        
        If m_rstOutboundsOff.RecordCount > 0 Then
            m_rstOutboundsOff.MoveFirst
        End If
        
        Do While Not m_rstOutboundsOff.EOF
            m_rstOutboundsOff![Quantity/Weight] = strAvailable
            m_rstOutboundsOff.MoveNext
        Loop
        
        m_rstOutboundsOff.Filter = adFilterNone
    
        If Val(jgxOutbounds.Value(ColIndex)) = 0 Then
            jgxOutbounds.Delete
        End If
        
        jgxOutbounds.Row = IIf(lngrow > jgxOutbounds.RowCount, jgxOutbounds.RowCount, lngrow)
        jgxOutbounds.Col = lngCol
        jgxOutbounds.SelStart = lngStart
        jgxOutbounds.SelLength = lngLen
    
    End If
    
End Sub

Private Sub jgxOutbounds_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX16.JSRetBoolean)
    
    Dim dblDifference As Double
    Dim lngCtr As Long
    Dim lngIndex As Long
    Dim lngSort As Long
    
    If ColIndex = jgxOutbounds.Columns("Outbound Qty").Index Then
    
        blnOutboundsChanged = True
        m_rstAvailableOff.Filter = adFilterNone
        m_rstAvailableOff.Filter = "In_ID = " & jgxOutbounds.Value(jgxOutbounds.Columns("In_ID").Index)
        dblDifference = Round(Val(jgxOutbounds.Value(ColIndex)) - Val(OldValue), Choose(lngHandling + 1, 0, 2, 3))
        
        If m_rstAvailableOff.EOF And m_rstAvailableOff.BOF Then
            If dblDifference > 0 Then
                MsgBox "There are no units left for outbound.", vbInformation + vbOKOnly, "Manual Outbound"
                Cancel = True
                Exit Sub
            Else
                m_rstAvailableOff.AddNew
                For lngCtr = 0 To m_rstAvailableOff.Fields.Count - 1
                    If UCase(m_rstAvailableOff.Fields(lngCtr).Name) <> "QTY FOR OUTBOUND" And _
                        UCase(m_rstAvailableOff.Fields(lngCtr).Name) <> "OUTBOUND EDITS" Then
                        
                        m_rstAvailableOff.Fields(lngCtr).Value = jgxOutbounds.Value(jgxOutbounds.Columns(m_rstAvailableOff.Fields(lngCtr).Name).Index)
                        
                    End If
                Next
                m_rstAvailableOff![Quantity/Weight] = Replace(CStr(-1 * dblDifference), ",", ".")
                m_rstAvailableOff![Outbound Edits] = dblDifference
                jgxOutbounds.Value(jgxOutbounds.Columns("Quantity/Weight").Index) = m_rstAvailableOff![Quantity/Weight]
                If InStr(jgxOutbounds.Value(ColIndex), ".") Then
                    jgxOutbounds.Value(ColIndex) = Replace(Format(Val(jgxOutbounds.Value(ColIndex)), "0.###"), ",", ".")
                End If
            End If
        Else
            'If Val(m_rstAvailableOff![Quantity/Weight]) - dblDifference < 0 Then
            '-------------------------------------------------------
            'CP v3.53
            'October 20, 2005
            If Val(m_rstAvailableOff![Quantity/Weight]) - dblDifference < 0 Then 'IIf(jgxOutbounds.Value(ColIndex) < 0, dblDifference * -1, dblDifference) < 0 Then
            '-------------------------------------------------------
                MsgBox "There are no units left for outbound.", vbInformation + vbOKOnly, "Manual Outbound"
                Cancel = True
                Exit Sub
            Else
                m_rstAvailableOff![Quantity/Weight] = Replace(CStr(Round(Val(m_rstAvailableOff![Quantity/Weight]) - dblDifference, Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
                m_rstAvailableOff![Outbound Edits] = Round(m_rstAvailableOff![Outbound Edits] + dblDifference, Choose(lngHandling + 1, 0, 2, 3))
                jgxOutbounds.Value(jgxOutbounds.Columns("Quantity/Weight").Index) = m_rstAvailableOff![Quantity/Weight]
                If InStr(jgxOutbounds.Value(ColIndex), ".") Then
                    jgxOutbounds.Value(ColIndex) = Replace(Format(Val(jgxOutbounds.Value(ColIndex)), "0.###"), ",", ".")
                End If
            End If
        End If
                
        If jgxAvailableStock.SortKeys.Count > 0 Then
            lngIndex = jgxAvailableStock.SortKeys(1).ColIndex
            lngSort = jgxAvailableStock.SortKeys(1).SortOrder
        End If
        m_rstAvailableOff.Filter = adFilterNone
        If strCorrectionMode <> "I" Then
            If chkShowZero.Value = vbUnchecked Then
                m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Qty for Outbound] > 0"
            End If
        End If
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        Call FormatAvailable
        If lngIndex > 0 Then
            jgxAvailableStock.SortKeys.Add lngIndex, lngSort
            jgxAvailableStock.RefreshSort
            lngIndex = 0
        End If
        
    End If
    
End Sub

Private Sub jgxOutbounds_BeforeDelete(ByVal Cancel As GridEX16.JSRetBoolean)
    
    Dim lngCtr As Long
    Dim lngIndex As Long
    Dim lngSort As Long
    
    cmdApply.Enabled = True
    
    If jgxOutbounds.Value(jgxOutbounds.Columns("Out_ID").Index) <> 0 Then
        ReDim Preserve alngDeleted(UBound(alngDeleted) + 1)
        alngDeleted(UBound(alngDeleted)) = jgxOutbounds.Value(jgxOutbounds.Columns("Out_ID").Index)
    End If
    
'    jgxAvailableStock.Update
    m_rstAvailableOff.Filter = adFilterNone
    m_rstAvailableOff.Filter = "In_ID = " & jgxOutbounds.Value(jgxOutbounds.Columns("In_ID").Index)
    
    If (Len(strCorrectionMode) > 0) Or (Len(strCorrectionMode) = 0 And Val(jgxOutbounds.Value(jgxOutbounds.Columns("Outbound Qty").Index)) > 0) Then
        If m_rstAvailableOff.RecordCount = 0 Then
            m_rstAvailableOff.AddNew
            For lngCtr = 0 To m_rstAvailableOff.Fields.Count - 1
                If UCase(m_rstAvailableOff.Fields(lngCtr).Name) <> "QTY FOR OUTBOUND" And _
                    UCase(m_rstAvailableOff.Fields(lngCtr).Name) <> "OUTBOUND EDITS" Then
                    
                    m_rstAvailableOff.Fields(lngCtr).Value = jgxOutbounds.Value(jgxOutbounds.Columns(m_rstAvailableOff.Fields(lngCtr).Name).Index)
                    
                End If
            Next
            m_rstAvailableOff![Quantity/Weight] = jgxOutbounds.Value(jgxOutbounds.Columns("Outbound Qty").Index)
            m_rstAvailableOff![Outbound Edits] = -1 * Val(jgxOutbounds.Value(jgxOutbounds.Columns("Outbound Qty").Index))
        Else
            m_rstAvailableOff![Quantity/Weight] = Replace(CStr(Round(Val(m_rstAvailableOff![Quantity/Weight]) + Val(jgxOutbounds.Value(jgxOutbounds.Columns("Outbound Qty").Index)), Choose(lngHandling + 1, 0, 2, 3))), ",", ".")
            m_rstAvailableOff![Outbound Edits] = Round(m_rstAvailableOff![Outbound Edits] - Val(jgxOutbounds.Value(jgxOutbounds.Columns("Outbound Qty").Index)), Choose(lngHandling + 1, 0, 2, 3))
        End If
        m_rstAvailableOff.Update
    End If
    
    blnSystemChanged = True
            
    If jgxAvailableStock.SortKeys.Count > 0 Then
        lngIndex = jgxAvailableStock.SortKeys(1).ColIndex
        lngSort = jgxAvailableStock.SortKeys(1).SortOrder
    End If
    m_rstAvailableOff.Filter = adFilterNone
    If strCorrectionMode <> "I" Then
        If chkShowZero.Value = vbUnchecked Then
            m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Qty for Outbound] > 0"
        End If
    End If
    Set jgxAvailableStock.ADORecordset = Nothing
    Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
    Call FormatAvailable
    If lngIndex > 0 Then
        jgxAvailableStock.SortKeys.Add lngIndex, lngSort
        jgxAvailableStock.RefreshSort
        lngIndex = 0
    End If
    
    blnSystemChanged = False

End Sub

Private Sub jgxOutbounds_Change()
    If Not blnSystemChanged Then
        cmdApply.Enabled = True
    End If
    
    'Ensures quantity entered does not exceed stock quantity.
    With jgxOutbounds
        Select Case .Col
            Case .Columns("Outbound Qty").Index
                'Decimal places allowed
                If lngHandling = 1 Then         'up to thousandth for Gross Wgt.
                    If InStr(.Value(.Col), ".") Then
                        If Len(Mid(.Value(.Col), InStr(.Value(.Col), ".") + 1)) > 2 Then
                            .Value(.Col) = Mid(.Value(.Col), 1, Len(.Value(.Col)) - 1)
                        End If
                    End If
                ElseIf lngHandling = 2 Then     'up to ten thousandth for Net Wgt.
                    If InStr(.Value(.Col), ".") Then
                        If Len(Mid(.Value(.Col), InStr(.Value(.Col), ".") + 1)) > 3 Then
                            .Value(.Col) = Mid(.Value(.Col), 1, Len(.Value(.Col)) - 1)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub jgxOutbounds_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)

    If jgxOutbounds.SortKeys.Count > 0 Then
        If jgxOutbounds.SortKeys.Item(1).ColIndex = Column.Index Then
            jgxOutbounds.SortKeys.Item(1).SortOrder = IIf(jgxOutbounds.SortKeys.Item(1).SortOrder = jgexSortAscending, jgexSortDescending, jgexSortAscending)
        Else
            jgxOutbounds.SortKeys.Clear
            jgxOutbounds.SortKeys.Add Column.Index, jgexSortAscending
        End If
    Else
        jgxOutbounds.SortKeys.Add Column.Index, jgexSortAscending
    End If

End Sub

Private Sub jgxOutbounds_KeyPress(KeyAscii As Integer)
    If jgxOutbounds.Col > 0 Then
        Select Case UCase(jgxOutbounds.Columns(jgxOutbounds.Col).Key)
            Case "OUTBOUND QTY"
                If Chr(KeyAscii) = "." Then
                    If lngHandling = 0 Then
                        KeyAscii = 0
                    Else
                        If InStr(CStr(IIf(IsNull(jgxOutbounds.Value(jgxOutbounds.Col)), "", jgxOutbounds.Value(jgxOutbounds.Col))), ".") Then
                            KeyAscii = 0
                        End If
                    End If
                ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
                    '-------------------------------------------------------
                    'CP v3.53
                    'October 20, 2005
                    If Len(strCorrectionMode) > 0 And KeyAscii = 45 Then
                        If InStr(CStr(IIf(IsNull(jgxOutbounds.Value(jgxOutbounds.Col)), "", jgxOutbounds.Value(jgxOutbounds.Col))), "-") Then
                            KeyAscii = 0
                        End If
                    Else
                        KeyAscii = 0
                    End If
                    '-------------------------------------------------------
                    'KeyAscii = 0
                ElseIf UCase(jgxOutbounds.Columns(jgxOutbounds.Col).Key) = "OUTBOUND QTY" And UCase(jgxOutbounds.Value(jgxOutbounds.Col)) = "<NEW>" Then
                    jgxOutbounds.SelStart = 0
                    jgxOutbounds.SelLength = Len(jgxOutbounds.Value(jgxOutbounds.Col))
                End If
            Case "JOB NO"
                'Only allows numberic + upper/lower case alpha.
                If Not (KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

Private Sub jgxOutbounds_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If Not jgxOutbounds.ADORecordset Is Nothing And jgxOutbounds.Row > 0 Then
        lngHandling = jgxOutbounds.Value(jgxOutbounds.Columns("Handling").Index)
    End If
End Sub

Private Sub jgxOutbounds_Validate(Cancel As Boolean)
    jgxOutbounds.Update
End Sub

Private Sub txtCommunalSettlement_Change()
    If jgxAvailableStock.RowCount > 0 Or jgxOutbounds.RowCount > 0 Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtCommunalSettlement_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDocNum_Change()
    
    If Trim(UCase(txtDocNum.Text)) <> Trim(UCase(strLastDocNum)) Then
    
        If lngOutDocID <> 0 Then

            If Not SavedIfChanged Then
                txtDocNum.Text = strLastDocNum
                Exit Sub
            End If
            
            blnSystemChanged = True
            txtProductNum.Text = ""
            Call ResetProduct
            blnSystemChanged = False
                        
            strLastDocNum = ""
            strLastDocType = ""
            strLastMRN = ""
            m_dteLastDate = 0
    
            Set jgxAvailableStock.ADORecordset = Nothing
            Call CreateAvailableOffline
            Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
            Call FormatAvailable
    
            Set jgxOutbounds.ADORecordset = Nothing
            Call CreateOutboundsOffline
            Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
            Call FormatOutbounds
        
        End If
        
        Call CheckOutDoc
        
    End If
        
    If jgxAvailableStock.RowCount > 0 Or jgxOutbounds.RowCount > 0 Then
        cmdApply.Enabled = True
    End If
    
End Sub

Private Sub txtDocNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtDocNum.Text <> strLastDocNum Then
    ElseIf Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    ElseIf Not (KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57) Then
        'KeyAscii = 0
    End If
End Sub

Private Sub txtDocNum_LostFocus()
    'Left pad the Document number if it has less than 7 characters.
    Dim bytCounter As Byte
    Dim bytNumberofZeros As Byte
    Dim strZeros As String
    
    If Len(Trim(txtDocNum.Text)) < 7 And Len(Trim(txtDocNum.Text)) <> 0 Then
        bytNumberofZeros = 7 - Len(txtDocNum.Text)
        For bytCounter = 1 To bytNumberofZeros
            strZeros = "0" & strZeros
        Next bytCounter
        txtDocNum.Text = strZeros & txtDocNum.Text
    End If
End Sub

Private Sub FormatAvailable()
    If strCorrectionMode = "I" Then jgxAvailableStock.Columns("Qty for Outbound").Caption = "Qty for Deduction"
    
    jgxAvailableStock.Columns("In_ID").Visible = False
    jgxAvailableStock.Columns("Outbound Edits").Visible = False
    jgxAvailableStock.Columns("Handling").Visible = False
    
    jgxAvailableStock.Columns("Entrepot No").Width = 1200
    jgxAvailableStock.Columns("Product No").Width = 1200
    jgxAvailableStock.Columns("Stock Card No").Width = 1200
    jgxAvailableStock.Columns("Document No").Width = 1200
    jgxAvailableStock.Columns("Quantity/Weight").Width = 1305
    jgxAvailableStock.Columns("Package Type").Width = 1200
    jgxAvailableStock.Columns("Batch No").Width = 1245
    jgxAvailableStock.Columns("Qty for Outbound").Width = 1350
    jgxAvailableStock.Columns("Job No").Width = 1200

    jgxAvailableStock.Columns("Entrepot No").Selectable = False
    jgxAvailableStock.Columns("Product No").Selectable = False
    jgxAvailableStock.Columns("Stock Card No").Selectable = False
    jgxAvailableStock.Columns("Document No").Selectable = False
    jgxAvailableStock.Columns("Quantity/Weight").Selectable = False
    jgxAvailableStock.Columns("Package Type").Selectable = False
    jgxAvailableStock.Columns("Batch No").Selectable = False
    
    jgxAvailableStock.Columns("Entrepot No").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Product No").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Stock Card No").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Document No").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Quantity/Weight").TextAlignment = jgexAlignRight
    jgxAvailableStock.Columns("Package Type").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Batch No").TextAlignment = jgexAlignLeft
    jgxAvailableStock.Columns("Qty for Outbound").TextAlignment = jgexAlignRight
    jgxAvailableStock.Columns("Job No").TextAlignment = jgexAlignLeft
    
    jgxAvailableStock.Columns("Entrepot No").SortType = jgexSortTypeString
    jgxAvailableStock.Columns("Product No").SortType = jgexSortTypeString
    jgxAvailableStock.Columns("Stock Card No").SortType = jgexSortTypeNumeric
    jgxAvailableStock.Columns("Quantity/Weight").SortType = jgexSortTypeNumeric
    jgxAvailableStock.Columns("Qty for Outbound").SortType = jgexSortTypeNumeric
    jgxAvailableStock.Columns("Document No").SortType = jgexSortTypeString
    jgxAvailableStock.Columns("Package Type").SortType = jgexSortTypeString
    jgxAvailableStock.Columns("Job No").SortType = jgexSortTypeString
    jgxAvailableStock.Columns("Batch No").SortType = jgexSortTypeString
        
    jgxAvailableStock.Columns("Job No").MaxLength = 50
    jgxAvailableStock.Columns("Qty for Outbound").MaxLength = IIf(lngHandling = 0, 6, 12)
        
    If jgxAvailableStock.RowCount > 0 And blnAvailableHasFocus Then
        cmdAddToOutbounds.Enabled = True
    Else
        cmdAddToOutbounds.Enabled = False
    End If
    
End Sub

Private Sub FormatOutbounds()
    If strCorrectionMode = "I" Then jgxOutbounds.Columns("Outbound Qty").Caption = "Inbound Qty"
    
    jgxOutbounds.Columns("Entrepot No").Width = 1545
    jgxOutbounds.Columns("Product No").Width = 1545
    jgxOutbounds.Columns("Stock Card No").Width = 1545
    jgxOutbounds.Columns("Quantity/Weight").Width = 1545
    jgxOutbounds.Columns("Package Type").Width = 1545
    jgxOutbounds.Columns("Batch No").Width = 1545
    jgxOutbounds.Columns("Outbound Qty").Width = 1545
    jgxOutbounds.Columns("Job No").Width = 1545
    
    jgxOutbounds.Columns("In_ID").Visible = False
    jgxOutbounds.Columns("Out_ID").Visible = False
    jgxOutbounds.Columns("Quantity/Weight").Visible = False
    jgxOutbounds.Columns("Document No").Visible = False
    jgxOutbounds.Columns("Handling").Visible = False
    
    jgxOutbounds.Columns("Entrepot No").Selectable = False
    jgxOutbounds.Columns("Product No").Selectable = False
    jgxOutbounds.Columns("Stock Card No").Selectable = False
    jgxOutbounds.Columns("Package Type").Selectable = False
    jgxOutbounds.Columns("Batch No").Selectable = False
    
    jgxOutbounds.Columns("Entrepot No").TextAlignment = jgexAlignLeft
    jgxOutbounds.Columns("Product No").TextAlignment = jgexAlignLeft
    jgxOutbounds.Columns("Stock Card No").TextAlignment = jgexAlignLeft
    jgxOutbounds.Columns("Quantity/Weight").TextAlignment = jgexAlignRight
    jgxOutbounds.Columns("Package Type").TextAlignment = jgexAlignLeft
    jgxOutbounds.Columns("Batch No").TextAlignment = jgexAlignLeft
    jgxOutbounds.Columns("Outbound Qty").TextAlignment = jgexAlignRight
    jgxOutbounds.Columns("Job No").TextAlignment = jgexAlignLeft
    
    jgxOutbounds.Columns("Entrepot No").SortType = jgexSortTypeString
    jgxOutbounds.Columns("Product No").SortType = jgexSortTypeString
    jgxOutbounds.Columns("Stock Card No").SortType = jgexSortTypeNumeric
    jgxOutbounds.Columns("Quantity/Weight").SortType = jgexSortTypeNumeric
    jgxOutbounds.Columns("Outbound Qty").SortType = jgexSortTypeNumeric
    jgxOutbounds.Columns("Package Type").SortType = jgexSortTypeString
    jgxOutbounds.Columns("Job No").SortType = jgexSortTypeString
    jgxOutbounds.Columns("Batch No").SortType = jgexSortTypeString
    
    jgxOutbounds.Columns("Job No").MaxLength = 50
    jgxOutbounds.Columns("Outbound Qty").MaxLength = IIf(lngHandling = 0, 6, 12)
    
End Sub

Public Sub MyLoad(ByRef Sadbel As ADODB.Connection, ByRef Taric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, _
                  ByVal UserID As Long, Optional ByVal IOCorrectionMode As String)
                  
    Set m_conSADBEL = Sadbel
    Set m_conTaric = Taric
    
    m_lngUserID = UserID
    
    intTaricProperties = TaricProperties
    strLanguage = Language
    lngResourceHandler = MyResourceHandler

    File1.Path = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
    File1.Pattern = "mdb_history*.mdb"
    
    'Applies appropriate caption changes according to correction mode selected
    If Len(IOCorrectionMode) > 0 Then
    
        'CSCLP-233 - BCo
        'This was added cuz switching between Manual Outbound and Correction mode only works in with English menus
        Select Case UCase$(Language)
            Case "DUTCH"
                Select Case IOCorrectionMode
                    Case "Correctie ingaand"
                        IOCorrectionMode = "Inbound Correction"
                    Case "Correctie uitgaand"
                        IOCorrectionMode = "Outbound Correction"
                End Select
            Case "FRENCH"
                Select Case IOCorrectionMode
                    Case "Corriger l'entrée"
                        IOCorrectionMode = "Inbound Correction"
                    Case "Corriger partance"
                        IOCorrectionMode = "Outbound Correction"
                End Select
        End Select
    
        'Glenn 3/30/2006 - changed position of letter to get since Caption has changed.
        strCorrectionMode = UCase$(Mid$(IOCorrectionMode, 1, 1))
        If strCorrectionMode = "I" Then
            Me.Caption = "Inbounds Correction"
            FraOutbounds.Caption = "Inbounds Correction"
            Label1.Caption = "Press Enter or click Add To Inbounds to add correction stock of the currently selected row."
            cmdAddToOutbounds.Caption = "A&dd To Inbounds"
        ElseIf strCorrectionMode = "O" Then
            Me.Caption = "Outbounds Correction"
            FraOutbounds.Caption = "Outbounds Correction"
            Label1.Caption = "Press Enter or click Add To Outbounds to add correction stock of the currently selected row."
            cmdAddToOutbounds.Caption = "A&dd To Outbounds"
        End If
    End If
    
    Me.Show vbModal
    
    Set m_conSADBEL = Nothing
    Set m_conTaric = Nothing
    
End Sub

Private Sub PopulateAvailable()
    Dim strSELECT As String
    Dim strFROM As String
    Dim strWhere As String
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    Dim rstDIA As ADODB.Recordset
    
    
    'SQL to get DIA records
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Inbounds!In_Batch_Num AS Batch_Num, "
        strSQL = strSQL & "Inbounds!Stock_ID AS Stock_ID, "
        strSQL = strSQL & "IIf(IsNull(Inbounds!In_Header),0,Inbounds!In_Header) AS Header, "
        strSQL = strSQL & "IIf(IsNull(Inbounds!In_Detail),0,Inbounds!In_Detail) AS Detail "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Inbounds INNER JOIN ("
        strSQL = strSQL & "StockCards INNER JOIN Products "
        strSQL = strSQL & "ON Products.Prod_ID = StockCards.Prod_ID) "
        strSQL = strSQL & "ON StockCards.Stock_ID = Inbounds.Stock_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "( "
        strSQL = strSQL & "Inbounds!In_Job_Num='CANCELLATION' "
        strSQL = strSQL & "OR "
        strSQL = strSQL & "Inbounds!In_Job_Num='DIA' "
        strSQL = strSQL & ") "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "Choose(Products!Prod_Handling + 1, Inbounds!In_Orig_Packages_Qty < 0, Inbounds!In_Orig_Gross_Weight < 0, Inbounds!In_Orig_Net_Weight < 0)"
    ADORecordsetOpen strSQL, m_conSADBEL, rstDIA, adOpenKeyset, adLockOptimistic
    
    
    'The strCorrectionMode = "I" condition ensures initial stock value is loaded for correction
    'The choose selects which initial stock field to use based on the products_handling value
        strSELECT = vbNullString
        strSELECT = strSELECT & "SELECT "
        strSELECT = strSELECT & "Inbounds!In_ID AS In_ID, "
        strSELECT = strSELECT & "Inbounds!In_Header AS Header, "
        strSELECT = strSELECT & "Inbounds!In_Detail AS Detail, "
        strSELECT = strSELECT & "Products!Prod_Handling AS Handling, "
        strSELECT = strSELECT & "Products!Prod_Num AS [Product No], "
        strSELECT = strSELECT & "Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num AS [Entrepot No], "
        strSELECT = strSELECT & "StockCards!Stock_ID AS Stock_ID, "
        strSELECT = strSELECT & "StockCards!Stock_Card_Num AS [Stock Card No], "
        strSELECT = strSELECT & "InboundDocs!InDoc_Num AS InDoc_Num, "
        strSELECT = strSELECT & "InboundDocs!InDoc_Type & '-' & InboundDocs!InDoc_Num AS [Document No], "
        strSELECT = strSELECT & "Inbounds!In_Avl_Qty_Wgt AS [Quantity/Weight], "
        strSELECT = strSELECT & "Inbounds!In_Orig_Packages_Type AS [Package Type], "
        strSELECT = strSELECT & "Inbounds!In_Batch_Num AS [Batch No] "
        
        strFROM = vbNullString
        strFROM = strFROM & "FROM "
        strFROM = strFROM & "InboundDocs INNER JOIN ("
        strFROM = strFROM & "Inbounds INNER JOIN ("
        strFROM = strFROM & "StockCards INNER JOIN ("
        strFROM = strFROM & "Products INNER JOIN Entrepots "
        strFROM = strFROM & "ON Products.Entrepot_ID = Entrepots.Entrepot_ID) "
        strFROM = strFROM & "ON Stockcards.Prod_ID = Products.Prod_ID) "
        strFROM = strFROM & "ON Inbounds.Stock_ID = Stockcards.Stock_ID) "
        strFROM = strFROM & "ON Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
        
        strWhere = vbNullString
        strWhere = strWhere & "WHERE "
        strWhere = strWhere & "Products!Prod_ID IN (" & IIf(Trim(strProdID) = "", 0, strProdID) & ") "
        strWhere = strWhere & "AND "
        strWhere = strWhere & "IIf(IsNull(Inbounds!In_Code), '', Inbounds!In_Code) NOT LIKE '%<<TEST>>' "
        strWhere = strWhere & "AND "
        strWhere = strWhere & "IIf(IsNull(Inbounds!In_Code), '', Inbounds!In_Code) NOT LIKE '%<<CLOSURE>>' "
        strWhere = strWhere & IIf(strCorrectionMode = "I", "AND InboundDocs!InDoc_Global <> -1 ", " ")
        
        strSQL = strSELECT & strFROM & strWhere & " AND Inbounds!In_Avl_Qty_Wgt>0"
        
    ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.MoveFirst
    
        Do While Not rstTemp.EOF
            rstDIA.Filter = adFilterNone
            rstDIA.Filter = "Batch_Num = '" & rstTemp!InDoc_Num & "' AND Stock_ID = " & rstTemp!Stock_ID & " AND Header = " & IIf(IsNull(rstTemp!Header), 0, rstTemp!Header) & " AND Detail = " & IIf(IsNull(rstTemp!Detail), 0, rstTemp!Detail)
            
            'IAN 05-20-2005
            'If record count is greater than zero then current record on rstTemp has been
            'cancelled by DIA, hence, don't include in the grid.
            If (chkShowZero.Value = vbChecked) Or (chkShowZero.Value = vbUnchecked And rstDIA.RecordCount = 0) Then
                m_rstAvailableOff.AddNew
                    m_rstAvailableOff!In_ID = rstTemp!In_ID
                    m_rstAvailableOff![Product No] = rstTemp![Product No]
                    m_rstAvailableOff!Handling = rstTemp!Handling
                    m_rstAvailableOff![Entrepot No] = rstTemp![Entrepot No]
                    m_rstAvailableOff![Stock Card No] = rstTemp![Stock Card No]
                    m_rstAvailableOff![Document No] = rstTemp![Document No]
                    m_rstAvailableOff![Quantity/Weight] = Replace(CStr(rstTemp![Quantity/Weight]), ",", ".")
                    m_rstAvailableOff![Package Type] = rstTemp![Package Type]
                    m_rstAvailableOff![Batch No] = rstTemp![Batch No]
                    m_rstAvailableOff![Qty for Outbound] = ""
                    m_rstAvailableOff![Outbound Edits] = 0
                m_rstAvailableOff.Update
            End If
            
            rstTemp.MoveNext
        Loop
    End If
    
    If (Len(strCorrectionMode) > 0) Then
        PopulateAvailableFromHistoryForCorrectionModes strSELECT, strWhere, rstTemp, rstDIA
    End If
    
    ADORecordsetClose rstTemp
    ADORecordsetClose rstDIA
    
    
    If (m_rstAvailableOff.RecordCount > 0 Or jgxAvailableStock.RowCount > 0) Then
        blnSystemChanged = True
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        Call FormatAvailable
        blnSystemChanged = False
    End If
    
End Sub

Private Sub PopulateAvailableFromHistoryForCorrectionModes(ByVal strSELECT As String, ByVal strWhere As String, ByVal rstTemp As ADODB.Recordset, ByVal rstDIA As ADODB.Recordset)
    Dim strSQL As String
    Dim strFROM As String
    Dim lngCtr As Long
    
    Dim strHistoryDBYear As String
    
    
    'The strCorrectionMode = "I" condition prevents displaying of stock records created using Initial Stock in Inbounds Correction
    'Purpose: cuz Inbounds Correction is used to make changes to stock movement sent to customs
    'Change "Z" to "I" when preliminary testing is complete
    
    '-------------------------------------------------------
    'CP v3.53
    'October 20, 2005
    If (UCase(Trim(strCorrectionMode)) = "O") Then
        'Only shows negative stock in Outbounds correction
        strWhere = strWhere & " AND Inbounds!In_Avl_Qty_Wgt <= 0 AND CHOOSE(Products!Prod_Handling + 1, Inbounds!In_Orig_Packages_Qty > 0, Inbounds!In_Orig_Gross_Weight > 0, Inbounds!In_Orig_Net_Weight > 0)"
    ElseIf strCorrectionMode = "I" Then
        'CSCLP-233
        'strWhere = strWhere & " AND Inbounds!In_Avl_Qty_Wgt = 0"
        strWhere = strWhere & " AND Inbounds!In_Avl_Qty_Wgt <= 0"
    End If
    '-------------------------------------------------------
    
    
    'This code is used by Inbounds/Outbounds Correction
    For lngCtr = 0 To File1.ListCount - 1
        strHistoryDBYear = Replace(File1.List(lngCtr), "mdb_history", vbNullString)
        strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
        
        
            strFROM = vbNullString
            strFROM = strFROM & "FROM "
            strFROM = strFROM & "InboundDocs INNER JOIN ("
            strFROM = strFROM & "Inbounds INNER JOIN ("
            strFROM = strFROM & "StockCards INNER JOIN ("
            strFROM = strFROM & "Products INNER JOIN Entrepots "
            strFROM = strFROM & "ON Products.Entrepot_ID = Entrepots.Entrepot_ID) "
            strFROM = strFROM & "ON Stockcards.Prod_ID = Products.Prod_ID) "
            strFROM = strFROM & "ON Inbounds.Stock_ID = Stockcards.Stock_ID) "
            strFROM = strFROM & "ON Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
            
            strSQL = strSELECT & strFROM & strWhere
            
            strSQL = Replace(strSQL, "Inbounds", "InboundHistory" & strHistoryDBYear & "_" & Format(m_lngUserID, "00"))
            strSQL = Replace(strSQL, "InboundDocs", "InboundDocHistory" & strHistoryDBYear & "_" & Format(m_lngUserID, "00"))
        
        ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
        If Not (rstTemp.EOF And rstTemp.BOF) Then
            rstTemp.MoveFirst
        
            Do While Not rstTemp.EOF
                rstDIA.Filter = adFilterNone
                rstDIA.Filter = "Batch_Num = '" & rstTemp!InDoc_Num & "' AND Stock_ID = " & rstTemp!Stock_ID & " AND Header = " & IIf(IsNull(rstTemp!Header), 0, rstTemp!Header) & " AND Detail = " & IIf(IsNull(rstTemp!Detail), 0, rstTemp!Detail)
                                        
                'IAN 05-20-2005
                'If record count is greater than zero then current record on rstTemp has been cancelled by DIA, hence, don't include in the grid.
                If (chkShowZero.Value = vbChecked) Or (chkShowZero.Value = vbUnchecked And rstDIA.RecordCount = 0) Then
                    m_rstAvailableOff.AddNew
                    m_rstAvailableOff!In_ID = rstTemp!In_ID
                    m_rstAvailableOff![Product No] = rstTemp![Product No]
                    m_rstAvailableOff!Handling = rstTemp!Handling
                    m_rstAvailableOff![Entrepot No] = rstTemp![Entrepot No]
                    m_rstAvailableOff![Stock Card No] = rstTemp![Stock Card No]
                    m_rstAvailableOff![Document No] = rstTemp![Document No]
                    m_rstAvailableOff![Quantity/Weight] = Replace(CStr(rstTemp![Quantity/Weight]), ",", ".")
                    m_rstAvailableOff![Package Type] = rstTemp![Package Type]
                    m_rstAvailableOff![Batch No] = rstTemp![Batch No]
                    m_rstAvailableOff![Qty for Outbound] = ""
                    m_rstAvailableOff![Outbound Edits] = 0
                    m_rstAvailableOff.Update
                End If
                rstTemp.MoveNext
            Loop
        End If
    Next lngCtr
    
    If (chkShowZero.Value = Unchecked) Then
        m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Outbound Edits] <> 0"
    Else
        m_rstAvailableOff.Filter = adFilterNone
    End If
    
End Sub

Private Sub CreateOutboundsOffline()
    ADORecordsetClose m_rstOutboundsOff
    
    Set m_rstOutboundsOff = New ADODB.Recordset
    m_rstOutboundsOff.CursorLocation = adUseClient
    
    m_rstOutboundsOff.Fields.Append "Entrepot No", adVarWChar, 50
    m_rstOutboundsOff.Fields.Append "Product No", adVarWChar, 50
    m_rstOutboundsOff.Fields.Append "Stock Card No", adVarWChar, 50, 102
    m_rstOutboundsOff.Fields.Append "Outbound Qty", adVarWChar, 50
    m_rstOutboundsOff.Fields.Append "Package Type", adVarWChar, 50, 102
    m_rstOutboundsOff.Fields.Append "Job No", adVarWChar, 50, 102
    m_rstOutboundsOff.Fields.Append "Batch No", adVarWChar, 50, 102
    m_rstOutboundsOff.Fields.Append "Out_ID", adInteger, 4, 90
    m_rstOutboundsOff.Fields.Append "In_ID", adInteger, 4, 118
    m_rstOutboundsOff.Fields.Append "Handling", adInteger, 1
    m_rstOutboundsOff.Fields.Append "Document No", adVarWChar, 50, 102
    m_rstOutboundsOff.Fields.Append "Quantity/Weight", adVarWChar, 50
    
    m_rstOutboundsOff.Open
        
End Sub

Private Sub CreateAvailableOffline()

    ADORecordsetClose m_rstAvailableOff
    
    Set m_rstAvailableOff = New ADODB.Recordset
    m_rstAvailableOff.CursorLocation = adUseClient
    
    m_rstAvailableOff.Fields.Append "Entrepot No", adVarWChar, 50
    m_rstAvailableOff.Fields.Append "Product No", adVarWChar, 50
    m_rstAvailableOff.Fields.Append "Stock Card No", adVarWChar, 50, 102
    m_rstAvailableOff.Fields.Append "Document No", adVarWChar, 50, 102
    m_rstAvailableOff.Fields.Append "Quantity/Weight", adVarWChar, 50
    m_rstAvailableOff.Fields.Append "Package Type", adVarWChar, 50, 102
    m_rstAvailableOff.Fields.Append "Batch No", adVarWChar, 50, 102
    m_rstAvailableOff.Fields.Append "Qty for Outbound", adVarWChar, 50
    m_rstAvailableOff.Fields.Append "Job No", adVarWChar, 50, 102
    m_rstAvailableOff.Fields.Append "Outbound Edits", adDouble
    m_rstAvailableOff.Fields.Append "In_ID", adInteger, 4, 90
    m_rstAvailableOff.Fields.Append "Handling", adInteger, 1
    
    m_rstAvailableOff.Open

End Sub
Private Sub PopulateOutbounds(Optional ByVal Prod_IDs As String = "")
    Dim rstTemp As ADODB.Recordset
    Dim strSQL As String
            
            
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Outbounds!Out_ID AS Out_ID, "
        strSQL = strSQL & "Outbounds!In_ID AS In_ID, "
        strSQL = strSQL & "Products!Prod_ID AS Prod_ID, "
        strSQL = strSQL & "Products!Prod_Num AS [Product No], "
        strSQL = strSQL & "Products!Prod_Handling AS Handling, "
        strSQL = strSQL & "InboundDocs!InDoc_Type & '-' & InboundDocs!InDoc_Num AS [Document No], "
        strSQL = strSQL & "Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num AS [Entrepot No], "
        strSQL = strSQL & "StockCards!Stock_Card_Num AS [Stock Card No], "
        strSQL = strSQL & "Inbounds!In_Avl_Qty_Wgt AS [Quantity/Weight], "
        strSQL = strSQL & "Inbounds!In_Orig_Packages_Type AS [Package Type], "
        strSQL = strSQL & "Inbounds!In_Batch_Num AS [Batch No], "
        strSQL = strSQL & "Outbounds!Out_Packages_Qty_Wgt AS [Outbound Qty], "
        strSQL = strSQL & "Outbounds!Out_Job_Num AS [Job No] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "OutboundDocs INNER JOIN ("
        strSQL = strSQL & "Outbounds INNER JOIN ("
        strSQL = strSQL & "InboundDocs INNER JOIN ("
        strSQL = strSQL & "Inbounds INNER JOIN ("
        strSQL = strSQL & "Stockcards INNER JOIN ("
        strSQL = strSQL & "Products INNER JOIN Entrepots "
        strSQL = strSQL & "ON Products.Entrepot_ID = Entrepots.Entrepot_ID) "
        strSQL = strSQL & "ON Stockcards.Prod_ID = Products.Prod_ID) "
        strSQL = strSQL & "ON Inbounds.Stock_ID = Stockcards.Stock_ID) "
        strSQL = strSQL & "ON InboundDocs.InDoc_ID = Inbounds.Indoc_ID) "
        strSQL = strSQL & "ON Outbounds.In_ID = Inbounds.In_ID) "
        strSQL = strSQL & "ON OutboundDocs.OutDoc_ID = Outbounds.OutDoc_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "OutboundDocs!OutDoc_ID = " & lngOutDocID & " "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "OutboundDocs!OutDoc_Global = 1 "
        
        If (Len(Trim$(Prod_IDs)) > 0) Then
            strSQL = strSQL & "AND Products!Prod_ID IN (" & Prod_IDs & ") "
        End If
        
        If (Len(strCorrectionMode) = 1) Then
            strSQL = strSQL & "AND Outbounds!Out_Code = '" & strCorrectionMode & "Correction' "
        Else
            strSQL = strSQL & "AND (Outbounds!Out_Code IS NULL OR UCASE(MID(Outbounds!Out_Code,2)) <> 'CORRECTION') "
        End If
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.MoveFirst
    
        Do While Not rstTemp.EOF
            
            m_rstOutboundsOff.AddNew
                If (Len(Trim$(Prod_IDs)) = 0) Then
                    If (Len(strProdID) = 0) Then
                        strProdID = rstTemp![Prod_ID]
                    Else
                        strProdID = IIf(InStr(strProdID, rstTemp![Prod_ID]) > 0, strProdID, strProdID & "," & rstTemp![Prod_ID])
                    End If
                End If
                
                m_rstOutboundsOff![Out_ID] = rstTemp![Out_ID]
                m_rstOutboundsOff![In_ID] = rstTemp![In_ID]
                m_rstOutboundsOff![Product No] = rstTemp![Product No]
                m_rstOutboundsOff![Handling] = rstTemp![Handling]
                m_rstOutboundsOff![Entrepot No] = rstTemp![Entrepot No]
                m_rstOutboundsOff![Stock Card No] = rstTemp![Stock Card No]
                m_rstOutboundsOff![Quantity/Weight] = Replace(CStr(rstTemp![Quantity/Weight]), ",", ".")
                m_rstOutboundsOff![Package Type] = rstTemp![Package Type]
                m_rstOutboundsOff![Batch No] = rstTemp![Batch No]
                m_rstOutboundsOff![Outbound Qty] = IIf(strCorrectionMode = "I", Replace(CStr(rstTemp![Outbound Qty]), ",", ".") * -1, Replace(CStr(rstTemp![Outbound Qty]), ",", "."))
                m_rstOutboundsOff![Job No] = rstTemp![Job No]
                m_rstOutboundsOff![Document No] = rstTemp![Document No]
            m_rstOutboundsOff.Update
            
            rstTemp.MoveNext
        Loop
    End If
    
    ADORecordsetClose rstTemp
    
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Inbounds!In_ID AS Inbound_ID, "
        strSQL = strSQL & "Outbounds!Out_ID AS Out_ID, "
        strSQL = strSQL & "Outbounds!In_ID AS In_ID, "
        strSQL = strSQL & "Outbounds!Out_Batch_Num AS [Batch No], "
        strSQL = strSQL & "Outbounds!Out_Job_Num AS [Job No], "
        strSQL = strSQL & "Outbounds!Out_Packages_Qty_Wgt AS [Outbound Qty] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "OutboundDocs LEFT JOIN ("
        strSQL = strSQL & "Outbounds LEFT JOIN ("
        strSQL = strSQL & "Inbounds LEFT JOIN Stockcards "
        strSQL = strSQL & "ON Inbounds.Stock_ID = Stockcards.Stock_ID) "
        strSQL = strSQL & "ON Outbounds.In_ID = Inbounds.In_ID) "
        strSQL = strSQL & "ON OutboundDocs.OutDoc_ID = Outbounds.OutDoc_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "OutboundDocs!OutDoc_ID = " & lngOutDocID & " "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "OutboundDocs!OutDoc_Global = 1 "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "IsNull(Inbounds!In_ID) "
        strSQL = strSQL & "AND NOT IsNull(Outbounds!Out_ID) "
        
        If (Len(strCorrectionMode) = 1) Then
            strSQL = strSQL & "AND Outbounds!Out_Code = '" & strCorrectionMode & "Correction'"
        Else
            strSQL = strSQL & "AND (Outbounds!Out_Code IS NULL OR UCASE(MID(Outbounds!Out_Code,2)) <> 'CORRECTION')"
        End If
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.MoveFirst
        
        Do While Not rstTemp.EOF
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "InboundDocHistory" & "_" & Format(m_lngUserID, "00") & ".InDoc_Type & '-' & InboundDocHistory" & "_" & Format(m_lngUserID, "00") & ".InDoc_Num AS [Document No], "
                strSQL = strSQL & "Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num AS [Entrepot No], "
                strSQL = strSQL & "StockCards!Stock_Card_Num AS [Stock Card No], "
                strSQL = strSQL & "InboundHistory" & "_" & Format(m_lngUserID, "00") & ".In_Avl_Qty_Wgt AS [Quantity/Weight], "
                strSQL = strSQL & "InboundHistory" & "_" & Format(m_lngUserID, "00") & ".In_Orig_Packages_Type AS [Package Type], "
                strSQL = strSQL & "Products!Prod_ID AS Prod_ID, "
                strSQL = strSQL & "Products!Prod_Num AS [Product No], "
                strSQL = strSQL & "Products!Prod_Handling AS Handling "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "InboundDocHistory" & "_" & Format(m_lngUserID, "00") & " INNER JOIN ("
                strSQL = strSQL & "InboundHistory" & "_" & Format(m_lngUserID, "00") & " INNER JOIN ("
                strSQL = strSQL & "StockCards INNER JOIN ("
                strSQL = strSQL & "Products INNER JOIN Entrepots "
                strSQL = strSQL & "ON Products.Entrepot_ID = Entrepots.Entrepot_ID) "
                strSQL = strSQL & "ON Stockcards.Prod_ID = Products.Prod_ID) "
                strSQL = strSQL & "ON InboundHistory" & "_" & Format(m_lngUserID, "00") & ".Stock_ID = Stockcards.Stock_ID) "
                strSQL = strSQL & "ON InboundHistory" & "_" & Format(m_lngUserID, "00") & ".InDoc_ID = InboundDocHistory" & "_" & Format(m_lngUserID, "00") & ".InDoc_ID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "InboundHistory" & "_" & Format(m_lngUserID, "00") & ".In_ID = " & rstTemp!In_ID & " "
                
                If (Len(Trim$(Prod_IDs)) > 0) Then
                    strSQL = strSQL & "AND StockCards!Prod_ID IN (" & Prod_IDs & ") "
                End If
                
                strSQL = strSQL & "ORDER BY InboundHistory" & "_" & Format(m_lngUserID, "00") & ".SourceDB"
            
            ADORecordsetOpen strSQL, m_conSADBEL, m_rstHistory, adOpenKeyset, adLockOptimistic
            If Not (m_rstHistory.EOF And m_rstHistory.BOF) Then
                m_rstHistory.MoveFirst
                
                If (Len(Trim$(Prod_IDs)) = 0) Then
                    If (Len(strProdID) = 0) Then
                        strProdID = m_rstHistory!Prod_ID
                    Else
                        strProdID = IIf(InStr(strProdID, m_rstHistory!Prod_ID) > 0, strProdID, strProdID & "," & m_rstHistory!Prod_ID)
                    End If
                End If
                
                m_rstOutboundsOff.AddNew
                    m_rstOutboundsOff![Out_ID] = rstTemp![Out_ID]
                    m_rstOutboundsOff![In_ID] = rstTemp![In_ID]
                    m_rstOutboundsOff![Batch No] = rstTemp![Batch No]
                    m_rstOutboundsOff![Outbound Qty] = IIf(strCorrectionMode = "I", Replace(CStr(rstTemp![Outbound Qty]), ",", ".") * -1, Replace(CStr(rstTemp![Outbound Qty]), ",", "."))
                    m_rstOutboundsOff![Job No] = rstTemp![Job No]
                    
                    m_rstOutboundsOff![Product No] = m_rstHistory![Product No]
                    m_rstOutboundsOff![Handling] = m_rstHistory![Handling]
                    m_rstOutboundsOff![Document No] = m_rstHistory![Document No]
                    m_rstOutboundsOff![Entrepot No] = m_rstHistory![Entrepot No]
                    m_rstOutboundsOff![Stock Card No] = m_rstHistory![Stock Card No]
                    m_rstOutboundsOff![Quantity/Weight] = Replace(CStr(m_rstHistory![Quantity/Weight]), ",", ".")
                    m_rstOutboundsOff![Package Type] = m_rstHistory![Package Type]
                m_rstOutboundsOff.Update
                
            End If
            
            ADORecordsetClose m_rstHistory
            
            rstTemp.MoveNext
        Loop
    End If
    
    ADORecordsetClose rstTemp
    
    
    If (m_rstOutboundsOff.RecordCount > 0 Or jgxOutbounds.RowCount > 0) Then
        blnSystemChanged = True
        Set jgxOutbounds.ADORecordset = Nothing
        Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
        blnSystemChanged = False
    End If
    
    Call FormatOutbounds
    
End Sub

Private Function ValidateOutbounds() As Boolean
    
    Select Case cboCodiType.ListIndex
        Case 0, 1, 2
            If (Len(Trim$(txtDocNum.Text)) = 0 Or Len(Trim$(cboDocType.Text)) = 0) Then
                MsgBox "Please provide data for the " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " Document fields.", vbOKOnly + vbInformation, Me.Caption '"Manual Outbound"
                ValidateOutbounds = False
                Exit Function
            End If
        Case 3, 5
            If Len(Trim$(txtMRN.Text)) = 0 Then
                MsgBox "Please provide data for the " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " Document fields.", vbOKOnly + vbInformation, Me.Caption '"Manual Outbound"
                ValidateOutbounds = False
                Exit Function
            End If
        Case 4
            If Len(Trim$(txtDocNum.Text)) = 0 Or Len(Trim$(cboDocType.Text)) = 0 Or Len(Trim$(txtMRN.Text)) = 0 Then
                MsgBox "Please provide data for the " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " Document fields.", vbOKOnly + vbInformation, Me.Caption '"Manual Outbound"
                ValidateOutbounds = False
                Exit Function
            End If
        
        Case 6, 7 ' PLDA Import, PLDA Export
            If Len(Trim$(txtDocNum.Text)) = 0 Or Len(Trim$(cboDocType.Text)) = 0 Then
                MsgBox "Please provide data for the " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " Document fields.", vbOKOnly + vbInformation, Me.Caption '"Manual Outbound"
                ValidateOutbounds = False
                Exit Function
            End If
    End Select
        
    If Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_History" & Right(dtpDate.Year, 2) & ".mdb") = "" Then

        MsgBox "The year entered in the " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " Document field is not valid.", vbInformation + vbOKOnly, Me.Caption '"Manual Outbound"
        ValidateOutbounds = False
        Exit Function
    End If
    
    If Trim(txtCommunalSettlement.Text) = "" Then
        MsgBox "Please provide value for Communal Settlement.", vbOKOnly + vbInformation, Me.Caption
        ValidateOutbounds = False
        Exit Function
    End If
    
    If Not (m_rstAvailableOff.EOF And m_rstAvailableOff.BOF) Then
'        jgxAvailableStock.Update
        m_rstAvailableOff.MoveFirst
    End If
    
    If strCorrectionMode = "I" Then
        m_rstAvailableOff.Filter = "[Quantity/Weight] > 0"
    Else
        m_rstAvailableOff.Filter = "[Qty for Outbound] > 0"
    End If
    
    Do While Not m_rstAvailableOff.EOF
        
        If Val(m_rstAvailableOff![Quantity/Weight]) < 0 Then
            'CSCLP-233 - Added IF..THEN
            If strCorrectionMode = "I" And m_rstAvailableOff.Fields("Outbound Edits").Value > 0 Then
                MsgBox "Quantity for  " & IIf(strCorrectionMode = "I", "Inbound", "Outbound") & " should not be greater than the Available Quantity." & vbCrLf & "Please check grid values.", vbExclamation & vbOKOnly, Me.Caption '"Manual Outbound"
                ValidateOutbounds = False
                Exit Function
            End If
        End If
        
        m_rstAvailableOff.MoveNext
        
    Loop
    m_rstAvailableOff.Filter = adFilterNone
    
    ValidateOutbounds = True
    
End Function

Private Sub ApplyChanges()
                
    Dim lngCtr As Long
    Dim lngrow As Long
    Dim rstTemp As ADODB.Recordset
    Dim lngIndex As Long
    Dim lngSort As Long
    
    lngrow = jgxAvailableStock.Row
    m_rstAvailableOff.Filter = adFilterNone
    If strCorrectionMode <> "I" Then
        If chkShowZero.Value = vbUnchecked Then
            m_rstAvailableOff.Filter = "[Qty for Outbound] > 0 OR [Outbound Edits] <> 0 "
        End If
    End If
    If Not (m_rstAvailableOff.EOF And m_rstAvailableOff.BOF) Then
        m_rstAvailableOff.MoveFirst
    
    
        Do While Not m_rstAvailableOff.EOF
                    
            ADORecordsetOpen "SELECT * FROM InboundHistory" & "_" & Format(m_lngUserID, "00") & " WHERE InboundHistory" & "_" & Format(m_lngUserID, "00") & ".In_ID = " & m_rstAvailableOff!In_ID & " ORDER BY SourceDB DESC", m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
            'rstTemp.Open "SELECT * FROM InboundHistory" & "_" & Format(m_lngUserID, "00") & " WHERE InboundHistory" & "_" & Format(m_lngUserID, "00") & ".In_ID = " & m_rstAvailableOff!In_ID & " ORDER BY SourceDB DESC", m_conSADBEL, adOpenKeyset, adLockOptimistic
                                    
            If Not (rstTemp.EOF And rstTemp.BOF) Then
                rstTemp.MoveFirst
                
                ADOConnectDB m_conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(rstTemp!SourceDB)
                'OpenADODatabase m_conHistory, strMDBpath & "\" & rstTemp!SourceDB & ".mdb"
                
                ADORecordsetOpen "SELECT * FROM Inbounds WHERE Inbounds!In_ID = " & m_rstAvailableOff!In_ID, _
                                m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
                'm_rstHistory.Open "SELECT * FROM Inbounds WHERE Inbounds!In_ID = " & m_rstAvailableOff!In_ID, _
                                m_conHistory, adOpenKeyset, adLockOptimistic
                
                ADORecordsetOpen "SELECT * FROM Inbounds Where Inbounds!In_ID = " & m_rstAvailableOff!In_ID, _
                                m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
                'm_rstSadbel.Open "SELECT * FROM Inbounds Where Inbounds!In_ID = " & m_rstAvailableOff!In_ID, _
                                m_conSADBEL, adOpenKeyset, adLockOptimistic
                
                'CSCLP-233
                'If m_rstHistory.RecordCount > 0 And m_rstSadbel.RecordCount > 0 Then
                If Not (m_rstHistory.EOF And m_rstHistory.BOF) Or Not (m_rstSadbel.EOF And m_rstSadbel.EOF) Then
                
                    If Val(m_rstAvailableOff![Qty for Outbound]) > 0 Or _
                        (Val(m_rstAvailableOff![Qty for Outbound]) < 0 And _
                        Len(strCorrectionMode) > 0) Then
                        
                        m_rstOutboundsOff.AddNew
                        m_rstOutboundsOff!Out_ID = 0
                        m_rstOutboundsOff!In_ID = m_rstAvailableOff!In_ID
                        m_rstOutboundsOff![Product No] = m_rstAvailableOff![Product No]
                        m_rstOutboundsOff!Handling = m_rstAvailableOff!Handling
                        m_rstOutboundsOff![Document No] = m_rstAvailableOff![Document No]
                        m_rstOutboundsOff![Entrepot No] = m_rstAvailableOff![Entrepot No]
                        m_rstOutboundsOff![Stock Card No] = m_rstAvailableOff![Stock Card No]
                        m_rstOutboundsOff![Package Type] = m_rstAvailableOff![Package Type]
                        m_rstOutboundsOff![Batch No] = m_rstAvailableOff![Batch No]
                        m_rstOutboundsOff![Outbound Qty] = m_rstAvailableOff![Qty for Outbound]
                        m_rstOutboundsOff![Job No] = m_rstAvailableOff![Job No]
                        m_rstOutboundsOff.Update
                                    
                    End If
                    
                    If strCorrectionMode = "I" Then
                        
                        'New available stocks = Old available stocks minus correction value
                        m_rstHistory!In_Avl_Qty_Wgt = Val(m_rstAvailableOff![Quantity/Weight]) 'm_rstHistory!In_Avl_Qty_Wgt  'Val(m_rstOutboundsOff![Outbound Qty])
                    Else
                        m_rstHistory!In_Avl_Qty_Wgt = Val(m_rstAvailableOff![Quantity/Weight])
                        m_rstHistory!In_TotalOut_Qty_Wgt = Round(m_rstHistory!In_TotalOut_Qty_Wgt + Val(m_rstAvailableOff![Qty for Outbound]) + m_rstAvailableOff![Outbound Edits], Choose(lngHandling + 1, 0, 2, 3))
                    End If
                    
                    rstTemp!In_Avl_Qty_Wgt = m_rstHistory!In_Avl_Qty_Wgt
                    
                    m_rstOutboundsOff.Filter = adFilterNone
                    m_rstOutboundsOff.Filter = "In_ID = " & m_rstAvailableOff!In_ID
                    
                    If m_rstOutboundsOff.RecordCount > 0 Then
                        m_rstOutboundsOff.MoveFirst
                    End If
                    
                    Do While Not m_rstOutboundsOff.EOF
                        m_rstOutboundsOff![Quantity/Weight] = m_rstAvailableOff![Quantity/Weight]
                        m_rstOutboundsOff.MoveNext
                    Loop
                    
                    m_rstOutboundsOff.Filter = adFilterNone
                    
                    'CSCLP-233 - added IF..THEN
                    If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                        m_rstSadbel.MoveFirst
                        
                        If Val(m_rstAvailableOff![Quantity/Weight]) <= 0 Then
                            If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                                If m_rstSadbel!In_Reserved_Qty_Wgt = 0 Then
                                    m_rstSadbel.Delete
                                    m_rstSadbel.Update
                                    
                                    ExecuteNonQuery m_conSADBEL, "DELETE * FROM Inbounds Where Inbounds!In_ID = " & m_rstAvailableOff!In_ID
                                Else
                                    If strCorrectionMode = "I" Then
                                        m_rstSadbel!In_Avl_Qty_Wgt = m_rstSadbel!In_Avl_Qty_Wgt - Val(m_rstOutboundsOff![Outbound Qty])
                                    Else
                                        m_rstSadbel!In_Avl_Qty_Wgt = Val(m_rstAvailableOff![Quantity/Weight])
                                        m_rstSadbel!In_TotalOut_Qty_Wgt = Round(m_rstSadbel!In_TotalOut_Qty_Wgt + Val(m_rstAvailableOff![Qty for Outbound]) + m_rstAvailableOff![Outbound Edits], Choose(lngHandling + 1, 0, 2, 3))
                                    End If
                                    m_rstSadbel.Update
                                    
                                    UpdateRecordset m_conSADBEL, m_rstSadbel, "Inbounds"
                                End If
                                
                                m_rstAvailableOff![Qty for Outbound] = ""
                                m_rstAvailableOff![Outbound Edits] = Empty
                                
                            End If
                            'This conditions allows Inbounds Correction to retain the record even when it is zeroed out
                            If Len(strCorrectionMode) = 0 Then
                                m_rstAvailableOff.Delete 'If chkShowZero.Value = vbUnchecked Then
                            End If
                        Else
                            If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                                m_rstSadbel.MoveFirst
                                
                                If strCorrectionMode = "I" Then
                                    m_rstSadbel!In_Avl_Qty_Wgt = Val(m_rstAvailableOff![Quantity/Weight]) 'm_rstSadbel!In_Avl_Qty_Wgt - Val(m_rstOutboundsOff![Outbound Qty])
                                Else
                                    m_rstSadbel!In_Avl_Qty_Wgt = Val(m_rstAvailableOff![Quantity/Weight])
                                    m_rstSadbel!In_TotalOut_Qty_Wgt = Round(m_rstSadbel!In_TotalOut_Qty_Wgt + Val(m_rstAvailableOff![Qty for Outbound]) + m_rstAvailableOff![Outbound Edits], Choose(lngHandling + 1, 0, 2, 3))
                                End If
                                m_rstSadbel.Update
                                
                                UpdateRecordset m_conSADBEL, m_rstSadbel, "Inbounds"
                            Else
                                m_rstSadbel.AddNew
                                For lngCtr = 0 To m_rstHistory.Fields.Count - 1
                                    m_rstSadbel.Fields(m_rstHistory.Fields(lngCtr).Name).Value = m_rstHistory.Fields(lngCtr).Value
                                Next
                                m_rstSadbel.Update
                                
                                InsertRecordset m_conSADBEL, m_rstSadbel, "Inbounds"
                            End If
                            
                            m_rstAvailableOff![Qty for Outbound] = ""
                            m_rstAvailableOff![Outbound Edits] = Empty

                        End If
                    End If
                    
                    m_rstHistory.Update
                                        
                    UpdateRecordset m_conHistory, m_rstHistory, "Inbounds"
                    
                    rstTemp.Update
                    
                    UpdateRecordset m_conSADBEL, rstTemp, "InboundHistory" & "_" & Format(m_lngUserID, "00")
                End If
                
                ADORecordsetClose m_rstHistory
                ADORecordsetClose m_rstSadbel
                
                ADODisconnectDB m_conHistory
            End If
            
            ADORecordsetClose rstTemp
            
            m_rstAvailableOff.MoveNext
            
        Loop
    End If
    
    m_rstAvailableOff.Filter = adFilterNone
    
    
    ADORecordsetClose m_rstSadbel
    ADORecordsetClose m_rstHistory
    ADORecordsetClose rstTemp
    
    ADODisconnectDB m_conHistory
    
    If jgxAvailableStock.SortKeys.Count > 0 Then
        lngIndex = jgxAvailableStock.SortKeys(1).ColIndex
        lngSort = jgxAvailableStock.SortKeys(1).SortOrder
    End If
    Set jgxAvailableStock.ADORecordset = Nothing
    Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
    Call FormatAvailable
    If lngIndex > 0 Then
        jgxAvailableStock.SortKeys.Add lngIndex, lngSort
        jgxAvailableStock.RefreshSort
        lngIndex = 0
    End If
    
    If blnAvailableIsLastFocus Then
        jgxAvailableStock.SetFocus
        blnAvailableIsLastFocus = False
    End If
    
    jgxAvailableStock.Row = IIf(lngrow > jgxAvailableStock.RowCount, jgxAvailableStock.RowCount, lngrow)
    jgxAvailableStock.Col = jgxAvailableStock.Columns("Qty for Outbound").Index
    jgxAvailableStock.SelStart = 0
    jgxAvailableStock.SelLength = 0
    jgxAvailableStock.EnsureVisible
    
    Call UpdateOutbounds
    
    If jgxOutbounds.SortKeys.Count > 0 Then
        lngIndex = jgxOutbounds.SortKeys(1).ColIndex
        lngSort = jgxOutbounds.SortKeys(1).SortOrder
    End If
    Set jgxOutbounds.ADORecordset = Nothing
    Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
    Call FormatOutbounds
    If lngIndex > 0 Then
        jgxOutbounds.SortKeys.Add lngIndex, lngSort
        lngIndex = 0
    End If
    
    'Re-applies filter if show zero is checked
    If chkShowZero.Value = Unchecked Then
        m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR [Outbound Edits] <> 0"
    
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        Call FormatAvailable
    End If
End Sub

Private Sub UpdateOutbounds()
    
    Dim lngCtr As Long
    Dim strSQL As String
    Dim bytCounter As Byte
    Dim bytNumberofZeros As Byte
    Dim strZeros As String
    Dim strHistoryYear As String
    Dim blnAddNew As Boolean
    Dim lngOutID As Long
    
    strHistoryYear = Right(dtpDate.Year, 2)
    
    If Len(Trim(txtDocNum.Text)) < 7 And Len(Trim(txtDocNum.Text)) <> 0 Then
        bytNumberofZeros = 7 - Len(txtDocNum.Text)
        For bytCounter = 1 To bytNumberofZeros
            strZeros = "0" & strZeros
        Next bytCounter
    End If
    
    ADOConnectDB m_conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, strHistoryYear
    'OpenADODatabase m_conHistory, strMDBpath & "\mdb_History" & strHistoryYear & ".mdb"
    
    If lngOutDocID = 0 And m_rstOutboundsOff.RecordCount > 0 Then
    
        strSQL = "SELECT * FROM OutboundDocs WHERE "
        strSQL = strSQL & IIf(cboDocType.Text <> "", "UCASE(OutDoc_Type) = '" & UCase(cboDocType.Text) & "'", "(ISNULL(OutDoc_Type) OR OutDoc_Type='')") & " AND "
        strSQL = strSQL & IIf(Trim$(txtDocNum.Text) <> "", "UCASE(OutDoc_Num) = '" & strZeros & UCase(txtDocNum.Text) & "'", "(ISNULL(OutDoc_Num) OR OutDoc_Num='')") & " AND "
        strSQL = strSQL & IIf(txtMRN.Text <> "", "UCASE(OutDoc_MRN) = '" & UCase(txtMRN.Text) & "'", "(ISNULL(OutDoc_MRN) OR OutDoc_MRN='')") & " AND "
        strSQL = strSQL & "OutDoc_Global=1"

        ADORecordsetOpen strSQL, m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
        'm_rstHistory.Open strSQL, m_conHistory, adOpenKeyset, adLockOptimistic
        
        ADORecordsetOpen "SELECT * FROM OutboundDocs", m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
        'm_rstSadbel.Open "SELECT * FROM OutboundDocs", m_conSADBEL, adOpenKeyset, adLockOptimistic
        
        If m_rstHistory.EOF And m_rstHistory.BOF Then
            
            ADORecordsetClose m_rstHistory
            
            strYear = dtpDate.Year
            
            m_rstSadbel.AddNew
            m_rstSadbel!OutDoc_Type = IIf(IsNull(cboDocType.Text), "", cboDocType.Text)
            m_rstSadbel!OutDoc_Num = IIf(IsNull(txtDocNum.Text), "", txtDocNum.Text)
            m_rstSadbel!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
            m_rstSadbel!OutDoc_MRN = IIf(IsNull(txtMRN.Text), "", txtMRN.Text)
            m_rstSadbel!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
            m_rstSadbel!OutDoc_Global = 1
            m_rstSadbel.Update
        
            lngOutDocID = InsertRecordset(m_conSADBEL, m_rstSadbel, "OutboundDocs")
            
            ADORecordsetOpen "SELECT * FROM OutBoundDocs WHERE OutDoc_ID = " & lngOutDocID, m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
            'm_rstHistory.Open "SELECT * FROM OutBoundDocs WHERE OutDoc_ID = " & lngOutDocID, m_conHistory, adOpenKeyset, adLockOptimistic
            
            ' GET AN UNUSED OutDoc_ID
            Do While m_rstHistory.RecordCount > 0
                
                m_rstSadbel.Filter = adFilterNone
                m_rstSadbel.Filter = "[OutDoc_ID] = " & lngOutDocID & " "
                If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                    m_rstSadbel.MoveFirst
                    m_rstSadbel.Delete
                    m_rstSadbel.Update
                End If
                
                ExecuteNonQuery m_conSADBEL, "DELETE * FROM [OutboundDocs] WHERE [OutDoc_ID] = " & lngOutDocID & " "
                
                m_rstSadbel.AddNew
                m_rstSadbel!OutDoc_Type = IIf(IsNull(cboDocType.Text), "", cboDocType.Text)
                m_rstSadbel!OutDoc_Num = IIf(IsNull(txtDocNum.Text), "", txtDocNum.Text)
                m_rstSadbel!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
                m_rstSadbel!OutDoc_MRN = IIf(IsNull(txtMRN.Text), "", txtMRN.Text)
                m_rstSadbel!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
                m_rstSadbel!OutDoc_Global = 1
                m_rstSadbel.Update
            
                lngOutDocID = InsertRecordset(m_conSADBEL, m_rstSadbel, "OutboundDocs")
                
                ADORecordsetClose m_rstHistory
                ADORecordsetOpen "SELECT * FROM OutBoundDocs WHERE OutDoc_ID = " & lngOutDocID, m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
                'm_rstHistory.Open "SELECT * FROM OutBoundDocs WHERE OutDoc_ID = " & m_rstSadbel!outDoc_ID, m_conHistory, adOpenKeyset, adLockOptimistic
            Loop

            'Glenn 3/30/2006
            If lngOutDocID = 0 Then
            'If m_rstSadbel!outDoc_ID = 0 Then
                MsgBox "Failed adding outbound document properly. Please contact your administrator.", vbInformation, IIf(Len(strCorrectionMode) = 0, "Manual Outbound", strCorrectionMode & "Correction")
            End If
            
            m_rstHistory.AddNew
            m_rstHistory!OutDoc_ID = lngOutDocID
            m_rstHistory!OutDoc_Type = IIf(IsNull(cboDocType.Text), "", cboDocType.Text)
            m_rstHistory!OutDoc_Num = IIf(IsNull(txtDocNum.Text), "", txtDocNum.Text)
            m_rstHistory!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
            m_rstHistory!OutDoc_MRN = IIf(IsNull(txtMRN.Text), "", txtMRN.Text)
            m_rstHistory!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
            m_rstHistory!OutDoc_Global = 1
            m_rstHistory.Update
            
            InsertRecordset m_conHistory, m_rstHistory, "OutboundDocs"
        
        Else
            m_rstSadbel.AddNew
            
            For lngCtr = 0 To m_rstHistory.Fields.Count - 1
                m_rstSadbel.Fields(lngCtr).Value = m_rstHistory.Fields(m_rstSadbel.Fields(lngCtr).Name).Value
            Next
            
            lngOutDocID = m_rstHistory!OutDoc_ID
            
            strYear = dtpDate.Year
            
            m_rstHistory!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
            m_rstSadbel!OutDoc_Date = m_rstHistory!OutDoc_Date
            m_rstHistory!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
            m_rstSadbel!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
                
            m_rstSadbel.Update
            
            InsertRecordset m_conSADBEL, m_rstSadbel, "OutboundDocs"
            
            m_rstHistory.Update
            
            UpdateRecordset m_conHistory, m_rstHistory, "OutboundDocs"
        End If
        
        ADORecordsetClose m_rstSadbel
        ADORecordsetClose m_rstHistory
    
    
    ElseIf lngOutDocID <> 0 Then
        
        ADORecordsetOpen "SELECT OutDoc_Date, OutDoc_Comm_Settlement FROM OutboundDocs WHERE OutDoc_ID=" & lngOutDocID, m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
        'm_rstSadbel.Open "SELECT OutDoc_Date, OutDoc_Comm_Settlement " & _
                "FROM OutboundDocs WHERE OutDoc_ID=" & lngOutDocID, m_conSADBEL, adOpenKeyset, adLockOptimistic
        
        m_rstSadbel!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
        m_rstSadbel!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
        m_rstSadbel.Update
        
        UpdateRecordset m_conSADBEL, m_rstSadbel, "OutboundDocs"
        
        ADORecordsetOpen "SELECT OutDoc_Date, OutDoc_Comm_Settlement FROM OutboundDocs WHERE OutDoc_ID=" & lngOutDocID, m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
        'm_rstHistory.Open "SELECT OutDoc_Date, OutDoc_Comm_Settlement " & _
                "FROM OutboundDocs WHERE OutDoc_ID=" & lngOutDocID, m_conHistory, adOpenKeyset, adLockOptimistic
        
        m_rstHistory!OutDoc_Date = Format(dtpDate.Value, "Short Date") & " " & Time
        m_rstHistory!OutDoc_Comm_Settlement = IIf(IsNull(txtCommunalSettlement.Text), "", txtCommunalSettlement.Text)
        m_rstHistory.Update
        
        UpdateRecordset m_conHistory, m_rstHistory, "OutboundDocs"
        
        ADORecordsetClose m_rstSadbel
        ADORecordsetClose m_rstHistory
        
    End If
    
    ADORecordsetOpen "SELECT * FROM Outbounds WHERE Outbounds!OutDoc_ID = " & lngOutDocID, m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
    'm_rstSadbel.Open "SELECT * FROM Outbounds WHERE Outbounds!OutDoc_ID = " & lngOutDocID, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    ADORecordsetOpen "SELECT * FROM Outbounds", m_conHistory, m_rstHistory, adOpenKeyset, adLockOptimistic
    'm_rstHistory.Open "SELECT * FROM Outbounds", m_conHistory, adOpenKeyset, adLockOptimistic
    
    If Not (m_rstOutboundsOff.EOF And m_rstOutboundsOff.BOF) Then
        
        m_rstOutboundsOff.MoveFirst
    
        Do While Not m_rstOutboundsOff.EOF
                    
            If m_rstOutboundsOff!Out_ID = 0 Then
                If Val(m_rstOutboundsOff![Outbound Qty]) > 0 Or (Val(m_rstOutboundsOff![Outbound Qty]) < 0 And Len(strCorrectionMode) > 0) Then
                    
                    m_rstSadbel.AddNew
                    m_rstSadbel!In_ID = m_rstOutboundsOff!In_ID
                    If Len(strCorrectionMode) > 0 Then
                        m_rstSadbel!Out_Code = strCorrectionMode & "Correction"
                    End If
                    m_rstSadbel!Out_Batch_Num = m_rstOutboundsOff![Batch No]
                    m_rstSadbel!Out_Job_Num = m_rstOutboundsOff![Job No]
                    m_rstSadbel!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstSadbel!OutDoc_ID = lngOutDocID
                
                    'Glenn 3/30/2006
                    If m_rstSadbel!OutDoc_ID = 0 Then
                        MsgBox "Failed to save outbound properly. Please contact your administrator.", vbInformation, IIf(Len(strCorrectionMode) = 0, "Manual Outbound", strCorrectionMode & "Correction")
                    End If
                    m_rstSadbel.Update
                    
                    lngOutID = InsertRecordset(m_conSADBEL, m_rstSadbel, "Outbounds")
                    
                    m_rstHistory.Filter = adFilterNone
                    m_rstHistory.Filter = "Out_ID = " & lngOutID
                    Do While m_rstHistory.RecordCount > 0
                        m_rstSadbel.Filter = adFilterNone
                        m_rstSadbel.Filter = "[OutID] = " & lngOutID & " "
                        If Not (m_rstSadbel.EOF And m_rstSadbel.BOF) Then
                            m_rstSadbel.MoveFirst
                            m_rstSadbel.Delete
                            m_rstSadbel.Update
                        End If
                        m_rstSadbel.Filter = adFilterNone
                        
                        ExecuteNonQuery m_conSADBEL, "DELETE * FROM [Outbounds] WHERE [Outbounds].[OutDoc_ID] = " & lngOutDocID & " AND [Out_ID] = " & lngOutID
                    
                        m_rstSadbel.AddNew
                        m_rstSadbel!In_ID = m_rstOutboundsOff!In_ID
                        If Len(strCorrectionMode) > 0 Then
                            m_rstSadbel!Out_Code = strCorrectionMode & "Correction"
                        End If
                        m_rstSadbel!Out_Batch_Num = m_rstOutboundsOff![Batch No]
                        m_rstSadbel!Out_Job_Num = m_rstOutboundsOff![Job No]
                        m_rstSadbel!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                        m_rstSadbel!OutDoc_ID = lngOutDocID
                    
                        'Glenn 3/30/2006
                        If m_rstSadbel!OutDoc_ID = 0 Then
                            MsgBox "Failed to save outbound properly. Please contact your administrator.", vbInformation, IIf(Len(strCorrectionMode) = 0, "Manual Outbound", strCorrectionMode & "Correction")
                        End If
                        m_rstSadbel.Update
                        
                        lngOutID = InsertRecordset(m_conSADBEL, m_rstSadbel, "Outbounds")
                        
                        m_rstHistory.Filter = adFilterNone
                        m_rstHistory.Filter = "Out_ID = " & lngOutID
                    Loop
                    m_rstHistory.Filter = adFilterNone
                    
                    
                    m_rstHistory.AddNew
                    m_rstHistory!Out_ID = lngOutID
                    
                    If Len(strCorrectionMode) > 0 Then
                        m_rstHistory!Out_Code = strCorrectionMode & "Correction"
                    End If
                    
                    m_rstHistory!In_ID = m_rstOutboundsOff!In_ID
                    m_rstHistory!Out_Batch_Num = m_rstOutboundsOff![Batch No]
                    m_rstHistory!Out_Job_Num = m_rstOutboundsOff![Job No]
                    m_rstHistory!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstHistory!OutDoc_ID = lngOutDocID
                    
                    'Glenn 3/30/2006
                    If m_rstHistory!OutDoc_ID = 0 Then
                        MsgBox "Failed to save outbound properly. Please contact your administrator.", vbInformation, IIf(Len(strCorrectionMode) = 0, "Manual Outbound", strCorrectionMode & "Correction")
                    End If
        
                    m_rstOutboundsOff!Out_ID = m_rstSadbel!Out_ID

                    m_rstHistory.Update
                    
                    UpdateRecordset m_conHistory, m_rstHistory, "Outbounds"
                    
                    m_rstOutboundsOff.Update
                End If
                
            Else
                m_rstSadbel.Filter = "Out_ID = " & m_rstOutboundsOff!Out_ID
                m_rstHistory.Filter = "Out_ID = " & m_rstOutboundsOff!Out_ID
                                
                If Val(m_rstOutboundsOff![Outbound Qty]) > 0 Then
                    m_rstSadbel!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstSadbel!Out_Job_Num = m_rstOutboundsOff![Job No]
                    m_rstHistory!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstHistory!Out_Job_Num = m_rstOutboundsOff![Job No]
                    
                    m_rstSadbel.Update
                    UpdateRecordset m_conSADBEL, m_rstSadbel, "Outbounds"
                    
                    m_rstHistory.Update
                    UpdateRecordset m_conHistory, m_rstHistory, "Outbounds"
                
                ElseIf Len(strCorrectionMode) > 0 Then
                    m_rstSadbel!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstSadbel!Out_Job_Num = m_rstOutboundsOff![Job No]
                    m_rstHistory!Out_Packages_Qty_Wgt = IIf(strCorrectionMode = "I", Val(m_rstOutboundsOff![Outbound Qty]) * -1, Val(m_rstOutboundsOff![Outbound Qty]))
                    m_rstHistory!Out_Job_Num = m_rstOutboundsOff![Job No]
                    
                    m_rstSadbel.Update
                    UpdateRecordset m_conSADBEL, m_rstSadbel, "Outbounds"
                    
                    m_rstHistory.Update
                    UpdateRecordset m_conHistory, m_rstHistory, "Outbounds"
                Else
                    m_rstSadbel.Delete
                    m_rstSadbel.Update
                    
                    ExecuteNonQuery m_conSADBEL, "DELETE * FROM [Outbounds] WHERE [Outbounds].[OutDoc_ID] = " & lngOutDocID & " AND [Out_ID] = " & m_rstOutboundsOff!Out_ID & " "
                    
                    m_rstHistory.Delete
                    m_rstHistory.Update
                    
                    ExecuteNonQuery m_conHistory, "DELETE * FROM [Outbounds] WHERE [Out_ID] = " & m_rstOutboundsOff!Out_ID & " "
                    
                    m_rstOutboundsOff.Delete
                End If
                
                
                
                m_rstSadbel.Filter = adFilterNone
                m_rstHistory.Filter = adFilterNone
                
            End If
            
            m_rstOutboundsOff.MoveNext
            
        Loop
    
    End If
    
    For lngCtr = 1 To UBound(alngDeleted)
        ExecuteNonQuery m_conSADBEL, "DELETE FROM Outbounds WHERE Out_ID = " & alngDeleted(lngCtr)
        ExecuteNonQuery m_conHistory, "DELETE FROM Outbounds WHERE Out_ID = " & alngDeleted(lngCtr)
    Next
    
    ReDim alngDeleted(0)
    
    If m_rstOutboundsOff.RecordCount = 0 And lngOutDocID <> 0 Then
        ADORecordsetClose m_rstSadbel
        
        ADORecordsetOpen "SELECT Out_ID FROM Outbounds WHERE OutDoc_ID = " & lngOutDocID, m_conSADBEL, m_rstSadbel, adOpenKeyset, adLockOptimistic
        'm_rstSadbel.Open "SELECT Out_ID FROM Outbounds WHERE OutDoc_ID = " & lngOutDocID, m_conSADBEL, adOpenKeyset, adLockOptimistic
        If m_rstSadbel.RecordCount = 0 Then
            ExecuteNonQuery m_conSADBEL, "DELETE FROM OutboundDocs WHERE OutDoc_ID = " & lngOutDocID
            ExecuteNonQuery m_conHistory, "DELETE FROM OutboundDocs WHERE OutDoc_ID = " & lngOutDocID
            lngOutDocID = 0
        End If
    End If
    
    ADORecordsetClose m_rstSadbel
    
    ADORecordsetClose m_rstHistory
    
    ADODisconnectDB m_conHistory
    
End Sub

Private Sub CheckYear()
    If lngOutDocID <> 0 Then
        If dtpDate.Year <> strYear Then
'            dtpDate.Year = strYear
        End If
    End If
End Sub

Private Sub cbodoctype_Click()
    
    If Trim(UCase(cboDocType.Text)) <> Trim(UCase(strLastDocType)) Then
    
        If lngOutDocID <> 0 Then

            If Not SavedIfChanged Then
                cboDocType.Text = strLastDocType
                Exit Sub
            End If
            
            blnSystemChanged = True
            txtProductNum.Text = ""
            Call ResetProduct
            blnSystemChanged = False
                        
            strLastDocNum = ""
            strLastDocType = ""
            strLastMRN = ""
            m_dteLastDate = 0
            
            Set jgxAvailableStock.ADORecordset = Nothing
            Call CreateAvailableOffline
            Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
            Call FormatAvailable
    
            Set jgxOutbounds.ADORecordset = Nothing
            Call CreateOutboundsOffline
            Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
            Call FormatOutbounds
        
        End If
        
        Call CheckOutDoc
        
    End If

    If jgxAvailableStock.RowCount > 0 Or jgxOutbounds.RowCount > 0 Then
        cmdApply.Enabled = True
    End If
    
End Sub

Private Sub cbodoctype_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMRN_Change()

    If Trim(UCase(txtMRN.Text)) <> Trim(UCase(strLastMRN)) Then
    
        If lngOutDocID <> 0 Then

            If Not SavedIfChanged Then
                txtMRN.Text = strLastMRN
                Exit Sub
            End If
            
            blnSystemChanged = True
            txtProductNum.Text = ""
            Call ResetProduct
            blnSystemChanged = False
            
            strLastDocNum = ""
            strLastDocType = ""
            strLastMRN = ""
            m_dteLastDate = 0
            
            Set jgxAvailableStock.ADORecordset = Nothing
            Call CreateAvailableOffline
            Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
            Call FormatAvailable
    
            Set jgxOutbounds.ADORecordset = Nothing
            Call CreateOutboundsOffline
            Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
            Call FormatOutbounds
        
        End If
        
        Call CheckOutDoc
        
    End If
    
    If jgxAvailableStock.RowCount > 0 Or jgxOutbounds.RowCount > 0 Then
        cmdApply.Enabled = True
    End If
    
End Sub

Private Sub txtMRN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtMRN.Text <> strLastMRN Then
    ElseIf Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    ElseIf Not (KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
        KeyAscii = 0
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub ResetProduct()
        
    lblProdDesc.Caption = ""
    lblCtryExportCode.Caption = ""
    lblCtryExportDesc.Caption = ""
    lblCtryOriginCode.Caption = ""
    lblCtryOriginDesc.Caption = ""
    lblTARICCode.Caption = ""
        
End Sub

Private Sub CreateLinkedTables()
    Dim strHistoryPath As String
    Dim strHistoryYear As String
    
    Dim lngDBCtr As Long
    
    'Inbounds
    'InboundDocs
    
    strHistoryPath = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_history??.mdb")
    ReDim Preserve m_astrHistoryDBs(0)
    Do Until Len(Trim(strHistoryPath)) = 0
        ReDim Preserve m_astrHistoryDBs(UBound(m_astrHistoryDBs) + 1)
        m_astrHistoryDBs(UBound(m_astrHistoryDBs)) = strHistoryPath
        strHistoryPath = Dir()
    Loop
    
    
    For lngDBCtr = 1 To UBound(m_astrHistoryDBs)
        strHistoryYear = Replace(m_astrHistoryDBs(lngDBCtr), "mdb_history", vbNullString)
        strHistoryYear = Replace(strHistoryYear, ".mdb", vbNullString)
        
        strHistoryPath = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & m_astrHistoryDBs(lngDBCtr)
        
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "InboundHistory" & strHistoryYear & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_HISTORY, "Inbounds", , GetHistoryDBYear(m_astrHistoryDBs(lngDBCtr))
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "InboundDocHistory" & strHistoryYear & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_HISTORY, "InboundDocs", , GetHistoryDBYear(m_astrHistoryDBs(lngDBCtr))
        'AddLinkedTableEx "InboundHistory" & strHistoryYear & "_" & Format(m_lngUserID, "00"), NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "Inbounds", strHistoryPath, G_Main_Password
        'AddLinkedTableEx "InboundDocHistory" & strHistoryYear & "_" & Format(m_lngUserID, "00"), NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "InboundDocs", strHistoryPath, G_Main_Password
    Next lngDBCtr
    
End Sub

Private Sub DeleteLinkedTempTables()
    Dim lngDBCtr As Long
    Dim strHistoryDBYear As String
    
    
    On Error Resume Next
    ' Delete linked tables
    
    ' Delete temp tables
    ExecuteNonQuery m_conSADBEL, "DROP TABLE InboundHistory" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE InboundDocHistory" & "_" & Format(m_lngUserID, "00")
    On Error GoTo 0
    
End Sub

Private Sub CreateInboundHistory()
    Dim strHistoryDBYear As String
    Dim strDBFile As String
    Dim lngDBCtr As Long
    Dim strSQL As String
    
    
    On Error Resume Next
    ExecuteNonQuery m_conSADBEL, "DROP TABLE InboundHistory" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE InboundDocHistory" & "_" & Format(m_lngUserID, "00")
    On Error GoTo 0
    
    
    ' Create temporary tables in sadbel
    ExecuteNonQuery m_conSADBEL, "CREATE TABLE InboundHistory" & "_" & Format(m_lngUserID, "00") & " (In_ID LONG, Stock_ID LONG, In_Avl_Qty_Wgt DOUBLE, In_Orig_Packages_Type TEXT (50), InDoc_ID LONG, SourceDB TEXT (50))"
    ExecuteNonQuery m_conSADBEL, "CREATE TABLE InboundDocHistory" & "_" & Format(m_lngUserID, "00") & " (InDoc_ID LONG, InDoc_Type TEXT (50), InDoc_Num TEXT (50), InDoc_Date Date, SourceDB TEXT (50))"
    
    
    For lngDBCtr = 1 To UBound(m_astrHistoryDBs)
        strHistoryDBYear = Replace(m_astrHistoryDBs(lngDBCtr), "mdb_history", vbNullString)
        strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
        
        strDBFile = m_astrHistoryDBs(lngDBCtr)
        
            strSQL = vbNullString
            strSQL = strSQL & "INSERT INTO "
            strSQL = strSQL & "InboundHistory" & "_" & Format(m_lngUserID, "00") & " "
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "In_ID, "
            strSQL = strSQL & "Stock_ID, "
            strSQL = strSQL & "In_Avl_Qty_Wgt, "
            strSQL = strSQL & "In_Orig_Packages_Type, "
            strSQL = strSQL & "InDoc_ID, "
            strSQL = strSQL & "'" & Left(strDBFile, Len(strDBFile) - 4) & "' AS SourceDB "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "InboundHistory" & strHistoryDBYear & "_" & Format(m_lngUserID, "00")
        On Error Resume Next
        ExecuteNonQuery m_conSADBEL, strSQL
        On Error GoTo 0
            
            strSQL = vbNullString
            strSQL = strSQL & "INSERT INTO "
            strSQL = strSQL & "InboundDocHistory" & "_" & Format(m_lngUserID, "00") & " "
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "InDoc_ID, "
            strSQL = strSQL & "InDoc_Type, "
            strSQL = strSQL & "InDoc_Num, "
            strSQL = strSQL & "InDoc_Date, "
            strSQL = strSQL & "'" & Left(strDBFile, Len(strDBFile) - 4) & "' AS SourceDB "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "InboundDocHistory" & strHistoryDBYear & "_" & Format(m_lngUserID, "00")
        On Error Resume Next
        ExecuteNonQuery m_conSADBEL, strSQL
        On Error GoTo 0
    Next lngDBCtr
    
End Sub

Private Function SavedIfChanged() As Boolean

    m_rstAvailableOff.Filter = adFilterNone
    m_rstAvailableOff.Filter = "[Qty for Outbound] > 0 OR [Outbound Edits] <> 0 "
    m_rstOutboundsOff.Filter = "Out_ID = 0"
    If m_rstAvailableOff.RecordCount > 0 Or blnOutboundsChanged Or m_rstOutboundsOff.RecordCount > 0 Then
        If MsgBox("Save pending stock for outbound?", vbYesNo + vbQuestion, "Manual Outbound") = vbYes Then
            Call cmdApply_Click
            If cmdApply.Enabled Then
                SavedIfChanged = False
                Exit Function
            End If
        End If
    End If
    m_rstOutboundsOff.Filter = adFilterNone
    m_rstAvailableOff.Filter = adFilterNone
    ReDim alngDeleted(0)
    blnOutboundsChanged = False
    SavedIfChanged = True
    
End Function

Private Sub AddRecordToOutbound()
    
    Dim lngCol As Long
    Dim lngrow As Long
    Dim lngIndex As Long
    Dim lngSort As Long
    
    If Val(jgxAvailableStock.Value(jgxAvailableStock.Columns("Qty for Outbound").Index)) = 0 Then Exit Sub
    
'    If Val(jgxAvailableStock.Value(jgxAvailableStock.Columns("Qty for Outbound").Index)) > 0 Then
                
        lngCol = jgxAvailableStock.Col
        lngrow = jgxAvailableStock.Row
        
        m_rstAvailableOff.Filter = "In_ID = " & jgxAvailableStock.Value(jgxAvailableStock.Columns("In_ID").Index)
        
        m_rstOutboundsOff.AddNew
        m_rstOutboundsOff!Out_ID = 0
        m_rstOutboundsOff!In_ID = m_rstAvailableOff!In_ID
        m_rstOutboundsOff![Product No] = m_rstAvailableOff![Product No]
        m_rstOutboundsOff!Handling = m_rstAvailableOff!Handling
        m_rstOutboundsOff![Document No] = m_rstAvailableOff![Document No]
        m_rstOutboundsOff![Entrepot No] = m_rstAvailableOff![Entrepot No]
        m_rstOutboundsOff![Stock Card No] = m_rstAvailableOff![Stock Card No]
        m_rstOutboundsOff![Package Type] = m_rstAvailableOff![Package Type]
        m_rstOutboundsOff![Batch No] = m_rstAvailableOff![Batch No]
        m_rstOutboundsOff![Outbound Qty] = m_rstAvailableOff![Qty for Outbound]
        m_rstOutboundsOff![Job No] = m_rstAvailableOff![Job No]
        m_rstOutboundsOff.Update
        
        m_rstOutboundsOff.Filter = "In_ID = " & m_rstAvailableOff!In_ID
        
        If m_rstOutboundsOff.RecordCount > 0 Then
            m_rstOutboundsOff.MoveFirst
        End If
        
        Do While Not m_rstOutboundsOff.EOF
            m_rstOutboundsOff![Quantity/Weight] = m_rstAvailableOff![Quantity/Weight]
            m_rstOutboundsOff.MoveNext
        Loop
        
        m_rstOutboundsOff.Filter = adFilterNone
        
        m_rstAvailableOff![Outbound Edits] = Round(m_rstAvailableOff![Outbound Edits] + Val(m_rstAvailableOff![Qty for Outbound]), Choose(lngHandling + 1, 0, 2, 3))
        m_rstAvailableOff![Qty for Outbound] = ""
        
        m_rstAvailableOff.Filter = adFilterNone
        If strCorrectionMode <> "I" Then
            If chkShowZero.Value = vbUnchecked Then
                m_rstAvailableOff.Filter = "[Quantity/Weight] > 0 OR ([Qty for Outbound] <> 0 AND [Qty for Outbound] <> '') OR [Outbound Edits] <> 0 "   '"[Quantity/Weight] > 0 OR [Qty for Outbound] > 0"
            End If
        End If
        If jgxOutbounds.SortKeys.Count > 0 Then
            lngIndex = jgxOutbounds.SortKeys(1).ColIndex
            lngSort = jgxOutbounds.SortKeys(1).SortOrder
        End If
        Set jgxOutbounds.ADORecordset = Nothing
        Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
        Call FormatOutbounds
        If lngIndex > 0 Then
            jgxOutbounds.SortKeys.Add lngIndex, lngSort
            lngIndex = 0
        End If
        
        If jgxAvailableStock.SortKeys.Count > 0 Then
            lngIndex = jgxAvailableStock.SortKeys(1).ColIndex
            lngSort = jgxAvailableStock.SortKeys(1).SortOrder
        End If
        Set jgxAvailableStock.ADORecordset = Nothing
        Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
        Call FormatAvailable
        If lngIndex > 0 Then
            jgxAvailableStock.SortKeys.Add lngIndex, lngSort
            jgxAvailableStock.RefreshSort
            lngIndex = 0
        End If
        
        jgxAvailableStock.Row = IIf(lngrow > jgxAvailableStock.RowCount, jgxAvailableStock.RowCount, lngrow)
        jgxAvailableStock.Col = IIf(jgxAvailableStock.Row <> 0, lngCol, 0)
        
'    End If

End Sub

Private Sub txtProductNum_Change()
    
    Dim rstTemp As ADODB.Recordset
    
    If Not blnSystemChanged And Trim(UCase(txtProductNum.Text)) <> Trim(UCase(strLastProdNum)) Then
        
        Call ResetProduct
        
        If Not SavedIfChanged Then
            blnSystemChanged = True
            txtProductNum.Text = strLastProdNum
            blnSystemChanged = False
            Exit Sub
        End If
        
        If Len(Trim$(txtProductNum.Text)) > 0 Then
            ADORecordsetOpen "SELECT Prod_ID, Prod_Desc, Taric_Code, Prod_Ctry_Origin, Prod_Ctry_Export, Prod_Handling FROM Products WHERE Prod_Num = '" & Trim$(txtProductNum.Text) & "'", m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
            'rstTemp.Open "SELECT Prod_ID, Prod_Desc, Taric_Code, Prod_Ctry_Origin, Prod_Ctry_Export, Prod_Handling FROM Products WHERE Prod_Num = '" & Trim$(txtProductNum.Text) & "'", m_conSADBEL, adOpenKeyset, adLockOptimistic
            
            If Not (rstTemp.EOF And rstTemp.BOF) Then
                rstTemp.MoveFirst
                
                strProdID = ""
                
                lblProdDesc.Caption = rstTemp!Prod_Desc
                lblTARICCode.Caption = rstTemp!Taric_Code
                lblCtryOriginCode.Caption = rstTemp!Prod_Ctry_Origin
                lblCtryExportCode.Caption = rstTemp!Prod_Ctry_Export
                lblCtryOriginDesc.Caption = GetCountryDesc(rstTemp!Prod_Ctry_Origin, m_conSADBEL, strLanguage)
                lblCtryExportDesc.Caption = GetCountryDesc(rstTemp!Prod_Ctry_Export, m_conSADBEL, strLanguage)
                lngHandling = rstTemp!Prod_Handling
                
                Do While Not rstTemp.EOF
                
                    If Len(Trim$(strProdID)) = 0 Then
                        strProdID = rstTemp!Prod_ID
                    Else
                        strProdID = strProdID & "," & rstTemp!Prod_ID
                    End If
                    
                    rstTemp.MoveNext
                
                Loop
                
                If lngOutDocID <> 0 Then
                    If m_rstOutboundsOff.RecordCount > 0 Then
                        Call CreateOutboundsOffline
                    End If
                    Call PopulateOutbounds(strProdID)
                End If
                
                m_rstAvailableOff.Filter = adFilterNone
                If m_rstAvailableOff.RecordCount > 0 Then
                    Call CreateAvailableOffline
                End If
                'Checks if user wanted to display zero stocks
                Call PopulateAvailable
                
            Else
            
                If m_rstOutboundsOff.RecordCount > 0 Then
                    Call CreateOutboundsOffline
                    Set jgxOutbounds.ADORecordset = Nothing
                    Set jgxOutbounds.ADORecordset = m_rstOutboundsOff
                End If
                
                m_rstAvailableOff.Filter = adFilterNone
                If m_rstAvailableOff.RecordCount > 0 Then
                    Call CreateAvailableOffline
                    Set jgxAvailableStock.ADORecordset = Nothing
                    Set jgxAvailableStock.ADORecordset = m_rstAvailableOff
                End If
                
            End If
            
            ADORecordsetClose rstTemp
                
        ElseIf lngOutDocID <> 0 Then
                    
            If m_rstOutboundsOff.RecordCount > 0 Then
                Call CreateOutboundsOffline
            End If
            Call PopulateOutbounds
            
            m_rstAvailableOff.Filter = adFilterNone
            If m_rstAvailableOff.RecordCount > 0 Then
                Call CreateAvailableOffline
            End If
            'Checks if user wanted to display zero stocks
            Call PopulateAvailable
            
        End If
        
        strLastProdNum = txtProductNum.Text
        
    End If
    
End Sub

Private Function IsStockClosed() As Boolean
    Dim rstStock As ADODB.Recordset
    Dim strSQL As String
    Dim strDate As String
    
    If FraOutbounds.Enabled = True And jgxOutbounds.ADORecordset.RecordCount > 0 Then
        jgxOutbounds.ADORecordset.MoveFirst
        Do While Not jgxOutbounds.ADORecordset.EOF
            strSQL = "SELECT TOP 1 OUTDOC_DATE FROM OUTBOUNDDOCS INNER JOIN (OUTBOUNDS INNER JOIN (INBOUNDS INNER JOIN (STOCKCARDS INNER JOIN (PRODUCTS INNER JOIN ENTREPOTS ON PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID) ON STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID) ON INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID) ON OUTBOUNDS.IN_ID = INBOUNDS.IN_ID) ON OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID WHERE UCASE(RIGHT(OUT_CODE,11))= '<<CLOSURE>>' AND DATEVALUE(OUTDOC_DATE) >= DATEVALUE('" & dtpDate.Value & "') AND ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM ='" & jgxOutbounds.ADORecordset.Fields("Entrepot No").Value & "'"
            
            ADORecordsetOpen strSQL, m_conSADBEL, rstStock, adOpenKeyset, adLockOptimistic
            'rstStock.Open strSQL, m_conSADBEL, adOpenKeyset, adLockReadOnly
            
            If Not (rstStock.BOF And rstStock.EOF) Then
                IsStockClosed = True
                Exit Do
            Else
                IsStockClosed = False
            End If
            
            ADORecordsetClose rstStock
            
            jgxOutbounds.ADORecordset.MoveNext
        Loop
        
        ADORecordsetClose rstStock
    End If
    
End Function




