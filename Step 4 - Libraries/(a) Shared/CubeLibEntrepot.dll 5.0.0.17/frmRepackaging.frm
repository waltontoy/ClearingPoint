VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmRepackaging 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repackaging"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmRepackaging.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Tag             =   "178"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Tag             =   "179"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Frame fraQtyRepack 
      Caption         =   "Quantity to Repackage"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   4740
      Width           =   9975
      Begin VB.TextBox txtGrossWt 
         Height          =   315
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   8
         Top             =   285
         Width           =   1815
      End
      Begin VB.TextBox txtNetWt 
         Height          =   315
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   7
         Top             =   285
         Width           =   1815
      End
      Begin VB.TextBox txtNoOfItems 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   6
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblGrossWeight 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Weight:"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label lblNetWeight 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight:"
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblNumItems 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Items:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   315
         Width           =   1575
      End
   End
   Begin VB.Frame fraNewStock 
      Caption         =   "New Stock"
      Height          =   3135
      Left            =   135
      TabIndex        =   21
      Top             =   5520
      Width           =   9975
      Begin GridEX16.GridEX jgxRepackage 
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4260
         Enabled         =   0   'False
         MethodHoldFields=   -1  'True
         Options         =   -1
         AllowColumnDrag =   0   'False
         RecordsetType   =   1
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         ColumnCount     =   6
         CardCaption1    =   -1  'True
         ColCaption1     =   "Package Type"
         ColCaption2     =   "Num of Items"
         ColCaption3     =   "Net Weight"
         ColCaption4     =   "Gross Weight"
         ColCaption5     =   "Job Num"
         ColWidth5       =   1680
         ColCaption6     =   "Batch Num"
         ColWidth6       =   1680
         ItemCount       =   0
         DataMode        =   1
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
      Begin VB.Label lblNotice 
         Caption         =   "Type in stock information on the top line and then press ENTER to add the stock to the list."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame fraAvailableStock 
      Caption         =   "Available Stock"
      Height          =   2775
      Left            =   105
      TabIndex        =   20
      Tag             =   "2181"
      Top             =   1920
      Width           =   9975
      Begin GridEX16.GridEX jgxAvailableStocks 
         Height          =   2445
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4313
         MethodHoldFields=   -1  'True
         Options         =   -1
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnCount     =   8
         CardCaption1    =   -1  'True
         ColCaption1     =   "Doc Number"
         ColWidth1       =   1200
         ColCaption2     =   "Doc Date"
         ColWidth2       =   1200
         ColCaption3     =   "Package Type"
         ColWidth3       =   1200
         ColCaption4     =   "Num of Items"
         ColWidth4       =   1200
         ColCaption5     =   "Net Weight"
         ColWidth5       =   1200
         ColCaption6     =   "Gross Weight"
         ColWidth6       =   1200
         ColCaption7     =   "Job Num"
         ColWidth7       =   1215
         ColCaption8     =   "Batch Num"
         ColWidth8       =   1215
         ItemCount       =   0
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Frame fraProduct 
      Caption         =   "Product to Repackage"
      Height          =   1365
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   9975
      Begin VB.CommandButton cmdProductNum 
         Caption         =   "..."
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblCtryOriginDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7560
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblProductNum 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblSC 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Card Number:"
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Tag             =   "2220"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblStockCardNum 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7080
         TabIndex        =   27
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblProduct 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Tag             =   "2274"
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label lblHandling 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8340
         TabIndex        =   19
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblCtryOriginCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   960
         Width           =   8175
      End
      Begin VB.Label lblTaricCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   3090
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "342"
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblTaric 
         BackStyle       =   0  'Transparent
         Caption         =   "Taric Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "2275"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblCtryOrigin 
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Origin:"
         Height          =   255
         Left            =   5520
         TabIndex        =   13
         Tag             =   "2195"
         Top             =   315
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEntrepotNum 
      Caption         =   "..."
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtEntrepotNum 
      Height          =   315
      Left            =   1800
      MaxLength       =   19
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblEntrepot 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrepot Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "2273"
      Top             =   150
      Width           =   1575
   End
End
Attribute VB_Name = "frmRepackaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_conSADBEL As ADODB.Connection
Private m_conTaric As ADODB.Connection

Private m_rstStocksOff As ADODB.Recordset
Private m_rstRepackOff As ADODB.Recordset

Private m_lngUserID As Long

Private strLanguage As String
Private intTaricProperties As Integer
Private strPrevious_Entrepot_Num As String

Private blnFormLoaded As Boolean
Private blnUserChanged As Boolean

Private blnGridIsNothing As Boolean
Private blnEntered As Boolean
Private blnCancelMove As Boolean
Private blnSystemChanged As Boolean
Private lngPack_Flag As Long
Private arrRow() As Variant

Private strSQLPack As String
Private pckList As PCubeLibPick.CPicklist
Private gsdList As PCubeLibPick.CGridSeed

Private m_lngInDoc As Long
Private m_lngInID As Long
Private m_alngInID2() As Long

Private m_strDocDate As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEntrepotNum_Click()
    Dim clsEntrepot As cEntrepot
    Set clsEntrepot = New cEntrepot
    
    With clsEntrepot
        'Calls Entrepot picklist.
        .ShowEntrepot Me, m_conSADBEL, True, strLanguage, ResourceHandler, Me.txtEntrepotNum.Name, Val(txtEntrepotNum.Tag), True
        
        'Initializes Product Number when an Entrepot is selected.
        If .Cancelled = False Then
            'If different Entrepot is selected, reset controls value.
            If Not (strPrevious_Entrepot_Num = txtEntrepotNum.Text) Then
                strPrevious_Entrepot_Num = txtEntrepotNum.Text
                
                Call ResetFields
            End If
        End If
    End With
    
    Set clsEntrepot = Nothing
End Sub

Public Sub MyLoad(ByRef connSadbel As ADODB.Connection, ByRef connTaric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, _
                  ByVal UserID As Long)
                  
    ResourceHandler = MyResourceHandler
    strLanguage = Language
    
    m_lngUserID = UserID
    
    modGlobals.LoadResStrings Me, True
    
    Set m_conSADBEL = connSadbel
    Set m_conTaric = connTaric
    
    intTaricProperties = TaricProperties
    
    'Lock handling textboxes.
    Call EnableHandling(999)
    
    Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    Dim conHistory As ADODB.Connection
    Dim conSADBEL As ADODB.Connection
    Dim conData As ADODB.Connection
    
    ADOConnectDB conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    ADOConnectDB conData, g_objDataSourceProperties, DBInstanceType_DATABASE_DATA
    
    If Len(Trim(Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(Year(Now), 2) & ".mdb"))) = 0 Then
        CreateHistoryMdb conSADBEL, conData, Right(Year(Now), 2)
    End If
    
    ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(Year(Now), 2)
    
    ADODisconnectDB conSADBEL
    ADODisconnectDB conData
    
    
    If jgxRepackage.ADORecordset Is Nothing Then
        ADODisconnectDB conHistory
        
        Unload Me
        
        Exit Sub
    End If
    
    If Not CheckIfRowToAdd() Then
        ADODisconnectDB conHistory
        
        Exit Sub
    End If
    
    If ExceedAvailableItemWt(Choose(Val(lblHandling.Tag) + 1, txtNoOfItems.Text, txtGrossWt.Text, txtNetWt.Text)) Then
        ADODisconnectDB conHistory
        
        Exit Sub
    End If
    
    If Not ValidPackages() Then
        ADODisconnectDB conHistory
        
        Exit Sub
    End If
    
    If Not ValidHandlingVal() Then
        ADODisconnectDB conHistory
        
        Exit Sub
    End If
    
    If jgxRepackage.RowCount > 0 Then
        If ValidateWeight = True Then
            Call OverwriteInboundsTable(m_conSADBEL, "SADBEL")
            Call UpdateInboundDocsTable(m_conSADBEL, "SADBEL")
            Call Cancelling(m_conSADBEL, conHistory, "SADBEL")
            Call UpdateInboundsTable(m_conSADBEL, "SADBEL")
            
            Call OverwriteInboundsTable(conHistory, "HISTORY")
            Call UpdateInboundDocsTable(conHistory, "HISTORY")
            Call Cancelling(conHistory, m_conSADBEL, "HISTORY")
            Call UpdateInboundsTable(conHistory, "HISTORY")
        Else
            ADODisconnectDB conHistory
            
            Exit Sub
        End If
    End If
    
    ADODisconnectDB conHistory
    
    Unload Me
    
End Sub

'This simply checks if the value for the product's handling type is not zero. IAN

Private Function ValidHandlingVal() As Boolean
    
    Dim lngIndex As Long
    Dim lngCounter As Long
    
    Select Case Val(lblHandling.Tag)
    
        Case 0
            lngIndex = jgxRepackage.Columns("Num of Items").Index
        Case 1
            lngIndex = jgxRepackage.Columns("Gross Weight").Index
        Case 2
            lngIndex = jgxRepackage.Columns("Net Weight").Index
            
    End Select
    
    blnSystemChanged = True
    jgxRepackage.MoveFirst
    
    For lngCounter = 1 To jgxRepackage.RowCount
    
        If IsNull(jgxRepackage.Value(lngIndex)) Or jgxRepackage.Value(lngIndex) = 0 Then
            MsgBox Translate(2296), vbInformation, Translate(2297)
            ValidHandlingVal = False
            jgxRepackage.SetFocus
            jgxRepackage.Col = lngIndex
            Exit Function
        End If
        
        jgxRepackage.MoveNext
        
    Next
    
    blnSystemChanged = False
    ValidHandlingVal = True
    
End Function

'This checks if the packaging type entered by the user are valid, hence, it is in the
'database. IAN

Private Function ValidPackages() As Boolean
    
    Dim rstPackagesOff As ADODB.Recordset
    Dim lngCounter As Long
    
    ADORecordsetOpen strSQLPack, m_conSADBEL, rstPackagesOff, adOpenKeyset, adLockOptimistic
    'Set rstPackagesOff = New ADODB.Recordset
    'rstPackagesOff.CursorLocation = adUseClient
    'rstPackagesOff.Open strSQLPack, m_conSADBEL, adOpenKeyset, adLockOptimistic
    'rstPackagesOff.ActiveConnection = Nothing
   
    If (rstPackagesOff.EOF And rstPackagesOff.BOF) Then
        ValidPackages = False
        
        ADORecordsetClose rstPackagesOff
        
        Exit Function
    End If
    
    If jgxRepackage.RowCount > 0 Then
        
        blnSystemChanged = True
        
        jgxRepackage.MoveFirst
        
        For lngCounter = 1 To jgxRepackage.RowCount
            rstPackagesOff.MoveFirst
            rstPackagesOff.Find "CODE = '" & IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)), "", jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)) & "'", , adSearchForward
            
            If rstPackagesOff.EOF Then
                MsgBox Translate(2298), vbInformation, Translate(2297)
                jgxRepackage.SetFocus
                jgxRepackage.Col = jgxRepackage.Columns("Package Type").Index
                ValidPackages = False
                
                ADORecordsetClose rstPackagesOff
                
                Exit Function
            End If
            
            jgxRepackage.MoveNext
        Next
            
        blnSystemChanged = False
        
    End If
    
    ValidPackages = True
    
    ADORecordsetClose rstPackagesOff
    
End Function

Private Sub cmdProductNum_Click()
    Dim clsStockProd As cStockProd
    Set clsStockProd = New cStockProd
                
    With clsStockProd
        'Used to automatically select previously selected item.
        If Len(lblStockCardNum.Tag) <> 0 Then .Stock_ID = lblStockCardNum.Tag
        If Len(lblProductNum.Tag) <> 0 Then .Product_ID = lblProductNum.Tag
        If Len(lblStockCardNum.Caption) > 0 Then .StockCardNo = lblStockCardNum.Caption
        .Entrepot_Num = txtEntrepotNum.Text
        
        'Calls Stock/Prod picklist.
        .ShowPicklist Me, m_conSADBEL, m_conTaric, strLanguage, intTaricProperties, ResourceHandler, True
        
        'Passes values to controls.
        If Trim(.Stock_ID) <> 0 And .Cancel = False Then
            lblProductNum.Caption = .ProductNo
            lblProductNum.Tag = .Product_ID
            lblTARICCode.Caption = .TaricCode
            lblDescription.Caption = .ProductDesc
            lblStockCardNum.Tag = .Stock_ID
            lblStockCardNum.Caption = .StockCardNo
            lblCtryOriginCode.Caption = .CtryOrigin
            lblCtryOriginDesc.Caption = GetCountryDesc(.CtryOrigin, m_conSADBEL, strLanguage)
            lblHandling.Tag = .ProductHandling
            
            'Enters appropriate Handling description.
            Select Case .ProductHandling
                Case 0
                    lblHandling.Caption = "Number of Items"
                    EnableHandling 0
                Case 1
                    lblHandling.Caption = "Gross Weight"
                    EnableHandling 1
                Case 2
                    lblHandling.Caption = "Net Weight"
                    EnableHandling 2
            End Select
                    
            blnGridIsNothing = True
            Set jgxRepackage.ADORecordset = Nothing 'IAN
            blnGridIsNothing = False
            Call CreateRepackFields
            Call PopulateAvailableStock(Val(lblStockCardNum.Tag)) 'IAN

        End If
    End With
    
    Set clsStockProd = Nothing
End Sub

Private Sub EnableHandling(ByVal intHandle As Integer)
    'Used for enabling and coloring handling input boxes.
    Select Case intHandle
        Case 0
            txtNoOfItems.Enabled = True
            txtGrossWt.Enabled = False
            txtNetWt.Enabled = False
            
            txtNoOfItems.BackColor = vbWhite
            txtGrossWt.BackColor = vbButtonFace
            txtNetWt.BackColor = vbButtonFace
        Case 1
            txtNoOfItems.Enabled = False
            txtGrossWt.Enabled = True
            txtNetWt.Enabled = False
            
            txtNoOfItems.BackColor = vbButtonFace
            txtGrossWt.BackColor = vbWhite
            txtNetWt.BackColor = vbButtonFace
        Case 2
            txtNoOfItems.Enabled = False
            txtGrossWt.Enabled = False
            txtNetWt.Enabled = True
    
            txtNoOfItems.BackColor = vbButtonFace
            txtGrossWt.BackColor = vbButtonFace
            txtNetWt.BackColor = vbWhite
        Case 999
            'Used by ResetFields
            txtNoOfItems.Enabled = False
            txtGrossWt.Enabled = False
            txtNetWt.Enabled = False
    
            txtNoOfItems.BackColor = vbButtonFace
            txtGrossWt.BackColor = vbButtonFace
            txtNetWt.BackColor = vbButtonFace
    End Select
End Sub

Private Sub ResetFields()
    'Empties controls used when new Entrepot is selected.
    lblProductNum.Caption = Empty
    lblProductNum.Tag = Empty
    lblTARICCode.Caption = Empty
    lblDescription.Caption = Empty
    lblStockCardNum.Caption = Empty
    lblStockCardNum.Tag = Empty
    lblCtryOriginCode.Caption = Empty
    lblCtryOriginDesc.Caption = Empty
    lblHandling.Caption = Empty
    Call EnableHandling(999)
    
    Set jgxAvailableStocks.ADORecordset = Nothing 'IAN
    Call CreateAvailableFields
    blnGridIsNothing = True
    Set jgxRepackage.ADORecordset = Nothing
    blnGridIsNothing = False
    Call CreateRepackFields
    jgxRepackage.Enabled = False
    
End Sub

'Always show the columns in the available grid for better user interface. IAN

Private Sub CreateAvailableFields()
    
    jgxAvailableStocks.DefaultColumnWidth = 1200
    jgxAvailableStocks.Columns.Clear
    jgxAvailableStocks.Columns.Add "Doc Number"
    jgxAvailableStocks.Columns.Add "Doc Date"
    jgxAvailableStocks.Columns.Add "Package Type"
    jgxAvailableStocks.Columns.Add "Num of Items"
    jgxAvailableStocks.Columns.Add "Net Weight"
    jgxAvailableStocks.Columns.Add "Gross Weight"
    jgxAvailableStocks.Columns.Add "Job Num"
'    jgxAvailableStocks.Columns("Job Num").Width = 1305
    
    jgxAvailableStocks.Columns.Add "Batch Num"
'    jgxAvailableStocks.Columns("Batch Num").Width = 1305
    
End Sub

'Alaways show the columns in the repackage grid for better user interface. IAN

Private Sub CreateRepackFields()
    
    jgxRepackage.DefaultColumnWidth = 1500
    jgxRepackage.Columns.Clear
    jgxRepackage.Columns.Add "Package Type"
    jgxRepackage.Columns.Add "Num of Items"
    jgxRepackage.Columns.Add "Net Weight"
    jgxRepackage.Columns.Add "Gross Weight"
    jgxRepackage.Columns.Add "Job Num"
    jgxRepackage.Columns.Add "Batch Num"
    
End Sub

Private Sub Form_Load()
            
    strSQLPack = "SELECT [PICKLIST MAINTENANCE " & strLanguage & "].CODE AS [Key Code]," & _
                      "[PICKLIST MAINTENANCE " & strLanguage & "].CODE AS [Code]," & _
                      "[PICKLIST MAINTENANCE " & strLanguage & "].[DESCRIPTION " & strLanguage & "] AS [Key Description] " & _
                      "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE " & strLanguage & "] " & _
                      "WHERE " & _
                      "([PICKLIST DEFINITION].[BOX CODE]= 'E3') AND " & _
                      "([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND " & _
                      "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE " & strLanguage & "].[internal code]) "
                      
    jgxAvailableStocks.DefaultColumnWidth = 1200
    
    blnFormLoaded = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Erase arrRow
    
    ADORecordsetClose m_rstStocksOff
    ADORecordsetClose m_rstRepackOff
    
    ADODisconnectDB m_conSADBEL
    ADODisconnectDB m_conTaric
End Sub

Private Sub jgxAvailableStocks_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    jgxAvailableStocks.SortKeys.Clear
    jgxAvailableStocks.SortKeys.Add Column.Index, jgexSortAscending
    jgxAvailableStocks.RefreshSort
End Sub


Private Sub jgxRepackage_BeforeUpdate(ByVal Cancel As GridEX16.JSRetBoolean)

    Dim lngCtr As Long
    Dim blnTemp As Boolean
    
    If jgxRepackage.ADORecordset Is Nothing Or blnGridIsNothing Then
        Exit Sub
    End If
    
    If jgxRepackage.Row = -1 Then
                
        If blnEntered Then
            
            If CheckInvalidRecord() Then
                Cancel = True
            Else
                blnSystemChanged = True
            End If
            
            blnEntered = False
            
            For lngCtr = 1 To jgxRepackage.Columns.Count
                arrRow(lngCtr) = jgxRepackage.Columns(lngCtr).DefaultValue
            Next
        ElseIf CheckIfRowToAdd() Then
                        
            For lngCtr = 1 To jgxRepackage.Columns.Count
                arrRow(lngCtr) = jgxRepackage.Columns(lngCtr).DefaultValue
            Next
        
            blnSystemChanged = True
            blnTemp = blnCancelMove
            blnCancelMove = False
            jgxRepackage.Delete
            blnCancelMove = blnTemp
            blnSystemChanged = False
            Exit Sub
        
        Else
                        
            For lngCtr = 1 To jgxRepackage.Columns.Count
                arrRow(lngCtr) = jgxRepackage.Value(lngCtr)
            Next
                        
            blnSystemChanged = True
            blnTemp = blnCancelMove
            blnCancelMove = False
            jgxRepackage.Delete
            blnCancelMove = blnTemp
            blnSystemChanged = False
            Exit Sub
            
        End If
        
    ElseIf jgxRepackage.Row > 0 Then
        
        If blnCancelMove Then
            jgxRepackage.EditMode = jgexEditModeOn
            jgxRepackage.SelStart = 0
            Cancel = True
        ElseIf CheckInvalidRecord() Then
            Cancel = True
        End If
                
    End If
    
End Sub

Private Sub jgxRepackage_Change()
    
    If jgxRepackage.Col = jgxRepackage.Columns("Package Type").Index Then
        jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index) = 0
        lngPack_Flag = 0
    End If

End Sub

Private Sub jgxRepackage_KeyDown(KeyCode As Integer, Shift As Integer)

    blnSystemChanged = False
    blnCancelMove = False
    
    If KeyCode = vbKeyReturn Then
        blnEntered = True
    ElseIf KeyCode = vbKeyF2 And jgxRepackage.Col = jgxRepackage.Columns("Package Type").Index Then
            
        Dim pckPackaging As PCubeLibPick.CPicklist
        Dim gsdPackaging As PCubeLibPick.CGridSeed
        
        Set pckPackaging = New CPicklist
        Set gsdPackaging = pckPackaging.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
                    
        With pckPackaging
            If Not IsNull(jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)) Then
                If jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index) <> "" Then
                    .Search True, "Key Code", jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)
                End If
            End If
            ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
            .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdPackaging, , , True, cpiKeyF2
            
            If Not .SelectedRecord Is Nothing Then
                jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index) = 1
                lngPack_Flag = 1
            Else
                KeyCode = 0
            End If
            
        End With
                    
        Set gsdPackaging = Nothing
        Set pckPackaging = Nothing
            
    End If

End Sub

Private Sub jgxRepackage_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    blnSystemChanged = False
    blnCancelMove = False

End Sub

Private Sub jgxRepackage_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)

    If Not blnSystemChanged And Not jgxRepackage.ADORecordset Is Nothing And LastRow <> 0 Then
        
        Dim lngrow As Long
        
        If LastCol = jgxRepackage.Columns("Package Type").Index And lngPack_Flag = 0 Then
            
            Dim strPack As String
            
            blnSystemChanged = True
            lngrow = jgxRepackage.Row
            jgxRepackage.Row = LastRow
            strPack = IIf(IsNull(jgxRepackage.Value(LastCol)), "", jgxRepackage.Value(LastCol))
            
            Set pckList = New CPicklist
            Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
                        
            With pckList
        
                If Trim(strPack) <> "" Then
                    .Search True, "Key Code", strPack
                End If
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyEnter
                                
                If Not .SelectedRecord Is Nothing Then
                    jgxRepackage.Value(LastCol) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index) = 1
                    lngPack_Flag = 1
                    jgxRepackage.Row = lngrow
                Else
                    jgxRepackage.Col = LastCol
                    jgxRepackage.EditMode = jgexEditModeOn
                    jgxRepackage.SelStart = 0
                    blnCancelMove = True
                End If
        
            End With
            
            Set pckList = Nothing
            Set gsdList = Nothing
                    
        End If
            
    ElseIf blnSystemChanged And blnCancelMove Then
        blnCancelMove = False
        jgxRepackage.Row = LastRow
        jgxRepackage.Col = LastCol
        jgxRepackage.EditMode = jgexEditModeOn
        jgxRepackage.SelStart = 0
    End If
    
    If LastRow <> jgxRepackage.Row Then
        
        If jgxRepackage.Row = -1 And Not jgxRepackage.ADORecordset Is Nothing And LastRow <> 0 Then
            Dim lngCounter As Long
            For lngCounter = LBound(arrRow) To UBound(arrRow)
                jgxRepackage.Value(lngCounter) = arrRow(lngCounter)
            Next
        End If
        
        lngPack_Flag = jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index)
        
    End If

End Sub

Private Sub txtEntrepotNum_GotFocus()
    'Stores current Entrepot Num in memory.
    strPrevious_Entrepot_Num = txtEntrepotNum.Text
End Sub

Private Sub txtEntrepotNum_LostFocus()
    'Performs a Entrepot Num memory refresh and ClearFields if new Entrepot Num detected.
    If Not (strPrevious_Entrepot_Num = txtEntrepotNum.Text) Then
        strPrevious_Entrepot_Num = txtEntrepotNum.Text
        Call ResetFields
    End If
End Sub

Private Sub CreateOfflineRecordset()
    
    'create offline recordset for repackaging grid. IAN
    ADORecordsetClose m_rstRepackOff
    
    
    Set m_rstRepackOff = New ADODB.Recordset
    
    m_rstRepackOff.CursorLocation = adUseClient
    
    m_rstRepackOff.Fields.Append "Package Type", adVarWChar, 50, adFldIsNullable
    m_rstRepackOff.Fields.Append "Num of Items", adVarWChar, 12, adFldIsNullable
    m_rstRepackOff.Fields.Append "Net Weight", adVarWChar, 12, adFldIsNullable
    m_rstRepackOff.Fields.Append "Gross Weight", adVarWChar, 12, adFldIsNullable
    m_rstRepackOff.Fields.Append "Job Num", adVarWChar, 50, adFldIsNullable
    m_rstRepackOff.Fields.Append "Batch Num", adVarWChar, 50, adFldIsNullable
    m_rstRepackOff.Fields.Append "Pack_Flag", adInteger
    
    m_rstRepackOff.Open
    
End Sub

'This procedure computes for other unit measurement once the value of the product
'handling changed. IAN

Private Sub FillQuantity(Optional ByVal blnAll As Boolean = False)

    If blnAll Then
        txtNoOfItems.Text = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Num of Items").Index)
        txtNetWt.Text = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Net Weight").Index)
        txtGrossWt.Text = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Gross Weight").Index)
    Else
        
        Dim dblTemp As Double
        
        Select Case UCase(lblHandling.Caption)
        
            Case "NUMBER OF ITEMS"
                
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Net Weight").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Num of Items").Index)) * Val(txtNoOfItems.Text)
                txtNetWt.Text = Replace(CStr(Round(dblTemp, 3) + IIf(Round(dblTemp, 3) < dblTemp, 0.001, 0)), ",", ".")
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Gross Weight").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Num of Items").Index)) * Val(txtNoOfItems.Text)
                txtGrossWt.Text = Replace(CStr(Round(dblTemp, 2) + IIf(Round(dblTemp, 2) < dblTemp, 0.01, 0)), ",", ".")
                
            Case "GROSS WEIGHT"
                
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Num of Items").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Gross Weight").Index)) * Val(txtGrossWt.Text)
                txtNoOfItems.Text = Round(dblTemp, 0) + IIf(Round(dblTemp, 0) < dblTemp, 1, 0)
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Net Weight").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Gross Weight").Index)) * Val(txtGrossWt.Text)
                txtNetWt.Text = Replace(CStr(Round(dblTemp, 3) + IIf(Round(dblTemp, 3) < dblTemp, 0.001, 0)), ",", ".")
            
            Case "NET WEIGHT"
                                
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Num of Items").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Net Weight").Index)) * Val(txtNetWt.Text)
                txtNoOfItems.Text = Round(dblTemp, 0) + IIf(Round(dblTemp, 0) < dblTemp, 1, 0)
                dblTemp = Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Gross Weight").Index)) / Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Net Weight").Index)) * Val(txtNetWt.Text)
                txtGrossWt.Text = Replace(CStr(Round(dblTemp, 2) + IIf(Round(dblTemp, 2) < dblTemp, 0.01, 0)), ",", ".")
                
        End Select
        
    End If
        
End Sub

'This procedure will determine and prompt the user if the quantity being repackaged
'exceeds the available units. IAN

Private Function ExceedAvailableItemWt(ByVal UnitVal As String) As Boolean

    If IIf(IsNull(UnitVal), 0, Val(UnitVal)) > IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns( _
        IIf(Val(lblHandling.Tag) = 0, "Num of Items", lblHandling.Caption)).Index)), 0, Val(jgxAvailableStocks.Value( _
        jgxAvailableStocks.Columns(IIf(Val(lblHandling.Tag) = 0, "Num of Items", lblHandling.Caption)).Index))) Then
        
        MsgBox Translate(2299), vbInformation + vbOKOnly, Translate(2297)
        ExceedAvailableItemWt = True
        
    End If
    
End Function

'Formats the repackage grid's text lengths, positioning, and default values. IAN

Private Sub FormatRepackGrid()
    jgxRepackage.Columns("Package Type").ButtonStyle = jgexButtonEllipsis
    jgxRepackage.Columns("Num of Items").TextAlignment = jgexAlignRight
    jgxRepackage.Columns("Net Weight").TextAlignment = jgexAlignRight
    jgxRepackage.Columns("Gross Weight").MaxLength = 12
    jgxRepackage.Columns("Num of Items").MaxLength = 6
    jgxRepackage.Columns("Net Weight").MaxLength = 12
    jgxRepackage.Columns("Package Type").MaxLength = 2
    jgxRepackage.Columns("Job Num").MaxLength = 50
    jgxRepackage.Columns("Batch Num").MaxLength = 50
    
    jgxRepackage.Columns("Gross Weight").TextAlignment = jgexAlignRight
    jgxRepackage.Columns("Job Num").DefaultValue = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Job Num").Index)
    jgxRepackage.Columns("Job Num").Width = 1680
    
    jgxRepackage.Columns("Batch Num").DefaultValue = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Batch Num").Index)
    jgxRepackage.Columns("Batch Num").Width = 1680
    
    jgxRepackage.Columns("Pack_Flag").DefaultValue = 0
    jgxRepackage.Columns("Num of Items").DefaultValue = "0"
    jgxRepackage.Columns("Gross Weight").DefaultValue = "0"
    jgxRepackage.Columns("Net Weight").DefaultValue = "0"
    
    jgxRepackage.Columns("Pack_Flag").Visible = False
End Sub

'This procedure will fill the available stock grid with values from the database and compute
'for other quantities aside from the one defined in the available quantity field in the
'database. IAN

Private Sub PopulateAvailableStock(ByVal Stock_ID As Long)
    
    Dim rstTempStocks As ADODB.Recordset
    Dim rstDIA As ADODB.Recordset
    Dim strSQL As String
    Dim lngCounter As Long
    
    'SQL to get DIA records
    strSQL = "SELECT Inbounds!In_Batch_Num AS Batch_Num " & _
        "FROM Inbounds INNER JOIN (StockCards INNER JOIN Products ON " & _
        "Products.Prod_ID = StockCards.Prod_ID) ON " & _
        "StockCards.Stock_ID = Inbounds.Stock_ID " & _
        "WHERE Inbounds!Stock_ID = " & Stock_ID & _
        " AND Inbounds!In_Job_Num='DIA' " & _
        "AND CHOOSE(Products!Prod_Handling + 1, Inbounds!In_Orig_Packages_Qty < 0, Inbounds!In_Orig_Gross_Weight < 0, Inbounds!In_Orig_Net_Weight < 0)"
    ADORecordsetOpen strSQL, m_conSADBEL, rstDIA, adOpenKeyset, adLockOptimistic
    'rstDIA.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic

    ADORecordsetClose m_rstStocksOff
    
    
    
    strSQL = "SELECT InBoundDocs!InDoc_Num AS [Doc Number], " & _
            "MIN(InBoundDocs!InDoc_Date) AS [Doc Date], " & _
            "Inbounds!In_Orig_Packages_Type AS [Package Type], " & _
            CreateSumSQL(Val(lblHandling.Tag) + 1) & _
            "Inbounds!In_Job_Num AS [Job Num], " & _
            "Inbounds!In_Batch_Num AS [Batch Num], " & _
            "InBoundDocs!InDoc_Office AS InDoc_Office, " & _
            "InBoundDocs!InDoc_SeqNum AS InDoc_SeqNum, " & _
            "InBoundDocs!InDoc_Cert_Type AS InDoc_Cert_Type, " & _
            "InBoundDocs!InDoc_Cert_Num AS InDoc_Cert_Num " & _
            "FROM Inbounds INNER JOIN InBoundDocs ON Inbounds.Indoc_ID = InBoundDocs.InDoc_ID " & _
            "WHERE Inbounds!Stock_ID = " & lblStockCardNum.Tag & _
            " AND Inbounds!In_Avl_Qty_Wgt > 0 " & _
            " AND IIF(ISNULL(Inbounds!In_Code),'', Inbounds!In_Code NOT LIKE '%<<TEST>>')" & _
            " AND IIF(ISNULL(Inbounds!In_Code),'', Inbounds!In_Code NOT LIKE '%<<CLOSURE>>')" & _
            "GROUP BY InBoundDocs!InDoc_Num, Inbounds!In_Orig_Packages_Type, Inbounds!In_Job_Num, Inbounds!In_Batch_Num, InBoundDocs!InDoc_Office, InBoundDocs!InDoc_SeqNum, InBoundDocs!InDoc_Cert_Type, InBoundDocs!InDoc_Cert_Num " & _
            "HAVING SUM(Inbounds!In_Avl_Qty_Wgt) > 0 " & _
            "ORDER BY InBoundDocs!InDoc_Num, Inbounds!In_Orig_Packages_Type, Inbounds!In_Job_Num, Inbounds!In_Batch_Num, InBoundDocs!InDoc_Office, InBoundDocs!InDoc_SeqNum, InBoundDocs!InDoc_Cert_Type, InBoundDocs!InDoc_Cert_Num"
            
    ADORecordsetOpen strSQL, m_conSADBEL, rstTempStocks, adOpenKeyset, adLockOptimistic
    'rstTempStocks.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    Set m_rstStocksOff = New ADODB.Recordset
    m_rstStocksOff.CursorLocation = adUseClient
    For lngCounter = 0 To rstTempStocks.Fields.Count - 1
        m_rstStocksOff.Fields.Append rstTempStocks.Fields(lngCounter).Name, rstTempStocks.Fields(lngCounter).Type, rstTempStocks.Fields(lngCounter).DefinedSize, rstTempStocks.Fields(lngCounter).Attributes
    Next
    m_rstStocksOff.Open
    
    If Not (rstTempStocks.EOF And rstTempStocks.BOF) Then
        rstTempStocks.MoveFirst
        
        Do While Not rstTempStocks.EOF
            rstDIA.Filter = adFilterNone
            rstDIA.Filter = "Batch_Num = '" & rstTempStocks![Doc Number] & "'"
                                    
            'IAN 05-21-2005
            'If record count is greater than zero then current record on rstTemp has been
            'cancelled by DIA, hence, don't include in the grid.
            If rstDIA.RecordCount = 0 Then
            
                m_rstStocksOff.AddNew
                
                For lngCounter = 0 To rstTempStocks.Fields.Count - 1
                    
                    Select Case UCase(rstTempStocks.Fields(lngCounter).Name)
                                  
                        Case "NUM OF ITEMS"
                            m_rstStocksOff.Fields(lngCounter).Value = CStr(Round(rstTempStocks.Fields(lngCounter).Value, 0) + IIf(Round(rstTempStocks.Fields(lngCounter).Value, 0) < CDbl(rstTempStocks.Fields(lngCounter).Value), 1, 0))
                        Case "NET WEIGHT"
                            m_rstStocksOff.Fields(lngCounter).Value = Replace(CStr(Round(rstTempStocks.Fields(lngCounter).Value, 3) + IIf(Round(rstTempStocks.Fields(lngCounter).Value, 3) < CDbl(rstTempStocks.Fields(lngCounter).Value), 0.001, 0)), ",", ".")
                        Case "GROSS WEIGHT"
                            m_rstStocksOff.Fields(lngCounter).Value = Replace(CStr(Round(rstTempStocks.Fields(lngCounter).Value, 2) + IIf(Round(rstTempStocks.Fields(lngCounter).Value, 2) < CDbl(rstTempStocks.Fields(lngCounter).Value), 0.01, 0)), ",", ".")
                        Case Else
                        
                            m_rstStocksOff.Fields(lngCounter).Value = rstTempStocks.Fields(lngCounter).Value
                    
                    End Select
                    
                Next
                
                m_rstStocksOff.Update
                
            End If
            
            rstTempStocks.MoveNext
                    
        Loop
        
    End If
    
    ADORecordsetClose rstDIA
    ADORecordsetClose rstTempStocks
    
    Set jgxAvailableStocks.ADORecordset = m_rstStocksOff
    
    Call FormatAvailableGrid
    
    Erase arrRow
    ReDim arrRow(1 To jgxRepackage.Columns.Count)
    
    If jgxAvailableStocks.RowCount = 0 Then
        txtGrossWt.Text = 0
        txtGrossWt.Enabled = False
        txtNetWt.Text = 0
        txtNetWt.Enabled = False
        txtNoOfItems.Text = 0
        txtNoOfItems.Enabled = False
        jgxRepackage.Enabled = False
    ElseIf jgxAvailableStocks.RowCount > 0 Then
        jgxRepackage.Enabled = True
    End If
            
End Sub

Private Function CreateSumSQL(ByVal lngHandling As Long) As String
            
    Dim strSQL As String
    
    strSQL = "CSTR(SUM( " & _
            Choose(lngHandling, "Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Orig_Packages_Qty / Inbounds!In_Orig_Gross_Weight * Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Orig_Packages_Qty / Inbounds!In_Orig_Net_Weight * Inbounds!In_Avl_Qty_Wgt") & _
            ")) AS [Num of Items], "
    strSQL = strSQL & "CStr(SUM( " & _
            Choose(lngHandling, "Inbounds!In_Orig_Net_Weight / Inbounds!In_Orig_Packages_Qty * Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Orig_Net_Weight / Inbounds!In_Orig_Gross_Weight * Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Avl_Qty_Wgt") & _
            ")) AS [Net Weight], "
    strSQL = strSQL & "CStr(SUM( " & _
            Choose(lngHandling, "Inbounds!In_Orig_Gross_Weight / Inbounds!In_Orig_Packages_Qty * Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Avl_Qty_Wgt", "Inbounds!In_Orig_Gross_Weight /  Inbounds!In_Orig_Net_Weight * Inbounds!In_Avl_Qty_Wgt") & _
            ")) AS [Gross Weight], "

    CreateSumSQL = strSQL
    
End Function

Private Sub FormatAvailableGrid()
    jgxAvailableStocks.Columns("Num of Items").TextAlignment = jgexAlignRight
    jgxAvailableStocks.Columns("Gross Weight").TextAlignment = jgexAlignRight
    jgxAvailableStocks.Columns("Net Weight").TextAlignment = jgexAlignRight
    jgxAvailableStocks.Columns("Doc Date").Format = "Short Date"
    jgxAvailableStocks.Columns("Num of Items").Width = 1100
    jgxAvailableStocks.Columns("Doc Date").Width = 1100
    jgxAvailableStocks.Columns("InDoc_Office").Visible = False
    jgxAvailableStocks.Columns("InDoc_SeqNum").Visible = False
    jgxAvailableStocks.Columns("InDoc_Cert_Type").Visible = False
    jgxAvailableStocks.Columns("InDoc_Cert_Num").Visible = False
End Sub

Private Sub jgxAvailableStocks_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
        
    'This will prompt the user for loss of data if there records in the repackage grid
    'and resets the grid if approved. IAN
    
    If jgxAvailableStocks.RowCount > 0 Then
        
        If Not blnUserChanged And blnFormLoaded Then
            
            If jgxRepackage.RowCount > 0 Then
            
                If MsgBox(Translate(2300), vbYesNo + vbQuestion, Translate(2301)) = vbNo Then
                    blnUserChanged = True
                    jgxAvailableStocks.Row = LastRow
                    blnUserChanged = False
                    Exit Sub
                End If
            
            End If
            
            Call FillQuantity(True)
            Call CreateOfflineRecordset
            blnGridIsNothing = True
            Set jgxRepackage.ADORecordset = Nothing
            blnGridIsNothing = False
            Set jgxRepackage.ADORecordset = m_rstRepackOff
            Call FormatRepackGrid
                
        End If
                                
    End If
    
End Sub

'Packaging type picklist
Private Sub jgxRepackage_ColButtonClick(ByVal ColIndex As Integer)
                    
    Dim pckPackaging As PCubeLibPick.CPicklist
    Dim gsdPackaging As PCubeLibPick.CGridSeed
    
    Set pckPackaging = New CPicklist
    Set gsdPackaging = pckPackaging.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
    
    With pckPackaging
        If Not IsNull(jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)) Then
            If jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index) <> "" Then
                .Search True, "Key Code", jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)
            End If
        End If
        
        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdPackaging, , , True, cpiKeyF2
        If Not .SelectedRecord Is Nothing Then
            jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
            jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index) = 1
            lngPack_Flag = 1
        End If
    End With
    
    Set gsdPackaging = Nothing
    Set pckPackaging = Nothing

End Sub

'This procedure serves as mask procedure for user text input for the numeric valued
'columns. This limits the decimal places and input text. IAN

Private Sub jgxRepackage_KeyPress(KeyAscii As Integer)

    If jgxRepackage.Col > 0 Then
    
        Select Case UCase(jgxRepackage.Columns(jgxRepackage.Col).Key)
        
            Case "NUM OF ITEMS", "NET WEIGHT", "GROSS WEIGHT"
                If Chr(KeyAscii) = "." Then
                    If UCase(jgxRepackage.Columns(jgxRepackage.Col).Key) = "NUM OF ITEMS" Then
                        KeyAscii = 0
                    ElseIf InStr(CStr(jgxRepackage.Value(jgxRepackage.Col)), ".") Then
                        KeyAscii = 0
                    End If
                ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
                    KeyAscii = 0
                ElseIf UCase(jgxRepackage.Columns(jgxRepackage.Col).Key) <> "NUM OF ITEMS" And IsNumeric(Chr(KeyAscii)) Then
                    Dim lngCount As Long
                    
                    lngCount = IIf(UCase(jgxRepackage.Columns(jgxRepackage.Col).Key) = "NET WEIGHT", 3, 2)
                    If Len(CStr(jgxRepackage.Value(jgxRepackage.Col))) - InStrRev(CStr(jgxRepackage.Value(jgxRepackage.Col)), ".") >= lngCount _
                        And InStr(CStr(jgxRepackage.Value(jgxRepackage.Col)), ".") > 0 And jgxRepackage.SelLength = 0 And _
                        jgxRepackage.SelStart >= InStr(CStr(jgxRepackage.Value(jgxRepackage.Col)), ".") Then
                        
                        KeyAscii = 0
                        
                    End If
                End If
                
        End Select
        
    End If
    
End Sub

Private Sub txtGrossWt_Change()
    
    If blnFormLoaded And txtGrossWt.Enabled And jgxAvailableStocks.RowCount > 0 Then
        Call FillQuantity
    End If
    
End Sub

Private Sub txtGrossWt_KeyPress(KeyAscii As Integer)

    If Chr(KeyAscii) = "." Then
        If InStr(txtGrossWt.Text, ".") Then
            KeyAscii = 0
        End If
    ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    ElseIf IsNumeric(Chr(KeyAscii)) Then
        If Len(txtGrossWt.Text) - InStr(txtGrossWt.Text, ".") >= 2 And _
            InStr(txtGrossWt.Text, ".") > 0 And txtGrossWt.SelLength = 0 And _
            txtGrossWt.SelStart >= InStr(txtGrossWt.Text, ".") Then
            
            KeyAscii = 0
            
        End If
    End If

End Sub

Private Sub txtGrossWt_Validate(Cancel As Boolean)

    If jgxAvailableStocks.RowCount > 0 Then 'IAN
        If ExceedAvailableItemWt(txtGrossWt.Text) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtNetWt_Change()
    
    If blnFormLoaded And txtNetWt.Enabled And jgxAvailableStocks.RowCount > 0 Then
        Call FillQuantity
    End If

End Sub

Private Sub txtNetWt_KeyPress(KeyAscii As Integer)

    If Chr(KeyAscii) = "." Then
        If InStr(txtNetWt.Text, ".") Then
            KeyAscii = 0
        End If
    ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    ElseIf IsNumeric(Chr(KeyAscii)) Then
        If Len(txtNetWt.Text) - InStr(txtNetWt.Text, ".") >= 3 And _
            InStr(txtNetWt.Text, ".") > 0 And txtNetWt.SelLength = 0 And _
            txtNetWt.SelStart >= InStr(txtNetWt.Text, ".") Then
            
            KeyAscii = 0
            
        End If
    End If

End Sub

Private Sub txtNetWt_Validate(Cancel As Boolean)

    If jgxAvailableStocks.RowCount > 0 Then
        If ExceedAvailableItemWt(txtNetWt.Text) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtNoOfItems_Change()
    
    If blnFormLoaded And txtNoOfItems.Enabled And jgxAvailableStocks.RowCount > 0 Then
        Call FillQuantity
    End If
    
End Sub

Private Sub txtNoOfItems_KeyPress(KeyAscii As Integer)

    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtNoOfItems_Validate(Cancel As Boolean)
    
    If jgxAvailableStocks.RowCount > 0 Then 'IAN
        If ExceedAvailableItemWt(txtNoOfItems.Text) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Function ValidateWeight()
    Dim rstDifferences As ADODB.Recordset
    Dim lngAllowableGrossDifference As Long
    Dim lngAllowableNetDifference As Long
    Dim dblOriginalNetWeight As Double
    Dim dblOriginalGrossWeight As Double
    Dim dblGrossWeightDiff As Double
    Dim dblNetWeightDiff As Double
    Dim dblGrossWeight As Double
    Dim dblNetWeight As Double
    Dim strOriginalGrossWt As String
    Dim strOriginalNetWt As String
    Dim strGrossWt As String
    Dim strNetWt As String
    Dim strError As String
    Dim strSQL As String
    Dim lngCounter As Long
    
    strError = ""
    dblGrossWeight = 0
    dblNetWeight = 0
    dblGrossWeightDiff = 0
    dblNetWeightDiff = 0
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "EntrepotProperties!Prop_Net_Diff As [Net Difference], "
        strSQL = strSQL & "EntrepotProperties!Prop_Gross_Diff As [Gross Difference], "
        strSQL = strSQL & "Prop_DisableNetCheck, "
        strSQL = strSQL & "Prop_DisableGrossCheck "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "EntrepotProperties "
    ADORecordsetOpen strSQL, m_conSADBEL, rstDifferences, adOpenKeyset, adLockOptimistic
    'rstDifferences.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    '<<< dandan 110907
    'Corrected checking for recordset records
    If (rstDifferences.EOF Or rstDifferences.BOF) Then
        lngAllowableGrossDifference = 0
        lngAllowableNetDifference = 0
    Else
        rstDifferences.MoveFirst
        
        lngAllowableNetDifference = IIf(IsNull(rstDifferences.Fields("Net Difference")), 0, rstDifferences.Fields("Net Difference"))
        lngAllowableGrossDifference = IIf(IsNull(rstDifferences.Fields("Gross Difference")), 0, rstDifferences.Fields("Gross Difference"))
    End If
    
    
    'Saving values of the recordset into a string is done to avoid the
    'computation error caused by the decimal point being removed
    'when the language is set to Dutch.
    strOriginalNetWt = IIf(IsNull(Val(txtNetWt.Text)), 0, Val(txtNetWt.Text))
    strOriginalGrossWt = IIf(IsNull(Val(txtGrossWt.Text)), 0, Val(txtGrossWt.Text))
    dblOriginalNetWeight = CDbl(strOriginalNetWt)
    dblOriginalGrossWeight = CDbl(strOriginalGrossWt)


    jgxRepackage.MoveFirst
    
    For lngCounter = 1 To jgxRepackage.RowCount
        'Saving values of the recordset into a string is done to avoid the
        'computation error caused by the decimal point being removed
        'when the language is set to Dutch.
        strGrossWt = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index))
        strNetWt = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index))
        dblGrossWeight = dblGrossWeight + Val(strGrossWt)
        dblNetWeight = dblNetWeight + Val(strNetWt)
        jgxRepackage.MoveNext
    Next
    
    If rstDifferences!Prop_DisableGrossCheck = 0 Then 'if false then get deviation
        If dblOriginalGrossWeight <> 0 Then
            'dblGrossWeightDiff = Round((dblOriginalGrossWeight - dblGrossWeight) / dblOriginalGrossWeight, 2)
            dblGrossWeightDiff = (dblOriginalGrossWeight - dblGrossWeight) / dblOriginalGrossWeight
        End If
    End If
    
    
    If rstDifferences!Prop_DisableNetCheck = 0 Then 'if false then get deviation
        If dblOriginalNetWeight <> 0 Then
            'dblNetWeightDiff = Round((dblOriginalNetWeight - dblNetWeight) / dblOriginalNetWeight, 3)
            dblNetWeightDiff = (dblOriginalNetWeight - dblNetWeight) / dblOriginalNetWeight
        End If
    End If
    
    
    If dblGrossWeightDiff > 0 Then
        If dblGrossWeightDiff > (lngAllowableGrossDifference / 100) Then
            strError = strError & Translate(2302) & Space(1) & lngAllowableGrossDifference & Translate(2303) & vbCrLf
        End If
    End If
    
    If dblNetWeightDiff > 0 Then
        If dblNetWeightDiff > (lngAllowableNetDifference / 100) Then
            strError = strError & Translate(2305) & Space(1) & lngAllowableNetDifference & "% allowable deviation. " & vbCrLf
        End If
    End If
    
    'add weight checking for repackaged gross should be >= repackaged net
    If dblGrossWeight < dblNetWeight Then
        strError = strError & "Repackaged gross is less than repackaged net." & vbCrLf
    End If
    
    If strError <> "" Then
        MsgBox strError, vbInformation + vbOKOnly, Translate(2307)
        ValidateWeight = False
    Else
        ValidateWeight = True
    End If
    
    ADORecordsetClose rstDifferences
    
End Function

Private Function InsertRepackagingInbounds() As Long

End Function

Private Sub Cancelling(ByRef SADBELDB As ADODB.Connection, _
                       ByRef HistoryDB As ADODB.Connection, _
                       ByVal DataType As String)
    Dim strSQL As String
    Dim strNetWt As String
    Dim strGrossWt As String
    Dim lngCounter As Long
    Dim lngRowCount As Long
    Dim rstRepackaged As ADODB.Recordset
    Dim lngInID As Long
    
    strSQL = "Select * from [Inbounds] "
    
    If DataType = "SADBEL" Then
        lngInID = 0
    Else
        lngInID = m_lngInID
    End If
    
    
    ADORecordsetOpen strSQL, SADBELDB, rstRepackaged, adOpenKeyset, adLockOptimistic
    'rstRepackaged.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic


    'Inbounds
    Do
         
        rstRepackaged.AddNew
        If DataType <> "SADBEL" Then
            rstRepackaged.Fields("In_ID").Value = lngInID
        End If
        rstRepackaged.Fields("In_Orig_Packages_Qty").Value = (-1) * Val(txtNoOfItems.Text)
        rstRepackaged.Fields("In_Orig_Packages_Type").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Package Type").Index)
        rstRepackaged.Fields("In_Orig_Net_Weight").Value = (-1) * Val(txtNetWt.Text)
        rstRepackaged.Fields("In_Orig_Gross_Weight").Value = (-1) * Val(txtGrossWt.Text)
        rstRepackaged.Fields("In_Batch_Num").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Batch Num").Index)
        rstRepackaged.Fields("In_Job_Num").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Job Num").Index)
        rstRepackaged.Fields("In_TotalOut_Qty_Wgt").Value = 0
        rstRepackaged.Fields("In_Reserved_Qty_Wgt").Value = 0
        rstRepackaged.Fields("Stock_ID").Value = lblStockCardNum.Tag
        rstRepackaged.Fields("InDoc_ID").Value = m_lngInDoc

        If lblHandling.Caption <> "" Then
            Select Case lblHandling.Caption
                Case "Number of Items"
                    rstRepackaged.Fields("In_Avl_Qty_Wgt").Value = (-1) * Val(txtNoOfItems.Text)
                Case "Gross Weight"
                    rstRepackaged.Fields("In_Avl_Qty_Wgt").Value = (-1) * Val(txtGrossWt.Text)
                Case "Net Weight"
                    rstRepackaged.Fields("In_Avl_Qty_Wgt").Value = (-1) * Val(txtNetWt.Text)
            End Select
        End If
        rstRepackaged.Update
        
        If DataType = "SADBEL" Then
            lngInID = InsertRecordset(SADBELDB, rstRepackaged, "Inbounds")
        Else
            lngInID = InsertRecordset(SADBELDB, rstRepackaged, "Inbounds")
            Debug.Assert m_lngInID <> lngInID
            ' TO DO FOR CP.NET - AUTONUMBER In_ID must be set to lngInID and not generated
            
        End If
       
    Loop While Not (IsIDUnique(SADBELDB, "In_ID", "Inbounds", lngInID) And _
                    IsIDUnique(HistoryDB, "In_ID", "Inbounds", lngInID)) Or _
               (DataType = "SADBEL")

    m_lngInID = lngInID
    
    ADORecordsetClose rstRepackaged
End Sub

Private Sub UpdateInboundsTable(ByRef ADOConnection As ADODB.Connection, _
                                ByVal DataType As String)
    Dim strSQL As String
    Dim strNetWt As String
    Dim strGrossWt As String
    Dim lngCounter As Long
    Dim lngRowCount As Long
    Dim rstToBeRepackaged As ADODB.Recordset
    Dim lngInID As Long
    
    strSQL = "SELECT * FROM [Inbounds] "
    
    ADORecordsetOpen strSQL, ADOConnection, rstToBeRepackaged, adOpenKeyset, adLockOptimistic
    'rstToBeRepackaged.Open strSQL, ADOConnection, adOpenKeyset, adLockOptimistic

    jgxRepackage.MoveFirst
    
    For lngCounter = 1 To jgxRepackage.RowCount

        If DataType = "SADBEL" Then
            lngInID = 0
        Else
            lngInID = m_alngInID2(lngCounter)
        End If
        
        Do
            rstToBeRepackaged.AddNew
            
            If DataType <> "SADBEL" Then
                rstToBeRepackaged.Fields("In_ID").Value = lngInID
            End If
             
            strNetWt = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index))
            strGrossWt = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index))
            
            rstToBeRepackaged.Fields("In_Orig_Packages_Qty").Value = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Num Of Items").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Num Of Items").Index))
            rstToBeRepackaged.Fields("In_Orig_Packages_Type").Value = jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)
            rstToBeRepackaged.Fields("In_Orig_Net_Weight").Value = Val(strNetWt)
            rstToBeRepackaged.Fields("In_Orig_Gross_Weight").Value = Val(strGrossWt)
            rstToBeRepackaged.Fields("In_Batch_Num").Value = jgxRepackage.Value(jgxRepackage.Columns("Batch Num").Index)
            rstToBeRepackaged.Fields("In_Job_Num").Value = jgxRepackage.Value(jgxRepackage.Columns("Job Num").Index)
            rstToBeRepackaged.Fields("In_TotalOut_Qty_Wgt").Value = 0
            rstToBeRepackaged.Fields("In_Reserved_Qty_Wgt").Value = 0
            rstToBeRepackaged.Fields("Stock_ID").Value = lblStockCardNum.Tag
            rstToBeRepackaged.Fields("InDoc_ID").Value = m_lngInDoc
            rstToBeRepackaged.Fields("In_Source_In_ID").Value = m_lngInID
            
            If lblHandling.Caption <> "" Then
                Select Case lblHandling.Caption
                    Case "Number of Items"
                        rstToBeRepackaged.Fields("In_Avl_Qty_Wgt").Value = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Num of Items").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Num of Items").Index))
                    Case "Gross Weight"
                        rstToBeRepackaged.Fields("In_Avl_Qty_Wgt").Value = Val(IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Gross Weight").Index)))
                    Case "Net Weight"
                        rstToBeRepackaged.Fields("In_Avl_Qty_Wgt").Value = Val(IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index)), 0, jgxRepackage.Value(jgxRepackage.Columns("Net Weight").Index)))
                End Select
            End If
            rstToBeRepackaged.Update
                        
            If DataType = "SADBEL" Then
                lngInID = InsertRecordset(ADOConnection, rstToBeRepackaged, "Inbounds")
            Else
                lngInID = InsertRecordset(ADOConnection, rstToBeRepackaged, "Inbounds")
                Debug.Assert m_alngInID2(lngCounter) <> lngInID
                ' TO DO FOR CP.NET - AUTONUMBER In_ID must be set to lngInID and not generated
                
            End If
        
        Loop While Not IsIDUnique(ADOConnection, "In_ID", "Inbounds", lngInID) Or _
                   (DataType = "SADBEL")

        m_alngInID2(lngCounter) = lngInID
        
        jgxRepackage.MoveNext
    
    Next lngCounter
    
    ADORecordsetClose rstToBeRepackaged

End Sub

Private Sub UpdateInboundDocsTable(ByRef ADOConnection As ADODB.Connection, _
                                   ByVal DataType As String)
    Dim strSQL As String
    Dim lngCounter As Long
    Dim lngRowCount As Long
    Dim rstToBeRepackaged As ADODB.Recordset
    Dim lngInDocID As Long
    Dim strDocDate As String

    strSQL = "SELECT * FROM [InboundDocs] "

    ADORecordsetOpen strSQL, ADOConnection, rstToBeRepackaged, adOpenKeyset, adLockOptimistic
    'rstToBeRepackaged.Open strSQL, ADOConnection, adOpenKeyset, adLockOptimistic
    
    If DataType = "SADBEL" Then
        lngInDocID = 0
        strDocDate = Now
    Else
        lngInDocID = m_lngInDoc
        strDocDate = m_strDocDate
    End If
    
    Do
        rstToBeRepackaged.AddNew
        If DataType <> "SADBEL" Then
            rstToBeRepackaged.Fields("InDoc_ID").Value = lngInDocID
        End If
        rstToBeRepackaged.Fields("InDoc_Type").Value = "IM7"
        rstToBeRepackaged.Fields("InDoc_Num").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Doc Number").Index)
        rstToBeRepackaged.Fields("InDoc_Office").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("InDoc_Office").Index)
        rstToBeRepackaged.Fields("InDoc_SeqNum").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("InDoc_SeqNum").Index)
        rstToBeRepackaged.Fields("InDoc_Cert_Type").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("InDoc_Cert_Type").Index)
        rstToBeRepackaged.Fields("InDoc_Cert_Num").Value = jgxAvailableStocks.Value(jgxAvailableStocks.Columns("InDoc_Cert_Num").Index)
        rstToBeRepackaged.Fields("InDoc_Date").Value = m_strDocDate
        rstToBeRepackaged.Update
            
        If DataType = "SADBEL" Then
            lngInDocID = InsertRecordset(ADOConnection, rstToBeRepackaged, "InboundDocs")
        Else
            lngInDocID = InsertRecordset(ADOConnection, rstToBeRepackaged, "InboundDocs")
            Debug.Assert m_lngInDoc <> lngInDocID
            ' TO DO FOR CP.NET - AUTONUMBER In_ID must be set to lngInID and not generated
            
        End If
            
    Loop While Not IsIDUnique(ADOConnection, "InDoc_ID", "InboundDocs", lngInDocID) Or _
               (DataType = "SADBEL")

    m_lngInDoc = lngInDocID
    m_strDocDate = strDocDate

    ADORecordsetClose rstToBeRepackaged

End Sub

Private Sub OverwriteInboundsTable(ByRef ADOConnection As ADODB.Connection, DataType As String)
    Dim strSQL As String
    Dim strNetWt As String
    Dim strGrossWt As String
    Dim strJobNumber As String
    Dim strBatchNumber As String
    Dim lngCounter As Long
    Dim lngRowCount As Long
    Dim lngSource_ID As Long
    
    Dim rstOverwrite As ADODB.Recordset
    Dim dblAvailableLeft As Double
    Dim dblAvailable As Double
    
    Dim rstInbounds As ADODB.Recordset
    Dim strCommand As String
    Dim lngIn_ID As Long
    Dim lngInDoc_ID As Long
    
        strSQL = " SELECT INBOUNDS.IN_ID AS IN_ID, INBOUNDS.In_Source_In_ID as [Source ID], " & _
                " INBOUNDS.In_TotalOut_Qty_Wgt AS [Total Out], " & _
                " INBOUNDS.InDoc_ID as [InDoc_ID], " & _
                " INBOUNDS.In_Reserved_Qty_Wgt AS [Reserved], " & _
                " INBOUNDS.In_Avl_Qty_Wgt AS [Available] " & _
                " FROM INBOUNDS INNER JOIN INBOUNDDOCS ON " & _
                " INBOUNDS.INDOC_ID = INBOUNDDOCS.INDOC_ID " & _
                " WHERE INBOUNDS.Stock_ID = " & lblStockCardNum.Tag & _
                " AND INBOUNDDOCS.InDoc_Num = '" & jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Doc Number").Index) & "'" & _
                " AND INBOUNDS.In_Orig_Packages_Type = '" & jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Package Type").Index) & "'" & _
                " AND INBOUNDS.In_Batch_Num "
        
        strBatchNumber = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Batch Num").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Batch Num").Index))
        strJobNumber = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Job Num").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Job Num").Index))
        
        If strBatchNumber = "" Then
            strSQL = strSQL & " IS NULL AND INBOUNDS.In_Job_Num "
        Else
            strSQL = strSQL & " = '" & strBatchNumber & "' AND INBOUNDS.In_Job_Num "
        End If
        
        If strJobNumber = "" Then
            strSQL = strSQL & "IS NULL"
        Else
            strSQL = strSQL & " = '" & strJobNumber & "'"
        End If
        
        strSQL = strSQL & " AND INBOUNDS.In_Avl_Qty_Wgt > 0"
    ADORecordsetOpen strSQL, ADOConnection, rstOverwrite, adOpenKeyset, adLockOptimistic
    'rstOverwrite.Open strSQL, ADOConnection, adOpenKeyset, adLockOptimistic

    If rstOverwrite.EOF And rstOverwrite.BOF Then
        'If there are no records found, exit the procedure.
        ADORecordsetClose rstOverwrite
        
        Exit Sub
    Else
        rstOverwrite.MoveFirst
        
        Select Case lblHandling.Caption
            Case "Number of Items"
                dblAvailableLeft = Val(txtNoOfItems.Text)
            Case "Gross Weight"
                dblAvailableLeft = Val(txtGrossWt.Text)
            Case "Net Weight"
                dblAvailableLeft = Val(txtNetWt.Text)
        End Select
        
        For lngCounter = 1 To rstOverwrite.RecordCount
            
            dblAvailable = rstOverwrite.Fields("Available").Value
            lngSource_ID = IIf(IsNull(rstOverwrite.Fields("Source ID").Value), 0, rstOverwrite.Fields("Source ID").Value)
            
            lngInDoc_ID = rstOverwrite.Fields("InDoc_ID").Value
            lngIn_ID = rstOverwrite.Fields("IN_ID").Value
            
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "[IN_ID] AS [IN_ID], "
                strCommand = strCommand & "[In_Source_In_ID] as [Source ID] "
                strCommand = strCommand & "[In_TotalOut_Qty_Wgt] AS [Total Out], "
                strCommand = strCommand & "[InDoc_ID] as [InDoc_ID], "
                strCommand = strCommand & "[In_Reserved_Qty_Wgt] AS [Reserved], "
                strCommand = strCommand & "[In_Avl_Qty_Wgt] AS [Available] "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[Inbounds] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "[In_ID] = " & lngIn_ID & " "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[INDOC_ID] = " & lngInDoc_ID & " "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[Stock_ID] = " & lblStockCardNum.Tag & " "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[InDoc_Num] = '" & jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Doc Number").Index) & "' "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[In_Orig_Packages_Type] = '" & jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Package Type").Index) & "' "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[In_Batch_Num] "
                
                If strBatchNumber = "" Then
                    strCommand = strCommand & "IS NULL "
                    strCommand = strCommand & "AND "
                    strCommand = strCommand & "[In_Job_Num] "
                Else
                    strCommand = strCommand & " = '" & strBatchNumber & "' "
                    strCommand = strCommand & "AND "
                    strCommand = strCommand & "[In_Job_Num] "
                End If
                
                If strJobNumber = "" Then
                    strCommand = strCommand & "IS NULL "
                Else
                    strCommand = strCommand & " = '" & strJobNumber & "' "
                End If
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[In_Avl_Qty_Wgt] > 0 "
        
            ADORecordsetOpen strCommand, ADOConnection, rstInbounds, adOpenKeyset, adLockOptimistic
            If Not (rstInbounds.EOF And rstInbounds.BOF) Then
                rstInbounds.MoveFirst
                
                With rstOverwrite
                    If dblAvailable < dblAvailableLeft Then
                        rstInbounds![Total Out] = rstInbounds![Total Out] + dblAvailable
                        dblAvailableLeft = dblAvailableLeft - rstInbounds.Fields("Available").Value
                        dblAvailable = 0
                    Else
                        rstInbounds![Total Out] = rstInbounds![Total Out] + dblAvailableLeft
                        dblAvailable = rstInbounds.Fields("Available").Value - dblAvailableLeft
                        dblAvailableLeft = 0
                    End If
                    
                    rstInbounds.Fields("Available").Value = dblAvailable
                    
                    'If there are no more available items in the record, delete it.
                    If rstInbounds.Fields("Available").Value = 0 Then
                        If DataType = "SADBEL" Then
                            If rstInbounds.Fields("Reserved").Value = 0 Then
                                .Delete
                                ' TO DO FOR CP.NET
                                ExecuteNonQuery ADOConnection, GetDeleteCommandFromSelect(strSQL, "Inbounds")
                                
                                'Delete the record if all of its repackaged items had been used up.
                                If lngSource_ID <> 0 Then Call UpdateTheNegativeAvailableQty(ADOConnection, lngSource_ID)
                                
                                'Delete an InboundDocs record if it's not referenced anymore
                                'by the records in the Inbounds table.
                                Call DeleteInboundDocs(ADOConnection, lngInDoc_ID)
                            Else
                                rstInbounds.Update
                                
                                UpdateRecordset ADOConnection, rstInbounds, "Inbounds"
                            End If
                        Else
                            rstInbounds.Update
                            UpdateRecordset ADOConnection, rstInbounds, "Inbounds"
                        End If
                    Else
                        rstInbounds.Update
                        UpdateRecordset ADOConnection, rstInbounds, "Inbounds"
                    End If
                    
                End With
                
            End If
            
            If dblAvailableLeft <> 0 Then
                If Not rstOverwrite.EOF Then
                    rstOverwrite.MoveNext
                End If
            Else
                Exit For
            End If
        Next
    End If
    
    ADORecordsetClose rstOverwrite
End Sub

Private Function IsIDUnique(ByRef ADOConnection As ADODB.Connection, _
                            ByVal FieldName As String, _
                            ByVal TableName As String, _
                            ByVal ID As Long) As Boolean
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

Private Sub UpdateTheNegativeAvailableQty(ByRef ADOConnection As ADODB.Connection, _
                                          ByVal lngSource_ID As Long)
    Dim strSQL As String
    Dim strSQL2 As String
    Dim blnToDelete As Boolean
    Dim rstUpdate As ADODB.Recordset
    Dim rstUpdate2 As ADODB.Recordset
    
        strSQL = "SELECT * FROM INBOUNDS WHERE INBOUNDS.In_Source_In_ID = " & lngSource_ID
    ADORecordsetOpen strSQL, ADOConnection, rstUpdate, adOpenKeyset, adLockOptimistic
    'rstUpdate.Open strSQL, ADOConnection, adOpenKeyset, adLockOptimistic
    
    'If there are no more records that point to a certain record, delete it.
    If rstUpdate.EOF And rstUpdate.BOF Then
            strSQL2 = " SELECT * FROM INBOUNDS WHERE INBOUNDS.In_ID = " & lngSource_ID
        ADORecordsetOpen strSQL2, ADOConnection, rstUpdate2, adOpenKeyset, adLockOptimistic
        'rstUpdate2.Open strSQL2, ADOConnection, adOpenKeyset, adLockOptimistic
        
        If Not (rstUpdate2.EOF And rstUpdate2.BOF) Then
            rstUpdate2.Delete
            
            ExecuteNonQuery ADOConnection, "DELETE * FROM INBOUNDS WHERE INBOUNDS.In_ID = " & lngSource_ID
        End If
        
        ADORecordsetClose rstUpdate2
        
    End If
    
    ADORecordsetClose rstUpdate
End Sub

Private Sub DeleteInboundDocs(ByRef ADOConnection As ADODB.Connection, _
                              ByVal lngInDoc_ID As Long)
    Dim strSQL As String
    Dim strSQL2 As String
    Dim rstInbounds As ADODB.Recordset
    Dim rstInboundDocs As ADODB.Recordset
    
        strSQL = " SELECT * FROM INBOUNDS WHERE InDoc_ID = " & lngInDoc_ID
    ADORecordsetOpen strSQL, ADOConnection, rstInbounds, adOpenKeyset, adLockOptimistic
    'rstInbounds.Open strSQL, ADOConnection, adOpenKeyset, adLockOptimistic
    
    'If the InDoc_ID of the InboundDocs table is not referenced anymore in the
    'Inbounds table, delete it.
    If rstInbounds.EOF And rstInbounds.BOF Then
            strSQL2 = " SELECT * FROM INBOUNDDOCS WHERE InDoc_ID = " & lngInDoc_ID
        ADORecordsetOpen strSQL2, ADOConnection, rstInboundDocs, adOpenKeyset, adLockOptimistic
        'rstInboundDocs.Open strSQL2, ADOConnection, adOpenKeyset, adLockOptimistic
        
        If Not (rstInboundDocs.EOF And rstInboundDocs.BOF) Then
            rstInboundDocs.Delete
            
            ExecuteNonQuery ADOConnection, "DELETE * FROM INBOUNDS WHERE InDoc_ID = " & lngInDoc_ID
        End If
        
        ADORecordsetClose rstInboundDocs
    End If
    
    ADORecordsetClose rstInbounds
End Sub

Private Function CheckIfRowToAdd() As Boolean

    Dim lngCtr As Long
    Dim lngCounter As Long
    Dim UserReply As VbMsgBoxResult
    Dim lngRowCount As Long
    Dim varValue As Variant
    
    If jgxRepackage.Enabled = False Or jgxRepackage.Row <> -1 Then
        CheckIfRowToAdd = True
        Exit Function
    End If
    
    For lngCtr = 1 To jgxRepackage.Columns.Count
        
        If m_rstRepackOff.Fields(jgxRepackage.Columns(lngCtr).DataField).Type = adInteger Then
            varValue = Val(jgxRepackage.Value(lngCtr))
        ElseIf m_rstStocksOff.Fields(jgxRepackage.Columns(lngCtr).DataField).Type = adBoolean Then
            varValue = CBool(jgxRepackage.Value(lngCtr))
        Else
            varValue = IIf(IsNull(jgxRepackage.Value(lngCtr)), "", jgxRepackage.Value(lngCtr))
        End If
        
        If varValue <> IIf(IsNull(jgxRepackage.Columns(lngCtr).DefaultValue), "", jgxRepackage.Columns(lngCtr).DefaultValue) Then
            UserReply = MsgBox("A record is waiting to be added. Would you like to add it now?", vbYesNoCancel + vbQuestion, Translate(2297))
            If UserReply = vbYes Then
                
                blnEntered = True
                lngRowCount = jgxRepackage.RowCount
                jgxRepackage.Update
                
                If jgxRepackage.RowCount = lngRowCount + 1 Then
                    CheckIfRowToAdd = True
                Else
                    CheckIfRowToAdd = False
                End If
                
            ElseIf UserReply = vbNo Then
                CheckIfRowToAdd = True
                For lngCounter = 1 To jgxRepackage.Columns.Count
                    jgxRepackage.Value(lngCounter) = IIf(IsNull(jgxRepackage.Columns(lngCounter).DefaultValue), "", jgxRepackage.Columns(lngCounter).DefaultValue)
                Next
            Else
                CheckIfRowToAdd = False
            End If
            Exit Function
        End If
    Next
    
    CheckIfRowToAdd = True
    
End Function

Private Function CheckInvalidRecord() As Boolean
                        
    If lngPack_Flag = 0 Then
                            
        Dim strPack As String
        
        strPack = IIf(IsNull(jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index)), "", jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index))
        
        Set pckList = New CPicklist
        Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
                    
        With pckList
    
            If Trim(strPack) <> "" Then
                .Search True, "Key Code", strPack
            End If
            ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
            .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyEnter
                            
            If Not .SelectedRecord Is Nothing Then
                jgxRepackage.Value(jgxRepackage.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                jgxRepackage.Value(jgxRepackage.Columns("Pack_Flag").Index) = 1
                lngPack_Flag = 1
            Else
                blnSystemChanged = True
                jgxRepackage.Col = jgxRepackage.Columns("Package Type").Index
                blnSystemChanged = False
                CheckInvalidRecord = True
                Set pckList = Nothing
                Set gsdList = Nothing
                Exit Function
            End If
    
        End With
        
        Set pckList = Nothing
        Set gsdList = Nothing
        
    End If

End Function
