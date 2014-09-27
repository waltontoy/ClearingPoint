VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmStockProdPicklist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products and Stock Cards"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "frmStockProdPicklist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStockCard 
      Caption         =   " Stock Cards "
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Tag             =   "2215"
      Top             =   2520
      Width           =   6615
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Tag             =   "149"
         Top             =   240
         Width           =   1095
      End
      Begin GridEX16.GridEX jgxPicklist 
         Height          =   2535
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4471
         TabKeyBehavior  =   1
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         Options         =   -1
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnCount     =   1
         CardCaption1    =   -1  'True
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Frame fraProduct 
      Caption         =   " Products "
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Tag             =   "2216"
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtTaricCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&More..."
         Height          =   375
         Left            =   5330
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtProductNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdProductNo 
         Caption         =   "..."
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtProductDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3315
      End
      Begin VB.Frame fraInfo 
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   5175
         Begin VB.TextBox txtCtryOrigin 
            Height          =   315
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   7
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox txtCtryExport 
            Height          =   315
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   10
            Top             =   540
            Width           =   495
         End
         Begin VB.CommandButton cmdCountry 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4680
            TabIndex        =   9
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdCountry 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   4680
            TabIndex        =   12
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox txtCtryOriginDesc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   2295
         End
         Begin VB.TextBox txtCtryExportDesc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   540
            Width           =   2295
         End
         Begin VB.Label lblCtryOrigin 
            BackStyle       =   0  'Transparent
            Caption         =   "Country of Origin:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Tag             =   "2195"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblCtryExport 
            BackStyle       =   0  'Transparent
            Caption         =   "Country of Export:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Tag             =   "2196"
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Label lblTaricCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Taric Code:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Tag             =   "2275"
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblProductNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Tag             =   "2274"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblProductDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description:"
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Tag             =   "2201"
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Tag             =   "426"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Tag             =   "179"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtJobNo 
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   60
      Width           =   2655
   End
   Begin VB.TextBox txtBatchNo 
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblBatchNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Number:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Tag             =   "2222"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblJobNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Job Number:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Tag             =   "2221"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmStockProdPicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form is called from:
'   1. Import Codisheet
'   2. Repackaging -> Product Number picklist
'   3. Summary Reports -> Filter type: Stock Card

Option Explicit

Private pckProducts As PCubeLibEntrepot.cProducts
Private pckStockProd As PCubeLibEntrepot.cStockProd
Private pckCountry As PCubeLibPick.CPicklist

Private m_rstFindSeveral As ADODB.Recordset
Public m_rstPass2GridOff As ADODB.Recordset
Public m_rstNewStockOff As ADODB.Recordset

Private blnRstIsNothing As Boolean
Private bytProdNumDeclined As Byte  '0-Yes, 1-No
Private bytProdFound As Byte        '0-No, 1-Yes
Private bytCtryOFound As Byte       '0-No, 1-Yes
Private bytCtryEFound As Byte       '0-No, 1-Yes
Private strProdNum As String
Private strCtryOrigin As String
Private strCtryExport As String
Private blnLesserForm As Boolean
Private alngDefEmptyVal(1 To 3) As Long     '1 - Box C1 (Country of Import)
                                            '2 - Box C2 (Country of Export)
                                            '3 - Box L1 (Taric Code)
Private bytCtryKeys As Byte         '0-No, 1-Yes (for VB bug in GotFocus not working)
Private strBlah As String

Public strStockCardNoHigh As String
Public bytNumbering As Byte
Public strStartingNum As String
Public strTaricCode As String
Public lngEntrepotID As Long
Public strEntrepotType As String
Public strEntrepotNum As String
Public blnCancelled As Boolean

Private blnWithEntrepotNum As Boolean
Private blnIsInitialStock As Boolean
Private strColumnName As String
Private lngSortCounter As Long

Private strCountryExp As String
Private strCountryOrig As String
Private strCountryExpDesc As String
Private strCountryOrigDesc As String
Private strBatchNum As String
Private strJobNum As String
Private strEntrepotName As String
Private m_blnFromSummaryReport As Boolean

Public Sub Pre_Load(ByRef cpiStockProd As PCubeLibEntrepot.cStockProd, ByRef Cancelled As Boolean, ByVal MyResourceHandler As Long, _
                    Optional ByVal blnDontShowBatchJob As Boolean, Optional ByVal blnInitialStock As Boolean, _
                    Optional ByVal blnWithEntrepot As Boolean, Optional ByVal strEntrepotNum As String, Optional blnFromSummaryReports As Boolean = False)
                    
    ResourceHandler = MyResourceHandler
    modGlobals.LoadResStrings Me, True
    
    'Flag not to close m_rstPass2GridOff in case Cancel is invoked immediately after load.
    blnRstIsNothing = True
    blnIsInitialStock = blnInitialStock
    blnWithEntrepotNum = blnWithEntrepot
    
    strEntrepotName = strEntrepotNum
    m_blnFromSummaryReport = blnFromSummaryReports
    
    Set pckStockProd = cpiStockProd
    Set pckProducts = New PCubeLibEntrepot.cProducts
    
    'Passes selected Entrepot Type-Num to cProduct so Products list will be filtered according to Entrepot.
    If Len(pckStockProd.Entrepot_Num) <> 0 And pckStockProd.Entrepot_Num <> "0" Then pckProducts.Entrepot_Num = pckStockProd.Entrepot_Num
    If Len(pckStockProd.TaricCode) <> 0 And pckStockProd.TaricCode <> "0" Then pckProducts.Taric_Code = pckStockProd.TaricCode
    If Len(pckStockProd.CtryOrigin) <> 0 And pckStockProd.CtryOrigin <> "0" Then pckProducts.Ctry_Origin = pckStockProd.CtryOrigin
    If Len(pckStockProd.CtryExport) <> 0 And pckStockProd.CtryExport <> "0" Then pckProducts.Ctry_Export = pckStockProd.CtryExport
    
    jgxPicklist.DefaultColumnWidth = 1200
    
    'Passes default empty values to array for recognition.
    '??? ano ginawa dito sa procedure na 'to???...comment ko na lng muna - alg======
    'GetDefaultEmptyVal alngDefEmptyVal(1), alngDefEmptyVal(2), alngDefEmptyVal(3)
    
    If blnIsInitialStock = False Then
        If blnWithEntrepotNum = False Then
            With pckStockProd
                'Loads values from codisheet.
                txtTaricCode.Text = .TaricCode
                strCountryOrig = .CtryOrigin
                strCountryExp = .CtryExport
                strBlah = GetCountryDesc(.CtryOrigin, .m_conSADBEL, .strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryOrigDesc = strBlah
                strBlah = GetCountryDesc(.CtryExport, .m_conSADBEL, .strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryExpDesc = strBlah
                
                'Will perform auto loading of value if probable product no is found in the memo field.
                ParseMemo pckStockProd.Memo, ":", vbCrLf
                
                txtJobNo.Text = .JobNo
                txtBatchNo.Text = .BatchNo
                strJobNum = txtJobNo.Text
                strBatchNum = txtBatchNo.Text
                
                'Flag that StockProd was called from Summary Reports.
                blnLesserForm = blnDontShowBatchJob
            
                If blnLesserForm = True Then
                    'Only performs load other info by Product ID if called from Summary Report.
                    If Not (.Product_ID = Empty) Then AutoLoadByProdID .Product_ID
                    'Hides Job Num, Batch Num and cmdNew controls.
                    HideBatchJob
                    'Cancel current selection.
                    jgxPicklist.RowSelected(jgxPicklist.Row) = False
                ElseIf blnLesserForm = False Then
                    'Only enables New if called from codisheet.
                    If Len(txtProductNo.Text) <> 0 Then cmdNew.Enabled = True
                End If
                
                'Moves selection according to Stock Card Number in Summary Report, if not empty.
                If Not (.StockCardNo = Empty) Then
                    Dim lngFind As Long
                    
                    If Not (jgxPicklist.ADORecordset.BOF And jgxPicklist.ADORecordset.EOF) Then jgxPicklist.ADORecordset.MoveFirst
                    'Obtains row index for the grid.rowselected property.
                    For lngFind = 1 To jgxPicklist.ADORecordset.RecordCount
                        If jgxPicklist.ADORecordset.Fields("Stock Card No").Value = pckStockProd.StockCardNo Then Exit For
                        jgxPicklist.ADORecordset.MoveNext
                    Next lngFind
                    
                    'Set row selected to row index matching Stock ID from property.
'                    jgxPicklist.RowSelected(lngFind) = True
                    jgxPicklist.MoveLast
                    If lngFind <= jgxPicklist.ADORecordset.RecordCount Then
                        jgxPicklist.Row = lngFind
                    Else
                        jgxPicklist.Row = 1
                    End If
                End If
            End With
        Else
            
            PopGrid2 pckStockProd.m_conSADBEL
            
            With frmStockProdPicklist
                .Caption = "Stock Cards"
                
                .cmdOK.Left = 5615
                .fraStockCard.Left = 120
                .cmdCancel.Left = .cmdOK.Left + 1320
                
                .fraStockCard.Height = .fraStockCard.Height + 350
                .Height = .fraStockCard.Height + (.cmdOK.Height * 2) + 480
                .jgxPicklist.Height = 2975
                
                .fraStockCard.Top = 120
                .cmdOK.Top = .fraStockCard.Height + 240
                .cmdCancel.Top = .fraStockCard.Height + 240
                
                .fraProduct.Visible = False
                .lblJobNo.Visible = False
                .lblBatchNo.Visible = False
                .txtJobNo.Visible = False
                .txtBatchNo.Visible = False
                .cmdNew.Visible = False
                
                .Width = 8380
                .fraStockCard.Width = 8040
                .jgxPicklist.Width = 7790
                .jgxPicklist.Columns("Entrepot Number").Width = 1505
                .jgxPicklist.Columns("Product Number").Width = 1505
                .jgxPicklist.Columns("Stock Card No").Width = 1400
                .jgxPicklist.Columns("Job No").Width = 1005
                .jgxPicklist.Columns("Batch No").Width = 1005
                .jgxPicklist.Columns("Doc Number").Width = 1505
            End With
        
            If jgxPicklist.RowCount > 0 Then
                Dim lngFind3 As Long
                
                jgxPicklist.ADORecordset.MoveFirst
                'Obtains row index for the grid.rowselected property.
                For lngFind3 = 1 To jgxPicklist.ADORecordset.RecordCount
                    If jgxPicklist.ADORecordset.Fields("Stock Card No").Value = pckStockProd.StockCardNo Then
                        If Trim(pckStockProd.Entrepot_Num) = "" Then
                            Exit For
                        ElseIf jgxPicklist.ADORecordset.Fields("Entrepot Number").Value = pckStockProd.Entrepot_Num Then
                            Exit For
                        End If
                    End If
                    jgxPicklist.ADORecordset.MoveNext
                Next lngFind3
                
                'Set row selected to row index matching Stock ID from property.
                jgxPicklist.MoveLast
                If lngFind3 <= jgxPicklist.ADORecordset.RecordCount Then
                    jgxPicklist.Row = lngFind3
                Else
                    jgxPicklist.Row = 1
                End If
                Call jgxPicklist_Click
            End If
        End If
    Else
        cmdNew.Enabled = True
        AutoLoadByProd pckStockProd.ProductNo
        PopGrid pckStockProd.Product_ID, pckStockProd.m_conSADBEL
        lngEntrepotID = GetEntrepot_ID(pckStockProd.Entrepot_Num, pckStockProd.m_conSADBEL)
        strEntrepotType = UCase(Mid$(pckStockProd.Entrepot_Num, 1, 1))
        strEntrepotNum = Mid$(pckStockProd.Entrepot_Num, 3)
        txtProductNo.Tag = pckStockProd.Product_ID
        txtProductNo.Text = pckStockProd.ProductNo
        
        With frmStockProdPicklist
            .fraStockCard.Height = .fraStockCard.Height - 40
            .Height = .fraStockCard.Height + (.cmdOK.Height * 2) + 480
            
            .Width = fraStockCard.Width + 360
            
            .fraStockCard.Top = 120
            .cmdOK.Top = .fraStockCard.Height + 240
            .cmdCancel.Top = .fraStockCard.Height + 240
            
            .cmdCancel.Left = .cmdNew.Left + 120
            .cmdOK.Left = .cmdCancel.Left - (.cmdOK.Width + 120)
        End With
        
        If jgxPicklist.RowCount > 0 Then
            Dim lngFind2 As Long
            
            jgxPicklist.ADORecordset.MoveFirst
            'Obtains row index for the grid.rowselected property.
            For lngFind2 = 1 To jgxPicklist.ADORecordset.RecordCount
                If jgxPicklist.ADORecordset.Fields("Stock Card No").Value = pckStockProd.StockCardNo Then Exit For
                jgxPicklist.ADORecordset.MoveNext
            Next lngFind2
            
            'Set row selected to row index matching Stock ID from property.
            jgxPicklist.MoveLast
            If lngFind2 <= jgxPicklist.ADORecordset.RecordCount Then
                jgxPicklist.Row = lngFind2
            Else
                jgxPicklist.Row = 1
            End If
            Call jgxPicklist_Click
        End If
    End If
    
    Me.Show vbModal
    Cancelled = blnCancelled
End Sub

Private Sub cmdCancel_Click()
    CleanUp True
    Unload Me
End Sub

Private Sub cmdCountry_Click(Index As Integer)
    Dim gsdCountry As PCubeLibPick.CGridSeed
    Dim strCountrySQL As String
    
    Set pckCountry = New CPicklist
    Set gsdCountry = New CGridSeed
    
    Set gsdCountry = pckCountry.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
    
    With pckStockProd
        ' The primary key is mentioned twice to conform to the design of the picklist class.
        strCountrySQL = "SELECT Code AS [Key Code], Code as [CODE], [Description " & IIf(UCase(.strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(.strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Key Description] " & _
                        "FROM [PICKLIST MAINTENANCE " & IIf(UCase(.strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(.strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] INNER JOIN [PICKLIST DEFINITION] ON " & _
                        "[PICKLIST MAINTENANCE " & IIf(UCase(.strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(.strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "].[INTERNAL CODE] = [PICKLIST DEFINITION].[INTERNAL CODE] " & _
                        "WHERE Document = 'Import' and [BOX CODE] = 'C2'"
    End With
    
    With pckCountry
        Select Case Index
            Case 0
                .Search True, "Key Code", Trim(strCountryOrig)
            Case 1
                .Search True, "Key Code", Trim(strCountryExp)
        End Select
        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick Me, cpiSimplePicklist, pckStockProd.m_conSADBEL, strCountrySQL, "Key Code", "Countries", vbModal, gsdCountry, , , True, cpiKeyF2
        
        If Not .SelectedRecord Is Nothing Then
            Select Case Index
                Case 0
                    strCountryOrig = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    strCountryOrigDesc = .SelectedRecord.RecordSource.Fields("Key Description").Value
                    txtCtryOrigin.Text = strCountryOrig
                    txtCtryOriginDesc.Text = strCountryOrigDesc
                Case 1
                    strCountryExp = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    strCountryExpDesc = .SelectedRecord.RecordSource.Fields("Key Description").Value
                    txtCtryExport.Text = strCountryExp
                    txtCtryExportDesc.Text = strCountryExpDesc
            End Select
        End If
    End With
    
    Set gsdCountry = Nothing
    Set pckCountry = Nothing
End Sub

Private Sub cmdDown_Click()
    If txtBatchNo.Visible And txtJobNo.Visible And fraProduct.Visible Then
        If cmdDown.Caption = "&More..." Then
            'Expand!
            fraProduct.Height = 2175
            
            cmdDown.Caption = "&Hide..."
            fraInfo.Visible = True
            
            txtTaricCode.Text = txtTaricCode.Text
            txtCtryExport.Text = strCountryExp
            txtCtryExportDesc.Text = strCountryExpDesc
            txtCtryOrigin.Text = strCountryOrig
            txtCtryOriginDesc.Text = strCountryOrigDesc
        Else
            'Hide!
            fraProduct.Height = 1575
            
            cmdDown.Caption = "&More..."
            fraInfo.Visible = False
        End If
        
        fraStockCard.Top = fraProduct.Height + txtBatchNo.Height + txtJobNo.Height + 360
        cmdOK.Top = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height + 480
        cmdCancel.Top = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height + 480
        
        frmStockProdPicklist.Height = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height
        frmStockProdPicklist.Height = frmStockProdPicklist.Height + (cmdOK.Height * 2) + 720
        
    ElseIf txtBatchNo.Visible = False And txtJobNo.Visible = False And fraProduct.Visible Then
        If cmdDown.Caption = "&More..." Then
            'Expand!
            fraProduct.Height = 2175
            
            cmdDown.Caption = "&Hide..."
            fraInfo.Visible = True
            
            txtTaricCode.Text = txtTaricCode.Text
            txtCtryExport.Text = strCountryExp
            txtCtryExportDesc.Text = strCountryExpDesc
            txtCtryOrigin.Text = strCountryOrig
            txtCtryOriginDesc.Text = strCountryOrigDesc
        Else
            'Hide!
            fraProduct.Height = 1575
            
            cmdDown.Caption = "&More..."
            fraInfo.Visible = False
        End If
        fraStockCard.Top = fraProduct.Height + 240
        cmdOK.Top = fraProduct.Height + fraStockCard.Height + 360
        cmdCancel.Top = fraProduct.Height + fraStockCard.Height + 360
    
        frmStockProdPicklist.Height = fraProduct.Height + fraStockCard.Height + (cmdOK.Height * 2) + 600
    End If
End Sub

Private Sub cmdNew_Click()
    Dim strStockID As String        'Storing Stock ID of newly created Stock Card.
    'Dim strStockCardNo As String
    Dim strStockCardNoHigh2 As String
    Dim blnNewStockcard As Boolean
    Dim lngLength As Long
    
    blnNewStockcard = False
    
    If Len(txtProductNo.Tag) = 0 Then
        'Just called on Product Number validation and/or prompt for new record addition.
        txtProductNo_LostFocus
        'Does not allow new addition if user cancels to avoid conflict.
        If bytProdNumDeclined = 0 Then Exit Sub
    End If

    Dim lngSafeLength As String
    Dim blnSafe As Boolean
    
    blnSafe = False
    
    'New SC# method.
    'This recordset is used in conjunction with m_rstPass2GridOff.
    'Previously was used and closed in PopGrid to get the <several>.
    'Now it is re-used to obtain the next safe Stock Card Number per Entrepot Number.
    'The IIF condition determines what to write in the Length field.
    'If the ceiling (9) of a certain length is met plus it does not contain leading zeros,
    'it will put LEN & "9" otherwise, it will only put the LEN under the Length field.
    ADORecordsetClose m_rstFindSeveral
    With m_rstFindSeveral
        
        ADORecordsetOpen "SELECT SC.Stock_ID AS [Stock ID], SC.Stock_Card_Num AS [Stock Card No], " & _
              "IIF(VAL(SC.Stock_Card_Num) = 9 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '19', " & "IIF(VAL(SC.Stock_Card_Num) = 99 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '29', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '39', " & "IIF(VAL(SC.Stock_Card_Num) = 9999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '49', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 99999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '59', " & "IIF(VAL(SC.Stock_Card_Num) = 999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '69', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 9999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '79', " & "IIF(VAL(SC.Stock_Card_Num) = 99999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '89', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 999999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '99', " & "LEN(SC.Stock_Card_Num))))))))))  AS [Length] " & _
              "FROM (Entrepots [E] INNER JOIN (StockCards [SC] INNER JOIN Products [P] " & "ON SC.Prod_ID = P.Prod_ID) ON E.Entrepot_ID = P.Entrepot_ID) " & _
              "WHERE E.Entrepot_ID = " & GetEntrepot_ID(pckStockProd.Entrepot_Num, pckStockProd.m_conSADBEL) & " " & "ORDER BY LEN(SC.Stock_Card_Num), SC.Stock_Card_Num", _
              pckStockProd.m_conSADBEL, m_rstFindSeveral, adOpenKeyset, adLockOptimistic
              
        '.Open "SELECT SC.Stock_ID AS [Stock ID], SC.Stock_Card_Num AS [Stock Card No], " & _
              "IIF(VAL(SC.Stock_Card_Num) = 9 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '19', " & "IIF(VAL(SC.Stock_Card_Num) = 99 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '29', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '39', " & "IIF(VAL(SC.Stock_Card_Num) = 9999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '49', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 99999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '59', " & "IIF(VAL(SC.Stock_Card_Num) = 999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '69', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 9999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '79', " & "IIF(VAL(SC.Stock_Card_Num) = 99999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '89', " & _
              "IIF(VAL(SC.Stock_Card_Num) = 999999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '99', " & "LEN(SC.Stock_Card_Num))))))))))  AS [Length] " & _
              "FROM (Entrepots [E] INNER JOIN (StockCards [SC] INNER JOIN Products [P] " & "ON SC.Prod_ID = P.Prod_ID) ON E.Entrepot_ID = P.Entrepot_ID) " & _
              "WHERE E.Entrepot_ID = " & GetEntrepot_ID(pckStockProd.Entrepot_Num, pckStockProd.m_conSADBEL) & " " & "ORDER BY LEN(SC.Stock_Card_Num), SC.Stock_Card_Num", _
              pckStockProd.m_conSADBEL, adOpenKeyset, adLockReadOnly

        'Begin safe length/value search with the length of configured Entrepot_Starting_Num.
        lngSafeLength = Len(strStartingNum)
        Do Until blnSafe = True
            'Retrieve involved records that have all digits as "9".
            If Not (.BOF And .EOF) Then .Filter = "[Length] = " & lngSafeLength & "9"
            
            If .BOF And .EOF Then
                'Flag safe when highest number for current Stock Card num length is not at ceiling.
                blnSafe = True
            Else
                'Otherwise, move to next Stock Card num length.
                lngSafeLength = lngSafeLength + 1
            End If
        Loop

''        If Not (m_rstPass2GridOff) Is Nothing Then .Filter = "[Length] = " & lngSafeLength --- 'Old SC# method.
        If Not (m_rstFindSeveral) Is Nothing Then .Filter = "[Length] = " & lngSafeLength
        If Not (.BOF And .EOF) Then
''            .Sort = "[Length], [Stock Card No] ASC"         --------------------------    'Old SC# method.
            If m_rstNewStockOff.RecordCount = 0 Then
'                .MoveLast
'                strStockCardNo = .Fields("Stock Card No").Value
                If UsedExistingStockcards(m_rstFindSeveral, strStartingNum) = True Then
                    strStockCardNoHigh = HighestStartingNumber(m_rstFindSeveral, strStartingNum)
                    blnNewStockcard = True
                Else
                    If UsedExistingStockcards(m_rstFindSeveral, strStartingNum, True) = True Then
                        strStockCardNoHigh = HighestStartingNumber(m_rstFindSeveral, strStartingNum, True)
                        blnNewStockcard = True
                    Else
                        strStockCardNoHigh = strStartingNum
                        blnNewStockcard = True
                    End If
                End If
            Else
                If UsedExistingStockcards(m_rstFindSeveral, strStartingNum) = True Then
                    strStockCardNoHigh = HighestStartingNumber(m_rstFindSeveral, strStartingNum)
                    blnNewStockcard = True
                Else
                    If UsedExistingStockcards(m_rstFindSeveral, strStartingNum, True) = True Then
                        strStockCardNoHigh = HighestStartingNumber(m_rstFindSeveral, strStartingNum, True)
                        blnNewStockcard = True
                    Else
                        strStockCardNoHigh = strStartingNum
                        blnNewStockcard = True
                    End If
                End If

                If UsedExistingStockcards(m_rstNewStockOff, strStartingNum) = True Then
                    strStockCardNoHigh2 = HighestStartingNumber(m_rstNewStockOff, strStartingNum)
                    blnNewStockcard = True
                Else
                    strStockCardNoHigh2 = strStartingNum
                    blnNewStockcard = True
                End If
                
                If Val(strStockCardNoHigh2) > Val(strStockCardNoHigh) Then
                    strStockCardNoHigh = strStockCardNoHigh2
                End If
            
                Do While UsedExistingStockcards(m_rstFindSeveral, strStockCardNoHigh, True) = True Or _
                    UsedExistingStockcards(m_rstNewStockOff, strStockCardNoHigh) = True
                    lngLength = Len(strStockCardNoHigh)
                    strStockCardNoHigh = strStockCardNoHigh + 1
                    
                    If Len(strStockCardNoHigh) < lngLength Then
                        strStockCardNoHigh = String$(lngLength - Len(strStockCardNoHigh), "0") & strStockCardNoHigh
                    End If
                Loop
                
            End If
            
            If Val(strStockCardNoHigh) >= Val(strStartingNum) Then
                If Not blnNewStockcard Then
                    lngLength = Len(strStockCardNoHigh)
                End If
            Else
                If lngSafeLength = Len(strStartingNum) Then
                    strStockCardNoHigh = strStartingNum
                ElseIf lngSafeLength >= 10 Then
                    'Default maximum length of Stock Card num has been reached.
                    strStockCardNoHigh = Empty
                Else
                    strStockCardNoHigh = "1" & String$(lngSafeLength - 1, "0")
                End If
            End If
        Else
            If m_rstNewStockOff.RecordCount > 0 Then
                If UsedExistingStockcards(m_rstNewStockOff, strStartingNum) = True Then
                    strStockCardNoHigh = HighestStartingNumber(m_rstNewStockOff, strStartingNum)
                Else
                    If lngSafeLength = Len(strStartingNum) Then
                        strStockCardNoHigh = strStartingNum
                    ElseIf lngSafeLength >= 10 Then
                        'Default maximum length of Stock Card num has been reached.
                        strStockCardNoHigh = Empty
                    Else
                        strStockCardNoHigh = "1" & String$(lngSafeLength - 1, "0")
                    End If
                End If
                
                Do While UsedExistingStockcards(m_rstFindSeveral, strStockCardNoHigh, True) = True Or _
                    UsedExistingStockcards(m_rstNewStockOff, strStockCardNoHigh) = True
                    lngLength = Len(strStockCardNoHigh)
                    strStockCardNoHigh = strStockCardNoHigh + 1
                    
                    If Len(strStockCardNoHigh) < lngLength Then
                        strStockCardNoHigh = String$(lngLength - Len(strStockCardNoHigh), "0") & strStockCardNoHigh
                    End If
                Loop
            Else
                If lngSafeLength = Len(strStartingNum) Then
                    strStockCardNoHigh = strStartingNum
                ElseIf lngSafeLength >= 10 Then
                    'Default maximum length of Stock Card num has been reached.
                    strStockCardNoHigh = Empty
                Else
                    strStockCardNoHigh = "1" & String$(lngSafeLength - 1, "0")
                End If
            End If
        End If
        
        If Not (.BOF Or .EOF) Then .Move jgxPicklist.Row, adBookmarkFirst
        
        frmStockcard.Pre_Load2 lngEntrepotID, strEntrepotType, strEntrepotNum, txtProductNo.Tag, txtProductNo.Text, bytNumbering, strStartingNum, strStockCardNoHigh, pckStockProd, ResourceHandler
        
        'Pass Stock ID to variable for use in .Find.
        If blnCancelled = False Then strStockID = m_rstPass2GridOff.Fields(0).Value
        
        m_rstPass2GridOff.Filter = ""
        'Only performs the commit and grid update when ok was clicked.
        If blnCancelled = False Then
            'Refresh display of grid.
            Set jgxPicklist.ADORecordset = m_rstPass2GridOff
            HideSomeFields
            
            'Find and move pointer to newly created Stock Card.
            If Len(strStockID) > 0 Then
                jgxPicklist.ADORecordset.Find "[Stock ID] = " & strStockID

                If jgxPicklist.ADORecordset.EOF = False Then jgxPicklist.MoveToBookmark (jgxPicklist.ADORecordset.Bookmark)
            End If
                        
            'Automatically performs the check on selected grid row to enable/disable Select button.
            jgxPicklist_Click
        End If
    End With
    
    ADORecordsetClose m_rstFindSeveral
End Sub

Private Sub cmdOK_Click()
    If Not blnIsInitialStock Then
        If Not blnWithEntrepotNum Then
            CtryCodeMod
            If Validation = False Then Exit Sub
        End If
    End If
    
    If Not blnWithEntrepotNum Then
        Pass2Class
        SaveStockCards
    Else
        Pass2Class2
    End If
    
    CleanUp False
    Unload Me
End Sub

Private Sub CleanUp(Cancel As Boolean)
    If blnRstIsNothing = False Then
        ADORecordsetClose m_rstPass2GridOff
    End If
    
    Set pckProducts = Nothing
    Set pckStockProd = Nothing
    
    ADORecordsetClose m_rstFindSeveral
    ADORecordsetClose m_rstPass2GridOff
    
    blnCancelled = Cancel
    
    ADORecordsetClose m_rstNewStockOff
End Sub

Private Function Validation() As Boolean
    Dim bytProblemsFlag As Byte
    Dim strProblems As String
    bytProblemsFlag = 0
    Validation = True
    
    'Only checks if called from codisheet.
    If blnLesserForm = False Then
        If Len(Trim(txtJobNo.Text)) = 0 Then
            Validation = False
            strProblems = strProblems & Space(5) & "* Job Number - Missing" & vbCrLf
            bytProblemsFlag = bytProblemsFlag + 1
        End If
        If Len(Trim(txtBatchNo.Text)) = 0 Then
            Validation = False
            strProblems = strProblems & Space(5) & "* Batch Number - Missing" & vbCrLf
            bytProblemsFlag = bytProblemsFlag + 2
        End If
    End If
    
    If Len(Trim(txtProductNo.Text)) = 0 Then
        Validation = False
        strProblems = strProblems & Space(5) & "* Product Number - Missing" & vbCrLf
        bytProblemsFlag = bytProblemsFlag + 4
    End If
    
    If Len(Trim(strCountryOrig)) = 0 Then
        Validation = False
        If Len(Trim(strCountryOrigDesc)) = 0 Then
            strProblems = strProblems & Space(5) & "* Country of Origin - Missing" & vbCrLf
        Else
            strProblems = strProblems & Space(5) & "* Country of Origin - Not In Database" & vbCrLf
        End If
        bytProblemsFlag = bytProblemsFlag + 8
    Else
        'Checks if entered Country Code corresponds with Product Number in products table.
        If Len(txtCtryOrigin.Tag) > 0 Then
            If Not (strCountryOrig = txtCtryOrigin.Tag) Then
                Validation = False
                strProblems = strProblems & Space(5) & "* Country of Origin - Inconsistent With Database" & vbCrLf
                bytProblemsFlag = bytProblemsFlag + 8
            End If
        End If
    End If
    
    If Len(Trim(strCountryExp)) = 0 Then
        Validation = False
        If Len(Trim(strCountryExpDesc)) = 0 Then
            strProblems = strProblems & Space(5) & "* Country of Export - Missing" & vbCrLf
        Else
            strProblems = strProblems & Space(5) & "* Country of Export - Not In Database" & vbCrLf
        End If
        bytProblemsFlag = bytProblemsFlag + 16
    Else
        'Checks if entered Country Code corresponds with Product Number in products table.
        If Len(txtCtryExport.Tag) > 0 Then
            If Not (strCountryExp = txtCtryExport.Tag) Then
                Validation = False
                strProblems = strProblems & Space(5) & "* Country of Export - Inconsistent With Database" & vbCrLf
                bytProblemsFlag = bytProblemsFlag + 16
            End If
        End If
    End If
    
    If Validation = False Then
        MsgBox Translate(2277) & vbCrLf & strProblems & vbCrLf, vbOKOnly + vbInformation, Translate(2278)
    End If
    
    'Checks flags and sets focus to top-most affected control in the list.
    Select Case bytProblemsFlag
        Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47, 49, 51, 53, 55, 57, 59, 61, 63
            txtJobNo.SetFocus
        Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62
            txtBatchNo.SetFocus
        Case 4, 12, 20, 28, 36, 44, 52, 60
            txtProductNo.SetFocus
        Case 8, 24, 40
            If fraInfo.Visible = True Then
                txtCtryOrigin.SetFocus
            Else
                Call cmdDown_Click
                txtCtryOrigin.SetFocus
            End If
        Case 16, 48
            If fraInfo.Visible = True Then
                txtCtryExport.SetFocus
            Else
                Call cmdDown_Click
                txtCtryExport.SetFocus
            End If
    End Select
End Function

Private Sub Pass2Class()
    With pckStockProd
        .CtryOrigin = strCountryOrig
        .CtryExport = strCountryExp
        .BatchNo = Trim$(txtBatchNo.Text)
        .JobNo = Trim$(txtJobNo.Text)
        .Product_ID = Trim$(txtProductNo.Tag)
        .ProductNo = Trim$(txtProductNo.Text)
        .TaricCode = Trim$(txtTaricCode.Text)
        .Stock_ID = jgxPicklist.Value(jgxPicklist.Columns("Stock ID").Index)
        .StockCardNo = jgxPicklist.Value(jgxPicklist.Columns("Stock Card No").Index)
        .ProductDesc = txtProductDesc.Text
    End With
End Sub

Private Sub cmdProductNo_Click()
    'Pass Taric Code to Stock Prod property for use in filtering of Products picklist.
    With pckProducts
        'If Taric Code is not empty.
        If Len(Trim(txtTaricCode.Text)) > 0 Then
            'If Taric Code value = zero.
            If Val(txtTaricCode.Text) = 0 Then
                'Perform a search to see if a Product matches Taric Code = 0 and Entrepot Num of BB.
                Dim strWhere As String
                Dim lngEntrepot As Long
                Dim rstTaricZero As ADODB.Recordset

                lngEntrepot = GetEntrepot_ID(pckStockProd.Entrepot_Num, pckStockProd.m_conSADBEL)
                
                strWhere = "WHERE "
                strWhere = strWhere & "Taric_Code = '0'"
                If lngEntrepot <> 0 Then strWhere = strWhere & " AND Entrepot_ID = " & lngEntrepot
                
                ADORecordsetOpen "SELECT Taric_Code, Entrepot_ID FROM Products " & _
                                  strWhere, pckStockProd.m_conSADBEL, rstTaricZero, adOpenKeyset, adLockOptimistic
                'rstTaricZero.Open "SELECT Taric_Code, Entrepot_ID FROM Products " & _
                                  strWhere, pckStockProd.m_conSADBEL, adOpenKeyset, adLockReadOnly
                'If match found, copy Taric Code to property for filter application.
                If Not (rstTaricZero.BOF And rstTaricZero.EOF) Then
                    .Taric_Code = txtTaricCode.Text
                'Otherwise, set to empty to disable filter.
                Else
                    .Taric_Code = Empty
                End If
                'Clean up.
                strWhere = Empty
                lngEntrepot = Empty
                
                ADORecordsetClose rstTaricZero

            'If Taric Code value > zero.
            Else
                'Passes value anyway (line below was original code).
                .Taric_Code = txtTaricCode.Text
            End If
        'If Taric Code is empty.
        Else
        End If
    
        If blnLesserForm = False Then
            'Passes Country Code values to class if value has length from 2 to 3 (max length).
            If Not (Len(Trim(strCountryOrig)) < 2) Then
                .Ctry_Origin = strCountryOrig
            Else
                .Ctry_Origin = Empty
            End If
            If Not (Len(Trim(strCountryExp)) < 2) Then
                .Ctry_Export = strCountryExp
            Else
                .Ctry_Export = Empty
            End If
        End If
    End With
    
    'Flag to close m_rstPass2GridOff on unload.
    blnRstIsNothing = False
    With pckStockProd
        pckProducts.ShowProducts 3, Me, .m_conSADBEL, .m_conTaric, _
                                    .strLanguage, .intTaricProperties, ResourceHandler
        
        'Prevents updating of Stock/Prod picklist when Product selection has been cancelled.
        If pckProducts.Cancelled = False Then
        
            strCountryOrig = pckProducts.Ctry_Origin
            txtCtryOrigin.Tag = pckProducts.Ctry_Origin
            strCountryExp = pckProducts.Ctry_Export
            txtCtryExport.Tag = pckProducts.Ctry_Export
            strCountryOrigDesc = pckProducts.Origin_Desc
            strCountryExpDesc = pckProducts.Export_Desc
            txtProductNo.Text = pckProducts.Product_Num
            txtProductNo.Tag = pckProducts.Product_ID
            txtProductDesc.Text = pckProducts.Prod_Desc
            txtTaricCode.Text = pckProducts.Taric_Code
            strTaricCode = pckProducts.Taric_Code
            If fraInfo.Visible = True Then
                txtCtryExport.Text = strCountryExp
                txtCtryExportDesc.Text = strCountryExpDesc
                txtCtryOrigin.Text = strCountryOrig
                txtCtryOriginDesc.Text = strCountryOrigDesc
            End If
            
            lngEntrepotID = pckProducts.Entrepot_ID
            
            strEntrepotType = pckProducts.Entrepot_Type
            strEntrepotNum = pckProducts.Entrepot_Num_Only
            bytNumbering = pckProducts.Numbering
            strStartingNum = pckProducts.StartingNum

            PopGrid Val(txtProductNo.Tag), .m_conSADBEL
            
            m_rstNewStockOff.Close
            Call InitializeRecordset
        End If
    End With
    
    'Simple checking to see if a product has been selected.
    'If so, enable the New button for creating stock cards based on selected product.
    If Len(txtProductNo.Tag) > 0 Then
        cmdNew.Enabled = True
    Else
        cmdNew.Enabled = False
    End If
    'Calls grid click to perform a check and enable OK if there's a selected item.
    jgxPicklist_Click
End Sub

Private Sub PopGrid(Prod_ID As Long, conn_Sadbel As ADODB.Connection)
    Dim strFirst(0 To 6) As String
    Dim strSecond(0 To 4) As String
    Dim blnSeveral(2 To 4) As Boolean               '0 not included since the "Stock Card No" field can't be <Several>.
    Dim lngCtr As Long
    Dim strIM7 As String
    Dim strStockCardSQL As String
    Dim strSCNumX As String

    'Temporary fix to avoid getting query error when WHERE condition has missing value.
    'Better than displaying stock cards that are not really associated with a product missing a Prod_ID.
    If Len(Trim(Prod_ID)) = 0 Then Prod_ID = 1999999999           'The 1999.. value can probably be replaced with a 0 instead.
    
    strStockCardSQL = "SELECT SC.Stock_ID AS [ID], SC.Stock_Card_Num AS [Stock Card No], " & _
                      "I.In_Job_Num AS [Job No], I.In_Batch_Num AS [Batch No], " & _
                      "ID.InDoc_Type AS [Doc Type], ID.InDoc_Num AS [Doc Num], " & _
                      "P.Prod_ID AS [Product ID], P.Prod_Num AS [Prod Num], SC.Prod_ID AS [Prod ID], " & _
                      "E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                      "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StockCard_Numbering AS [Numbering], " & _
                      "E.Entrepot_Starting_Num AS [Starting Num], " & _
                      "P.Prod_Ctry_Origin AS [Country Origin], P.Prod_Ctry_Export AS [Country Export] " & _
                      "FROM (Entrepots [E] INNER JOIN (StockCards [SC] INNER JOIN Products [P] " & _
                      "ON SC.Prod_ID = P.Prod_ID) ON E.Entrepot_ID = P.Entrepot_ID) " & _
                      "LEFT JOIN " & _
                      "(Inbounds [I] LEFT JOIN InboundDocs [ID] on I.Indoc_ID = ID.Indoc_ID) " & _
                      "ON SC.Stock_ID = I.Stock_ID " & _
                      "WHERE SC.Prod_ID = " & Trim(Prod_ID) & " " & _
                      "ORDER BY P.Prod_ID, SC.Stock_ID, I.In_Job_Num, SC.Stock_Card_Num"

    
    'Query for m_rstFindSeveral.
    ADORecordsetOpen strStockCardSQL, conn_Sadbel, m_rstFindSeveral, adOpenKeyset, adLockOptimistic
    'm_rstFindSeveral.Open strStockCardSQL, conn_Sadbel, adOpenKeyset, adLockOptimistic
    
    'Prepare second recordset for manual population.
    With m_rstPass2GridOff
        .CursorLocation = adUseClient
        .Fields.Append "Stock ID", adVarNumeric, 10
        .Fields.Append "Stock Card No", adVarChar, 10
        .Fields.Append "Job No", adVarChar, 50
        .Fields.Append "Batch No", adVarChar, 50
        .Fields.Append "Doc Number", adVarChar, 100
        .Fields.Append "Entrepot ID", adVarChar, 10
        .Fields.Append "Product ID", adVarChar, 10
        .Fields.Append "New", adBoolean
        .Fields.Append "Length", adVarChar, 10
        
        .Open
    End With
    '---------------- Locating those severals ----------------
    With m_rstFindSeveral
        If .RecordCount > 0 Then .MoveFirst
        'Process primary recordset until last record or Stock Card No.
        Do Until .EOF = True
            '==========================================================================
            '===== m_rstPass2GridOff Fields    =====       ===== m_rstFindSeveral Fields =====
            '===== (strFirst & strSecond) =====       =====                       =====
            '=====  0 = Stock ID          =====       =====  0 = ID               =====
            '=====  1 = Stock Card No     =====       =====  1 = Stock Card No    =====
            '=====  2 = Job No            =====       =====  2 = Job No           =====
            '=====  3 = Batch No          =====       =====  3 = Batch No         =====
            '=====  4 = IM7               =====       =====  4 = Doc Type         =====
            '=====  5 = Entrepot ID       =====       =====  5 = Doc Num          =====
            '=====  6 = Product ID        =====       =====  6 = Product ID       =====
            '=====  7 = New               =====       =====  7 = Prod Num         =====
            '=====  8 = Length            =====       =====  8 = Entrepot ID      =====
            '=====                        =====       =====  9 = Entrepot Type    =====
            '=====                        =====       ===== 10 = Entrepot Num     =====
            '=====                        =====       ===== 11 = Numbering        =====
            '=====                        =====       ===== 12 = Starting Num     =====
            '=====                        =====       ===== 13 = Country Origin   =====
            '=====                        =====       ===== 14 = Country Export   =====
            '==========================================================================
            
            'Stores first unique instance of:
            'Stock ID | Stock Card No | Job No | Batch No
            For lngCtr = 0 To 6
                Select Case lngCtr
                    Case Is < 4
                        If Not IsNull(.Fields(lngCtr).Value) Then
                            strFirst(lngCtr) = .Fields(lngCtr).Value
                        Else
                            'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                            strFirst(lngCtr) = Empty
                        End If
                    Case 4
                        If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                            'Sets format of IM7 to "Type"-"Num".
                            strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                        Else
                            strIM7 = Empty
                        End If
                        strFirst(lngCtr) = strIM7
                    Case 5
                        strFirst(lngCtr) = .Fields("Entrepot ID").Value
                    Case 6
                        strFirst(lngCtr) = .Fields("Product ID").Value
                End Select
            Next lngCtr
            .MoveNext
            
            'If there's more than 1 record in the rst.
            If .RecordCount > 1 And Not .EOF Then
                'Stores second instance (possibly unique).
                'Stock ID | Stock Card No | Job No | Batch No
                For lngCtr = 0 To 4
                    If Not IsNull(.Fields(lngCtr).Value) Then
                        If Not lngCtr = 4 Then
                            strSecond(lngCtr) = .Fields(lngCtr).Value
                        Else
                            If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                                'Sets format of IM7 to "Type"-"Num".
                                strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                            End If
                            strSecond(lngCtr) = strIM7
                        End If
                    Else
                        'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                        strSecond(lngCtr) = Empty
                    End If
                Next lngCtr
            'In case there's only 1 record in the rst or an EOF.
            Else
                'Stores dummy second instance (definitely unique).
                For lngCtr = 0 To 4
                    strSecond(lngCtr) = "dummy"
                Next lngCtr
            End If
            
            'Check if first and second are alike.
            'Perform record transfer when distinct.
            If strFirst(1) <> strSecond(1) Then
                'Copy first instance to a new recordset.
                'This is the same action perform after processing multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 6
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    strSCNumX = Empty
                End With
            'When First = Second.
            ElseIf strFirst(1) = strSecond(1) Then
                'Initialize/reset "Several" flag.
                For lngCtr = 2 To 4
                    blnSeveral(lngCtr) = False
                Next lngCtr
                    
                'Proceed with locating severals for every alike Stock Card Num.
                Do While strFirst(1) = strSecond(1)
                    'Cycles through other fields to identify which records will be labelled "Several".
                    For lngCtr = 2 To 4
                        'Used boolean flags to avoid too much string comparisons.
                        If blnSeveral(lngCtr) = False Then
                            'Copies next record to strSecond for comparison with strFirst.
                            If Not IsNull(.Fields(lngCtr).Value) Then           'CUSTOM
                                If Not lngCtr = 4 Then
                                    strSecond(lngCtr) = .Fields(lngCtr).Value
                                Else
                                    If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                                        'Sets format of IM7 to "Type"-"Num".
                                        strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                                    End If
                                    strSecond(lngCtr) = strIM7
                                End If
                            Else
                                'MsgBox "Null value encountered in " & lngCtr & "."
                                strSecond(lngCtr) = Empty
                            End If
                                                                
                            'When strFirst != strSecond, a flag is raised to avoid processing this field.
                            If strFirst(lngCtr) <> strSecond(lngCtr) Then
                                blnSeveral(lngCtr) = True
                                'Stores the value "Several" for recording in first recordset.
                                strFirst(lngCtr) = "<Several>"
                                strSecond(lngCtr) = Empty
                            End If
                        End If
                    Next lngCtr

                    'Goes to next record or exits if last record of last distinct set.
                    .MoveNext
                    If .EOF = True Then Exit Do
                    strSecond(1) = .Fields(1).Value
                    
                    'Just adds the Entrepot ID and Product ID to the recordset.
                    strFirst(5) = .Fields("Entrepot ID").Value
                    strFirst(6) = .Fields("Product ID").Value
                Loop
                                
                'Copy first instance to a new recordset. After multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 6
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    strSCNumX = Empty
                End With
            End If
        Loop
    End With
    m_rstFindSeveral.Close
    'Commented to prevent new entries from being moved to top of grid.
'    m_rstPass2GridOff.Sort = "[Stock ID], [Stock Card No]"
    m_rstPass2GridOff.Sort = "[Length], [Stock Card No]"
    '------------ Finished locating those severals ------------
    Set jgxPicklist.ADORecordset = m_rstPass2GridOff
    HideSomeFields
End Sub

Private Sub Form_Activate()
    ControlResizing
    If txtJobNo.Visible = True Then
        If Len(txtJobNo.Text) > 0 Then
            txtJobNo.ForeColor = vbRed
        End If
    End If
    
    If txtBatchNo.Visible = True Then
        If Len(txtBatchNo.Text) > 0 Then
            txtBatchNo.ForeColor = vbRed
        End If
    End If
End Sub

Private Sub Form_Load()
    'Initialize flags for validating Product Number entered with database.
    bytProdFound = 1
    bytProdNumDeclined = 0
    blnCancelled = True
    
    InitializeRecordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'These values are for jgxPicklist sorting.
    strColumnName = ""
    lngSortCounter = 0
End Sub

Private Sub jgxPicklist_Click()
    'Enables/Disables Select button when a grid row is selected.
    'No grid rows are selected if there are no rows in the grid.
    If jgxPicklist.Row = 0 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub HideSomeFields()
    'Hides the Stock ID and New fields.
    With jgxPicklist
        .Columns(1).Visible = False
        .Columns(6).Visible = False
        .Columns(7).Visible = False
        .Columns(8).Visible = False
        .Columns(9).Visible = False
    End With
End Sub

Private Sub ParseMemo(Memo As String, MainDelim As String, SubDelim As String)
    'Filter memo field to obtain L1 inserted data.
    Dim lngBlah As Long
    Dim strMemo As String
    If Len(Memo) <> 0 Then
        'Copy value of memo to another var for manipulation.
        strMemo = Memo
        'Point tracer to first entry delimiter "JOB NO.:"
        lngBlah = InStr(1, Memo, "JOB")
        'If not found, then set tracer to 1 (since a value of 0 or -1 will result in error).
        If lngBlah > 0 Then
            strMemo = Mid$(Memo, lngBlah)
            'Point tracer to the value after delimiter.
            lngBlah = InStr(1, strMemo, SubDelim) - (InStr(1, strMemo, MainDelim) + 1)
            'If not found, then set tracer to 0. This will cause result to be empty.
            If lngBlah < 0 Then lngBlah = 0
            'Get Job Number.
            pckStockProd.JobNo = Trim$(Mid$(strMemo, InStr(1, strMemo, MainDelim) + 1, lngBlah))
        End If
        
        'Point tracer to entry delimiter "BATCH NO.:"
        lngBlah = InStr(1, strMemo, "BATCH")
        'If not found, then set tracer to 1 (since a value of 0 or -1 will result in error).
        If lngBlah > 0 Then
            strMemo = Mid$(strMemo, lngBlah)
            'Point tracer to the value after delimiter.
            lngBlah = InStr(1, strMemo, SubDelim) - (InStr(1, strMemo, MainDelim) + 1)
            'If not found, then set tracer to 0. This will cause result to be empty.
            If lngBlah < 0 Then lngBlah = 0
            'Get Batch Number.
            pckStockProd.BatchNo = Trim$(Mid$(strMemo, InStr(1, strMemo, MainDelim) + 1, lngBlah))
        End If
        
        'Point tracer to entry delimiter "PRODUCT NO.:"
        lngBlah = InStr(1, strMemo, "PROD")
        'If found, proceed with extraction to trace delimiter value.
        If lngBlah > 0 Then
            strMemo = Mid$(strMemo, lngBlah)
            'Point tracer to the value after delimiter.
            lngBlah = InStr(1, strMemo, SubDelim) - (InStr(1, strMemo, MainDelim) + 1)
            'If found, proceed with extraction for use with filter.
            If lngBlah < 0 Then lngBlah = 0
                'Get Product Number.
            pckStockProd.ProductNo = Trim$(Mid$(strMemo, InStr(1, strMemo, MainDelim) + 1, lngBlah))
        End If
        
        'Check if Product Number is in memo field, since having Stock Card is pointless without it.
        If Len(pckStockProd.ProductNo) > 0 Then
            'Point tracer to entry delimiter "STOCK NO.:"
            lngBlah = InStr(1, strMemo, "STOCK")
            'If found, proceed with extraction to trace delimiter value.
            If lngBlah > 0 Then
                strMemo = Mid$(strMemo, lngBlah)
                'Point tracer to the value after delimiter.
                If InStr(1, strMemo, SubDelim) > 0 Then
                    'If a vbCrLf was used for additional stuff after Stock Number.
                    lngBlah = InStr(1, strMemo, SubDelim) - (InStr(1, strMemo, MainDelim) + 1)
                Else
                    'If no vbCrLf was used for additional stuff after Stock Number.
                    lngBlah = InStr(1, strMemo, MainDelim) + 1
                End If
                'If found, proceed with extraction for use with filter.
                If lngBlah < 0 Then lngBlah = 0
                'Get Stock Number.
                pckStockProd.StockCardNo = Trim$(Mid$(strMemo, InStr(1, strMemo, MainDelim) + 1, lngBlah))
                'Only passes a valid number to Stock Number grid pointer.
                If IsNumeric(Val(pckStockProd.StockCardNo)) = True Then
                    pckStockProd.StockCardNo = pckStockProd.StockCardNo
                Else
                    pckStockProd.StockCardNo = Empty
                End If
            End If
        End If
        
        '---------
        
        'Uses the product num to find the product id and the other related fields.
        AutoLoadByProd pckStockProd.ProductNo
        
        'Will only perform the grid display if the product id is in the tag property.
        If Len(txtProductNo.Tag) <> 0 Then PopGrid Val(txtProductNo.Tag), pckStockProd.m_conSADBEL
        'Automatically performs the check on selected grid row to enable/disable Select button.
        jgxPicklist_Click
    Else
        'Allows passing of correct Entrepot Number for use with Stock Card maintainance form.
        strEntrepotType = UCase(Mid$(pckStockProd.Entrepot_Num, 1, 1))
        strEntrepotNum = Mid$(pckStockProd.Entrepot_Num, 3)
    End If
End Sub

Private Sub AutoLoadByProd(ProductNo As String)
    'Passes values to appropriate boxes based on Stock Prod properties' value passed from codisheet.
    Dim rstAutoLoad As ADODB.Recordset
    Dim strConnectionString As String
    
    txtProductNo.Text = ProductNo
    strEntrepotType = UCase(Mid$(pckStockProd.Entrepot_Num, 1, 1))
    strEntrepotNum = Mid$(pckStockProd.Entrepot_Num, 3)

    strConnectionString = pckStockProd.m_conSADBEL.ConnectionString
    pckStockProd.m_conSADBEL.Close
    pckStockProd.m_conSADBEL.Open strConnectionString
    
    With rstAutoLoad
        'Uses Right Join so Entrepot settings can be extracted even without any Products.
        ADORecordsetOpen "SELECT P.Prod_ID AS [Prod_ID], P.Prod_Handling as [Prod_Handling], P.Prod_Num AS [Product No], " & _
              "P.Taric_Code AS [Taric Code], P.Entrepot_ID AS [Entrepot ID], P.Prod_Desc AS [Product Desc], " & _
              "P.Prod_Ctry_Origin AS [Country Origin], P.Prod_Ctry_Export AS [Country Export], " & _
              "E.Entrepot_Type AS [Entrepot Type], E.Entrepot_Num AS [Entrepot Num], " & _
              "E.Entrepot_StockCard_Numbering AS [Numbering], E.Entrepot_Starting_Num AS [Starting Num] " & _
              "FROM Products [P] RIGHT JOIN Entrepots [E] " & _
              "ON P.Entrepot_ID = E.Entrepot_ID " & _
              "WHERE E.Entrepot_Type = '" & strEntrepotType & "' " & _
              "AND E.Entrepot_Num = '" & strEntrepotNum & _
              "' AND P.Prod_Num = '" & ProductNo & "'", _
              pckStockProd.m_conSADBEL, rstAutoLoad, adOpenKeyset, adLockOptimistic
              
        '.Open "SELECT P.Prod_ID AS [Prod_ID], P.Prod_Handling as [Prod_Handling], P.Prod_Num AS [Product No], " & _
              "P.Taric_Code AS [Taric Code], P.Entrepot_ID AS [Entrepot ID], P.Prod_Desc AS [Product Desc], " & _
              "P.Prod_Ctry_Origin AS [Country Origin], P.Prod_Ctry_Export AS [Country Export], " & _
              "E.Entrepot_Type AS [Entrepot Type], E.Entrepot_Num AS [Entrepot Num], " & _
              "E.Entrepot_StockCard_Numbering AS [Numbering], E.Entrepot_Starting_Num AS [Starting Num] " & _
              "FROM Products [P] RIGHT JOIN Entrepots [E] " & _
              "ON P.Entrepot_ID = E.Entrepot_ID " & _
              "WHERE E.Entrepot_Type = '" & strEntrepotType & "' " & _
              "AND E.Entrepot_Num = '" & strEntrepotNum & _
              "' AND P.Prod_Num = '" & ProductNo & "'", _
              pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
        
            'To load selected Product related settings.
            If Not (.BOF And .EOF) Then
                .MoveFirst
                
                pckStockProd.ProductHandling = .Fields("Prod_Handling").Value

                bytNumbering = .Fields("Numbering").Value
                strStartingNum = .Fields("Starting Num").Value
                
                strCountryOrig = .Fields("Country Origin").Value
                strCountryExp = .Fields("Country Export").Value
                txtCtryOrigin.Tag = .Fields("Country Origin").Value
                txtCtryExport.Tag = .Fields("Country Export").Value
            
                strBlah = GetCountryDesc(strCountryOrig, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryOrigDesc = strBlah
                strBlah = GetCountryDesc(strCountryExp, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryExpDesc = strBlah
    
                txtProductNo.Tag = .Fields("Prod_ID").Value
                txtProductDesc.Text = .Fields("Product Desc").Value
                'Mod to load Taric Code from codisheet instead from Products table.
                txtTaricCode.Text = pckStockProd.TaricCode
                strTaricCode = pckStockProd.TaricCode
''                txtTaricCode.Text = .Fields("Taric Code").Value
''                strTaricCode = .Fields("Taric Code").Value

                lngEntrepotID = .Fields("Entrepot ID").Value
            Else
                'Prevents loading error with inconsistent Memo field and Entrepot Number (BB).
                'Scenario: User selects valid Stock Card, changes BB value and open Stock/Prod again.
                pckStockProd.StockCardNo = Empty
            End If
'        End If
        
        .Close
    End With
    
    ADORecordsetClose rstAutoLoad
End Sub

Private Sub AutoLoadByProdID(ProductID As Long)
    'Passes values to appropriate boxes based on Product ID passed from Summary Report.
    Dim rstAutoLoad As ADODB.Recordset
    
    txtProductNo.Tag = ProductID

    ADORecordsetOpen "SELECT P.Prod_ID AS [Prod_ID], P.Prod_Num AS [Product No], P.Prod_Handling as [Prod_Handling],  " & _
              "P.Taric_Code AS [Taric Code], P.Entrepot_ID AS [Entrepot ID], P.Prod_Desc AS [Product Desc], " & _
              "P.Prod_Ctry_Origin AS [Country Origin], P.Prod_Ctry_Export AS [Country Export], " & _
              "E.Entrepot_Type AS [Entrepot Type], E.Entrepot_Num AS [Entrepot Num], " & _
              "E.Entrepot_StockCard_Numbering AS [Numbering], E.Entrepot_Starting_Num AS [Starting Num] " & _
              "FROM Products [P] INNER JOIN Entrepots [E] " & _
              "ON P.Entrepot_ID = E.Entrepot_ID " & _
              "WHERE P.Prod_ID = " & ProductID, _
              pckStockProd.m_conSADBEL, rstAutoLoad, adOpenKeyset, adLockOptimistic
              
    With rstAutoLoad
        '.Open "SELECT P.Prod_ID AS [Prod_ID], P.Prod_Num AS [Product No], P.Prod_Handling as [Prod_Handling],  " & _
              "P.Taric_Code AS [Taric Code], P.Entrepot_ID AS [Entrepot ID], P.Prod_Desc AS [Product Desc], " & _
              "P.Prod_Ctry_Origin AS [Country Origin], P.Prod_Ctry_Export AS [Country Export], " & _
              "E.Entrepot_Type AS [Entrepot Type], E.Entrepot_Num AS [Entrepot Num], " & _
              "E.Entrepot_StockCard_Numbering AS [Numbering], E.Entrepot_Starting_Num AS [Starting Num] " & _
              "FROM Products [P] INNER JOIN Entrepots [E] " & _
              "ON P.Entrepot_ID = E.Entrepot_ID " & _
              "WHERE P.Prod_ID = " & ProductID, _
              pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
  
        If Not (.BOF And .EOF) Then
            .MoveFirst
            
            strCountryOrig = .Fields("Country Origin").Value
            strCountryExp = .Fields("Country Export").Value
            txtCtryOrigin.Tag = .Fields("Country Origin").Value
            txtCtryExport.Tag = .Fields("Country Export").Value

            strBlah = GetCountryDesc(strCountryOrig, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
            If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryOrigDesc = strBlah
            strBlah = GetCountryDesc(strCountryExp, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
            If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryExpDesc = strBlah
            
            txtProductNo.Text = .Fields("Product No").Value
            txtProductDesc.Text = .Fields("Product Desc").Value
            txtTaricCode.Text = .Fields("Taric Code").Value
            strTaricCode = .Fields("Taric Code").Value
            lngEntrepotID = .Fields("Entrepot ID").Value
            bytNumbering = .Fields("Numbering").Value
            strStartingNum = .Fields("Starting Num").Value
            
            pckStockProd.ProductHandling = .Fields("Prod_Handling").Value
        End If
    End With
    
    ADORecordsetClose rstAutoLoad

    'Will only perform the grid display if the product id is in the tag property.
    If Len(txtProductNo.Tag) <> 0 Then PopGrid Val(txtProductNo.Tag), pckStockProd.m_conSADBEL
    'Automatically performs the check on selected grid row to enable/disable Select button.
    jgxPicklist_Click
End Sub

Private Sub HideBatchJob()
    cmdCountry(0).Enabled = False
    cmdCountry(1).Enabled = False
    txtCtryOrigin.Enabled = False
    txtCtryExport.Enabled = False
    txtCtryOriginDesc.Enabled = False
    txtCtryExportDesc.Enabled = False
    txtProductDesc.Enabled = False
    
    txtBatchNo.Visible = False
    txtJobNo.Visible = False
    lblBatchNo.Visible = False
    lblJobNo.Visible = False
    cmdNew.Visible = False
    
    jgxPicklist.Left = 700
End Sub

Private Sub jgxPicklist_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    
    'Assign the sort type per column.
'    If blnIsInitialStock = False And blnWithEntrepotNum = True Then
            
    If ColumnExists("Entrepot Number") Then
        jgxPicklist.Columns("Entrepot Number").SortType = jgexSortTypeString
    End If
    
    If ColumnExists("Product Number") Then
        jgxPicklist.Columns("Product Number").SortType = jgexSortTypeString
    End If
            
    If ColumnExists("Stock Card No") Then
        jgxPicklist.Columns("Stock Card No").SortType = jgexSortTypeNumeric
    End If
    
    If ColumnExists("Job No") Then
        jgxPicklist.Columns("Job No").SortType = jgexSortTypeString
    End If
    
    If ColumnExists("Batch No") Then
        jgxPicklist.Columns("Batch No").SortType = jgexSortTypeString
    End If
    
    If ColumnExists("Doc Number") Then
        jgxPicklist.Columns("Doc Number").SortType = jgexSortTypeString
    End If
        
    If strColumnName = "" Then
        lngSortCounter = 2
        strColumnName = Column
    Else
        If strColumnName = Column Then
            lngSortCounter = lngSortCounter + 1
        Else
            strColumnName = Column
            lngSortCounter = 2
        End If
    End If
    
    If (lngSortCounter Mod 2) = 0 Then
        jgxPicklist.SortKeys.Clear
        jgxPicklist.SortKeys.Add Column.Index, jgexSortAscending
    Else
        jgxPicklist.SortKeys.Clear
        jgxPicklist.SortKeys.Add Column.Index, jgexSortDescending
    End If
    
    jgxPicklist.RefreshSort
'    End If
End Sub

Private Sub jgxPicklist_DblClick()
    If cmdOK.Enabled = True Then cmdOK_Click
End Sub

Private Sub jgxPicklist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call jgxPicklist_DblClick
    End If
End Sub

Private Sub txtBatchNo_Change()
    If strBatchNum <> txtBatchNo.Text And txtBatchNo.ForeColor = vbRed Then
        txtBatchNo.ForeColor = txtProductNo.ForeColor
    End If
End Sub

Private Sub txtBatchNo_GotFocus()
    txtBatchNo.SelStart = 0
    txtBatchNo.SelLength = Len(txtBatchNo.Text)
End Sub

Private Sub txtBatchNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtProductNo.SetFocus
    End If
End Sub

Private Sub txtCtryExport_Change()
    strCountryExp = txtCtryExport.Text
End Sub

Private Sub txtCtryExportDesc_Change()
    strCountryExpDesc = txtCtryExportDesc.Text
End Sub

Private Sub txtCtryOrigin_Change()
    strCountryOrig = txtCtryOrigin.Text
End Sub

Private Sub txtCtryOrigin_GotFocus()
    'Sets flag to signify execution (for VB bug in GotFocus not working).
    bytCtryKeys = 1

    strCtryOrigin = strCountryOrig
    
    txtCtryOrigin.SelStart = 0
    txtCtryOrigin.SelLength = Len(txtCtryOrigin.Text)
End Sub

Private Sub txtCtryOrigin_KeyDown(KeyCode As Integer, Shift As Integer)
    'First key triggers GotFocus only if GotFocus was not executed earlier (for VB bug in GotFocus not working).
    If bytCtryKeys = 0 Then txtCtryOrigin_GotFocus
        
    If KeyCode = vbKeyF2 Then
        cmdCountry_Click (0)
    ElseIf KeyCode = 13 Then
        txtCtryExport.SetFocus
    End If
End Sub

Private Sub txtCtryOrigin_LostFocus()
    'Resets Country Code textbox workaround counter (for VB bug in GotFocus not working).
    bytCtryKeys = 0
    
    'Performs check and auto description loading for Country.
    If (strCtryOrigin = strCountryOrig) Then
        If Len(strCountryOrigDesc) > 0 And Len(strCountryOrig) = 3 Then
            bytCtryOFound = 1
        Else
            bytCtryOFound = 0
        End If
    ElseIf Val(strCountryOrig) = 0 Then
        strCountryOrigDesc = Empty
    Else
        If Len(strCountryOrig) = 3 Then
            bytCtryOFound = 0
        Else
            'Prompts to open picklist.
            'If MsgBox("Country of Origin Code not found in database.  " & _
                      "Please ensure the code entered already exists." & vbCrLf & _
                      "Would you like to open the Country of Origin picklist?", vbYesNo + vbInformation, _
                      "Stock Card / Products") = vbYes Then
            If MsgBox(Translate(2229) & vbCrLf & _
                      Translate(2230), vbYesNo + vbInformation, _
                      Translate(2278)) = vbYes Then
                      
                cmdCountry_Click (0)
            Else
                'Revert to previous country code.
                strCountryOrig = strCtryOrigin
            End If
            
            bytCtryOFound = 1
        End If
    End If
    
    If bytCtryOFound = 0 And Len(strCountryOrig) = 3 Then
        strBlah = GetCountryDesc(strCountryOrig, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
        'Validates entered code based on description.
        If strBlah = "ALL YOUR BASE ARE BELONG TO US" Then
            'Prompts to open picklist.
            If MsgBox(Translate(2229) & vbCrLf & _
                      Translate(2230), vbYesNo + vbInformation, _
                      Translate(2278)) = vbYes Then
                cmdCountry_Click (0)
            Else
                'Revert to previous country code.
                strCountryOrig = strCtryOrigin
            End If
        Else
            strCountryOrigDesc = strBlah
        End If
        strBlah = Empty
        bytCtryOFound = 1
    End If
End Sub

Private Sub txtCtryExport_GotFocus()
    'Sets flag to signify execution (for VB bug in GotFocus not working).
    bytCtryKeys = 1
        
    strCtryExport = txtCtryExport.Text
    txtCtryExport.SelStart = 0
    txtCtryExport.SelLength = Len(txtCtryExport.Text)
End Sub

Private Sub txtCtryExport_KeyDown(KeyCode As Integer, Shift As Integer)
    'First key triggers GotFocus only if GotFocus was not executed earlier (for VB bug in GotFocus not working).
    If bytCtryKeys = 0 Then txtCtryExport_GotFocus
    
    If KeyCode = vbKeyF2 Then
        cmdCountry_Click (1)
    ElseIf KeyCode = 13 Then
'        txtJobNo.SetFocus
        If jgxPicklist.Visible = True And jgxPicklist.Enabled = True Then
            If jgxPicklist.RowCount > 0 Then
                jgxPicklist.SetFocus
            Else
                If cmdNew.Visible = True And cmdNew.Enabled = True Then
                    cmdNew.SetFocus
                Else
                    txtCtryExport.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtCtryExport_LostFocus()
    'Resets Country Code textbox workaround counter (for VB bug in GotFocus not working).
    bytCtryKeys = 0
        
    'Performs check and auto description loading for Country.
    If (strCtryExport = strCountryExp) Then
        If Len(strCountryExpDesc) > 0 And Len(strCountryExp) = 3 Then
            bytCtryEFound = 1
        Else
            bytCtryEFound = 0
        End If
    ElseIf Val(strCountryExp) = 0 Then
        strCountryExpDesc = Empty
    Else
        If Len(strCountryExp) = 3 Then
            bytCtryEFound = 0
        Else
            'Prompts to open picklist.
            'If MsgBox("Country of Export Code not found in datatbase.  " & _
                      "Please ensure the code entered already exists." & vbCrLf & _
                      "Would you like to open the Country of Export picklist?", vbYesNo + vbInformation, _
                      "Stock Card / Products") = vbYes Then
            If MsgBox(Translate(2272) & vbCrLf & _
                      Translate(2231), vbYesNo + vbInformation, _
                      Translate(2278)) = vbYes Then
                      
                cmdCountry_Click (1)
            Else
                'Revert to previous country code.
                txtCtryExport.Text = strCtryExport
            End If
            
            bytCtryEFound = 1
        End If
    End If
    
    If bytCtryEFound = 0 And Len(strCountryExp) = 3 Then
        strBlah = GetCountryDesc(strCountryExp, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
        'Validates entered code based on description.
        If strBlah = "ALL YOUR BASE ARE BELONG TO US" Then
            'Prompts to open picklist.
            If MsgBox(Translate(2272) & vbCrLf & _
                       Translate(2231), vbYesNo + vbInformation, _
                       Translate(2278)) = vbYes Then
                cmdCountry_Click (1)
            Else
                'Revert to previous country code.
                strCountryExp = strCtryExport
            End If
        Else
            strCountryExpDesc = strBlah
        End If
        strBlah = Empty
        bytCtryEFound = 1
    End If
End Sub

Private Sub txtCtryOriginDesc_Change()
    strCountryOrigDesc = txtCtryOriginDesc.Text
End Sub

Private Sub txtJobNo_Change()
    If strJobNum <> txtJobNo.Text And txtJobNo.ForeColor = vbRed Then
        txtJobNo.ForeColor = txtProductNo.ForeColor
    End If
End Sub

Private Sub txtJobNo_GotFocus()
    txtJobNo.SelStart = 0
    txtJobNo.SelLength = Len(txtJobNo.Text)
End Sub

Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtBatchNo.SetFocus
    End If
End Sub

Private Sub txtProductNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdProductNo_Click
    ElseIf KeyCode = 13 Then
        txtProductNo_LostFocus
        
        If fraInfo.Visible = True Then
            If txtCtryOrigin.Enabled = True Then
                txtCtryOrigin.SetFocus
            Else
                If jgxPicklist.Enabled And jgxPicklist.Visible And jgxPicklist.RowCount > 0 Then
                    jgxPicklist.SetFocus
                Else
                    If cmdNew.Visible = True And cmdNew.Enabled = True Then
                        cmdNew.SetFocus
                    Else
                        txtProductNo.SetFocus
                    End If
                End If
            End If
        Else
            If jgxPicklist.Visible = True And jgxPicklist.Enabled = True And jgxPicklist.RowCount > 0 Then
                jgxPicklist.SetFocus
            Else
                If cmdNew.Visible = True And cmdNew.Enabled = True Then
                    cmdNew.SetFocus
                Else
                    txtProductNo.SetFocus
                End If
            End If
        End If
    End If
End Sub


Private Sub txtProductNo_GotFocus()
    strProdNum = txtProductNo.Text
    txtProductNo.SelStart = 0
    txtProductNo.SelLength = Len(txtProductNo.Text)
End Sub

Private Sub txtProductNo_LostFocus()
    Dim strSQLWhere As String
    
    'Resets recordset checking if Product Number textbox was changed.
    If Not (strProdNum = txtProductNo.Text) Then
        bytProdFound = 0
        bytProdNumDeclined = 0
    End If
        
    'Only creates a recordset and prompt for quick registration if value was not registered.
    If bytProdFound = 0 And bytProdNumDeclined = 0 Then
        Dim rstProdNumVerify As ADODB.Recordset
        
        lngEntrepotID = GetEntrepot_ID(pckStockProd.Entrepot_Num, pckStockProd.m_conSADBEL)
        
        strSQLWhere = "WHERE "
        If blnLesserForm = False Then strSQLWhere = strSQLWhere & "P.Entrepot_ID = " & lngEntrepotID & " AND "
        strSQLWhere = strSQLWhere & "Prod_Num = '" & txtProductNo.Text & "'"
                
        ADORecordsetOpen "SELECT Prod_ID AS [Prod ID], Prod_Num AS [Prod Num], Prod_Handling as [Prod_Handling], Prod_Desc AS [Prod Desc], " & _
                  "Taric_Code AS [Taric Code], Prod_Ctry_Origin AS [Ctry Origin], " & _
                  "Prod_Ctry_Export AS [Ctry Export], P.Entrepot_ID AS [Entrepot ID], " & _
                  "E.Entrepot_StockCard_Numbering AS [Numbering], " & _
                  "E.Entrepot_Starting_Num AS [Starting Num] " & _
                  "FROM Products [P] INNER JOIN Entrepots [E] ON P.Entrepot_ID = E.Entrepot_ID " & _
                  strSQLWhere, _
                  pckStockProd.m_conSADBEL, rstProdNumVerify, adOpenKeyset, adLockOptimistic
                  
        With rstProdNumVerify
            '.Open "SELECT Prod_ID AS [Prod ID], Prod_Num AS [Prod Num], Prod_Handling as [Prod_Handling], Prod_Desc AS [Prod Desc], " & _
                  "Taric_Code AS [Taric Code], Prod_Ctry_Origin AS [Ctry Origin], " & _
                  "Prod_Ctry_Export AS [Ctry Export], P.Entrepot_ID AS [Entrepot ID], " & _
                  "E.Entrepot_StockCard_Numbering AS [Numbering], " & _
                  "E.Entrepot_Starting_Num AS [Starting Num] " & _
                  "FROM Products [P] INNER JOIN Entrepots [E] ON P.Entrepot_ID = E.Entrepot_ID " & _
                  strSQLWhere, _
                  pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
            
            'Checks if new Product Number.
            If (.BOF And .EOF) Then
                bytProdFound = 0                        'Flag that Product Number is not in DB.
                cmdNew.Enabled = False                  'Disables new Stock Card creation button.
                
                'First determines if can perform a quick Product registration based on pre-requisite fields.
                'Pre-requisite fields are as follows:
                '   [Codisheet BB]  - Entrepot Type-Num
                '   [Codisheet C1]  - Country of Origin
                '   [Codisheet C2]  - Country of Export
                '   [StockProd]     - Taric Code
                '   [StockProd]     - Product Description
                '** If any of the above fields are blank, entering a new Product Number will prompt
                '** to load the Product picklist instead of performing a quick Product registration.
                '--------------------------------------------------------------------------------------
                If Len(Trim(strCountryOrig)) < 2 Or Len(Trim(strCountryExp)) < 2 _
                Or Len(Trim$(txtProductDesc.Text)) = 0 Or Len(pckStockProd.Entrepot_Num) = 0 Then
                    'If called from codisheet (blnLesserForm = False).
                    If blnLesserForm = False Then
                        
                        If MsgBox(Translate(2279) & vbCrLf & _
                                  Translate(2281), vbYesNo + vbQuestion, Translate(2278)) = vbYes Then

                            'Sets products tag to empty first just in case users cancels Products picklist.
                            txtProductNo.Tag = Empty
                            cmdProductNo_Click

                            'Enables new Stock Card creation button when user selects a valid Product from the picklist.
                            If Len(txtProductNo.Tag) <> 0 Then cmdNew.Enabled = True
                        Else
                            'Sets products tag to empty as flag the product number entered is invalid.
                            txtProductNo.Tag = Empty
                        End If
'                        strProblems = Empty
                        
                    'If called from Summary Report (blnLesserForm = True).
                    '** blnLesserForm is not checked on second half of this
                    '** procedure since it's unlikely it will pass there.
                    ElseIf blnLesserForm = True Then
                        'Informs user about problems encountered and presents an alternate option.
                        If MsgBox(Translate(2279) & vbCrLf & _
                                  Translate(2288), vbYesNo + vbQuestion, Translate(2278)) = vbYes Then
                            
                            'Sets products tag to empty first just in case users cancels Products picklist.
                            txtProductNo.Tag = Empty
                            cmdProductNo_Click
                        Else
                            'Sets products tag to empty as flag the product number entered is invalid.
                            txtProductNo.Tag = Empty
                        End If
                    End If
                    
                    'Checks Product ID in property tag to determine what filter to apply in Grid's recordset.
                    If Len(txtProductNo.Tag) <> 0 Then
                        PopGrid txtProductNo.Tag, pckStockProd.m_conSADBEL
                    ElseIf Len(txtProductNo.Tag) = 0 And bytProdNumDeclined = 0 Then
                        PopGrid 0, pckStockProd.m_conSADBEL
                    End If
                    'Performs a row check on Grid to enable select as necessary.
                    jgxPicklist_Click
                                        
                    ADORecordsetClose rstProdNumVerify
                    
                    Exit Sub
                End If
                
                'Performs a quick add if pre-requisite fields have valid values.
                '--------------------------------------------------------------------------------------
                If Len(Trim(txtProductNo.Text)) > 0 Then
                    If MsgBox(Translate(2279) & vbCrLf & _
                              Translate(2281), vbYesNo + vbQuestion, Translate(2278)) = vbYes Then
                    
                        txtProductNo.Tag = Empty
                        cmdProductNo_Click
                    Else
                        'Sets products tag to empty as flag the product number entered is invalid.
                        txtProductNo.Tag = Empty
                    End If
                End If
            Else
                bytProdFound = 1                        'Flag that Product Number is in DB.
                
                pckStockProd.ProductHandling = .Fields("Prod_Handling").Value
                
                'Passes Product ID to Tag property for recording purposes.
                
                txtProductNo.Tag = .Fields("Prod ID").Value
                txtTaricCode.Text = .Fields("Taric Code").Value
                txtProductDesc.Text = .Fields("Prod Desc").Value
                strCountryOrig = .Fields("Ctry Origin").Value
                strCountryExp = .Fields("Ctry Export").Value
                txtCtryOrigin.Tag = .Fields("Ctry Origin").Value
                txtCtryExport.Tag = .Fields("Ctry Export").Value
                bytNumbering = .Fields("Numbering").Value
                strStartingNum = .Fields("Starting Num").Value
                
                strBlah = GetCountryDesc(.Fields("Ctry Origin").Value, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryOrigDesc = strBlah
                strBlah = GetCountryDesc(.Fields("Ctry Export").Value, pckStockProd.m_conSADBEL, pckStockProd.strLanguage)
                If Not (strBlah = "ALL YOUR BASE ARE BELONG TO US") Then strCountryExpDesc = strBlah
            End If
            
            'Clean up.
            ADORecordsetClose rstProdNumVerify
            
            'Enables new Stock Card creation button when user selects a valid Product from the picklist.
            If Len(txtProductNo.Tag) <> 0 And blnLesserForm = False Then
                cmdNew.Enabled = True
            Else
                cmdNew.Enabled = False
            End If
            
            'Checks Product ID in property tag to determine what filter to apply in Grid's recordset.
            If Len(txtProductNo.Tag) <> 0 Then
                PopGrid txtProductNo.Tag, pckStockProd.m_conSADBEL
            ElseIf Len(txtProductNo.Tag) = 0 And bytProdNumDeclined = 0 Then
                'Condition prevents lost focus from running PopGrid even when grid is already empty.
                If jgxPicklist.ADORecordset.RecordCount > 0 Then PopGrid 0, pckStockProd.m_conSADBEL

            End If
            'Performs a row check on Grid to enable select as necessary.
            jgxPicklist_Click
        End With
    End If
End Sub

Private Sub CtryCodeMod()
    Dim rstUpdate As ADODB.Recordset
    
    'This is used to apply Country Code changes to the Products table when inconsistency found.
    Dim bytProblemsFlag As Byte
    Dim strProblems As String
    bytProblemsFlag = 0

    'Locate errors and construct error message.
    If Len(strCountryOrig) > 0 And Len(txtCtryOrigin.Tag) > 0 Then
        If Not (strCountryOrig = txtCtryOrigin.Tag) Then
            strProblems = strProblems & Space(5) & "* Country of Origin" & vbCrLf
            bytProblemsFlag = bytProblemsFlag + 1
        End If
    End If
    
    If Len(strCountryExp) > 0 And Len(txtCtryExport.Tag) > 0 Then
        If Not (strCountryExp = txtCtryExport.Tag) Then
            strProblems = strProblems & Space(5) & "* Country of Export" & vbCrLf
            bytProblemsFlag = bytProblemsFlag + 2
        End If
    End If
    
    If bytProblemsFlag > 0 Then
        'Performs the Products table modification if user accepts.
        If MsgBox(Translate(2284) & vbCrLf & _
                  Translate(2285) & vbCrLf & strProblems & _
                  vbCrLf & Translate(2286), _
                  vbYesNo + vbQuestion, Translate(2278)) = vbYes Then
                    
            ADORecordsetOpen "SELECT Prod_ID, Prod_Ctry_Origin, Prod_Ctry_Export " & _
                           "FROM Products WHERE Prod_ID = " & txtProductNo.Tag, _
                           pckStockProd.m_conSADBEL, rstUpdate, adOpenKeyset, adLockOptimistic
            'rstUpdate.Open "SELECT Prod_ID, Prod_Ctry_Origin, Prod_Ctry_Export " & _
                           "FROM Products WHERE Prod_ID = " & txtProductNo.Tag, _
                           pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
            
            If Not (rstUpdate.EOF And rstUpdate.BOF) Then
                rstUpdate.MoveFirst
                
                'Applies the changes both to the table and to the tag.
                If Not (bytProblemsFlag = 2) Then
                    rstUpdate.Fields("Prod_Ctry_Origin").Value = strCountryOrig
                    txtCtryOrigin.Tag = strCountryOrig
                End If
                If Not (bytProblemsFlag = 1) Then
                    rstUpdate.Fields("Prod_Ctry_Export").Value = strCountryExp
                    txtCtryExport.Tag = strCountryExp
                End If
                rstUpdate.Update
                
                UpdateRecordset pckStockProd.m_conSADBEL, rstUpdate, "Products"
            End If
            
            ADORecordsetClose rstUpdate
        Else
            'What happens when user decides to cancel.
        End If
    End If
End Sub
    
Private Sub InitializeRecordset()
    'Initialize the recordset that will contain the temporary stockcard/s.
    
    With m_rstNewStockOff
        .CursorLocation = adUseClient
        .Fields.Append "Stock ID", adVarNumeric, 10
        .Fields.Append "Stock Card No", adVarChar, 10
        .Fields.Append "Product ID", adVarChar, 10
        .Open
    End With
End Sub

Private Sub SaveStockCards()
    Dim rstStockCard As ADODB.Recordset
    Dim strSQL As String
    Dim lngDifference As Long
    Dim lngCounter As Long
    
    Dim lngTempStockID As Long

    'Save the newly added stockcard/s into the database.
        strSQL = "SELECT Stock_ID AS [Stock ID], Stock_Card_Num AS [Stock Card No], Prod_ID AS [Product ID] FROM StockCards "
    
    ADORecordsetOpen strSQL, pckStockProd.m_conSADBEL, rstStockCard, adOpenKeyset, adLockOptimistic
    'rstStockCard.Open strSQL, pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
    

    If Not (m_rstNewStockOff.EOF And m_rstNewStockOff.BOF) Then
        m_rstNewStockOff.MoveFirst
        
        If Not rstStockCard.EOF Or Not rstStockCard.BOF Then
            rstStockCard.MoveLast
        End If

        For lngCounter = 1 To m_rstNewStockOff.RecordCount
            With rstStockCard
                .AddNew
                '    .Fields("Stock ID").Value = m_rstNewStockOff.Fields("Stock ID").Value
                .Fields("Stock Card No").Value = m_rstNewStockOff.Fields("Stock Card No").Value
                .Fields("Product ID").Value = m_rstNewStockOff.Fields("Product ID").Value
                
                'lngTempStockID = .Fields("Stock ID").Value
                
                .Update
                
                lngTempStockID = InsertRecordset(pckStockProd.m_conSADBEL, rstStockCard, "StockCards")
                
                ' Update Grid Stock_ID
                m_rstNewStockOff.Fields("Stock ID").Value = lngTempStockID
                
                If Not m_rstNewStockOff.EOF Then
                    m_rstNewStockOff.MoveNext
                End If
            End With
        Next lngCounter
    End If

    ADORecordsetClose rstStockCard
End Sub

Private Function UsedExistingStockcards(rstExistingStockcards As ADODB.Recordset, _
                                        strStartingNumber As String, _
                                        Optional blnRemoveFilter As Boolean) As Boolean
    If blnRemoveFilter = True Then
        rstExistingStockcards.Filter = adFilterNone
    End If
    
    With rstExistingStockcards
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While .EOF = False
                If .Fields("Stock Card No").Value = strStartingNumber Then
                    UsedExistingStockcards = True
                    Exit Do
                ElseIf .Fields("Stock Card No").Value > strStartingNumber Then
                    UsedExistingStockcards = False
                End If
                .MoveNext
            Loop
        End If
    End With
End Function

Private Function HighestStartingNumber(rstExistingStockcards As ADODB.Recordset, _
                                        strStartingNumber As String, _
                                        Optional blnRemoveFilter As Boolean) As String
    Dim lngLength As Long
    
    lngLength = Len(strStartingNumber)
    
    If blnRemoveFilter = True Then
        rstExistingStockcards.Filter = adFilterNone
    End If
    
    With rstExistingStockcards
        .MoveFirst
        
        Do While .EOF = False
            If .Fields("Stock Card No").Value = strStartingNumber Then
                strStartingNumber = Val(strStartingNumber) + 1
                If Len(strStartingNumber) < lngLength Then
                    strStartingNumber = String$(lngLength - Len(strStartingNumber), "0") & strStartingNumber
                End If
            Else
                strStartingNumber = Val(strStartingNumber)
                If Len(strStartingNumber) < lngLength Then
                    strStartingNumber = String$(lngLength - Len(strStartingNumber), "0") & strStartingNumber
                End If
            End If
            .MoveNext
        Loop
    End With
    
    HighestStartingNumber = strStartingNumber
End Function

Private Sub PopGrid2(ByRef ADOSadbel As ADODB.Connection)
    Dim strFirst(0 To 8) As String
    Dim strSecond(4 To 8) As String
    Dim blnSeveral(6 To 8) As Boolean               '0 not included since the "Stock Card No" field can't be <Several>.
    Dim lngCtr As Long
    Dim strIM7 As String
    Dim strSQL As String
    Dim strSCNumX As String
    
    
    If (m_blnFromSummaryReport = True) Then
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "E.Entrepot_ID AS [Entrepot ID], "
        strSQL = strSQL & "E.Entrepot_Type & '-' & E.Entrepot_Num AS [Entrepot Name], "
        strSQL = strSQL & "P.Prod_ID AS [Product ID], "
        strSQL = strSQL & "P.Prod_Num AS [Product Number], "
        strSQL = strSQL & "SC.Stock_ID AS [ID], "
        strSQL = strSQL & "SC.Stock_Card_Num AS [Stock Card No], "
        strSQL = strSQL & "I.In_Job_Num AS [Job No], "
        strSQL = strSQL & "I.In_Batch_Num AS [Batch No], "
        strSQL = strSQL & "ID.InDoc_Type AS [Doc Type], "
        strSQL = strSQL & "ID.InDoc_Num AS [Doc Num], "
        strSQL = strSQL & "E.Entrepot_StockCard_Numbering AS [Numbering], "
        strSQL = strSQL & "E.Entrepot_Starting_Num AS [Starting Num] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "(Entrepots [E] INNER JOIN ("
        strSQL = strSQL & "StockCards [SC] INNER JOIN Products [P] "
        strSQL = strSQL & "ON SC.Prod_ID = P.Prod_ID) "
        strSQL = strSQL & "ON E.Entrepot_ID = P.Entrepot_ID) "
        strSQL = strSQL & "LEFT JOIN ("
        strSQL = strSQL & "Inbounds [I] LEFT JOIN InboundDocs [ID] "
        strSQL = strSQL & "ON I.Indoc_ID = ID.Indoc_ID) "
        strSQL = strSQL & "ON SC.Stock_ID = I.Stock_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "E.Entrepot_Type & '-' & E.Entrepot_Num = '" & ProcessQuotes(Trim(strEntrepotName)) & "' "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "P.Prod_ID, "
        strSQL = strSQL & "SC.Stock_ID, "
        strSQL = strSQL & "I.In_Job_Num, "
        strSQL = strSQL & "SC.Stock_Card_Num "
        
    Else
        
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "E.Entrepot_ID AS [Entrepot ID], "
        strSQL = strSQL & "E.Entrepot_Type & '-' & E.Entrepot_Num AS [Entrepot Name], "
        strSQL = strSQL & "P.Prod_ID AS [Product ID], "
        strSQL = strSQL & "P.Prod_Num AS [Product Number], "
        strSQL = strSQL & "SC.Stock_ID AS [ID], "
        strSQL = strSQL & "SC.Stock_Card_Num AS [Stock Card No], "
        strSQL = strSQL & "I.In_Job_Num AS [Job No], "
        strSQL = strSQL & "I.In_Batch_Num AS [Batch No], "
        strSQL = strSQL & "ID.InDoc_Type AS [Doc Type], "
        strSQL = strSQL & "ID.InDoc_Num AS [Doc Num], "
        strSQL = strSQL & "E.Entrepot_StockCard_Numbering AS [Numbering], "
        strSQL = strSQL & "E.Entrepot_Starting_Num AS [Starting Num] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "("
        strSQL = strSQL & "Entrepots [E] INNER JOIN ("
        strSQL = strSQL & "StockCards [SC] INNER JOIN Products [P] "
        strSQL = strSQL & "ON SC.Prod_ID = P.Prod_ID) "
        strSQL = strSQL & "ON E.Entrepot_ID = P.Entrepot_ID) "
        strSQL = strSQL & "LEFT JOIN ("
        strSQL = strSQL & "Inbounds [I] LEFT JOIN InboundDocs [ID] "
        strSQL = strSQL & "ON I.Indoc_ID = ID.Indoc_ID) "
        strSQL = strSQL & "ON SC.Stock_ID = I.Stock_ID "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "P.Prod_ID, "
        strSQL = strSQL & "SC.Stock_ID, "
        strSQL = strSQL & "I.In_Job_Num, "
        strSQL = strSQL & "SC.Stock_Card_Num "
    End If

    ADORecordsetOpen strSQL, ADOSadbel, m_rstFindSeveral, adOpenKeyset, adLockOptimistic
    
    'Secondary recordset: contains the word <Several>.
    Set m_rstPass2GridOff = New ADODB.Recordset
    
    'Prepare second recordset for manual population.
    With m_rstPass2GridOff
        .CursorLocation = adUseClient
        .Fields.Append "Entrepot ID", adVarNumeric, 10
        .Fields.Append "Entrepot Number", adVarChar, 50
        .Fields.Append "Product ID", adVarNumeric, 10
        .Fields.Append "Product Number", adVarChar, 50
        .Fields.Append "Stock ID", adVarNumeric, 10
        .Fields.Append "Stock Card No", adVarChar, 10
        .Fields.Append "Job No", adVarChar, 50
        .Fields.Append "Batch No", adVarChar, 50
        .Fields.Append "Doc Number", adVarChar, 100
        .Fields.Append "New", adBoolean
        .Fields.Append "Length", adVarChar, 10
        .Open
    End With
    '---------------- Locating those severals ----------------
    With m_rstFindSeveral
        If .RecordCount > 0 Then .MoveFirst
        'Process primary recordset until last record or Stock Card No.
        Do Until .EOF = True
            
            'Stores first unique instance of:
            'Stock ID | Stock Card No | Job No | Batch No
            For lngCtr = 0 To 8
                Select Case lngCtr
                    Case Is < 8
                        If Not IsNull(.Fields(lngCtr).Value) Then
                            strFirst(lngCtr) = .Fields(lngCtr).Value
                        Else
                            'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                            strFirst(lngCtr) = Empty
                        End If
                    Case 8
                        If Not (IsNull(.Fields(8).Value) And IsNull(.Fields(9).Value)) Then
                            'Sets format of IM7 to "Type"-"Num".
                            strIM7 = .Fields(8).Value & "-" & .Fields(9).Value
                        Else
                            strIM7 = Empty
                        End If
                        strFirst(lngCtr) = strIM7
                    
                End Select
            Next lngCtr
            .MoveNext
            
            'If there's more than 1 record in the rst.
            If .RecordCount > 1 And Not .EOF Then
                'Stores second instance (possibly unique).
                'Stock ID | Stock Card No | Job No | Batch No
                For lngCtr = 4 To 8
                    If Not IsNull(.Fields(lngCtr).Value) Then
                        If Not lngCtr = 8 Then
                            strSecond(lngCtr) = .Fields(lngCtr).Value
                        Else
                            If Not (IsNull(.Fields(8).Value) And IsNull(.Fields(9).Value)) Then
                                'Sets format of IM7 to "Type"-"Num".
                                strIM7 = .Fields(8).Value & "-" & .Fields(9).Value
                            End If
                            strSecond(lngCtr) = strIM7
                        End If
                    Else
                        'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                        strSecond(lngCtr) = Empty
                    End If
                Next lngCtr
            'In case there's only 1 record in the rst or an EOF.
            Else
                'Stores dummy second instance (definitely unique).
                For lngCtr = 4 To 8
                    strSecond(lngCtr) = "dummy"
                Next lngCtr
            End If
            
            'Check if first and second are alike.
            'Perform record transfer when distinct.
            If strFirst(5) <> strSecond(5) Then
                'Copy first instance to a new recordset.
                'This is the same action perform after processing multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 8
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    strSCNumX = Empty
                End With
            'When First = Second.
            ElseIf strFirst(5) = strSecond(5) Then
                'Initialize/reset "Several" flag.
                For lngCtr = 6 To 8
                    blnSeveral(lngCtr) = False
                Next lngCtr
                    
                'Proceed with locating severals for every alike Stock Card Num.
                Do While strFirst(5) = strSecond(5)
                    'Cycles through other fields to identify which records will be labelled "Several".
                    For lngCtr = 6 To 8
                        'Used boolean flags to avoid too much string comparisons.
                        If blnSeveral(lngCtr) = False Then
                            'Copies next record to strSecond for comparison with strFirst.
                            If Not IsNull(.Fields(lngCtr).Value) Then           'CUSTOM
                                If Not lngCtr = 8 Then
                                    strSecond(lngCtr) = .Fields(lngCtr).Value
                                Else
                                    If Not (IsNull(.Fields(8).Value) And IsNull(.Fields(9).Value)) Then
                                        'Sets format of IM7 to "Type"-"Num".
                                        strIM7 = .Fields(8).Value & "-" & .Fields(9).Value
                                    End If
                                    strSecond(lngCtr) = strIM7
                                End If
                            Else
                                'MsgBox "Null value encountered in " & lngCtr & "."
                                strSecond(lngCtr) = Empty
                            End If
                                                                
                            'When strFirst != strSecond, a flag is raised to avoid processing this field.
                            If strFirst(lngCtr) <> strSecond(lngCtr) Then
                                blnSeveral(lngCtr) = True
                                'Stores the value "Several" for recording in first recordset.
                                strFirst(lngCtr) = "<Several>"
                                strSecond(lngCtr) = Empty
                            End If
                        End If
                    Next lngCtr

                    'Goes to next record or exits if last record of last distinct set.
                    .MoveNext
                    If .EOF = True Then Exit Do
                    strSecond(5) = .Fields(5).Value
                    
                Loop
                                
                'Copy first instance to a new recordset. After multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 8
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    strSCNumX = Empty
                End With
            End If
        Loop
    End With
    
    ADORecordsetClose m_rstFindSeveral

    'Commented to prevent new entries from being moved to top of grid.
    m_rstPass2GridOff.Sort = "[Length], [Stock Card No]"
    '------------ Finished locating those severals ------------
    Set jgxPicklist.ADORecordset = m_rstPass2GridOff
    HideSomeFields2
End Sub

Private Sub HideSomeFields2()
    'Hides the Stock ID and New fields.
    With jgxPicklist
        .Columns(1).Visible = False
        .Columns(3).Visible = False
        .Columns(5).Visible = False
        .Columns(10).Visible = False
        .Columns(11).Visible = False
    End With
End Sub

Private Sub Pass2Class2()
    With pckStockProd
        .BatchNo = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Batch No").Index)), "", jgxPicklist.Value(jgxPicklist.Columns("Batch No").Index))
        .JobNo = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Job No").Index)), "", jgxPicklist.Value(jgxPicklist.Columns("Job No").Index))
        .Product_ID = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Product ID").Index)), 0, jgxPicklist.Value(jgxPicklist.Columns("Product ID").Index))
        .ProductNo = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Product Number").Index)), "", jgxPicklist.Value(jgxPicklist.Columns("Product Number").Index))
        .Stock_ID = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Stock ID").Index)), 0, jgxPicklist.Value(jgxPicklist.Columns("Stock ID").Index))
        .StockCardNo = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Stock Card No").Index)), "", jgxPicklist.Value(jgxPicklist.Columns("Stock Card No").Index))
        .Entrepot_Num = IIf(IsNull(jgxPicklist.Value(jgxPicklist.Columns("Stock Card No").Index)), "", jgxPicklist.Value(jgxPicklist.Columns("Entrepot Number").Index))
    End With
End Sub

Private Function ColumnExists(ColumnName As String) As Boolean
    'Check if column caption exists.
    Dim lngCounter As Long
    
    For lngCounter = 1 To jgxPicklist.Columns.Count
        If jgxPicklist.Columns.Item(lngCounter).Caption = ColumnName Then
            ColumnExists = True
            Exit Function
        Else
            ColumnExists = False
        End If
    Next lngCounter
End Function

Private Sub ControlResizing()
    If txtBatchNo.Visible And txtJobNo.Visible And fraProduct.Visible Then
        If cmdDown.Caption = "&Hide..." Then
            'Expand!
            fraProduct.Height = 2175
            
            fraInfo.Visible = True
            
            txtTaricCode.Text = txtTaricCode.Text
            txtCtryExport.Text = strCountryExp
            txtCtryExportDesc.Text = strCountryExpDesc
            txtCtryOrigin.Text = strCountryOrig
            txtCtryOriginDesc.Text = strCountryOrigDesc
        Else
            'Hide!
            fraProduct.Height = 1575
            
            fraInfo.Visible = False
        End If
        
        fraStockCard.Top = fraProduct.Height + txtBatchNo.Height + txtJobNo.Height + 360
        cmdOK.Top = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height + 480
        cmdCancel.Top = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height + 480
        
        frmStockProdPicklist.Height = fraProduct.Height + fraStockCard.Height + txtBatchNo.Height + txtJobNo.Height
        frmStockProdPicklist.Height = frmStockProdPicklist.Height + (cmdOK.Height * 2) + 720
        
    ElseIf txtBatchNo.Visible = False And txtJobNo.Visible = False And fraProduct.Visible Then
        If cmdDown.Caption = "&Hide..." Then
            'Expand!
            fraProduct.Height = 2175
            
            fraInfo.Visible = True
            
            txtTaricCode.Text = txtTaricCode.Text
            txtCtryExport.Text = strCountryExp
            txtCtryExportDesc.Text = strCountryExpDesc
            txtCtryOrigin.Text = strCountryOrig
            txtCtryOriginDesc.Text = strCountryOrigDesc
        Else
            'Hide!
            fraProduct.Height = 1575
            
            fraInfo.Visible = False
        End If
        
        fraProduct.Top = 120
        fraStockCard.Top = fraProduct.Height + 240
        cmdOK.Top = fraProduct.Height + fraStockCard.Height + 360
        cmdCancel.Top = fraProduct.Height + fraStockCard.Height + 360
    
        frmStockProdPicklist.Height = fraProduct.Height + fraStockCard.Height + (cmdOK.Height * 2) + 600
    End If
End Sub
