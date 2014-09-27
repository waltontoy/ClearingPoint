VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmInitialStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Initial Stock"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmInitialStock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      Tag             =   "180"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtDocType 
      Height          =   315
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Tag             =   "179"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Tag             =   "178"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame fraProduct 
      Caption         =   "Product Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   9495
      Begin VB.Label lblTaric 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblHand 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6720
         TabIndex        =   25
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblCtryOriginDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7200
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCtryExportCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6720
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblCtryExportDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7200
         TabIndex        =   22
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblCtryOriginCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6720
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblCountryOrigin 
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Origin:"
         Height          =   255
         Left            =   5160
         TabIndex        =   12
         Tag             =   "2195"
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblTaricCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Taric Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Tag             =   "2275"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "2201"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblCountryExport 
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Export:"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Tag             =   "2196"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblHandling 
         BackStyle       =   0  'Transparent
         Caption         =   "Handling:"
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Tag             =   "2219"
         Top             =   990
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   9240
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.Frame fraStock 
      Caption         =   "Stock"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   9495
      Begin GridEX16.GridEX jgxStock 
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3836
         MethodHoldFields=   -1  'True
         Options         =   -1
         AllowColumnDrag =   0   'False
         RecordsetType   =   1
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         ColumnCount     =   8
         ColButtonStyle1 =   1
         CardCaption1    =   -1  'True
         ColCaption1     =   "Product Number"
         ColKey1         =   "Product Number"
         ColWidth1       =   1305
         ColButtonStyle2 =   1
         ColDefaultValue2=   "<New>"
         ColCaption2     =   "Stock Card No"
         ColKey2         =   "Stock Card No"
         ColWidth2       =   1200
         ColCaption3     =   "Num of Items"
         ColKey3         =   "Num of Items"
         ColWidth3       =   1095
         ColSortType3    =   2
         ColCaption4     =   "Gross Weight"
         ColKey4         =   "Gross Weight"
         ColWidth4       =   1095
         ColCaption5     =   "Net Weight"
         ColKey5         =   "Net Weight"
         ColWidth5       =   1005
         ColButtonStyle6 =   1
         ColCaption6     =   "Package Type"
         ColKey6         =   "Package Type"
         ColWidth6       =   1200
         ColCaption7     =   "Job Num"
         ColKey7         =   "Job Num"
         ColWidth7       =   1005
         ColCaption8     =   "Batch Num"
         ColKey8         =   "Batch Num"
         ColWidth8       =   1005
         ItemCount       =   0
         DataMode        =   1
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: New stocks with no outbound movements yet are marked in red.  All columns of these rows can be edited."
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   9255
      End
      Begin VB.Label lblNotice 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Type in stock information on the top line and then press ENTER to add the stock to the list."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.TextBox txtDocNum 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      MaxLength       =   25
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtEntrepotNum 
      Height          =   315
      Left            =   1800
      MaxLength       =   19
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblInboundDoc 
      BackStyle       =   0  'Transparent
      Caption         =   "Inbound Document:"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   150
      Width           =   1695
   End
   Begin VB.Label lblEntrepotNum 
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
Attribute VB_Name = "frmInitialStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngUserID As Long

Private m_conSADBEL As ADODB.Connection
Private m_conTaric As ADODB.Connection

Private m_rstStocksOff As ADODB.Recordset

Private WithEvents pckInboundDocsPicklist As PCubeLibPick.CPicklist
Attribute pckInboundDocsPicklist.VB_VarHelpID = -1
'Private ResourceHandler As Long
Private m_strLanguage As String
Private m_intTaricProperties As Integer

Private blnGridIsNothing As Boolean
Private blnEntered As Boolean
Private blnCancelMove As Boolean
Private blnSystemChanged As Boolean
Private lngProd_ID As Long
Private lngPack_Flag As Long
Private lngStockID As Long

Private strSQLPack As String
Private pckList As PCubeLibPick.CPicklist
Private gsdList As PCubeLibPick.CGridSeed
Private strStartingNum As String
Private bytNumbering As Byte
Private lngLastSeqNumber As Long

Private strNetWt As String
Private strGrossWt As String
Private strQuantity As String

Private m_alngDeleted() As Long
Private m_arrRow() As Variant
Private blnBypass As Boolean

Public Sub MyLoad(ByRef Sadbel As ADODB.Connection, ByRef Taric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, _
                  ByVal UserID As Long)
                  
    modGlobals.ResourceHandler = MyResourceHandler
    modGlobals.LoadResStrings Me, True
    
    m_strLanguage = Language
    m_lngUserID = UserID
    
    Set m_conSADBEL = Sadbel
    Set m_conTaric = Taric
    
    m_intTaricProperties = TaricProperties
    
    strSQLPack = "SELECT [PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Key Code]," & _
              "[PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Code]," & _
              "[PICKLIST MAINTENANCE " & m_strLanguage & "].[DESCRIPTION " & m_strLanguage & "] AS [Key Description] " & _
              "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE " & m_strLanguage & "] " & _
              "WHERE " & _
              "([PICKLIST DEFINITION].[BOX CODE]= 'E3') AND " & _
              "([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND " & _
              "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE " & m_strLanguage & "].[internal code]) "
               
    Me.Show vbModal

End Sub

Private Sub FormatStockGrid()
    
    jgxStock.Columns("Product Number").ButtonStyle = jgexButtonEllipsis
    jgxStock.Columns("Stock Card No").ButtonStyle = jgexButtonEllipsis
    jgxStock.Columns("Package Type").ButtonStyle = jgexButtonEllipsis
    jgxStock.Columns("Num of Items").TextAlignment = jgexAlignRight
    jgxStock.Columns("Gross Weight").TextAlignment = jgexAlignRight
    jgxStock.Columns("Net Weight").TextAlignment = jgexAlignRight
    jgxStock.Columns("Num of Items").MaxLength = 6
    jgxStock.Columns("Gross Weight").MaxLength = 12
    jgxStock.Columns("Net Weight").MaxLength = 12
    jgxStock.Columns("Batch Num").MaxLength = 50
    jgxStock.Columns("Job Num").MaxLength = 50
    jgxStock.Columns("Product Number").MaxLength = 50
    jgxStock.Columns("Stock Card No").MaxLength = 10
    jgxStock.Columns("Package Type").MaxLength = 2
    
    jgxStock.Columns("Stock_ID").DefaultValue = 0
    jgxStock.Columns("Prod_ID").DefaultValue = 0
    jgxStock.Columns("In_ID").DefaultValue = 0
    jgxStock.Columns("Pack_Flag").DefaultValue = 0
    jgxStock.Columns("Num of Items").DefaultValue = "0"
    jgxStock.Columns("Gross Weight").DefaultValue = "0"
    jgxStock.Columns("Net Weight").DefaultValue = "0"
    jgxStock.Columns("Handling").DefaultValue = 0
    jgxStock.Columns("NonEdit").DefaultValue = False
    
    jgxStock.Columns("Stock_ID").Visible = False
    jgxStock.Columns("Prod_ID").Visible = False
    jgxStock.Columns("In_ID").Visible = False
    jgxStock.Columns("Pack_Flag").Visible = False
    jgxStock.Columns("NonEdit").Visible = False
    jgxStock.Columns("Handling").Visible = False
    
    jgxStock.Columns("Product Number").Width = 1305
    jgxStock.Columns("Stock Card No").Width = 1200
    jgxStock.Columns("Num of Items").Width = 1095
    jgxStock.Columns("Gross Weight").Width = 1095
    jgxStock.Columns("Net Weight").Width = 1005
    jgxStock.Columns("Package Type").Width = 1200
    jgxStock.Columns("Job Num").Width = 1005
    jgxStock.Columns("Batch Num").Width = 1005
    
    jgxStock.Columns("Product Number").SortType = jgexSortTypeString
    jgxStock.Columns("Stock Card No").SortType = jgexSortTypeNumeric
    jgxStock.Columns("Num of Items").SortType = jgexSortTypeNumeric
    jgxStock.Columns("Gross Weight").SortType = jgexSortTypeNumeric
    jgxStock.Columns("Net Weight").SortType = jgexSortTypeNumeric
    jgxStock.Columns("Package Type").SortType = jgexSortTypeString
    jgxStock.Columns("Job Num").SortType = jgexSortTypeString
    jgxStock.Columns("Batch Num").SortType = jgexSortTypeString
    
End Sub

Private Sub cmdApply_Click()
    
    If Not CheckIfRowToAdd() Then
        Exit Sub
    End If
    
    If Not NotOkToSave Then
        'RACHELLE 092705: add checking for closure date
        If IsStockClosed = True Then
            MsgBox "A closure has already been done that includes this document.", vbInformation + vbOKOnly, "Initial Stock"
            Exit Sub
        End If
    
        UpdateTables
        cmdApply.Enabled = False
    Else
        jgxStock.SetFocus
    End If
    
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim strPreviousEntrepotID As String
    Dim clsEntrepot As cEntrepot
    
    Dim gsdPicklist As PCubeLibPick.CGridSeed
    Dim strSQL As String
    
    
    Select Case Index
        Case 0
            Set clsEntrepot = New cEntrepot
            strPreviousEntrepotID = txtEntrepotNum.Tag
            clsEntrepot.ShowEntrepot Me, m_conSADBEL, True, m_strLanguage, ResourceHandler, Me.txtEntrepotNum.Name, Val(txtEntrepotNum.Tag)
            
            If clsEntrepot.Cancelled = False Then
                If Len(strPreviousEntrepotID) > 0 And Not (strPreviousEntrepotID = txtEntrepotNum.Tag) Then
                    txtEntrepotNum.Text = clsEntrepot.SelectedEntrepot
                    
                    Call ResetControlValues(True)
                    
                    cmdBrowse(1).Enabled = True
                End If
            End If
            
            Set clsEntrepot = Nothing
        Case 1
        
            If txtEntrepotNum.Text <> "" Then
            
                lngLastSeqNumber = GetLastSeqNum(txtEntrepotNum.Text, m_conSADBEL)
        
                Set pckInboundDocsPicklist = New PCubeLibPick.CPicklist
                Set gsdPicklist = pckInboundDocsPicklist.SeedGrid("Doc Type", 800, "Left", "Sequence Number", 1300, "Left", "Document Number", 2954, "Left")
                
                strSQL = "Select DISTINCT InboundDocs.InDoc_ID AS ID, InDoc_Type AS [Doc Type], InDoc_Global, InboundDocs.InDoc_ID AS InDoc_ID, InboundDocs.InDoc_SeqNum as [Sequence Number], InDoc_Num as [Document Number], InDoc_Date as DocDate, InDoc_Office as DocOffice, InDoc_Cert_Type as Cert_Type, InDoc_Cert_Num as Cert_Num from InboundDocs LEFT JOIN (Inbounds left JOIN (StockCards LEFT join (Products left Join Entrepots on Products.Entrepot_ID = Entrepots.Entrepot_ID) on StockCards.Prod_ID = Products.Prod_ID)  on Inbounds.Stock_ID = StockCards.Stock_ID) On InboundDocs.InDoc_ID = Inbounds.InDoc_ID  where (InDoc_Global = -1 and ((Entrepot_Type & '-' & Entrepot_Num) = '" & txtEntrepotNum.Text & "' or Entrepots.Entrepot_ID IS NULL))"
                
                With pckInboundDocsPicklist
                    .MainSQL = "Select InDoc_ID AS ID, InDoc_Type AS [Doc Type], InDoc_Global, InDoc_ID, InboundDocs.InDoc_SeqNum as [Sequence Number], InDoc_Num as [Document Number], InDoc_Date as DocDate, InDoc_Office as DocOffice, InDoc_Cert_Type as Cert_Type, InDoc_Cert_Num as Cert_Num from InboundDocs"

                    .GetTopWhere = "InboundDocs.InDoc_ID"
                    .Search True, "Document Number", txtDocNum.Text
                    .Pick Me, cpiFilterCatalog, m_conSADBEL, strSQL, "ID", "Inbounds", vbModal, gsdPicklist, , , True, cpiKeyF2
                    
                    If .CancelTrans = False Then
                    
                        blnBypass = True
                        
                        If Not .SelectedRecord Is Nothing Then
                            txtDocType.Text = .SelectedRecord.RecordSource.Fields("Doc Type")
                            txtDocNum.Text = .SelectedRecord.RecordSource.Fields("Document Number")
                        End If
                        
                        UpdateLastSeqNum txtEntrepotNum.Text, m_conSADBEL, lngLastSeqNumber
                        
                        AddMissingInboundDocsInHistory m_conSADBEL
                    Else
                        blnBypass = False
                    End If
                End With
                
                Set pckInboundDocsPicklist = Nothing
                Set gsdPicklist = Nothing

            End If
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not CheckIfRowToAdd() Then
        Exit Sub
    End If
    
    If fraStock.Enabled Then
        If Not NotOkToSave Then
            'RACHELLE 092705: add checking for closure date
            If IsStockClosed = True Then
                MsgBox "A closure has already been done that includes this document.", vbInformation + vbOKOnly, "Initial Stock"
                Exit Sub
            End If

            UpdateTables

            Unload Me
        Else
            jgxStock.SetFocus
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    ReDim m_alngDeleted(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Erase m_arrRow
    
    ADORecordsetClose m_rstStocksOff
    
End Sub

Private Sub jgxStock_BeforeDelete(ByVal Cancel As GridEX16.JSRetBoolean)
    
    If jgxStock.Row > 0 Then
        cmdApply.Enabled = True
    End If

    If jgxStock.Value(jgxStock.Columns("In_ID").Index) <> 0 Then
        
        If m_alngDeleted(UBound(m_alngDeleted)) <> 0 Then
            ReDim Preserve m_alngDeleted(UBound(m_alngDeleted) + 1)
        End If
        m_alngDeleted(UBound(m_alngDeleted)) = jgxStock.Value(jgxStock.Columns("In_ID").Index)
        
    End If
    
End Sub

Private Sub jgxStock_BeforeUpdate(ByVal Cancel As GridEX16.JSRetBoolean)

    Dim lngCtr As Long
    Dim blnTemp As Boolean
    
    If jgxStock.ADORecordset Is Nothing Or blnGridIsNothing Then
        Exit Sub
    End If
    
    If jgxStock.Row = -1 Then
                
        If blnEntered Then
            
            If CheckInvalidRecord() Then
                Cancel = True
            Else
                blnSystemChanged = True
                jgxStock.Col = jgxStock.Columns("Product Number").Index
            End If
            
            blnEntered = False
            
            For lngCtr = 1 To jgxStock.Columns.Count
                m_arrRow(lngCtr) = jgxStock.Columns(lngCtr).DefaultValue
            Next
        ElseIf CheckIfRowToAdd() Then
                        
            For lngCtr = 1 To jgxStock.Columns.Count
                m_arrRow(lngCtr) = jgxStock.Columns(lngCtr).DefaultValue
            Next
        
            blnSystemChanged = True
            blnTemp = blnCancelMove
            blnCancelMove = False
            jgxStock.Delete
            blnCancelMove = blnTemp
            blnSystemChanged = False
            Exit Sub
        
        Else
                        
            For lngCtr = 1 To jgxStock.Columns.Count
                m_arrRow(lngCtr) = jgxStock.Value(lngCtr)
            Next
                        
            blnSystemChanged = True
            blnTemp = blnCancelMove
            blnCancelMove = False
            jgxStock.Delete
            blnCancelMove = blnTemp
            blnSystemChanged = False
            Exit Sub
            
        End If
        
    ElseIf jgxStock.Row > 0 Then
        
        If blnCancelMove Then
            jgxStock.EditMode = jgexEditModeOn
            jgxStock.SelStart = 0
            Cancel = True
        ElseIf CheckInvalidRecord() Then
            Cancel = True
        End If
                
    End If
    
    If Cancel = False Then
        cmdApply.Enabled = True
    End If
    
End Sub

Private Sub jgxStock_Change()
    
    If jgxStock.Row > 0 Then
        cmdApply.Enabled = True
    End If
    
    Select Case jgxStock.Col
    
        Case jgxStock.Columns("Package Type").Index
            jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 0
            lngPack_Flag = 0
        Case jgxStock.Columns("Stock Card No").Index
            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
            lngStockID = 0
        Case jgxStock.Columns("Product Number").Index
            jgxStock.Value(jgxStock.Columns("Prod_ID").Index) = 0
            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
            lngProd_ID = 0
            lngStockID = 0
            
    End Select
    
End Sub

Private Sub jgxStock_ColButtonClick(ByVal ColIndex As Integer)
    
    Set pckList = New CPicklist
    Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
        
    Select Case ColIndex
        
        Case jgxStock.Columns("Product Number").Index
            
            If Not ProdPick(False) Then
                If jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0 Then
                    jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                    jgxStock.EditMode = jgexEditModeOn
                    jgxStock.SelStart = 0
                End If
            End If
            
        Case jgxStock.Columns("Stock Card No").Index
            
            Call StockPick
            
        Case jgxStock.Columns("Package Type").Index
                                  
            With pckList
                    
                If Not IsNull(jgxStock.Value(jgxStock.Columns("Package Type").Index)) Then
                    If jgxStock.Value(jgxStock.Columns("Package Type").Index) <> "" Then
                        .Search True, "Key Code", jgxStock.Value(jgxStock.Columns("Package Type").Index)
                    End If
                End If
                
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyF2
                If Not .SelectedRecord Is Nothing Then
                    jgxStock.Value(jgxStock.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 1
                    lngPack_Flag = 1
                End If
            End With
        
    End Select
    
    Set gsdList = Nothing
    Set pckList = Nothing
    
End Sub

Private Sub jgxStock_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX16.JSRetBoolean)
    
    If jgxStock.AllowDelete = False Then
        
        If UCase(jgxStock.Columns(ColIndex).Key) = "STOCK CARD NO" Or UCase(jgxStock.Columns(ColIndex).Key) = "PRODUCT NUMBER" Or _
            UCase(jgxStock.Columns(ColIndex).Key) = "PACKAGE TYPE" Or (UCase(jgxStock.Columns(ColIndex).Key) = "GROSS WEIGHT" And _
            jgxStock.Value(jgxStock.Columns("Handling").Index) <> 1) Or (UCase(jgxStock.Columns(ColIndex).Key) = "NUM OF ITEMS" And _
            jgxStock.Value(jgxStock.Columns("Handling").Index) <> 0) Or (UCase(jgxStock.Columns(ColIndex).Key) = "NET WEIGHT" And _
            jgxStock.Value(jgxStock.Columns("Handling").Index) <> 2) Then
            
            Cancel = True
            
        End If
        
    End If

    If ColIndex = jgxStock.Columns("Stock Card No").Index And jgxStock.Row = -1 And _
        jgxStock.Value(jgxStock.Columns("Prod_ID").Index) = 0 Then
        
        Cancel = True
        Exit Sub
        
    End If
    
End Sub

Private Sub jgxStock_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    
    If jgxStock.SortKeys.Count > 0 Then
        If jgxStock.SortKeys.Item(1).ColIndex = Column.Index Then
            jgxStock.SortKeys.Item(1).SortOrder = IIf(jgxStock.SortKeys.Item(1).SortOrder = jgexSortAscending, jgexSortDescending, jgexSortAscending)
        Else
            jgxStock.SortKeys.Clear
            jgxStock.SortKeys.Add Column.Index, jgexSortAscending
        End If
    Else
        jgxStock.SortKeys.Add Column.Index, jgexSortAscending
    End If
    
End Sub

Private Sub jgxStock_KeyPress(KeyAscii As Integer)

    If jgxStock.Col > 0 Then
    
        Select Case UCase(jgxStock.Columns(jgxStock.Col).Key)
        
            Case "NUM OF ITEMS", "NET WEIGHT", "GROSS WEIGHT", "STOCK CARD NO"
                If Chr(KeyAscii) = "." Then
                    If UCase(jgxStock.Columns(jgxStock.Col).Key) = "NUM OF ITEMS" Or UCase(jgxStock.Columns(jgxStock.Col).Key) = "STOCK CARD NO" Then
                        KeyAscii = 0
                    ElseIf InStr(CStr(IIf(IsNull(jgxStock.Value(jgxStock.Col)), "", jgxStock.Value(jgxStock.Col))), ".") Then
                        KeyAscii = 0
                    End If
                ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
                    KeyAscii = 0
                ElseIf UCase(jgxStock.Columns(jgxStock.Col).Key) <> "NUM OF ITEMS" And IsNumeric(Chr(KeyAscii)) And UCase(jgxStock.Columns(jgxStock.Col).Key) <> "STOCK CARD NO" Then
                    Dim lngCount As Long
                    
                    lngCount = IIf(UCase(jgxStock.Columns(jgxStock.Col).Key) = "NET WEIGHT", 3, 2)
                    If Len(CStr(IIf(IsNull(jgxStock.Value(jgxStock.Col)), "", jgxStock.Value(jgxStock.Col)))) - _
                        InStrRev(CStr(IIf(IsNull(jgxStock.Value(jgxStock.Col)), "", jgxStock.Value(jgxStock.Col))), ".") >= lngCount _
                        And InStr(CStr(IIf(IsNull(jgxStock.Value(jgxStock.Col)), "", jgxStock.Value(jgxStock.Col))), ".") > 0 And _
                        jgxStock.SelLength = 0 And jgxStock.SelStart >= InStr(CStr(IIf(IsNull(jgxStock.Value(jgxStock.Col)), "", _
                        jgxStock.Value(jgxStock.Col))), ".") Then
                        
                        KeyAscii = 0
                        
                    End If
                ElseIf UCase(jgxStock.Columns(jgxStock.Col).Key) = "STOCK CARD NO" And UCase(jgxStock.Value(jgxStock.Col)) = "<NEW>" Then
                    jgxStock.SelStart = 0
                    jgxStock.SelLength = Len(jgxStock.Value(jgxStock.Col))
                End If
                
        End Select
        
    End If
    
End Sub

Private Sub jgxStock_KeyDown(KeyCode As Integer, Shift As Integer)

    blnSystemChanged = False
    blnCancelMove = False
    
    If KeyCode = vbKeyReturn Then
        blnEntered = True
    ElseIf KeyCode = vbKeyF2 Then
    
        Set pckList = New CPicklist
        Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
            
        Select Case jgxStock.Col
            
            Case jgxStock.Columns("Product Number").Index
                
                If Not ProdPick(False) Then
                    If jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0 Then
                        jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                        jgxStock.EditMode = jgexEditModeOn
                        jgxStock.SelStart = 0
                    End If
                End If
                
            Case jgxStock.Columns("Stock Card No").Index
                
                Call StockPick
                
            Case jgxStock.Columns("Package Type").Index
                                      
                With pckList
                    
                    If Not IsNull(jgxStock.Value(jgxStock.Columns("Package Type").Index)) Then
                        If jgxStock.Value(jgxStock.Columns("Package Type").Index) <> "" Then
                            .Search True, "Key Code", jgxStock.Value(jgxStock.Columns("Package Type").Index)
                        End If
                    End If
                    
                    ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                    .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyF2
                    If Not .SelectedRecord Is Nothing Then
                        jgxStock.Value(jgxStock.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                        jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 1
                        lngPack_Flag = 1
                    End If
                End With
            
        End Select
        
        Set gsdList = Nothing
        Set pckList = Nothing
    
    End If

End Sub

Private Sub jgxStock_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    blnSystemChanged = False
    blnCancelMove = False
End Sub

Private Sub jgxStock_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    
    If jgxStock.Row <> LastRow And jgxStock.Row <> 0 And Not jgxStock.ADORecordset Is Nothing Then
        Dim strSQL As String
        Dim rstProd As ADODB.Recordset
                
        strSQL = "SELECT Products!Prod_Handling AS Prod_Handling, " & _
            "Products!Prod_Num AS Prod_Num, " & _
            "Products!Prod_Desc AS Prod_Desc, " & _
            "Products!Taric_Code AS Taric_Code, " & _
            "Products!Prod_Ctry_Origin AS Origin_Code, " & _
            "M1.[Description " & IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS Origin_Desc, " & _
            "Products!Prod_Ctry_Export AS Export_Code, " & _
            "M2.[Description " & IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS Export_Desc " & _
            "FROM (Products INNER JOIN " & _
            "[Picklist Maintenance " & IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 ON Products.Prod_Ctry_Origin = M1.Code) " & _
            "INNER JOIN [Picklist Maintenance " & IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 ON Products.Prod_Ctry_Export = M2.Code " & _
            "WHERE Products!Prod_ID =" & jgxStock.Value(jgxStock.Columns("Prod_ID").Index) & _
            " AND M1.[Internal Code] = '8.29801619052887E+19' AND M2.[Internal Code] = '8.29801619052887E+19'"

        ADORecordsetOpen strSQL, m_conSADBEL, rstProd, adOpenKeyset, adLockOptimistic
        'rstProd.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
        
        If Not (rstProd.EOF And rstProd.BOF) Then
            
            rstProd.MoveFirst
            
            lblCtryExportCode.Caption = rstProd!Export_Code
            lblCtryExportDesc.Caption = rstProd!Export_Desc
            lblCtryOriginCode.Caption = rstProd!Origin_Code
            lblCtryOriginDesc.Caption = rstProd!Origin_Desc
            lblDesc.Caption = rstProd!Prod_Desc
            lblHand.Caption = Choose(rstProd!Prod_Handling + 1, "Number of Items", "Gross Weight", "Net Weight")
            lblTaric.Caption = rstProd!Taric_Code
            
        End If
        
        ADORecordsetClose rstProd
        
        If jgxStock.Value(jgxStock.Columns("NonEdit").Index) Then
            jgxStock.AllowDelete = False
        Else
            jgxStock.AllowDelete = True
        End If
        
    End If
    
    If Not blnSystemChanged And _
        Not jgxStock.ADORecordset Is Nothing And _
        LastRow <> 0 Then
        
        Dim lngrow As Long
        
        Select Case LastCol
            
            Case jgxStock.Columns("Stock Card No").Index
                
                If lngStockID = 0 And lngProd_ID <> 0 And jgxStock.Col <> jgxStock.Columns("Product Number").Index Then
                    
                    Dim strStock As String
                        
                    blnSystemChanged = True
                    lngrow = jgxStock.Row
                    jgxStock.Row = LastRow
                    strStock = IIf(IsNull(jgxStock.Value(LastCol)), "", jgxStock.Value(LastCol))
                    
                    If UniqueStockcard(False, jgxStock.Value(jgxStock.Columns("Prod_ID").Index), strStock) Then
                        If Not StockPick() Then
                            jgxStock.Col = LastCol
                            jgxStock.EditMode = jgexEditModeOn
                            jgxStock.SelStart = 0
                            blnCancelMove = True
                        Else
                            jgxStock.Row = lngrow
                        End If
                    Else
                        If lngStockID <> 0 Then
                            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
                            jgxStock.Row = lngrow
                        Else
                            jgxStock.Col = LastCol
                            jgxStock.EditMode = jgexEditModeOn
                            jgxStock.SelStart = 0
                            blnCancelMove = True
                        End If
                    End If
                    
                End If
                
            Case jgxStock.Columns("Product Number").Index
            
                If lngProd_ID = 0 Then
                
                    blnSystemChanged = True
                    lngrow = jgxStock.Row
                    jgxStock.Row = LastRow
                    
                    If ProdPick(True) Then
                        jgxStock.Col = LastCol
                        jgxStock.EditMode = jgexEditModeOn
                        jgxStock.SelStart = 0
                        blnCancelMove = True
                    ElseIf jgxStock.Value(jgxStock.Columns("Stock_ID").Index) <= 0 Then
                        jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                        jgxStock.SelStart = 0
                        jgxStock.EditMode = jgexEditModeOn
                        blnCancelMove = True
                    Else
                        jgxStock.Row = lngrow
                    End If
                                        
                End If
                
            Case jgxStock.Columns("Package Type").Index
            
                If lngPack_Flag = 0 Then
                                        
                    Dim strPack As String
                    
                    blnSystemChanged = True
                    lngrow = jgxStock.Row
                    jgxStock.Row = LastRow
                    strPack = IIf(IsNull(jgxStock.Value(LastCol)), "", jgxStock.Value(LastCol))
                    
                    Set pckList = New CPicklist
                    Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
                                
                    With pckList
                
                        If Trim(strPack) <> "" Then
                            .Search True, "Key Code", strPack
                        End If
                        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                        .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyEnter
                                        
                        If Not .SelectedRecord Is Nothing Then
                            jgxStock.Value(LastCol) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                            jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 1
                            lngPack_Flag = 1
                            jgxStock.Row = lngrow
                        Else
                            jgxStock.Col = LastCol
                            jgxStock.EditMode = jgexEditModeOn
                            jgxStock.SelStart = 0
                            blnCancelMove = True
                        End If
                
                    End With
                    
                    Set pckList = Nothing
                    Set gsdList = Nothing
                    
                End If
                
        End Select
            
    ElseIf blnSystemChanged And blnCancelMove Then
        blnCancelMove = False
        jgxStock.Row = LastRow
        jgxStock.Col = LastCol
        jgxStock.EditMode = jgexEditModeOn
        jgxStock.SelStart = 0
    End If
    
    If LastRow <> jgxStock.Row Then
        
        If jgxStock.Row = -1 And Not jgxStock.ADORecordset Is Nothing And LastRow <> 0 Then
            Dim lngCounter As Long
            For lngCounter = LBound(m_arrRow) To UBound(m_arrRow)
                jgxStock.Value(lngCounter) = m_arrRow(lngCounter)
            Next
        End If
        
        Call SetIDs
        
    End If
    
End Sub




Private Sub pckInboundDocsPicklist_BeforeDelete(ByVal BaseName As String, ByVal ID As Variant, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Dim rstInbounds As ADODB.Recordset

    ADORecordsetOpen "Select * from Inbounds where InDoc_ID = " & ID, m_conSADBEL, rstInbounds, adOpenKeyset, adLockOptimistic
    'rstInbounds.Open "Select * from Inbounds where InDoc_ID = " & ID, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
    If Not (rstInbounds.BOF And rstInbounds.EOF) Then
        MsgBox "Unable to delete selected item. Stocks are already associated with this record.", vbInformation, "Initial Stocks"
        Cancel = True
    Else
        Cancel = False
    End If
    
    ADORecordsetClose rstInbounds
End Sub

Private Sub txtDocNum_Change()
    If blnBypass = True Then
        blnBypass = False
        Call InitStocks
        
    ElseIf ExistingDocNum Then
        Call InitStocks
    Else
        lblCtryExportCode.Caption = ""
        lblCtryExportDesc.Caption = ""
        lblCtryOriginCode.Caption = ""
        lblCtryOriginDesc.Caption = ""
        lblDesc.Caption = ""
        lblHand.Caption = ""
        lblTaric.Caption = ""
        
        blnGridIsNothing = True
        Set jgxStock.ADORecordset = Nothing
        blnGridIsNothing = False
        Call CreateGridCol
        fraStock.Enabled = False
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtEntrepotNum_Change()
    
    Dim strSQL As String
    Dim rstEntrepot As ADODB.Recordset
    
        strSQL = "SELECT * FROM ENTREPOTS Where ENTREPOT_TYPE & '-' & ENTREPOT_NUM = '" & txtEntrepotNum.Text & "'"
    ADORecordsetOpen strSQL, m_conSADBEL, rstEntrepot, adOpenKeyset, adLockOptimistic
    'rstEntrepot.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If Not (rstEntrepot.EOF And rstEntrepot.BOF) Then
        
        txtDocNum.Enabled = True
        cmdBrowse(1).Enabled = True
        Call txtDocNum_Change
        
    Else
        If Len(txtEntrepotNum.Text) > 0 Then
            Call ResetControlValues(True)
        Else
            Call ResetControlValues(False)
        End If
'        txtDocNum.Enabled = False
        cmdBrowse(1).Enabled = False
        blnGridIsNothing = True
        Set jgxStock.ADORecordset = Nothing
        blnGridIsNothing = False
        Call CreateGridCol
        fraStock.Enabled = False
        cmdApply.Enabled = False
    
    End If
    
    ADORecordsetClose rstEntrepot
End Sub

Private Function ProductExist(ByVal strProd, ByRef Handling As Integer) As Boolean
        
    Dim strSQL As String
    Dim rstProd As ADODB.Recordset
            
    strSQL = "SELECT Products!Prod_ID AS Prod_ID, Products!Prod_Handling AS Handling " & _
        "FROM (Products INNER JOIN " & "[Picklist Maintenance " & _
        IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & _
        "] AS M1 ON Products.Prod_Ctry_Origin = M1.Code) " & "INNER JOIN [Picklist Maintenance " & _
        IIf(UCase(m_strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(m_strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & _
        "] AS M2 ON Products.Prod_Ctry_Export = M2.Code " & _
        "WHERE Products!Prod_Num = '" & strProd & _
        "' AND M1.[Internal Code] = '8.29801619052887E+19' AND M2.[Internal Code] = '8.29801619052887E+19' "

    ADORecordsetOpen strSQL, m_conSADBEL, rstProd, adOpenKeyset, adLockOptimistic
    'rstProd.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If Not (rstProd.EOF And rstProd.BOF) Then
        rstProd.MoveFirst
        
        ProductExist = True
        lngProd_ID = rstProd!Prod_ID
        Handling = rstProd!Handling
    End If
    
    ADORecordsetClose rstProd
End Function

Private Sub InitStocks()
    
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    Dim lngCtr As Long
    Dim jsFormat As JSFmtCondition
    
    strSQL = "SELECT Inbounds!In_ID AS In_ID, " & _
            "Products!Prod_ID AS Prod_ID, " & _
            "StockCards!Stock_ID AS Stock_ID, " & _
            "Products!Prod_Num AS [Product Number], " & _
            "Products!Prod_Handling AS Handling, " & _
            "StockCards!Stock_Card_Num AS [Stock Card No], " & _
            "Inbounds!In_Avl_Qty_Wgt AS Available, " & _
            "Inbounds!In_Orig_Packages_Qty AS Qty, " & _
            "Inbounds!In_Orig_Gross_Weight AS Gross, " & _
            "Inbounds!In_Orig_Net_Weight AS Net, " & _
            "Inbounds!In_Orig_Packages_Type AS [Package Type], " & _
            "Inbounds!In_Batch_Num AS [Batch Num], " & _
            "Inbounds!In_Job_Num AS [Job Num], " & _
            "(CHOOSE(Products!Prod_Handling +1, Inbounds!In_Orig_Packages_Qty, " & _
            "Inbounds!In_Orig_Gross_Weight, Inbounds!In_Orig_Net_Weight)>Inbounds!In_Avl_Qty_Wgt) AS NonEdit " & _
            "FROM InBoundDocs INNER JOIN (Inbounds INNER JOIN (StockCards INNER JOIN " & _
            "(Products INNER JOIN Entrepots ON Products.Entrepot_ID = Entrepots.Entrepot_ID) " & _
            "ON StockCards.Prod_ID = Products.Prod_ID) ON Inbounds.Stock_ID = StockCards.Stock_ID) ON " & _
            "InBoundDocs.Indoc_ID = Inbounds.Indoc_ID " & _
            "WHERE InBoundDocs!Indoc_Num='" & txtDocNum.Text & _
            "' AND InBoundDocs!Indoc_Global AND " & _
            "Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num = '" & txtEntrepotNum.Text & _
            "' AND IIF(ISNULL(Inbounds!In_Avl_Qty_Wgt),0,Inbounds!In_Avl_Qty_Wgt)>0 "
    
    ADORecordsetOpen strSQL, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    Set jgxStock.ADORecordset = Nothing
    
    ADORecordsetClose m_rstStocksOff
    
    
    Set m_rstStocksOff = New ADODB.Recordset
    m_rstStocksOff.CursorLocation = adUseClient
    
    m_rstStocksOff.Fields.Append "Product Number", rstTemp.Fields("Product Number").Type, rstTemp.Fields("Product Number").DefinedSize, rstTemp.Fields("Product Number").Attributes
    m_rstStocksOff.Fields.Append "Stock Card No", rstTemp.Fields("Stock Card No").Type, rstTemp.Fields("Stock Card No").DefinedSize, rstTemp.Fields("Stock Card No").Attributes
    m_rstStocksOff.Fields.Append "Num of Items", adVarWChar, 6
    m_rstStocksOff.Fields.Append "Gross Weight", adVarWChar, 12
    m_rstStocksOff.Fields.Append "Net Weight", adVarWChar, 12
    m_rstStocksOff.Fields.Append "Package Type", rstTemp.Fields("Package Type").Type, rstTemp.Fields("Package Type").DefinedSize, rstTemp.Fields("Package Type").Attributes
    m_rstStocksOff.Fields.Append "Job Num", rstTemp.Fields("Job Num").Type, rstTemp.Fields("Job Num").DefinedSize, rstTemp.Fields("Job Num").Attributes
    m_rstStocksOff.Fields.Append "Batch Num", rstTemp.Fields("Batch Num").Type, rstTemp.Fields("Batch Num").DefinedSize, rstTemp.Fields("Batch Num").Attributes
    m_rstStocksOff.Fields.Append "In_ID", adInteger
    m_rstStocksOff.Fields.Append "Prod_ID", adInteger
    m_rstStocksOff.Fields.Append "Stock_ID", adInteger
    m_rstStocksOff.Fields.Append "Pack_Flag", adInteger
    m_rstStocksOff.Fields.Append "NonEdit", adBoolean
    m_rstStocksOff.Fields.Append "Handling", adInteger
    m_rstStocksOff.Open
    
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        
        rstTemp.MoveFirst
                
        Do While Not rstTemp.EOF
            
            m_rstStocksOff.AddNew
            
            m_rstStocksOff!In_ID = rstTemp!In_ID
            m_rstStocksOff!Prod_ID = rstTemp!Prod_ID
            m_rstStocksOff!Stock_ID = rstTemp!Stock_ID
            m_rstStocksOff!Pack_Flag = 1
            m_rstStocksOff![Product Number] = rstTemp![Product Number]
            m_rstStocksOff![Stock Card No] = rstTemp![Stock Card No]
            m_rstStocksOff![Package Type] = IIf(IsNull(rstTemp![Package Type]), "", rstTemp![Package Type])
            m_rstStocksOff![Batch Num] = IIf(IsNull(rstTemp![Batch Num]), "", rstTemp![Batch Num])
            m_rstStocksOff![Job Num] = IIf(IsNull(rstTemp![Job Num]), "", rstTemp![Job Num])
            m_rstStocksOff!NonEdit = rstTemp!NonEdit
            m_rstStocksOff!Handling = rstTemp!Handling
            
            Call FillQuantity(rstTemp!Handling, rstTemp!Qty, rstTemp!Gross, rstTemp!Net, rstTemp!Available)

            m_rstStocksOff![Num of Items] = strQuantity
            m_rstStocksOff![Gross Weight] = strGrossWt
            m_rstStocksOff![Net Weight] = strNetWt
            m_rstStocksOff.Update
            
            rstTemp.MoveNext
        Loop
        
    End If
    
    ADORecordsetClose rstTemp
    
    Set jgxStock.ADORecordset = m_rstStocksOff
    
    Set jsFormat = jgxStock.FmtConditions.Add(jgxStock.Columns("NonEdit").Index, jgexNotEqual, True)
    jsFormat.FormatStyle.ForeColor = vbRed
    
    Set jsFormat = Nothing
    
    Call FormatStockGrid
    
    Erase m_arrRow
    ReDim m_arrRow(1 To jgxStock.Columns.Count)
    
    For lngCtr = 1 To jgxStock.Columns.Count
        m_arrRow(lngCtr) = jgxStock.Columns(lngCtr).DefaultValue
    Next
    
    fraStock.Enabled = True
    cmdApply.Enabled = False
    
End Sub

Private Function ProdPick(ByVal AutoUnload As Boolean) As Boolean
            
    Dim pckProducts As PCubeLibEntrepot.cProducts
    Dim strProdNum As String
    Dim intHandling As Integer

    strProdNum = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Product Number").Index)), "", jgxStock.Value(jgxStock.Columns("Product Number").Index))
    
    If AutoUnload Then
        If ProductExist(strProdNum, intHandling) Then
        
            jgxStock.Value(jgxStock.Columns("Prod_ID").Index) = lngProd_ID
            jgxStock.Value(jgxStock.Columns("Handling").Index) = intHandling
            
            If UniqueStockcard(False, lngProd_ID, IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", (jgxStock.Value(jgxStock.Columns("Stock Card No").Index)))) Then
                If NewSequential(jgxStock.Value(jgxStock.Columns("Product Number").Index)) Then
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = "<New>"
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = -1
                    lngStockID = -1
                Else
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = ""
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
                    lngStockID = 0
                End If
            Else
                If lngStockID <> 0 Then
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
                ElseIf NewSequential(jgxStock.Value(jgxStock.Columns("Product Number").Index)) Then
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = "<New>"
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = -1
                    lngStockID = -1
                Else
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = ""
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
                    lngStockID = 0
                End If
            End If
            
            Exit Function
            
        End If
        
    End If
    
    Set pckProducts = New PCubeLibEntrepot.cProducts
    
    With pckProducts
                    
        .Entrepot_Num = txtEntrepotNum.Text
        .Product_Num = strProdNum
        .ShowProducts 1, Me, m_conSADBEL, m_conTaric, m_strLanguage, _
                                    m_intTaricProperties, ResourceHandler, strProdNum

        'Prevents updating of Stock/Prod picklist when Product selection has been cancelled.
        If .Cancelled = False Then
            
                                
            jgxStock.Value(jgxStock.Columns("Product Number").Index) = .Product_Num
            jgxStock.Value(jgxStock.Columns("Prod_ID").Index) = .Product_ID
            jgxStock.Value(jgxStock.Columns("Handling").Index) = .Prod_Handling
            lngProd_ID = .Product_ID
            lblDesc.Caption = .Prod_Desc
            lblHand.Caption = Choose(.Prod_Handling + 1, "Number of Items", "Gross Weight", "Net Weight")
            lblTaric.Caption = .Taric_Code
            lblCtryOriginCode.Caption = .Ctry_Origin
            lblCtryExportCode.Caption = .Ctry_Export
            lblCtryOriginDesc.Caption = .Origin_Desc
            lblCtryExportDesc.Caption = .Export_Desc
            
            If UniqueStockcard(False, .Product_ID, IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", (jgxStock.Value(jgxStock.Columns("Stock Card No").Index)))) Then
                
                If NewSequential(.Product_Num) Then
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = "<New>"
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = -1
                    lngStockID = -1
                Else
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = ""
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
                    lngStockID = 0
                End If
            
            Else
                
                If lngStockID <> 0 Then
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
                ElseIf NewSequential(.Product_Num) Then
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = "<New>"
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = -1
                    lngStockID = -1
                Else
                    jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = ""
                    jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = 0
                    lngStockID = 0
                End If
                            
            End If
            
        End If
                        
        ProdPick = .Cancelled
        
    End With
    
    Set pckProducts = Nothing

End Function

Private Sub GetStartingNo()
    Dim strSQL As String
    Dim rstStartingNo As ADODB.Recordset
    
    strSQL = "Select Entrepots.Entrepot_Starting_Num AS [Starting Num], " & _
        "Entrepots.Entrepot_Stockcard_Numbering AS [Stockcard Numbering] " & _
        "FROM Products RIGHT JOIN Entrepots ON Products.Entrepot_ID = " & _
        "Entrepots.Entrepot_ID WHERE (Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = '" & Trim(txtEntrepotNum.Text) & "' " & _
        "AND Products.Prod_ID = " & jgxStock.Value(jgxStock.Columns("Prod_ID").Index)
        
    ADORecordsetOpen strSQL, m_conSADBEL, rstStartingNo, adOpenKeyset, adLockOptimistic
    'rstStartingNo.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If Not (rstStartingNo.BOF And rstStartingNo.EOF) Then
        rstStartingNo.MoveFirst
        
        strStartingNum = rstStartingNo.Fields("Starting Num").Value
        bytNumbering = rstStartingNo.Fields("Stockcard Numbering").Value
    End If
    
    ADORecordsetClose rstStartingNo
End Sub

Private Function NewStockcardNo(Optional strStockcardNo As String) As String
    Dim strSQL As String
    Dim blnSafe As Boolean
    Dim strStockCardNoHigh As String
    Dim lngSafeLength As Long
    Dim rstStockCard As ADODB.Recordset
    
    If strStockcardNo = "" Then
        Select Case bytNumbering
            Case 0
                
                blnSafe = False
                strSQL = "SELECT SC.Stock_ID AS [Stock ID], SC.Stock_Card_Num AS [Stock Card No], " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 9 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '19', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 99 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '29', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '39', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 9999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '49', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 99999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '59', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '69', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 9999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '79', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 99999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '89', " & _
                          "IIF(VAL(SC.Stock_Card_Num) = 999999999 AND LEN(SC.Stock_Card_Num) = LEN(VAL(SC.Stock_Card_Num)), '99', " & _
                          "LEN(SC.Stock_Card_Num))))))))))  AS [Length] " & _
                          "FROM (Entrepots [E] INNER JOIN (StockCards [SC] INNER JOIN Products [P] " & _
                          "ON SC.Prod_ID = P.Prod_ID) ON E.Entrepot_ID = P.Entrepot_ID) " & _
                          "WHERE E.Entrepot_ID = " & GetEntrepot_ID(Trim(txtEntrepotNum.Text), m_conSADBEL) & " " & _
                          "ORDER BY LEN(SC.Stock_Card_Num), SC.Stock_Card_Num"
                ADORecordsetOpen strSQL, m_conSADBEL, rstStockCard, adOpenKeyset, adLockOptimistic
                'rstStockCard.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
                
                GetStartingNo
                
                lngSafeLength = Len(strStartingNum)
                
                With rstStockCard
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
                
                    If Not (rstStockCard) Is Nothing Then .Filter = "[Length] = " & lngSafeLength
                    
                    If Not (.BOF And .EOF) Then
                        If UsedExistingStockcards(rstStockCard, strStartingNum) = True Then
                            strStockCardNoHigh = HighestStartingNumber(rstStockCard, strStartingNum)
                        Else
                            If UsedExistingStockcards(rstStockCard, strStartingNum, True) = True Then
                                strStockCardNoHigh = HighestStartingNumber(rstStockCard, strStartingNum, True)
                            Else
                                strStockCardNoHigh = strStartingNum
                            End If
                        End If
                    Else
                        If lngSafeLength = Len(strStartingNum) Then
                            strStockCardNoHigh = strStartingNum
                        ElseIf lngSafeLength >= 10 Then
                            strStockCardNoHigh = Empty
                        Else
                            strStockCardNoHigh = "1" & String$(lngSafeLength - 1, "0")
                        End If
                    End If
                End With
                
                ADORecordsetClose rstStockCard
                
                NewStockcardNo = strStockCardNoHigh
            Case 1
                NewStockcardNo = Empty
        End Select
    Else
        NewStockcardNo = strStockcardNo
    End If

    Call SaveStockCards(NewStockcardNo)
    
End Function

Private Sub pckInboundDocsPicklist_BtnClick(Record As PCubeLibPick.CRecord, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Select Case Button
        Case cpiAdd, cpiModify, cpiCopy
            frmInboundDoc.MyLoad m_conSADBEL, Record.RecordSource, Button, Cancel, 1, lngLastSeqNumber, pckInboundDocsPicklist, m_strLanguage

    End Select
    
End Sub

Private Function GetLastSeqNum(ByVal strEntrepotNum As String, connSadbel As ADODB.Connection) As Long
    Dim rstTmp As ADODB.Recordset
    
    ADORecordsetOpen "Select Entrepot_LastSeqNum from Entrepots where Entrepot_Type & '-' & Entrepot_Num = '" & strEntrepotNum & "'", connSadbel, rstTmp, adOpenKeyset, adLockOptimistic
    'rstTmp.Open "Select Entrepot_LastSeqNum from Entrepots where Entrepot_Type & '-' & Entrepot_Num = '" & strEntrepotNum & "'", connSadbel, adOpenForwardOnly, adLockReadOnly
    
    If Not (rstTmp.BOF And rstTmp.EOF) Then
        rstTmp.MoveFirst
        If IsNull(rstTmp!Entrepot_LastSeqNum) Then
            GetLastSeqNum = 0
        Else
            GetLastSeqNum = rstTmp!Entrepot_LastSeqNum
        End If
    Else
        GetLastSeqNum = 0
    End If
    
    ADORecordsetClose rstTmp
End Function

Private Sub UpdateLastSeqNum(ByVal strEntrepotNum As String, ByVal connSadbel As ADODB.Connection, ByVal lngLastSeqNum As Long)

    Dim rstTmp As ADODB.Recordset
        
    ADORecordsetOpen "Select Entrepot_LastSeqNum from Entrepots where Entrepot_Type & '-' & Entrepot_Num = '" & strEntrepotNum & "'", connSadbel, rstTmp, adOpenKeyset, adLockOptimistic
    'rstTmp.Open "Select Entrepot_LastSeqNum from Entrepots where Entrepot_Type & '-' & Entrepot_Num = '" & strEntrepotNum & "'", connSadbel, adOpenKeyset, adLockPessimistic
    
    If Not (rstTmp.BOF And rstTmp.EOF) Then
        rstTmp.MoveFirst
        
        rstTmp!Entrepot_LastSeqNum = lngLastSeqNum
        rstTmp.Update
    End If
    
    ADORecordsetClose rstTmp
End Sub

Private Function NewSequential(ByVal ProdNum As String) As Boolean
    
    Dim rstTemp As ADODB.Recordset
    
    ADORecordsetOpen "SELECT Entrepots!Entrepot_StockCard_Numbering AS Sequential " & _
                "FROM Entrepots WHERE Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num = '" & _
                txtEntrepotNum.Text & "'", m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open "SELECT Entrepots!Entrepot_StockCard_Numbering AS Sequential " & _
                "FROM Entrepots WHERE Entrepots!Entrepot_Type & '-' & Entrepots!Entrepot_Num = '" & _
                txtEntrepotNum.Text & "'", m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If Not (rstTemp.EOF And rstTemp.BOF) Then
    
        rstTemp.MoveFirst
        
        If IIf(IsNull(rstTemp!Sequential), 0, rstTemp!Sequential) <> 0 Then
            
            ADORecordsetClose rstTemp
            
            NewSequential = False
            
            Exit Function
        End If
    End If
    
    ADORecordsetClose rstTemp
    
    ADORecordsetOpen "SELECT StockCards!Stock_ID AS Stock_ID FROM StockCards INNER JOIN " & _
                "Products ON StockCards.Prod_ID = Products.Prod_ID WHERE Products!Prod_Num = '" & _
                ProdNum & "'", m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open "SELECT StockCards!Stock_ID AS Stock_ID FROM StockCards INNER JOIN " & _
                "Products ON StockCards.Prod_ID = Products.Prod_ID WHERE Products!Prod_Num = '" & _
                ProdNum & "'", m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If Not (rstTemp.EOF And rstTemp.BOF) Then
        NewSequential = False
    Else
        NewSequential = True
    End If
    
    ADORecordsetClose rstTemp

End Function

Private Sub FillQuantity(ByVal Handling As Integer, ByVal Qty As Double, ByVal Gross As Double, ByVal Net As Double, ByVal Available As Double)

    Dim dblTemp As Double
    
    Select Case Handling
    
        Case 0
            
            dblTemp = Net / Qty * Available
            strNetWt = Replace(CStr(Round(dblTemp, 3) + IIf(Round(dblTemp, 3) < dblTemp, 0.001, 0)), ",", ".")
            dblTemp = Gross / Qty * Available
            strGrossWt = Replace(CStr(Round(dblTemp, 2) + IIf(Round(dblTemp, 2) < dblTemp, 0.01, 0)), ",", ".")
            strQuantity = Replace(CStr(Available), ",", ".")
            
        Case 1
            
            dblTemp = Qty / Gross * Available
            strQuantity = Round(dblTemp, 0) + IIf(Round(dblTemp, 0) < dblTemp, 1, 0)
            dblTemp = Net / Gross * Available
            strNetWt = Replace(CStr(Round(dblTemp, 3) + IIf(Round(dblTemp, 3) < dblTemp, 0.001, 0)), ",", ".")
            strGrossWt = Replace(CStr(Available), ",", ".")
            
        Case 2
                            
            dblTemp = Qty / Net * Available
            strQuantity = Round(dblTemp, 0) + IIf(Round(dblTemp, 0) < dblTemp, 1, 0)
            dblTemp = Gross / Net * Available
            strGrossWt = Replace(CStr(Round(dblTemp, 2) + IIf(Round(dblTemp, 2) < dblTemp, 0.01, 0)), ",", ".")
            strNetWt = Replace(CStr(Available), ",", ".")
            
    End Select
        
End Sub

Private Function StockPick() As Boolean

    Dim clsStockProd As cStockProd
    Dim strStock As String
    
    strStock = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", jgxStock.Value(jgxStock.Columns("Stock Card No").Index))
    Set clsStockProd = New cStockProd
    
    With clsStockProd
        .Entrepot_Num = txtEntrepotNum.Text
        .StockCardNo = strStock
        .Product_ID = jgxStock.Value(jgxStock.Columns("Prod_ID").Index)
        .ProductNo = jgxStock.Value(jgxStock.Columns("Product Number").Index)
        .Stock_ID = jgxStock.Value(jgxStock.Columns("Stock_ID").Index)
        .StockCardNo = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", jgxStock.Value(jgxStock.Columns("Stock Card No").Index))
        
        With frmStockProdPicklist
            .fraProduct.Visible = False
            .lblJobNo.Visible = False
            .lblBatchNo.Visible = False
            .txtJobNo.Visible = False
            .txtBatchNo.Visible = False
            .Caption = "Stock Cards"
            .fraStockCard.Left = 120
            .cmdOK.Left = 3000
            .cmdCancel.Left = .cmdOK.Left + 1320
            .Width = 5745
        End With
        
        .ShowPicklist Me, m_conSADBEL, m_conTaric, m_strLanguage, m_intTaricProperties, ResourceHandler, True, True
                
        If .Cancel = False Then
            jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = .StockCardNo
            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = .Stock_ID
            lngStockID = .Stock_ID
        End If
        
        StockPick = Not .Cancel
        
    End With
        
    Set clsStockProd = Nothing

End Function

Private Sub SaveStockCards(NewStockcardNo As String)
    Dim rstStockCard As ADODB.Recordset
    Dim strSQL As String
    Dim lngDifference As Long
    Dim lngCounter As Long
    
    'Save the newly added stockcard/s into the database.
    strSQL = "SELECT Stock_ID AS [Stock ID], Stock_Card_Num AS [Stock Card No], Prod_ID AS [Product ID] FROM [StockCards] "
         
    ADORecordsetOpen strSQL, m_conSADBEL, rstStockCard, adOpenKeyset, adLockOptimistic
    'rstStockCard.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    With rstStockCard
        .AddNew
        .Fields("Stock Card No").Value = NewStockcardNo
        .Fields("Product ID").Value = jgxStock.Value(jgxStock.Columns("Prod_ID").Index)
        'lngStockID = .Fields("Stock ID").Value
        .Update
        
        lngStockID = InsertRecordset(m_conSADBEL, rstStockCard, "StockCards")
    End With
    
    ADORecordsetClose rstStockCard
End Sub

Private Function UniqueStockcard(IsToBeSaved As Boolean, lngProdID As Long, strStockcardNo As String) As Boolean
    Dim rstStock As ADODB.Recordset
    Dim strSQlStock As String
    
        strSQlStock = "Select STOCKCARDS.Stock_ID As [Stock_ID], STOCKCARDS.Stock_Card_Num As [Stockcard] FROM STOCKCARDS "
        strSQlStock = strSQlStock & " INNER JOIN (Products INNER JOIN Entrepots ON "
        strSQlStock = strSQlStock & " Products.Entrepot_ID = Entrepots.Entrepot_ID) ON "
        strSQlStock = strSQlStock & " Stockcards.Prod_ID = Products.Prod_ID "
        strSQlStock = strSQlStock & " WHERE Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '"
        strSQlStock = strSQlStock & txtEntrepotNum.Text & "' AND STOCKCARDS.Stock_Card_Num = '" & strStockcardNo & "'"
    ADORecordsetOpen strSQlStock, m_conSADBEL, rstStock, adOpenKeyset, adLockOptimistic
    'rstStock.Open strSQlStock, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If rstStock.EOF And rstStock.BOF Then
        If IsToBeSaved = True Then
            UniqueStockcard = True
            SaveStockCards (strStockcardNo)
        Else
            UniqueStockcard = True
        End If
    Else
        If StockcardExistsInProduct(lngProdID, strStockcardNo) = True Then
            UniqueStockcard = False
            lngStockID = rstStock.Fields("Stock_ID").Value
        Else
            UniqueStockcard = False
            lngStockID = 0
            If IsToBeSaved = True Then
                Call StockPick
            End If
        End If
    
    End If
    
    ADORecordsetClose rstStock
End Function

Private Sub pckInboundDocsPicklist_TempIDChange(GenerateNew As Boolean, ByVal NewID As Long)
    Dim rstTmp As ADODB.Recordset
    
    Dim conHistory As ADODB.Connection
    Dim conData As ADODB.Connection
    Dim conSADBEL As ADODB.Connection
                
    ADOConnectDB conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    ADOConnectDB conData, g_objDataSourceProperties, DBInstanceType_DATABASE_DATA
    
    'Check if the database exists. If it does not, create a new one. 1/5/05 rac
    If Len(Trim(Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(Year(Now), 2) & ".mdb"))) = 0 Then
        CreateHistoryMdb conSADBEL, conData, Right(Year(Now), 2)
    End If
    
    ADODisconnectDB conSADBEL
    ADODisconnectDB conData
    
    ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(Year(Now), 2)
    
    ADORecordsetOpen "Select InDoc_ID from InboundDocs where InDoc_ID = " & NewID, conHistory, rstTmp, adOpenKeyset, adLockOptimistic
    'rstTmp.Open "Select InDoc_ID from InboundDocs where InDoc_ID = " & NewID, conHistory, adOpenForwardOnly, adLockReadOnly
    
    If Not (rstTmp.BOF And rstTmp.EOF) Then
        GenerateNew = True
    Else
        GenerateNew = False
    End If
    
    ADORecordsetClose rstTmp
    
    ADODisconnectDB conHistory
End Sub

Private Sub AddMissingInboundDocsInHistory(ByVal SADBELDB As ADODB.Connection)
    Dim rstSadbel As ADODB.Recordset
    Dim rstHistory As ADODB.Recordset
    Dim conHistory As ADODB.Connection
    Dim fld As ADODB.Field

    Dim conData As ADODB.Connection
    Dim conSADBEL As ADODB.Connection
                
    ADOConnectDB conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    ADOConnectDB conData, g_objDataSourceProperties, DBInstanceType_DATABASE_DATA
    
    'Check if the database exists. If it does not, create a new one. 1/5/05 rac
    If Len(Trim(Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(Year(Now), 2) & ".mdb"))) = 0 Then
        CreateHistoryMdb conSADBEL, conData, Right(Year(Now), 2)
    End If
    
    ADODisconnectDB conSADBEL
    ADODisconnectDB conData
    
    ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(Year(Now), 2)
    
    SADBELDB.Close
    SADBELDB.Open
    
    ADORecordsetOpen "Select * from InboundDocs where InDoc_Global = -1", SADBELDB, rstSadbel, adOpenKeyset, adLockOptimistic
    'rstSadbel.Open "Select * from InboundDocs where InDoc_Global = -1", SADBELDB, adOpenKeyset, adLockReadOnly
    
    ADORecordsetOpen "Select * from InboundDocs where InDoc_Global = -1", conHistory, rstHistory, adOpenKeyset, adLockPessimistic
    'rstHistory.Open "Select * from InboundDocs where InDoc_Global = -1", conHistory, adOpenKeyset, adLockPessimistic
    
    Do While Not rstSadbel.EOF
        If Not (rstHistory.BOF And rstHistory.EOF) Then rstHistory.MoveFirst
        
        rstHistory.Find "InDoc_ID =" & rstSadbel.Fields("InDoc_ID").Value
        
        If rstHistory.EOF Then
            rstHistory.AddNew
            
            For Each fld In rstSadbel.Fields
                rstHistory.Fields(fld.Name).Value = fld.Value
            Next
            rstHistory.Update
            
            InsertRecordset conHistory, rstHistory, "InboundDocs"
        Else
            
            For Each fld In rstSadbel.Fields
                If UCase(fld.Name) <> "INDOC_ID" Then
                    rstHistory.Fields(fld.Name).Value = fld.Value
                End If
            Next
            rstHistory.Update
            
            UpdateRecordset conHistory, rstHistory, "InboundDocs"
        End If
        
        rstSadbel.MoveNext
    Loop
    
    ADORecordsetClose rstSadbel
    ADORecordsetClose rstHistory
    
    ADODisconnectDB conHistory
End Sub

Public Sub UpdateTables()
    Dim i As Long
    Dim rstInbound As ADODB.Recordset
'    Dim rstInboundHistory As ADODB.Recordset
    Dim rstInboundDocs As ADODB.Recordset
'    Dim conHistory As ADODB.Connection
    Dim strSQL As String
    Dim lngInDoc_ID As Long
    Dim dblDiff As Double
    Dim dblMultiplier As Double
    Dim dblOrigQty As Double
    Dim dblOrigGross As Double
    Dim dblOrigNet As Double
    
    Dim dblNewAvlQtyWgt As Double
    Dim dblNewOrigQty As Double
    Dim dblNewOrigGross As Double
    Dim dblNewOrigNet As Double
    Dim blnDeleteRecord As Boolean
    
    Dim blnAddNegativeRecord As Boolean
        
    jgxStock.Update
    
    Me.MousePointer = vbHourglass
        
    If Not (m_rstStocksOff.BOF And m_rstStocksOff.EOF) Then
    
        m_rstStocksOff.MoveFirst
        
        lngInDoc_ID = GetInDoc_ID
        
        '================== start updating inbounds table =====================================
        ADORecordsetOpen "Select * from Inbounds", m_conSADBEL, rstInbound, adOpenKeyset, adLockOptimistic
        'rstInbound.Open "Select * from Inbounds", m_conSADBEL, adOpenKeyset, adLockPessimistic
    
        With rstInbound
    
            For i = 1 To m_rstStocksOff.RecordCount
                blnDeleteRecord = False
                blnAddNegativeRecord = False
                
                If Not (.BOF And .EOF) Then .MoveFirst

                .Find "In_ID = " & m_rstStocksOff!In_ID
                
                If Not rstInbound.EOF Then   'kapag nakita ang record
                
                    If m_rstStocksOff!NonEdit = -1 Then
                        Select Case m_rstStocksOff!Handling
                        
                            Case 0  ' "Number of Items"
                                If CDbl(Val(m_rstStocksOff![Num of Items])) <> CDbl(rstInbound!In_Avl_Qty_Wgt) Then
                                    If CDbl(Val(m_rstStocksOff![Num of Items])) = 0 Then
                                        blnDeleteRecord = True
                                    End If
                                
                                    blnAddNegativeRecord = True
                                    If !In_Avl_Qty_Wgt <> !In_Orig_Packages_Qty Then    'means nabawasan na
                                        dblDiff = !In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Num of Items]))
                                        
                                        dblOrigQty = !In_Orig_Packages_Qty - dblDiff
                                        dblOrigGross = !In_Orig_Gross_Weight - (dblDiff * (!In_Orig_Gross_Weight / !In_Orig_Packages_Qty))
                                        dblOrigNet = !In_Orig_Net_Weight - (dblDiff * (!In_Orig_Net_Weight / !In_Orig_Packages_Qty))
                                        
                                        
                                    End If
                                    dblMultiplier = CDbl(Val(m_rstStocksOff![Num of Items])) / CDbl(rstInbound!In_Avl_Qty_Wgt)
                                    
                                    dblNewAvlQtyWgt = Abs(!In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Num of Items])))
                                    
                                    dblNewOrigGross = Abs(dblDiff * (!In_Orig_Gross_Weight / !In_Orig_Packages_Qty))
                                    dblNewOrigNet = Abs(dblDiff * (!In_Orig_Net_Weight / !In_Orig_Packages_Qty))
                                    dblNewOrigQty = dblNewAvlQtyWgt
                                    
                                    !In_Avl_Qty_Wgt = CDbl(Val(m_rstStocksOff![Num of Items]))
                                    
                                    !In_Orig_Packages_Qty = Round(dblOrigQty, 0)
                                    !In_Orig_Gross_Weight = Round(dblOrigGross, 2)
                                    !In_Orig_Net_Weight = Round(dblOrigNet, 3)
                                    
'                                    CopyToHistory rstInbound, rstInboundHistory, 0
                                    CopyToHistory rstInbound, 0
                                    
                                End If
                                
                            Case 1  '"Gross Weight"
                                If CDbl(Val(m_rstStocksOff![Gross Weight])) <> rstInbound!In_Avl_Qty_Wgt Then
                                    If CDbl(Val(m_rstStocksOff![Gross Weight])) = 0 Then
                                        blnDeleteRecord = True
                                    End If

                                    blnAddNegativeRecord = True
                                    If !In_Avl_Qty_Wgt <> !In_Orig_Gross_Weight Then    'means nabawasan na

                                        dblDiff = !In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Gross Weight]))

                                        dblOrigGross = !In_Orig_Gross_Weight - dblDiff
                                        dblOrigQty = !In_Orig_Packages_Qty - (dblDiff * (!In_Orig_Packages_Qty / !In_Orig_Gross_Weight))
                                        dblOrigNet = !In_Orig_Net_Weight - (dblDiff * (!In_Orig_Net_Weight / !In_Orig_Gross_Weight))
                                    
                                    End If

                                    dblMultiplier = CDbl(Val(m_rstStocksOff![Gross Weight])) / CDbl(rstInbound!In_Avl_Qty_Wgt)
                                    dblNewAvlQtyWgt = Abs(!In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Gross Weight])))
                                    
                                    dblNewOrigNet = Abs(dblDiff * (!In_Orig_Net_Weight / !In_Orig_Gross_Weight))
                                    dblNewOrigQty = Abs(dblDiff * (!In_Orig_Packages_Qty / !In_Orig_Gross_Weight))
                                    dblNewOrigGross = dblNewAvlQtyWgt
                                    !In_Avl_Qty_Wgt = CDbl(Val(m_rstStocksOff![Gross Weight]))
                                    
                                    !In_Orig_Packages_Qty = Round(dblOrigQty, 0)
                                    !In_Orig_Gross_Weight = Round(dblOrigGross, 2)
                                    !In_Orig_Net_Weight = Round(dblOrigNet, 3)
                                    
                                    CopyToHistory rstInbound, 0
                                    
                                End If
                            
                            Case 2  '"Net Weight"
                                If CDbl(Val(m_rstStocksOff![Net Weight])) <> rstInbound!In_Avl_Qty_Wgt Then
                                    If CDbl(Val(m_rstStocksOff![Net Weight])) = 0 Then
                                        blnDeleteRecord = True
                                    End If

                                    blnAddNegativeRecord = True
                                    If !In_Avl_Qty_Wgt <> !In_Orig_Net_Weight Then    'means nabawasan na
                                        dblDiff = !In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Net Weight]))

                                        dblOrigNet = !In_Orig_Net_Weight - dblDiff
                                        dblOrigQty = !In_Orig_Packages_Qty - (dblDiff * (!In_Orig_Packages_Qty / !In_Orig_Net_Weight))
                                        dblOrigGross = !In_Orig_Gross_Weight - (dblDiff * (!In_Orig_Gross_Weight / !In_Orig_Net_Weight))
                                        
                                    End If
                                    dblMultiplier = CDbl(Val(m_rstStocksOff![Net Weight])) / CDbl(rstInbound!In_Avl_Qty_Wgt)
                                    
                                    dblNewAvlQtyWgt = Abs(!In_Avl_Qty_Wgt - CDbl(Val(m_rstStocksOff![Net Weight])))
                                    
                                    
                                    dblNewOrigGross = Abs(dblDiff * (!In_Orig_Gross_Weight / !In_Orig_Net_Weight))
                                    dblNewOrigQty = Abs(dblDiff * (!In_Orig_Packages_Qty / !In_Orig_Net_Weight))
                                    dblNewOrigNet = dblNewAvlQtyWgt
                                    
                                    !In_Avl_Qty_Wgt = CDbl(Val(m_rstStocksOff![Net Weight]))
                                    !In_Orig_Net_Weight = Round(dblOrigNet, 3)
                                    !In_Orig_Gross_Weight = Round(dblOrigGross, 2)
                                    !In_Orig_Packages_Qty = Round(dblOrigQty, 0)
                                    
                                    CopyToHistory rstInbound, 0
                                    
                                End If
                            
                        End Select
                        
                        If blnDeleteRecord = False Then

                            !In_Batch_Num = IIf(IsNull(m_rstStocksOff![Batch Num]), "", m_rstStocksOff![Batch Num])
                            !In_Job_Num = IIf(IsNull(m_rstStocksOff![Job Num]), "", m_rstStocksOff![Job Num])
                            .Update
                            
                            UpdateRecordset m_conSADBEL, rstInbound, "Inbounds"
                        Else
                            .Delete
                            
                            ExecuteNonQuery m_conSADBEL, "DELETE * FROM [INBOUNDS] WHERE [IN_ID] = " & m_rstStocksOff!In_ID
                        End If
                        
                        
                        '========================================== Add negative record ===================================
                        If blnAddNegativeRecord = True Then
                            .AddNew
                            
                            !In_Batch_Num = IIf(IsNull(m_rstStocksOff![Batch Num]), "", m_rstStocksOff![Batch Num])
                            !In_Job_Num = IIf(IsNull(m_rstStocksOff![Job Num]), "", m_rstStocksOff![Job Num])
                            
                            !In_Orig_Packages_Qty = Round(dblNewOrigQty, 0)
                            !In_Orig_Gross_Weight = Round(dblNewOrigGross, 2)
                            !In_Orig_Net_Weight = Round(dblNewOrigGross, 3)
                            
                            !In_Orig_Packages_Type = IIf(IsNull(m_rstStocksOff![Package Type]), "", m_rstStocksOff![Package Type])
                            
                            !In_Avl_Qty_Wgt = Round(Abs(dblNewAvlQtyWgt) * -1, 0)
                            !Stock_ID = m_rstStocksOff![Stock_ID]
                            
                            !InDoc_ID = lngInDoc_ID
                            
                            'Glenn 3/30/2006
                            If !InDoc_ID = 0 Then
                                MsgBox "Failed adding inbound document properly. Please contact your administrator.", vbInformation, "Initial Stock"
                            End If
        
                            CopyToHistory rstInbound, 1
                            
                            .Update
                            
                            InsertRecordset m_conSADBEL, rstInbound, "Inbounds"
                        End If
                        '===================================================================================================

                    Else    'meaning existing ang record pero hindi pa naeedit
                                            
                        
                        Select Case m_rstStocksOff!Handling
                        
                            Case 0  '"Number of Items"
                                If Round(CDbl(Val(m_rstStocksOff![Num of Items])), 0) = 0 Then
                                    blnDeleteRecord = True
                                End If
                            Case 1  '"Gross Weight"
                                If Round(CDbl(Val(m_rstStocksOff![Gross Weight])), 2) = 0 Then
                                    blnDeleteRecord = True
                                End If
                            Case 2  '"Net Weight"
                                If Round(CDbl(Val(m_rstStocksOff![Net Weight])), 3) = 0 Then
                                     blnDeleteRecord = True
                                End If
                               
                        End Select
                        
                        If blnDeleteRecord Then
                            
                            'if record is for deletion then don't update mdb_sadbel...update only mdb_history
                            UpdateHistory rstInbound!In_ID
                            
                            .Delete
                            
                            ExecuteNonQuery m_conSADBEL, "DELETE * FROM [INBOUNDS] WHERE [IN_ID] = " & m_rstStocksOff!In_ID
                            
                        Else
                            !In_Batch_Num = IIf(IsNull(m_rstStocksOff![Batch Num]), "", m_rstStocksOff![Batch Num])
                            !In_Job_Num = IIf(IsNull(m_rstStocksOff![Job Num]), "", m_rstStocksOff![Job Num])
                            
                            !In_Orig_Packages_Qty = Round(CDbl(Val(m_rstStocksOff![Num of Items])), 0)
                            !In_Orig_Gross_Weight = Round(CDbl(Val(m_rstStocksOff![Gross Weight])), 2)
                            !In_Orig_Net_Weight = Round(CDbl(Val(m_rstStocksOff![Net Weight])), 3)
                            !In_Orig_Packages_Type = m_rstStocksOff![Package Type]
                            !Stock_ID = m_rstStocksOff![Stock_ID]
                            
                            Select Case m_rstStocksOff!Handling
                                Case 0  '"Number of Items"
                                    !In_Avl_Qty_Wgt = !In_Orig_Packages_Qty
                                Case 1  '"Gross Weight"
                                    !In_Avl_Qty_Wgt = !In_Orig_Gross_Weight
                                Case 2  '"Net Weight"
                                   !In_Avl_Qty_Wgt = !In_Orig_Net_Weight
                                   
                            End Select
                            
                            CopyToHistory rstInbound, 0
                            
                            .Update
                            
                            UpdateRecordset m_conSADBEL, rstInbound, "Inbounds"
                            
                        End If
                        
                    End If
                    
                    
                        
                Else 'record is not found
                
                    Select Case m_rstStocksOff!Handling
                        Case 0  '"Number of Items"
                            If Round(CDbl(Val(m_rstStocksOff![Num of Items])), 0) = 0 Then
                                blnDeleteRecord = True
                            End If
                        Case 1  '"Gross Weight"
                            If Round(CDbl(Val(m_rstStocksOff![Gross Weight])), 2) = 0 Then
                                blnDeleteRecord = True
                            End If
                        Case 2  '"Net Weight"
                            If Round(CDbl(Val(m_rstStocksOff![Net Weight])), 3) = 0 Then
                                 blnDeleteRecord = True
                            End If
                           
                    End Select
                    
                                        
                    If blnDeleteRecord Then
                    
                        'if record is for deletion then don't update mdb_sadbel...update only mdb_history
'                        UpdateHistory rstInbound!In_ID
                        
                    Else
                                        
                        .AddNew
                        !In_Batch_Num = IIf(IsNull(m_rstStocksOff![Batch Num]), "", m_rstStocksOff![Batch Num])
                        !In_Job_Num = IIf(IsNull(m_rstStocksOff![Job Num]), "", m_rstStocksOff![Job Num])
                        
                        !In_Orig_Packages_Qty = Round(CDbl(Val(m_rstStocksOff![Num of Items])), 0)
                        
                        !In_Orig_Gross_Weight = Round(CDbl(Val(m_rstStocksOff![Gross Weight])), 2)
                        
                        !In_Orig_Net_Weight = Round(CDbl(Val(m_rstStocksOff![Net Weight])), 3)
                        !In_Orig_Packages_Type = m_rstStocksOff![Package Type]
                        
                        !In_TotalOut_Qty_Wgt = 0
                        !In_Reserved_Qty_Wgt = 0
                            
                        Select Case m_rstStocksOff!Handling
                            Case 0  '"Number of Items"
                                !In_Avl_Qty_Wgt = !In_Orig_Packages_Qty
                            Case 1  '"Gross Weight"
                                !In_Avl_Qty_Wgt = !In_Orig_Gross_Weight
                            Case 2  '"Net Weight"
                               !In_Avl_Qty_Wgt = !In_Orig_Net_Weight
                               
                        End Select
                        
                        !Stock_ID = m_rstStocksOff!Stock_ID
                        !InDoc_ID = lngInDoc_ID
                        
                        'Glenn 3/30/2006
                        If !InDoc_ID = 0 Then
                            MsgBox "Failed adding inbound document properly. Please contact your administrator.", vbInformation, "Initial Stock"
                        End If
        
                        CopyToHistory rstInbound, 1
                        
                        .Update
                        
                        m_rstStocksOff!In_ID = InsertRecordset(m_conSADBEL, rstInbound, "Inbounds")
                        
                    End If
                    
                End If
                
                m_rstStocksOff.MoveNext
            Next
            
                
        End With
        
        ADORecordsetClose rstInbound
    
        Me.MousePointer = vbNormal
        
    End If
    
    
    'delete from Inbounds
    If m_alngDeleted(0) <> 0 Then
        '================== start updating inbounds table =====================================
        
        ADORecordsetOpen "Select * from Inbounds", m_conSADBEL, rstInbound, adOpenKeyset, adLockOptimistic
    
        For i = 0 To UBound(m_alngDeleted)
            If Not (rstInbound.BOF And rstInbound.EOF) Then
                rstInbound.MoveFirst
            End If
            rstInbound.Find "In_ID = " & m_alngDeleted(i)
            
            If Not rstInbound.EOF Then
                UpdateHistory rstInbound!In_ID, True
                
                rstInbound.Delete
                
                ExecuteNonQuery m_conSADBEL, "DELETE * FROM [Inbounds] WHERE [In_ID] = " & m_alngDeleted(i)
            End If
        Next
        
        Erase m_alngDeleted
        ReDim m_alngDeleted(0)
    End If
    
    If m_alngDeleted(0) <> 0 Or m_rstStocksOff.RecordCount > 0 Then
        DeleteInInboundDocs
    End If
    
    Me.MousePointer = vbNormal
    
End Sub

Private Sub UpdateHistory(ByVal In_ID As Long, Optional ByVal blnDeleted As Boolean = False)
            
    Dim conHistory As ADODB.Connection
    Dim rstInboundHistory As ADODB.Recordset
    Dim blnInboundFound As Boolean
    Dim strDB As String
    Dim strDBHistory As String
    Dim strSQL As String
    
    blnInboundFound = False
    
    strDB = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_history" & Right(Year(Date), 2) & ".mdb")
    
    If strDB <> "" Then
        
        ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(strDB)
        'OpenADODatabase conHistory, NoBackSlash(g_objDataSourceProperties.TracefilePath), strDB
        
            strSQL = "SELECT * FROM Inbounds WHERE In_ID = " & In_ID
        ADORecordsetOpen strSQL, conHistory, rstInboundHistory, adOpenKeyset, adLockOptimistic
        'rstInboundHistory.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
        
        If Not (rstInboundHistory.EOF And rstInboundHistory.BOF) Then
        
            blnInboundFound = True
        Else
            ADORecordsetClose rstInboundHistory
        End If
    End If
    
    If blnInboundFound = False Then
    
        strDBHistory = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_history??.mdb")
        Do While blnInboundFound = False And strDBHistory <> ""
            If conHistory.State = adStateOpen Then
                conHistory.Close
            End If
            If strDBHistory <> strDB Then
                
                ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(strDBHistory)
                'OpenADODatabase conHistory, NoBackSlash(g_objDataSourceProperties.TracefilePath), strDBHistory

                ADORecordsetOpen strSQL, conHistory, rstInboundHistory, adOpenKeyset, adLockOptimistic
                'rstInboundHistory.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
                
                If Not (rstInboundHistory.EOF And rstInboundHistory.BOF) Then
                
                    blnInboundFound = True
                    Exit Do
                Else
                    ADORecordsetClose rstInboundHistory
                End If
                
            End If
            strDBHistory = Dir()
        Loop
    End If
    
    With rstInboundHistory
        
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            If blnDeleted = False Then
                !In_Batch_Num = IIf(IsNull(m_rstStocksOff![Batch Num]), "", m_rstStocksOff![Batch Num])
                
                !In_Job_Num = IIf(IsNull(m_rstStocksOff![Job Num]), "", m_rstStocksOff![Job Num])
                !In_Orig_Packages_Qty = Round(CDbl(Val(m_rstStocksOff![Num of Items])), 0)
                !In_Orig_Gross_Weight = Round(CDbl(Val(m_rstStocksOff![Gross Weight])), 2)
                !In_Orig_Net_Weight = Round(CDbl(Val(m_rstStocksOff![Net Weight])), 3)
                !In_Orig_Packages_Type = m_rstStocksOff![Package Type]
                
                !Stock_ID = m_rstStocksOff![Stock_ID]
                
                Select Case m_rstStocksOff!Handling
                    Case 0  '"Number of Items"
                        !In_Avl_Qty_Wgt = !In_Orig_Packages_Qty
                    Case 1  '"Gross Weight"
                        !In_Avl_Qty_Wgt = !In_Orig_Gross_Weight
                    Case 2  '"Net Weight"
                       !In_Avl_Qty_Wgt = !In_Orig_Net_Weight
                       
                End Select

            Else
                !In_Orig_Packages_Qty = 0
                !In_Orig_Gross_Weight = 0
                !In_Orig_Net_Weight = 0
                !In_Avl_Qty_Wgt = 0
                
            End If

            .Update
            
            UpdateRecordset conHistory, rstInboundHistory, "Inbounds"
        End If
        
    End With
    
    ADORecordsetClose rstInboundHistory
    
    ADODisconnectDB conHistory

End Sub

Private Function UsedExistingStockcards(rstExistingStockcards As ADODB.Recordset, _
                                        strStartingNumber As String, _
                                        Optional blnRemoveFilter As Boolean) As Boolean
    If blnRemoveFilter = True Then
        rstExistingStockcards.Filter = adFilterNone
    End If
    
    With rstExistingStockcards
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
    
    If Len(strStartingNumber) < lngLength Then
        strStartingNumber = String$(lngLength - Len(strStartingNumber), "0") & strStartingNumber
    End If
    
    HighestStartingNumber = strStartingNumber
End Function

Private Function NotOkToSave() As Boolean
    
    Dim strInvalidData As String
    Dim lngTempID As Long
    Dim intHandling As Integer
    Dim strCode As String
    
    If jgxStock.Row > 0 Then
    
        If lngProd_ID = 0 Then
            
            strCode = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Product Number").Index)), "", jgxStock.Value(jgxStock.Columns("Product Number").Index))
            
            If ProductExist(strCode, intHandling) Then
                jgxStock.Value(jgxStock.Columns("Prod_ID").Index) = lngProd_ID
                jgxStock.Value(jgxStock.Columns("Handling").Index) = intHandling
            Else
                blnSystemChanged = True
                jgxStock.Col = jgxStock.Columns("Product Number").Index
                blnSystemChanged = False
                jgxStock.SetFocus
                jgxStock.EditMode = jgexEditModeOn
                jgxStock.SelStart = 0
                
                strInvalidData = strInvalidData & vbCrLf & "   * Product Number"
                NotOkToSave = True
            End If
        End If
                    
        If lngStockID = 0 Then
                            
            strCode = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", jgxStock.Value(jgxStock.Columns("Stock Card No").Index))
            
            If strCode <> "" Then
                If Not UniqueStockcard(False, jgxStock.Value(jgxStock.Columns("Prod_ID").Index), strCode) Then
                    If lngStockID = 0 Then
                        If Not NotOkToSave Then
                            blnSystemChanged = True
                            jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                            jgxStock.EditMode = jgexEditModeOn
                            jgxStock.SelStart = 0
                            blnSystemChanged = False
                            NotOkToSave = True
                        End If
                        strInvalidData = strInvalidData & vbCrLf & "   * Stock Card Number"
                        NotOkToSave = True
                    Else
                        jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
                    End If
                Else
                    If Not NotOkToSave Then
                        blnSystemChanged = True
                        jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                        jgxStock.EditMode = jgexEditModeOn
                        jgxStock.SelStart = 0
                        blnSystemChanged = False
                        NotOkToSave = True
                    End If
                    strInvalidData = strInvalidData & vbCrLf & "   * Stock Card Number"
                    NotOkToSave = True
                End If
            Else
                If Not NotOkToSave Then
                    blnSystemChanged = True
                    jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                    jgxStock.EditMode = jgexEditModeOn
                    jgxStock.SelStart = 0
                    blnSystemChanged = False
                    NotOkToSave = True
                End If
                strInvalidData = strInvalidData & vbCrLf & "   * Stock Card Number"
                NotOkToSave = True
            End If
        ElseIf lngStockID = -1 Then
            jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = NewStockcardNo
            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
        End If
                                                    
        If lngPack_Flag = 0 Then
            
            If Not ValidPackageType Then
            
                If Not NotOkToSave Then
                    jgxStock.Col = jgxStock.Columns("Package Type").Index
                    jgxStock.SetFocus
                    jgxStock.EditMode = jgexEditModeOn
                    jgxStock.SelStart = 0
                End If
                
                strInvalidData = strInvalidData & vbCrLf & "   * Package Type"
                NotOkToSave = True
                        
            Else
                jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 1
                lngPack_Flag = 1
            End If
        End If
        
    End If
    
    If NotOkToSave Then
        MsgBox "Please correct the following fields:" & Space(13) & vbCrLf & strInvalidData, vbInformation + vbOKOnly, "Initial Stocks"
    End If
    
End Function

Private Function IsStockClosed() As Boolean
    Dim rstStock As ADODB.Recordset
    Dim strSQL As String
    Dim strDate As String
    
        strSQL = "SELECT INDOC_DATE AS [DATE] FROM INBOUNDDOCS WHERE INDOC_NUM = '" & txtDocNum.Text & "' ORDER BY INDOC_DATE DESC"
    ADORecordsetOpen strSQL, m_conSADBEL, rstStock, adOpenKeyset, adLockOptimistic
    'rstStock.Open strSQL, m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If Not (rstStock.BOF And rstStock.EOF) Then
        strDate = rstStock.Fields("Date").Value
    End If
    
    ADORecordsetClose rstStock
    
        strSQL = "SELECT TOP 1 OUTDOC_DATE FROM OUTBOUNDDOCS INNER JOIN (OUTBOUNDS INNER JOIN (INBOUNDS INNER JOIN (STOCKCARDS INNER JOIN (PRODUCTS INNER JOIN ENTREPOTS ON PRODUCTS.ENTREPOT_ID = ENTREPOTS.ENTREPOT_ID) ON STOCKCARDS.PROD_ID = PRODUCTS.PROD_ID) ON INBOUNDS.STOCK_ID = STOCKCARDS.STOCK_ID) ON OUTBOUNDS.IN_ID = INBOUNDS.IN_ID) ON OUTBOUNDS.OUTDOC_ID = OUTBOUNDDOCS.OUTDOC_ID WHERE UCASE(RIGHT(OUT_CODE,11))= '<<CLOSURE>>' AND DATEVALUE(OUTDOC_DATE) >= DATEVALUE('" & strDate & "') AND ENTREPOTS.ENTREPOT_TYPE & '-' & ENTREPOTS.ENTREPOT_NUM ='" & txtEntrepotNum.Text & "'"
    ADORecordsetOpen strSQL, m_conSADBEL, rstStock, adOpenKeyset, adLockOptimistic
    'rstStock.Open strSQL, m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If Not (rstStock.BOF And rstStock.EOF) Then
        IsStockClosed = True
    Else
        IsStockClosed = False
    End If
    
    ADORecordsetClose rstStock
End Function

Private Sub SetIDs()
    
    lngStockID = jgxStock.Value(jgxStock.Columns("Stock_ID").Index)
    lngProd_ID = jgxStock.Value(jgxStock.Columns("Prod_ID").Index)
    lngPack_Flag = jgxStock.Value(jgxStock.Columns("Pack_Flag").Index)
    
End Sub

Private Sub CopyToHistory(ByVal rstSource As ADODB.Recordset, ByVal bytAction As Byte)
'bytAction: 0-> Edit, 1 -> AddNewRecord
    Dim fld As ADODB.Field
    Dim conHistory As ADODB.Connection
    Dim rstDestination As ADODB.Recordset
    Dim strDBHistory As String
    Dim strDB As String
    Dim strSQL As String
    Dim blnInboundFound As Boolean
    
    blnInboundFound = False
    
    strDB = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_history" & Right(Year(Date), 2) & ".mdb")
    
    If strDB <> "" Then
            
        ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(strDB)
        'OpenADODatabase conHistory, NoBackSlash(g_objDataSourceProperties.TracefilePath), strDB
        
            strSQL = "SELECT * FROM Inbounds WHERE In_ID = " & rstSource!In_ID
        ADORecordsetOpen strSQL, conHistory, rstDestination, adOpenKeyset, adLockOptimistic
        'rstDestination.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
        If Not (rstDestination.EOF And rstDestination.BOF) Or bytAction = 1 Then
            blnInboundFound = True
        Else
            ADORecordsetClose rstDestination
        End If
    End If
    
    If blnInboundFound = False Then
        strDBHistory = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\mdb_history??.mdb")
        Do While blnInboundFound = False And strDBHistory <> ""
            ADODisconnectDB conHistory
            
            If strDBHistory <> strDB Then
                
                ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(strDBHistory)
                'OpenADODatabase conHistory, NoBackSlash(g_objDataSourceProperties.TracefilePath), strDBHistory
                        
                ADORecordsetOpen strSQL, conHistory, rstDestination, adOpenKeyset, adLockOptimistic
                'rstDestination.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
                
                If Not (rstDestination.EOF And rstDestination.BOF) Or bytAction = 1 Then
                    blnInboundFound = True
                    Exit Do
                Else
                    ADORecordsetClose rstDestination
                End If
                
            End If
            strDBHistory = Dir()
        Loop
    End If
    
    If bytAction = 0 Then
        If Not rstDestination.EOF Then
            For Each fld In rstDestination.Fields
                If UCase(fld.Name) <> "IN_ID" Then
                    fld.Value = rstSource.Fields(fld.Name).Value
                End If
            Next
            
            rstDestination.Update
        End If
        
        UpdateRecordset conHistory, rstDestination, "Inbounds"
    Else
        rstDestination.AddNew
        For Each fld In rstDestination.Fields
            fld.Value = rstSource.Fields(fld.Name).Value
        Next
        
        rstDestination.Update
        
        InsertRecordset conHistory, rstDestination, "Inbounds"
    End If
    
    Set fld = Nothing
    
    ADORecordsetClose rstDestination
    
    ADODisconnectDB conHistory
End Sub

Private Function StockcardExistsInProduct(lngProductID As Long, strStockcardNo As String) As Boolean
    Dim strSQLOfStockcard As String
    Dim rstStockcardInProduct As ADODB.Recordset
    
        strSQLOfStockcard = " Select Stockcards.Stock_Card_Num As Stockcard FROM STOCKCARDS "
        strSQLOfStockcard = strSQLOfStockcard & " INNER JOIN PRODUCTS ON Stockcards.Prod_ID = "
        strSQLOfStockcard = strSQLOfStockcard & " Products.Prod_ID WHERE Products.Prod_ID = "
        strSQLOfStockcard = strSQLOfStockcard & lngProductID
        strSQLOfStockcard = strSQLOfStockcard & " AND Stockcards.Stock_Card_Num = '"
        strSQLOfStockcard = strSQLOfStockcard & strStockcardNo & "'"
    ADORecordsetOpen strSQLOfStockcard, m_conSADBEL, rstStockcardInProduct, adOpenKeyset, adLockOptimistic
    'rstStockcardInProduct.Open strSQLOfStockcard, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If rstStockcardInProduct.EOF And rstStockcardInProduct.BOF Then
        StockcardExistsInProduct = False
    Else
        StockcardExistsInProduct = True
    End If
    
    ADORecordsetClose rstStockcardInProduct
End Function

Private Sub CreateGridCol()
    jgxStock.Columns.Clear
    jgxStock.Columns.Add "Product Number", , , "Product Number"
    jgxStock.Columns.Add "Stock Card No", , , "Stock Card No"
    jgxStock.Columns.Add "Num of Items", , , "Num of Items"
    jgxStock.Columns.Add "Gross Weight", , , "Gross Weight"
    jgxStock.Columns.Add "Net Weight", , , "Net Weight"
    jgxStock.Columns.Add "Package Type", , , "Package Type"
    jgxStock.Columns.Add "Job Num", , , "Job Num"
    jgxStock.Columns.Add "Batch Num", , , "Batch Num"
    
    jgxStock.Columns("Product Number").Width = 1305
    jgxStock.Columns("Stock Card No").Width = 1200
    jgxStock.Columns("Num of Items").Width = 1095
    jgxStock.Columns("Gross Weight").Width = 1095
    jgxStock.Columns("Net Weight").Width = 1005
    jgxStock.Columns("Package Type").Width = 1200
    jgxStock.Columns("Job Num").Width = 1005
    jgxStock.Columns("Batch Num").Width = 1005
    
End Sub

Private Function ExistingDocNum() As Boolean
    Dim strSQL As String
    Dim rstInboundDocs As ADODB.Recordset

        strSQL = "Select DISTINCT InboundDocs.InDoc_ID  from InboundDocs LEFT JOIN (Inbounds left JOIN (StockCards LEFT join (Products left Join Entrepots on Products.Entrepot_ID = Entrepots.Entrepot_ID) on StockCards.Prod_ID = Products.Prod_ID)  on Inbounds.Stock_ID = StockCards.Stock_ID) On InboundDocs.InDoc_ID = Inbounds.InDoc_ID  where (InDoc_Global = -1 and InDoc_Num = '" & txtDocNum.Text & "' and ((Entrepot_Type & '-' & Entrepot_Num) = '" & txtEntrepotNum.Text & "' or Entrepots.Entrepot_ID IS NULL))"
    ADORecordsetOpen strSQL, m_conSADBEL, rstInboundDocs, adOpenKeyset, adLockOptimistic
    'rstInboundDocs.Open strSQL, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
    
    If Not (rstInboundDocs.BOF And rstInboundDocs.EOF) Then
        ExistingDocNum = True
    Else
        ExistingDocNum = False
    End If
    
    ADORecordsetClose rstInboundDocs
End Function

Private Function CheckInvalidRecord() As Boolean
                    
    Dim strStock As String
    
    If lngProd_ID = 0 Then
    
        If ProdPick(True) Then
            blnSystemChanged = True
            jgxStock.Col = jgxStock.Columns("Product Number").Index
            jgxStock.EditMode = jgexEditModeOn
            jgxStock.SelStart = 0
            blnSystemChanged = False
            CheckInvalidRecord = True
            Exit Function
        End If
                            
    End If
                
    If lngStockID = 0 Then
                        
        strStock = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Stock Card No").Index)), "", jgxStock.Value(jgxStock.Columns("Stock Card No").Index))
        
        If strStock = "" Then
            If Not StockPick() Then
                blnSystemChanged = True
                jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                jgxStock.EditMode = jgexEditModeOn
                jgxStock.SelStart = 0
                blnSystemChanged = False
                CheckInvalidRecord = True
                Exit Function
            End If
        ElseIf strStock <> "" Then
            If Not UniqueStockcard(blnEntered, jgxStock.Value(jgxStock.Columns("Prod_ID").Index), strStock) Then
                If lngStockID = 0 Then
                    blnSystemChanged = True
                    jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                    jgxStock.EditMode = jgexEditModeOn
                    jgxStock.SelStart = 0
                    blnSystemChanged = False
                    CheckInvalidRecord = True
                    Exit Function
                End If
            ElseIf Not blnEntered Then
                If Not StockPick() Then
                    blnSystemChanged = True
                    jgxStock.Col = jgxStock.Columns("Stock Card No").Index
                    jgxStock.EditMode = jgexEditModeOn
                    jgxStock.SelStart = 0
                    blnSystemChanged = False
                    CheckInvalidRecord = True
                    Exit Function
                End If
            End If
            jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
        End If
    ElseIf lngStockID = -1 Then
        jgxStock.Value(jgxStock.Columns("Stock Card No").Index) = NewStockcardNo
        jgxStock.Value(jgxStock.Columns("Stock_ID").Index) = lngStockID
    End If
    
    If lngPack_Flag = 0 Then
                            
        Dim strPack As String
        
        strPack = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Package Type").Index)), "", jgxStock.Value(jgxStock.Columns("Package Type").Index))
        
        Set pckList = New CPicklist
        Set gsdList = pckList.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
                    
        With pckList
    
            If Trim(strPack) <> "" Then
                .Search True, "Key Code", strPack
            End If
            ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
            .Pick Me, cpiSimplePicklist, m_conSADBEL, strSQLPack, "Code", "Codes", vbModal, gsdList, , , True, cpiKeyEnter
                            
            If Not .SelectedRecord Is Nothing Then
                jgxStock.Value(jgxStock.Columns("Package Type").Index) = .SelectedRecord.RecordSource.Fields("Key Code").Value
                jgxStock.Value(jgxStock.Columns("Pack_Flag").Index) = 1
                lngPack_Flag = 1
            Else
                blnSystemChanged = True
                jgxStock.Col = jgxStock.Columns("Package Type").Index
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

Private Function ValidPackageType() As Boolean
    
    Dim rstPackages As ADODB.Recordset
    Dim strPackage As String
    
    strPackage = IIf(IsNull(jgxStock.Value(jgxStock.Columns("Package Type").Index)), "", jgxStock.Value(jgxStock.Columns("Package Type").Index))
    
    ADORecordsetOpen "SELECT [PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Key Code]," & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Code]," & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].[DESCRIPTION " & m_strLanguage & "] AS [Key Description] " & _
                      "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE " & m_strLanguage & "] " & _
                      "WHERE " & _
                      "([PICKLIST DEFINITION].[BOX CODE]= 'E3') AND " & _
                      "([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND " & _
                      "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE " & m_strLanguage & "].[internal code]) AND " & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].CODE = '" & strPackage & "'", _
                      m_conSADBEL, rstPackages, adOpenKeyset, adLockOptimistic
    'rstPackages.Open "SELECT [PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Key Code]," & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].CODE AS [Code]," & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].[DESCRIPTION " & m_strLanguage & "] AS [Key Description] " & _
                      "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE " & m_strLanguage & "] " & _
                      "WHERE " & _
                      "([PICKLIST DEFINITION].[BOX CODE]= 'E3') AND " & _
                      "([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND " & _
                      "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE " & m_strLanguage & "].[internal code]) AND " & _
                      "[PICKLIST MAINTENANCE " & m_strLanguage & "].CODE = '" & strPackage & "'", _
                      m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    If Not (rstPackages.EOF And rstPackages.BOF) Then
        ValidPackageType = False
    Else
        ValidPackageType = True
    End If
    
    ADORecordsetClose rstPackages
End Function

Private Sub DeleteInInboundDocs()
    Dim rstTemp As ADODB.Recordset
    Dim lngInDoc_ID As Long

    If m_alngDeleted(0) <> 0 Or jgxStock.RowCount > 0 Then
        lngInDoc_ID = GetInDoc_ID
        
        ADORecordsetOpen "Select * from Inbounds where InDoc_ID = " & lngInDoc_ID, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
        'rstTemp.Open "Select * from Inbounds where InDoc_ID = " & lngInDoc_ID, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
        
        If rstTemp.BOF And rstTemp.EOF Then
            ExecuteNonQuery m_conSADBEL, "Delete from InboundDocs where InDoc_ID = " & lngInDoc_ID
        End If
                
        ADORecordsetClose rstTemp
    End If

End Sub

Private Function GetInDoc_ID() As Long
    Dim rstInboundDocs As ADODB.Recordset
    Dim strSQL As String

    '================== get the corresponding InDoc_ID =====================================
        strSQL = "Select DISTINCT InboundDocs.InDoc_ID  from InboundDocs LEFT JOIN (Inbounds left JOIN (StockCards LEFT join (Products left Join Entrepots on Products.Entrepot_ID = Entrepots.Entrepot_ID) on StockCards.Prod_ID = Products.Prod_ID)  on Inbounds.Stock_ID = StockCards.Stock_ID) On InboundDocs.InDoc_ID = Inbounds.InDoc_ID  where (InDoc_Global = -1 and InDoc_Num = '" & txtDocNum.Text & "' and ((Entrepot_Type & '-' & Entrepot_Num) = '" & txtEntrepotNum.Text & "' or Entrepots.Entrepot_ID IS NULL))"
    ADORecordsetOpen strSQL, m_conSADBEL, rstInboundDocs, adOpenKeyset, adLockOptimistic
    'rstInboundDocs.Open strSQL, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
    
    If Not (rstInboundDocs.BOF And rstInboundDocs.EOF) Then
        rstInboundDocs.MoveFirst
        
        GetInDoc_ID = rstInboundDocs!InDoc_ID
    End If
    
    ADORecordsetClose rstInboundDocs
End Function

Private Function CheckIfRowToAdd() As Boolean

    Dim lngCtr As Long
    Dim lngCounter As Long
    Dim UserReply As VbMsgBoxResult
    Dim lngRowCount As Long
    Dim varValue As Variant
    
    If fraStock.Enabled = False Or jgxStock.Row <> -1 Then
        CheckIfRowToAdd = True
        Exit Function
    End If
    
    For lngCtr = 1 To jgxStock.Columns.Count
        
        If m_rstStocksOff.Fields(jgxStock.Columns(lngCtr).DataField).Type = adInteger Then
            varValue = Val(jgxStock.Value(lngCtr))
        ElseIf m_rstStocksOff.Fields(jgxStock.Columns(lngCtr).DataField).Type = adBoolean Then
            varValue = CBool(jgxStock.Value(lngCtr))
        Else
            varValue = IIf(IsNull(jgxStock.Value(lngCtr)), "", jgxStock.Value(lngCtr))
        End If
        
        If varValue <> IIf(IsNull(jgxStock.Columns(lngCtr).DefaultValue), "", jgxStock.Columns(lngCtr).DefaultValue) Then
            UserReply = MsgBox("A record is waiting to be added. Would you like to add it now?", vbYesNoCancel + vbQuestion, "Initial Stocks")
            If UserReply = vbYes Then
                
                blnEntered = True
                lngRowCount = jgxStock.RowCount
                jgxStock.Update
                
                If jgxStock.RowCount = lngRowCount + 1 Then
                    CheckIfRowToAdd = True
                    cmdApply.Enabled = True
                Else
                    CheckIfRowToAdd = False
                End If
                
            ElseIf UserReply = vbNo Then
                CheckIfRowToAdd = True
                For lngCounter = 1 To jgxStock.Columns.Count
                    jgxStock.Value(lngCounter) = IIf(IsNull(jgxStock.Columns(lngCounter).DefaultValue), "", jgxStock.Columns(lngCounter).DefaultValue)
                Next
            Else
                CheckIfRowToAdd = False
            End If
            Exit Function
        End If
    Next
    
    CheckIfRowToAdd = True
    
End Function

Private Function GenerateUniqueID(ByVal rstInboundHistory As ADODB.Recordset) As Long
'

    Dim lngUniqueID As Long
    Dim blnNotUnique As Boolean
    Dim strSQL  As String

    Do While True

        Randomize
        lngUniqueID = CLng((-2147483646 * Rnd) + 1)   ' Generate random value between -2147483646 and 1.
        
        If Not (rstInboundHistory.BOF And rstInboundHistory.EOF) Then
            rstInboundHistory.MoveFirst
        End If
        
        rstInboundHistory.Find "In_ID = " & lngUniqueID
        
        If rstInboundHistory.EOF Then
            GenerateUniqueID = lngUniqueID
            Exit Function
        End If

        Randomize
        lngUniqueID = CLng((2147483646 * Rnd) + 1)   ' Generate random value between 1 and 2147483646.
        
        If Not (rstInboundHistory.BOF And rstInboundHistory.EOF) Then
            rstInboundHistory.MoveFirst
        End If
        
        rstInboundHistory.Find "In_ID = " & lngUniqueID
        
        If rstInboundHistory.EOF Then
            GenerateUniqueID = lngUniqueID
            Exit Function
        End If


    Loop

    '

End Function

Private Sub ResetControlValues(ByVal blnEntrepotHasValue As Boolean)
    'Reset Document Number
    txtDocType.Text = ""
    txtDocNum.Text = ""
    txtDocType.Enabled = blnEntrepotHasValue
    txtDocNum.Enabled = blnEntrepotHasValue

    'Reset Product Info
    lblCtryExportCode.Caption = ""
    lblCtryExportDesc.Caption = ""
    lblCtryOriginCode.Caption = ""
    lblCtryOriginDesc.Caption = ""
    lblDesc.Caption = ""
    lblHand.Caption = ""
    lblTaric.Caption = ""
End Sub
