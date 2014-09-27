VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmResetReservedStocks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reset Reserved Stocks"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmResetReservedStocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox filList 
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Tag             =   "178"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Tag             =   "179"
      Top             =   3720
      Width           =   1215
   End
   Begin GridEX16.GridEX jgxGrid 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5741
      TabKeyBehavior  =   1
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnCount     =   7
      CardCaption1    =   -1  'True
      ColCaption1     =   "In_ID"
      ColKey1         =   "In_ID"
      ColVisible1     =   0   'False
      ColCaption2     =   "Entrepot No"
      ColKey2         =   "Entrepot No"
      ColWidth2       =   1155
      ColCaption3     =   "Product No"
      ColKey3         =   "Product No"
      ColWidth3       =   1155
      ColCaption4     =   "Stock Card No"
      ColKey4         =   "Stock Card Number"
      ColWidth4       =   1305
      ColCaption5     =   "Reserved Stocks"
      ColKey5         =   "Reserved Stocks"
      ColWidth5       =   1395
      ColCaption6     =   "Stocks to Reset"
      ColKey6         =   "Stocks to Reset"
      ColWidth6       =   1395
      ColCaption7     =   "Available Stocks"
      ColKey7         =   "Available Stocks"
      ColWidth7       =   1395
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
   End
   Begin VB.Frame fraResetStocks 
      Caption         =   "Reset Stocks"
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmResetReservedStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
        
    Private m_lngUserID As Long
    
    Private m_conSADBEL As ADODB.Connection
    Private m_conEDIFACT As ADODB.Connection
    
    Private m_rstStocksOff As ADODB.Recordset
    Private m_rstStockcardOff As ADODB.Recordset
    
    Private Const G_CONST_EDINCTS1_TYPE = "EDI NCTS"        ' = cImport = "Import"
    Private Const G_CONST_NCTS1_TYPE = "Transit NCTS"        ' = cImport = "Import"
    Private Const G_CONST_NCTS2_TYPE = "Combined NCTS"       ' = cImport = "Import"

    Public Enum IE29Values
        enuIE29Val_NotFromIE29 = 1
        
        enuIEVal_IE43_Marks_And_Numbers
        enuIEVal_IE43_Number_of_Packages
        enuIEVal_IE43_Kind_of_Packages
        enuIEVal_IE43_Container_Numbers
        enuIEVal_IE43_Description_of_Goods
        enuIEVal_IE43_Sensitivity_Code
        enuIEVal_IE43_Sensitive_Quantity
        enuIEVal_IE43_Country_of_Dispatch_Export
        enuIEVal_IE43_Country_of_Destination
        enuIEVal_IE43_CO_Departure                      'LOC+118(2)
        enuIEVal_IE43_Gross_Mass
        enuIEVal_IE43_Net_Mass
        enuIEVal_IE43_Additional_Information
        enuIEVal_IE43_Consignor_TIN
        enuIEVal_IE43_Consignor_Name
        enuIEVal_IE43_Consignor_Street_And_Number
        enuIEVal_IE43_Consignor_Postal_Code
        enuIEVal_IE43_Consignor_City
        enuIEVal_IE43_Consignor_Country
        enuIEVal_IE43_Consignee_TIN
        enuIEVal_IE43_Consignee_Name
        enuIEVal_IE43_Consignee_Street_And_Number
        enuIEVal_IE43_Consignee_Postal_Code
        enuIEVal_IE43_Consignee_City
        enuIEVal_IE43_Consignee_Country
        enuIEVal_IE43_Document_Type
        enuIEVal_IE43_Document_Reference
        enuIEVal_IE43_Document_Complement_Information
        enuIEVal_IE43_Detail_Number
        enuIEVal_IE43_Commodity_Code
        
        enuIE29Val_MessageIdentification                'UNH(1)
        enuIE29Val_ReferenceNumber                      'BGM(5)
        enuIE29Val_AuthorizedLocationOfGoods            'LOC+14(6)
        enuIE29Val_DeclarationPlace                     'LOC+91(5)
        enuIE29Val_COReferencNumber                     'LOC+168(2) - CO = Customs Office
        enuIE29Val_COName                               'LOC+168(5) - CO = Customs Office
        enuIE29Val_COCountry                            'LOC+168(6) - CO = Customs Office
        enuIE29Val_COStreetAndNumber                    'LOC+168(9) - CO = Customs Office
        enuIE29Val_COPostalCode                         'LOC+168(10) - CO = Customs Office
        enuIE29Val_COCity                               'LOC+168(13) - CO = Customs Office
        enuIE29Val_COLanguage                           'LOC+168(14) - CO = Customs Office
        enuIE29Val_DateApproval                         'DTM+148(2)
        enuIE29Val_DateIssuance                         'DTM+182(2)
        enuIE29Val_DateControl                          'DTM+9(2)
        enuIEVal_IE29_DateLimitTransit                  'DTM+268(2)
        enuIE29Val_ReturnCopy                           'GIS 62(2)
        enuIE29Val_BindingItinerary                     'FTX+ABL(6)
        enuIE29Val_NotValidForEC                        'PCI+19(2)
        enuIE29Val_TPName                               'NAD+AF(10) - TP = Transit Principal
        enuIE29Val_TPStreetAndNumber                    'NAD+AF(16) - TP = Transit Principal
        enuIE29Val_TPCity                               'NAD+AF(20) - TP = Transit Principal
        enuIE29Val_TPPostalCode                         'NAD+AF(22) - TP = Transit Principal
        enuIE29Val_TPCountry                            'NAD+AF(23) - TP = Transit Principal
        enuIE29Val_ControlledBy                         'NAD+EI(2)
        
        enuIEVal_IE28_TPTIN                             'NAD+AF(2)  - TP = Transit Principal
        enuIEVal_IE28_TPName                            'NAD+AF(10) - TP = Transit Principal
        enuIEVal_IE28_TPStreetAndNumber                 'NAD+AF(16) - TP = Transit Principal
        enuIEVal_IE28_TPCity                            'NAD+AF(20) - TP = Transit Principal
        enuIEVal_IE28_TPPostalCode                      'NAD+AF(22) - TP = Transit Principal
        enuIEVal_IE28_TPCountry                         'NAD+AF(23) - TP = Transit Principal
    End Enum

    Public Enum eTabType
        eTab_Header = 1
        eTab_Detail = 2
    End Enum

Public Sub My_Load(ByRef conSADBEL As ADODB.Connection, conEdifact As ADODB.Connection, ByVal UserID As Long)
    
    Set m_conSADBEL = conSADBEL
    Set m_conEDIFACT = conEdifact
        
    m_lngUserID = UserID

    ' Create linked tables for master tables
    CreateLinkedTableMaster
    
    Me.Show vbModal
    
End Sub

Private Sub CreateLinkedTableMaster()
    
    '<<< dandan 112806
    '<<< Create link table
    CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "MDB_MASTER" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_DATA, "MASTER"
    CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_DATA, "MASTERNCTS"
    CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_EDIFACT, "MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00"), DBInstanceType_DATABASE_DATA, "MASTEREDINCTS"
    'AddLinkedTableEx "MDB_MASTER" & "_" & Format(m_lngUserID, "00"), NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "MASTER", NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_data.mdb", G_Main_Password
    'AddLinkedTableEx "MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00"), NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "MASTERNCTS", NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_data.mdb", G_Main_Password
    'AddLinkedTableEx "MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00"), NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\EDIFACT.mdb", G_Main_Password, "MASTEREDINCTS", NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_data.mdb", G_Main_Password
    
End Sub

Private Sub DropLinkedTableMaster()
    
    On Error Resume Next
    ExecuteNonQuery m_conSADBEL, "DROP TABLE MDB_MASTER" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conSADBEL, "DROP TABLE MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00")
    ExecuteNonQuery m_conEDIFACT, "DROP TABLE MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00")
    On Error GoTo 0
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'If there are records in the grid, update the database.
    If jgxGrid.Enabled = True Then
        UpdateStock
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Initialize the offline recordsets and the grid.
    InitializeOfflineRecordset m_rstStocksOff, False
    InitializeGrid
    InitializeOfflineRecordset m_rstStockcardOff, True
    
End Sub

Private Sub InitializeOfflineRecordset(ByRef RecordsetOff As ADODB.Recordset, blnToUpdate As Boolean)

    Set RecordsetOff = New ADODB.Recordset
    RecordsetOff.CursorLocation = adUseClient
    
    If blnToUpdate = False Then
        RecordsetOff.Fields.Append "In_ID", adVarWChar, 50
        RecordsetOff.Fields.Append "In_Code", adVarWChar, 50
        RecordsetOff.Fields.Append "Entrepot No", adVarWChar, 50
        RecordsetOff.Fields.Append "Product No", adVarWChar, 50
        RecordsetOff.Fields.Append "Handling", adVarWChar, 50
        RecordsetOff.Fields.Append "Stock Card No", adVarWChar, 50
        
        RecordsetOff.Fields.Append "Document Number", adVarWChar, 50
        RecordsetOff.Fields.Append "Batch Number", adVarWChar, 50
        
        RecordsetOff.Fields.Append "Reserved Stocks", adVarWChar, 50
        RecordsetOff.Fields.Append "Reserved Stocks2", adVarWChar, 50
        RecordsetOff.Fields.Append "Stocks to Reset", adVarWChar, 50
        RecordsetOff.Fields.Append "Available Stocks", adVarWChar, 50
        RecordsetOff.Fields.Append "Available Stocks2", adVarWChar, 50
    Else
        RecordsetOff.Fields.Append "In_ID", adVarWChar, 50
        RecordsetOff.Fields.Append "Stocks to Reset", adVarWChar, 50
    End If
    
    RecordsetOff.Open
    
End Sub

Private Sub InitializeGrid()
    Dim strSQL As String
    Dim lngCounter As Long
    Dim dblReserveStocks As Double
    Dim rstStock As ADODB.Recordset
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "INBOUNDS.In_ID As In_ID, "
        strSQL = strSQL & "INBOUNDS.In_Code As In_Code, "
        strSQL = strSQL & "ENTREPOTS.Entrepot_Type & '-' & ENTREPOTS.Entrepot_Num AS [Entrepot No], "
        strSQL = strSQL & "PRODUCTS.Prod_Num AS [Product No], "
        strSQL = strSQL & "PRODUCTS.Prod_Handling AS [Handling], "
        strSQL = strSQL & "STOCKCARDS.Stock_Card_Num AS [Stock Card No], "
        strSQL = strSQL & "INBOUNDS.In_Reserved_Qty_Wgt AS [Reserved Stocks], "
        strSQL = strSQL & "INBOUNDS.In_Avl_Qty_Wgt AS [Available Stocks], "
        strSQL = strSQL & "INBOUNDS.In_Batch_Num AS [Batch Number], "
        strSQL = strSQL & "INBOUNDDOCS.InDoc_Num AS [Document Number] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "(INBOUNDS INNER JOIN INBOUNDDOCS ON INBOUNDS.InDoc_ID = INBOUNDDOCS.InDoc_ID)"
            strSQL = strSQL & "INNER JOIN (STOCKCARDS "
                strSQL = strSQL & "INNER JOIN (PRODUCTS "
                    strSQL = strSQL & "INNER JOIN ENTREPOTS ON PRODUCTS.Entrepot_ID = ENTREPOTS.Entrepot_ID) "
                strSQL = strSQL & "ON STOCKCARDS.Prod_ID = PRODUCTS.Prod_ID) "
            strSQL = strSQL & "ON INBOUNDS.Stock_ID = STOCKCARDS.Stock_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "INBOUNDS.In_Reserved_Qty_Wgt <> 0 "
        strSQL = strSQL & "GROUP BY "
        strSQL = strSQL & "INBOUNDS.In_ID, "
        strSQL = strSQL & "INBOUNDS.In_Code, "
        strSQL = strSQL & "ENTREPOTS.Entrepot_Type & '-' & ENTREPOTS.Entrepot_Num, "
        strSQL = strSQL & "PRODUCTS.Prod_Num, "
        strSQL = strSQL & "PRODUCTS.Prod_Handling, "
        strSQL = strSQL & "STOCKCARDS.Stock_Card_Num, "
        strSQL = strSQL & "INBOUNDS.In_Reserved_Qty_Wgt, "
        strSQL = strSQL & "INBOUNDS.In_Avl_Qty_Wgt, "
        strSQL = strSQL & "INBOUNDS.In_Batch_Num, "
        strSQL = strSQL & "INBOUNDDOCS.InDoc_Num "
    ADORecordsetOpen strSQL, m_conSADBEL, rstStock, adOpenKeyset, adLockOptimistic
    'rstStock.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    With rstStock
        If Not (.BOF And .EOF) Then
            .MoveFirst
            
            Do While Not .EOF
                'Add records to the offline recordset.
                m_rstStocksOff.AddNew
                
                m_rstStocksOff.Fields("In_ID").Value = .Fields("In_ID").Value
                m_rstStocksOff.Fields("In_Code").Value = IIf(IsNull(.Fields("In_Code").Value), "", .Fields("In_Code").Value)
                m_rstStocksOff.Fields("Entrepot No").Value = .Fields("Entrepot No").Value
                m_rstStocksOff.Fields("Product No").Value = .Fields("Product No").Value
                m_rstStocksOff.Fields("Handling").Value = .Fields("Handling").Value
                m_rstStocksOff.Fields("Stock Card No").Value = .Fields("Stock Card No").Value
                
                m_rstStocksOff.Fields("Batch Number").Value = .Fields("Batch Number").Value
                m_rstStocksOff.Fields("Document Number").Value = .Fields("Document Number").Value
                
                dblReserveStocks = Round(.Fields("Reserved Stocks").Value - (Round(CheckForOutbox(m_rstStocksOff.Fields("In_ID").Value, m_rstStocksOff.Fields("In_Code").Value, m_rstStocksOff.Fields("Handling").Value), Choose(.Fields("Handling").Value + 1, 0, 3, 3))), Choose(.Fields("Handling").Value + 1, 0, 3, 3))
                        
                If .Fields("Handling").Value <> 0 Then
                    m_rstStocksOff.Fields("Reserved Stocks").Value = Replace(Format(dblReserveStocks, "0.###"), ",", ".")
                    m_rstStocksOff.Fields("Reserved Stocks2").Value = Replace(Format(dblReserveStocks, "0.###"), ",", ".")
                    m_rstStocksOff.Fields("Available Stocks").Value = Replace(Format(.Fields("Available Stocks").Value, "0.###"), ",", ".")
                    m_rstStocksOff.Fields("Available Stocks2").Value = Replace(Format(.Fields("Available Stocks").Value, "0.###"), ",", ".")
                Else
                    m_rstStocksOff.Fields("Reserved Stocks").Value = dblReserveStocks
                    m_rstStocksOff.Fields("Reserved Stocks2").Value = dblReserveStocks
                    m_rstStocksOff.Fields("Available Stocks").Value = .Fields("Available Stocks").Value
                    m_rstStocksOff.Fields("Available Stocks2").Value = .Fields("Available Stocks").Value
                End If
                m_rstStocksOff.Update
                
                .MoveNext
            Loop
        End If
    End With
    
    ADORecordsetClose rstStock
    
    'Make sure that the records shown on the grid have reserved stocks value
    'greater than 0.
    m_rstStocksOff.Filter = "[Reserved Stocks] > 0"
    
    'If there are still records on the recordset even after filtering it,
    'show them on the grid.
    If m_rstStocksOff.RecordCount > 0 Then
        Set jgxGrid.ADORecordset = Nothing
        Set jgxGrid.ADORecordset = m_rstStocksOff
        
        EditableGrid
    Else
        jgxGrid.Enabled = False
    End If
    
End Sub

'This procedure checks for the existence of documents in the Outbox folders.
Private Function CheckForOutbox(ByVal In_ID As String, In_Code As String, Handling As Long)
    Dim strSQL As String
    Dim lngCounter As Long
    Dim lngCtr As Long
    Dim rstOutbox As ADODB.Recordset
    Dim dblOutbound As Double
    
    Dim lngAttempts As Long

    dblOutbound = 0
    For lngCtr = 1 To 6
        Select Case lngCtr
            Case 1 'IMPORT
                strSQL = vbNullString
                strSQL = strSQL & "Select IMPORT.CODE AS CODE, MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[IMPORT DETAIL].In_ID AS In_ID, "
                strSQL = strSQL & "[IMPORT DETAIL].T6 AS [Number of Packages], "
                strSQL = strSQL & "[IMPORT DETAIL].M1 AS [Gross Weight], "
                strSQL = strSQL & "[IMPORT DETAIL].M2 AS [Net Weight] "
                strSQL = strSQL & "FROM (IMPORT "
                strSQL = strSQL & "INNER JOIN ([IMPORT HEADER] "
                strSQL = strSQL & "INNER JOIN [IMPORT DETAIL] "
                strSQL = strSQL & "ON [IMPORT HEADER].CODE = [IMPORT DETAIL].CODE "
                strSQL = strSQL & "AND [IMPORT HEADER].HEADER=[IMPORT DETAIL].HEADER) "
                strSQL = strSQL & "ON IMPORT.CODE =[IMPORT HEADER].CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTER" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".CODE = IMPORT.CODE "
                strSQL = strSQL & "WHERE MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = 'WL2' "
                strSQL = strSQL & "AND [IMPORT DETAIL].In_ID = " & In_ID
                
            Case 2 'EXPORT
                strSQL = vbNullString
                strSQL = strSQL & "Select EXPORT.CODE AS CODE, MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[EXPORT DETAIL].In_ID AS In_ID, "
                strSQL = strSQL & "[EXPORT DETAIL].S3 AS [Number of Packages], [EXPORT DETAIL].M1 AS [Gross Weight], "
                strSQL = strSQL & "[EXPORT DETAIL].M2 AS [Net Weight] "
                strSQL = strSQL & "FROM (EXPORT "
                strSQL = strSQL & "INNER JOIN ([EXPORT HEADER] "
                strSQL = strSQL & "INNER JOIN [EXPORT DETAIL] "
                strSQL = strSQL & "ON [EXPORT HEADER].CODE=[EXPORT DETAIL].CODE "
                strSQL = strSQL & "AND [EXPORT HEADER].HEADER=[EXPORT DETAIL].HEADER) "
                strSQL = strSQL & "ON EXPORT.CODE =[EXPORT HEADER].CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTER" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".CODE = EXPORT.CODE "
                strSQL = strSQL & "WHERE MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = 'WL2' "
                strSQL = strSQL & "AND [EXPORT DETAIL].In_ID = " & In_ID
                
            Case 3 'TRANSIT
                strSQL = vbNullString
                strSQL = strSQL & "Select TRANSIT.CODE AS CODE, MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[TRANSIT DETAIL].In_ID AS In_ID, "
                strSQL = strSQL & "[TRANSIT DETAIL].S3 AS [Number of Packages], "
                strSQL = strSQL & "[TRANSIT DETAIL].M1 AS [Gross Weight], [TRANSIT DETAIL].M2 AS [Net Weight] "
                strSQL = strSQL & "FROM (TRANSIT "
                strSQL = strSQL & "INNER JOIN ([TRANSIT HEADER] "
                strSQL = strSQL & "INNER JOIN [TRANSIT DETAIL] "
                strSQL = strSQL & "ON [TRANSIT HEADER].CODE = [TRANSIT DETAIL].CODE "
                strSQL = strSQL & "AND [TRANSIT HEADER].HEADER = [TRANSIT DETAIL].HEADER) "
                strSQL = strSQL & "ON TRANSIT.CODE =[TRANSIT HEADER].CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTER" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".CODE = TRANSIT.CODE "
                strSQL = strSQL & "WHERE MDB_MASTER" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = 'WL2' "
                strSQL = strSQL & "AND [TRANSIT DETAIL].In_ID = " & In_ID
                
            Case 4 'SADBEL NCTS
                strSQL = vbNullString
                strSQL = strSQL & "Select NCTS.CODE AS CODE, MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[NCTS DETAIL].In_ID AS In_ID, "
                strSQL = strSQL & "[NCTS DETAIL COLLI].S3 AS [Number of Packages], "
                strSQL = strSQL & "[NCTS DETAIL].M1 AS [Gross Weight], [NCTS DETAIL].M2 AS [Net Weight] "
                strSQL = strSQL & "FROM (NCTS "
                strSQL = strSQL & "INNER JOIN ([NCTS HEADER] "
                strSQL = strSQL & "INNER JOIN ([NCTS DETAIL] "
                strSQL = strSQL & "INNER JOIN [NCTS DETAIL COLLI] "
                strSQL = strSQL & "ON ([NCTS DETAIL COLLI].CODE = [NCTS DETAIL].CODE "
                strSQL = strSQL & "AND [NCTS DETAIL COLLI].DETAIL = [NCTS DETAIL].DETAIL)) "
                strSQL = strSQL & "ON [NCTS HEADER].CODE=[NCTS DETAIL].CODE) "
                strSQL = strSQL & "ON NCTS.CODE =[NCTS HEADER].CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".CODE = NCTS.CODE "
                strSQL = strSQL & "WHERE MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = 'WL2' "
                strSQL = strSQL & "AND [NCTS DETAIL].In_ID = " & In_ID & " "
                strSQL = strSQL & "AND [NCTS DETAIL COLLI].ORDINAL = 1"
                
            Case 5 'COMBINED NCTS
                strSQL = vbNullString
                strSQL = strSQL & "Select [COMBINED NCTS].CODE AS CODE, MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[COMBINED NCTS DETAIL].In_ID AS In_ID, "
                strSQL = strSQL & "[COMBINED NCTS DETAIL COLLI].S3 AS [Number of Packages], "
                strSQL = strSQL & "[COMBINED NCTS DETAIL GOEDEREN].M1 AS [Gross Weight], "
                strSQL = strSQL & "[COMBINED NCTS DETAIL GOEDEREN].M2 AS [Net Weight] "
                strSQL = strSQL & "FROM ([COMBINED NCTS] "
                strSQL = strSQL & "INNER JOIN ([COMBINED NCTS HEADER] "
                strSQL = strSQL & "INNER JOIN ([COMBINED NCTS DETAIL] "
                strSQL = strSQL & "INNER JOIN ([COMBINED NCTS DETAIL COLLI] "
                strSQL = strSQL & "INNER JOIN [COMBINED NCTS DETAIL GOEDEREN] "
                strSQL = strSQL & "ON [COMBINED NCTS DETAIL COLLI].CODE=[COMBINED NCTS DETAIL GOEDEREN].CODE "
                strSQL = strSQL & "AND [COMBINED NCTS DETAIL COLLI].DETAIL=[COMBINED NCTS DETAIL GOEDEREN].DETAIL) "
                strSQL = strSQL & "ON [COMBINED NCTS DETAIL COLLI].CODE =[COMBINED NCTS DETAIL].CODE "
                strSQL = strSQL & "AND [COMBINED NCTS DETAIL COLLI].DETAIL =[COMBINED NCTS DETAIL].DETAIL) "
                strSQL = strSQL & "ON [COMBINED NCTS HEADER].CODE=[COMBINED NCTS DETAIL].CODE) "
                strSQL = strSQL & "ON [COMBINED NCTS].CODE =[COMBINED NCTS HEADER].CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".CODE = [COMBINED NCTS].CODE "
                strSQL = strSQL & "WHERE MDB_MASTERNCTS" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = 'WL2' "
                strSQL = strSQL & "AND [COMBINED NCTS DETAIL].In_ID = " & In_ID & " "
                strSQL = strSQL & "AND [COMBINED NCTS DETAIL COLLI].ORDINAL = 1 "
                
            Case 6 'EDI DEPARTURE - EDIFACT ANG CONNECTION
                strSQL = vbNullString
                strSQL = strSQL & "Select DATA_NCTS.CODE AS CODE, MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00") & ".DTYPE AS DTYPE, "
                strSQL = strSQL & "[DATA_NCTS_DETAIL].In_ID AS In_ID "
                strSQL = strSQL & "FROM (DATA_NCTS "
                strSQL = strSQL & "INNER JOIN DATA_NCTS_DETAIL "
                strSQL = strSQL & "ON DATA_NCTS_DETAIL.CODE = DATA_NCTS.CODE) "
                strSQL = strSQL & "INNER JOIN MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "ON MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00") & ".CODE = [DATA_NCTS].CODE "
                strSQL = strSQL & "WHERE MDB_MASTEREDINCTS" & "_" & Format(m_lngUserID, "00") & ".[TREE ID] = '31ED' "
                strSQL = strSQL & "AND [DATA_NCTS_DETAIL].In_ID = " & In_ID
                
        End Select
        
        If (Len(Trim(strSQL)) > 0) Then
            If (lngCtr <> 6) Then
                On Error GoTo ErrorHandler
                ADORecordsetOpen strSQL, m_conSADBEL, rstOutbox, adOpenKeyset, adLockOptimistic
                'rstOutbox.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                strSQL = vbNullString
            Else
                On Error GoTo ErrorHandler
                ADORecordsetOpen strSQL, m_conEDIFACT, rstOutbox, adOpenKeyset, adLockOptimistic
                'rstOutbox.Open strSQL, m_conEDIFACT, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                strSQL = vbNullString
            End If
        End If
        
        If Not (rstOutbox.BOF And rstOutbox.EOF) Then
            With rstOutbox
                .MoveFirst
                For lngCounter = 0 To .RecordCount - 1
                    If .Fields("DType") = 11 Then 'Edifact
                        dblOutbound = dblOutbound + GetValue(.Fields("Code").Value, .Fields("In_ID").Value, Handling)
                    Else
                        Select Case Handling
                            Case 0
                                    dblOutbound = dblOutbound + IIf(IsNull(.Fields("Number of Packages").Value), 0, .Fields("Number of Packages").Value)
                            Case 1
                                    dblOutbound = dblOutbound + IIf(IsNull(.Fields("Gross Weight").Value), 0, .Fields("Gross Weight").Value)
                            Case 2
                                    dblOutbound = dblOutbound + IIf(IsNull(.Fields("Net Weight").Value), 0, .Fields("Net Weight").Value)
                        End Select
                    End If
                    
                    If Not .EOF Then
                        .MoveNext
                    End If
                Next lngCounter
            End With
        End If
        
        ADORecordsetClose rstOutbox
    Next
    
    
    CheckForOutbox = dblOutbound
    
    ADORecordsetClose rstOutbox
    
    Exit Function
    
ErrorHandler:
    If (Err.Number = -2147217865) Then
        'The Microsoft Jet database engine cannot find the input table or query <table>.  Make sure it exists and that its name is spelled correctly.
        lngAttempts = lngAttempts + 1
        If (lngAttempts > 10) Then
            CreateLinkedTableMaster
            lngAttempts = 0
        End If
        
        Err.Clear
        Resume
        
    Else
        
    End If
    
End Function

Private Sub EditableGrid()
    'Set all of the columns except the 'No To Items Reset' to be not selectable.
    jgxGrid.Columns("Entrepot No").Selectable = False
    jgxGrid.Columns("Product No").Selectable = False
    jgxGrid.Columns("Stock Card No").Selectable = False
    jgxGrid.Columns("Reserved Stocks").Selectable = False
    jgxGrid.Columns("Available Stocks").Selectable = False
    
    jgxGrid.Columns("Batch Number").Selectable = False
    jgxGrid.Columns("Document Number").Selectable = False
    
    'Set the widths.
    jgxGrid.Columns("In_ID").Width = 1200
    jgxGrid.Columns("Entrepot No").Width = 1100
    jgxGrid.Columns("Product No").Width = 1300
    jgxGrid.Columns("Stock Card No").Width = 1300
    jgxGrid.Columns("Reserved Stocks").Width = 1400
    jgxGrid.Columns("Stocks to Reset").Width = 1300
    jgxGrid.Columns("Available Stocks").Width = 1400
    jgxGrid.Columns("Document Number").Width = 2000
    
    'Set the visibility property of some of the grid's columns.
    jgxGrid.Columns("In_ID").Visible = False
    jgxGrid.Columns("In_Code").Visible = False
    jgxGrid.Columns("Handling").Visible = False
    jgxGrid.Columns("Available Stocks2").Visible = False
    jgxGrid.Columns("Reserved Stocks2").Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ADORecordsetClose m_rstStocksOff
    ADORecordsetClose m_rstStockcardOff
    
    DropLinkedTableMaster
    
    Set m_conSADBEL = Nothing
    Set m_conEDIFACT = Nothing
End Sub

Private Sub jgxGrid_AfterColEdit(ByVal ColIndex As Integer)
    Dim lngCounter As Long
    Dim blnToAdd As Boolean

    If ColIndex = jgxGrid.Columns("Stocks to Reset").Index Then
        If Not IsNumeric(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)) Then
            jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index) = 0
        End If
            
        'If the record has not been added yet, add. If it has been, edit.
        If m_rstStockcardOff.RecordCount = 0 Then
        
            m_rstStockcardOff.AddNew
            m_rstStockcardOff.Fields("In_ID").Value = jgxGrid.Value(jgxGrid.Columns("In_ID").Index)
            m_rstStockcardOff.Fields("Stocks to Reset").Value = jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)
            m_rstStockcardOff.Update
            
        ElseIf m_rstStockcardOff.RecordCount > 0 Then
        
            m_rstStockcardOff.MoveFirst
            For lngCounter = 1 To m_rstStockcardOff.RecordCount
                If m_rstStockcardOff.Fields("In_ID").Value = jgxGrid.Value(jgxGrid.Columns("In_ID").Index) Then
                    'Edit
                    blnToAdd = False
                    Exit For
                Else
                    'Add
                    blnToAdd = True
                End If
                If Not m_rstStockcardOff.EOF Then
                    m_rstStockcardOff.MoveNext
                End If
            Next lngCounter
            
            If blnToAdd = False Then
                'm_rstStockcardOff.Fields("In_ID").Value = jgxGrid.Value(jgxGrid.Columns("In_ID").Index)
                m_rstStockcardOff.Fields("Stocks to Reset").Value = jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)
                m_rstStockcardOff.Update
            Else
                m_rstStockcardOff.AddNew
                m_rstStockcardOff.Fields("In_ID").Value = jgxGrid.Value(jgxGrid.Columns("In_ID").Index)
                m_rstStockcardOff.Fields("Stocks to Reset").Value = jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)
                m_rstStockcardOff.Update
            End If
        End If
        
        jgxGrid.Value(jgxGrid.Columns("Available Stocks").Index) = Val(jgxGrid.Value(jgxGrid.Columns("Available Stocks2").Index)) + Format(Val(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)), "0.###")
        jgxGrid.Value(jgxGrid.Columns("Available Stocks").Index) = Replace(Round(jgxGrid.Value(jgxGrid.Columns("Available Stocks").Index), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3)), ",", ".")
        
        jgxGrid.Value(jgxGrid.Columns("Reserved Stocks").Index) = Val(jgxGrid.Value(jgxGrid.Columns("Reserved Stocks2").Index)) - Format(Val(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)), "0.###")
        jgxGrid.Value(jgxGrid.Columns("Reserved Stocks").Index) = Replace(Round(jgxGrid.Value(jgxGrid.Columns("Reserved Stocks").Index), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3)), ",", ".")
        
    End If
    
        
End Sub

Private Sub UpdateStock()
    Dim rstUpdate As ADODB.Recordset
    Dim conHistory As ADODB.Connection
    
    Dim strSQL As String
    Dim lngCounter As Long
    Dim lngCtr As Long

    'Get the database path.
    filList.Path = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
    filList.Pattern = "mdb_History*.mdb"
    
    With m_rstStockcardOff
        If .RecordCount > 0 Then
            .MoveFirst
            
            For lngCounter = 1 To m_rstStockcardOff.RecordCount
            
                'Update Sadbel record.
                    strSQL = " Select INBOUNDS.In_ID As In_ID, INBOUNDS.In_Reserved_Qty_Wgt AS [Reserved Stocks], INBOUNDS.In_Avl_Qty_Wgt AS [Available Stocks] FROM INBOUNDS WHERE INBOUNDS.In_ID = " & m_rstStockcardOff.Fields("In_ID").Value
                ADORecordsetOpen strSQL, m_conSADBEL, rstUpdate, adOpenKeyset, adLockOptimistic
                'rstUpdate.Open strSQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
                
                If Not (rstUpdate.EOF And rstUpdate.BOF) Then
                    rstUpdate.MoveFirst
                    
                    rstUpdate.Fields("Reserved Stocks").Value = Round(rstUpdate.Fields("Reserved Stocks").Value - Val(m_rstStockcardOff.Fields("Stocks to Reset").Value), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3))
                    rstUpdate.Fields("Available Stocks").Value = Round(rstUpdate.Fields("Available Stocks").Value + Val(m_rstStockcardOff.Fields("Stocks to Reset").Value), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3))
                    rstUpdate.Update
                    
                    UpdateRecordset m_conSADBEL, rstUpdate, "Inbounds"
                End If
                
                ADORecordsetClose rstUpdate
                '******************************************************
                
                'Update History record.
                '******************************************************
                For lngCtr = 0 To filList.ListCount - 1
                    ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, GetHistoryDBYear(filList.List(lngCtr))
                    'OpenADODatabase conHistory, NoBackSlash(g_objDataSourceProperties.TracefilePath), filList.List(lngCtr)
                    
                    ADORecordsetOpen strSQL, conHistory, rstUpdate, adOpenKeyset, adLockOptimistic
                    'rstUpdate.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
                    
                    If Not (rstUpdate.BOF And rstUpdate.BOF) Then
                        rstUpdate.MoveFirst
                        
                        rstUpdate.Fields("Reserved Stocks").Value = Round(rstUpdate.Fields("Reserved Stocks").Value - Val(m_rstStockcardOff.Fields("Stocks to Reset").Value), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3))
                        rstUpdate.Fields("Available Stocks").Value = Round(rstUpdate.Fields("Available Stocks").Value + Val(m_rstStockcardOff.Fields("Stocks to Reset").Value), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3))
                        rstUpdate.Update
                        
                        UpdateRecordset conHistory, rstUpdate, "Inbounds"
                        
                        ADORecordsetClose rstUpdate
                        
                        ADODisconnectDB conHistory
                        
                        Exit For
                    End If
                    
                    ADORecordsetClose rstUpdate
                        
                    ADODisconnectDB conHistory
                Next lngCtr
                '******************************************************
                
                If Not m_rstStockcardOff.EOF Then
                    m_rstStockcardOff.MoveNext
                End If
            Next lngCounter
        End If
    End With
    
    ADORecordsetClose rstUpdate
                        
    ADODisconnectDB conHistory

End Sub

Private Sub jgxGrid_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX16.JSRetBoolean)
    If ColIndex = jgxGrid.Columns("Stocks to Reset").Index Then
        If Not IsNumeric(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)) Then
            If InStr(1, jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index), ".") Then
                jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index) = Replace(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index), ".", ",")
            End If
        End If
        If Val(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)) > Val(jgxGrid.Value(jgxGrid.Columns("Reserved Stocks").Index)) Then
            jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index) = jgxGrid.Value(jgxGrid.Columns("Reserved Stocks").Index)
        Else
            jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index) = Round(Val(jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)), Choose(jgxGrid.Value(jgxGrid.Columns("Handling").Index) + 1, 0, 3, 3))
            If InStr(1, jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index), ",") Then
                jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index) = Replace((jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index)), ",", ".")
            End If
        End If
   End If
End Sub

Private Sub jgxGrid_KeyPress(KeyAscii As Integer)
    'Allow decimal point input depending on the handling type.
    If jgxGrid.Value(jgxGrid.Columns("Handling").Index) = 0 Then
        If Chr(KeyAscii) = "." Then
            KeyAscii = 0
        ElseIf Not IsNumeric(Chr(KeyAscii)) Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    ElseIf jgxGrid.Value(jgxGrid.Columns("Handling").Index) = 1 Or _
        jgxGrid.Value(jgxGrid.Columns("Handling").Index) = 2 Then
        If Chr(KeyAscii) = "." Then
            If InStr(1, jgxGrid.Value(jgxGrid.Columns("Stocks to Reset").Index), ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        ElseIf Not IsNumeric(Chr(KeyAscii)) Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Public Function GetValueFromClass(ByVal EDIClass As EdifactMessage, _
                                  ByVal MapRecordset As ADODB.Recordset, _
                                  ByVal IE29Value As IE29Values, _
                                  ByVal ValueSource As String, _
                                  ByVal TabNumber As Long) As Variant
    Dim varReturnValue                  As Variant
    
    Dim strSegmentKey                   As String
    Dim strSegmentKeyCST                As String
    Dim strBoxCodeValue                 As String
    Dim strLeftSubString                As String
    Dim strRightSubString               As String
    Dim strMiddleSubString              As String
    Dim strSpaces                       As String
    Dim astrSegmentValues()             As String
    
    Dim lngValuesCount                  As Long
    Dim lngBoxCodeInstance              As Long
    Dim lngSegmentInstance              As Long
    Dim lngCST_NCTS_IEM_TMS_ID          As Long
    Dim lngNCTS_IEM_TMS_ID              As Long
    Dim lngNCTS_IEM_TMS_IDAlternative   As Long
    Dim lngDataItemOrdinal              As Long
    Dim lngDataItemsCount               As Long
    
    Dim blnValueIsComplete              As Boolean
    Dim blnContinueLoop                 As Boolean
    Dim blnIsNumeric                    As Boolean
    Dim blnSegmentInHeader              As Boolean
    
    ReDim astrSegmentValues(0)
    astrSegmentValues(0) = ""
    lngValuesCount = 0
    lngDataItemsCount = 1
    Call GetValueFromClassAssertions(EDIClass, IE29Value)
    Select Case IE29Value
        Case IE29Values.enuIE29Val_NotFromIE29
            lngBoxCodeInstance = 1 'TabNumber
            
            MapRecordset.Filter = "NCTS_IEM_MAP_Source = #" & ValueSource & "#"
            If MapRecordset.RecordCount > 0 Then
                lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                Select Case GetTabType(G_CONST_EDINCTS1_TYPE, ValueSource)
                    Case eTabType.eTab_Header
                        Do
                            If Not MapRecordset.BOF And Not MapRecordset.EOF Then
                                lngSegmentInstance = GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance)
                                blnValueIsComplete = False
                                strSegmentKey = "S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(lngSegmentInstance)
                                If EDIClass.GetSegmentIndex(strSegmentKey) > 0 Then
                                    lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                                    blnValueIsComplete = (MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0)
                                    If MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0 Then
                                        strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                        blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                    Else
                                        Debug.Assert MapRecordset.Fields("NCTS_IEM_MAP_Length").Value > 0
                                        strMiddleSubString = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                        blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                        Debug.Assert Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1
                                        If Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1 Then
                                            strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1) - Len(strMiddleSubString))
                                        Else
                                            strSpaces = ""
                                        End If
                                        strMiddleSubString = strMiddleSubString & strSpaces
                                        strLeftSubString = Left(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1)
                                        strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1) - Len(strLeftSubString))
                                        strLeftSubString = strLeftSubString & strSpaces
                                        strRightSubString = Mid(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_Length").Value + 1)
                                        strBoxCodeValue = strLeftSubString & strMiddleSubString & strRightSubString
                                    End If
                                End If
                            End If
                            
                            If Not blnValueIsComplete Then
                                MapRecordset.MoveNext
                                blnValueIsComplete = MapRecordset.EOF
                                If blnValueIsComplete Then
                                    MapRecordset.MoveFirst
                                End If
                            End If
                            'blnContinueLoop = (Not blnValueIsComplete) And Not MapRecordset.EOF
                            If blnValueIsComplete Then
                                'If IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or strSegmentKeyCST = strSegmentKey Then
                                    ReDim Preserve astrSegmentValues(lngValuesCount)
                                    If blnIsNumeric Then
                                        If Trim(strBoxCodeValue) = vbNullString Then
                                            astrSegmentValues(lngValuesCount) = "0"
                                        Else
                                            astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                        End If
                                    Else
                                        astrSegmentValues(lngValuesCount) = strBoxCodeValue
                                    End If
                                    lngValuesCount = lngValuesCount + 1
                                'End If
                                lngBoxCodeInstance = lngBoxCodeInstance + 1
                            End If
                            blnContinueLoop = (EDIClass.GetSegmentIndex("S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance))) > 0)
                        Loop While blnContinueLoop
                    Case eTabType.eTab_Detail
                        lngBoxCodeInstance = TabNumber
                        Select Case EDIClass.MessageType
                            Case ENCTSMessageType.EMsg_IE15
                                lngCST_NCTS_IEM_TMS_ID = 33
                            Case ENCTSMessageType.EMsg_IE29
                                lngCST_NCTS_IEM_TMS_ID = 126
                            Case ENCTSMessageType.EMsg_IE43
                                lngCST_NCTS_IEM_TMS_ID = 260
                                Debug.Assert False
                            Case ENCTSMessageType.EMsg_IE51
                                lngCST_NCTS_IEM_TMS_ID = 187
                                Debug.Assert False
                            ' <<<<
                            Case ENCTSMessageType.EMsg_IE13
                                lngCST_NCTS_IEM_TMS_ID = 442
                            Case Else
                                Debug.Assert False
                        End Select
                        
                        strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(TabNumber)
                        Debug.Assert EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0
                        If EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0 Then
                            Do
                                If Not MapRecordset.BOF And Not MapRecordset.EOF Then
                                    lngSegmentInstance = GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance)
                                    blnValueIsComplete = False
                                    strSegmentKey = "S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(lngSegmentInstance)
                                    If EDIClass.GetSegmentIndex(strSegmentKey) > 0 Then
                                        If IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or strSegmentKeyCST = strSegmentKey Then
                                            lngDataItemOrdinal = MapRecordset.Fields("NCTS_IEM_MAP_EDI_ITM_ORDINAL").Value
                                            blnValueIsComplete = (MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0)
                                            If MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value = 0 Then
                                                strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                                blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                            Else
                                                Debug.Assert MapRecordset.Fields("NCTS_IEM_MAP_Length").Value
                                                strMiddleSubString = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                                                blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                                                Debug.Assert Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1
                                                If Len(strMiddleSubString) <= MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1 Then
                                                    strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_Length").Value - MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value + 1) - Len(strMiddleSubString))
                                                Else
                                                    strSpaces = ""
                                                End If
                                                strMiddleSubString = strMiddleSubString & strSpaces
                                                strLeftSubString = Left(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1)
                                                strSpaces = Space((MapRecordset.Fields("NCTS_IEM_MAP_StartPosition").Value - 1) - Len(strLeftSubString))
                                                strLeftSubString = strLeftSubString & strSpaces
                                                strRightSubString = Mid(strBoxCodeValue, MapRecordset.Fields("NCTS_IEM_MAP_Length").Value + 1)
                                                strBoxCodeValue = strLeftSubString & strMiddleSubString & strRightSubString
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If Not blnValueIsComplete Then
                                    MapRecordset.MoveNext
                                    blnValueIsComplete = MapRecordset.EOF
                                    If blnValueIsComplete Then
                                        MapRecordset.MoveFirst
                                    End If
                                End If
                                'blnContinueLoop = (Not blnValueIsComplete) And Not MapRecordset.EOF
                                If blnValueIsComplete Then
                                    If IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or strSegmentKeyCST = strSegmentKey Then
                                        ReDim Preserve astrSegmentValues(lngValuesCount)
                                        If blnIsNumeric Then
                                            If Trim(strBoxCodeValue) = vbNullString Then
                                                astrSegmentValues(lngValuesCount) = "0"
                                            Else
                                                astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                            End If
                                        Else
                                            astrSegmentValues(lngValuesCount) = strBoxCodeValue
                                        End If
                                        lngValuesCount = lngValuesCount + 1
                                    End If
                                    lngBoxCodeInstance = lngBoxCodeInstance + 1
                                End If
                                
                                blnContinueLoop = (EDIClass.GetSegmentIndex("S_" & CStr(MapRecordset.Fields("NCTS_IEM_TMS_ID").Value) & "_" & CStr(GetSegmentInstance(MapRecordset.Fields("NCTS_IEM_MAP_Source").Value, lngBoxCodeInstance))) > 0)
                            Loop Until Not blnContinueLoop
                        End If
                End Select
            Else
                Select Case ValueSource
                    Case "F<DETAIL COUNT>"
                        lngCST_NCTS_IEM_TMS_ID = 33
                        lngSegmentInstance = 0
                        strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance + 1)
                        Do While EDIClass.GetSegmentIndex(strSegmentKeyCST) > 0
                            lngSegmentInstance = lngSegmentInstance + 1
                            strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance + 1)
                        Loop
                        ReDim Preserve astrSegmentValues(0)
                        astrSegmentValues(0) = lngSegmentInstance
                    Case Else
                        Debug.Assert False
                End Select
            End If
        
        Case Else
            Select Case IE29Value
                Case enuIEVal_IE43_Marks_And_Numbers To enuIEVal_IE43_Commodity_Code
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE43(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
                Case enuIE29Val_MessageIdentification To enuIE29Val_ControlledBy
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE29(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
                Case enuIEVal_IE28_TPTIN To enuIEVal_IE28_TPCountry
                    Call SetNCTS_IEM_TMS_IDAndOrdinalIE28(IE29Value, lngNCTS_IEM_TMS_ID, lngNCTS_IEM_TMS_IDAlternative, lngDataItemOrdinal, lngDataItemsCount)
            End Select
            
    End Select
    
    'Dim lngIndex As Long
    
    Dim blnSegmentExists As Boolean
    
    If IE29Value <> enuIE29Val_NotFromIE29 Then
        Select Case GetTabTypeNonIE15(IE29Value)
            Case eTabType.eTab_Header
                lngValuesCount = 0
                lngSegmentInstance = 1
                blnIsNumeric = False
                Do
                    strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                    blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    blnSegmentInHeader = (blnSegmentExists And (lngNCTS_IEM_TMS_IDAlternative > 0))
                    If Not blnSegmentExists And lngNCTS_IEM_TMS_IDAlternative <> 0 Then
                        blnSegmentInHeader = False
                        lngNCTS_IEM_TMS_ID = lngNCTS_IEM_TMS_IDAlternative
                        strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                        blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    End If
                    If blnSegmentExists Then

                            strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                            blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                            ReDim Preserve astrSegmentValues(lngValuesCount)
                            If blnIsNumeric Then
                                If Trim(strBoxCodeValue) = vbNullString Then
                                    astrSegmentValues(lngValuesCount) = "0"
                                Else
                                    astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                End If
                            Else
                                astrSegmentValues(lngValuesCount) = strBoxCodeValue
                            End If
                            lngValuesCount = lngValuesCount + 1
                    End If
                    
                    lngSegmentInstance = lngSegmentInstance + 1
                Loop Until Not blnSegmentExists

                
                
            Case eTabType.eTab_Detail
                Select Case EDIClass.MessageType
                    Case ENCTSMessageType.EMsg_IE15
                        '----->  OOPS!!! restricted area hehehe. must not have reached this area
                        Debug.Assert False
                    Case ENCTSMessageType.EMsg_IE29
                        lngCST_NCTS_IEM_TMS_ID = 126
                    Case ENCTSMessageType.EMsg_IE43
                        lngCST_NCTS_IEM_TMS_ID = 260
                    Case ENCTSMessageType.EMsg_IE51
                        lngCST_NCTS_IEM_TMS_ID = 187
                        Debug.Assert False
                    Case ENCTSMessageType.EMsg_IE28
                        lngCST_NCTS_IEM_TMS_ID = 0
                    Case Else
                        Debug.Assert False
                End Select
                If lngCST_NCTS_IEM_TMS_ID = 0 Then
                    strSegmentKeyCST = ""
                Else
                    strSegmentKeyCST = "S_" & CStr(lngCST_NCTS_IEM_TMS_ID) & "_" & CStr(TabNumber)
                End If
                lngValuesCount = 0
                lngSegmentInstance = 1
                blnIsNumeric = False
                Do
                    strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                    blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    blnSegmentInHeader = (blnSegmentExists And (lngNCTS_IEM_TMS_IDAlternative > 0))
                    If Not blnSegmentExists And lngNCTS_IEM_TMS_IDAlternative <> 0 Then
                        blnSegmentInHeader = False
                        lngNCTS_IEM_TMS_ID = lngNCTS_IEM_TMS_IDAlternative
                        strSegmentKey = "S_" & CStr(lngNCTS_IEM_TMS_ID) & "_" & CStr(lngSegmentInstance)
                        blnSegmentExists = (EDIClass.GetSegmentIndex(strSegmentKey) > 0)
                    End If
                    If blnSegmentExists Then
                        If (IsDecendant(EDIClass, strSegmentKeyCST, strSegmentKey) Or (strSegmentKeyCST = strSegmentKey) Or (strSegmentKeyCST = "")) Or blnSegmentInHeader Then
                            strBoxCodeValue = EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).Value
                            blnIsNumeric = (Left(EDIClass.Segments(strSegmentKey).SDataItems(lngDataItemOrdinal).NCTSDataFormat, 1) = "N")
                            ReDim Preserve astrSegmentValues(lngValuesCount)
                            If blnIsNumeric Then
                                If Trim(strBoxCodeValue) = vbNullString Then
                                    astrSegmentValues(lngValuesCount) = "0"
                                Else
                                    astrSegmentValues(lngValuesCount) = Trim(strBoxCodeValue)
                                End If
                            Else
                                astrSegmentValues(lngValuesCount) = strBoxCodeValue
                            End If
                            lngValuesCount = lngValuesCount + 1
                        End If
                    End If
                    lngSegmentInstance = lngSegmentInstance + 1
                Loop Until Not blnSegmentExists
        End Select
    End If
    varReturnValue = astrSegmentValues
    GetValueFromClass = varReturnValue
End Function

Private Function GetValueFromClassAssertions(ByVal EDIClass As EdifactMessage, _
                                             ByVal IEMValue As IE29Values) As Boolean
    Select Case IEMValue
        Case enuIE29Val_NotFromIE29
            Debug.Assert EDIClass.MessageType = EMsg_IE15
        Case enuIEVal_IE43_Marks_And_Numbers, _
             enuIEVal_IE43_Number_of_Packages, enuIEVal_IE43_Kind_of_Packages, _
             enuIEVal_IE43_Container_Numbers, _
             enuIEVal_IE43_Description_of_Goods, _
             enuIEVal_IE43_Sensitivity_Code, enuIEVal_IE43_Sensitive_Quantity, _
             enuIEVal_IE43_Country_of_Dispatch_Export, enuIEVal_IE43_Country_of_Destination, _
             enuIEVal_IE43_CO_Departure, _
             enuIEVal_IE43_Gross_Mass, enuIEVal_IE43_Net_Mass, _
             enuIEVal_IE43_Additional_Information, _
             enuIEVal_IE43_Consignor_TIN, enuIEVal_IE43_Consignor_Name, enuIEVal_IE43_Consignor_Street_And_Number, enuIEVal_IE43_Consignor_Postal_Code, enuIEVal_IE43_Consignor_City, enuIEVal_IE43_Consignor_Country, _
             enuIEVal_IE43_Consignee_TIN, enuIEVal_IE43_Consignee_Name, enuIEVal_IE43_Consignee_Street_And_Number, enuIEVal_IE43_Consignee_Postal_Code, enuIEVal_IE43_Consignee_City, enuIEVal_IE43_Consignee_Country, _
             enuIEVal_IE43_Document_Type, enuIEVal_IE43_Document_Reference, enuIEVal_IE43_Document_Complement_Information, _
             enuIEVal_IE43_Detail_Number, _
             enuIEVal_IE43_Commodity_Code
            Debug.Assert EDIClass.MessageType = EMsg_IE43
        Case enuIE29Val_MessageIdentification, _
             enuIE29Val_ReferenceNumber, _
             enuIE29Val_AuthorizedLocationOfGoods, _
             enuIE29Val_DeclarationPlace, _
             enuIE29Val_COReferencNumber, enuIE29Val_COName, enuIE29Val_COCountry, enuIE29Val_COStreetAndNumber, enuIE29Val_COPostalCode, enuIE29Val_COCity, enuIE29Val_COLanguage, _
             enuIE29Val_DateApproval, enuIE29Val_DateIssuance, enuIE29Val_DateControl, enuIEVal_IE29_DateLimitTransit, _
             enuIE29Val_ReturnCopy, _
             enuIE29Val_BindingItinerary, _
             enuIE29Val_NotValidForEC, _
             enuIE29Val_TPName, enuIE29Val_TPStreetAndNumber, enuIE29Val_TPCity, enuIE29Val_TPPostalCode, enuIE29Val_TPCountry, _
             enuIE29Val_ControlledBy
            Debug.Assert EDIClass.MessageType = EMsg_IE29
        Case enuIEVal_IE28_TPTIN, enuIEVal_IE28_TPName, enuIEVal_IE28_TPStreetAndNumber, enuIEVal_IE28_TPCity, enuIEVal_IE28_TPPostalCode, enuIEVal_IE28_TPCountry
            Debug.Assert EDIClass.MessageType = EMsg_IE28
    End Select
End Function

Public Function GetTabType(ByVal DocumentTypeIdentifier As String, ByVal BoxCode As String) As eTabType
    Dim enuReturnValue As eTabType
    
    Select Case DocumentTypeIdentifier
        Case G_CONST_EDINCTS1_TYPE
            Select Case BoxCode
                Case "A4", "A5", "A6", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF", _
                     "B7", "B1", "B8", "B2", "B3", "B9", "BA", "B5", _
                     "C2", "C3", "C4", "C5", _
                     "X4", "X5", "X1", "X2", "X6", "X3", "X7", "X8", _
                     "E1", "EJ", "E3", "EK", "E4", "E5", "E6", "E7", "EM", "EN", "EO", "E8", "EA", "EC", "EE", "EG", "EI"
                    enuReturnValue = eTab_Header
                    
                Case "U6", "U2", "U3", "U4", "U8", "U7", _
                     "W6", "W7", "W1", "W2", "W4", "W3", "W5", _
                     "L7", "L1", "L8", _
                     "M1", "M2", "M9", _
                     "S1", "S2", "S4", "S3", "S5", "S6", "S7", "S8", "S9", "SA", "SB", _
                     "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", _
                     "Y1", "Y2", "Y3", "Y4", "Y5", _
                     "Z1", "Z2", "Z3", "Z4", _
                     "T7"
                    enuReturnValue = eTab_Detail
                    
                Case Else
                    '----->  Box code must either be found in either of the two categories above.
                    Debug.Assert False
            End Select
        Case Else
            '----->  No code yet
            Debug.Assert False
    End Select
    GetTabType = enuReturnValue
End Function

Private Function GetSegmentInstance(ByVal BoxCode As String, _
                                    ByVal BoxCodeInstance As Long) As Long
    Dim lngReturnValue As Long

    Select Case BoxCode
        Case "A4", "A5", "A6", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF"
            lngReturnValue = BoxCodeInstance
        Case "B7", "B1", "B8", "B2", "B3", "B9", "BA", "B5"
            lngReturnValue = BoxCodeInstance
        Case "C2", "C3", "C4", "C5"
            lngReturnValue = BoxCodeInstance
        Case "X4", "X5", "X1", "X2", "X6", "X3", "X7", "X8"
            lngReturnValue = BoxCodeInstance
        Case "E8", "EA", "EC", "EE", "EG", "EI"
            lngReturnValue = ((BoxCodeInstance - 1) * 6) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Header) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "E8", eTab_Header) + 1)
        Case "E1", "E3", "EK"
            lngReturnValue = BoxCodeInstance
        Case "E4", "E5", "E6", "E7", "EM", "EN"
            lngReturnValue = ((BoxCodeInstance - 1) * 6) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Header) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "E4", eTab_Header) + 1)
        Case "W1" To "W6", "U1" To "U8"
            lngReturnValue = BoxCodeInstance
        Case "V1", "V3", "V5", "V7"
            lngReturnValue = ((BoxCodeInstance - 1) * 4) + (((IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "V1", eTab_Detail)) / 2) + 1)
        Case "V2", "V4", "V6", "V8"
            lngReturnValue = ((BoxCodeInstance - 1) * 4) + (((IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "V2", eTab_Detail)) / 2) + 1)
        Case "L1", "L7", "L8", "M1", "M2", "M9"
            lngReturnValue = BoxCodeInstance
        Case "S1"
            lngReturnValue = BoxCodeInstance
        Case "S2", "S3", "S4"
            lngReturnValue = BoxCodeInstance
        Case "S6", "S7", "S8", "S9", "SA"
            lngReturnValue = ((BoxCodeInstance - 1) * 5) + (IndexTextNCTS(G_CONST_EDINCTS1_TYPE, BoxCode, eTab_Detail) - IndexTextNCTS(G_CONST_EDINCTS1_TYPE, "S6", eTab_Detail) + 1)
        Case "Y2", "Y3", "Y4"
            lngReturnValue = BoxCodeInstance
        Case "Z1", "Z2", "Z3"
            lngReturnValue = BoxCodeInstance
    End Select
    GetSegmentInstance = lngReturnValue
End Function

Private Function IsDecendant(ByVal EDIClass As EdifactMessage, ByVal RootSegmentKey As String, DecendantSegmentKey As String) As Boolean
    Dim blnReturnValue As Boolean
    
    Dim blnContinueSearch As Boolean
    
    blnReturnValue = False
    Debug.Assert EDIClass.GetSegmentIndex(DecendantSegmentKey) > 0
    If EDIClass.GetSegmentIndex(DecendantSegmentKey) > 0 Then
        If Trim(EDIClass.Segments(DecendantSegmentKey).KeyParent) <> vbNullString Then
            If EDIClass.Segments(DecendantSegmentKey).KeyParent = RootSegmentKey Then
                blnReturnValue = True
            Else
                blnReturnValue = IsDecendant(EDIClass, RootSegmentKey, EDIClass.Segments(DecendantSegmentKey).KeyParent)
            End If
        Else
            blnReturnValue = False
        End If
    End If
    IsDecendant = blnReturnValue
End Function

Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE43(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    NCTS_IEM_TMS_IDAlternative = 0
    Select Case IEMessageValue
        Case enuIEVal_IE43_Marks_And_Numbers                'PCI+28(2,3)
            NCTS_IEM_TMS_ID = 268
            DataItemOrdinal = 2
            DataItemsCount = 2
        Case enuIEVal_IE43_Number_of_Packages, enuIEVal_IE43_Kind_of_Packages
            NCTS_IEM_TMS_ID = 267
            Select Case IEMessageValue
                Case enuIEVal_IE43_Number_of_Packages               'PAC+6(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Kind_of_Packages                 'PAC+6(9)
                    DataItemOrdinal = 9
            End Select
        Case enuIEVal_IE43_Container_Numbers                'RFF+AAQ
            NCTS_IEM_TMS_ID = 269
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Description_of_Goods             'FTX+AAA(6,7,8,9)
            NCTS_IEM_TMS_ID = 261
            DataItemOrdinal = 6
            DataItemsCount = 4
        Case enuIEVal_IE43_Sensitivity_Code, enuIEVal_IE43_Sensitive_Quantity
            NCTS_IEM_TMS_ID = 273
            Select Case IEMessageValue
                Case enuIEVal_IE43_Sensitivity_Code                 'GIR+3+AP(5)
                    DataItemOrdinal = 5
                Case enuIEVal_IE43_Sensitive_Quantity               'GIR+3+AP(2)
                    DataItemOrdinal = 2
            End Select
        Case enuIEVal_IE43_Country_of_Dispatch_Export       'LOC+35(2)
            NCTS_IEM_TMS_ID = 246
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Country_of_Destination           'LOC+36(2)
            NCTS_IEM_TMS_ID = 247
            DataItemOrdinal = 2
        Case enuIEVal_IE43_CO_Departure                     'LOC+118(2)
            NCTS_IEM_TMS_ID = 244
            DataItemOrdinal = 2
        Case enuIEVal_IE43_Gross_Mass                       'MEA+WT+AAB+KGM(7)
            NCTS_IEM_TMS_ID = 263
            DataItemOrdinal = 7
        Case enuIEVal_IE43_Net_Mass                         'MEA+WT+AAA+KGM(7)
            NCTS_IEM_TMS_ID = 264
            DataItemOrdinal = 7
        Case enuIEVal_IE43_Additional_Information           'FTX+ACB(11)
            NCTS_IEM_TMS_ID = 272
            DataItemOrdinal = 11
        Case enuIEVal_IE43_Consignor_TIN, enuIEVal_IE43_Consignor_Name, enuIEVal_IE43_Consignor_Street_And_Number, _
             enuIEVal_IE43_Consignor_Postal_Code, enuIEVal_IE43_Consignor_City, enuIEVal_IE43_Consignor_Country
            NCTS_IEM_TMS_ID = 258
            NCTS_IEM_TMS_IDAlternative = 266
            Select Case IEMessageValue
                Case enuIEVal_IE43_Consignor_TIN                    'NAD+CZ(2)
                    DataItemOrdinal = 2
                Case enuIEVal_IE43_Consignor_Name                   'NAD+CZ(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Consignor_Street_And_Number      'NAD+CZ(16)
                    DataItemOrdinal = 16
                Case enuIEVal_IE43_Consignor_Postal_Code            'NAD+CZ(22)
                    DataItemOrdinal = 22
                Case enuIEVal_IE43_Consignor_City                   'NAD+CZ(20)
                    DataItemOrdinal = 20
                Case enuIEVal_IE43_Consignor_Country                'NAD+CZ(23)
                    DataItemOrdinal = 23
            End Select
        Case enuIEVal_IE43_Consignee_TIN, enuIEVal_IE43_Consignee_Name, enuIEVal_IE43_Consignee_Street_And_Number, _
             enuIEVal_IE43_Consignee_Postal_Code, enuIEVal_IE43_Consignee_City, enuIEVal_IE43_Consignee_Country
            NCTS_IEM_TMS_ID = 256
            NCTS_IEM_TMS_IDAlternative = 265
            Select Case IEMessageValue
                Case enuIEVal_IE43_Consignee_TIN                    'NAD+CN(2)
                    DataItemOrdinal = 2
                Case enuIEVal_IE43_Consignee_Name                   'NAD+CN(10)
                    DataItemOrdinal = 10
                Case enuIEVal_IE43_Consignee_Street_And_Number      'NAD+CN(16)
                    DataItemOrdinal = 16
                Case enuIEVal_IE43_Consignee_Postal_Code            'NAD+CN(22)
                    DataItemOrdinal = 22
                Case enuIEVal_IE43_Consignee_City                   'NAD+CN(20)
                    DataItemOrdinal = 20
                Case enuIEVal_IE43_Consignee_Country                'NAD+CN(23)
                    DataItemOrdinal = 23
            End Select
        Case enuIEVal_IE43_Document_Type, enuIEVal_IE43_Document_Reference, enuIEVal_IE43_Document_Complement_Information
            NCTS_IEM_TMS_ID = 270
            Select Case IEMessageValue
                Case enuIEVal_IE43_Document_Type                    'DOC+916(4)
                    DataItemOrdinal = 4
                Case enuIEVal_IE43_Document_Reference               'DOC+916(5)
                    DataItemOrdinal = 5
                Case enuIEVal_IE43_Document_Complement_Information  'DOC+916(7)
                    DataItemOrdinal = 7
            End Select
        Case enuIEVal_IE43_Detail_Number                    'CST(1)
            NCTS_IEM_TMS_ID = 260
            DataItemOrdinal = 1
        Case enuIEVal_IE43_Commodity_Code                   'CST(2)
            NCTS_IEM_TMS_ID = 260
            DataItemOrdinal = 2
        Case Else
            Debug.Assert False
    End Select
End Sub

Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE29(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    NCTS_IEM_TMS_IDAlternative = 0
    Select Case IEMessageValue
        Case enuIE29Val_MessageIdentification               'UNH(1)
            NCTS_IEM_TMS_ID = 83
            DataItemOrdinal = 1
        Case enuIE29Val_ReferenceNumber                     'BGM(5)
            NCTS_IEM_TMS_ID = 84
            DataItemOrdinal = 5
        Case enuIE29Val_AuthorizedLocationOfGoods           'LOC+14(6)
            NCTS_IEM_TMS_ID = 86
            DataItemOrdinal = 6
        Case enuIE29Val_DeclarationPlace                    'LOC+91(5)
            NCTS_IEM_TMS_ID = 93
            DataItemOrdinal = 5
        Case enuIE29Val_COReferencNumber, enuIE29Val_COName, enuIE29Val_COCountry, enuIE29Val_COStreetAndNumber, enuIE29Val_COPostalCode, enuIE29Val_COCity, enuIE29Val_COLanguage
            NCTS_IEM_TMS_ID = 87
            Select Case IEMessageValue
                Case enuIE29Val_COReferencNumber                    'LOC+168(2) - CO = Customs Office
                    DataItemOrdinal = 2
                Case enuIE29Val_COName                              'LOC+168(5) - CO = Customs Office
                    DataItemOrdinal = 5
                Case enuIE29Val_COCountry                           'LOC+168(6) - CO = Customs Office
                    DataItemOrdinal = 6
                Case enuIE29Val_COStreetAndNumber                   'LOC+168(9) - CO = Customs Office
                    DataItemOrdinal = 9
                Case enuIE29Val_COPostalCode                        'LOC+168(10) - CO = Customs Office
                    DataItemOrdinal = 10
                Case enuIE29Val_COCity                              'LOC+168(13) - CO = Customs Office
                    DataItemOrdinal = 13
                Case enuIE29Val_COLanguage                          'LOC+168(14) - CO = Customs Office
                    DataItemOrdinal = 14
            End Select
        Case enuIE29Val_DateApproval                         'DTM+148(2)
            NCTS_IEM_TMS_ID = 95
            DataItemOrdinal = 2
        Case enuIE29Val_DateIssuance                         'DTM+182(2)
            NCTS_IEM_TMS_ID = 96
            DataItemOrdinal = 2
        Case enuIE29Val_DateControl                          'DTM+9(2)
            NCTS_IEM_TMS_ID = 98
            DataItemOrdinal = 2
        Case enuIEVal_IE29_DateLimitTransit                  'DTM+268(2)
            NCTS_IEM_TMS_ID = 97
            DataItemOrdinal = 2
        Case enuIE29Val_ReturnCopy                           'GIS 62(2)
            NCTS_IEM_TMS_ID = 100
            DataItemOrdinal = 2
        Case enuIE29Val_BindingItinerary                     'FTX+ABL(6)
            NCTS_IEM_TMS_ID = 103
            DataItemOrdinal = 6
        Case enuIE29Val_NotValidForEC                        'PCI+19(2)
            NCTS_IEM_TMS_ID = 112
            DataItemOrdinal = 2
        Case enuIE29Val_TPName, enuIE29Val_TPStreetAndNumber, enuIE29Val_TPCity, enuIE29Val_TPPostalCode, enuIE29Val_TPCountry
            NCTS_IEM_TMS_ID = 119
            Select Case IEMessageValue
                Case enuIE29Val_TPName                               'NAD+AF(10) - TP = Transit Principal
                    DataItemOrdinal = 10
                Case enuIE29Val_TPStreetAndNumber                    'NAD+AF(16) - TP = Transit Principal
                    DataItemOrdinal = 16
                Case enuIE29Val_TPCity                               'NAD+AF(20) - TP = Transit Principal
                    DataItemOrdinal = 20
                Case enuIE29Val_TPPostalCode                         'NAD+AF(22) - TP = Transit Principal
                    DataItemOrdinal = 22
                Case enuIE29Val_TPCountry                            'NAD+AF(23) - TP = Transit Principal
                    DataItemOrdinal = 23
            End Select
        Case enuIE29Val_ControlledBy                         'NAD+EI(2)
            NCTS_IEM_TMS_ID = 123
            DataItemOrdinal = 2
    End Select
End Sub

Private Sub SetNCTS_IEM_TMS_IDAndOrdinalIE28(ByVal IEMessageValue As IE29Values, _
                                             ByRef NCTS_IEM_TMS_ID As Long, _
                                             ByRef NCTS_IEM_TMS_IDAlternative As Long, _
                                             ByRef DataItemOrdinal As Long, _
                                             ByRef DataItemsCount As Long)
    NCTS_IEM_TMS_IDAlternative = 0
    NCTS_IEM_TMS_ID = 78
    Select Case IEMessageValue
        Case enuIEVal_IE28_TPTIN                             'NAD+AF(2)  - TP = Transit Principal
            DataItemOrdinal = 2
        Case enuIEVal_IE28_TPName                            'NAD+AF(10) - TP = Transit Principal
            DataItemOrdinal = 10
        Case enuIEVal_IE28_TPStreetAndNumber                 'NAD+AF(16) - TP = Transit Principal
            DataItemOrdinal = 16
        Case enuIEVal_IE28_TPCity                            'NAD+AF(20) - TP = Transit Principal
            DataItemOrdinal = 20
        Case enuIEVal_IE28_TPPostalCode                      'NAD+AF(22) - TP = Transit Principal
            DataItemOrdinal = 22
        Case enuIEVal_IE28_TPCountry                         'NAD+AF(23) - TP = Transit Principal
            DataItemOrdinal = 23
    End Select
End Sub

Public Function GetTabTypeNonIE15(IEMessageValue As IE29Values) As eTabType
    Dim enuReturnValue As eTabType
    
    Select Case IEMessageValue
        Case enuIEVal_IE43_Marks_And_Numbers
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Number_of_Packages
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Kind_of_Packages
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Container_Numbers
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Description_of_Goods
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Sensitivity_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Sensitive_Quantity
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Country_of_Dispatch_Export
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Country_of_Destination
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_CO_Departure
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Gross_Mass
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Net_Mass
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Additional_Information
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_TIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Name
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Street_And_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Postal_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_City
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignor_Country
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_TIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Name
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Street_And_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Postal_Code
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_City
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Consignee_Country
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Type
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Reference
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Document_Complement_Information
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Detail_Number
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE43_Commodity_Code
            enuReturnValue = eTab_Detail
        
        Case enuIE29Val_MessageIdentification
            enuReturnValue = eTab_Detail
        Case enuIE29Val_ReferenceNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_AuthorizedLocationOfGoods
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DeclarationPlace
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COReferencNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COName
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COCountry
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COPostalCode
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COCity
            enuReturnValue = eTab_Detail
        Case enuIE29Val_COLanguage
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateApproval
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateIssuance
            enuReturnValue = eTab_Detail
        Case enuIE29Val_DateControl
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE29_DateLimitTransit
            'enuReturnValue = eTab_Detail
            enuReturnValue = eTab_Header
        Case enuIE29Val_ReturnCopy
            enuReturnValue = eTab_Detail
        Case enuIE29Val_BindingItinerary
            enuReturnValue = eTab_Detail
        Case enuIE29Val_NotValidForEC
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPName
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPCity
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPPostalCode
            enuReturnValue = eTab_Detail
        Case enuIE29Val_TPCountry
            enuReturnValue = eTab_Detail
        Case enuIE29Val_ControlledBy
            enuReturnValue = eTab_Detail
        
        Case enuIEVal_IE28_TPTIN
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPName
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPStreetAndNumber
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPCity
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPPostalCode
            enuReturnValue = eTab_Detail
        Case enuIEVal_IE28_TPCountry
            enuReturnValue = eTab_Detail
    End Select
    GetTabTypeNonIE15 = enuReturnValue
End Function

Public Function IndexTextNCTS(ByVal DocumentTypeIdentifier As String, _
                                ByVal BoxCode As String, _
                                ByVal TabType As eTabType, Optional ByRef blnNotFound As Boolean) As Integer
        
    Select Case DocumentTypeIdentifier
        Case G_CONST_NCTS1_TYPE, G_CONST_EDINCTS1_TYPE

                If TabType = eTab_Header Then
                    Select Case BoxCode
                        Case "A4": IndexTextNCTS = 0
                        Case "A5": IndexTextNCTS = 1
                        Case "A6": IndexTextNCTS = 2
                        Case "A8": IndexTextNCTS = 3
                        Case "A9": IndexTextNCTS = 4
                        Case "AA": IndexTextNCTS = 5
                        Case "AB": IndexTextNCTS = 6
                        Case "AC": IndexTextNCTS = 7
                        Case "AD": IndexTextNCTS = 8
                        Case "AE": IndexTextNCTS = 9
                        Case "AF": IndexTextNCTS = 10
                        
                        Case "B7": IndexTextNCTS = 11
                        Case "B1": IndexTextNCTS = 12
                        Case "B8": IndexTextNCTS = 13
                        Case "B2": IndexTextNCTS = 14
                        Case "B3": IndexTextNCTS = 15
                        Case "B9": IndexTextNCTS = 16
                        Case "BA": IndexTextNCTS = 17
                        Case "B5": IndexTextNCTS = 18
                                    
                        Case "C2": IndexTextNCTS = 19
                        Case "C3": IndexTextNCTS = 20
                        Case "C4": IndexTextNCTS = 21
                        Case "C5": IndexTextNCTS = 22
                    
                        Case "X4": IndexTextNCTS = 23
                        Case "X5": IndexTextNCTS = 24
                        Case "X1": IndexTextNCTS = 25
                        Case "X2": IndexTextNCTS = 26
                        Case "X6": IndexTextNCTS = 27
                        Case "X3": IndexTextNCTS = 28
                        Case "X7": IndexTextNCTS = 29
                        Case "X8": IndexTextNCTS = 30
            
                        Case "E1": IndexTextNCTS = 31
                        Case "EJ": IndexTextNCTS = 32
                        Case "E3": IndexTextNCTS = 33
                        Case "EK": IndexTextNCTS = 34
                        Case "E4": IndexTextNCTS = 35
                        Case "E5": IndexTextNCTS = 36
                        Case "E6": IndexTextNCTS = 37
                        Case "E7": IndexTextNCTS = 38
                        Case "EM": IndexTextNCTS = 39
                        Case "EN": IndexTextNCTS = 40
                        Case "EO": IndexTextNCTS = 41
                        Case "E8": IndexTextNCTS = 42
                        Case "EA": IndexTextNCTS = 43
                        Case "EC": IndexTextNCTS = 44
                        Case "EE": IndexTextNCTS = 45
                        Case "EG": IndexTextNCTS = 46
                        Case "EI": IndexTextNCTS = 47
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                        
                    End Select
                Else
                    Select Case BoxCode
                        Case "U6": IndexTextNCTS = 0
                        Case "U2": IndexTextNCTS = 1
                        Case "U3": IndexTextNCTS = 2
                        Case "U4": IndexTextNCTS = 3
                        Case "U8": IndexTextNCTS = 4
                        Case "U7": IndexTextNCTS = 5
                        
                        Case "W6": IndexTextNCTS = 6
                        Case "W7": IndexTextNCTS = 7
                        Case "W1": IndexTextNCTS = 8
                        Case "W2": IndexTextNCTS = 9
                        Case "W4": IndexTextNCTS = 10
                        Case "W3": IndexTextNCTS = 11
                        Case "W5": IndexTextNCTS = 12
                        
                        Case "L7": IndexTextNCTS = 13
                        Case "L1": IndexTextNCTS = 14
                        Case "L8": IndexTextNCTS = 15
                        
                        Case "M1": IndexTextNCTS = 16
                        Case "M2": IndexTextNCTS = 17
                        Case "M9": IndexTextNCTS = 18
                        
                        Case "S1": IndexTextNCTS = 19
                        Case "S2": IndexTextNCTS = 20
                        Case "S4": IndexTextNCTS = 21
                        Case "S3": IndexTextNCTS = 22
                        Case "S5": IndexTextNCTS = 23
                        Case "S6": IndexTextNCTS = 24
                        Case "S7": IndexTextNCTS = 25
                        Case "S8": IndexTextNCTS = 26
                        Case "S9": IndexTextNCTS = 27
                        Case "SA": IndexTextNCTS = 28
                        Case "SB": IndexTextNCTS = 29
                        
                        Case "V1": IndexTextNCTS = 30
                        Case "V2": IndexTextNCTS = 31
                        Case "V3": IndexTextNCTS = 32
                        Case "V4": IndexTextNCTS = 33
                        Case "V5": IndexTextNCTS = 34
                        Case "V6": IndexTextNCTS = 35
                        Case "V7": IndexTextNCTS = 36
                        Case "V8": IndexTextNCTS = 37
            
                        Case "Y1": IndexTextNCTS = 38
                        Case "Y2": IndexTextNCTS = 39
                        Case "Y3": IndexTextNCTS = 40
                        Case "Y4": IndexTextNCTS = 41
                        Case "Y5": IndexTextNCTS = 42
                    
                        Case "Z1": IndexTextNCTS = 43
                        Case "Z2": IndexTextNCTS = 44
                        Case "Z3": IndexTextNCTS = 45
                        Case "Z4": IndexTextNCTS = 46
                    
                        Case "T7": IndexTextNCTS = 47
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                    End Select
                End If
                

        Case G_CONST_NCTS2_TYPE

                If TabType = eTab_Header Then
                    Select Case BoxCode
                        Case "A1": IndexTextNCTS = 0
                        Case "A2": IndexTextNCTS = 1
                        Case "A4": IndexTextNCTS = 2
                        Case "A5": IndexTextNCTS = 3
                        Case "A6": IndexTextNCTS = 4
                        Case "A7": IndexTextNCTS = 5
                        Case "A8": IndexTextNCTS = 6
                        Case "A9": IndexTextNCTS = 7
                        Case "AA": IndexTextNCTS = 8
                        Case "AB": IndexTextNCTS = 9
                        Case "AC": IndexTextNCTS = 10
                        Case "AD": IndexTextNCTS = 11
                        Case "AE": IndexTextNCTS = 12
                        Case "AF": IndexTextNCTS = 13
                    
                        Case "B7": IndexTextNCTS = 14
                        Case "B1": IndexTextNCTS = 15
                        Case "B8": IndexTextNCTS = 16
                        Case "B2": IndexTextNCTS = 17
                        Case "B3": IndexTextNCTS = 18
                        Case "B9": IndexTextNCTS = 19
                        Case "B4": IndexTextNCTS = 20
                        Case "BA": IndexTextNCTS = 21
                        Case "B5": IndexTextNCTS = 22
                        Case "B6": IndexTextNCTS = 23
                                                           
                        Case "C1": IndexTextNCTS = 24
                        Case "C2": IndexTextNCTS = 25
                        Case "C3": IndexTextNCTS = 26
                        Case "C4": IndexTextNCTS = 27
                        Case "C5": IndexTextNCTS = 28
                    
                        Case "D1": IndexTextNCTS = 29
                        Case "D2": IndexTextNCTS = 30
                        Case "D3": IndexTextNCTS = 31
                        Case "D4": IndexTextNCTS = 32
                        Case "D5": IndexTextNCTS = 33
                        Case "D6": IndexTextNCTS = 34
                        Case "D7": IndexTextNCTS = 35
                    
                        Case "F1": IndexTextNCTS = 36
                        Case "G1": IndexTextNCTS = 37
                        Case "H1": IndexTextNCTS = 38
                        Case "J1": IndexTextNCTS = 39
                        Case "F2": IndexTextNCTS = 40
                        Case "G2": IndexTextNCTS = 41
                        Case "H2": IndexTextNCTS = 42
                        Case "J2": IndexTextNCTS = 43
                        Case "F3": IndexTextNCTS = 44
                        Case "G3": IndexTextNCTS = 45
                        Case "H3": IndexTextNCTS = 46
                        Case "J3": IndexTextNCTS = 47
                    
                        Case "K1": IndexTextNCTS = 48
                        Case "K2": IndexTextNCTS = 49
                        Case "K3": IndexTextNCTS = 50
                        Case "K4": IndexTextNCTS = 51
                        Case "K5": IndexTextNCTS = 52
                        Case "K6": IndexTextNCTS = 53
                    
                        Case "X4": IndexTextNCTS = 54
                        Case "X5": IndexTextNCTS = 55
                        Case "X1": IndexTextNCTS = 56
                        Case "X2": IndexTextNCTS = 57
                        Case "X6": IndexTextNCTS = 58
                        Case "X3": IndexTextNCTS = 59
                        Case "X7": IndexTextNCTS = 60
                        Case "X8": IndexTextNCTS = 61
            
                        Case "E1": IndexTextNCTS = 62
                        Case "EJ": IndexTextNCTS = 63
                        Case "E3": IndexTextNCTS = 64
                        Case "EK": IndexTextNCTS = 65
                        Case "E4": IndexTextNCTS = 66
                        Case "E5": IndexTextNCTS = 67
                        Case "E6": IndexTextNCTS = 68
                        Case "E7": IndexTextNCTS = 69
                        Case "EM": IndexTextNCTS = 70
                        Case "EN": IndexTextNCTS = 71
                        Case "EO": IndexTextNCTS = 72
                        Case "E8": IndexTextNCTS = 73
                        Case "EA": IndexTextNCTS = 74
                        Case "EC": IndexTextNCTS = 75
                        Case "EE": IndexTextNCTS = 76
                        Case "EG": IndexTextNCTS = 77
                        Case "EI": IndexTextNCTS = 78
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                    End Select
                Else
                    Select Case BoxCode
                        Case "U6": IndexTextNCTS = 0
                        Case "U7": IndexTextNCTS = 1
                        Case "U2": IndexTextNCTS = 2
                        Case "U3": IndexTextNCTS = 3
                        Case "U4": IndexTextNCTS = 4
                        Case "U8": IndexTextNCTS = 5
                        Case "U5": IndexTextNCTS = 6

                        Case "W6": IndexTextNCTS = 7
                        Case "W7": IndexTextNCTS = 8
                        Case "W1": IndexTextNCTS = 9
                        Case "W2": IndexTextNCTS = 10
                        Case "W4": IndexTextNCTS = 11
                        Case "W3": IndexTextNCTS = 12
                        Case "W5": IndexTextNCTS = 13
                                            
                        Case "L1": IndexTextNCTS = 14
                        Case "L2": IndexTextNCTS = 15
                        Case "L3": IndexTextNCTS = 16
                        Case "L4": IndexTextNCTS = 17
                        Case "L5": IndexTextNCTS = 18
                        Case "L6": IndexTextNCTS = 19
                        Case "L8": IndexTextNCTS = 20
                        
                        Case "M1": IndexTextNCTS = 21
                        Case "M2": IndexTextNCTS = 22
                        Case "M9": IndexTextNCTS = 23
                        Case "M3": IndexTextNCTS = 24
                        Case "M4": IndexTextNCTS = 25
                        Case "M5": IndexTextNCTS = 26
                    
                        Case "S1": IndexTextNCTS = 27
                        Case "S2": IndexTextNCTS = 28
                        Case "S4": IndexTextNCTS = 29
                        Case "S3": IndexTextNCTS = 30
                        Case "S5": IndexTextNCTS = 31
                        Case "S6": IndexTextNCTS = 32
                        Case "S7": IndexTextNCTS = 33
                        Case "S8": IndexTextNCTS = 34
                        Case "S9": IndexTextNCTS = 35
                        Case "SA": IndexTextNCTS = 36
                        Case "SB": IndexTextNCTS = 37
                    
                        Case "V1": IndexTextNCTS = 38
                        Case "V2": IndexTextNCTS = 39
                        Case "V3": IndexTextNCTS = 40
                        Case "V4": IndexTextNCTS = 41
                        Case "V5": IndexTextNCTS = 42
                        Case "V6": IndexTextNCTS = 43
                        Case "V7": IndexTextNCTS = 44
                        Case "V8": IndexTextNCTS = 45

                        Case "Y1": IndexTextNCTS = 46
                        Case "Y2": IndexTextNCTS = 47
                        Case "Y3": IndexTextNCTS = 48
                        Case "Y4": IndexTextNCTS = 49
                        Case "Y5": IndexTextNCTS = 50
                    
                        Case "Z1": IndexTextNCTS = 51
                        Case "Z2": IndexTextNCTS = 52
                        Case "Z3": IndexTextNCTS = 53
                        Case "Z4": IndexTextNCTS = 54

                        Case "M6": IndexTextNCTS = 55
                        Case "M7": IndexTextNCTS = 56
                        Case "M8": IndexTextNCTS = 57
                
                        Case "N1": IndexTextNCTS = 58
                        Case "O1": IndexTextNCTS = 59
                        Case "P1": IndexTextNCTS = 60
                        Case "Q1": IndexTextNCTS = 61
                        Case "N2": IndexTextNCTS = 62
                        Case "O2": IndexTextNCTS = 63
                        Case "P2": IndexTextNCTS = 64
                        Case "Q2": IndexTextNCTS = 65
                        Case "N3": IndexTextNCTS = 66
                        Case "O3": IndexTextNCTS = 67
                        Case "P3": IndexTextNCTS = 68
                        Case "Q3": IndexTextNCTS = 69
                    
                        Case "R1": IndexTextNCTS = 70
                        Case "R2": IndexTextNCTS = 71
                        Case "R3": IndexTextNCTS = 72
                        Case "R4": IndexTextNCTS = 73
                        Case "R5": IndexTextNCTS = 74
                        Case "R6": IndexTextNCTS = 75
                        Case "R7": IndexTextNCTS = 76
                        Case "R8": IndexTextNCTS = 77
                        Case "R9": IndexTextNCTS = 78
                        Case "RA": IndexTextNCTS = 79
                        
                        Case "T1": IndexTextNCTS = 80
                        Case "T2": IndexTextNCTS = 81
                        Case "T3": IndexTextNCTS = 82
                        Case "T4": IndexTextNCTS = 83
                        Case "T5": IndexTextNCTS = 84
                        Case "T6": IndexTextNCTS = 85
                        Case "T7": IndexTextNCTS = 86
                        
                        Case Else
                            blnNotFound = True
                            Debug.Assert False
                    End Select
                End If
    End Select
End Function

Public Sub GetEIE15Data(ByVal lngData_NCTS_ID As Long, _
                        ByVal strCode As String, _
                        ByRef cpiEDI As EdifactMessage, ByRef rstIEMap As ADODB.Recordset)
    
    Dim strSQL As String
    Dim strStatus As String
    Dim rstEDI As ADODB.Recordset
    
    strStatus = "Document"
    
    'cpiEDI.DBLocation = NoBackSlash(g_objDataSourceProperties.TracefilePath)
    'cpiEDI.DBName = "EDIFACT.MDB"
    cpiEDI.ConnectEDIDB g_objDataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    
    
    strSQL = "SELECT TOP 1 DATA_NCTS.[LOGID DESCRIPTION],DATA_NCTS.[DATE LAST MODIFIED],DATA_NCTS.[DOCUMENT NAME], " & _
                " DATA_NCTS.USERNAME, DATA_NCTS.MRN, DATA_NCTS.[DATE SEND],DATA_NCTS.CODE, " & _
                " DATA_NCTS_MESSAGES.DATA_NCTS_MSG_ID AS DATA_NCTS_MSG_ID FROM DATA_NCTS INNER JOIN " & _
                " DATA_NCTS_MESSAGES ON DATA_NCTS.DATA_NCTS_ID = DATA_NCTS_MESSAGES.DATA_NCTS_ID " & _
                " WHERE DATA_NCTS_MESSAGES.DATA_NCTS_MSG_StatusType = 'Document' AND " & _
                " DATA_NCTS_MESSAGES.NCTS_IEM_ID = 5 AND DATA_NCTS.DATA_NCTS_ID = " & lngData_NCTS_ID & " ORDER " & _
                " BY DATA_NCTS_MESSAGES.DATA_NCTS_MSG_DATE desc "
    
    ADORecordsetOpen strSQL, m_conEDIFACT, rstEDI, adOpenKeyset, adLockOptimistic

    With rstEDI
        If Not (.EOF And .BOF) Then
            .MoveFirst

            cpiEDI.PrepareMessageFromDatabase !DATA_NCTS_MSG_ID, strCode
        End If
    End With
    
    OpenMapRecordset m_conEDIFACT, rstIEMap
    
End Sub

Public Sub OpenMapRecordset(ByRef ADOActiveConnection As ADODB.Connection, ByRef MapRecordset As ADODB.Recordset)
    Dim SQLMap As String
    
        SQLMap = "SELECT * FROM NCTS_IEM_MAP WHERE NCTS_IEM_ID = " & CStr(NCTS_IEM_ID_IE15)
    ADORecordsetOpen SQLMap, ADOActiveConnection, MapRecordset, adOpenKeyset, adLockOptimistic
    'MapRecordset.Open SQLMap, ADOActiveConnection, adOpenKeyset, adLockReadOnly
End Sub

Private Function GetValue(ByVal In_Code As String, In_ID As String, Handling As Long) As Double
    Dim strSQL As String
    Dim dblOutbound As Double
    Dim rst As ADODB.Recordset
    Dim cpiEDI As EdifactMessage
    Dim rstIEMMap As ADODB.Recordset
    
    strSQL = "SELECT *, [DATA_NCTS_DETAIL].DETAIL AS DETAIL FROM [DATA_NCTS] " & _
            " INNER JOIN ([DATA_NCTS_HEADER] " & _
            " INNER JOIN [DATA_NCTS_DETAIL] ON " & _
            " ([DATA_NCTS_HEADER].HEADER = [DATA_NCTS_DETAIL].HEADER) AND ([DATA_NCTS_HEADER].CODE = [DATA_NCTS_DETAIL].CODE)) " & _
            " ON [DATA_NCTS].CODE = [DATA_NCTS_HEADER].CODE" & _
            " WHERE [DATA_NCTS].CODE = '" & In_Code & "'" & _
            " AND [DATA_NCTS_DETAIL].IN_ID = " & In_ID & _
            " ORDER BY [DATA_NCTS_DETAIL].DETAIL"
    ADORecordsetOpen strSQL, m_conEDIFACT, rst, adOpenKeyset, adLockOptimistic
    'rst.Open strSQL, m_conEDIFACT, adOpenForwardOnly, adLockReadOnly
    
    dblOutbound = 0
    
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do While Not rst.EOF
            If rst!In_ID <> 0 Then
                Select Case Handling
                    Case 0
                        Set cpiEDI = New EdifactMessage
                        
                        GetEIE15Data rst!Data_NCTS_ID, In_Code, cpiEDI, rstIEMMap
                        dblOutbound = dblOutbound + CDbl(Val(GetValueFromClass(cpiEDI, rstIEMMap, enuIE29Val_NotFromIE29, "S3", rst!Detail)(0)))
                        
                        Set cpiEDI = Nothing
                        
                        ADORecordsetClose rstIEMMap
                        
                    Case 1
                        Set cpiEDI = New EdifactMessage
                        
                        GetEIE15Data rst!Data_NCTS_ID, In_Code, cpiEDI, rstIEMMap
                        dblOutbound = dblOutbound + CDbl(Val(GetValueFromClass(cpiEDI, rstIEMMap, enuIE29Val_NotFromIE29, "M1", rst!Detail)(0)))
                    
                        Set cpiEDI = Nothing
                        
                        ADORecordsetClose rstIEMMap
                    
                    Case 2
                                    
                        Set cpiEDI = New EdifactMessage
                        
                        GetEIE15Data rst!Data_NCTS_ID, In_Code, cpiEDI, rstIEMMap
                        dblOutbound = dblOutbound + CDbl(Val(GetValueFromClass(cpiEDI, rstIEMMap, enuIE29Val_NotFromIE29, "M2", rst!Detail)(0)))
                
                        Set cpiEDI = Nothing
                        
                        ADORecordsetClose rstIEMMap
                
                End Select
            End If
            rst.MoveNext
        Loop
    End If
    
    GetValue = dblOutbound
    
    ADORecordsetClose rst
End Function

Private Sub jgxGrid_Validate(Cancel As Boolean)
    jgxGrid.Update
End Sub
