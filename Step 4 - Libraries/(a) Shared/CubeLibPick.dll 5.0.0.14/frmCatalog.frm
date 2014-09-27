VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCatalog 
   Caption         =   "Picklist"
   ClientHeight    =   5745
   ClientLeft      =   375
   ClientTop       =   3240
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSeeOnLine 
      Caption         =   "See List &Online"
      Height          =   375
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   8
      Left            =   8760
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   7
      Left            =   8040
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   6
      Left            =   7320
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   5
      Left            =   6600
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2460
      TabIndex        =   37
      Top             =   3780
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2460
      TabIndex        =   36
      Top             =   3300
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2460
      TabIndex        =   35
      Top             =   2820
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2460
      TabIndex        =   34
      Top             =   2340
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3660
      TabIndex        =   39
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   2
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   1
      Left            =   1380
      TabIndex        =   11
      Top             =   120
      Width           =   3000
   End
   Begin VB.TextBox txtSearchFields 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1300
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   4
      Left            =   6240
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   2
      Left            =   6240
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter 5"
      Height          =   195
      Index           =   4
      Left            =   6240
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter 4"
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter 3"
      Height          =   195
      Index           =   2
      Left            =   6240
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter 2"
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter 1"
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   25
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1850
      TabIndex        =   24
      Top             =   4680
      Width           =   1215
   End
   Begin GridEX16.GridEX jgxPicklist 
      Height          =   4095
      Left            =   45
      TabIndex        =   19
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7223
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnCount     =   2
      CardCaption1    =   -1  'True
      ColWidth1       =   1305
      ColEditType1    =   0
      ColWidth2       =   2745
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
   End
   Begin VB.CommandButton cmdCatalogOps 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   23
      Top             =   2635
      Width           =   1215
   End
   Begin VB.CommandButton cmdCatalogOps 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   20
      Top             =   1150
      Width           =   1215
   End
   Begin VB.CommandButton cmdCatalogOps 
      Caption         =   "&Modify"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   21
      Top             =   1645
      Width           =   1215
   End
   Begin VB.CommandButton cmdCatalogOps 
      Caption         =   "&Copy"
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   22
      Top             =   2140
      Width           =   1215
   End
   Begin TabDlg.SSTab tabCatalog 
      Height          =   4935
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "C&lients"
      TabPicture(0)   =   "frmCatalog.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblListDescription"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFilter(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFilter(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFilter(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFilter(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFilter(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "EcmdFilter"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdFilter"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkClearFilter"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CheckBox chkClearFilter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         Picture         =   "frmCatalog.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Clear Filter"
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   315
         Left            =   4680
         Picture         =   "frmCatalog.frx":0166
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Apply Filter"
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox EcmdFilter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         Picture         =   "frmCatalog.frx":02B0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter 1"
         Height          =   210
         Index           =   4
         Left            =   4560
         TabIndex        =   32
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter 1"
         Height          =   210
         Index           =   3
         Left            =   4560
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter 1"
         Height          =   210
         Index           =   2
         Left            =   4560
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter 1"
         Height          =   210
         Index           =   1
         Left            =   4560
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter 1"
         Height          =   210
         Index           =   0
         Left            =   4560
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblListDescription 
         Caption         =   "The list below shows all existing Clients."
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Label lblRecord 
      AutoSize        =   -1  'True
      Caption         =   "Record 1 of n"
      Height          =   195
      Left            =   2460
      TabIndex        =   33
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
' constants
Private Const CMD_OK = 0
Private Const CMD_CANCEL = 1
Private Const CMD_SEELISTONLINE = 2

Private Const LEFT_ALIGN = 0
Private Const RIGHT_ALIGN = 1
Private Const CENTER_ALIGN = 2

Private Const FIRST_FILTER = 0
Private Const SECOND_FILTER = 1
Private Const THIRD_FILTER = 2
Private Const FOURTH_FILTER = 3

Private Const CPI_TRUE = -1
Private Const CPI_FALSE = 0
Private Const CPI_AUTOCANCEL = 1

' windows metrics
Private Const STD_BORDER_WIDTH = 15
Private Const STD_CAPTION_HEIGHT = 270
Private Const STD_MENU_HEIGHT = 270

' user enums
Private Enum RefreshType
    cpiRequery = 0
    cpiRefresh = 1
End Enum

' enums for task commandbuttons
Private Enum enuTask
    CMD_ADD = 0
    CMD_MODIFY = 1
    CMD_COPY = 2
    CMD_DELETE = 3
    
End Enum

' enum constants
Private cpiActiveStatus As cpiActiveStatusConstants
Private enuStyle As PicklistStyle
Private jgexHitTestConstant As jgexHitTestConstants

' native variables

' user - jgxPicklist_ColumnHeaderClick, FormInitialize
Dim lngTempAutoNoCtr As Long
    ' user - ShowPicklist, InitAddTrans
Dim lngAbsolutePosition As Long
    ' user - cmdCatalogOps_Click, Form_Activate, Form_Unload, RefreshGrid, jgxPicklist_Click,
    '               jgxPicklist_DblClick, jgxPicklist_KeyUp, jgxPicklist_RowColChange, TransactAddCopy
    '               blnReconcileAbsPos, LoadGrid, SetSelectedRecord
Dim blnDBAutoNumber As Boolean
    ' user - cmdCatalogOps_Click, CommitTransaction, TransactAddCopy, ProcessRecord, ApplyFetchGrid
    '               InitAddTrans, GetTopOne
Dim blnChangeDueToGridClick As Boolean
    ' user - txtSearchFields_Change, SetHeaderTextBoxes, FormInitialize
Dim blnGridIsEmpty As Boolean
    ' user - RefreshGrid , txtSearchFields_Click, FormInitialize, SetHeaderTextBoxes
Dim blnFormActivated As Boolean
    ' user - Form_Activate,optFilter_Click,ApplyGridFilter,FormInitialize,chkClearFilter_Click,chkFilter_Click,
Dim blnCanceled As Boolean
    ' user - ShowPicklist , cmdOK_Click, cmdTransact_Click, Form_Unload, jgxPicklist_DblClick,
    '            ProcessRecord , RetrieveRecord, FormInitialize, SetSelectedRecord
Dim blnInitEnd As Boolean
    ' user - ShowPicklist , PopulateGrid, ApplyFetchGrid
Dim blnTaskClick As Boolean
    ' user - cmdCatalogOps_Click , ApplyGridFilter
Dim blnIsCheck As Boolean
    ' user - chkFilter_Click , ApplyGridFilter
Dim blnCloseBtnClick As Boolean
    ' user - cmdTransact_Click , Form_Unload
Dim blnSkipTextChange As Boolean
    ' user - Form_Activate , RefreshGrid, txtSearchFields_Change, UpdateSearchField
Dim blnIsInTrans As Boolean
    ' user - cmdCatalogOps_Click , Form_Unload, RunDeleteTrans, SetSelectedRecord
Dim blnHaltExecution As Boolean
    ' user - ShowPicklist , RefreshGrid, FetchGridRecords, CreateRstToGrid, CheckAutoSearch

Dim varTempDBID As Variant
    ' user - cmdCatalogOps_Click , blnReconcileAbsPos, InitAddTrans, RunDeleteTrans
' critical variables for PK tracing
Dim varPKValueInGrid As Variant
    ' user - cmdCatalogOps_Click
Dim varPKValueInDB As Variant
    ' user - cmdCatalogOps_Click
Dim varPKValueInTrans As Variant
    ' user - cmdCatalogOps_Click

Dim strRecordsList As String
    ' user - ShowPicklist , dcbFilter_Change, CheckIfFilter
Dim strPluralEntity As String
    ' user - ShowPicklist , InitTextFields, InitTabList, ApplyTextFields
Dim strDummySettings As String
    ' user - InitGrid , SetGridColumns
Dim strLastColumnFilter As String
    ' user - cmdFilter_Click , InitGridFilter, CreateRstToGrid, FormInitialize
Dim strNewColumnFilter As String
    ' user - cmdFilter_Click , FormInitialize
Dim strLastColumnSort As String
    ' user - jgxPicklist_ColumnHeaderClick , ApplyGridFilter, FormInitialize
Dim strHeaderFilter As String
    ' user - chkFilter_Click , optFilter_Click, InitCheckOptions, InitRadioOption,
    '    InitGridRecFilter , InitGridFilter, ApplyGridFilter, FormInitialize, InitControl
Dim strPKFieldAliasInSQL As String
    ' user - ShowPicklist , cmdCatalogOps_Click, TransactAddCopy, TransactEdit,
    '    TransactDelete , RetrieveRecord, CreateRstToGrid

' janus objects
Dim jgxActiveColumn As GridEX16.JSColumn
    ' user - RemoveObjects , RefreshGrid, jgxPicklist_ColumnHeaderClick, SetGridSetting

' form objects
Dim frmOwnerForm As Form
    ' user - ShowPicklist , cmdTransact_Click, RemoveObjects, FetchGridRecords, CommitTransaction,
    '    ApplyGridFilter , InitRecordset, GetMaxAutoInDB, blnCancelDeleteOp, GetTopOne, CheckRecPos,
    '    GetDBID , ResetFormHeight

' ado objects
Dim conDBConnection As ADODB.Connection
    ' user - ShowPicklist , cmdCatalogOps_Click, Form_Unload, RemoveObjects, FetchGridRecords,
    '    CommitTransaction , TransactAddCopy, UpdateDataCombo, InitDataCombo, ProcessRecord,
    '    RetrieveRecord , CreateRstToGrid, InitButtons, blnReconcileAbsPos, InitAddTrans,
    '    GetSelectedRec , blnPKinDB, CheckAutoSearch
Dim rstRecordsList As ADODB.Recordset
    ' user - chkClearFilter_Click , chkFilter_Click, cmdCatalogOps_Click, cmdFilter_Click, Form_Activate, Form_Unload,
    '    RemoveObjects , RefreshGrid, jgxPicklist_Click, jgxPicklist_DblClick, jgxPicklist_KeyDown, jgxPicklist_KeyUp,
    '    jgxPicklist_RowColChange , optFilter_Click, txtSearchFields_Change, SetHeaderTextBoxes, UpdateGrid,
    '    CommitTransaction , TransactAddCopy, InitGrid, RetrieveRecord, CreateRstToGrid,
    '    ApplyGridFilter , InitControls, blnReconcileAbsPos, RunAddTrans, RunCopyTrans, RunEditTrans, RunDeleteTrans,
    '    GetSelectedRec , LoadGrid, UpdateSearchField, ReSynchRecord, SetGridProperty, RepositionRst, SetGridSetting,
    '    SetSelectedRecord
Dim rstFilterRecordsets() As ADODB.Recordset
    ' user - UpdateDataCombo , InitDataCombo
Dim rstFieldModel As ADODB.Recordset
    ' user - cmdCatalogOps_Click , RemoveObjects, FetchGridRecords
Dim rstCurrentDB As ADODB.Recordset
    ' user - cmdCatalogOps_Click , RemoveObjects

' cubepoint objects
Dim WithEvents clsPicklist As CPicklist
Attribute clsPicklist.VB_VarHelpID = -1
    ' user - ShowPicklist , cmdCatalogOps_Click, cmdOK, cmdTransact_Click, Form_Activate, Form_Unload, RemoveObjects,
    '    RefreshGrid , jgxPicklist_DblClick, FetchGridRecords, CommitTransaction, TransactAddCopy, TransactEdit,
    '    TransactDelete , InitTabList, ProcessRecord, RetrieveRecord, InitGridRecFilter, CreateRstToGrid, SetGridColumns,
    '    ApplyTextFields , blnRecordIsDirty, InitControls, InitButtons, blnReconcileAbsPos, RunAddTrans, RunCopyTrans,
    '    RunEditTrans , RunDeleteTrans, CheckChildTrans, ReconcileRecord, GetTransPKValue, InitButton, GetSelectedRec,
    '    blnPKinDB , CheckAutoSearch, IsExactRecord, UpdateBaseSql, LoadGrid, GridSelectedField_Sort, SetGridProperty,
    '    InitFilterInterface , SetSelectedRecord, ReconcileTransPK
Dim clsRecordset As CRecordset
    ' user - Form_Unload , RemoveObjects, UpdateDataCombo, InitDataCombo, FormInitialize
Dim clsGridSeed As CGridSeed
    ' user - ShowPicklist , RemoveObjects, SetHeaderTextBoxes, InitGrid, InitDataComboExt, LoadColumnValue,
    '    ColumnSimplePick , CreateRstToGrid, SetGridColumns, SetGridProperty
Dim clsFilter As CPicklistFilter
    ' user - ShowPicklist , chkFilter_Click, cmdCatalogOps_Click, dcbFilter_Change, Form_Activate, RemoveObjects,
    '    optFilter_Click , UpdateDataCombo, InitCheckBoxCombo, InitCheckOptions, InitComboRecords, InitRadioOption,
    '    CheckPicklistStyle , InitDataCombo, InitDataComboExt, CheckIfFilter, InitGridRecFilter, InitGridFilter, ApplyGridFilter,
    '    InitControl , blnReconcileAbsPos, CheckAutoSearch, GetActiveFilter, UpdateOption
Dim clsRecord As CRecord
    ' user - ShowPicklist , cmdCatalogOps_Click, Form_Unload, RemoveObjects, TransactAddCopy,
    '    TransactEdit , TransactDelete, ProcessRecord, InitAddTrans, RunAddTrans, RunCopyTrans, RunEditTrans,
    '    RunDeleteTrans, GetTransIndex, CheckChildTrans, ReconcileRecord, GetSelectedRec, CheckAutoSearch,
    '    InitFilterInterface, SetSelectedRecord
    
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

'=> IAN 09-15-04
'This is placed to eliminate changing the minimum size of the form when these variables
'are declare globally.

Private g_FormWid
Private g_FormHgt

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single
Private blnDoNotOpen As Boolean 'vince - answer to the problem where frmCatalog is unresponsive
Private strColumnWidth() As String

Private blnHasFirstActivated As Boolean
Private lngXCoordinate As Long
Private lngYCoordinate As Long
    
Private mDontCommitFields As String

'<<< dandan 110607
Private mvarWebLink As String
Private mvarWindowKey As String
Private mvarTemplateDBConnection As ADODB.Connection
Private mvarUserID As Long
Private m_blnRunOnce As Boolean



Public Function ShowPicklist(ByRef OwnerForm As Object, _
                            ByVal Style As PicklistStyle, _
                            ByRef DBConnection As ADODB.Connection, _
                            ByVal RecordsListSQL As String, _
                            ByVal PKFieldAliasInSQL As String, _
                            ByRef PickListToInterface As CPicklist, _
                            ByRef PluralEntity As String, _
                            Optional ByRef GridSeed As CGridSeed = Nothing, _
                            Optional ByRef Filter As CPicklistFilter = Nothing, _
                            Optional ByVal strDontCommitFields As String, _
                            Optional ByVal UserID As Long, _
                            Optional ByRef TemplateDBConnection As ADODB.Connection, _
                            Optional ByVal WindowKey As String, _
                            Optional ByVal WebLink As String) _
                            As Boolean
    
    Set conDBConnection = DBConnection
    Set clsGridSeed = GridSeed
    Set clsFilter = Filter
    Set clsPicklist = PickListToInterface
    
    ' important class, most active
    Set clsRecord = New CRecord 'initilialize the host
    mDontCommitFields = strDontCommitFields
    enuStyle = Style
    strRecordsList = RecordsListSQL
    strPluralEntity = PluralEntity
    strPKFieldAliasInSQL = PKFieldAliasInSQL
    blnCanceled = True
    
    '<<< dandan 110607
    'Added checking for weblink button
    'set cmdseelistonline to visible if weblink is not 0, else disable

    mvarWebLink = IIf(Len(Trim(WebLink)) > 0, Trim(WebLink), vbNullString)

    
    mvarWindowKey = IIf(Len(Trim(WindowKey)) > 0, Trim(WindowKey), vbNullString)
    Set mvarTemplateDBConnection = TemplateDBConnection
    mvarUserID = UserID
    
    ' set to true
    ShowPicklist = True
    ' init temp key
    lngTempAutoNoCtr = 0
    ' interfaces - set reference to calling form
    Set frmOwnerForm = OwnerForm
   
   ' pass the filter
    If ((clsFilter Is Nothing) = False) Then
        Set clsRecord.ActiveFilters = clsFilter.PicklistFilters
        ' initialize option/checkbox interface
        Call InitFilterInterface
    End If
    
    ' check if autosearch
    Call CheckAutoSearch
    
    If (blnHaltExecution = True) Then
        ShowPicklist = False
        Exit Function
    End If
    
    If (clsPicklist.AutoSearch = True) Then
    
        Select Case clsPicklist.ActiveKey
        
            Case cpiKeyEnter, cpiKeyTabEnter
            
                If ((cpiActiveStatus = cpiManyRecord) Or (cpiActiveStatus = cpiNotFound) Or (cpiActiveStatus = cpiOneRecord)) Then
                    
                    ' critical flag
                    blnInitEnd = False
                    
                    ' initialize everything here
                    Call FormInitialize
                    
                    ' update grid setting
                    Call LoadGrid(clsPicklist.ActiveStatus, clsPicklist.SearchField, clsPicklist.SearchValue)
                    Call UpdateSearchField
                    
                    ' error trapper
                    If (blnHaltExecution = True) Then
                        ShowPicklist = False
                        Exit Function
                    End If
                
                ElseIf cpiActiveStatus = cpiOneRecordExact Then
                    Unload Me
                    Exit Function
                End If
            
            Case cpiKeyF2
            
                ' critical flag
                blnInitEnd = False
                
                ' initialize everything here
                Call FormInitialize
                
                ' update grid setting
                Call LoadGrid(clsPicklist.ActiveStatus, clsPicklist.SearchField, clsPicklist.SearchValue)
                Call UpdateSearchField
                
                ' error trapper
                If (blnHaltExecution = True) Then
                    ShowPicklist = False
                    Exit Function
                End If
            
            Case cpiKeyTabNoAction
        
        End Select
        
    ElseIf (clsPicklist.AutoSearch = False) Then
        
        ' critical flag
        blnInitEnd = False
        
        ' initialize everything here
        Call FormInitialize
        
        ' error trapper
        If (blnHaltExecution = True) Then
            ShowPicklist = False
            Exit Function
        End If
    End If
    
    Set Me.Icon = OwnerForm.Icon
    
    


End Function

Private Sub chkClearFilter_Click()

    If (blnFormActivated = True) Then
    
        ' update filter value
        If (chkClearFilter.Value = vbUnchecked) Then
            
            RefreshGrid cpiRefresh, , True
            jgxPicklist.SetFocus
            
            If (rstRecordsList.RecordCount = 0) Then
            
                cmdCatalogOps(CMD_MODIFY).Enabled = False
                cmdCatalogOps(CMD_COPY).Enabled = False
                cmdCatalogOps(CMD_DELETE).Enabled = False
                
            ElseIf (rstRecordsList.RecordCount > 0) Then
            
                cmdCatalogOps(CMD_MODIFY).Enabled = True
                cmdCatalogOps(CMD_COPY).Enabled = True
                cmdCatalogOps(CMD_DELETE).Enabled = True
                
            End If
            
            chkClearFilter.ToolTipText = "Apply Filter"
            
        ElseIf (chkClearFilter.Value = vbChecked) Then
            
            Call cmdFilter_Click
            chkClearFilter.ToolTipText = "Remove Filter"
        
        End If
        
    End If

End Sub

Private Sub chkClearFilter_GotFocus()

    cmdFilter.Tag = "unclick"
    
End Sub

Private Sub chkFilter_Click(Index As Integer)

    Dim lngFilterCtr As Long
    Dim intIndex As Integer
    Dim strFilter As String

    If (blnFormActivated = True) Then

        If (chkFilter(Index).Value = vbChecked) Then
        
            strFilter = "Tag <> 'D' AND " & clsFilter.PicklistFilters(Index + 1).Filter
            blnIsCheck = True
            
        ElseIf (chkFilter(Index).Value = vbUnchecked) Then
        
            strFilter = "Tag <> 'D' AND " & clsFilter.PicklistFilters(Index + 1).Filter
            strFilter = Replace(strFilter, "True", "False", , , vbTextCompare)
            blnIsCheck = False
            
        End If
    
        strHeaderFilter = strFilter
        RefreshGrid cpiRefresh
    
        Exit Sub
    
        If ((clsFilter Is Nothing) = False) Then
            
            Select Case clsFilter.FilterType
                
                Case cpiRadioOptions
                
                Case cpiComboRecords
                
                Case cpiCheckOptions
                    
                    If (chkFilter(Index).Value = vbUnchecked) Then
                        
                        strHeaderFilter = "Tag <> 'D'"
                        rstRecordsList.Filter = strHeaderFilter
                        RefreshGrid cpiRefresh
                        jgxPicklist.Refresh
                    
                    End If
                
                Case Else
            
            End Select
            
        End If
        
        strHeaderFilter = ""
        
        ' hard -coded
        For lngFilterCtr = 0 To clsFilter.FilterCount - 1
            
            If ((chkFilter(lngFilterCtr).Visible = True) And (chkFilter(lngFilterCtr).Value = vbChecked)) Then
                
                strHeaderFilter = strHeaderFilter & clsFilter.PicklistFilters(lngFilterCtr + 1).Filter & " OR "
            
            End If
        
        Next lngFilterCtr
        
        If (Trim$(strHeaderFilter) <> "") Then
            
            strHeaderFilter = Left$(strHeaderFilter, Len(strHeaderFilter) - 4)
            
        ElseIf (Trim$(strHeaderFilter) = "") Then
            ' L1 - start
            strHeaderFilter = Left$(clsFilter.PicklistFilters(1).Filter _
                                        , Len(clsFilter.PicklistFilters(1).Filter) - 3)
            ' L1 - end
        End If
    
    ElseIf (blnFormActivated = False) Then
    
        strFilter = "Tag <> 'D' AND " & clsFilter.PicklistFilters(Index + 1).Filter
        strFilter = Replace(strFilter, "True", "False", , , vbTextCompare)
        RefreshGrid cpiRefresh
        jgxPicklist.Refresh

    End If

End Sub

Private Sub ReconcileGridPos()

    ' reconcile grid positions
    If (rstRecordsList.RecordCount <> 0) Then
        ' set proper recordset position
        Call RepositionRst
        
        If (rstRecordsList.AbsolutePosition <> lngAbsolutePosition) Then
        
            If (lngAbsolutePosition > 0) Then
            
                rstRecordsList.AbsolutePosition = lngAbsolutePosition
                
            End If
            
        End If
        
    End If

End Sub

Private Sub ResetGridFilter(ByRef strTempFilter As String, ByRef varTempBookmark As Variant)

    ' remove filter if filter is catalog
    If ((clsFilter Is Nothing) = False) Then
    
        'If (enuStyle = cpiFilterCatalog) Then
        
            If (cmdFilter = vbUnchecked) Then
            
                If (rstRecordsList.RecordCount <> 0) Then
                
                    ' check position
                    Call RepositionRst
                    strTempFilter = rstRecordsList.Filter
                    varTempBookmark = rstRecordsList.Bookmark
                    rstRecordsList.Filter = ""
                    rstRecordsList.Bookmark = varTempBookmark
                    lngAbsolutePosition = rstRecordsList.AbsolutePosition
                    
                End If
                
            End If
            
        'End If
        
    End If

End Sub

Private Sub GetPrimaryID(ByRef varTempDBID As Variant)
'
    Select Case enuStyle

        Case cpiFilterCatalog
    
            varTempDBID = GetMaxAutoInDB(conDBConnection, GetTableName(clsPicklist.BaseSQL), _
                                        clsPicklist.PKFieldBaseName)
            varTempDBID = varTempDBID + 1
        
            If (clsRecord.RecordSource.RecordCount = 0) Then
            
                clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, varTempDBID
                
            ElseIf (clsRecord.RecordSource.RecordCount = 1) Then
            
                clsRecord.RecordSource.Fields(clsPicklist.PKFieldBaseName).Value = varTempDBID
                
            End If
    
        Case cpiCatalog, cpiSimplePicklist
    
            varTempDBID = GetMaxAutoInDB(conDBConnection, GetTableName(clsPicklist.BaseSQL), _
                                        clsPicklist.PKFieldBaseName)
            
            ' if not empty
            If (varTempDBID = 0) Then
                clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, "1"
                varTempDBID = "1"
            ElseIf varTempDBID <> 0 Then
                varTempDBID = varTempDBID + 1
                clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, varTempDBID
            End If
    
    End Select
    
End Sub

Private Sub GetCurrentRecord(ByRef varTempDBID As Variant, _
                    ByRef strCurrentKey As String, _
                    ByRef strTempSQL As String)

    blnIsInTrans = True
    Set clsRecord.RecordSource = GetTopOne(conDBConnection, strTempSQL, False, _
                                                        clsPicklist.PKFieldBaseName, 0)
    Set clsRecord.RecordSource = RstCopy(clsRecord.RecordSource, True, 0, 0, 1)
    
    ' get topindex in trans
    varTempDBID = GetMaxAutoInTrans(clsPicklist.Transactions, 0, clsPicklist.PKFieldBaseName)
    
    If blnDBAutoNumber = True Then
        clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, varTempDBID
    ElseIf blnDBAutoNumber = False Then
        clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, varTempDBID
    End If
    
    Set clsRecord.OldRecordSource = RstCopy(clsRecord.RecordSource, True, 0, 0, 1)
    
    strCurrentKey = "S" & varTempDBID

End Sub

Private Sub GetNewRecord(ByRef varTempDBID As Variant, _
                    ByRef strCurrentKey As String, _
                    ByRef strTempSQL As String)

    
    ' if no data in DB and no filter then create new
    Set clsRecord.RecordSource = GetTopOne(conDBConnection, strTempSQL, False, _
                                    clsPicklist.PKFieldBaseName, 0)
                                        
    
    'If (clsRecord.RecordSource.RecordCount <> 0) Then
    ' disconnect
        Set clsRecord.RecordSource = RstCopy(clsRecord.RecordSource, True, 0, 0, 1)
    'ElseIf (clsRecord.RecordSource.RecordCount = 0) Then
    '    clsRecord.RecordSource.AddNew
    'End If
    
    If (blnDBAutoNumber = True) Then
        
        ' no data in db
        If ((clsFilter Is Nothing) = True) Then
        
            clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, 1
            varTempDBID = 1
        
        ElseIf ((clsFilter Is Nothing) = False) Then
        
            ' get temp DB ID
            Call GetPrimaryID(varTempDBID)
            
        End If
        
    ElseIf (blnDBAutoNumber = False) Then
    
        clsRecord.RecordSource.AddNew clsPicklist.PKFieldBaseName, "1"
        varTempDBID = "1"
    
    End If
    
    Set clsRecord.OldRecordSource = RstCopy(clsRecord.RecordSource, True, 0, 0, 1)
    
    strCurrentKey = "S" & varTempDBID

End Sub

Private Sub GetRecordPos()

    If (rstRecordsList.RecordCount <> 0) Then
    
     ' reconcile position
     If rstRecordsList.AbsolutePosition = adPosEOF Then
        rstRecordsList.MoveLast
        lngAbsolutePosition = rstRecordsList.AbsolutePosition
     End If
     
     If rstRecordsList.AbsolutePosition = adPosBOF Then
        rstRecordsList.MoveFirst
        lngAbsolutePosition = rstRecordsList.AbsolutePosition
     End If
     
     ' important
     If rstRecordsList.AbsolutePosition <> lngAbsolutePosition Then
        If rstRecordsList.EOF = True Then
           rstRecordsList.MoveLast
        End If
        If lngAbsolutePosition > rstRecordsList.RecordCount Then
           lngAbsolutePosition = rstRecordsList.AbsolutePosition
        ElseIf lngAbsolutePosition <= 0 Then
           lngAbsolutePosition = rstRecordsList.AbsolutePosition
        End If
        rstRecordsList.AbsolutePosition = lngAbsolutePosition
     End If

    End If

End Sub

Private Function ValidateOperation(ByRef varTempDBID As Variant, _
                        ByRef lngTempOldPos As Long, _
                        ByRef lngTempPos As Long, _
                        ByRef strCurrentKey As String, _
                        ByRef blnZeroRec As Boolean, _
                        ByRef enuOperation As RecordOperation, _
                        ByRef strSelectedTag As String, _
                        ByRef strTempFilter As String) As Boolean
'
    ' set the id here
    clsRecord.TempDBID = varTempDBID
    
    ' save recordsources position after passing
    lngTempOldPos = clsRecord.OldRecordSource.AbsolutePosition
    lngTempPos = clsRecord.RecordSource.AbsolutePosition
   
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++ Before Passing the clsRecord    ++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   
   'pass the mother transaction
    Set clsRecord.MotherTransactions = clsPicklist.Transactions
   
    Screen.MousePointer = vbDefault

    If clsPicklist.ButtonClick(clsRecord, enuOperation) = True Then
        ' cancel transaction
        ' the child form has no committed operation.
        ' msgBox "Transaction has been canceled."
        ' do nothing
        ValidateOperation = False
        
        
    Else
    
        ValidateOperation = True
        Screen.MousePointer = vbDefault
        Call CheckChildTrans
        ' return to original position
        
        If clsRecord.OldRecordSource.RecordCount <> 0 Then
            clsRecord.OldRecordSource.AbsolutePosition = _
            IIf(lngTempOldPos > clsRecord.OldRecordSource.RecordCount, _
            clsRecord.OldRecordSource.RecordCount, lngTempOldPos)
        End If
        
        If clsRecord.RecordSource.RecordCount <> 0 Then
            clsRecord.RecordSource.AbsolutePosition = lngTempPos
        End If
        
        ' proceed transaction
        Select Case enuOperation
            Case cpiRecordAdd, cpiRecordCopy
                clsPicklist.TransactionCtr = clsPicklist.TransactionCtr + 1
            Call RunAddTrans(clsRecord, clsPicklist, strCurrentKey, blnZeroRec)
                ' do nothing
            Case cpiRecordEdit
                Call RunEditTrans(clsRecord, strSelectedTag, strCurrentKey)
            Case cpiRecordDelete
                Call RunDeleteTrans(rstRecordsList, clsRecord, clsPicklist, strCurrentKey)
                lngTempPos = lngAbsolutePosition
                
                If (rstRecordsList.RecordCount <> 0) Then
                    rstRecordsList.AbsolutePosition = lngAbsolutePosition
                End If
                
                RepositionRst
                'GoTo RefreshJGXGrid
                Call RunRefreshGrid(strTempFilter, enuOperation)
                lngAbsolutePosition = lngTempPos
            Case Else
        End Select
    End If

End Function

Private Sub RunRefreshGrid(ByRef strTempFilter As String, _
                                                        ByRef enuOperation As RecordOperation)
'
    Dim strSearchField As String, strSearchValue As String
    
    'If (clsPicklist.AutoSearch = True) Then
        If rstRecordsList.RecordCount <> 0 Then
            strSearchField = rstRecordsList.Fields(clsPicklist.SearchField).Name
            strSearchValue = rstRecordsList.Fields(clsPicklist.SearchField).Value
        End If
    'End If
   
    ' re-apply original filter if filter catalog
    If Not (clsFilter Is Nothing) Then
        'If enuStyle = cpiFilterCatalog Then
            If cmdFilter = vbUnchecked Then
                If rstRecordsList.RecordCount <> 0 Then
                    rstRecordsList.Filter = strTempFilter
                End If
            End If
        'End If
    End If

    If (enuOperation <> cpiRecordDelete) Then
        lngAbsolutePosition = rstRecordsList.AbsolutePosition
    ElseIf (enuOperation = cpiRecordDelete) Then
        If (rstRecordsList.RecordCount <> 0) Then
            If (jgxPicklist.Row <= rstRecordsList.RecordCount) Then
                rstRecordsList.AbsolutePosition = jgxPicklist.Row
                If clsPicklist.AutoSearch = True Then
                    strSearchField = rstRecordsList.Fields(clsPicklist.SearchField).Name
                    strSearchValue = rstRecordsList.Fields(clsPicklist.SearchField).Value
                End If
            ElseIf jgxPicklist.Row > rstRecordsList.RecordCount Then
                If rstRecordsList.RecordCount > 0 Then
                    rstRecordsList.AbsolutePosition = rstRecordsList.RecordCount
                    If clsPicklist.AutoSearch = True Then
                        strSearchField = rstRecordsList.Fields(clsPicklist.SearchField).Name
                        strSearchValue = rstRecordsList.Fields(clsPicklist.SearchField).Value
                    End If
                End If
            End If
        End If
    End If
   
    ' update grid
    RefreshGrid cpiRefresh
      
    ' update selected record
    If chkClearFilter.Value = vbUnchecked Then
        If jgxPicklist.ADORecordset.RecordCount > 0 Then
            Call LoadGrid(cpiOneRecordExact, strSearchField, strSearchValue)
            Call UpdateSearchField
        End If
    End If

End Sub

Private Function CheckGridRecord(ByRef blnZeroRec As Boolean, _
                    ByRef strCurrentKey As String, _
                    ByRef strTempSQL As String, _
                    ByRef enuOperation As RecordOperation, _
                    ByRef lngTempOldPos As Long, _
                    ByRef lngTempPos As Long, _
                    ByRef strSelectedTag As String, _
                    ByRef strTempFilter As String) As Boolean
'
    ' check if db is empty
    If rstRecordsList.RecordCount = 0 Then
        
        If clsPicklist.Transactions.Count = 0 Then
            
            blnZeroRec = True
            
            Call GetNewRecord(varTempDBID, strCurrentKey, strTempSQL)
            
            If enuOperation = cpiRecordAdd Then
                'GoTo AddRec
                'CheckGridRecord = ValidateOperation(varTempDBID, lngTempOldPos, lngTempPos, strCurrentKey, blnZeroRec, _
                                                            enuOperation, strSelectedTag, strTempFilter)
            'ElseIf enuOperation = cpiRecordEdit Then
            '    Exit Sub
            End If
        
        ' no rec with trans
        ElseIf clsPicklist.Transactions.Count <> 0 Then
            
            Call GetCurrentRecord(varTempDBID, strCurrentKey, strTempSQL)
            
            If enuOperation = cpiRecordAdd Then
                ' goto AddRec
                'CheckGridRecord = ValidateOperation(varTempDBID, lngTempOldPos, lngTempPos, strCurrentKey, blnZeroRec, _
                                                        enuOperation, strSelectedTag, strTempFilter)
            
            'ElseIf enuOperation = cpiRecordEdit Then
            '    Exit Sub
            End If
            
        End If
        
    End If

End Function

Private Function UpdateOperation(ByRef blnZeroRec As Boolean, _
                                                ByRef strTempSQL As String, _
                                                ByRef enuOperation, _
                                                ByRef lngIndex As Long) As Boolean

    Dim rstDummy As ADODB.Recordset

    UpdateOperation = True

    If ((blnZeroRec = False) And (rstRecordsList.RecordCount <> 0)) Then
        ' get the first record that has proper criteria
        
        ' O nasa database
        If (rstRecordsList!Tag = "O") Then
            
            ' find record from database if tag = 'O'
            Set rstDummy = New ADODB.Recordset
            Set rstDummy = GetTopOne(conDBConnection, strTempSQL, False, _
            clsPicklist.PKFieldBaseName, varTempDBID)
            
            If (rstDummy.State = adStateOpen) Then
            
                If (rstDummy.RecordCount = 0) Then
                    ' check in transaction collection
                    If clsPicklist.Transactions.Count <> 0 Then
                        blnIsInTrans = True
                    End If
                ElseIf rstDummy.RecordCount <> 0 Then
                    ' get from db if found
                    Set clsRecord.RecordSource = RstCopy(rstDummy, True, 0, 0, 1)
                End If
        
            ElseIf (rstDummy.State = adStateClosed) Then
            
                MsgBox "Error in SQL syntax or fieldname not found: " & strTempSQL, vbInformation, "SQL Error"
                Err.Clear
                Set clsRecord.RecordSource = rstRecordsList
                UpdateOperation = False
            
            End If
        
        ElseIf rstRecordsList!Tag = "A" Then
        
            If (enuOperation = cpiRecordDelete) Then
            
            ' temporary recordset for deleting
            Set clsRecord.RecordSource = RstCopy(rstRecordsList, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1)
            
            ElseIf enuOperation <> cpiRecordDelete Then
            
                If rstFieldModel.BOF = True Then
                    rstFieldModel.MoveFirst
                End If
                
                If rstFieldModel.EOF = True Then
                    rstFieldModel.MoveLast
                End If
            
                Set clsRecord.RecordSource = RstCopy(rstFieldModel, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1, True)
                clsRecord.RecordSource.Fields(clsPicklist.PKFieldBaseName).Value = _
                                                                rstRecordsList.Fields(clsPicklist.PKFieldBaseName).Value
                clsRecord.RecordSource.Fields(clsPicklist.PKFieldAlias).Value = _
                                                                rstRecordsList.Fields(clsPicklist.PKFieldAlias).Value
            End If
        
        ElseIf rstRecordsList!Tag = "M" Then
        
            Set clsRecord.RecordSource = RstCopy(rstRecordsList, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1)
            lngIndex = GetTransIndex(clsPicklist.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
            
            If lngIndex = 0 Then
                Set clsRecord.RecordSource = RstCopy(rstRecordsList, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1)
            Else
                ' reconcile transaction and grid record
                Set clsRecord.RecordSource = clsPicklist.Transactions(lngIndex).RecordSource
            End If
            
        End If
        
        Set rstDummy = Nothing
        
    ElseIf (blnZeroRec = True) Then
        
        ' initial make a copy the current selected record to Clsrecord.Recordsource - wrong in absolute position
        'Set clsRecord.RecordSource = RstCopy(rstRecordsList, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1)
    
    ElseIf ((blnZeroRec = False) And (rstRecordsList.RecordCount = 0)) Then
    End If

    Set rstDummy = Nothing


End Function

Private Sub UpdateGridRecord(ByRef enuOperation As RecordOperation, _
                                                    ByRef strCurrentKey As String, _
                                                    ByRef blnZeroRec As Boolean, _
                                                    ByRef strSelectedTag As String, _
                                                    ByRef lngIndex As Long)
'
    Select Case enuOperation
    
        Case cpiRecordAdd
            ' +1 to PK and Remove existing Values; set tag to 'A'
            ' initialize values when before passing
            If (clsRecord.RecordSource.RecordCount <> 0) Then
                Call InitAddValues(clsRecord.RecordSource)
                ' update key - initialize before add/copy transaction
                strCurrentKey = InitAddTrans(clsRecord, clsPicklist, strPKFieldAliasInSQL)
            End If
            
        Case cpiRecordCopy ' +1 to PK and retain existing values set tag to 'A'
            ' reconcile the values for transaction and DB
            ' set clsrecord to this temporary field record
            'Set clsRecord.RecordSource = RstCopy(rstRecordsList, True, lngAbsolutePosition - 1, lngAbsolutePosition - 1, 1)
            Call ReconcileRecord
            ' get new key
            strCurrentKey = InitAddTrans(clsRecord, clsPicklist, strPKFieldAliasInSQL)
            
        Case cpiRecordEdit ' retain PK, retain existing values set tag to 'A'
        
            If (blnZeroRec = True) Then
                Exit Sub
            End If
            
            strSelectedTag = rstRecordsList!Tag
            
            If strSelectedTag = "O" Then
                ' do nothing
            ElseIf strSelectedTag = "A" Then
                ' set clsrecord to this temporary field record
                lngIndex = GetTransIndex(clsPicklist.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
                Set clsRecord.RecordSource = clsPicklist.Transactions(lngIndex).RecordSource
            ElseIf rstRecordsList!Tag = "M" Then
                ' change record source to transaction collection
                lngIndex = GetTransIndex(clsPicklist.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
                Set clsRecord.RecordSource = clsPicklist.Transactions(lngIndex).RecordSource
            End If
        
        Case cpiRecordDelete
            ' do nothing
    End Select

End Sub

Private Sub RunOperation(ByVal Index As Integer)

    Dim strFieldValue As String
    Dim strSelectedTag As String
    Dim strCurrentKey As String
    Dim strTempSQL As String
    Dim strTempFilter As String
    
    Dim lngDataType As Long
    Dim lngCurrentValue As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngTempOldPos As Long
    Dim lngTempPos As Long
    Dim lngIndex As Long
    Dim lngAbsPosRec As Long
    
    Dim blnZeroRec As Boolean
    Dim blnRecordExists As Boolean
    Dim blnValidateOperation As Boolean
    
    Dim varTempBookmark As Variant
   
    Dim enuOperation As RecordOperation
    
    ' if button is not visible then exit sub
    If (cmdCatalogOps(Index).Visible = False) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' save last absolute position value
    lngAbsPosRec = lngAbsolutePosition
    
    ' checkbox flag patches
    blnTaskClick = True
    
    ' check proper operation
    enuOperation = Choose(Index + 1, cpiRecordAdd, cpiRecordEdit, cpiRecordCopy, cpiRecordDelete)
    
    strTempSQL = RegenerateSQL(clsPicklist.BaseSQL, False)
    
    ' if grid is zero then check transaction
    blnZeroRec = False
    blnIsInTrans = False
    
    ' check if grid is zero
    
    ' set clsrecord.recordsource to nothing first
    Set clsRecord.RecordSource = Nothing
    
    ' set grid pos
    Call ReconcileGridPos
    
    ' reset filters
    Call ResetGridFilter(strTempFilter, varTempBookmark)
   
   ' check if delete operation is cancelled
    If (blnCancelDeleteOp(enuOperation) = True) Then
        ' cancel delete operation
        Exit Sub
    End If
   
    ' check if db is empty
    If (enuOperation <> cpiRecordEdit) Then
    
        blnValidateOperation = CheckGridRecord(blnZeroRec, strCurrentKey, strTempSQL, enuOperation, lngTempOldPos, _
                                                            lngTempPos, strSelectedTag, strTempFilter)
                                            
    'ElseIf (enuOperation = cpiRecordEdit) Then
    
        'Exit Sub
        
    End If
   
   Screen.MousePointer = vbHourglass
   
    ' set proper position
   Call GetRecordPos
   
   If (rstRecordsList.RecordCount <> 0) Then
        ' get primary key values from grid
        varPKValueInGrid = rstRecordsList.Fields(clsPicklist.PKFieldBaseName).Value
    End If
    
    ' to be added later
   Set rstCurrentDB = rstDBCopy(conDBConnection, strTempSQL _
                                , adOpenKeyset, adLockOptimistic, True)
   
    ' no rec in db , trans not zero
    If (rstCurrentDB.RecordCount <> 0) Then
        varPKValueInDB = rstCurrentDB.Fields(clsPicklist.PKFieldBaseName).Value
    End If
    
    varPKValueInTrans = GetTransPKValue
    
    ' there will be three important recordsets for tracing this procedure
    ' first: the current grid records - rstRecordslist
    ' second: the current record in the DB - rstCurrentDB 'connected
    ' third: the records in the transaction collection - clsPicklist.transactions(1 to n).recordsource looper function
    
    If ((rstCurrentDB.RecordCount <> 0) And (rstRecordsList.RecordCount <> 0)) Then
        ' get current key
        strCurrentKey = "S" & CStr(rstRecordsList.Fields(strPKFieldAliasInSQL).Value)
        varTempDBID = rstRecordsList.Fields(strPKFieldAliasInSQL).Value
    End If
    
    ' to be used also by user
    Call SetOldRecordPos(rstRecordsList, clsPicklist.LoadAllRecord, lngStart, lngEnd)
    
    ' load records for copying
    If clsPicklist.LoadAllRecord = False Then
        lngAbsolutePosition = 1
    End If
    
    If (rstRecordsList.RecordCount <> 0) Then
        Set clsRecord.OldRecordSource = RstCopy(rstRecordsList, True _
                                                        , lngStart, lngEnd, lngAbsolutePosition)
                                                        
        lngAbsolutePosition = rstRecordsList.AbsolutePosition
   
   End If
   
    ' update operation
    If UpdateOperation(blnZeroRec, strTempSQL, enuOperation, lngIndex) = True Then
   
        ' update recordlist
          Call UpdateGridRecord(enuOperation, strCurrentKey, blnZeroRec, strSelectedTag, lngIndex)
        
         Call ValidateOperation(varTempDBID, lngTempOldPos, lngTempPos, strCurrentKey, blnZeroRec, _
                      enuOperation, strSelectedTag, strTempFilter)
        
        Call RunRefreshGrid(strTempFilter, enuOperation)
        
        jgxPicklist.SetFocus

    End If

    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCatalogOps_Click(Index As Integer)
    If Not ByPassButton(Index) Then
        Call RunOperation(Index)
    End If
   
    gHW = Me.hwnd

    If IsCompiled Then
        Unhook
        Hook g_FormWid, g_FormHgt
    End If
End Sub

Private Sub cmdFilter_Click()

    Dim intSearchBoxCtr As Integer
    Dim varValue As Variant
    
    cmdFilter.Tag = "clicked"
    chkClearFilter.Value = vbChecked
    
    strNewColumnFilter = ""
    
    For intSearchBoxCtr = 0 To txtSearchFields.Count - 1
        
        ' check if it is the selecte textboxd
        If (txtSearchFields(intSearchBoxCtr).BackColor = &HC0FFFF) Then
        
            If (Trim$(txtSearchFields(intSearchBoxCtr).Text) <> "") Then
            
                varValue = Trim$(txtSearchFields(intSearchBoxCtr).Text)
                
                strNewColumnFilter = GetWhereClause(txtSearchFields(intSearchBoxCtr).Tag, _
                                    txtSearchFields(intSearchBoxCtr).Text, _
                                    rstRecordsList.Fields(txtSearchFields(intSearchBoxCtr).Tag).Type)
            End If
            
            Exit For
        End If
        
    Next intSearchBoxCtr
    
    strLastColumnFilter = strNewColumnFilter
    
    RefreshGrid cpiRefresh
    
    jgxPicklist.SetFocus
    
    If (rstRecordsList.RecordCount = 0) Then
    
        cmdCatalogOps(CMD_MODIFY).Enabled = False
        cmdCatalogOps(CMD_COPY).Enabled = False
        cmdCatalogOps(CMD_DELETE).Enabled = False
        
    ElseIf (rstRecordsList.RecordCount <> 0) Then
    
        cmdCatalogOps(CMD_MODIFY).Enabled = True
        cmdCatalogOps(CMD_COPY).Enabled = True
        cmdCatalogOps(CMD_DELETE).Enabled = True
        
    End If
    
    chkClearFilter.Enabled = True

End Sub

Private Sub cmdOK_Click()

    blnCanceled = False
    
    ' return check for cancel trans to be use by the programmer
    clsPicklist.CancelTrans = blnCanceled
    
    Call SetSelectedRecord

End Sub


Private Sub cmdSeeOnLine_Click()
    Dim m_clsWebLink As CIEExplore
    '<<< dandan 110607
    'Launch internet browser using mvarweblink
    Set m_clsWebLink = New CIEExplore
    
    m_clsWebLink.OpenURL Trim(mvarWebLink), Me
End Sub

Private Sub cmdTransact_Click(Index As Integer)

    Dim varAns As Variant
    
    
    If (Index = CMD_OK) Then
    
        Screen.MousePointer = vbHourglass
        cmdOK.Value = True
        
        ' return check for cancel trans to be use by the programmer
        If (clsPicklist.CancelTrans = False) Then 'if cancele exit
        
            'refresh owner form
            Me.Hide
            ' :o) frmOwnerForm.Refresh
            Unload Me
        
        End If
        
        Screen.MousePointer = vbDefault
        

    
    ElseIf (Index = CMD_CANCEL) Then
        
        blnCloseBtnClick = True
        
        If Not clsPicklist Is Nothing Then
            If ((clsPicklist.Transactions Is Nothing) = False) Then
            
                If (clsPicklist.Transactions.Count > 0) Then
                
                    ' check if AUTOCANCEL or NOT
                    If (clsPicklist.AutoUnload = CPI_FALSE) Then
                    
                        ' roylann 09-02-2002
                        varAns = MsgBox("Cancel will cause all previously made changes to be ignored. " _
                        & " Are you sure you want to ignore these changes?", vbYesNo + vbInformation, _
                        Me.Caption & " - Cancel Transaction")
                          
                        If varAns = vbYes Then
                        
                            Screen.MousePointer = vbHourglass
                            blnCanceled = True
                            
                            ' return check for cancel trans to be use by the programmer
                            clsPicklist.CancelTrans = blnCanceled
                            ' refresh owner form
                            Me.Hide
                            frmOwnerForm.MousePointer = vbDefault
                            ' :o) frmOwnerForm.Refresh
                            Unload Me
                            
                        End If
            
                    ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
            
                        Screen.MousePointer = vbHourglass
                
                        ' return check for cancel trans to be use by the programmer
                        If (clsPicklist.CancelTrans = False) Then
                        
                            ' refresh owner form
                            Me.Hide
                            ' :o) frmOwnerForm.Refresh
                            frmOwnerForm.MousePointer = vbDefault
                            blnCanceled = False
                            Unload Me
                        
                        End If
            
                    End If
            
                ElseIf (clsPicklist.Transactions.Count = 0) Then
            
                    Screen.MousePointer = vbHourglass
                    blnCanceled = True
            
                    ' return check for cancel trans to be use by the programmer
                    clsPicklist.CancelTrans = blnCanceled
                    ' refresh owner form
                    Screen.MousePointer = vbDefault
                    frmOwnerForm.MousePointer = vbDefault
                    ' :o) frmOwnerForm.Refresh
                    Me.Hide
                    Unload Me
            
                End If
                
            ElseIf ((clsPicklist.Transactions Is Nothing) = True) Then
            
                Screen.MousePointer = vbHourglass
                blnCanceled = True
                
                ' return check for cancel trans to be use by the programmer
                clsPicklist.CancelTrans = blnCanceled
                ' refresh owner form
                Screen.MousePointer = vbDefault
                frmOwnerForm.MousePointer = vbDefault
                ' :o) frmOwnerForm.Refresh
                Me.Hide
                Unload Me
                
            End If
        End If
    End If
    
End Sub

'Private Sub dcbFilter_Change(Index As Integer)
'
'    Dim strSQL As String
'
'    If Val(dcbFilter(Index).BoundText) = 0 Then
'        Exit Sub
'    End If
'
'    strSQL = ""
'
'    If (Index < (clsFilter.PicklistFilters.Count - 1)) = True Then
'        Call UpdateDataCombo(strSQL, Index)
'    ElseIf (Index < (clsFilter.PicklistFilters.Count - 1)) = False Then
'        ' Code to Repopulate Grid
'        RefreshGrid cpiRequery, strRecordsList
'    End If
'
'End Sub

' Fan-Out: CheckAutoSearch
Private Function InsertWhere(ByVal strBaseSql As String, ByVal strWhere As String) As String

    Dim strReturn As String
    Dim strSplit() As String
    Dim blnExist As Boolean
    Dim blnWhere As Boolean
    
    If Trim(strWhere) <> "" Then
        blnExist = (InStr(1, UCase$(strBaseSql), " ORDER BY ", vbTextCompare) <> 0)
        blnWhere = (InStr(1, UCase$(strBaseSql), " WHERE ", vbTextCompare) <> 0)
        
        ' check if ORDER BY clause exist
        If (blnExist = True) Then
            strSplit = Split(strBaseSql, " ORDER BY ", , vbTextCompare)
            
            If (blnWhere = False) Then
                strReturn = strSplit(0) & " WHERE " & strWhere & " ORDER BY" & strSplit(1)
            Else
                strReturn = strSplit(0) & " AND " & strWhere & " ORDER BY" & strSplit(1)
            End If
        ElseIf (blnExist = False) Then
            If (blnWhere = False) Then
                strReturn = strBaseSql & " WHERE " & strWhere
            Else
                strReturn = strBaseSql & " AND " & strWhere
            End If
        End If
    Else
        strReturn = strBaseSql
    End If
    
    InsertWhere = strReturn
   
End Function

Private Sub Form_Activate()
    
    Dim intFieldIndex As Integer
    Dim intDelay As Integer
    Dim strSendKeys As String
    Dim lngRowPos As Long
   
    Dim jgxColumn As JSColumn
    
    blnFormActivated = True
    
    If (jgxPicklist.RowCount > 0) Then
        Call SetHeaderTextBoxes
    End If
    
    If (Trim$(CStr(Me.Tag)) = "") Then
    
        Me.Tag = "1"
        
        If (txtSearchFields(2).Visible = True) Then
            
            txtSearchFields(2).SetFocus
        
        ElseIf (txtSearchFields(2).Visible = False) Then
            
            If (txtSearchFields(0).Visible = True) Then
                txtSearchFields(0).SetFocus
            End If
        
        End If
    
    End If
    
    If (rstRecordsList.RecordCount = 0) Then
        cmdCatalogOps(CMD_MODIFY).Enabled = False
        cmdCatalogOps(CMD_COPY).Enabled = False
        cmdCatalogOps(CMD_DELETE).Enabled = False
    Else
        jgxPicklist.MoveLast
        jgxPicklist.MoveFirst
        jgxPicklist.Refresh
        
    End If
   
    ' init selected grid record
   blnSkipTextChange = True
   Call SetGridProperty
   blnSkipTextChange = False
   
    If (clsPicklist.AutoSearch = True) Then
    
        jgxPicklist.SetFocus
        
        If (rstRecordsList.RecordCount <> 0) Then
        
            ' solution #1 - start
            jgxPicklist.Row = lngAbsolutePosition
            rstRecordsList.AbsolutePosition = lngAbsolutePosition
            Call jgxPicklist_Click
            ' then ascending
            GridSelectedField_Sort
            ' solution #1 - end
            
        End If
    
    End If
      
    If (clsFilter Is Nothing) = False Then
    
        If ((clsFilter.FilterType = cpiCheckOptions) Or (clsFilter.FilterType = cpiRadioOptions)) Then
        
            Call UpdateOption
        
        End If
    
    End If
    
    chkClearFilter.Enabled = False
    blnDoNotOpen = False
    
    'Uncommented by BCo for minimum form size restriction
    'Save handle to the form.
    gHW = Me.hwnd
    'Begin subclassing.
    
    'The condition statement was inserted to eliminate going to the procedure when
    'the form was already loaded and activated before. =>IAN 09-15-04
    
    If Not blnHasFirstActivated Then
        SaveSizes
        
        ' Set the minimum width and height to be the
        ' same as the initial height and width
        If IsCompiled Then
            Hook g_FormWid, g_FormHgt
        End If
    End If
    
    'vince - get the Column width and the Column name and put it in the dynarray
    If Not blnHasFirstActivated Then
        ReDim strColumnWidth(0)
        For Each jgxColumn In jgxPicklist.Columns
            If jgxColumn.Width > 0 And jgxColumn.Visible = True And Trim(jgxColumn.Caption) <> "" Then
                strColumnWidth(UBound(strColumnWidth)) = Trim(jgxColumn.Caption) & "$$" & jgxColumn.Width
                ReDim Preserve strColumnWidth(UBound(strColumnWidth) + 1)
            End If
        Next
        ReDim Preserve strColumnWidth(UBound(strColumnWidth) - 1)
        blnHasFirstActivated = True
        'Me.BorderStyle = 1  'sizable
    End If
    
        '<<< dandan 110607
    ' Retrieve window settings from database
    Dim clsWinSettings As PCubeLibWinSet.IWindows
        
    If Not m_blnRunOnce Then
        If (Len(Trim(mvarWindowKey)) > 0) Then
            Set clsWinSettings = New PCubeLibWinSet.IWindows
            clsWinSettings.LoadWindowSettings mvarTemplateDBConnection, mvarUserID, mvarWindowKey, Me
            Set clsWinSettings = Nothing
            m_blnRunOnce = True
        End If
    End If
    
    'cmdSeeOnLine.ZOrder 0
    If Len(Trim(mvarWebLink)) > 0 Then
        cmdSeeOnLine.Visible = True
        cmdSeeOnLine.ToolTipText = mvarWebLink
    Else
        cmdSeeOnLine.Visible = False
    End If
    'cmdSeeOnLine.Visible = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyReturn) Then
        Call jgxPicklist_DblClick
    End If
    
    If jgxPicklist.Visible Then
        If Not jgxPicklist.ADORecordset Is Nothing Then
            If jgxPicklist.ADORecordset.RecordCount > 0 Then
                On Error Resume Next    ' Handle EOF or BOF
                Select Case KeyCode
                    Case vbKeyDown
                        jgxPicklist.SetFocus
                        jgxPicklist.MoveNext
                    Case vbKeyUp
                        jgxPicklist.SetFocus
                        jgxPicklist.MovePrevious
                End Select
                On Error GoTo 0
            End If
        End If
    End If
End Sub


Private Sub Form_Load()

    
    'Me.BorderStyle = 1  'sizable
    Me.jgxPicklist.Refresh
End Sub

Private Sub Form_Resize()
    ResizeControls
    'Me.jgxPicklist.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim clsReturnRecord As CRecord
    Dim rstReturn As ADODB.Recordset
    
    
    Screen.MousePointer = vbHourglass
    
    '<<< dandan 110607
    ' Save window settings from database
    Dim clsWinSettings As PCubeLibWinSet.IWindows
        
    If (Len(Trim(mvarWindowKey)) > 0) Then
        Set clsWinSettings = New PCubeLibWinSet.IWindows
        clsWinSettings.SaveWindowSettings mvarTemplateDBConnection, mvarUserID, mvarWindowKey, Me
        Set clsWinSettings = Nothing
    End If
    
    'Hide
    If Not clsPicklist Is Nothing Then
        clsPicklist.OwnerForm.Refresh
        
        If ((blnCloseBtnClick = False) And (clsPicklist.AutoUnload = CPI_TRUE)) Then
            blnCanceled = False
        End If
        
        If (clsPicklist.AutoSearch = True) Then
            If (cpiActiveStatus = cpiOneRecordExact) Then
                If (clsPicklist.Transactions.Count > 0) Then
                    'If ((clsReturnRecord Is Nothing) = False) Then
                        Call ProcessRecord(clsReturnRecord, Cancel)
                    'End If
                End If
                GoTo DestroyObject
            End If
        End If
    End If
    
    ' process transaction records if any
    If ((clsRecord Is Nothing) = True) Then
    
        ' no child transactions
        conDBConnection.RollbackTrans
    
    ElseIf ((clsRecord Is Nothing) = False) Then
    
        Call ProcessRecord(clsReturnRecord, Cancel)
        
        If (blnCanceled = False) Then
        
            ' return record if simple picklist
            Set clsReturnRecord = New CRecord
            Set rstReturn = New ADODB.Recordset
        
            If (blnIsInTrans = False) Then
                
                ' reconcile positions
                If (rstRecordsList.AbsolutePosition <> lngAbsolutePosition) Then
                    If (rstRecordsList.RecordCount <> 0) Then
                        If (lngAbsolutePosition > 0) Then
                            rstRecordsList.AbsolutePosition = lngAbsolutePosition
                        End If
                    End If
                End If
        
                ' check if in trans
                Set rstReturn = GetSelectedRec
        
            End If
        
            Set clsReturnRecord.RecordSource = rstReturn
        
            If ((clsPicklist.SelectedRecord Is Nothing) = True) Then
                ' no selected, no record
                If (rstRecordsList.RecordCount = 0) Then
                    Set clsPicklist.SelectedRecord = Nothing
                Else
            
                    If (clsPicklist.AutoUnload = CPI_AUTOCANCEL) Then
                        Set clsPicklist.SelectedRecord = clsReturnRecord
                    ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
                        Set clsPicklist.SelectedRecord = Nothing
                        clsPicklist.CancelTrans = True
                    End If
            
                End If
            End If
            
            Set clsPicklist.Transactions = clsRecord.ChildTransactions
        
        End If
    End If
    
    ' return the grid recordset to be used by the programmer
    rstRecordsList.Filter = ""
    If ((clsPicklist.LoadAllRecord = True) And (rstRecordsList.RecordCount <> 0) And (blnCanceled = False)) Then
        rstRecordsList.MoveFirst
        Set clsPicklist.GridRecord = RstCopy(rstRecordsList, True, 0, rstRecordsList.RecordCount - 1, 1)
    End If
    
    ADORecordsetClose rstRecordsList
    'clsRecordset.cpiClose rstRecordsList

DestroyObject:
    
    Set clsReturnRecord = Nothing
    Set rstReturn = Nothing
        
    blnOperator = True

    Screen.MousePointer = vbDefault
    
    Call RemoveObjects
        
    If IsCompiled Then
        'Stop subclassing.
        Unhook
    End If
End Sub

' Fan-Out: Form_Unload
Private Sub RemoveObjects()

    Dim intCtr As Integer
    
    Set jgxActiveColumn = Nothing
    Set frmOwnerForm = Nothing
    
    ADODisconnectDB conDBConnection
    ADORecordsetClose rstRecordsList
    ADORecordsetClose rstFieldModel
    ADORecordsetClose rstCurrentDB
    
    ' destroy cubepoint objects
    Set clsRecordset = Nothing
    Set clsGridSeed = Nothing
    Set clsFilter = Nothing
    Set clsRecord = Nothing
    
    ' return the pick
    clsPicklist.PickEnd = True
    Set clsPicklist = Nothing
        
End Sub

' Fan-Out: CheckIfFilter,chkClearFilter_Click,chkFilter_Click,cmdCatalogOps_Click,cmdFilter_Click,
                        '  dcbFilter_Change,optFilter_Click
Private Sub RefreshGrid(ByVal Repaint As RefreshType, _
                    Optional ByVal Source As String = "", _
                    Optional ByVal DisableFilter As Boolean = False, _
                    Optional ByVal AdditionalFilter As String = "")
                                             
    Dim strFilterToApply As String
    Dim strCurrentFilter As String
    Dim strTempFormat() As String
    
    Dim intColumnCtr As Integer
    Dim intSearchTextIndex As Integer
    Dim intIndex As Integer
    Dim intLoopIndex As Integer
    
    Dim lngVisibleColumnCtr As Long
    
    Dim clsWhere As CFilterNodes
    
    Static blnIsFilter As Boolean
    Static lngTempAbsPos As Long
    
    ReDim strTempFormat(jgxPicklist.Columns.Count)
    
    lngVisibleColumnCtr = 0
    intSearchTextIndex = 0
    strFilterToApply = ""
    
    ' initialize grid recordset filter value
    Call InitGridRecFilter(Repaint, intSearchTextIndex, lngVisibleColumnCtr, intColumnCtr, _
            strFilterToApply, strCurrentFilter, Source, DisableFilter, clsWhere)
    
    If Trim$(AdditionalFilter) <> "" Then
        If Trim$(strCurrentFilter) = "" Then
            strCurrentFilter = AdditionalFilter
        End If
    End If
    ' populate the grid
    Call PopulateGrid(Repaint, intSearchTextIndex, lngVisibleColumnCtr, intColumnCtr, _
            strFilterToApply, strCurrentFilter, Source, DisableFilter)
    
    If (blnHaltExecution = True) Then
        Exit Sub
    End If
   
    ' apply filter to grid recordset
   Call ApplyGridFilter(Repaint, intSearchTextIndex, lngVisibleColumnCtr, intColumnCtr, _
        strFilterToApply, strCurrentFilter, Source, DisableFilter, clsWhere)

    Set clsWhere = Nothing

    ' save the last position before filtering
    If (chkClearFilter.Value = vbUnchecked) Then
    
        If (blnIsFilter = False) Then
            lngTempAbsPos = lngAbsolutePosition
        End If
        
        blnIsFilter = False
        
    ElseIf (chkClearFilter.Value = vbChecked) Then
        blnIsFilter = True
    End If
    
    ' refresh start
    jgxPicklist.Visible = False
    
    Set clsPicklist.GridRecord = rstRecordsList
    Set jgxPicklist.ADORecordset = rstRecordsList
    jgxPicklist.Rebind
    Set jgxPicklist.ADORecordset = rstRecordsList
    jgxPicklist.Refresh
    
    ' set setting for grid columns
    Call SetGridColumns(Repaint, intSearchTextIndex, lngVisibleColumnCtr, intColumnCtr, _
                                        strFilterToApply, strCurrentFilter, Source, DisableFilter)
    
    ' sets the alignment formats etc of the grid
    blnSkipTextChange = True
    Call SetGridProperty
    blnSkipTextChange = False
    
    ' check if grid has no records displayed
    If (jgxPicklist.ItemCount > 0) Then
        blnGridIsEmpty = False
    ElseIf (jgxPicklist.ItemCount = 0) Then
        blnGridIsEmpty = True
    End If
    
    ' set grid sortorders
    If ((jgxActiveColumn Is Nothing) = False) Then
    
        If (rstRecordsList.RecordCount <> 0) Then
            
            ' start restoring
            Call jgxPicklist_ColumnHeaderClick(jgxActiveColumn)
            Call jgxPicklist_ColumnHeaderClick(jgxActiveColumn)
        
        End If
    
    End If
    
    If (jgxPicklist.ADORecordset.RecordCount <> 0) Then
        
        If (lngTempAbsPos > 0) Then
            
            If (chkClearFilter.Value = vbUnchecked) Then
            
                rstRecordsList.AbsolutePosition = lngTempAbsPos
                Call RepositionRst
                
                If (rstRecordsList.RecordCount = 1) Then
                    lngTempAbsPos = 1
                End If
                
                If (lngTempAbsPos > rstRecordsList.RecordCount) Then
                    lngTempAbsPos = rstRecordsList.RecordCount
                End If
                
                jgxPicklist.MoveToBookmark (rstRecordsList.Bookmark)
                
                Do While (rstRecordsList.Bookmark < lngTempAbsPos)
                
                    jgxPicklist.Row = lngTempAbsPos
                    
                    ' safety factor in case there is an infinite loop
                    intLoopIndex = intLoopIndex + 1
                    
                    If (intLoopIndex = 100) Then
                        Exit Do
                    End If
                
                Loop
                
                If (rstRecordsList.RecordCount <> 0) Then
                    
                    If (rstRecordsList.AbsolutePosition <> lngTempAbsPos) Then
                        rstRecordsList.AbsolutePosition = lngTempAbsPos
                        lngAbsolutePosition = lngTempAbsPos
                    End If
                
                End If
            
            End If
        
        End If
    
    End If
    
    jgxPicklist.Visible = True
    
    If (enuStyle = cpiFilterCatalog) Then
        
        ' check filter enabled
        If (rstRecordsList.RecordCount = 0) Then
            
            If (chkClearFilter.Value = vbUnchecked) Then
                chkClearFilter.Enabled = False
            End If
            
            cmdFilter.Enabled = False
        
        ElseIf (rstRecordsList.RecordCount <> 0) Then
            
            chkClearFilter.Enabled = True
            cmdFilter.Enabled = True
        
        End If
    
    End If
    
    ' check if there is zero record
    If (clsPicklist.AutoUnload = CPI_AUTOCANCEL) Then
    
        If (rstRecordsList.RecordCount = 0) Then
            cmdTransact(CMD_ADD).Enabled = False
        ElseIf (rstRecordsList.RecordCount <> 0) Then
            cmdTransact(CMD_ADD).Enabled = True
        End If
    
    ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
    
        If (rstRecordsList.RecordCount = 0) Then
            cmdTransact(CMD_ADD).Enabled = False
        ElseIf (rstRecordsList.RecordCount <> 0) Then
            cmdTransact(CMD_ADD).Enabled = True
        End If
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub jgxPicklist_Click()

    If (rstRecordsList.RecordCount <> 0) Then
    
        rstRecordsList.AbsolutePosition = lngAbsolutePosition
        
        If ((enuStyle = cpiFilterCatalog) Or _
            (jgxPicklist.ADORecordset.Filter = "Tag <> 'D'")) Then
        
            Call SetHeaderTextBoxes
        
        ElseIf (enuStyle = cpiSimplePicklist) Then
            Call SetHeaderTextBoxes
        End If
    
    End If
   
End Sub

Private Sub jgxPicklist_ColButtonClick(ByVal ColIndex As Integer)
    blnDoNotOpen = True 'vince
End Sub

Private Sub jgxPicklist_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX16.JSRetBoolean)
    
    Debug.Print "WOOHOOO"
    blnDoNotOpen = True 'vince
    
    
End Sub

Private Sub jgxPicklist_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    Static intLastSort As Integer
    
    
    
    blnDoNotOpen = True 'vince
    
    
    
    ' save the last column header clicked
    Set jgxActiveColumn = Column
    
    If (Column.SortOrder = intLastSort) Then
    
        strLastColumnSort = "[" & Column.Caption & "] ASC"
        
    ElseIf (Column.SortOrder <> intLastSort) Then
    
        If Column.SortOrder = intLastSort Then
            strLastColumnSort = "[" & Column.Caption & "] DESC"
        Else
            strLastColumnSort = "[" & Column.Caption & "] ASC"
        End If
    
    End If
        
    With jgxPicklist
    
        If (intLastSort = 0) Then
            .SortKeys.Clear
            .SortKeys.Add Column.Index, jgexSortAscending
            intLastSort = 1
        ElseIf (intLastSort <> 0) Then
            If (intLastSort = 1) Then
                .SortKeys.Clear
                .SortKeys.Add Column.Index, -1
                intLastSort = -1
            ElseIf (intLastSort <> 1) Then
                .SortKeys.Clear
                .SortKeys.Add Column.Index, 1
                intLastSort = 1
            End If
            
        End If
        
    End With
    
End Sub

Private Sub jgxPicklist_DblClick()
    
    Dim lngIndexArray As Long
    Dim strTmpArray() As String
    Dim objJSColumn As JSColumn
    
    If Not blnDoNotOpen Then
        If Not rstRecordsList Is Nothing Then
            If (rstRecordsList.RecordCount <> 0) Then
            
                Call RepositionRst
                
                
                If (rstRecordsList.AbsolutePosition <> lngAbsolutePosition) Then
                
                    If (lngAbsolutePosition > 0) Then
                        rstRecordsList.AbsolutePosition = lngAbsolutePosition
                    End If
                    
                End If
                
                Select Case enuStyle
                
                    Case cpiFilterCatalog, cpiSimplePicklist
                        If Not clsPicklist Is Nothing Then
                            If (clsPicklist.AutoUnload = CPI_AUTOCANCEL) Then
                            
                                blnCanceled = False
                                Call cmdTransact_Click(CMD_OK)
                                
                            ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
                            
                                blnCanceled = False
                                Call cmdTransact_Click(CMD_OK)
                            
                            ElseIf (clsPicklist.AutoUnload = CPI_FALSE) Then
                            
                                Call cmdCatalogOps_Click(CMD_MODIFY)
                            
                            End If
                        End If
                    
                    Case cpiCatalog
                        
                        If Not clsPicklist Is Nothing Then
                            If (clsPicklist.AutoUnload = CPI_AUTOCANCEL) Then
                                
                                blnCanceled = False
                                Call cmdTransact_Click(CMD_OK)
                            
                            ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
                            
                                blnCanceled = False
                                Call cmdTransact_Click(CMD_OK)
                            
                            ElseIf (clsPicklist.AutoUnload = CPI_FALSE) Then
                            
                                If ((jgexHitTestConstant = jgexHTColumnHeader) = False) Then
                                
                                    Call cmdCatalogOps_Click(CMD_MODIFY)
                                    
                                End If
                            
                            End If
                        End If
                        
                    Case Else
                    
                        ' trap error
                        'Stop
                        
                End Select
                
            End If
        End If
    Else
        blnDoNotOpen = False
        If jgexHitTestConstant = 2 Then
            For lngIndexArray = 0 To UBound(strColumnWidth)
                strTmpArray() = Split(strColumnWidth(lngIndexArray), "$$")
                Set objJSColumn = jgxPicklist.ColFromPoint(lngXCoordinate, lngYCoordinate)
                If Not objJSColumn Is Nothing Then
                    If UCase(Trim(strTmpArray(0))) = UCase(Trim(objJSColumn.Caption)) Then
                        jgxPicklist.ColFromPoint(lngXCoordinate, lngYCoordinate).Width = CLng(strTmpArray(1))
                    End If
                Else
                    Debug.Assert False
                End If
            Next
        End If
    End If
    
End Sub

Private Sub jgxPicklist_GroupByBoxHeaderClick(ByVal Group As GridEX16.JSGroup)
    blnDoNotOpen = True
End Sub

Private Sub jgxPicklist_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyReturn) Then
    
        If (rstRecordsList.RecordCount <> 0) Then
        
            Call jgxPicklist_DblClick
            KeyCode = 0
            
        End If
        
    ElseIf (KeyCode = vbKeyEscape) Then
    
        Call cmdTransact_Click(CMD_CANCEL)
        
    End If
   
End Sub

Private Sub jgxPicklist_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If rstRecordsList.RecordCount <> 0 Then
    
        Select Case KeyCode
        
            Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
                ' reconcile true position of grid
                If rstRecordsList.AbsolutePosition <> lngAbsolutePosition Then
                    If lngAbsolutePosition > 0 Then
                        rstRecordsList.AbsolutePosition = lngAbsolutePosition
                    End If
                End If
                
                SetHeaderTextBoxes
            
        End Select
        
    End If

End Sub

Private Sub jgxPicklist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRowFromPoint As Long
    Dim rstClone As ADODB.Recordset
    
    ' test object pointed by mouse
    jgexHitTestConstant = jgxPicklist.HitTest(x, y)
     lngXCoordinate = x
    lngYCoordinate = y
    
    If jgexHitTestConstant = jgexHTCell Then
        lngRowFromPoint = jgxPicklist.RowFromPoint(lngXCoordinate, lngYCoordinate)
        'jgxPicklist.MoveToRowIndex lngRowFromPoint
        Set rstClone = jgxPicklist.ADORecordset.Clone
        rstClone.Bookmark = jgxPicklist.RowBookmark(lngRowFromPoint)
        
        If Not IsNull(rstClone.Fields(jgxPicklist.Columns(jgxPicklist.ColFromPoint(lngXCoordinate, lngYCoordinate)).Caption).Value) Then
            jgxPicklist.ToolTipText = rstClone.Fields(jgxPicklist.Columns(jgxPicklist.ColFromPoint(lngXCoordinate, lngYCoordinate)).Caption).Value
        Else
            jgxPicklist.ToolTipText = ""
        End If

        'jgxPicklist.MoveToBookmark jgxPicklist.RowBookmark(lngRowFromPoint)
        'jgxPicklist.ToolTipText = jgxPicklist.Value(jgxPicklist.Columns(jgxPicklist.ColFromPoint(lngXCoordinate, lngYCoordinate)).Index)
    End If
End Sub

Private Sub jgxPicklist_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    
    If ((rstRecordsList Is Nothing) = False) Then
        If (rstRecordsList.RecordCount <> 0) Then
            If (lngAbsolutePosition <> rstRecordsList.AbsolutePosition) Then
                lngAbsolutePosition = rstRecordsList.AbsolutePosition
            End If
        End If
    End If

End Sub
   
Private Sub optFilter_Click(Index As Integer)
    
    Dim lngFilterCtr As Long
    
    If (blnFormActivated = True) Then
    
        chkClearFilter.Value = vbUnchecked
        strHeaderFilter = ""
        
        ' remove all state
        For lngFilterCtr = 1 To clsFilter.FilterCount
            clsFilter.PicklistFilters(lngFilterCtr).State = False
            clsFilter.PicklistFilters(lngFilterCtr).Value = False
        Next lngFilterCtr
        
        For lngFilterCtr = 1 To clsFilter.FilterCount
            If optFilter(lngFilterCtr - 1).Visible And optFilter(lngFilterCtr - 1).Value = True Then
                strHeaderFilter = strHeaderFilter & clsFilter.PicklistFilters(lngFilterCtr).Filter
                clsFilter.PicklistFilters(lngFilterCtr).State = True
                Exit For
            End If
        Next lngFilterCtr
        
        RefreshGrid cpiRefresh
        
    End If
        
    If ((rstRecordsList Is Nothing) = True) Then
        ' do nothing
    ElseIf ((rstRecordsList Is Nothing) = False) Then
    
        If (rstRecordsList.RecordCount = 0) Then
            cmdCatalogOps(CMD_MODIFY).Enabled = False
            cmdCatalogOps(CMD_COPY).Enabled = False
            cmdCatalogOps(CMD_DELETE).Enabled = False
            ' clear textfields
            Call UpdateSearchField
        ElseIf (rstRecordsList.RecordCount <> 0) Then
            cmdCatalogOps(CMD_MODIFY).Enabled = True
            cmdCatalogOps(CMD_COPY).Enabled = True
            cmdCatalogOps(CMD_DELETE).Enabled = True
            ' clear textfields
            Call UpdateSearchField
        End If
    End If
    
End Sub


Private Sub txtSearchFields_Change(Index As Integer)
   
    Dim strSQL As String
    Dim strWhereStart As String
    Dim strWhereEnd As String
    Dim varValue As Variant
    
    'edited by tonio 09292003
    'If ((blnChangeDueToGridClick = False) And ((Index = 0) Or (Index = 1)) And (blnSkipTextChange = False)) Then
    If ((blnChangeDueToGridClick = False) And (blnSkipTextChange = False)) Then
        
        Call LoadGrid(cpiManyRecord, txtSearchFields(Index).Tag, txtSearchFields(Index).Text)
        Call UpdateSearchField(Index)
        
        If (jgxPicklist.ADORecordset.EOF = False) Then
        
            jgxPicklist.MoveToBookmark jgxPicklist.ADORecordset.Bookmark
            
        ElseIf (jgxPicklist.ADORecordset.EOF = True) Then
        
            If (rstRecordsList.RecordCount <> 0) Then
            
                jgxPicklist.ADORecordset.MoveFirst
                jgxPicklist.ADORecordset.Find strSQL, , adSearchForward
                
            End If
        
        End If
    
    End If
   
End Sub

' Fan-Out: SetGridProperty
Private Function GetColIndex(ByRef coliRef As GridEX16.JSColumns, ByVal intiRefIndex As Integer) As Integer
   
    ' returns the matched index of the txtfields compare with visible columns of the grid
    Dim intIndex As Integer
    Dim intFound As Integer
    
    For intIndex = 1 To coliRef.Count
    
        If (coliRef(intIndex).Visible = True) Then
        
            intFound = intFound + 1
            
            If (intFound = intiRefIndex) Then
            
                GetColIndex = intIndex
                Exit Function
                
            End If
            
        End If
    
    Next intIndex
    
    GetColIndex = 0
   
End Function

' Fan-Out: Form_Activate,jgxPicklist_Click,jgxPicklist_KeyUp
Public Sub SetHeaderTextBoxes()
    
    Dim lngColumnCtr As Long
    
    If (rstRecordsList.AbsolutePosition = adPosUnknown) Then
        Exit Sub
    End If
    
    If (rstRecordsList.EOF = True) Then
        rstRecordsList.MoveLast
    End If
        
    If (blnGridIsEmpty = False) Then
    
        blnChangeDueToGridClick = True
        
        For lngColumnCtr = 0 To clsGridSeed.GridColumns.Count - 1
        
            txtSearchFields(lngColumnCtr).Text = _
                                        IIf(IsNull(jgxPicklist.ADORecordset.Fields(clsGridSeed.GridColumns(lngColumnCtr + 1) _
                                        .ColumnFieldAias).Value), "", _
                                        jgxPicklist.ADORecordset.Fields(clsGridSeed.GridColumns(lngColumnCtr + 1) _
                                        .ColumnFieldAias).Value)
            
        Next lngColumnCtr
        
        blnChangeDueToGridClick = False
        
    End If
    
End Sub

' Fan-Out: txtSearchFields_Click
Public Sub EnableDisableButtons(ByVal GridIsEmpty As Boolean)
    
    cmdCatalogOps(1).Enabled = Not GridIsEmpty
    cmdCatalogOps(2).Enabled = Not GridIsEmpty
    cmdCatalogOps(3).Enabled = Not GridIsEmpty

End Sub

Private Sub txtSearchFields_Click(Index As Integer)
    Call EnableDisableButtons(blnGridIsEmpty)
End Sub

Private Sub txtSearchFields_GotFocus(Index As Integer)
    Dim lngSearchBoxCtr As Long
    
'Commented and mod by BCo
'To allow picklist filtering via first and second text box.
'    If Index <> 0 And Index <> 1 Then
'        For lngSearchBoxCtr = 2 To 8
        For lngSearchBoxCtr = 0 To 8
            If lngSearchBoxCtr = Index Then
                txtSearchFields(lngSearchBoxCtr).BackColor = &HC0FFFF
            Else
                txtSearchFields(lngSearchBoxCtr).BackColor = &H80000005
            End If
        Next lngSearchBoxCtr
'    End If
    
    txtSearchFields(Index).SelStart = 0
    txtSearchFields(Index).SelLength = Len(txtSearchFields(Index))
End Sub

' Fan-Out: RunEditTrans
Public Sub UpdateGrid(ByRef ModifiedRecord As CRecord)
    
    Dim fldDummy As ADODB.Field
    
    ' selected fields
    For Each fldDummy In jgxPicklist.ADORecordset.Fields
    
        If (UCase$(fldDummy.Name) <> "TAG") Then
        
            jgxPicklist.ADORecordset.Fields(fldDummy.Name).Value = _
                                ModifiedRecord.RecordSource.Fields(fldDummy.Name).Value
            
        ElseIf (UCase$(fldDummy.Name) = "TAG") Then
        
            If (jgxPicklist.ADORecordset.Fields(fldDummy.Name).Value <> "A") Then
            
                jgxPicklist.ADORecordset.Fields(fldDummy.Name).Value = "M"
                
            End If
            
        End If
        
    Next 'fldDummy
    
    Set fldDummy = Nothing
    
End Sub

' Fan-Out: CreateRstToGrid, ApplyFetchGrid
Private Function FetchGridRecords(ByRef Source As String, _
                        ByRef conToUse As ADODB.Connection, _
                        ByRef rstToOpen As ADODB.Recordset, _
                        ByRef CursorType As CursorTypeEnum, _
                        ByRef LockType As LockTypeEnum, _
                        ByRef PKFieldName As String, _
                        ByRef GridSeed As CGridSeed) As String
                                                               
    Dim rstDummy As ADODB.Recordset
    Dim fldDummy As ADODB.Field
    
    Dim strSourceLeftOfFrom As String
    Dim strSourceRightOfFrom As String
    Dim strOriginalSource As String
    Dim strFinalSource As String
    Dim strRunSQL As String
    
    ' counter for rstToOpen recordset's fields
    Dim intFieldIndex As Integer
    
    'Dim clsRecordset As CRecordset
    
    ' ************************************************ '
    
    ' model field
    Set rstFieldModel = New ADODB.Recordset
    strRunSQL = RegenerateSQL(clsPicklist.BaseSQL, False)
    
    'Set clsRecordset = New CRecordset
    
    ADORecordsetOpen strRunSQL, conDBConnection, rstFieldModel, adOpenKeyset, adLockOptimistic
    'clsRecordset.cpiOpen strRunSQL, conDBConnection, rstFieldModel, adOpenForwardOnly, adLockPessimistic, , True
    'Set rstFieldModel = rstDBCopy(conDBConnection, strRunSQL, _
                                    adOpenForwardOnly, adLockPessimistic, True)
    
    'Set clsRecordset = Nothing
    
    On Error GoTo ERROR_HANDLER_BOOKMARK
    
    If (rstFieldModel.RecordCount = 0) Then
        rstFieldModel.AddNew
        rstFieldModel.Update
    End If
    
    strOriginalSource = clsPicklist.BaseSQL
    
    ADORecordsetOpen strOriginalSource, conToUse, rstDummy, adOpenKeyset, adLockOptimistic
    
    'Set rstDummy = New ADODB.Recordset
    'rstDummy.Open strOriginalSource, conToUse
    
    ' return the first field name usually the Primary Key
'    FetchGridRecords = rstDummy.Fields(PKFieldName).Properties(0).Value
    FetchGridRecords = IIf(IsNull(rstDummy.Fields(PKFieldName).Properties(0).Value), PKFieldName, rstDummy.Fields(PKFieldName).Properties(0).Value)
    
    Call ApplyFetchGrid(Source, conToUse, rstToOpen, CursorType, LockType, PKFieldName _
                                , GridSeed, rstDummy, fldDummy, strSourceLeftOfFrom _
                                , strSourceRightOfFrom, strOriginalSource, strFinalSource, intFieldIndex)
    
    ADORecordsetClose rstDummy

    Set fldDummy = Nothing
    
    Exit Function
    
ERROR_HANDLER_BOOKMARK:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation, frmOwnerForm.Caption
    ' error flag
    blnHaltExecution = True
    
End Function

' Fan-Out: ProcessRecord
Public Sub CommitTransaction(ByRef DBConnection As ADODB.Connection, _
                                                        ByVal SQLCommit As String)
    
    Dim strPKFieldAlias As String
    Dim strSQLCommit As String
        
    Dim lngTransactionCnt As Long
    
    Dim rstRecordsCommit As ADODB.Recordset
    Dim fldRecordsCommit As ADODB.Field
    
    'Dim clsRecordsetCommit As CRecordset
    Dim straCommitFields() As String
    
    Dim strTableName As String
    Dim strCommandCommit As String
    
    Dim lngCtr As Long
    ' ************************************************ '
    
    If ((clsPicklist.Transactions Is Nothing) = False) Then
    
        If (clsPicklist.Transactions.Count > 0) Then
    
            'Set clsRecordsetCommit = New CRecordset
            
            If Len(Trim(clsPicklist.MainSQL)) = 0 Then
                strSQLCommit = RegenerateSQL(SQLCommit, True)  ' new code
            Else
                strSQLCommit = clsPicklist.MainSQL
            End If
            
            strTableName = GetSQLFromTable(strSQLCommit)
            
            ADORecordsetOpen strSQLCommit, DBConnection, rstRecordsCommit, adOpenKeyset, adLockOptimistic
            'Call clsRecordsetCommit.cpiOpen(strSQLCommit, DBConnection, _
                                                    rstRecordsCommit, adOpenKeyset, adLockOptimistic)
            
            For Each fldRecordsCommit In rstRecordsCommit.Fields
            
                If (UCase$(fldRecordsCommit.Properties(0).Value) = UCase$(clsPicklist.PKFieldBaseName)) Then
                    strPKFieldAlias = fldRecordsCommit.Name
                    Exit For
                End If
                
            Next ' fldRecordsCommit
                        
            For lngTransactionCnt = 1 To clsPicklist.Transactions.Count
            
                Select Case clsPicklist.Transactions(lngTransactionCnt).Status
                
                    Case cpiStateDeleted
                        '------------------------------------------------------------------
                        'Entrepot>Products picklist triggers a delete on both Entrepots & Products tables.
                        'Since we only want to delete a record in Products, this checks if an alternate
                        'SQL has been passed for single table deletion. If empty, just proceeds normally.
                        If Len(clsPicklist.DeleteSQL4InnerJoinCases) = 0 Then
                        '------------------------------------------------------------------
                            
                            rstRecordsCommit.MoveFirst
                            rstRecordsCommit.Find GetRecordCriteria(clsPicklist.PKFieldBaseName, _
                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                            .Fields(strPKFieldAlias).Value, _
                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                            .Fields(strPKFieldAlias).Type), , adSearchForward
                            
                            If (rstRecordsCommit.EOF = False) Then
                                rstRecordsCommit.Delete
                                
                                    strCommandCommit = vbNullString
                                    strCommandCommit = strCommandCommit & "DELETE "
                                    strCommandCommit = strCommandCommit & " * "
                                    strCommandCommit = strCommandCommit & "FROM "
                                    strCommandCommit = strCommandCommit & "[" & strTableName & "] "
                                    strCommandCommit = strCommandCommit & "WHERE "
                                    strCommandCommit = strCommandCommit & GetRecordCriteria(clsPicklist.PKFieldBaseName, _
                                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                                            .Fields(strPKFieldAlias).Value, _
                                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                                            .Fields(strPKFieldAlias).Type)
                                ExecuteNonQuery conDBConnection, strCommandCommit
                            End If
                                
                            'strTableName
                        '------------------------------------------------------------------
                        'Creates a new recordset for single table deletion purposes.
                        Else
                            Dim rstAltDelete As ADODB.Recordset
                            
                            strTableName = GetSQLFromTable(clsPicklist.DeleteSQL4InnerJoinCases)
                            
                            ADORecordsetOpen clsPicklist.DeleteSQL4InnerJoinCases, conDBConnection, rstAltDelete, adOpenKeyset, adLockOptimistic
                            'Set rstAltDelete = New ADODB.Recordset
                            'rstAltDelete.Open clsPicklist.DeleteSQL4InnerJoinCases, conDBConnection, adOpenKeyset, adLockOptimistic
                            
                            rstAltDelete.MoveFirst
                            rstAltDelete.Find GetRecordCriteria(clsPicklist.PKFieldBaseName, _
                                              clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                              .Fields(strPKFieldAlias).Value, _
                                              clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                              .Fields(strPKFieldAlias).Type), , adSearchForward
                            If (rstAltDelete.EOF = False) Then
                                rstAltDelete.Delete
                                
                                    strCommandCommit = vbNullString
                                    strCommandCommit = strCommandCommit & "DELETE "
                                    strCommandCommit = strCommandCommit & " * "
                                    strCommandCommit = strCommandCommit & "FROM "
                                    strCommandCommit = strCommandCommit & "[" & strTableName & "] "
                                    strCommandCommit = strCommandCommit & "WHERE "
                                    strCommandCommit = strCommandCommit & GetRecordCriteria(clsPicklist.PKFieldBaseName, _
                                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                                            .Fields(strPKFieldAlias).Value, _
                                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                                            .Fields(strPKFieldAlias).Type)
                                ExecuteNonQuery conDBConnection, strCommandCommit
                                
                            End If
                            
                            ADORecordsetClose rstAltDelete
                            'rstAltDelete.Close
                            'Set rstAltDelete = Nothing
                        End If
                        '------------------------------------------------------------------
                    Case cpiStateModified
                        'Initialize the string containing fields not affected by rst event.
                        straCommitFields = Split(mDontCommitFields, ",")
                        
                        rstRecordsCommit.MoveFirst
                        rstRecordsCommit.Find GetRecordCriteria(clsPicklist.PKFieldBaseName, _
                                                        clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                        .Fields(strPKFieldAlias).Value, _
                                                        clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                        .Fields(strPKFieldAlias).Type), , adSearchForward
                        
                        If (rstRecordsCommit.EOF = False) Then
                            
                            Dim rst As ADODB.Recordset
                                
                            
                            Set rst = New ADODB.Recordset
                            Set rst = rstDBCopy(conDBConnection, RegenerateSQL(clsPicklist.BaseSQL, False), _
                                                            adOpenKeyset, adLockOptimistic, False)
                            
                            For Each fldRecordsCommit In rstRecordsCommit.Fields
                                'Perform check to avoid speficied fields from being affected by rst action.
                                For lngCtr = 0 To UBound(straCommitFields)
                                    If UCase(straCommitFields(lngCtr)) = UCase(fldRecordsCommit.Name) Then GoTo SKIPCOMMIT2
                                Next

                                
                                If (rst.Fields(fldRecordsCommit.Name).Properties(2).Value = False) Then
    
                                    rstRecordsCommit.Fields(fldRecordsCommit.Name).Value = _
                                                                            clsPicklist.Transactions(lngTransactionCnt).RecordSource _
                                                                            .Fields(fldRecordsCommit.Name).Value
                                        
                                End If
                                
SKIPCOMMIT2:                Next
                        
                            rstRecordsCommit.Update
                            
                            
                            ExecuteRecordset ExecuteRecordsetConstant.Update, DBConnection, rstRecordsCommit, strTableName
                            
                            ADORecordsetClose rst
                            'Set rst = Nothing
                            
                        End If
                        
                    Case cpiStateNew
                        straCommitFields = Split(mDontCommitFields, ",")
                        
                        rstRecordsCommit.AddNew
                        For Each fldRecordsCommit In rstRecordsCommit.Fields
                            For lngCtr = 0 To UBound(straCommitFields)
                                If UCase(straCommitFields(lngCtr)) = UCase(fldRecordsCommit.Name) Then GoTo SKIPCOMMIT
                            Next
                            
                            
                            
                            If (blnDBAutoNumber = True) Then
                                rstRecordsCommit.Fields(fldRecordsCommit.Name).Value = _
                                IIf(IsNull(clsPicklist.Transactions(lngTransactionCnt).RecordSource.Fields(fldRecordsCommit.Name).Value), _
                                        0, clsPicklist.Transactions(lngTransactionCnt).RecordSource.Fields(fldRecordsCommit.Name).Value)
                            Else
                                rstRecordsCommit.Fields(fldRecordsCommit.Name).Value = _
                                IIf(IsNull(clsPicklist.Transactions(lngTransactionCnt).RecordSource.Fields(fldRecordsCommit.Name).Value), _
                                        "", clsPicklist.Transactions(lngTransactionCnt).RecordSource.Fields(fldRecordsCommit.Name).Value)
                            End If
                        
                        
SKIPCOMMIT:              Next

                        On Error GoTo RollBack_Record
                        rstRecordsCommit.Update
                        
                        InsertRecordset DBConnection, rstRecordsCommit, strTableName
                        
                End Select
                
            Next lngTransactionCnt
            
            ADORecordsetClose rstRecordsCommit
            'Call clsRecordsetCommit.cpiClose(rstRecordsCommit)
        
            'Set rstRecordsCommit = Nothing
            'Set clsRecordsetCommit = Nothing
        
        End If
    
    End If
    
    Exit Sub
    
RollBack_Record:

    ' rollback transaction here
    MsgBox Err.Description & ". The record of " & _
                rstRecordsCommit.Fields(clsPicklist.PKFieldAlias).Name & "=" & _
                CStr(rstRecordsCommit.Fields(clsPicklist.PKFieldAlias).Value) & _
                " cannot be saved.", vbInformation, Me.Caption
                                       
   Resume Next

End Sub

' Fan-Out: dcbFilter_Change
'Private Sub UpdateDataCombo(ByRef strSQL As String, ByRef Index As Integer)
'
'    Dim blnExist As Boolean
'
'    ' code to repopulate succeeding filter datacombos
'    strSQL = clsFilter.PicklistFilters(Index + 2).Filter
'    blnExist = (InStr(1, strSQL, " WHERE ") > 0)
'
'    ' check if where clause exist
'    If (blnExist = True) Then
'        strSQL = strSQL & " AND "
'    ElseIf (blnExist = False) Then
'        strSQL = strSQL & " WHERE "
'    End If
'
'    strSQL = strSQL & rstFilterRecordsets(Index).Fields(0).Name & " = " _
'                                        & Val(dcbFilter(Index).BoundText)
'
'    clsRecordset.cpiOpen strSQL, conDBConnection, rstFilterRecordsets(Index + 1), _
'                                        adOpenKeyset, adLockReadOnly
'
'    If (rstFilterRecordsets(Index + 1).RecordCount > 0) Then
'        rstFilterRecordsets(Index + 1).MoveFirst
'    ElseIf (rstFilterRecordsets(Index + 1).RecordCount = 0) Then
'        Exit Sub
'    End If
'
'    Set dcbFilter(Index + 1).DataSource = Nothing
'    Set dcbFilter(Index + 1).RowSource = Nothing
'    dcbFilter(Index + 1).BoundColumn = ""
'    dcbFilter(Index + 1).BoundText = ""
'    dcbFilter(Index + 1).DataField = ""
'    dcbFilter(Index + 1).ListField = ""
'    dcbFilter(Index + 1).Refresh
'
'    Set dcbFilter(Index + 1).DataSource = rstFilterRecordsets(Index + 1)
'    Set dcbFilter(Index + 1).RowSource = rstFilterRecordsets(Index + 1)
'    dcbFilter(Index + 1).BoundColumn = rstFilterRecordsets(Index + 1).Fields(0).Name
'    dcbFilter(Index + 1).BoundText = rstFilterRecordsets(Index + 1).Fields(0).Name
'    dcbFilter(Index + 1).DataField = rstFilterRecordsets(Index + 1).Fields(1).Name
'    dcbFilter(Index + 1).ListField = rstFilterRecordsets(Index + 1).Fields(1).Name
'
'    dcbFilter(Index + 1).Text = rstFilterRecordsets(Index + 1).Fields(1).Value
'
'End Sub

' Fan-Out: CheckPicklistStyle
Private Sub InitCheckBoxCombo(ByRef intColumnCtr As Integer, _
                                                    ByRef dblLastRightPosition As Double, _
                                                    ByRef dblDummyTop As Double, _
                                                    ByRef dblTextSearchTopOffset As Double)
      
    ' initialize checkboxes, combo boxes
    For intColumnCtr = 1 To clsFilter.FilterCount
    
        Select Case clsFilter.FilterType
        
            Case cpiCheckOptions
            
                Call InitCheckOptions(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
            
            Case cpiComboRecords
            
                'Call InitComboRecords(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
            
            Case cpiRadioOptions
            
                Call InitRadioOption(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
        
        End Select
    
    Next intColumnCtr

End Sub

' Fan-Out: InitCheckBoxCombo
Private Sub InitCheckOptions(ByRef intColumnCtr As Integer, _
                                                   ByRef dblLastRightPosition As Double, _
                                                   ByRef dblDummyTop As Double, _
                                                   ByRef dblTextSearchTopOffset As Double)

    ' start here -  check if both
    Dim blnVisible As Boolean
    
    blnVisible = clsFilter.PicklistFilters(intColumnCtr).Visible
    chkFilter(intColumnCtr - 1).Visible = blnVisible
    chkFilter(intColumnCtr - 1).Left = tabCatalog.Left + 90 + 180
    chkFilter(intColumnCtr - 1).Caption = clsFilter.PicklistFilters(intColumnCtr).FilterCaption
    chkFilter(intColumnCtr - 1).Value = IIf(clsFilter.PicklistFilters(intColumnCtr).State, _
                                                        vbChecked, vbUnchecked)
    chkFilter(intColumnCtr - 1).Top = dblDummyTop
    
    If (clsFilter.PicklistFilters(intColumnCtr).State = True) Then
    
        strHeaderFilter = strHeaderFilter & clsFilter.PicklistFilters(intColumnCtr).Filter & " OR "
        
    End If
    
    dblDummyTop = dblDummyTop + chkFilter(intColumnCtr - 1).Height + 45

End Sub

' Fan-Out: InitCheckBoxCombo
'Private Sub InitComboRecords(ByRef intColumnCtr As Integer, _
'                                                ByRef dblLastRightPosition As Double, _
'                                                ByRef dblDummyTop As Double, _
'                                                ByRef dblTextSearchTopOffset As Double)
'
'   dcbFilter(intColumnCtr - 1).Visible = clsFilter.PicklistFilters(intColumnCtr).Visible
'   dcbFilter(intColumnCtr - 1).Left = tabCatalog.Left + 90 + 1750
'   dcbFilter(intColumnCtr - 1).Width = 2750
'   dcbFilter(intColumnCtr - 1).Top = dblDummyTop
'
'   lblFilter(intColumnCtr - 1).Visible = True
'   lblFilter(intColumnCtr - 1).Left = tabCatalog.Left + 90 + 180
'   lblFilter(intColumnCtr - 1).Width = 2750
'   lblFilter(intColumnCtr - 1).Caption = clsFilter.PicklistFilters(intColumnCtr).FilterCaption _
'                                                                     & " :"
'   lblFilter(intColumnCtr - 1).Top = dcbFilter(intColumnCtr - 1).Top + 55 - tabCatalog.Top
'
'   dblDummyTop = dblDummyTop + dcbFilter(intColumnCtr - 1).Height + 45
'
'End Sub

' Fan-Out: InitCheckBoxCombo
Private Sub InitRadioOption(ByRef intColumnCtr As Integer, _
                                                ByRef dblLastRightPosition As Double, _
                                                ByRef dblDummyTop As Double, _
                                                ByRef dblTextSearchTopOffset As Double)
   
    optFilter(intColumnCtr - 1).Visible = CBool(clsFilter.PicklistFilters(intColumnCtr).Visible)
    optFilter(intColumnCtr - 1).Left = tabCatalog.Left + 90 + 180
    optFilter(intColumnCtr - 1).Caption = clsFilter.PicklistFilters(intColumnCtr).FilterCaption
    optFilter(intColumnCtr - 1).Top = dblDummyTop
    optFilter(intColumnCtr - 1).Value = clsFilter.PicklistFilters(intColumnCtr).State
    
    If (clsFilter.PicklistFilters(intColumnCtr).State = True) Then
        
        strHeaderFilter = clsFilter.PicklistFilters(intColumnCtr).Filter & " OR "
    
    End If
    
    
    If (clsFilter.PicklistFilters(intColumnCtr).Visible = True) Then
    
        dblDummyTop = dblDummyTop + optFilter(intColumnCtr - 1).Height + 45
    
    End If

End Sub

' Fan-Out: FormInitialize
Private Sub CheckPicklistStyle(ByRef intColumnCtr As Integer, _
                    ByRef dblLastRightPosition As Double, _
                    ByRef dblDummyTop As Double, _
                    ByRef dblTextSearchTopOffset _
                    As Double)

    ' check if there is a filter
    If (((clsFilter Is Nothing) = False) And (enuStyle <> cpiSimplePicklist)) Then
    
        dblDummyTop = lblListDescription.Top + 180
        
        Call InitCheckBoxCombo(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
        
        ' initialize label fields
        If (clsFilter.FilterCount > 0) Then
        
            Call InitControl
        
        End If
        
    End If

End Sub

' Fan-Out: CheckPicklistStyle
Private Sub InitControl()

    Select Case clsFilter.FilterType
    
        Case cpiCheckOptions
            If clsFilter.PicklistFilters(clsFilter.FilterCount).Visible = True Then
            
                lblListDescription.Top = chkFilter(clsFilter.FilterCount - 1).Top _
                    + chkFilter(clsFilter.FilterCount - 1).Height + 90
                    
            End If
            
            If Trim$(strHeaderFilter) <> "" Then
            
                strHeaderFilter = Left(strHeaderFilter, Len(strHeaderFilter) - 4)
                
            End If
        
        Case cpiComboRecords
        
'            If dcbFilter(clsFilter.FilterCount - 1).Visible = True Then
'
'                lblListDescription.Top = dcbFilter(clsFilter.FilterCount - 1).Top _
'                        + dcbFilter(clsFilter.FilterCount - 1).Height + 90
'
'            End If
        
        Case cpiRadioOptions
        
            If clsFilter.PicklistFilters(clsFilter.FilterCount).Visible = True Then
            
                lblListDescription.Top = optFilter(clsFilter.FilterCount - 1).Top _
                        + optFilter(clsFilter.FilterCount - 1).Height + 90
            
            End If
            
            If Trim$(strHeaderFilter) <> "" Then
            
                strHeaderFilter = Left(strHeaderFilter, Len(strHeaderFilter) - 4)
                
            End If
        
    End Select

End Sub

' Fan-Out: FormInitialize
Private Sub InitTextFields(ByRef intColumnCtr As Integer, _
                                            ByRef dblLastRightPosition As Double, _
                                            ByRef dblDummyTop As Double, _
                                            ByRef dblTextSearchTopOffset As Double)

    Select Case enuStyle
    
        Case cpiSimplePicklist
            cmdTransact(0).Caption = "Select"
            Me.Caption = "Picklist - " & strPluralEntity
            
        Case cpiCatalog, cpiFilterCatalog
        
            ' apply text fields
            Call ApplyTextFields(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
            ' initialize tab list
            Call InitTabList(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
            
    End Select

End Sub

' Fan-Out: FormInitialize
Private Sub InitGrid(ByRef intColumnCtr As Integer, _
                                 ByRef dblLastRightPosition As Double, _
                                 ByRef dblDummyTop As Double, _
                                 ByRef dblTextSearchTopOffset As Double)

    dblLastRightPosition = jgxPicklist.Left
    
    If ((clsGridSeed Is Nothing) = False) Then
    
        Set jgxPicklist.ADORecordset = Nothing
        ' reference to grid
        Set jgxPicklist.ADORecordset = rstRecordsList
        
        For intColumnCtr = 1 To clsGridSeed.GridColumns.Count
        
            strDummySettings = strDummySettings & "||" _
                                            & clsGridSeed.GridColumns.Item(intColumnCtr).ColumnFieldAias
            ' load column values from class
            Call LoadColumnValue(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
        
        Next intColumnCtr
        
        strDummySettings = strDummySettings & "||"
    
    End If
    
End Sub

' Fan-Out: InitDataComboExt,FormInitialize
Private Sub InitDataCombo(ByRef intColumnCtr As Integer, _
                                            ByRef dblLastRightPosition As Double, _
                                            ByRef dblDummyTop As Double, _
                                            ByRef dblTextSearchTopOffset As Double)

    If ((clsFilter Is Nothing) = False) And (enuStyle <> cpiSimplePicklist) Then
    
        If (clsFilter.FilterType = cpiComboRecords) Then
        
            ReDim rstFilterRecordsets(clsFilter.FilterCount)
            
            ADORecordsetOpen clsFilter.PicklistFilters(0 + 1).Filter, conDBConnection, rstFilterRecordsets(0), adOpenKeyset, adLockOptimistic
            'clsRecordset.cpiOpen clsFilter.PicklistFilters(0 + 1).Filter, conDBConnection, rstFilterRecordsets(0), adOpenKeyset, adLockReadOnly
            
            If (rstFilterRecordsets(0).RecordCount > 0) Then
            
                rstFilterRecordsets(0).MoveFirst
                
            End If
            
'            Set dcbFilter(0).DataSource = rstFilterRecordsets(0)
'            Set dcbFilter(0).RowSource = rstFilterRecordsets(0)
'            dcbFilter(0).BoundColumn = rstFilterRecordsets(0).Fields(0).Name
'            dcbFilter(0).BoundText = rstFilterRecordsets(0).Fields(0).Name
'            dcbFilter(0).DataField = rstFilterRecordsets(0).Fields(1).Name
'            dcbFilter(0).ListField = rstFilterRecordsets(0).Fields(1).Name
'            dcbFilter(0).Text = rstFilterRecordsets(0).Fields(1).Value
        
        End If
    
    End If
    
End Sub

' Fan-Out: FormInitialize
Private Sub InitDataComboExt(ByRef intColumnCtr As Integer, _
                                                ByRef dblLastRightPosition As Double, _
                                                ByRef dblDummyTop As Double, _
                                                ByRef dblTextSearchTopOffset As Double)

    If (clsGridSeed.GridColumns.Count > 2) Then
    
        jgxPicklist.Width = dblLastRightPosition - jgxPicklist.Left
        
    End If
    
    lblListDescription.Width = jgxPicklist.Width
    
    If ((clsFilter Is Nothing) = False) And (enuStyle <> cpiSimplePicklist) Then
    
        For intColumnCtr = 0 To clsFilter.FilterCount - 1
        
            Select Case clsFilter.FilterType
            
                Case cpiCheckOptions
                    chkFilter(intColumnCtr).Width = jgxPicklist.Width - chkFilter(intColumnCtr).Left
                
                Case cpiRadioOptions
                    'optFilter(intColumnCtr).Width = Abs(jgxPicklist.Width - optFilter(intColumnCtr).Left)
                    
                    If tabCatalog.Width > jgxPicklist.Width Then
                        optFilter(intColumnCtr).Width = Abs(tabCatalog.Width - optFilter(intColumnCtr).Left)
                    Else
                        optFilter(intColumnCtr).Width = Abs(jgxPicklist.Width - optFilter(intColumnCtr).Left)
                    End If
            End Select
            
        Next intColumnCtr
    
    End If

End Sub

' Fan-Out: FormInitialize
Private Sub InitPickStyle(ByRef intColumnCtr As Integer, _
                                    ByRef dblLastRightPosition As Double, _
                                    ByRef dblDummyTop As Double, _
                                    ByRef dblTextSearchTopOffset As Double)
    
    Select Case enuStyle
    
        Case cpiSimplePicklist
        
            cmdTransact(1).Left = jgxPicklist.Left + jgxPicklist.Width - 1215
            cmdTransact(0).Left = cmdTransact(1).Left - 100 - 1215
            
            '<<< dandan 110707
            cmdSeeOnLine.Top = cmdTransact(0).Top
            cmdSeeOnLine.Left = txtSearchFields(0).Left
        
        Case cpiCatalog, cpiFilterCatalog
        
            tabCatalog.Width = 90 + jgxPicklist.Width + 210 + cmdCatalogOps(0).Width
            
            cmdCatalogOps(0).Top = jgxPicklist.Top + 315
            cmdCatalogOps(1).Top = cmdCatalogOps(0).Top + cmdCatalogOps(0).Height + 120
            cmdCatalogOps(2).Top = cmdCatalogOps(1).Top + cmdCatalogOps(1).Height + 120
            cmdCatalogOps(3).Top = cmdCatalogOps(2).Top + cmdCatalogOps(2).Height + 120
            
            cmdCatalogOps(0).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(1).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(2).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(3).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            
            cmdFilter.Left = cmdCatalogOps(0).Left + 240
            chkClearFilter.Left = cmdFilter.Left + 480
            
            cmdTransact(1).Left = tabCatalog.Width + tabCatalog.Left - 1215
            cmdTransact(0).Left = cmdTransact(1).Left - 100 - 1215
                    
            cmdTransact(0).Top = tabCatalog.Height + 180
            cmdTransact(1).Top = tabCatalog.Height + 180
            
            '<<< dandan 110707
            cmdSeeOnLine.Top = tabCatalog.Height + 180
            cmdSeeOnLine.Left = txtSearchFields(0).Left
            
            Me.Height = tabCatalog.Height + cmdTransact(0).Height + 270 + 375
            
    End Select

   Call CheckIfFilter(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)

End Sub

' Fan-Out: InitPickStyle
Private Sub CheckIfFilter(ByRef intColumnCtr As Integer, _
                ByRef dblLastRightPosition As Double, _
                ByRef dblDummyTop As Double, _
                ByRef dblTextSearchTopOffset _
                As Double)
    
    'Mod by BCo
    'Fixes minor positioning problem with Consignor/Consignee OK and Cancel cmd.
    'Me.Width = cmdTransact(1).Left + cmdTransact(1).Width + 145
    Me.Width = cmdTransact(1).Left + cmdTransact(1).Width + 230
    
    If ((clsFilter Is Nothing) = False) Then
    
        If ((enuStyle = cpiSimplePicklist) Or ((enuStyle <> cpiSimplePicklist) _
            And (clsFilter.FilterType <> cpiComboRecords))) Then
            
            ' load SQL command
            RefreshGrid cpiRequery, strRecordsList
            
        End If
        
    ElseIf ((clsFilter Is Nothing) = True) Then
    
        RefreshGrid cpiRequery, strRecordsList
        
    End If

End Sub

' Fan-Out: InitTextFields
Private Sub InitTabList(ByRef intColumnCtr As Integer, _
                                    ByRef dblLastRightPosition As Double, _
                                    ByRef dblDummyTop As Double, _
                                    ByRef dblTextSearchTopOffset _
                                    As Double)

   Select Case enuStyle
       
       Case cpiCatalog
           
           jgxPicklist.Top = tabCatalog.Top + lblListDescription.Top + lblListDescription.Height
           tabCatalog.Height = jgxPicklist.Top + jgxPicklist.Height
           Me.Caption = "Catalog - " & strPluralEntity
           
       Case cpiFilterCatalog
           
           jgxPicklist.Top = txtSearchFields(0).Top + txtSearchFields(0).Height
           tabCatalog.Height = jgxPicklist.Top + jgxPicklist.Height
           
           'IF statement commented by alg
           'If (clsPicklist.Columns.Count > 2) Then
               
               cmdFilter.Visible = True
               chkClearFilter.Visible = True
           
           'End If
           
           Me.Caption = "Search Catalog - " & strPluralEntity
           
   End Select

End Sub

' Fan-Out: InitGrid
Private Sub LoadColumnValue(ByRef intColumnCtr As Integer, _
                                                ByRef dblLastRightPosition As Double, _
                                                ByRef dblDummyTop As Double, _
                                                ByRef dblTextSearchTopOffset As Double)

    Select Case enuStyle
    
        Case cpiSimplePicklist
        
            Call ColumnSimplePick(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
            
        Case cpiCatalog
        
            dblLastRightPosition = dblLastRightPosition _
                                        + clsGridSeed.GridColumns.Item(intColumnCtr).ColumnWidth
        
        Case cpiFilterCatalog
        
            txtSearchFields(intColumnCtr - 1).Visible = True
        
            If (clsGridSeed.GridColumns.Count > 2) Then
            
                txtSearchFields(intColumnCtr - 1).Width = _
                clsGridSeed.GridColumns.Item(intColumnCtr).ColumnWidth
                
            End If
        
            txtSearchFields(intColumnCtr - 1).Left = 5 + dblLastRightPosition
        
            If (clsGridSeed.GridColumns.Count > 2) Then
            
                dblLastRightPosition = dblLastRightPosition + 10 _
                                        + clsGridSeed.GridColumns.Item(intColumnCtr).ColumnWidth
                                        
            ElseIf (clsGridSeed.GridColumns.Count <= 2) Then
                
                dblLastRightPosition = dblLastRightPosition + 10 _
                                        + txtSearchFields(intColumnCtr - 1).Width
            End If
    
    End Select

End Sub

' Fan-Out: LoadColumnValue
Private Sub ColumnSimplePick(ByRef intColumnCtr As Integer, _
                    ByRef dblLastRightPosition As Double, _
                    ByRef dblDummyTop As Double, _
                    ByRef dblTextSearchTopOffset As Double)

    txtSearchFields(intColumnCtr - 1).Visible = True
    txtSearchFields(intColumnCtr - 1).Left = dblLastRightPosition - 5
    
    If ((clsGridSeed.GridColumns.Count > 2) = True) Then
    
        txtSearchFields(intColumnCtr - 1).Width = _
            clsGridSeed.GridColumns.Item(intColumnCtr).ColumnWidth
        
        dblLastRightPosition = dblLastRightPosition + 10 _
                                    + clsGridSeed.GridColumns.Item(intColumnCtr).ColumnWidth
        
    ElseIf ((clsGridSeed.GridColumns.Count > 2) = False) Then
    
        dblLastRightPosition = dblLastRightPosition _
                                    + 10 + txtSearchFields(intColumnCtr - 1).Width
        
    End If

End Sub

' Fan-Out: Form_Unload
Private Sub ProcessRecord(ByRef clsReturnRecord As CRecord, ByRef Cancel As Integer)

    Dim intTransIndex As Integer
    Dim intChildTransIndex As Integer
    Dim varDBID As Variant
    
    If ((enuStyle = cpiCatalog) Or (enuStyle = cpiFilterCatalog)) Then
        
        If (blnCanceled = True) Then
        
            conDBConnection.RollbackTrans
        
        ElseIf (blnCanceled = False) Then
        
            ' reconcile primary keys
            If (blnDBAutoNumber = True) Then
                Call ReconcileTransPK
            End If
            
            CommitTransaction conDBConnection, clsPicklist.BaseSQL
            ' commiting the transaction
            conDBConnection.CommitTrans
            ' after commiting
            
            ' get the new db id from the database
            ' set clsRecord DBID into new ID from DB
            If (blnDBAutoNumber = True) Then
            
                If ((clsPicklist.Transactions Is Nothing) = False) Then
                
                    For intTransIndex = 1 To clsPicklist.Transactions.Count
                    
                        varDBID = GetDBID(conDBConnection _
                                                    , clsPicklist.BaseSQL _
                                                    , clsPicklist.PKFieldBaseName _
                                                    , clsRecord.TempDBID)
                    Next intTransIndex
                    
                End If
        
            End If
        
        End If
    
    End If

End Sub

' Fan-Out: RefreshGrid
Private Sub InitGridRecFilter(ByVal Repaint As RefreshType, _
                                            ByRef intSearchTextIndex As Integer, _
                                            ByRef lngVisibleColumnCtr As Long, _
                                            ByRef intColumnCtr As Integer, _
                                            ByRef strFilterToApply As String, _
                                            ByRef strCurrentFilter As String, _
                                            Optional ByVal Source As String, _
                                            Optional ByVal DisableFilter As Boolean = False, _
                                            Optional ByRef clsWhere As CFilterNodes)
    
    Dim strTempFilter As String
    
'    Select Case clsPicklist.PicklistStyle
    
'        Case cpiSimplePicklist, cpiCatalog
        
'            Set clsWhere = New cpiFilterNodes
'            clsWhere.Add "Tag <> 'D'"
'            clsWhere.RefreshFilter
'            strTempFilter = GetProperFilter(clsWhere)
'            strHeaderFilter = Trim$(strTempFilter)
'            strFilterToApply = strHeaderFilter
        
'        Case cpiFilterCatalog
        
            If ((clsFilter Is Nothing) = False) Then
            
                ' load proper filter here
                Call InitGridFilter(clsWhere)
                strTempFilter = GetProperFilter(clsWhere)
                strHeaderFilter = Trim$(strTempFilter)
                strFilterToApply = strHeaderFilter
            
            ElseIf ((clsFilter Is Nothing) = True) Then
                Set clsWhere = New CFilterNodes
                clsWhere.Add "Tag <> 'D'"
                clsWhere.RefreshFilter
                strTempFilter = GetProperFilter(clsWhere)
                strHeaderFilter = Trim$(strTempFilter)
                strFilterToApply = strHeaderFilter
            End If
    
'    End Select
   
End Sub

' Fan-Out: InitGridRecFilter
Private Function GetProperFilter(ByRef clsWhere As CFilterNodes) As String
   
   Dim intFilterCtr As Integer
   Dim strReturn As String
   
   For intFilterCtr = 1 To clsWhere.FilterCount
   
      strReturn = strReturn & " OR (" & clsWhere.FilterList(intFilterCtr) & ")"
      
   Next intFilterCtr
   
   strReturn = Right$(strReturn, Len(strReturn) - 4)
   
   GetProperFilter = strReturn

End Function


' Fan-Out: InitGridRecFilter
Private Sub InitGridFilter(ByRef clsWhere As CFilterNodes)

    Dim intFilterCtr As Integer
    
    Set clsWhere = New CFilterNodes
    
    ' first filter
    clsWhere.Add "Tag <> 'D'"
    
    ' check if filter object is enabled
    If ((clsFilter Is Nothing) = False) Then
    
        If (clsFilter.PicklistFilters.Count <> 0) Then
        
            For intFilterCtr = 1 To clsFilter.PicklistFilters.Count
            
                If (clsFilter.PicklistFilters(intFilterCtr).State = True) Then
                
                    If (Trim$(clsFilter.PicklistFilters(intFilterCtr).Filter) <> "") Then
                    
                        clsWhere.Add clsFilter.PicklistFilters(intFilterCtr).Filter
                        
                    End If
                    
                End If
                
            Next intFilterCtr
            
        End If
        
    End If
    
    If (strLastColumnFilter <> "") Then
    
        If (chkClearFilter.Value = vbChecked) Then
            clsWhere.Add strLastColumnFilter
        ElseIf (chkClearFilter.Value = vbUnchecked) Then
            ' do nothing
        End If
        
    End If
    
    clsWhere.RefreshFilter
    
End Sub



' Fan-Out: PopulateGrid
Private Sub CreateRstToGrid(ByVal Repaint As RefreshType, _
                    ByRef intSearchTextIndex As Integer, _
                    ByRef lngVisibleColumnCtr As Long, _
                    ByRef intColumnCtr As Integer, _
                    ByRef strFilterToApply As String, _
                    ByRef strCurrentFilter As String, _
                    Optional ByVal Source As String, _
                    Optional ByVal DisableFilter As Boolean = False)

    Dim blnFlag As Boolean
    Dim strFilter As String

    ' get the field base name
    clsPicklist.PKFieldBaseName = FetchGridRecords(Source, conDBConnection, _
                                                                    rstRecordsList, adOpenKeyset, _
                                                                    adLockOptimistic, strPKFieldAliasInSQL, clsGridSeed)
    
    If (blnHaltExecution = False) Then
    
        rstRecordsList.Filter = adFilterNone
        
        blnFlag = ((chkClearFilter.Value = vbUnchecked) Or (DisableFilter = True) Or (Trim$(strLastColumnFilter) = ""))
        
        If (blnFlag = True) Then
            
            rstRecordsList.Filter = strFilterToApply
        
        ElseIf (blnFlag = False) Then
            
            strFilter = Trim$(strLastColumnFilter)
            If (strFilter = "") Then
                strFilter = strFilterToApply
            ElseIf (strFilter <> "") Then
                strFilter = strLastColumnFilter & " AND " & strFilterToApply
            End If
            
            rstRecordsList.Filter = strFilter
        
        End If
        
    End If
    
End Sub

' users - RefreshGrid
Private Sub ApplyGridFilter(ByVal Repaint As RefreshType _
                    , ByRef intSearchTextIndex As Integer _
                    , ByRef lngVisibleColumnCtr As Long _
                    , ByRef intColumnCtr As Integer _
                    , ByRef strFilterToApply As String _
                    , ByRef strCurrentFilter As String _
                    , Optional ByVal Source As String _
                    , Optional ByVal DisableFilter As Boolean = False _
                    , Optional ByRef clsWhere As CFilterNodes)
    
    Dim enuFilterType As enuPicklistFilter
    Dim strRunFilter As String
    
    If (clsFilter Is Nothing) = True Then
        ' not initialized
        Exit Sub
    End If
   
    If ((rstRecordsList Is Nothing) = False) Then
        
        rstRecordsList.Filter = adFilterNone
        
        If (enuStyle = cpiFilterCatalog) Then
        
            'radio options
            If (clsFilter.FilterType = cpiCheckOptions) Then
                
                If (blnFormActivated = True) Then 'blnFormActivated Start
                    'hard-coded
                    If (blnTaskClick = False) Then 'blnTaskList start
                        
                        If Trim$(strCurrentFilter) <> "" Then
                            rstRecordsList.Filter = strFilterToApply & " AND " & "(" & strCurrentFilter & ") "
                        Else
                            rstRecordsList.Filter = strFilterToApply
                        End If
                    
                    ElseIf (blnTaskClick = True) Then 'blnTaskList start
                    
                        'refresh checkbox
                        If (blnIsCheck = False) Then      'blncheck start
                        
                            strFilterToApply = "Tag <> 'D' AND " & clsFilter.PicklistFilters(1).Filter
                            strFilterToApply = Replace(strFilterToApply, "True", "False", , , vbTextCompare)
                            
                            If Trim$(strCurrentFilter) <> "" Then
                                rstRecordsList.Filter = strFilterToApply & " AND " & "(" & strCurrentFilter & ") "
                            Else
                                rstRecordsList.Filter = strFilterToApply
                            End If
                            
                            blnTaskClick = False
                            
                        ElseIf (blnIsCheck = True) Then
                        
                            strFilterToApply = "Tag <> 'D' AND " & clsFilter.PicklistFilters(1).Filter
                            
                            If Trim$(strCurrentFilter) <> "" Then
                                rstRecordsList.Filter = strFilterToApply & " AND " & "(" & strCurrentFilter & ") "
                            Else
                                rstRecordsList.Filter = strFilterToApply
                            End If
                            
                            blnTaskClick = False
                            
                        End If         'blncheck end
                    
                    End If      'blnTaskList end
                    
                ElseIf (blnFormActivated = False) Then
                
                    'hard-coded
                    strFilterToApply = "Tag <> 'D' AND " & clsFilter.PicklistFilters(1).Filter
                    strFilterToApply = Replace(strFilterToApply, "True", "False", , , vbTextCompare)
                    
                    If Trim$(strCurrentFilter) <> "" Then
                        rstRecordsList.Filter = strFilterToApply & " AND " & "(" & strCurrentFilter & ") "
                    Else
                        rstRecordsList.Filter = strFilterToApply
                    End If
                    
                End If ' check options end
            
            ElseIf (clsFilter.FilterType = cpiRadioOptions) Then  ' Radio options
            
                If (chkClearFilter.Value = vbChecked) Then
                
                    strRunFilter = strFilterToApply
                    rstRecordsList.Filter = ""
                    If Trim$(strCurrentFilter) <> "" Then
                        rstRecordsList.Filter = strRunFilter & " AND " & "(" & strCurrentFilter & ") "
                    Else
                        rstRecordsList.Filter = strRunFilter
                    End If
                
                ElseIf (chkClearFilter.Value = vbUnchecked) Then
                
                    strRunFilter = clsWhere.FilterGroup
                    If Trim$(strCurrentFilter) <> "" Then
                        rstRecordsList.Filter = strRunFilter & " AND " & "(" & strCurrentFilter & ") "
                    Else
                        rstRecordsList.Filter = strRunFilter
                    End If
                    
                End If
                
            End If  ' radio options end
                
        ElseIf (enuStyle = cpiCatalog) Then  ' cpiCatalog start
                
                If (blnFormActivated = False) Then ' patches
                    strHeaderFilter = "Tag <> 'D' AND " & clsFilter.PicklistFilters(1).Filter
                    strHeaderFilter = Replace(strHeaderFilter, "True", "False", , , vbTextCompare)
                End If
                
                If Trim$(strCurrentFilter) <> "" Then
                    rstRecordsList.Filter = strHeaderFilter & " AND " & "(" & strCurrentFilter & ") "
                Else
                    rstRecordsList.Filter = strHeaderFilter
                End If
                
        End If
            
        ' apply sort on grid recordset
        If (Trim$(strLastColumnSort) <> "") Then
            rstRecordsList.Sort = strLastColumnSort
        End If
            
    End If  'clsfilter end
   
End Sub

' Fan-Out: RefreshGrid
Private Sub SetGridColumns(ByVal Repaint As RefreshType, _
                    ByRef intSearchTextIndex As Integer, _
                    ByRef lngVisibleColumnCtr As Long, _
                    ByRef intColumnCtr As Integer, _
                    ByRef strFilterToApply As String, _
                    ByRef strCurrentFilter As String, _
                    Optional ByVal Source As String, _
                    Optional ByVal DisableFilter As Boolean = False)

    Dim blnExist As Boolean

    For intColumnCtr = 1 To jgxPicklist.Columns.Count
    
        ' check if the char exist
        blnExist = (InStr(1, strDummySettings, "||" & jgxPicklist.Columns(intColumnCtr).Caption & "||") <= 0)
    
        If (blnExist = True) Then
        
            jgxPicklist.Columns(intColumnCtr).Visible = False
            
        ElseIf (blnExist = False) Then
        
            If (clsGridSeed.GridColumns.Count > 2) Then
            
                jgxPicklist.Columns(jgxPicklist.Columns(intColumnCtr).Caption).Width = _
                clsGridSeed.GridColumns.Item(jgxPicklist.Columns(intColumnCtr).Caption).ColumnWidth
                
            ElseIf (clsGridSeed.GridColumns.Count <= 2) Then
            
                Select Case lngVisibleColumnCtr
                
                    Case 0
                        jgxPicklist.Columns(jgxPicklist.Columns(intColumnCtr).Caption).Width = 1300
                    Case 1
                        jgxPicklist.Columns(jgxPicklist.Columns(intColumnCtr).Caption).Width = 3000
                        
                End Select
            
                lngVisibleColumnCtr = lngVisibleColumnCtr + 1
                
            End If
            
            ' match tags of search textboxes with respective grid columns
            txtSearchFields(intSearchTextIndex).Tag = jgxPicklist.Columns(intColumnCtr).Caption
            intSearchTextIndex = intSearchTextIndex + 1
            
        End If
        
    Next intColumnCtr

    ' invisible all columns
    Call SetGridVisible(False)
    
    ' set proper visible grid
    For intColumnCtr = 1 To clsPicklist.Columns.Count
    
        jgxPicklist.Columns(clsPicklist.Columns(intColumnCtr).ColumnFieldAias).Visible = True
        txtSearchFields(intColumnCtr - 1).Tag = clsPicklist.Columns(intColumnCtr).ColumnFieldAias
        
    Next intColumnCtr

End Sub

' Fan-Out: RefreshGrid
Private Sub PopulateGrid(ByVal Repaint As RefreshType _
                                                   , ByRef intSearchTextIndex As Integer _
                                                   , ByRef lngVisibleColumnCtr As Long _
                                                   , ByRef intColumnCtr As Integer _
                                                   , ByRef strFilterToApply As String _
                                                   , ByRef strCurrentFilter As String _
                                                   , Optional ByVal Source As String _
                                                   , Optional ByVal DisableFilter As Boolean = False)
    
    Select Case Repaint
    
        Case cpiRequery
        
            'create recordset to populate the grid with
            Call CreateRstToGrid(Repaint, intSearchTextIndex, lngVisibleColumnCtr, intColumnCtr, _
                strFilterToApply, strCurrentFilter, Source, DisableFilter)
            Case cpiRefresh
        
    End Select
    
End Sub

' users - InitTextFields
Private Sub ApplyTextFields(ByRef intColumnCtr As Integer _
                , ByRef dblLastRightPosition As Double _
                , ByRef dblDummyTop As Double _
                , ByRef dblTextSearchTopOffset As Double)
   
    Dim strCaption As String
    
    tabCatalog.Visible = True
    tabCatalog.Caption = "&" & strPluralEntity
    
    strCaption = clsPicklist.Caption
    strCaption = Replace(strCaption, "&", "&&", , , vbTextCompare)
    tabCatalog.Caption = strCaption
    
    lblListDescription.Caption = "List of " & strCaption
    
    dblTextSearchTopOffset = lblListDescription.Top + lblListDescription.Height + 45
    
    txtSearchFields(0).Top = dblTextSearchTopOffset
    txtSearchFields(1).Top = dblTextSearchTopOffset
    txtSearchFields(2).Top = dblTextSearchTopOffset
    txtSearchFields(3).Top = dblTextSearchTopOffset
    txtSearchFields(4).Top = dblTextSearchTopOffset
    txtSearchFields(5).Top = dblTextSearchTopOffset
    txtSearchFields(6).Top = dblTextSearchTopOffset
    txtSearchFields(7).Top = dblTextSearchTopOffset
    txtSearchFields(8).Top = dblTextSearchTopOffset
    
    cmdFilter.Top = dblTextSearchTopOffset
    chkClearFilter.Top = dblTextSearchTopOffset
    
    txtSearchFields(0).Visible = False
    txtSearchFields(1).Visible = False
    txtSearchFields(2).Visible = False
    txtSearchFields(3).Visible = False
    txtSearchFields(4).Visible = False
    txtSearchFields(5).Visible = False
    txtSearchFields(6).Visible = False
    txtSearchFields(7).Visible = False
    txtSearchFields(8).Visible = False
    
    jgxPicklist.Left = tabCatalog.Left + 90

End Sub

' users - FetchGridRecords
Private Sub ApplyFetchGrid(ByRef Source As String _
                    , ByRef conToUse As ADODB.Connection _
                    , ByRef rstToOpen As ADODB.Recordset _
                    , ByRef CursorType As CursorTypeEnum _
                    , ByRef LockType As LockTypeEnum _
                    , ByRef PKFieldName As String _
                    , ByRef GridSeed As CGridSeed _
                    , ByRef rstDummy As ADODB.Recordset _
                    , ByRef fldDummy As ADODB.Field _
                    , ByRef strSourceLeftOfFrom As String _
                    , ByRef strSourceRightOfFrom As String _
                    , ByRef strOriginalSource As String _
                    , ByRef strFinalSource As String _
                    , ByRef intFieldIndex As Integer)

    ' intFieldIndex is counter for rstToOpen recordset's fields

    Dim varFields() As Variant
    Dim varValues() As Variant
    
    ' apply fetch grid
    Set rstToOpen = New ADODB.Recordset
    
    ' add fields to rstToOpen
    ReDim varFields(rstDummy.Fields.Count)
    ReDim varValues(rstDummy.Fields.Count)
    
    For intFieldIndex = 0 To rstDummy.Fields.Count - 1
    
        ' get fieldname
        varFields(intFieldIndex) = rstDummy.Fields(intFieldIndex).Name
        rstToOpen.Fields.Append rstDummy.Fields(intFieldIndex).Name, rstDummy.Fields(intFieldIndex).Type, _
        rstDummy.Fields(intFieldIndex).DefinedSize, rstDummy.Fields(intFieldIndex).Attributes
    
    Next intFieldIndex
    
    ' get the last intFieldIndex
    varFields(intFieldIndex) = "Tag"
    varValues(intFieldIndex) = "O"
    
    ' original - add one unbound field name tag
    rstToOpen.Fields.Append "Tag", adVarWChar, 1, adFldIsNullable
    ' open rst as disconnected recordset
    rstToOpen.Open
    
    ' alternative method
    While Not rstDummy.EOF
        For intFieldIndex = 0 To rstDummy.Fields.Count - 1
            varValues(intFieldIndex) = rstDummy.Fields(intFieldIndex).Value
        Next intFieldIndex
        rstToOpen.AddNew varFields, varValues
        rstDummy.MoveNext
    Wend
    
    ' global autonumber
    blnDBAutoNumber = rstDummy.Fields(PKFieldName).Properties("ISAUTOINCREMENT").Value
    'critical flag
    blnInitEnd = True

End Sub

' Fan-Out: GetAutoNoInDB,cmdCatalogOps_Click
Private Function GetMaxAutoInTrans(ByVal clsiTrans As CTransactions, _
                                                            ByVal lngiTempKey As Long, _
                                                            ByVal striPKField As String) As Long
   
   Dim lngMaxIndex  As Long
      
   Dim clsRecordCtr As CRecord
   
   lngMaxIndex = lngiTempKey
   
   For Each clsRecordCtr In clsiTrans
   
      If clsRecordCtr.RecordSource.Fields(striPKField).Value > lngMaxIndex Then
         lngMaxIndex = clsRecordCtr.RecordSource.Fields(striPKField).Value
      End If
      
   Next  'clsRecordCtr
   
   GetMaxAutoInTrans = lngMaxIndex 'return the value
   
   Set clsRecordCtr = Nothing
   
End Function

' Fan-Out: cmdCatalogOps_Click
Private Sub SetOldRecordPos(ByRef rstiRecord As ADODB.Recordset, _
                    ByRef blniLoadAllrecord As Boolean, _
                    ByRef lngioStart As Long, _
                    ByRef lngioEnd As Long)
    
    ' load current record only
    If (blniLoadAllrecord = False) Then
        ' if not all loaded then load the selected record
        lngioStart = rstiRecord.AbsolutePosition - 1
        lngioEnd = lngioStart
    ElseIf (blniLoadAllrecord = True) Then
        ' load all records
        lngioStart = 0
        lngioEnd = rstiRecord.RecordCount - 1
    End If
   
End Sub

' Fan-Out: GetAutoNoInDB,GetDBID,cmdCatalogOps_Click
Private Function GetMaxAutoInDB(ByRef conConnection As ADODB.Connection, _
                                                        ByVal striTable As String, _
                                                        ByVal striPKField As String) As Long

   Dim rstTemp As ADODB.Recordset
   Dim strRunSQL As String
   
   Set rstTemp = New ADODB.Recordset
   
   ' command to get the maximum autonumber in the primary key field
   strRunSQL = "SELECT  Max(" & striPKField & ") FROM " & IIf(InStr(1, striTable, "("), Replace(striTable, "(", vbNullString), striTable) 'allanent nov7
   
   On Error GoTo ERROR_MSG
   
   ADORecordsetOpen strRunSQL, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
   'rstTemp.Open strRunSQL, conConnection    ' run query
   
   GetMaxAutoInDB = IIf(IsNull(rstTemp.Fields(0).Value), 0, rstTemp.Fields(0).Value)
   
   ADORecordsetClose rstTemp
   
   Exit Function
   
ERROR_MSG:

   MsgBox Err.Description, vbInformation, frmOwnerForm.Caption
   Resume Next
   
End Function

' Fan-Out: GetAutoNoInDB,GetDBID,cmdCatalogOps_Click
Private Function GetTableName(ByVal striSQL As String) As String
   
    Dim strTable As String
    Dim strSplitTable() As String
    Dim intOrderLoc As Integer
    
    strSplitTable = Split(striSQL, " FROM ", , vbTextCompare)
    
    'mark 12092002 for positioning and ordering of picklist, to be removed if there is already a solution
    intOrderLoc = InStr(1, UCase$(strSplitTable(1)), "ORDER BY")
    
    If (intOrderLoc > 0) Then
        intOrderLoc = intOrderLoc - 1
    ElseIf (intOrderLoc = 0) Then
        intOrderLoc = Len(strSplitTable(1))
    End If
    
    strTable = Mid$(strSplitTable(1), 1, intOrderLoc)
    '-------------------------------------------------------------------------------------------------------------------
    
    strSplitTable = Split(strTable, " WHERE ", , vbTextCompare)
    strTable = strSplitTable(0)
    
    ' check if inner join exist
    If (InStr(1, strTable, " INNER JOIN ", vbTextCompare) <> 0) Then
    
    ' count the no of inner join, total tables=total no. of inner join + 1
    strTable = GetFirstTable(" INNER JOIN ", strTable)
    
    ' check if the table is matrix ,
    ElseIf (InStr(1, strTable, ",", vbTextCompare) <> 0) Then
    
        strTable = GetFirstTable(",", strTable)
        
    ElseIf (InStr(1, strTable, " LEFT JOIN ", vbTextCompare) <> 0) Then
    
        strTable = GetFirstTable(" LEFT JOIN ", strTable)
        
    ElseIf (InStr(1, strTable, " RIGHT JOIN ", vbTextCompare) <> 0) Then
    
        strTable = GetFirstTable(" RIGHT JOIN ", strTable)
        
    End If
    
    ' return the first table
    While Left(strTable, 1) = "("
        strTable = Mid(strTable, 2)
    Wend
    GetTableName = Trim$(strTable)
    
    
End Function

' Fan-Out: ProcessRecord
Private Function GetFirstTable(ByRef strSearch As String, ByRef strFrom As String) As String
   
    Dim strTable As String
    Dim strSplitTable() As String
    Dim clsString As CStringExtension
    
    Set clsString = New CStringExtension
    clsString.GetStringPosition strSearch, strFrom
    strSplitTable = Split(strFrom, strSearch, , vbTextCompare)
    
    ' return the first table
    GetFirstTable = Trim$(strSplitTable(0))
    
    Set clsString = Nothing

End Function

' Fan-Out: InitAddTrans,TransactAddCopy
Private Function GetAutoNoInDB(ByRef clsiTrans As CTransactions, _
                                ByRef coniDBConnection As ADODB.Connection, _
                                ByRef striSQL As String, _
                                ByRef striPKField As String) As Long
   
   Dim lngkey As Long
   Dim lngTempKey As Long
   Dim strTable  As String
   Dim blnGenerateNew As Boolean
       
   ' get the table name from the SQL statement
   strTable = GetTableName(striSQL)  'ERROR HERE WHEN INNER JOIN
   ' get the highest Autonumber from the database
   lngTempKey = GetMaxAutoInDB(coniDBConnection, strTable, striPKField)
   ' check if there is a added/deleted/modified records in Transaction collection
   ' that has a higher autonumber in lngTempKey
   lngTempKey = GetMaxAutoInTrans(clsiTrans, lngTempKey, striPKField)
   ' get index from transaction collection
   lngkey = lngTempKey + 1
   
   Do While blnGenerateNew = False
        clsPicklist.CheckID blnGenerateNew, lngkey
        If blnGenerateNew = False Then
            Exit Do
        End If
        lngkey = lngkey + 1
   Loop
   
   GetAutoNoInDB = lngkey  ' return
   
End Function

' Fan-Out: ShowPicklist
Private Sub FormInitialize()
   
    Dim intColumnCtr As Integer
    
    Dim dblLastRightPosition As Double
    Dim dblDummyTop As Double
    Dim dblTextSearchTopOffset As Double
    
    Screen.MousePointer = vbHourglass
    
    blnCanceled = True
    blnFormActivated = False
    blnChangeDueToGridClick = False
    blnGridIsEmpty = True
    strHeaderFilter = ""
    strLastColumnFilter = ""
    strNewColumnFilter = ""
    strLastColumnSort = ""
    
    dblTextSearchTopOffset = lblListDescription.Top
    
    ' initialise recordset opener
    Set clsRecordset = New CRecordset
    
    ' check pickstyle
    Call CheckPicklistStyle(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    
    ' initialize textfields
    Call InitTextFields(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    
    ' initialize grid using GridSeed
    Call InitGrid(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    
    ' populate datacombo boxes
    Call InitDataCombo(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    Call InitDataComboExt(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    
    ' initialize pickstyle
    Call InitPickStyle(intColumnCtr, dblLastRightPosition, dblDummyTop, dblTextSearchTopOffset)
    
    ' get windopws metrics on
    ' registry directory HKEY_CURRERTUSER\Control Panel\Desktop\Windows Metrics
    ' the following: BorderWidth, CaptionHeight, MenuHeight
    Call ResetFormHeight
    
    ' initialize controls
    Call InitControls
    
    Screen.MousePointer = vbDefault
    
End Sub

' Fan-Out: FormInitialize
Private Sub InitControls()

    ' init user defined buttons
    Call InitButtons
    Call InitButton
    
    ' initialize ok button
    If (clsPicklist.AutoUnload = CPI_AUTOCANCEL) Then
        cmdTransact(CMD_OK).Caption = "&Select"
    ElseIf (clsPicklist.AutoUnload = CPI_TRUE) Then
        cmdTransact(CMD_OK).Caption = "&Select"
        cmdTransact(CMD_CANCEL).Caption = "&Close"
        
    End If
    
    ' patches
    If (rstRecordsList.RecordCount = 0) Then
        cmdTransact(CMD_OK).Enabled = False
    ElseIf (rstRecordsList.RecordCount > 0) Then
        cmdTransact(CMD_OK).Enabled = True
    End If

End Sub

' Fan-Out: InitControls
Private Sub InitButtons()
         
   ' interface button settings
   Select Case enuStyle
      
      Case cpiSimplePicklist
         cmdCatalogOps(CMD_ADD).Visible = False
         cmdCatalogOps(CMD_MODIFY).Visible = False
         cmdCatalogOps(CMD_COPY).Visible = False
         cmdCatalogOps(CMD_DELETE).Visible = False
      
      Case cpiCatalog
         conDBConnection.BeginTrans
         cmdCatalogOps(CMD_ADD).Enabled = clsPicklist.AddButton
         cmdCatalogOps(CMD_DELETE).Enabled = clsPicklist.DeleteButton
         cmdCatalogOps(CMD_MODIFY).Enabled = clsPicklist.ModifyButton
         cmdCatalogOps(CMD_COPY).Enabled = clsPicklist.CopyButton
         
         cmdCatalogOps(CMD_ADD).Visible = clsPicklist.AddButtonVisible
         cmdCatalogOps(CMD_DELETE).Visible = clsPicklist.DeleteButtonVisible
         cmdCatalogOps(CMD_MODIFY).Visible = clsPicklist.ModifyButtonVisible
         cmdCatalogOps(CMD_COPY).Visible = clsPicklist.CopyButtonVisible
         
      
      Case cpiFilterCatalog
         conDBConnection.BeginTrans
         cmdCatalogOps(CMD_ADD).Enabled = clsPicklist.AddButton
         cmdCatalogOps(CMD_DELETE).Enabled = clsPicklist.DeleteButton
         cmdCatalogOps(CMD_MODIFY).Enabled = clsPicklist.ModifyButton
         cmdCatalogOps(CMD_COPY).Enabled = clsPicklist.CopyButton

         cmdCatalogOps(CMD_ADD).Visible = clsPicklist.AddButtonVisible
         cmdCatalogOps(CMD_DELETE).Visible = clsPicklist.DeleteButtonVisible
         cmdCatalogOps(CMD_MODIFY).Visible = clsPicklist.ModifyButtonVisible
         cmdCatalogOps(CMD_COPY).Visible = clsPicklist.CopyButtonVisible
          
          'Tooltip
         cmdFilter.ToolTipText = "Filter By Selection"
         chkClearFilter.ToolTipText = "Apply Filter"

   End Select
    
    
    'initialize ok button
   If (enuStyle = cpiFilterCatalog) Then
   
      cmdTransact(CMD_OK).Default = False
      
   End If
   
   'Cancel
   cmdTransact(CMD_CANCEL).Cancel = True
    
    cmdSeeOnLine.ZOrder 0
    If Len(Trim(mvarWebLink)) > 0 Then
        cmdSeeOnLine.Visible = True
        cmdSeeOnLine.ToolTipText = mvarWebLink
    Else
        cmdSeeOnLine.Visible = False
    End If
    
End Sub



' Fan-Out: InitAddValues
Private Sub InitAddValues(ByRef rstiBlank As ADODB.Recordset)
    
    Dim fldDummy As ADODB.Field

    If (rstiBlank.RecordCount <> 0) Then
        rstiBlank.MoveFirst

       For Each fldDummy In rstiBlank.Fields
   
      Select Case fldDummy.Type
         Case 2      ' Integer
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 3      ' Long
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 4      ' Single
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 5      ' Double
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 6      ' Currency
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 7      ' Date
            ' Do nothing
         Case 11     ' Boolean
            ' Do nothing
         Case 17     ' Byte
            rstiBlank.Fields(fldDummy.Name) = 0
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = 0
         Case 202    ' Text
            rstiBlank.Fields(fldDummy.Name) = ""
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = ""
         Case 203    ' Memo
            rstiBlank.Fields(fldDummy.Name) = ""
            'mvarADOFields.Item(fldDummy.Properties(0).Value).Value = ""
         Case 205    ' OLE Object
            ' Do nothing
         Case Else
            ' Do nothing
      End Select
      
    Next ' fldDummy

    Set fldDummy = Nothing
   
    End If

End Sub




' users -cmdCatalogOps_Click
Private Function blnCancelDeleteOp(ByVal blniOperation As RecordOperation) As Boolean
   
    Dim varAns As Variant
    
    If (blniOperation = cpiRecordDelete) Then
    
        Screen.MousePointer = vbDefault
        
        varAns = MsgBox("Are you sure you want to delete this record? " _
                    , vbQuestion + vbYesNo, Me.Caption & " - Delete Record")
        
        If (varAns = vbNo) Then
            blnCancelDeleteOp = True
        ElseIf (varAns <> vbNo) Then
            Screen.MousePointer = vbHourglass
        End If
    
    End If

End Function

' Fan-Out: cmdCatalogOps_Click
Private Function InitAddTrans(ByRef clsRecord As CRecord, _
                                                    ByRef clsiPick As CPicklist, _
                                                    ByRef striPKAliasName As String) _
                                                    As String
   
    ' Public - blnDBAutoNumber - check if autonumber
    ' Public - lngTempAutoNoCtr - transaction counter if not autonumber
    
    Dim lngNewPKValue As Long
    
    If (blnDBAutoNumber = True) Then
    
        If ((clsPicklist.Transactions Is Nothing) = False) Then
        
            lngNewPKValue = GetAutoNoInDB(clsiPick.Transactions, conDBConnection, _
                                            clsiPick.BaseSQL, clsiPick.PKFieldBaseName)
        End If
        
    ElseIf (blnDBAutoNumber = False) Then
        
        lngTempAutoNoCtr = lngTempAutoNoCtr + 1 ' modular
        lngNewPKValue = lngTempAutoNoCtr
    
    End If
        
    InitAddTrans = "S" & CStr(lngNewPKValue)
        
    If (blnDBAutoNumber = True) Then
        
        clsRecord.RecordSource.Fields(striPKAliasName).Value = lngNewPKValue                      ' ID generic -pampagulo
        clsRecord.RecordSource.Fields(clsiPick.PKFieldBaseName).Value = lngNewPKValue    'rep_id
        varTempDBID = lngNewPKValue
    
    End If
   
End Function

' Fan-Out: cmdCatalogOps_Click
Private Sub RunAddTrans(ByRef clsRecord As CRecord, _
                    ByRef clsiTrans As CPicklist, _
                    ByRef striKey As String, _
                    ByRef blnioZeroRec _
                    As Boolean)
      
    Dim fldDummy As ADODB.Field
    
    ' update transaction counter
    clsRecord.RecordSource.Fields(clsPicklist.PKFieldAlias).Value = _
                    clsRecord.RecordSource.Fields(clsPicklist.PKFieldBaseName).Value
    
    ' add transaction to transaction collection
    clsPicklist.Transactions.Add clsRecord.ADOFields, clsRecord.RecordSource, _
                    clsRecord.RecordSQL, cpiStateNew, _
                    clsiTrans.TransactionCtr, _
                    clsRecord.OldRecordSource
    
    ' add record to grid recordset
    rstRecordsList.AddNew
    
    For Each fldDummy In rstRecordsList.Fields
    
        If (UCase$(fldDummy.Name) <> "TAG") Then
        
            rstRecordsList.Fields(fldDummy.Name).Value = _
                                clsRecord.RecordSource.Fields(fldDummy.Name).Value
            
        End If
        
    Next 'fldDummy
    
    
    ' set state to 'A' = add state
    rstRecordsList!Tag = "A"
    
    ' update grid
    rstRecordsList.Update
    
    If (rstRecordsList.RecordCount > 0) Then
    
        cmdCatalogOps(CMD_ADD).Enabled = True
        cmdCatalogOps(CMD_DELETE).Enabled = True
        cmdCatalogOps(CMD_MODIFY).Enabled = True
        cmdCatalogOps(CMD_COPY).Enabled = True
        
    End If
   
   Set fldDummy = Nothing
   
End Sub

' Fan-Out: cmdCatalogOps_Click
Private Sub RunEditTrans(ByRef clsRecord As CRecord, _
                    ByRef striSelectedTag As String, _
                    ByRef striKey As String)
   
    Dim intIndex As Integer
    
    If (striSelectedTag = "O") Then
    
        clsPicklist.TransactionCtr = clsPicklist.TransactionCtr + 1
        
        ' set state into modified
        clsPicklist.Transactions.Add clsRecord.ADOFields, clsRecord.RecordSource _
                                                        , clsRecord.RecordSQL, cpiStateModified _
                                                        , striKey, clsRecord.OldRecordSource
        rstRecordsList!Tag = "M"
        
    ElseIf (striSelectedTag = "A") Then
    
        intIndex = GetTransIndex(clsPicklist.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
        
        If (intIndex <> 0) Then
        
            Set clsPicklist.Transactions(intIndex).RecordSource = clsRecord.RecordSource
            
        End If
        
    ElseIf (striSelectedTag = "M") Then
    
        intIndex = GetTransIndex(clsPicklist.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
        ' change record source to transaction collection
        Set clsPicklist.Transactions(intIndex).RecordSource = clsRecord.RecordSource
        
    End If
    
    ' refresh grid with data modified
    Call UpdateGrid(clsRecord)

End Sub

' Fan-Out: cmdCatalogOps_Click
Private Sub RunDeleteTrans(ByRef rstiRecord As ADODB.Recordset, _
                    ByRef clsRecord As CRecord, _
                    ByRef clsiTrans As CPicklist, _
                    ByRef striKey _
                    As String)
                                                       
    Dim lngIndex As Long
    Dim blnExist As Boolean
                                
    If (UCase$(rstiRecord!Tag) = "A") Then
    
        lngIndex = GetTransIndex(clsiTrans.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
        clsiTrans.Transactions.Remove lngIndex
        
        ' to be upadted  for multiuser usage
        blnIsInTrans = True
        
        rstiRecord!Tag = "D"
        rstRecordsList.MoveFirst
    
    ' from DB
    ElseIf (UCase$(rstiRecord!Tag) = "O") Then
    
        ' change tag into 'D'
        rstiRecord!Tag = "D"
        clsiTrans.Transactions.Add clsRecord.ADOFields, clsRecord.RecordSource _
                                        , clsRecord.RecordSQL, cpiStateDeleted _
                                         , striKey, clsRecord.OldRecordSource
                                         
    ' already in transaction
    ElseIf (rstiRecord!Tag = "M") Then
        
        ' if in db then check if in db
        blnExist = blnPKinDB(varTempDBID)
        
        If (blnExist = True) Then
        
            rstiRecord!Tag = "D"  ' change tag into 'D'
            
            ' change state from modified into deleted
            lngIndex = GetTransIndex(clsiTrans.Transactions, clsRecord, clsPicklist.PKFieldBaseName)
            
            If (striKey <> clsiTrans.Transactions(lngIndex).Key) Then
            
                clsiTrans.Transactions(lngIndex).Status = cpiStateDeleted
                
            ElseIf (striKey = clsiTrans.Transactions(lngIndex).Key) Then
            
                clsiTrans.Transactions(striKey).Status = cpiStateDeleted
                
            End If
            
        'else in transaction only
        ElseIf (blnExist = False) Then
        
            clsiTrans.Transactions.Remove striKey
            rstiRecord.Delete
            
        End If
    
    End If
    
    If (blnIsInTrans = False) Then
    
        If ((rstiRecord.AbsolutePosition = adPosUnknown) And (rstiRecord.RecordCount > 0)) Then
            rstiRecord.MoveFirst
        End If
        
    End If
    
    If (rstRecordsList.RecordCount = 0) Then
    
        cmdCatalogOps(CMD_MODIFY).Enabled = False
        cmdCatalogOps(CMD_COPY).Enabled = False
        cmdCatalogOps(CMD_DELETE).Enabled = False
        
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

' Fan-Out: cmdCatalogOps_Click,RunEditTrans,RunDeleteTrans
Private Function GetTransIndex(ByRef clsiTrans As CTransactions, _
                                                        ByRef clsRecord As CRecord, _
                                                        ByRef striPKValue As String) As Long

    Dim lngIndex As Long
    
    If (ErrorPatch(clsRecord.RecordSource) = True) Then
    
        Exit Function
        
    End If
    
    For lngIndex = 1 To clsiTrans.Count
    
        If clsiTrans(lngIndex).RecordSource(striPKValue).Value = _
            IIf(IsNull(clsRecord.RecordSource(striPKValue).Value), "", clsRecord.RecordSource(striPKValue)) Then
        
            GetTransIndex = lngIndex
            Exit Function
            
        End If
        
    Next lngIndex
    
End Function


' Fan-Out: GetSelectedRec,blnPKinDB,cmdCatalogOps_Click
Private Function GetTopOne(ByRef conConnection As ADODB.Connection, _
                                                ByRef striSQL As String, _
                                                ByRef blniNoTop As Boolean, _
                                                ByVal striPKField As String, _
                                                ByVal striValue) _
                                                As ADODB.Recordset
   
    Dim strTempSQL As String
    Dim rstTemp As ADODB.Recordset
    Dim strOrderBy As String
    Dim strWhere As String
    Dim varTemp As Variant
    
    Set rstTemp = New ADODB.Recordset
    
'    If blnOperator Then
'        Set GetTopOne = New ADODB.Recordset 'allan pick
'
'        GetTopOne.Open clsPicklist.BaseSQL & " AND [OPERATOR_VENTURENUMBER]= '" & jgxPicklist.Value(5) & "'", conConnection, adOpenKeyset, adLockOptimistic 'allan feb22
'    Else
        strTempSQL = striSQL
        
        ' split to get order by
        varTemp = Split(UCase$(strTempSQL), " ORDER BY ", , vbTextCompare)
        
        If (UBound(varTemp) <> 0) Then
        
            strOrderBy = varTemp(1)
            strTempSQL = varTemp(0)
            
        ElseIf (UBound(varTemp) = 0) Then
        
            strOrderBy = ""
            
        End If
        
        ' split to get the WHERE
        varTemp = Split(UCase$(strTempSQL), " WHERE ", , vbTextCompare)
        
        If (UBound(varTemp) <> 0) Then
        
            strWhere = varTemp(1)
            strTempSQL = varTemp(0)
            
        ElseIf (UBound(varTemp) = 0) Then
        
            strWhere = ""
            
        End If
        
        ' get the *
        strTempSQL = RegenerateSQL(strTempSQL, blniNoTop)
        
        ' set the where
        If (blnDBAutoNumber = True) Then
            If Len(Trim(clsPicklist.GetTopWhere)) = 0 Then
                strTempSQL = strTempSQL & " WHERE [" & striPKField & "] = " & striValue
            Else
                strTempSQL = strTempSQL & " WHERE [" & clsPicklist.GetTopWhere & "] = " & striValue
            
            End If
            
        ElseIf (blnDBAutoNumber = False) Then
            If Len(Trim(clsPicklist.GetTopWhere)) = 0 Then
                strTempSQL = strTempSQL & " WHERE [" & striPKField & "] = '" & striValue & "'"
            Else
                strTempSQL = strTempSQL & " WHERE [" & clsPicklist.GetTopWhere & "] = '" & striValue & "'"
               
            End If
        End If
        
        ' attached the other wheres
        If strWhere <> "" Then strTempSQL = strTempSQL & " AND " & strWhere
        
        ' attached the order by
        If strOrderBy <> "" Then strTempSQL = strTempSQL & " ORDER BY " & strOrderBy
        
        rstTemp.CursorLocation = adUseClient
        
        On Error GoTo ERROR_MSG:
        ADORecordsetOpen strTempSQL, conConnection, rstTemp, adOpenKeyset, adLockOptimistic
        'rstTemp.Open strTempSQL, conConnection, adOpenKeyset, adLockOptimistic
        
        Set GetTopOne = rstTemp
        
        ADORecordsetClose rstTemp
        'Set rstTemp = Nothing
        
        Exit Function
        
ERROR_MSG:
        
        MsgBox Err.Description, vbInformation, frmOwnerForm.Caption
        Resume Next
    'End If 'allan pick
    
End Function

' Fan-Out: ProcessRecord
Private Function GetDBID(ByRef conConnection As ADODB.Connection, _
                            ByRef striSQL As String, _
                            ByRef striPKField As String, _
                            ByRef tempID) As Variant

    Dim rstDummy As ADODB.Recordset
    Dim strRunSQL As String
    
    Set rstDummy = New ADODB.Recordset
    
    If (tempID <> Empty) Then
    
        strRunSQL = AppendWhere(striSQL, IIf(clsPicklist.GetTopWhere <> "", clsPicklist.GetTopWhere, striPKField) & " = " & tempID)
    
    End If
    
    On Error GoTo ERROR_MSG
    ' initial connection
    
    ADORecordsetOpen strRunSQL, conConnection, rstDummy, adOpenKeyset, adLockOptimistic
    'rstDummy.Open strRunSQL, conConnection, adOpenKeyset, adLockOptimistic
    
    ' check if same record
    If (rstDummy.RecordCount = 0) Then
    
        tempID = GetMaxAutoInDB(conConnection, GetTableName(striSQL), striPKField)
        strRunSQL = AppendWhere(striSQL, IIf(clsPicklist.GetTopWhere <> "", clsPicklist.GetTopWhere, striPKField) & " = " & tempID)
        
        ADORecordsetOpen strRunSQL, conConnection, rstDummy, adOpenKeyset, adLockOptimistic
        'rstDummy.Close
        'rstDummy.Open strRunSQL, conConnection, adOpenKeyset, adLockOptimistic
        
    End If
    
    If (rstDummy.RecordCount <> 0) Then
    
        GetDBID = rstDummy.Fields(striPKField).Value
    
    End If
    
    ADORecordsetClose rstDummy
    'Set rstDummy = Nothing
    
    Exit Function
    
ERROR_MSG::
    
    MsgBox Err.Description, vbInformation, frmOwnerForm.Caption
    Resume Next

End Function

' Fan-Out: GetDBID
' Fan-In:
'
Private Function AppendWhere(ByVal striSQL As String, ByVal striWhere As String) As String

    Dim strRunSQL As String
    Dim strRunSplit() As String
    Dim strWhere As String
    Dim strOrderBy As String
    
    Dim intOrderLoc As Integer
    Dim intOrderEndLoc As Integer
    
    Dim blnExist As Boolean
    
    strRunSQL = UCase$(striSQL)
    blnExist = (InStr(1, strRunSQL, "WHERE", vbTextCompare) <> 0)
'    strRunSQL = UCase$(striSQL)
    
    ' check if WHERE clause exist
    If blnExist = True Then
        strRunSplit = Split(strRunSQL, "WHERE", , vbTextCompare)
        strRunSQL = strRunSplit(0) & " WHERE (" & striWhere & ") AND " & strRunSplit(1)
    ElseIf blnExist = False Then
        strRunSQL = striSQL & " WHERE " & striWhere
    End If
    
    'mark 12092002 , to be deleted if there's a solution already for auto positioning & ordering of picklists
    intOrderLoc = InStr(1, UCase$(strRunSQL), "ORDER BY")
    If (intOrderLoc > 0) Then
    
        intOrderEndLoc = InStr(intOrderLoc + 9, strRunSQL, " ")
        
        If (intOrderEndLoc = 0) Then
            intOrderEndLoc = Len(strRunSQL) + 1
        End If
        
        strOrderBy = Mid$(strRunSQL, intOrderLoc, intOrderEndLoc - intOrderLoc)
        strRunSQL = Replace(strRunSQL, Trim$(strOrderBy), " ")
        strRunSQL = strRunSQL & " " & strOrderBy & " "
        
    End If
    '--------------------------------------------------------------------------------------------------------------------
    
    AppendWhere = strRunSQL

End Function

' Fan-Out: cmdCatalogOps_Click
Private Sub CheckChildTrans()

    Static intLastTrans As Integer
    Dim intIndex As Integer
    
    If (clsRecord.ChildTransactions.Count > 0) Then
    
        For intIndex = intLastTrans To clsRecord.ChildTransactions.Count - 1
        
            clsRecord.ChildTransactions(intIndex + 1).DBID = clsRecord.TempDBID
            
        Next intIndex
        
        intLastTrans = intIndex
        
    ElseIf (clsRecord.ChildTransactions.Count = 0) Then
    
        intLastTrans = 0
        
    End If

End Sub

' Fan-Out: cmdCatalogOps_Click
Private Sub ReconcileRecord()
         
    ' reconcile the values for transaction and DB
    Dim intTransIndex As Integer
    Dim blnExistInTrans As Boolean
    
    blnExistInTrans = False
    
    For intTransIndex = 1 To clsPicklist.Transactions.Count
    
        If (clsPicklist.Transactions(intTransIndex).RecordSource.Fields(clsPicklist.PKFieldBaseName).Value = _
                                                            clsRecord.RecordSource.Fields(clsPicklist.PKFieldBaseName).Value) Then
            blnExistInTrans = True
            Exit For
            
        End If
        
    Next intTransIndex
    
    If (blnExistInTrans = True) Then
    
        Set clsRecord.RecordSource = RstCopy(clsPicklist.Transactions(intTransIndex).RecordSource, True, 0, 0, 1)
        
    End If

End Sub

' Fan-Out: cmdCatalogOps_Click
Private Function GetTransPKValue()
         
   ' reconcile the values for transaction and DB
   Dim intTransIndex As Integer
   Dim strTransPK As String

    If ((clsPicklist.Transactions Is Nothing) = False) Then
    
       For intTransIndex = 1 To clsPicklist.Transactions.Count
       
          strTransPK = strTransPK _
                & clsPicklist.Transactions(intTransIndex).RecordSource.Fields(clsPicklist.PKFieldBaseName).Value & "|"
                
       Next intTransIndex
    
       If (strTransPK = "") Then
       
          GetTransPKValue = "Empty"
          
       ElseIf (strTransPK <> "") Then
       
          GetTransPKValue = Split(strTransPK, "|", , vbTextCompare)
          
       End If
       
    End If
      
End Function

Private Sub InitButton()

    ' invisible buttons
    If ((clsPicklist.AddButton = False) And (clsPicklist.PicklistStyle <> cpiSimplePicklist)) Then
    
        cmdAdd.Move cmdCatalogOps(CMD_ADD).Left, cmdCatalogOps(CMD_ADD).Top, _
                                    cmdCatalogOps(CMD_ADD).Width, cmdCatalogOps(CMD_ADD).Height
        cmdAdd.Caption = "&Add"
        cmdAdd.Visible = True
        cmdAdd.Enabled = False
        cmdCatalogOps(CMD_ADD).Visible = False
        
    End If
    
    If ((clsPicklist.ModifyButton = False) And (clsPicklist.PicklistStyle <> cpiSimplePicklist)) Then
    
        cmdModify.Move cmdCatalogOps(CMD_MODIFY).Left, cmdCatalogOps(CMD_MODIFY).Top, _
                                        cmdCatalogOps(CMD_MODIFY).Width, cmdCatalogOps(CMD_MODIFY).Height
        cmdModify.Caption = "&Modify"
        cmdModify.Visible = True
        cmdModify.Enabled = False
        cmdCatalogOps(CMD_MODIFY).Visible = False
    
    End If
    
    If ((clsPicklist.CopyButton = False) And (clsPicklist.PicklistStyle <> cpiSimplePicklist)) Then
    
        cmdCopy.Move cmdCatalogOps(CMD_COPY).Left, cmdCatalogOps(CMD_COPY).Top, _
                                        cmdCatalogOps(CMD_COPY).Width, cmdCatalogOps(CMD_COPY).Height
        cmdCopy.Caption = "&Copy"
        cmdCopy.Visible = True
        cmdCopy.Enabled = False
        cmdCatalogOps(CMD_COPY).Visible = False
    
    End If
    
    If ((clsPicklist.DeleteButton = False) And (clsPicklist.PicklistStyle <> cpiSimplePicklist)) Then
    
        cmdDelete.Move cmdCatalogOps(CMD_DELETE).Left, cmdCatalogOps(CMD_DELETE).Top, _
                                            cmdCatalogOps(CMD_DELETE).Width, cmdCatalogOps(CMD_DELETE).Height
        cmdDelete.Caption = "&Delete"
        cmdDelete.Visible = True
        cmdDelete.Enabled = False
        cmdCatalogOps(CMD_DELETE).Visible = False
        
    End If
    
    If Len(Trim(mvarWebLink)) > 0 Then
        cmdSeeOnLine.Visible = True
    End If
End Sub

' Fan-Out: SetSelectedRecord,Form_Unload
Private Function GetSelectedRec() As ADODB.Recordset
            
    Dim rstDummy As ADODB.Recordset
    
    ' not all deleted
    If (rstRecordsList.RecordCount <> 0) Then
        
        Call RepositionRst
        
        ' check if in db
        'If blnOperator = False Then 'allan feb26
            Set rstDummy = GetTopOne(conDBConnection, clsPicklist.BaseSQL, True, clsPicklist.PKFieldBaseName, _
                                        rstRecordsList.Fields(clsPicklist.PKFieldBaseName).Value)
'        Else
'
'            Set rstDummy = New ADODB.Recordset
'
'            rstDummy.Open clsPicklist.BaseSQL & " AND [OPERATOR_VENTURENUMBER]= '" & jgxPicklist.Value(5) & "'", conDBConnection, adOpenKeyset, adLockOptimistic 'allan feb22
'
'        End If

        ' if not found then get the grid in trans
        If (rstDummy.RecordCount = 0) Then
        
            ' if not found
            Set rstDummy = clsRecord.RecordSource
        
        ElseIf (rstDummy.RecordCount <> 0) Then
        
            ' reconcile pk value
            If (rstRecordsList.Fields("Tag").Value = "M") Then
            
                Set rstDummy = clsRecord.RecordSource
                
            End If
            
        End If
        
        Set GetSelectedRec = rstDummy
        
    ElseIf (rstRecordsList.RecordCount = 0) Then
    
        Set GetSelectedRec = Nothing
    
    End If
    
    Set rstDummy = Nothing

End Function

' users - RunDeleteTrans
Private Function blnPKinDB(ByVal striValue As String) As Boolean
      
    Dim rstDummy As ADODB.Recordset
    Dim blnReturn As Boolean
    
    Set rstDummy = New ADODB.Recordset
    
    Set rstDummy = GetTopOne(conDBConnection, clsPicklist.BaseSQL, False, clsPicklist.PKFieldBaseName, striValue)
    
    If (rstDummy.RecordCount <> 0) Then
        blnReturn = True
    ElseIf (rstDummy.RecordCount = 0) Then
        blnReturn = False
    End If
    
    blnPKinDB = blnReturn
    
    If Not rstDummy Is Nothing Then
        If rstDummy.State = adStateOpen Then
            rstDummy.Close
        End If
        Set rstDummy = Nothing
    End If
    
End Function

' Fan=Out: ShowPicklist
Private Sub CheckAutoSearch()
   
    Dim strTempFilter As String
    Dim strFirstChar As String
    Dim strRunSQL As String
    Dim strTempSQL As String
    
    Dim rstDummy As ADODB.Recordset
    
    ' ************************************************ '

    If (clsPicklist.AutoSearch = True) Then
    
        Set rstDummy = New ADODB.Recordset
        
        strTempSQL = RegenerateSQL(clsPicklist.BaseSQL, True)
        
        ' check filter
        If ((clsFilter Is Nothing) = True) Then
            
            strRunSQL = strTempSQL
        
        ElseIf ((clsFilter Is Nothing) = False) Then
            
            strTempFilter = GetActiveFilter(clsRecord.ActiveFilters)
            ' set proper run sql
            strRunSQL = InsertWhere(strTempSQL, strTempFilter)
        
        End If
        
        Set rstDummy = rstDBCopy(conDBConnection, strRunSQL, adOpenKeyset, adLockOptimistic, False)
        
        ' check if db is empty
        If (rstDummy.State = adStateClosed) Then
            blnHaltExecution = True
            Exit Sub
        End If
        
        If (rstDummy.RecordCount <> 0) Then
            
            ' init recordset
            rstDummy.MoveFirst
            strFirstChar = Left$(clsPicklist.SearchValue, 1)
            
            ' check if
            strTempFilter = GetWhereClause(clsPicklist.SearchField, clsPicklist.SearchValue _
                                    , rstDummy.Fields(clsPicklist.SearchField).Type)
            rstDummy.Filter = strTempFilter
            
            ' default ascending
            If (rstDummy.RecordCount = 0) Then
            
                cpiActiveStatus = cpiNotFound
                
                ' update base sql
                Call UpdateBaseSql(rstDummy)
                
            ElseIf (rstDummy.RecordCount = 1) Then
            
                cpiActiveStatus = IsExactRecord(rstDummy)
                
            ElseIf (rstDummy.RecordCount > 1) Then
            
                ' check if there is only one exact record
                strTempFilter = GetWhereClause(clsPicklist.SearchField, clsPicklist.SearchValue _
                                        , rstDummy.Fields(clsPicklist.SearchField).Type, True)
                rstDummy.Filter = strTempFilter
                
                If (rstDummy.RecordCount = 1) Then
                
                    cpiActiveStatus = IsExactRecord(rstDummy)
                    
                ElseIf (rstDummy.RecordCount <> 1) Then
                
                    cpiActiveStatus = cpiManyRecord
                    Call UpdateBaseSql(rstDummy)
            
                End If
                
            End If
            
        End If
            
        Set rstDummy = Nothing
        clsPicklist.ActiveStatus = cpiActiveStatus
    
    End If
   
End Sub

' Fan-Out: CheckAutoSearch
Private Function IsExactRecord(ByRef rstDummy As ADODB.Recordset) As cpiActiveStatusConstants

    Dim blnFlag As Boolean
    
    blnFlag = UCase$(rstDummy.Fields(clsPicklist.SearchField).Value) = UCase$(clsPicklist.SearchValue)
    
    If (blnFlag = True) Then
    
        IsExactRecord = cpiOneRecordExact
        
        ' return if enter key else nothing - update returned record
        If (clsPicklist.ActiveKey = cpiKeyEnter) Then
        
            Set clsPicklist.SelectedRecord = New CRecord
            Set clsPicklist.SelectedRecord.RecordSource = RstCopy(rstDummy, True, 0, 0, 1)
        
        End If
        
    ElseIf (blnFlag = False) Then
        
        IsExactRecord = cpiOneRecord
    
    End If
    
End Function

' Fan-Out: CheckAutoSearch
Private Sub UpdateBaseSql(ByRef rstField As ADODB.Recordset)

    Dim strField As String
    Dim strSplit() As String
    Dim blnExist As Boolean
    
    On Error GoTo Error_Handler
    If IsNull(rstField.Fields(clsPicklist.SearchField).Properties("BASECOLUMNNAME").Value) Then
        strField = clsPicklist.SearchField
    Else
        strField = rstField.Fields(clsPicklist.SearchField).Properties("BASECOLUMNNAME").Value
    End If
    
    blnExist = (InStr(1, UCase$(clsPicklist.BaseSQL), " ORDER BY ", vbTextCompare) = 0)
    
    ' set proper order by clause if not found
    If (blnExist = True) Then
        clsPicklist.BaseSQL = clsPicklist.BaseSQL & " ORDER BY [" & strField & "]"
    ElseIf (blnExist = False) Then
        strSplit = Split(UCase$(clsPicklist.BaseSQL), " ORDER BY ", , vbTextCompare)
        clsPicklist.BaseSQL = strSplit(0) & " ORDER BY [" & strField & "], " & strSplit(1)
    End If
   
    Exit Sub
    
Error_Handler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    Err.Clear
    
End Sub

' Fan-Out: ShowPicklist,cmdCatalogOps_Click,txtSearchFields_Change
Private Function LoadGrid(ByVal ActiveStatus As cpiActiveStatusConstants, _
                                          ByVal SearchField As String, _
                                          ByVal SearchValue As String) As Boolean

    Dim lngRowPos As Long
    Dim lngLastPosition As Long
    Dim intFieldCtr As Integer
    Dim strTempFilter As String
    
    ' exit when record is empty
    If (jgxPicklist.ADORecordset.RecordCount = 0) Then
        Exit Function
    End If
    
    lngLastPosition = lngAbsolutePosition
    
    Select Case ActiveStatus
    
        Case cpiManyRecord
        
            ' position to the first record
            jgxPicklist.ADORecordset.MoveFirst
            
            If (Trim$(SearchValue) <> "") Then
            
                ' set proper where
                strTempFilter = GetWhereClause(SearchField, SearchValue _
                                     , jgxPicklist.ADORecordset.Fields(SearchField).Type, False)
                                     
                jgxPicklist.ADORecordset.Filter = adFilterNone
                jgxPicklist.ADORecordset.Filter = strTempFilter
                If jgxPicklist.ADORecordset.RecordCount > 0 Then
                    RefreshGrid cpiRefresh, , , strTempFilter
                Else
                    jgxPicklist.ADORecordset.Filter = adFilterNone
                    RefreshGrid cpiRefresh
                End If
                jgxPicklist.Refresh
                
                'jgxPicklist.ADORecordset.Find strTempFilter, , adSearchForward, 0
                
                If (jgxPicklist.ADORecordset.EOF = True) Then
                    If jgxPicklist.ADORecordset.RecordCount > 0 Then
                        jgxPicklist.ADORecordset.AbsolutePosition = lngLastPosition
                    End If
                End If
                
                If jgxPicklist.ADORecordset.RecordCount > 0 Then
                    lngRowPos = jgxPicklist.ADORecordset.AbsolutePosition
                
                
                    If ((jgxPicklist.ADORecordset.EOF = False) And (lngRowPos <> 0)) Then
                        ' stop
                        Do While (jgxPicklist.Row <> lngRowPos)
                            jgxPicklist.Row = lngRowPos
                        Loop
                    End If
                End If
                
            ElseIf (Trim$(SearchValue) = "") Then
                    
                jgxPicklist.ADORecordset.Filter = adFilterNone
                strTempFilter = ""
                RefreshGrid cpiRefresh, , , strTempFilter
                jgxPicklist.Refresh
                lngRowPos = 1
                jgxPicklist.Row = lngRowPos
                
            End If
        
        Case cpiNotFound
            jgxPicklist.ADORecordset.Filter = adFilterNone
            strTempFilter = ""
            RefreshGrid cpiRefresh, , , strTempFilter
            jgxPicklist.Refresh
            jgxPicklist.ADORecordset.MoveFirst
            
            
            lngRowPos = jgxPicklist.ADORecordset.AbsolutePosition
            jgxPicklist.Row = lngRowPos
        
        Case cpiOneRecord, cpiOneRecordExact
        
            ' position to the first record
            jgxPicklist.ADORecordset.MoveFirst
            
            If (SearchValue <> "") Then
                ' set proper where
                strTempFilter = GetWhereClause(SearchField, SearchValue _
                                     , jgxPicklist.ADORecordset.Fields(SearchField).Type, False)
                
                If (jgxPicklist.ADORecordset.RecordCount > 0) Then
                
                    jgxPicklist.ADORecordset.Find strTempFilter, , adSearchForward, 0
                    
                    If (jgxPicklist.ADORecordset.EOF = True) Then
                    
                        jgxPicklist.ADORecordset.AbsolutePosition = lngLastPosition
                        lngRowPos = jgxPicklist.ADORecordset.AbsolutePosition
                        
                    ElseIf (jgxPicklist.ADORecordset.EOF = False) Then
                    
                        lngRowPos = jgxPicklist.ADORecordset.AbsolutePosition
                        
                    End If
                
                End If
                
            End If
            
            If (((jgxPicklist.ADORecordset.EOF = False) And (lngRowPos <> 0)) = True) Then
            
                ' stop
                Do While (jgxPicklist.Row <> lngRowPos)
                    jgxPicklist.Row = lngRowPos
                Loop
                
            ElseIf (((jgxPicklist.ADORecordset.EOF = False) And (lngRowPos <> 0)) = False) Then
            
                lngRowPos = lngLastPosition
                
            End If
    
    End Select
    
    ' patches
    If (rstRecordsList.AbsolutePosition <> lngRowPos) Then
    
        If (jgxPicklist.ADORecordset.RecordCount > 0) Then
            rstRecordsList.AbsolutePosition = lngRowPos
        End If
    End If
    
    ' resynchronize the records
    Call ReSynchRecord(lngRowPos)
    
End Function

' Fan-Out: ShowPicklist,cmdCatalogOps_Click,optFilter_Click,txtSearchFields_Change
Private Sub UpdateSearchField(Optional ByVal ExceptFieldIndex As Integer = -1)

    Dim intFieldCtr As Integer
    
    ' exit when record is empty
    If (rstRecordsList.RecordCount <> 0) Then
    
        blnSkipTextChange = True
        
        For intFieldCtr = 0 To txtSearchFields.Count - 1
        
            If (txtSearchFields(intFieldCtr).Tag <> "") Then
            
                If (intFieldCtr <> ExceptFieldIndex) Then
                
                    txtSearchFields(intFieldCtr).Text = IIf(IsNull(rstRecordsList.Fields(txtSearchFields(intFieldCtr).Tag).Value), "", _
                                        rstRecordsList.Fields(txtSearchFields(intFieldCtr).Tag).Value)
                    
                End If
                
            End If
        
        Next intFieldCtr
        
        blnSkipTextChange = False
    
    ElseIf (rstRecordsList.RecordCount = 0) Then
    
        For intFieldCtr = 0 To txtSearchFields.Count - 1
        
            blnSkipTextChange = True
            txtSearchFields(intFieldCtr).Text = ""
            blnSkipTextChange = False
            
        Next intFieldCtr
        
    End If
    
End Sub

' Fan-Out: LoadGrid
Private Sub ReSynchRecord(ByVal lngNewPos As Long)
   
    If (jgxPicklist.ADORecordset.RecordCount > 0) Then
        Do While (jgxPicklist.Row <> lngNewPos)
   
            jgxPicklist.Row = lngNewPos
      
        Loop
   
        rstRecordsList.AbsolutePosition = lngNewPos
    End If
   
End Sub

' DEBUGGER
Private Sub Debug_Position(strMsg As String)
'
   Debug.Print strMsg
   Debug.Print "lngAbsolutePosition=" & lngAbsolutePosition
   Debug.Print "rstRecordslist.AbsolutePosition=" & rstRecordsList.AbsolutePosition
   Debug.Print "jgxpicklist.Row=" & jgxPicklist.Row
   Debug.Print "jgxPicklist.RowIndex=" & jgxPicklist.RowIndex(jgxPicklist.Row)
   Debug.Print "jgxPicklist.ADORecordset.AbsolutePosition=" & jgxPicklist.ADORecordset.AbsolutePosition

End Sub

' Fan-Out: Form_Activate
Private Sub GridSelectedField_Sort()

   ' sort the grid in ascending order
   Call jgxPicklist_ColumnHeaderClick(jgxPicklist.Columns(clsPicklist.SearchField))
   jgxPicklist.RefreshSort
   jgxPicklist.RefreshRowIndex (jgxPicklist.RowIndex(jgxPicklist.Row))
   jgxPicklist.RefreshRowBookmark (jgxPicklist.ADORecordset.Bookmark)
   jgxPicklist.RefreshGroups

End Sub

' Fan-Out: cmdFilter_Click,CheckAutoSearch,LoadGrid
Private Function GetWhereClause(ByVal FieldAlias As String, _
                              ByVal FindValue As Variant, _
                              FieldType As DataTypeEnum, _
                              Optional ByVal blnExactRec As Boolean, _
                              Optional ByVal sWhereEnd As String) As String
   
   Dim strWhereStart As String
   Dim strWhereEnd As String
   Dim strWhere As String
   Dim strValue As String
   
   FindValue = Replace(FindValue, "'", "''", , , vbTextCompare)
   FindValue = Replace(FindValue, "%", "%%", , , vbTextCompare)
   FindValue = Replace(FindValue, "*", "**", , , vbTextCompare)
   
   strValue = Trim$(FindValue)
   
   ' set the proper SQL WHERE parameters
     'field aliases
   Select Case FieldType
      
        ' adSmallInt = 2
        ' adInteger = 3
        ' adSingle = 4
        ' adDouble = 5
        ' adCurrency = 6
        ' adIDispatch = 9
        ' adDecimal = 14
        ' adTinyInt= 16
        ' adUnsignedTinyInt = 17
        ' adUnsignedSmallInt = 18
        ' adUnsignedInt = 19
        ' adBigInt = 20
        ' adUnsignedBigInt = 21
        ' adBinary = 128
        ' adNumeric = 131
        ' adChapter = 136
        ' adPropVariant= 138
        ' adLongVarBinary = 205
        Case adUnsignedInt, adUnsignedSmallInt, adInteger, adUnsignedTinyInt, _
                adSingle, adDouble, adCurrency, adNumeric, adDecimal, _
                adLongVarBinary, adIDispatch, adPropVariant, adChapter, _
                adBinary, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt
                
            strWhereStart = " = "
            strWhereEnd = ""
            
            ' check if the value is numeric otherwise set it to 0
            If (IsNumeric(strValue) = False) Then
               strValue = "0"
            ' check if var value exceeds the integer limit
            ElseIf ((FindValue > (2 ^ 16)) Or (FindValue < -(2 ^ 16))) Then
               strValue = "0"
            End If
      
        ' adBSTR = 8                                            ' String
        ' dbText = 10
        ' dbMemo = 12
        ' adChar = 129
        ' adWChar = 130
        ' adVarChar = 200
        ' adLongVarChar = 201
        ' adVarWChar = 202
        ' adLongVarWChar= 203
        ' adVarBinary = 204
        Case adChar, adLongVarChar, adVarBinary, adVarChar, _
                adVarWChar, adWChar, 10, 12, adBSTR
         
            If (blnExactRec = False) Then
               strWhereStart = " LIKE '" & "*"
               'strWhereStart = " LIKE #" & "*"
               
               If (strValue <> "") Then
                  strWhereEnd = sWhereEnd & "*" & "' "
                  'strWhereEnd = sWhereEnd & "*" & "# "
               
               ElseIf (strValue = "") Then
                  'strWhereEnd = Chr$(255) & "# "
                  strWhereEnd = "*' "
               
               End If
            
            ElseIf (blnExactRec = True) Then
                           
               'strWhereStart = " = #"
               strWhereStart = " = '"
               
               
               If (strValue <> "") Then
                   strWhereEnd = sWhereEnd & "*" & "' "
                   'strWhereEnd = sWhereEnd & "*" & "# "
               
               ElseIf (strValue = "") Then
                  
                  'strWhereEnd = Chr$(255) & "# "
                  'strWhereEnd = Chr$(255) & "*' "
                  strWhereEnd = "*' "
                   
                   
               End If
            End If
         
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            strWhereStart = " = #"
            strWhereEnd = "# "
            
            ' check if the value is date otherwise set it to date today
            If (IsDate(strValue) = False) Then
               strValue = CStr(Date)
            End If
                    
   End Select
             
   If LenB(Trim$(strValue)) > 0 Then
        strWhere = "[" & FieldAlias & "]" & strWhereStart & strValue & strWhereEnd
    Else
        strWhere = vbNullString
    End If
   GetWhereClause = strWhere

End Function

' Fan-Out: Form_Activate,RefreshGrid
Private Sub SetGridProperty()

    Dim intGridIndex As Integer
    Dim intIndex As Integer
    
    For intGridIndex = 1 To clsGridSeed.GridColumns.Count
    
        intIndex = GetColIndex(jgxPicklist.Columns, intGridIndex)
        jgxPicklist.Columns(intIndex).Format = clsGridSeed.GridColumns(intGridIndex).Format
        jgxPicklist.Columns(intIndex).Width = clsGridSeed.GridColumns(intGridIndex).ColumnWidth
        jgxPicklist.Columns(intIndex).Caption = clsGridSeed.GridColumns(intGridIndex).ColumnFieldAias
        
        ' adder
        Select Case UCase$(clsGridSeed.GridColumns(intGridIndex).ColumnAlignment)
        
            Case "LEFT"
                jgxPicklist.Columns(intIndex).HeaderAlignment = jgexAlignLeft
                jgxPicklist.Columns(intIndex).TextAlignment = jgexAlignLeft
                txtSearchFields(intGridIndex - 1).Alignment = vbLeftJustify
                
            Case "RIGHT"
                jgxPicklist.Columns(intIndex).HeaderAlignment = jgexAlignRight
                jgxPicklist.Columns(intIndex).TextAlignment = jgexAlignRight
                txtSearchFields(intGridIndex - 1).Alignment = vbRightJustify
                
            Case "CENTER"
                jgxPicklist.Columns(intIndex).HeaderAlignment = jgexAlignCenter
                jgxPicklist.Columns(intIndex).TextAlignment = jgexAlignCenter
                txtSearchFields(intGridIndex - 1).Alignment = vbCenter
        
        End Select
    
    Next intGridIndex
    
    ' patches
    If (rstRecordsList.RecordCount > 0) Then
        cmdTransact(CMD_OK).Enabled = True
    End If
    
    ' reconcile patches here
    If (rstRecordsList.RecordCount = 0) Then
    
        cmdCatalogOps(CMD_MODIFY).Enabled = False
        cmdCatalogOps(CMD_COPY).Enabled = False
        cmdCatalogOps(CMD_DELETE).Enabled = False
        
    ElseIf (rstRecordsList.RecordCount <> 0) Then
    
        cmdCatalogOps(CMD_MODIFY).Enabled = clsPicklist.ModifyButton
        cmdCatalogOps(CMD_COPY).Enabled = clsPicklist.CopyButton
        cmdCatalogOps(CMD_DELETE).Enabled = clsPicklist.DeleteButton
        
    End If

End Sub

' Fan-Out: SetGridSetting,cmdCatalogOps_Click,RefreshGrid,jgxPicklist_DblClick,GetSelectedRec,LoadGrid
Private Sub RepositionRst()
   
    If (rstRecordsList.RecordCount <> 0) Then
    
        If (rstRecordsList.EOF = True) Then
            rstRecordsList.MoveLast
        End If
    
        If (rstRecordsList.BOF = True) Then
            rstRecordsList.MoveFirst
        End If
        
    End If
   
End Sub

' Fan-Out: SetGridColumns
Private Sub SetGridVisible(blnIsVisible As Boolean)

    Dim intGridIndex As Integer
    
    For intGridIndex = 1 To jgxPicklist.Columns.Count
    
        jgxPicklist.Columns(intGridIndex).Visible = blnIsVisible
        
    Next intGridIndex
   
End Sub

' Fan-Out: ShowPicklist
Private Sub InitFilterInterface()

    Dim intFilterIndex As Integer
    
    For intFilterIndex = 1 To clsRecord.ActiveFilters.Count
    
        Select Case clsPicklist.PicklistFilter.FilterType
        
            Case cpiCheckOptions
            
                chkFilter(intFilterIndex - 1).Enabled = clsRecord.ActiveFilters(intFilterIndex).Enabled
                
            Case cpiRadioOptions
            
                optFilter(intFilterIndex - 1).Enabled = clsRecord.ActiveFilters(intFilterIndex).Enabled
        
        End Select
    
    Next intFilterIndex

End Sub

' Fan-Out: CheckAutoSearch
Private Function GetActiveFilter(clsActiveFilters As CPicklistFilters) As String

    Dim intFilterIndex As Integer
    Dim varTempFilter() As String
    Dim strTempFilter As String
    Dim intIndex As Integer
    
    ReDim varTempFilter(clsActiveFilters.Count - 1)
    intIndex = 0
    
    For intFilterIndex = 1 To UBound(varTempFilter) + 1
    
        If (clsActiveFilters(intFilterIndex).State = True) Then
        
           varTempFilter(intIndex) = clsActiveFilters(intFilterIndex).Filter
           intIndex = intIndex + 1
           
        End If
        
    Next intFilterIndex
    
    If (clsFilter.FilterType = cpiRadioOptions) Then
    
        strTempFilter = ConnectFilters("OR", varTempFilter)
        
    ElseIf (clsFilter.FilterType = cpiCheckOptions) Then
    
        strTempFilter = ConnectFilters("AND", varTempFilter)
        
    End If
   
   GetActiveFilter = strTempFilter
   
End Function

' Fan-Out: GetActiveFilter
Private Function ConnectFilters(ByRef Operand As String, ByRef FilterItems() As String) As String

    Dim intFilterIndex As Integer
    Dim strTempFilter As String
    
    strTempFilter = ""
    
    For intFilterIndex = 0 To UBound(FilterItems)
    
        If (FilterItems(intFilterIndex) <> "") Then
        
            strTempFilter = strTempFilter & " " & Operand & " " & FilterItems(intFilterIndex)
            
        End If
    
    Next intFilterIndex
    
    If (strTempFilter <> "") Then
    
        strTempFilter = Right$(strTempFilter, Len(strTempFilter) - Len(Operand) - 2)
        
    End If
    
    ConnectFilters = strTempFilter

End Function

' Fan-Out: cmdOK_Click
Private Sub SetSelectedRecord()

    Dim lngTransIndex As Integer
    Dim varPk As String
    Dim clsReturnRecord As CRecord
    Dim rstReturn As ADODB.Recordset
    
    If (blnCanceled = False) Then
    
        Set clsReturnRecord = New CRecord
        Set rstReturn = New ADODB.Recordset
        
        If (blnIsInTrans = False) Then
        
            ' reconcile positions
            If (rstRecordsList.AbsolutePosition <> lngAbsolutePosition) Then
            
                If (rstRecordsList.RecordCount <> 0) Then
                
                    If (lngAbsolutePosition > 0) Then
                    
                        rstRecordsList.AbsolutePosition = lngAbsolutePosition
                        
                    End If
                
                End If
            
            End If
        
            ' check if in trans
            Set rstReturn = GetSelectedRec
    
        ElseIf (blnIsInTrans = True) Then
        
            ' get in transaction
            Set rstReturn = clsRecord.RecordSource
            
        End If
            
        Set clsReturnRecord.RecordSource = rstReturn
        
        'no selected, no record
        If (rstRecordsList.RecordCount = 0) Then
        
            Set clsPicklist.SelectedRecord = Nothing
            
        ElseIf (rstRecordsList.RecordCount <> 0) Then
        
            Set clsPicklist.SelectedRecord = clsReturnRecord
            
        End If
    
    End If
    
End Sub

' Fan-Out: ProcessRecord
Private Sub ReconcileTransPK()
   
    Dim intTransIndex As Integer
    Dim intCheckIndex As Integer
    Dim strTempPK As String
    
    If ((clsPicklist.Transactions Is Nothing) = False) Then
    
        ' loop through all transactions
        For intTransIndex = 1 To clsPicklist.Transactions.Count - 1
        
            For intCheckIndex = intTransIndex + 1 To clsPicklist.Transactions.Count
            
                If (clsPicklist.Transactions(intTransIndex).RecordSource(clsPicklist.PKFieldBaseName).Value = _
                            clsPicklist.Transactions(intCheckIndex).RecordSource(clsPicklist.PKFieldBaseName).Value) Then
                
                    clsPicklist.Transactions(intCheckIndex).RecordSource(clsPicklist.PKFieldBaseName).Value = _
                                clsPicklist.Transactions(intCheckIndex).RecordSource(clsPicklist.PKFieldBaseName).Value + 1
                    clsPicklist.Transactions(intCheckIndex).RecordSource(clsPicklist.PKFieldAlias).Value = _
                                clsPicklist.Transactions(intCheckIndex).RecordSource(clsPicklist.PKFieldBaseName).Value
                
                End If
            
            Next intCheckIndex
        
        Next intTransIndex
    
    End If
    
End Sub

' Fan-Out: Form_Activate
Private Sub UpdateOption()

    Dim intIndex As Integer
    
    For intIndex = 1 To clsFilter.FilterCount
    
        If (clsFilter.PicklistFilters(intIndex).Value = True) Then
        
            If (clsFilter.FilterType = cpiCheckOptions) Then
            
                chkFilter(intIndex - 1).Value = Abs(clsFilter.PicklistFilters(intIndex).Value)
                
            ElseIf (clsFilter.FilterType = cpiRadioOptions) Then
            
                optFilter(intIndex - 1).Value = clsFilter.PicklistFilters(intIndex).Value
                Exit For
                
            End If
            
        End If
        
    Next intIndex
    
End Sub

' Fan-Out: FormInitialize
Private Sub ResetFormHeight()
    
    Dim clsRegistry As CRegistry
    
    Dim strRegErrMsg As String
    Dim strRootKeyPath As String
    Dim strSearchKey As String
    Dim strSearchValueName As String
    
    Dim varBorderWidth As Variant
    Dim varCaptionHeight As Variant
    Dim varMenuHeight As Variant
    
    Dim blnExist As Boolean
    Dim blnFlag As Boolean
    
    On Error GoTo ERROR_REG
    
    Set clsRegistry = New CRegistry
    
    strRegErrMsg = "The registry key/value cannot be found.  Your registry may be corrupted."
    
    ' check if Windows Control Panel\Desktop\Windows Metrics Folder exist
    strRootKeyPath = "Control Panel\Desktop"
    strSearchKey = "WindowMetrics"
    
    blnExist = clsRegistry.RegistryKeyExists(cpiCurrentUser, "", strRootKeyPath, strSearchKey, cpiUserDefined, True)
    
    If (blnExist = True) Then
    
        strRootKeyPath = strRootKeyPath & "\" & strSearchKey
        
        ' get metrics
        blnFlag = clsRegistry.GetRegistry(cpiCurrentUser, "", strRootKeyPath, "BorderWidth", cpiUserDefined)
        If (blnFlag = True) Then
            varBorderWidth = clsRegistry.RegistryValue
        End If
        
        blnFlag = clsRegistry.GetRegistry(cpiCurrentUser, "", strRootKeyPath, "CaptionHeight", cpiUserDefined)
        If (blnFlag = True) Then
            varCaptionHeight = clsRegistry.RegistryValue
        End If
        
        blnFlag = clsRegistry.GetRegistry(cpiCurrentUser, "", strRootKeyPath, "MenuHeight", cpiUserDefined)
        If (blnFlag = True) Then
            varMenuHeight = clsRegistry.RegistryValue
        End If
        
        ' set true height
        'Height = Height - (STD_CAPTION_HEIGHT + Abs(varCaptionHeight))
        
        If (varCaptionHeight <> "") And (IsNumeric(varCaptionHeight) = True) Then
            
            If (clsPicklist.PicklistStyle <> cpiSimplePicklist) Then
                Height = Height - (STD_CAPTION_HEIGHT + (varCaptionHeight))
            ElseIf (clsPicklist.PicklistStyle = cpiSimplePicklist) Then
                Height = Height - (STD_CAPTION_HEIGHT - (varCaptionHeight))
            End If
            
        End If
        
    ElseIf (blnExist = False) Then
    
        'MsgBox strRegErrMsg, vbInformation, frmOwnerForm.Caption
        
    End If
    
    Set clsRegistry = Nothing
    
    Exit Sub
    
ERROR_REG:
    
    Set clsRegistry = Nothing
    
End Sub

' Save the form's and controls' dimensions.
Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control

    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                .Left = ctl.X1
                .Top = ctl.Y1
                .Width = ctl.X2 - ctl.X1
                .Height = ctl.Y2 - ctl.Y1
            Else
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    g_FormWid = Me.Width / 15
    
    m_FormHgt = ScaleHeight
    g_FormHgt = Me.Height / 15
    
    
End Sub

' Arrange the controls for the new size.
Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single

    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Or Not blnFormActivated Then Exit Sub
    
    On Error Resume Next
    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            ElseIf TypeOf ctl Is CommandButton Then
                'Refers to OK, Cancel, Close, Apply, etc.. below the sstab/picklist control
                Select Case enuStyle
                    Case cpiSimplePicklist
                        If ctl.Top > (jgxPicklist.Top + jgxPicklist.Height) Then
                            'Ensures horizontal distance between buttons and their
                            'relative position to bottom side of form are constant
                            ctl.Top = ScaleHeight - m_FormHgt + .Top
                        End If
                    Case cpiCatalog, cpiFilterCatalog
                        If ctl.Top > (tabCatalog.Top + tabCatalog.Height) Then
                            'Ensures horizontal distance between buttons and their
                            'relative position to bottom side of form are constant
                            ctl.Top = ScaleHeight - m_FormHgt + .Top
                        End If
                End Select
                
                'Refers to other command buttons above the form
                'Ensures vertical distance between buttons and their
                'relative position to right side of form are constant
                '<<< dandan 110707
                'Add ed checking for url button
                If (ctl.Name = "cmdSeeOnLine") Then
                    cmdSeeOnLine.Left = txtSearchFields(0).Left
                Else
                    ctl.Left = ScaleWidth - m_FormWid + .Left
                End If
            ElseIf TypeOf ctl Is CheckBox Then
                'Ensures vertical distance between buttons and their
                'relative position to right side of form are constant
                ctl.Left = ScaleWidth - m_FormWid + .Left
            ElseIf TypeOf ctl Is TextBox Then
                'Textbox size increments proportional to form size increments
                If UCase$(Trim$(ctl.Name)) = "TXTSEARCHFIELDS" Then
                    Select Case ctl.Index
                        Case 0
                            ctl.Width = ScaleWidth - m_FormWid + .Width
                            'txtSearchFields(1).Left = txtSearchFields(1).Left + (ScaleWidth - m_FormWid)
                            'txtSearchFields(1).Width = txtSearchFields(1).Width + (ScaleWidth - m_FormWid)
                        Case Else
                            ctl.Width = ScaleWidth - m_FormWid + .Width
                    End Select
                Else
                    ctl.Width = ScaleWidth - m_FormWid + .Width
                End If
            ElseIf Not (TypeOf ctl Is Label) Then
                'Any other object size increments proprotional to form size increments
                On Error Resume Next
                ctl.Width = ScaleWidth - m_FormWid + .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = ScaleHeight - m_FormHgt + .Height
                End If
                On Error Resume Next
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl
    'On Error GoTo 0
End Sub

Private Function ByPassButton(Index) As Boolean

    Dim blnCancel As Boolean
    Dim enuOperation As RecordOperation


    enuOperation = Choose(Index + 1, cpiRecordAdd, cpiRecordEdit, cpiRecordCopy, cpiRecordDelete)

   If enuOperation = cpiRecordDelete Then
        clsPicklist.ProcPriorBtnClick clsPicklist.PKFieldBaseName, rstRecordsList.Fields(clsPicklist.PKFieldBaseName).Value, enuOperation, blnCancel
        ByPassButton = blnCancel
    End If


End Function

Private Sub txtSearchFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = "#" Then
        KeyAscii = 0
    End If
End Sub

Private Sub RepositionControls()
    Select Case enuStyle
    
        Case cpiSimplePicklist
            cmdTransact(1).Left = jgxPicklist.Left + jgxPicklist.Width - 1215
            cmdTransact(0).Left = cmdTransact(1).Left - 100 - 1215
            
            '<<< dandan 110707
            cmdSeeOnLine.Top = cmdTransact(0).Top
            cmdSeeOnLine.Left = txtSearchFields(0).Left
            
        Case cpiCatalog, cpiFilterCatalog
        
            tabCatalog.Width = 90 + jgxPicklist.Width + 210 + cmdCatalogOps(0).Width
            
            cmdCatalogOps(0).Top = jgxPicklist.Top + 315
            cmdCatalogOps(1).Top = cmdCatalogOps(0).Top + cmdCatalogOps(0).Height + 120
            cmdCatalogOps(2).Top = cmdCatalogOps(1).Top + cmdCatalogOps(1).Height + 120
            cmdCatalogOps(3).Top = cmdCatalogOps(2).Top + cmdCatalogOps(2).Height + 120
            
            cmdCatalogOps(0).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(1).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(2).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            cmdCatalogOps(3).Left = jgxPicklist.Left + jgxPicklist.Width + 105
            
            cmdFilter.Left = cmdCatalogOps(0).Left + 240
            chkClearFilter.Left = cmdFilter.Left + 480
            
            cmdTransact(1).Left = tabCatalog.Width + tabCatalog.Left - 1215
            cmdTransact(0).Left = cmdTransact(1).Left - 100 - 1215
                    
            cmdTransact(0).Top = tabCatalog.Height + 180
            cmdTransact(1).Top = tabCatalog.Height + 180
            
            '<<< dandan 110707
            cmdSeeOnLine.Top = tabCatalog.Height + 180
            cmdSeeOnLine.Left = txtSearchFields(0).Left
            
            Me.Height = tabCatalog.Height + cmdTransact(0).Height + 270 + 375
    End Select
End Sub
