VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStockcardPicklist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Cards"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8445
   Icon            =   "frmStockcardPicklist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Stock Cards"
      TabPicture(0)   =   "frmStockcardPicklist.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "jgxPicklist"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAdd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFilter(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFilter(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFilter(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFilter(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   4
         Left            =   4920
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   3
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   1200
         Width           =   1215
      End
      Begin GridEX16.GridEX jgxPicklist 
         Height          =   4935
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8705
         HideSelection   =   2
         MethodHoldFields=   -1  'True
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
End
Attribute VB_Name = "frmStockcardPicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pckStockCard As PCubeLibEntrepot.cStockCard
Private blnCancel As Boolean
Public blnAddCancel As Boolean

Private Sub cmdAdd_Click()
    frmStockcard.Pre_Load pckStockCard, ResourceHandler
        'Restores the default filter and sort of the original recordset.
        'Prevents the records from moving around when adding new entries.
        If Len(pckStockCard.m_rstPass2GridOff.Filter) <> 0 Then
            pckStockCard.m_rstPass2GridOff.Filter = ""
            pckStockCard.m_rstPass2GridOff.Sort = ""
            'Commented so newest item will appear at the bottom and become highlighted by .rowselected.
            'pckStockCard.m_rstPass2GridOff.Sort = "[Stock ID], [Stock Card No]"
            'pckStockCard.m_rstPass2GridOff.Sort = "[Length], [Stock Card No]"
        End If
    If blnAddCancel = False Then
        'Re-set to refresh display.
        Set jgxPicklist.ADORecordset = pckStockCard.m_rstPass2GridOff
        jgxPicklist.RowSelected(jgxPicklist.RowCount) = True
        HideSomeFields
        cmdOK.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    CleanUp2Exit True
End Sub

Public Sub Pre_Load(ByRef cpiStockCard As PCubeLibEntrepot.cStockCard, ByRef Cancelled As Boolean)
    Set pckStockCard = cpiStockCard
    'Uses filtered recordset for Grid.
    Set jgxPicklist.ADORecordset = cpiStockCard.m_rstPass2GridOff
    
    'Prevents user from clicking enter if picklist is empty.
    If jgxPicklist.ADORecordset Is Nothing Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
    'Added so new entries will always appear at the bottom instead of somewhere due to sorting.
    jgxPicklist.AutomaticArrange = False
    HideSomeFields
    Me.Show vbModal
    Cancelled = blnCancel
End Sub

Private Sub HideSomeFields()
    'Hides the Stock ID, Product ID, Entrepot ID, New and Length fields.
    With jgxPicklist
        .Columns(1).Visible = False
        .Columns(6).Visible = False
        .Columns(7).Visible = False
        .Columns(8).Visible = False
        .Columns(9).Visible = False
    End With
End Sub

Private Sub cmdOK_Click()
    CleanUp2Exit False
End Sub

Private Sub CleanUp2Exit(Cancel As Boolean)
    'Clean up.
    Set pckStockCard = Nothing
    blnCancel = Cancel
    Unload Me
End Sub
