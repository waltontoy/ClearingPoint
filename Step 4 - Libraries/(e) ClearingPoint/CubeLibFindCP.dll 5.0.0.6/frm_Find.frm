VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{343F59D0-FE0F-11D0-A89A-0000C02AC6DB}#1.0#0"; "sstbars.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Find 
   Caption         =   "Find: All Files"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6645
   Icon            =   "frm_Find.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Tag             =   "654"
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstSorter 
      Height          =   255
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   240
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "624"
      Text            =   "There are no items to show in this view."
      Top             =   3105
      Visible         =   0   'False
      Width           =   3435
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   675
      Left            =   5265
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1875
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   56
      FullHeight      =   45
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   33
      Top             =   5175
      Visible         =   0   'False
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNewSearch 
      Caption         =   "Ne&w Search"
      Height          =   345
      Left            =   5085
      TabIndex        =   5
      Tag             =   "616"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5085
      TabIndex        =   4
      Tag             =   "615"
      Top             =   900
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F&ind Now"
      Default         =   -1  'True
      Height          =   345
      Left            =   5085
      TabIndex        =   3
      Tag             =   "614"
      Top             =   480
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   60
      TabIndex        =   7
      Top             =   45
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "610"
      TabPicture(0)   =   "frm_Find.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboName"
      Tab(0).Control(1)=   "icbLookIn"
      Tab(0).Control(2)=   "icbType"
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "747"
      TabPicture(1)   =   "frm_Find.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel2"
      Tab(1).Control(1)=   "SSPanel1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "612"
      TabPicture(2)   =   "frm_Find.frx":05C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "icbBox"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtValue"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboCondition"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkAllFields"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.CheckBox chkAllFields 
         Caption         =   "Use full &grid on next search"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   1920
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         Left            =   -73770
         TabIndex        =   0
         ToolTipText     =   "Wildcards available: * for any number of character, ? for single character. E.g. *title*0? = untitled_2005"
         Top             =   585
         Width           =   3420
      End
      Begin VB.ComboBox cboCondition 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1035
         Width           =   1950
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   1440
         Width           =   1950
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1305
         Left            =   -74985
         TabIndex        =   21
         Top             =   1155
         Width           =   4665
         _Version        =   65536
         _ExtentX        =   8229
         _ExtentY        =   2302
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Index           =   1
            Left            =   2660
            TabIndex        =   32
            Top             =   825
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   0
            Left            =   1530
            TabIndex        =   12
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   16515073
            CurrentDate     =   36098
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Index           =   0
            Left            =   2660
            TabIndex        =   22
            Top             =   420
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   0
            Left            =   2265
            TabIndex        =   15
            Text            =   "1"
            Top             =   390
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&between"
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Tag             =   "618"
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton Option2 
            Caption         =   "during the previou&s"
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Tag             =   "620"
            Top             =   420
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&during the previous"
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Tag             =   "621"
            Top             =   810
            Width           =   1890
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   1
            Left            =   3300
            TabIndex        =   13
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   16515073
            CurrentDate     =   36098
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   1
            Left            =   2265
            TabIndex        =   17
            Text            =   "1"
            Top             =   795
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "and"
            Height          =   195
            Index           =   0
            Left            =   2940
            TabIndex        =   25
            Tag             =   "619"
            Top             =   60
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "month(s)"
            Height          =   195
            Index           =   1
            Left            =   2940
            TabIndex        =   24
            Tag             =   "622"
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "day(s)"
            Height          =   195
            Index           =   2
            Left            =   2940
            TabIndex        =   23
            Tag             =   "623"
            Top             =   855
            Width           =   420
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   690
         Left            =   -74955
         TabIndex        =   26
         Top             =   405
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   1217
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Begin VB.ComboBox Combo5 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   330
            Width           =   1815
         End
         Begin VB.OptionButton optAll 
            Caption         =   "&All files"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Tag             =   "662"
            Top             =   60
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optFind 
            Caption         =   "Fi&nd all files"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Tag             =   "617"
            Top             =   420
            Width           =   1260
         End
      End
      Begin MSComctlLib.ImageCombo icbBox 
         Height          =   330
         Left            =   1080
         TabIndex        =   18
         Top             =   600
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo icbLookIn 
         Height          =   330
         Left            =   -73770
         TabIndex        =   2
         Top             =   1515
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imgImages"
      End
      Begin MSComctlLib.ImageCombo icbType 
         Height          =   330
         Left            =   -73770
         TabIndex        =   1
         Top             =   1080
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imgImages"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Type:"
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   35
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Condition"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Tag             =   "554"
         Top             =   1095
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Box"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Tag             =   "329"
         Top             =   645
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   29
         Tag             =   "496"
         Top             =   645
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Look in:"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   28
         Tag             =   "473"
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Tag             =   "451"
         Top             =   1545
         Width           =   405
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   65541
      ToolBarsCount   =   1
      ToolsCount      =   39
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frm_Find.frx":05DE
      ToolBars        =   "frm_Find.frx":10BC4
   End
   Begin MSComctlLib.ImageList imgImages 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":10C9B
            Key             =   "Approved"
            Object.Tag             =   "DO NOT REMOVE THIS CAN'T FIND .ICO equivalent"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":110F3
            Key             =   "Rejected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":11BBF
            Key             =   "Outbox"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":12013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":1232D
            Key             =   "Deleted"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":12647
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":12961
            Key             =   "Archive"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":12C7B
            Key             =   "Drafts"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":130CF
            Key             =   "Templates"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":13533
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":13E0F
            Key             =   "ClearingPoint"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":146EB
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":14C85
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":1521F
            Key             =   "Transit"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":157B9
            Key             =   "NCTS"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":15D53
            Key             =   "Combined"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":162ED
            Key             =   "EDIDepartures"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":16887
            Key             =   "EDIArrivals"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":16E21
            Key             =   "PLDA Import"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Find.frx":173BB
            Key             =   "PLDA Export"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItemsFound 
      Height          =   2415
      Left            =   30
      TabIndex        =   6
      Top             =   2655
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgImages"
      SmallIcons      =   "imgImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Document"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "In Folder"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Modified Index"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frm_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LISTSEPARATOR = "----------------------"

Private Enum enuDocType
    eDocAny = 0
    edocimport = 1
    eDocExport = 2
    eDocOTS = 3
    eDocNCTS = 4
    edoccombined = 5
    eDocEDIDepartures = 6
    eDocEDIARRIVALS = 7
    eDocPLDAImport = 8
    eDocPLDACombined = 9
End Enum

Private Enum enuLookIn
    eLookInAll = 0
    eLookInTemplates = 1
    eLookInApproved = 2
    eLookInArchive = 3
    eLookInDRAFTS = 4
    eLookInOutbox = 5
    eLookInRejected = 6
    eLookInDeleted = 7
    'CSCLP-248
    eLookInToBePrinted = 8                  'SADBEL, SADBEL NCTS
    eLookInReleased = 9                     'PLDA Import, PLDA Export
    eLookInSent = 10                        'NCTS Departure, PLDA Import, PLDA Export
    eLookInEmergencyProcedure = 11          'NCTS Departure, PLDA Import, PLDA Export
    eLookInExitEC = 12                      'NCTS Departure, PLDA Export
    eLookInGuarantee = 13                   'NCTS Departure
    eLookInUnderControl = 14                'NCTS Departure
    eLookInReleases = 15                    'NCTS Departure
    eLookInReleaseRejected = 16             'NCTS Departure
    eLookInCancelled = 17                   'NCTS Departure, PLDA Import, PLDA Export
    eLookInWrittenOff = 18                  'NCTS Departure
    eLookInAmendmentSent = 19               'NCTS Departure
    eLookInAmendmentAccepted = 20           'NCTS Departure
    eLookInAmendmentRejected = 21           'NCTS Departure
    eLookInArchives = 22                    'NCTS Arrivals
    eLookInArrivalNotificationSent = 23     'NCTS Arrivals
    eLookInArrivalNotificationRejected = 24 'NCTS Arrivals
    eLookInUnloadingPermitted = 25          'NCTS Arrivals
    eLookInUnloadingRemarksSent = 26        'NCTS Arrivals
    eLookInUnloadingRemarksRejected = 27    'NCTS Arrivals
End Enum

Private Enum enuCondition
    eContains = 0
    eIsExactly = 1
    eDoesnotContain = 2
    eIsEmpty = 3
    eIsNotEmpty = 4
End Enum

Private Enum enuGroup
    eGroupMain = 0  '--> IMPORT, EXPORT, etc...
    eGroupHeader = 1    '--> IMPORT HEADER, EXPORT HEADER, TRANSIT HEADER, etc...
    eGroupDetail = 2    '--> IMPORT DETAIL, EXPORT DETAIL, etc.
End Enum

Private Const BOXESWITHINSTANCES As String = "S6|S7|S8|S9|SA|E4|E5|E6|E7|EM|EN|AL|AM"

Private enuDocType As enuDocType
Private enuLookIn As enuLookIn

Private Const IMPORTHEADER As String = "A1A2A3A4A5A6A7A8A9B1B2B3B4B5B6B7B8B9BABBC1C2C3C4C5C6D1D2D3D4D5D6D7D8D9DADBE1E2E3E4F1G1H1J1F2G2H2J2F3G3H3J3K1K2K3K4K5K6U1U2U3U4U5"
Private Const IMPORTDETAIL As String = "L1L2L3L4M1M2M3M4M5M6M7M8M9MAN1O1P1Q1N2O2P2Q2N3O3P3Q3R1R2R3R4R5R6R7R8R9RAS1T1T2T3T4T5T6T7"

Private Const EXPORTHEADER As String = "A1A2A3A4A5A6A7B1B2B3B4B5B6C1C2C3C4D1D2D3D4D5D6D7E1E2E3E4E5E6E7E8E9EAEBECEDEEEFEGEHEIEJF1G1H1J1F2G2H2J2F3G3H3J3K1K2K3K4K5K6U2U3U4U5W1W2W3X1X2X3"
Private Const EXPORTDETAIL As String = "L1L2L3L4L5L6M1M2M3M4M5M6M7M8N1O1P1Q1N2O2P2Q2N3O3P3Q3R1R2R3R4R5R6R7R8R9RAS1S2S3S4S5S6T1T2T3T4T5T6T7U1"

Private Const OTSHEADER As String = "A1A2A3A4A5A6A7B1B2B3B4B5B6C1C2C3C4D1D2D3D4D5D6D7E1E2E3E4E5E6E7E8E9EAEBECEDEEEFEGEHEIEJF1G1H1J1F2G2H2J2F3G3H3J3K1K2K3K4K5K6U2U3U4U5W1W2W3X1X2X3"
Private Const OTSDETAIL As String = "L1L2L3L4L5L6M1M2M3M4M5M6M7M8N1O1P1Q1N2O2P2Q2N3O3P3Q3R1R2R3R4R5R6R7R8R9RAS1S2S3S4S5S6T1T2T3T4T5T6T7U1"

Private Const NCTSHEADER As String = "A4A5A6A8A9AAABACADAEAFB7B1B8B2B3B9BAB5C2C3C4C5X4X5X1X2X6X3X7X8E1EJE3EKE4E5E6E7EMENEOE8EAECEEEGEI"
Private Const NCTSDETAIL As String = "U6U2U3U4U8U7W6W7W1W2W4W3W5L7L1L8M1M2M9S1S2S4S3S5S6S7S8S9SASBV1V2V3V4V5V6V7V8Y1Y2Y3Y4Y5Z1Z2Z3Z4T7"

Private Const ZEKERHEID As String = "E1EJE3EKE4E5E6E7EMENEO"
Private Const COLLI  As String = "S2S4S3S5"
Private Const Container As String = "S6S7S8S9SASB"
Private Const DOCUMENTEN As String = "Y1Y2Y3Y4Y5"
Private Const BIJZONDERE As String = "Z1Z2Z3Z4"
Private Const GEVEOLIGE As String = "V1V2V3V4V5V6V7V8"
Private Const GEODEREN As String = "L1L2L3L4L5L6L7L8M1M2M9M3M4M5S1"

Private Const COMBINEDHEADER  As String = "A1A2A4A5A6A7A8A9AAABACADAEAFB7B1B8B2B3B9B4BAB5B6C1C2C3C4C5D1D2D3D4D5D6D7F1G1H1J1F2G2H2J2F3G3H3J3K1K2K3K4K5K6X4X5X1X2X6X3X7X8E1EJE3EKE4E5E6E7EMENEOE8EAECEEEGEI"
Private Const COMBINEDDETAIL As String = "U6U7U2U3U4U8U5W6W7W1W2W4W3W5L1L2L3L4L5L6L8M1M2M9M3M4M5S1S2S4S3S5S6S7S8S9SASBV1V2V3V4V5V6V7V8Y1Y2Y3Y4Y5Z1Z2Z3Z4M6M7M8N1O1P1Q1N2O2P2Q2N3O3P3Q3R1R2R3R4R5R6R7R8R9RAT1T2T3T4T5T6T7"

Private Const PLDAIMPORTHEADER  As String = "A1A2A9ACA3A4A5A6A8AAABD1D2D3D4DAC1C7C2C3D5D7D8D9C4C5C6C9CACBB1B4B5B2B6XEX1XDX2X3X4X5X7X6X8X9XAXBXCE1E2E3XF"
Private Const PLDAIMPORTDETAIL As String = "L1L2L3L4L5L6L7LBL8L9LAN1N2N3NDNEN4NFNGNHN5N9N7N8NBS1S2S3S4S5S6M1M2M3M4M5O5O6OBO7O8OCO9OAODO1O2O3O4U1U2U3TZT8T9T1T2VEV1V2V3V4V5V7V6V8R1R2R3R5R6R8R9Q1Q2Q3Q4QBQCQ5Q7Q8Q9QAT3T4T5T6P1P2P5T7" 'joy 9/12/2006 changed TZ and T7 'joy 9/12/2006 changed T7 and TZ

Private Const PLDAIMPORTZEGELS As String = "E1E2E3"
Private Const PLDAIMPORTCONTAINER As String = "S4S5S6"
Private Const PLDAIMPORTBIJZONDERE As String = "P1P2P5"
Private Const PLDAIMPORTDOCUMENTEN As String = "Q1Q2Q3Q4Q5Q7Q8Q9QAQBQC"
Private Const PLDAIMPORTZELF As String = "U1U2U3"

'rachelle 082806 - for new PLDA Import codisheet design
Private Const PLDAIMPORTHEADERHANDELAARS As String = "X1X2X3X4X5X6X7X8X9XAXBXCXDXEXF"
Private Const PLDAIMPORTDETAILHANDELAARS As String = "V1V2V3V4V5V6V7V8VE"
Private Const PLDAIMPORTDETAILBEREKENINGS As String = "TZT8T9" 'joy 9/12/2006 changed T7 to TZ

Private Const PLDACOMBINEDHEADER  As String = "A1A2A9ACA3A4A5A6A7A8AAABD1D2D3D4DBC4C5C6C2C3C7D5D8D9D6D7E1E2E3XEX1XFXDX2X3X4X5X7X6X8X9XAXBXC"
Private Const PLDACOMBINEDDETAIL As String = "L1L2L3L4L5L6LBL8L9LAN1N2N3NDNEN4NFNGNHN9N7NBNCS1S2S3S4S5S6M1M2O2O6OBO3O4M3M4M5VEV1V2V3V4V5V7V6V8P1P2P5R1R2R3R5R6R8R9Q1Q2Q3Q4QBQCQ5Q7Q8Q9QAT7" 'joy 9/12/2006 changed TZ and T7

Private Const PLDACOMBINEDZEGELS As String = "E1E2E3"
Private Const PLDACOMBINEDHEADERHANDELAARS As String = "XEX1XFXDX2X3X4X5X7X6X8X9XAXBXC"
Private Const PLDACOMBINEDCONTAINER As String = "S4S5S6"
Private Const PLDACOMBINEDBIJZONDERE As String = "P1P2P5"
Private Const PLDACOMBINEDDOCUMENTEN As String = "Q1Q2Q3Q4QBQCQ5Q7Q8Q9QA"
Private Const PLDACOMBINEDDETAILHANDELAARS As String = "VEV1V2V3V4V5V7V6V8"

Private rstOfflineTemp As ADODB.Recordset

'Private WithEvents mMainCls As MainCls
Private mblnCancel As Boolean

Private intActiveSortKey As Integer
Private intActiveSortOrder As Integer

Private conFind As ADODB.Connection
Private strSettingWidth As String
'Private strSettingHeaders As String
'Private strSettingAlignment As String
Private strSort As String
Private strFormWidth As String
Private rstFind As ADODB.Recordset
Private blnShowFields As Boolean
Private blnJustLoaded As Boolean
Private blnJustLoaded2 As Boolean
Private blnJustLoaded3 As Boolean
Private strOldDocType As String
Private strView As String
Private strPosition As String
Private Const DOC_ANY = "**T7"
Private Const DOC_IMPORT = "**A1**A2**A3**A4**A5**A6**A7**A8**A9**B1**B2**B3**B4**B5**B6**B7**B8**B9**BA**BB**C1**C2**C3**C4**C5**C6**D1**D2**D3**D4**D5**D6**D7**D8**D9**DA**DB**E1**E2**E3**E4**F1**F2**F3**G1**G2**G3**H1**H2**H3**J1**J2**J3**K1**K2**K3**K4**K5**K6**L1**L2**L3**L4**M1**M2**M3**M4**M5**M6**M7**M8**M9**MA**N1**N2**N3**O1**O2**O3**P1**P2**P3**Q1**Q2**Q3**R1**R2**R3**R4**R5**R6**R7**R8**R9**RA**S1**T1**T2**T3**T4**T5**T6**T7**U1**U2**U3**U4**U5"
Private Const DOC_EXPORT = "**A1**A2**A3**A4**A5**A6**A7**B1**B2**B3**B4**B5**B6**C1**C2**C3**C4**D1**D2**D3**D4**D5**D6**D7**E1**E2**E3**E4**E5**E6**E7**E8**E9**EA**EB**EC**ED**EE**EF**EG**EH**EI**EJ**F1**F2**F3**G1**G2**G3**H1**H2**H3**J1**J2**J3**K1**K2**K3**K4**K5**K6**L1**L2**L3**L4**L5**L6**M1**M2**M3**M4**M5**M6**M7**M8**N1**N2**N3**O1**O2**O3**P1**P2**P3**Q1**Q2**Q3**R1**R2**R3**R4**R5**R6**R7**R8**R9**RA**S1**S2**S3**S4**S5**S6**T1**T2**T3**T4**T5**T6**T7**U1**U2**U3**U4**U5**W1**W2**W3**X1**X2**X3"
Private Const DOC_OTS = "**A1**A2**A3**A4**A5**A6**A7**B1**B2**B3**B4**B5**B6**C1**C2**C3**C4**D1**D2**D3**D4**D5**D6**D7**E1**E2**E3**E4**E5**E6**E7**E8**E9**EA**EB**EC**ED**EE**EF**EG**EH**EI**EJ**F1**F2**F3**G1**G2**G3**H1**H2**H3**J1**J2**J3**K1**K2**K3**K4**K5**K6**L1**L2**L3**L4**L5**L6**M1**M2**M3**M4**M5**M6**M7**M8**N1**N2**N3**O1**O2**O3**P1**P2**P3**Q1**Q2**Q3**R1**R2**R3**R4**R5**R6**R7**R8**R9**RA**S1**S2**S3**S4**S5**S6**T1**T2**T3**T4**T5**T6**T7**U1**U2**U3**U4**U5**W1**W2**W3**X1**X2**X3"
Private Const DOC_SADBELNCTS = "**A4**A5**A6**A8**A9**AA**AB**AC**AD**AE**AF**B1**B2**B3**B5**B7**B8**B9**BA**C2**C3**C4**C5**E1**E3**E4**E5**E6**E7**E8**EA**EC**EE**EG**EI**EJ**EK**EM**EN**EO**L1**L7**L8**M1**M2**M9**S1**S2**S3**S4**S5**S6**S7**S8**S9**SA**SB**T7**U2**U3**U4**U6**U7**U8**V1**V2**V3**V4**V5**V6**V7**V8**W1**W2**W3**W4**W5**W6**W7**X1**X2**X3**X4**X5**X6**X7**X8**Y1**Y2**Y3**Y4**Y5**Z1**Z2**Z3**Z4"
Private Const DOC_COMBINEDNCTS = "**A1**A2**A4**A5**A6**A7**A8**A9**AA**AB**AC**AD**AE**AF**B1**B2**B3**B4**B5**B6**B7**B8**B9**BA**C1**C2**C3**C4**C5**D1**D2**D3**D4**D5**D6**D7**E1**E3**E4**E5**E6**E7**E8**EA**EC**EE**EG**EI**EJ**EK**EM**EN**EO**F1**F2**F3**G1**G2**G3**H1**H2**H3**J1**J2**J3**K1**K2**K3**K4**K5**K6**L1**L2**L3**L4**L5**L6**L8**M1**M2**M3**M4**M5**M6**M7**M8**M9**N1**N2**N3**O1**O2**O3**P1**P2**P3**Q1**Q2**Q3**R1**R2**R3**R4**R5**R6**R7**R8**R9**RA**S1**S2**S3**S4**S5**S6**S7**S8**S9**SA**SB**T1**T2**T3**T4**T5**T6**T7**U2**U3**U4**U5**U6**U7**U8**V1**V2**V3**V4**V5**V6**V7**V8**W1**W2**W3**W4**W5**W6**W7**X1**X2**X3**X4**X5**X6**X7**X8**Y1**Y2**Y3**Y4**Y5**Z1**Z2**Z3**Z4"
Private Const DOC_EDINCTS = "**A4**A5**A6**A8**A9**AA**AB**AC**AD**AE**AF**B1**B2**B3**B5**B7**B8**B9**BA**C2**C3**C4**C5**E1**E3**E4**E5**E6**E7**E8**EA**EC**EE**EG**EI**EJ**EK**EM**EN**EO**L1**L7**L8**M1**M2**M9**S1**S2**S3**S4**S5**S6**S7**S8**S9**SA**SB**T7**U2**U3**U4**U6**U7**U8**V1**V2**V3**V4**V5**V6**V7**V8**W1**W2**W3**W4**W5**W6**W7**X1**X2**X3**X4**X5**X6**X7**X8**Y1**Y2**Y3**Y4**Y5**Z1**Z2**Z3**Z4"
Private Const DOC_EDINCTS2 = "**AG**AH**AJ**AK**AL**AM**BC**BD**BF**BG**BH**BI**BJ**BK**BL**C7**C8**C9**CA**CB**EP**EQ**ER**ES**MC**ME**MR**SB**SC**SD**SE**SF**SG**T7**W8**W9**WA**WB**WD**WE"

Private Const DOC_PLDAIMPORT = "**A1**A2**A9**AC**A3**A4**A5**A6**A8**AA**AB**D1**D2**D3**D4**DA**C1**C7**C2**C3**D5**D7**D8**D9**C4**C5**C6**C9**CA**CB**B1**B4**B5**B2**B6**XE**X1**XD**X2**X3**X4**X5**X7**X6**X8**X9**XA**XB**XC**E1**E2**E3**XF**L1**L2**L3**L4**L5**L6**L7**LB**L8**L9**LA**N1**N2**N3**ND**NE**N4**NF**NG**NH**N5**N9**N7**N8**NB**S1**S2**S3**S4**S5**S6**M1**M2**M3**M4**M5**O5**O6**OB**O7**O8**OC**O9**OA**OD**O1**O2**O3**O4**U1**U2**U3**TZ**T8**T9**T1**T2**VE**V1**V2**V3**V4**V5**V7**V6**V8**R1**R2**R3**R5**R6**R8**R9**Q1**Q2**Q3**Q4**QB**QC**Q5**Q7**Q8**Q9**QA**T3**T4**T5**T6**P1**P2**P5**T7" 'joy 9/12/2006 changed TZ and T7 'joy 9/12/2006 changed T7 to TZ
'Private Const DOC_PLDACOMBINED = "**A1**A2**A3**A4**A5**A6**A7**A8**A9**AA**AB**AC**AD**AE**AF**AG**AH**AJ**AK**C2**C3**C4**C5**C6**C7**C8**D1**D2**D3**D4**D5**D6**D7**D8**D9**DA**DB**DC**E1**E2**E3**E4**F1**F2**F3**F4**F5**F6**F7**F8**F9**FA**FB**FC**FD**FE**FF**G2**G3**G4**G5**G6**G7**G8**G9**GA**GB**GC**H1**H2**H3**H4**H5**H6**H7**H8**H9**HA**HB**HC**L1**L2**L3**L4**L5**L6**L8**L9**LA**LB**M1**M2**M3**M4**M5**N1**N2**N3**N4**N7**N9**NA**NB**NC**ND**NE**NF**NG**NH**O2**O3**O4**OE**OF**P1**P2**P3**P4**P5**Q1**Q2**Q3**Q4**Q5**Q6**Q7**Q8**Q9**R1**R2**R3**R5**R6**R8**R9**S1**S2**S3**S4**S5**S6**S7**S8**S9**SA**SB**SC**SD**SE**T7**V1**V2**V3**V4**V5**V6**V7**V8**V9**VA**VB**VC**W1**W2**W3**W4**W5**W6**W7**W8**W9**WA**WB**WC**X1**X2**X3**X4**X5**X6**X7**X8**X9**XA**XB**XC**XD**XE**Y1**Y2**Y3**Y4**Y5**Y6**Y7**Y8**Y9**YA**YB**YC**Z1**Z2**Z3**Z4**Z5**Z6**Z7"
Private Const DOC_PLDACOMBINED = "**A1**A2**A9**AC**A3**A4**A5**A6**A7**A8**AA**AB**D1**D2**D3**D4**DB**C4**C5**C6**C2**C3**C7**D5**D8**D9**D6**D7**E1**E2**E3**XE**X1**XF**XD**X2**X3**X4**X5**X7**X6**X8**X9**XA**XB**XC**L1**L2**L3**L4**L5**L6**LB**L8**L9**LA**N1**N2**N3**ND**NE**N4**NF**NG**NH**N9**N7**NB**NC**S1**S2**S3**S4**S5**S6**M1**M2**O2**O6**OB**O3**O4**M3**M4**M5**VE**V1**V2**V3**V4**V5**V7**V6**V8**P1**P2**P5**R1**R2**R3**R5**R6**R8**R9**Q1**Q2**Q3**Q4**QB**QC**Q5**Q7**Q8**Q9**QA**T7" 'joy 9/12/2006 changed TZ and T7
Private lngFieldCounter As Long
Private strPLDAFields As String
Private strSQLFields() As String

Private m_blnSearchInProgress As Boolean

Private m_objDataSourceProperties As CDataSourceProperties

Public Function ShowForm(ByRef OwnerForm As Object, ByVal Application As Object, _
                         ByRef DataSourceProperties As CDataSourceProperties)
                         
    Set m_objDataSourceProperties = DataSourceProperties
    
    Set Me.Icon = OwnerForm.Icon
    Me.Show vbModal
End Function

Private Sub InitListView(ByVal strFieldsToSetWidth As String)
'    Dim strSettingHeaders2() As String
    Dim lngCounter As Long
    Dim lngCtr As Long
    Dim strTempPosition() As String
    Dim strSettingPosition() As String

    '====== Prepare Column Name header ==========
    lngFieldCounter = 0
    lvwItemsFound.ColumnHeaders.Clear
    lvwItemsFound.ColumnHeaders.Add , , Translate(496)
    lvwItemsFound.ColumnHeaders.Add , , Translate(635)
    lvwItemsFound.ColumnHeaders.Add , , Translate(625)
    lvwItemsFound.ColumnHeaders.Add , , Left(Translate(272), Len(Translate(272)) - 1) '626 userno
    lvwItemsFound.ColumnHeaders.Add , , Translate(611)
    lvwItemsFound.ColumnHeaders.Add , , ""
    
    If Not (Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocAny Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocNCTS Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocPLDAImport Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocPLDACombined) Then
        lvwItemsFound.ColumnHeaders.Add , , Translate(437)
        lngFieldCounter = lngFieldCounter + 1
    End If
    
    lvwItemsFound.ColumnHeaders.Add , , Translate(713)
    lvwItemsFound.ColumnHeaders.Add , , Translate(715)
    lvwItemsFound.ColumnHeaders.Add , , Translate(742)
    
    If Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Then
        lvwItemsFound.ColumnHeaders.Add , , "Date Last Received"
        lngFieldCounter = lngFieldCounter + 1
    End If
    
    lvwItemsFound.ColumnHeaders.Add , , "Date Printed"
    lvwItemsFound.ColumnHeaders.Add , , "LogID Description"
    lvwItemsFound.ColumnHeaders.Add , , "Error String"
    lvwItemsFound.ColumnHeaders.Add , , Translate(423)
    
    If Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocNCTS Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.edoccombined Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Or _
        Right(icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Then
        lvwItemsFound.ColumnHeaders.Add , , "MRN"
        lngFieldCounter = lngFieldCounter + 1
    End If
    '============================================
    
    For lngCounter = 14 + lngFieldCounter To rstOfflineTemp.Fields.Count - 1
        lvwItemsFound.ColumnHeaders.Add lngCounter, , rstOfflineTemp.Fields(lngCounter).Name
    Next lngCounter
    
    If strPosition <> "" Then
        strSettingPosition = Split(strPosition, "*****")
        ReDim strTempPosition(UBound(strSettingPosition), 0 To 1)
        
        For lngCtr = 0 To UBound(strSettingPosition)
            For lngCounter = 0 To UBound(strSettingPosition)
                If strSettingPosition(lngCounter) = lngCtr + 1 Then
                    strTempPosition(lngCtr, 0) = lngCounter + 1
                    strTempPosition(lngCtr, 1) = lngCtr + 1
                    Exit For
                End If
            Next lngCounter
        Next lngCtr
        
        For lngCounter = 0 To UBound(strTempPosition)
            lvwItemsFound.ColumnHeaders(CLng(strTempPosition(lngCounter, 0))).Position = strTempPosition(lngCounter, 1)
        Next lngCounter
    End If
    
    If strFieldsToSetWidth <> "" Then
        SetListViewWidths strFieldsToSetWidth
    Else
        SetListViewWidths ""
    End If
End Sub

Private Sub SetListViewWidths(ByVal strFieldsToSet As String)
    Dim strSettingWidth2() As String
    Dim strFieldsToSetWidth() As String
    Dim lngCounter As Long
    Dim lngCounter2 As Long
    
    For lngCounter = 7 To rstOfflineTemp.Fields.Count - 1
        lvwItemsFound.ColumnHeaders(lngCounter).Width = 0
    Next lngCounter
    
    If strSettingWidth <> "" Then
        If strFieldsToSet = "" Then
            strSettingWidth2 = Split(strSettingWidth, "*****")
            For lngCounter = 0 To UBound(strSettingWidth2)
                'CSCLP-248
                'Added Val() since passed property value uses comman (,) as decimal resulting in error - possibly regional settings
                lvwItemsFound.ColumnHeaders(lngCounter + 1).Width = Val(strSettingWidth2(lngCounter))
                'lvwItemsFound.ColumnHeaders(lngCounter + 1).Width = strSettingWidth2(lngCounter)
            Next lngCounter
        Else
            strFieldsToSetWidth = Split(strFieldsToSet, "*")
            For lngCounter = 6 To UBound(strFieldsToSetWidth)
                For lngCounter2 = 7 To lvwItemsFound.ColumnHeaders.Count
                    If lvwItemsFound.ColumnHeaders(lngCounter2).Text = strFieldsToSetWidth(lngCounter) Then
                        lvwItemsFound.ColumnHeaders(lngCounter2).Width = 1440
                        Exit For
                    End If
                Next lngCounter2
            Next lngCounter
        End If
    Else
        If strFieldsToSet = "" Then
            'Default
            For lngCounter = 1 To lvwItemsFound.ColumnHeaders.Count
                If lngCounter = 6 Then
                    lvwItemsFound.ColumnHeaders(lngCounter).Width = 0
                ElseIf lngCounter < 6 Then
                    lvwItemsFound.ColumnHeaders(lngCounter).Width = 1440
                ElseIf lngCounter > 6 Then
                    lvwItemsFound.ColumnHeaders(lngCounter).Width = 0
                End If
            Next lngCounter
        Else
            strFieldsToSetWidth = Split(strFieldsToSet, "*")
            For lngCounter = 6 To UBound(strFieldsToSetWidth)
                For lngCounter2 = 7 To lvwItemsFound.ColumnHeaders.Count
                    If lvwItemsFound.ColumnHeaders(lngCounter2).Text = strFieldsToSetWidth(lngCounter) Then
                        lvwItemsFound.ColumnHeaders(lngCounter2).Width = 1440
                        Exit For
                    End If
                Next lngCounter2
            Next lngCounter
        End If
    End If
    
End Sub

Private Sub InitControls()
    Dim strDocName As String
    Dim strHistoryDBFile As String
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    Dim strYear As String
    Dim lngCtr As Long

    ' ======== Prepare previously seached doc name ================
    strDocName = Trim(GetSetting(AppTitle, "Settings", "FindString"))
    j = CountChr(strDocName, "~^~")
    For i = 1 To j
        k = InStr(1, strDocName, "~^~")

        cboName.AddItem (Mid(strDocName, 1, k - 1))
        strDocName = Right(strDocName, Len(strDocName) - (k + 2))
    Next
    '==============================================================


    ' =================== fill doc type ================================
    icbType.ComboItems.Clear
    icbType.ComboItems.Add , "D" & enuDocType.eDocAny, "Any"   '--> translate
    icbType.ComboItems.Add , "D" & enuDocType.edocimport, "Import", GetImageIndex(1)
    icbType.ComboItems.Add , "D" & enuDocType.eDocExport, "Export", GetImageIndex(2)
    icbType.ComboItems.Add , "D" & enuDocType.eDocOTS, "OTS", GetImageIndex(3)
    icbType.ComboItems.Add , "D" & enuDocType.eDocNCTS, "Sadbel NCTS", GetImageIndex(7)
    icbType.ComboItems.Add , "D" & enuDocType.edoccombined, "Combined NCTS", GetImageIndex(9)
    icbType.ComboItems.Add , "D" & enuDocType.eDocEDIDepartures, "EDI Departures", GetImageIndex(11)
    icbType.ComboItems.Add , "D" & enuDocType.eDocEDIARRIVALS, "EDI Arrivals", GetImageIndex(12)
    icbType.ComboItems.Add , "D" & enuDocType.eDocPLDAImport, "PLDA Import", GetImageIndex(14)
    'CSCLP-248
    icbType.ComboItems.Add , "D" & enuDocType.eDocPLDACombined, "PLDA Export", GetImageIndex(18)
    'icbType.ComboItems.Add , "D" & enuDocType.eDocPLDACombined, "PLDA Combined", GetImageIndex(18)
    ' ====================================================================
    
    icbType.ComboItems.Item("D" & enuDocType.eDocAny).Selected = True
    
    'fill Search In
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInAll, AppTitle, imgImages.ListImages.Item("ClearingPoint").Index
    
    icbLookIn.ComboItems.Add , , LISTSEPARATOR
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInTemplates, Translate(347), imgImages.ListImages.Item("Templates").Index   '"Templates"
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInApproved, "Approved Documents", imgImages.ListImages.Item("Approved").Index
    
    If Dir(cAppPath & "\mdb_history" & Right(Year(Now), 2) & ".mdb") <> "" Or Dir(cAppPath & "\mdb_EDIhistory" & Right(Year(Now), 2) & ".mdb") <> "" Then
        strYear = Year(Now)
        'icbLookIn.ComboItems.Add , "D" & strYear, Translate(1074) & Space(4) & Year(Now), imgImages.ListImages.Item("Archive").Index
        icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInArchive, Translate(1074), imgImages.ListImages.Item("Archive").Index 'allan nov27 to remove Archive 2007
    End If
    
    'Drafts
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInDRAFTS, Translate(969), imgImages.ListImages.Item("Drafts").Index
    
    icbLookIn.ComboItems.Add , , LISTSEPARATOR
    
    'Outbox
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInOutbox, Translate(970), imgImages.ListImages.Item("Outbox").Index
    
    'Rejected
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInRejected, Translate(968), imgImages.ListImages.Item("Rejected").Index
    
    'Deleted
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInDeleted, Translate(348), imgImages.ListImages.Item("Deleted").Index
    
        
    'CSCLP-248
    icbLookIn.ComboItems.Add , , LISTSEPARATOR
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInToBePrinted, Translate(902)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInReleased, Translate(2388)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInSent, Translate(1386)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInEmergencyProcedure, "Emergency Procedure"
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInExitEC, Translate(2462)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInGuarantee, Translate(2309)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInUnderControl, Translate(1375)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInReleases, Translate(1376)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInReleaseRejected, Translate(1377)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInCancelled, Translate(1378)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInWrittenOff, Translate(1379)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInAmendmentSent, Translate(2171)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInAmendmentAccepted, Translate(2172)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInAmendmentRejected, Translate(2173)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInArchives, Translate(757)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInArrivalNotificationSent, Translate(1381)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInArrivalNotificationRejected, Translate(1382)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInUnloadingPermitted, Translate(1383)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInUnloadingRemarksSent, Translate(1384)
    icbLookIn.ComboItems.Add , "D" & enuLookIn.eLookInUnloadingRemarksRejected, Translate(1385)
    icbLookIn.ComboItems.Add , , LISTSEPARATOR
    
    '=========== For History =========================================
    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")

    Do Until strHistoryDBFile = ""

        'If strHistoryDBFile <> "" And UCase(Right(strHistoryDBFile, 6)) <> UCase(Right(Year(Now), 2) & ".mdb") Then
        If strHistoryDBFile <> "" Then  'allan nov27 to put archive 2007
            If Left(Right(strHistoryDBFile, 6), 2) <= "50" Then
                strYear = "20" & Left(Right(strHistoryDBFile, 6), 2)
            Else
                strYear = "19" & Left(Right(strHistoryDBFile, 6), 2)
            End If
            lstSorter.AddItem strYear
        End If
        strHistoryDBFile = Dir()    ' Get next History DB file.
    Loop
    
    strHistoryDBFile = Dir(cAppPath & "\mdb_EDIHistory*.mdb")

    Do Until strHistoryDBFile = ""

        If strHistoryDBFile <> "" And UCase(Right(strHistoryDBFile, 6)) <> UCase(Right(Year(Now), 2) & ".mdb") And UCase(strHistoryDBFile) <> "MDB_EDIHISTORY.MDB" Then
            If Left(Right(strHistoryDBFile, 6), 2) <= "50" Then
                strYear = "20" & Left(Right(strHistoryDBFile, 6), 2)
            Else
                strYear = "19" & Left(Right(strHistoryDBFile, 6), 2)
            End If
            For lngCtr = 0 To lstSorter.ListCount - 1
                If UCase(lstSorter.List(lngCtr)) = UCase(strYear) Then
                    Exit For
                End If
            Next
            If lngCtr >= lstSorter.ListCount Then
                lstSorter.AddItem strYear
            End If
        End If
        strHistoryDBFile = Dir()    ' Get next History DB file.
    Loop
    
    For i = lstSorter.ListCount To 1 Step -1
        icbLookIn.ComboItems.Add , "D" & lstSorter.List(i - 1), Translate(1074) & Space(4) & lstSorter.List(i - 1), imgImages.ListImages.Item("Archive").Index
        
    Next
    
    
    '==================================================================

    icbLookIn.ComboItems.Item("D" & enuLookIn.eLookInAll).Selected = True
    
    
    ' ============ Conditions ================
    UpdateConditionList 0
    '=========================================

    
    Combo5(1).Clear
    Combo5(1).AddItem (Trim(Translate(714)))
    Combo5(1).AddItem (Trim(Translate(713)))
    Combo5(1).AddItem (Trim(Translate(742)))
    Combo5(1).ListIndex = 0
    Combo5(1).BackColor = vbButtonFace    'RGB(192, 192, 192)
    
    '====== Initialize DatePickers =======
    DTPicker1(0).Value = Date
    DTPicker1(1).Value = Date
    '======================================
        
    InitListView ""
    
    Me.Height = 3550
    lvwItemsFound.Visible = False

End Sub

Private Sub cboName_LostFocus()
    'Added by BCo 2006-08-31
    'Auto-complete
    cboName.SelLength = 0
End Sub

Private Sub cboName_KeyPress(KeyAscii As Integer)
    'Added by BCo 2006-08-31
    'Auto-complete
    Dim strSearchText As String
    Dim strEnteredText As String
    Dim lngLength As Long
    Dim lngIndex As Long
    Dim lngCounter As Long
    
    On Error GoTo ErrorHandler
    If cboName.ListCount = 0 Then Exit Sub
    With cboName
        If .SelStart > 0 Then strEnteredText = Left$(.Text, .SelStart)

        Select Case KeyAscii
            Case vbKeyReturn
                If .ListIndex > -1 Then
                    .SelStart = 0
                    .SelLength = Len(.List(.ListIndex))
                    Exit Sub
                End If
            Case vbKeyEscape
                .Text = Empty
                Exit Sub
            Case vbKeyBack
                If Len(strEnteredText) > 1 Then
                    strSearchText = LCase$(Left$(strEnteredText, Len(strEnteredText) - 1))
                Else
                    strEnteredText = Empty
                    KeyAscii = 0
                    .Text = Empty
                End If
                Exit Sub
            Case Else
                strSearchText = LCase$(strEnteredText & Chr(KeyAscii))
        End Select
    
        lngIndex = -1
        lngLength = Len(strSearchText)
    
        For lngCounter = 0 To .ListCount - 1
            If LCase$(Left$(.List(lngCounter), lngLength)) = strSearchText Then
                lngIndex = lngCounter
                Exit For
            End If
        Next lngCounter
    
        If lngIndex > -1 Then
            .ListIndex = lngIndex
            .SelStart = Len(strSearchText)
            .SelLength = Len(.List(lngIndex)) - Len(strSearchText)
        KeyAscii = 0
        Else
            'Beep
        End If
    End With
    Exit Sub
    
ErrorHandler:
    KeyAscii = 0
    Beep
End Sub

Private Sub cmdFind_Click()
    Dim strFormWidth2() As String
    
    If CriteriaNotSpecified Then Exit Sub

    'Added by BCo 2006-08-31 (temporary since there's partial code that uses registry for persistent logging)
    'Adds names used to cboName list, except spaces or blanks
    Dim lngCounter As Long
    Dim bytMatch As Byte
    If Len(Trim$(cboName.Text)) > 0 Then                                        'No process on dud criteria
        If cboName.ListCount > 0 Then                                           'No process on empty list
            For lngCounter = 1 To cboName.ListCount                             'Check combo list
                If Len(cboName.Text) = Len(cboName.List(lngCounter - 1)) Then   'Length comparison first for performance
                    If cboName.Text = cboName.List(lngCounter - 1) Then         'String comparison for accuracy
                        bytMatch = 1                                            'Found!
                        Exit For                                                'No further processing needed
                    End If
                End If
            Next
            If bytMatch = 0 Then cboName.AddItem (cboName.Text)                 'Only add if not found
        Else
            cboName.AddItem (cboName.Text)                                      'Only add if combo is empty
        End If
    End If

    lvwItemsFound.ListItems.Clear

    cmdStop.Enabled = True
    
    Me.StatusBar1.Visible = True
    cmdFind.Enabled = False
    cmdStop.Enabled = True
    cmdNewSearch.Enabled = False 'allan dec13
    
    icbType.Enabled = False
    icbLookIn.Enabled = False
    cboName.Enabled = False
    
    If Me.WindowState = 0 Then If Not lvwItemsFound.Visible Then Me.Height = 6260
    
    lvwItemsFound.Visible = True
    
    If Me.WindowState = 2 And lvwItemsFound.Visible = True Then 'allan dec4

        With lvwItemsFound
          .Top = SSTab1.Top + SSTab1.Height + 108
          '.Height = me.h Me.StatusBar1.Top - (Me.SSTab1.Top + Me.SSTab1.Height + 108 + 110)  ' last 100 bottom margin
          .Width = Me.Width - 100
          .Height = Me.Height - (.Top + StatusBar1.Height + 850)
          
        End With

    End If
    
    If strFormWidth <> "" And Not Me.WindowState = 2 Then 'allan dec4
        strFormWidth2 = Split(strFormWidth, "*****")
        Me.Width = Val(strFormWidth2(0))
        Me.Height = Val(strFormWidth2(1))
        Me.Top = Val(strFormWidth2(2))
        Me.Left = Val(strFormWidth2(3))
        strFormWidth = ""
    End If
    
    CreateOfflineRecordset
    
    'CSCLP-248
    If chkAllFields.Value = vbUnchecked Then
        '5 default fields
        strFields = "*Name*Document*In folder*User Name*Date Modified*"
        'Date tab
        If optFind.Value = True Then
            strFields = strFields & "*" & Combo5(1).Text
        End If
        'Advanced tab
        If Len(icbBox.Text) > 0 Then
            strFields = strFields & "*" & Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)
        End If
        
        SetListViewWidths strFields
    End If
    
    m_blnSearchInProgress = True
    BeginSearch
    m_blnSearchInProgress = False
    
    If Len(Dir(cAppPath & "\bmp\findfile.avi")) > 0 Then
        Animation1.Stop
    End If
    
    cmdFind.Enabled = True
    cmdStop.Enabled = False
    cmdNewSearch.Enabled = True 'allan dec4
    
    icbType.Enabled = True
    icbLookIn.Enabled = True
    cboName.Enabled = True
    
    Me.StatusBar1.Panels(1).Text = Trim(CStr(lvwItemsFound.ListItems.Count)) + " file(s) found"
End Sub

Private Sub cmdNewSearch_Click()
    
    'Modified by Mugs on 08-23-2006. Placed If-Then statement to check the number of list items in the list view.
    If lvwItemsFound.ListItems.Count > 0 Then
        If MsgBox(Translate(700), vbInformation + vbOKCancel, Translate(827)) = 1 Then
            Me.WindowState = 0
            If lvwItemsFound.Visible Then lvwItemsFound.Visible = False
            Me.StatusBar1.Visible = False
            Me.Height = 3550
            lvwItemsFound.ListItems.Clear
            
            cmdStop_Click
            cboName.Text = vbNullString
            cboName.Enabled = True
            cboName.SetFocus
            
        End If
    Else
        Me.WindowState = 0
        If lvwItemsFound.Visible Then lvwItemsFound.Visible = False
        Me.StatusBar1.Visible = False
        Me.Height = 3550
    End If
    
    m_blnSearchInProgress = False
End Sub

Private Sub cmdStop_Click()
    mblnCancel = True
    m_blnSearchInProgress = False
    If lvwItemsFound.ListItems.Count = 0 Then
        SSActiveToolBars1.Tools("ID_Open").Enabled = False
        SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = False
    Else
        SSActiveToolBars1.Tools("ID_Open").Enabled = True
        SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = True
    End If
End Sub

Private Sub Form_Load()
    LoadResStrings Me, True
    
    m_blnSearchInProgress = False
    
    blnShowFields = False
    blnJustLoaded = True
    blnJustLoaded2 = True
    blnJustLoaded3 = True
    strOldDocType = "Any"
    strView = ""
    
    '#####
    ConnectToDB m_objDataSourceProperties
    
    '#####
    LoadListView
    
    '#####
    CreateOfflineRecordset
    
    InitControls
    
    optAll_Click
    
    clsFindForm.OpenOnly = False
    SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = SSActiveToolBars1.Tools("ID_Open").Enabled
    
    'Added by BCo 2006-05-03
    'Licensing
    With SSActiveToolBars1
        SSActiveToolBars1.Tools("ID_Import").Enabled = clsFindForm.LicSADI
        SSActiveToolBars1.Tools("ID_Export").Enabled = clsFindForm.LicSADET
        SSActiveToolBars1.Tools("ID_Transit").Enabled = clsFindForm.LicSADET
        SSActiveToolBars1.Tools("ID_NCTS").Enabled = clsFindForm.LicSADTC
        SSActiveToolBars1.Tools("ID_Combined").Enabled = clsFindForm.LicSADTC
        SSActiveToolBars1.Tools("ID_EDIDepartures").Enabled = clsFindForm.LicNCTS
        SSActiveToolBars1.Tools("ID_EDIArrivals").Enabled = clsFindForm.LicNCTS
        SSActiveToolBars1.Tools("ID_PLDAImport").Enabled = clsFindForm.LicPLDAI
        SSActiveToolBars1.Tools("ID_PLDACombined").Enabled = clsFindForm.LicPLDAC
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngCounter As Long
    
    'Added by BCo 2006-09-01
    'Proper saving of cboName list to registry for persistent logging.
    Dim strFindString As String
    For lngCounter = 1 To cboName.ListCount
        strFindString = cboName.List(lngCounter - 1) & "~^~" & strFindString
    Next
    SaveSetting AppTitle, "Settings", "FindString", strFindString
 
    SaveSpecs icbType.SelectedItem, True
    
    ADODisconnectDB rstOfflineTemp
    ADODisconnectDB rstFind

    ADODisconnectDB g_conSADBEL
    ADODisconnectDB g_conData
    ADODisconnectDB g_conEDIFACT
    ADODisconnectDB g_conTemplate
    
    If blnEDIHistoryExisting = True Then
        For lngCounter = 0 To UBound(g_conEDIHistory)
            ADODisconnectDB g_conEDIHistory(lngCounter)
        Next lngCounter
    End If
    
    Set CallingForm.clsFind = Nothing
End Sub

Private Sub icbType_Change()
    Dim blnLVW As Boolean

    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny Then
        icbBox.Enabled = False
        UpdateConditionList 0
    Else
        icbBox.Enabled = True
        UpdateConditionList 1
        UpdateBoxList
    End If

    If strOldDocType <> icbType.SelectedItem Then
        blnLVW = lvwItemsFound.Visible
        If blnLVW = True Then
            lvwItemsFound.Visible = False
            Me.MousePointer = vbHourglass
        End If
    
        SaveSpecs strOldDocType
        strOldDocType = icbType.SelectedItem
        lvwItemsFound.ListItems.Clear
        Me.StatusBar1.Panels(1).Text = ""
        
        LoadListView
        CreateOfflineRecordset
        InitListView ""
    
        If blnLVW = True Then
            lvwItemsFound.Visible = True
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub BeginSearch()
    Dim strHistoryDBFile As String
    Dim strYear As String
    Dim lngCounter As Long
    
    mblnCancel = False
    
    If Len(Dir(cAppPath & "\bmp\findfile.avi")) > 0 Then
        Animation1.Open cAppPath & "\bmp\findfile.avi"
        Animation1.Play
    End If
    
    StatusBar1.Visible = True
    
    'Call mMainCls.LongTask(0.01, 0.01)   '14.4, 0.66
    
    If mblnCancel = True Then Exit Sub
    
    Select Case icbType.SelectedItem.Key
        Case "D" & enuDocType.eDocAny
            Select Case icbLookIn.SelectedItem.Key
                Case "D" & enuLookIn.eLookInAll
                    'List of folders
                    '1. Approved and Printed/ TobePrinted
                    '2. Archive
                    '3. Deleted
                    '4. Drafts
                    '5. Outbox
                    '6. Rejected
                    '7. Templates
                    
                    '==== Approved and Printed/Tobe Printed, Deleted, Drafts, Outbox, Rejected ===='
                    
                    '#####
                    SearchIn "IMPORT", 1, "", "SDI1", "SDI2", "DD", "WL1", "WL2", "DE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "EXPORT", 2, "", "SDE1", "SDE2", "DD", "WL1", "WL2", "DE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "TRANSIT", 3, "", "SDT1", "SDT2", "DD", "WL1", "WL2", "DE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "NCTS", 7, "", "SDN1", "SDN2", "DD", "WL1", "WL2", "DE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "COMBINED NCTS", 9, "", "SDC1", "SDC2", "DD", "WL1", "WL2", "DE"
                    If mblnCancel = True Then Exit Sub

                    SearchInEDI 11, 5, "", "34ED", "35ED", "36ED", "-1ED", "30ED", "31ED", "32ED", "33ED", "37ED", "39ED"

                    SearchInEDI 12, 2, "", "47ED", "48ED", "-2ED", "43ED", "44ED", "45ED", "46ED", "49ED"
                    SearchInEDI 12, 11, "", "47ED", "48ED", "-2ED", "43ED", "44ED", "45ED", "46ED", "49ED"

                    SearchIn "PLDA IMPORT", 14, "", "SDX1", "SDX2", "SDX3", "DD", "WL1", "WL2", "DE", "PX01"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "PLDA COMBINED", 18, "", "SDZ1", "SDZ2", "SDZ3", "DD", "WL1", "WL2", "DE", "PZ01"
                    If mblnCancel = True Then Exit Sub

                    '=============================================================================='


                    '==== Archive ===='
                    'Added by BCo 2006-08-31
                    'Search in all available mdb_history DBs
                    Dim arrHistories() As String
                    
                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")                  'Get first history filename
                    Do
                        strHistoryDBFile = strHistoryDBFile & "," & Dir()                   'CSV history filename
                    Loop Until Right$(strHistoryDBFile, 1) = ","                            'Process until Dir() is empty
                    strHistoryDBFile = Mid$(strHistoryDBFile, 1, Len(strHistoryDBFile) - 1) 'Erase marker used with Dir()
                    arrHistories = Split(strHistoryDBFile, ",")                             'Partition results
                    
                    For lngCounter = LBound(arrHistories) To UBound(arrHistories)           'Span across array
                        strYear = Mid$(arrHistories(lngCounter), 12, 2)                     'Get 2-digit year from history filename
'Original code
'--------------
''                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")
''                    Do Until strHistoryDBFile = ""
''                        strHistoryDBFile = Dir()    ' Get next History DB file.
''
''                        If strHistoryDBFile <> "" Then
''                            strYear = Left(Right(strHistoryDBFile, 6), 2)
                            SearchIn "IMPORT", 1, strYear, "HI" & strYear & "I"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "EXPORT", 2, strYear, "HI" & strYear & "E"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "TRANSIT", 3, strYear, "HI" & strYear & "T"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "NCTS", 7, strYear, "HI" & strYear & "N"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "COMBINED NCTS", 9, strYear, "HI" & strYear & "C"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "PLDA IMPORT", 14, strYear, "HI" & strYear & "X"
                            If mblnCancel = True Then Exit Sub

                            SearchIn "PLDA COMBINED", 18, strYear, "HI" & strYear & "Z"
                            If mblnCancel = True Then Exit Sub
''                        End If
''                    Loop
'--------------
                    Next
                    
                    'For EDI Archive
                    'Added by Rachelle on 06/16/2005
                    
                    'For Archive that are not in EDI History dbs..
                    '#####
                    SearchInEDI 11, 5, "", "40ED"
                    SearchInEDI 12, 2, "", "50ED"
                    SearchInEDI 12, 11, "", "50ED"
                    
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            If IsNumeric(Mid(File1.List(lngCounter), 15, 2)) And _
                                Right(File1.List(lngCounter), 3) = "mdb" Then
                                If Mid(File1.List(lngCounter), 15, 2) <> Right(Year(Now), 2) Then
                                    SearchInEDI 11, 5, Mid(File1.List(lngCounter), 15, 2), "40ED"
                                    SearchInEDI 12, 2, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                    SearchInEDI 12, 11, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                End If
                            End If
                        Next lngCounter
                    End If
                    '======================================================'

                    '==== Templates ======================================='
                    SearchInEDI 11, 5, "", "41ED"
                    If mblnCancel = True Then Exit Sub

                    SearchInEDI 12, 2, "", "51ED"
                    If mblnCancel = True Then Exit Sub

                    SearchInEDI 12, 11, "", "51ED"
                    If mblnCancel = True Then Exit Sub


                    SearchIn "IMPORT", 1, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "EXPORT", 2, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "TRANSIT", 3, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "NCTS", 7, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "COMBINED NCTS", 9, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "PLDA IMPORT", 14, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    SearchIn "PLDA COMBINED", 18, "", "TE"
                    If mblnCancel = True Then Exit Sub

                    '========== for template subfolders ====================
                    '#####
                    SearchInSubfolders "MASTEREDINCTS", 11
                    SearchInSubfolders "MASTEREDINCTS2", 12
                    SearchInSubfolders "MASTEREDINCTSIE44", 12
                    SearchInSubfolders "MASTER", 1, 2, 3
                    SearchInSubfolders "MASTERNCTS", 7, 9
                    SearchInSubfolders "MASTERPLDA", 14, 18
                    '=======================================================

                    SearchInSubfolders "MASTER", 1, 2, 3
                    If mblnCancel = True Then Exit Sub

                    SearchInSubfolders "MASTERNCTS", 7, 9
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInSubfolders "MASTERPLDA", 14, 18
                    If mblnCancel = True Then Exit Sub
                    '======================================================'
                    
                Case "D" & enuLookIn.eLookInApproved
                    SearchInEDI 11, 5, "", "34ED", "35ED", "36ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "46ED", "47ED", "49ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 11, "", "46ED", "47ED", "49ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "IMPORT", 1, "", "SDI1", "SDI2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "SDE1", "SDE2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "SDT1", "SDT2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "SDN1", "SDN2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "SDC1", "SDC2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA IMPORT", 14, "", "SDX1", "SDX2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "SDZ1", "SDZ2"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInOutbox
                    SearchInEDI 11, 5, "", "31ED", "32ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "44ED", "45ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "44ED", "45ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchIn "IMPORT", 1, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                                
                    SearchIn "PLDA IMPORT", 14, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                                
                    SearchIn "PLDA COMBINED", 18, "", "WL2"
                    If mblnCancel = True Then Exit Sub
                                
                Case "D" & enuLookIn.eLookInDRAFTS
                    SearchInEDI 11, 5, "", "30ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "43ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "43ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchIn "IMPORT", 1, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                
                    SearchIn "PLDA IMPORT", 14, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "WL1"
                    If mblnCancel = True Then Exit Sub
                
                
                Case "D" & enuLookIn.eLookInRejected
                    SearchInEDI 11, 5, "", "33ED", "37ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "49ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "49ED"
                    If mblnCancel = True Then Exit Sub
                
                
                    SearchIn "IMPORT", 1, "", "DE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "DE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "DE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "DE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "DE"
                    If mblnCancel = True Then Exit Sub
                
                    SearchIn "PLDA IMPORT", 14, "", "DE"
                    If mblnCancel = True Then Exit Sub
                
                    SearchIn "PLDA COMBINED", 18, "", "DE"
                    If mblnCancel = True Then Exit Sub
                
                
                Case "D" & enuLookIn.eLookInDeleted
                    SearchInEDI 11, 5, "", "-1ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "-2ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "-2ED"
                    If mblnCancel = True Then Exit Sub
                
                
                    SearchIn "IMPORT", 1, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA IMPORT", 14, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "DD"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInTemplates
                    SearchInEDI 11, 5, "", "41ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInEDI 12, 2, "", "51ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "51ED"
                    If mblnCancel = True Then Exit Sub
                
                
                    SearchIn "IMPORT", 1, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA IMPORT", 14, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "TE"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInSubfolders "MASTER", 1, 2, 3
                    If mblnCancel = True Then Exit Sub
                    
                    SearchInSubfolders "MASTERNCTS", 7, 9
                    If mblnCancel = True Then Exit Sub
                
                    SearchInSubfolders "MASTERPLDA", 14, 18
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInArchive 'allan nov27
                
                    '==== Archive ===='
                    'Added by BCo 2006-08-31
                    'Search in all available mdb_history DBs
                                        
                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")                  'Get first history filename
                    Do
                        strHistoryDBFile = strHistoryDBFile & "," & Dir()                   'CSV history filename
                    Loop Until Right$(strHistoryDBFile, 1) = ","                            'Process until Dir() is empty
                    strHistoryDBFile = Mid$(strHistoryDBFile, 1, Len(strHistoryDBFile) - 1) 'Erase marker used with Dir()
                    arrHistories = Split(strHistoryDBFile, ",")                             'Partition results
                    
                    For lngCounter = LBound(arrHistories) To UBound(arrHistories)           'Span across array
                        strYear = Mid$(arrHistories(lngCounter), 12, 2)                     'Get 2-digit year from history filename

                            SearchIn "IMPORT", 1, strYear, "HI" & strYear & "I"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "EXPORT", 2, strYear, "HI" & strYear & "E"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "TRANSIT", 3, strYear, "HI" & strYear & "T"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "NCTS", 7, strYear, "HI" & strYear & "N"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "COMBINED NCTS", 9, strYear, "HI" & strYear & "C"
                            If mblnCancel = True Then Exit Sub
                            
                            SearchIn "PLDA IMPORT", 14, strYear, "HI" & strYear & "X"
                            If mblnCancel = True Then Exit Sub

                            SearchIn "PLDA COMBINED", 18, strYear, "HI" & strYear & "Z"
                            If mblnCancel = True Then Exit Sub

                    Next
                    
                    SearchInEDI 11, 5, "", "40ED"
                    SearchInEDI 12, 2, "", "50ED"
                    SearchInEDI 12, 11, "", "50ED"
                    
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            If IsNumeric(Mid(File1.List(lngCounter), 15, 2)) And _
                                Right(File1.List(lngCounter), 3) = "mdb" Then
                                If Mid(File1.List(lngCounter), 15, 2) <> Right(Year(Now), 2) Then
                                    SearchInEDI 11, 5, Mid(File1.List(lngCounter), 15, 2), "40ED"
                                    SearchInEDI 12, 2, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                    SearchInEDI 12, 11, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                End If
                            End If
                        Next lngCounter
                    End If
                
                'CSCLP-248
                '-------------------------------------------------
                Case "D" & enuLookIn.eLookInToBePrinted                     'SADBEL IET, NCTS, Combined NCTS
                    SearchIn "IMPORT", 1, "", "SDI2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "EXPORT", 2, "", "SDE2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "TRANSIT", 3, "", "SDT2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "NCTS", 7, "", "SDN2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "COMBINED NCTS", 9, "", "SDC2"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInReleased                        'PLDA IE
                    SearchIn "PLDA IMPORT", 14, "", "SDX2"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "SDZ2"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInSent                            'Departure, PLDA IE
                    SearchInEDI 11, 5, "", "32ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA IMPORT", 14, "", "PX01"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "PZ01"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInEmergencyProcedure              'PLDA IE
                    SearchIn "PLDA IMPORT", 14, "", "EP"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "EP"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInExitEC                          'PLDA E
                    SearchIn "PLDA COMBINED", 18, "", "SDZ4"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInGuarantee                       'Departure
                    SearchInEDI 11, 5, "", "55ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInUnderControl                    'Departure
                    SearchInEDI 11, 5, "", "35ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInReleases                        'Departure
                    SearchInEDI 11, 5, "", "36ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInReleaseRejected                 'Departure
                    SearchInEDI 11, 5, "", "37ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInCancelled                       'PLDA IE, Departure
                    SearchInEDI 11, 5, "", "39ED"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA IMPORT", 14, "", "SDX3"
                    If mblnCancel = True Then Exit Sub
                    
                    SearchIn "PLDA COMBINED", 18, "", "SDZ3"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInWrittenOff                      'Departure
                    SearchInEDI 11, 5, "", "40ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInAmendmentSent                   'Departure
                    SearchInEDI 11, 5, "", "52ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInAmendmentRejected               'Departure
                    SearchInEDI 11, 5, "", "53ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInAmendmentAccepted               'Departure
                    SearchInEDI 11, 5, "", "54ED"
                    If mblnCancel = True Then Exit Sub
                    
                Case "D" & enuLookIn.eLookInArchives                        'Arrival
                    SearchInEDI 12, 2, "", "50ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "50ED"
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInArrivalNotificationSent         'Arrival
                    SearchInEDI 12, 2, "", "45ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "45ED"
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInArrivalNotificationRejected     'Arrival
                    SearchInEDI 12, 2, "", "46ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "46ED"
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInUnloadingPermitted              'Arrival
                    SearchInEDI 12, 2, "", "47ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "47ED"
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInUnloadingRemarksSent            'Arrival
                    SearchInEDI 12, 2, "", "48ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "48ED"
                    If mblnCancel = True Then Exit Sub
                
                Case "D" & enuLookIn.eLookInUnloadingRemarksRejected        'Arrival
                    SearchInEDI 12, 2, "", "49ED"
                    If mblnCancel = True Then Exit Sub
                
                    SearchInEDI 12, 11, "", "49ED"
                    If mblnCancel = True Then Exit Sub
                
                '-------------------------------------------------
                
                Case Else   'for Archive
                    '**** icbLookIn.SelectedItem.Key = "D" & YYYYY  ******`

                    strYear = Mid(CStr(icbLookIn.SelectedItem.Key), 4)
                    
                    SearchIn "IMPORT", 1, strYear, "HI" & strYear & "I"
                    SearchIn "EXPORT", 1, strYear, "HI" & strYear & "E"
                    SearchIn "TRANSIT", 1, strYear, "HI" & strYear & "T"
                    SearchIn "NCTS", 1, strYear, "HI" & strYear & "N"
                    SearchIn "COMBINED NCTS", 1, strYear, "HI" & strYear & "C"
                    SearchIn "PLDA IMPORT", 1, strYear, "HI" & strYear & "X"
                    SearchIn "PLDA COMBINED", 1, strYear, "HI" & strYear & "Z"
                    
                    'For EDI Archive
                    'Added by Rachelle on 06/16/2005
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            If Mid(File1.List(lngCounter), 15, 2) = strYear And Right(File1.List(lngCounter), 3) = "mdb" Then
                                SearchInEDI 11, 5, Mid(File1.List(lngCounter), 15, 2), "40ED"
                                SearchInEDI 12, 2, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                SearchInEDI 12, 11, Mid(File1.List(lngCounter), 15, 2), "50ED"
                            End If
                        Next lngCounter
                    End If
            End Select
        
        Case "D" & enuDocType.edocimport
            '#####
            SearchAccdngToDocType "IMPORT", "I", 1
            
        Case "D" & enuDocType.eDocExport
            SearchAccdngToDocType "EXPORT", "E", 2
            
        Case "D" & enuDocType.eDocOTS
            SearchAccdngToDocType "TRANSIT", "T", 3
           
        Case "D" & enuDocType.eDocNCTS
            SearchAccdngToDocType "NCTS", "N", 7
            
        Case "D" & enuDocType.edoccombined
            SearchAccdngToDocType "COMBINED NCTS", "C", 9
            
        Case "D" & enuDocType.eDocEDIDepartures
            SearchAccdngToDocType "EDI DEPARTURES", "", 11
            
        Case "D" & enuDocType.eDocEDIARRIVALS
            SearchAccdngToDocType "EDI ARRIVALS", "", 12
        
        Case "D" & enuDocType.eDocPLDAImport
            SearchAccdngToDocType "PLDA IMPORT", "X", 14
    
        Case "D" & enuDocType.eDocPLDACombined
            SearchAccdngToDocType "PLDA COMBINED", "Z", 18
    
    End Select
    
    
End Sub

Private Function CriteriaNotSpecified() As Boolean
    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny Then
        
    End If
End Function

Private Function SqlWhere(TableName As String, Optional BoxCode As String, Optional bytDocType As Byte) As String

Dim strTemp As String
Dim strFieldToUse As String
Dim strValue As String
Dim intDataType As Integer
Dim strSQLWhere As String
Dim strTableToOpen As String
Dim strFieldTemp As String
Dim strWildCardCharCheck As String

    '============ for DATE CONDITION ====================
    If optFind.Value = True Then
        strTemp = ""
        strTemp = DateCondition(TableName)
        strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
    End If
    '====================================================

    '============ for DOCUMENT NAME CONDITION ===========
    strTemp = ""
    strWildCardCharCheck = cboName.Text
    If strWildCardCharCheck <> "" Then
        If InStr(1, strWildCardCharCheck, "*") > 0 Or _
            InStr(1, strWildCardCharCheck, "?") > 0 Then
                
            strTemp = " [DOCUMENT NAME] LIKE " & Chr(39) & ProcessQuotes(strWildCardCharCheck) & Chr(39) & " "
        Else
            strTemp = " [DOCUMENT NAME] = " & Chr(39) & ProcessQuotes(strWildCardCharCheck) & Chr(39) & " "
        End If
        
    End If

    strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
    '====================================================
    
    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny Then
    
        '=================== for BOXCODE + CONDITION + VALUE ==========
        If txtValue.Text <> "" Then
            Select Case bytDocType
                Case 1
                    If InStr(1, IMPORTHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                    End If
                Case 2
                    If InStr(1, EXPORTHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                    End If
                
                Case 3
                    If InStr(1, OTSHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                    End If

                Case 7
                    If InStr(1, NCTSHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[NCTS HEADER]."
                        
                        If InStr(1, ZEKERHEID, BoxCode) > 0 Then
                            strFieldToUse = "[NCTS HEADER ZEKERHEID]."
                        End If
                    Else
                        strFieldToUse = "[NCTS DETAIL]."
                        
                        If InStr(1, COLLI, BoxCode) > 0 Then
                            strFieldToUse = "[NCTS DETAIL COLLI]."
                        ElseIf InStr(1, Container, BoxCode) > 0 Then
                            strFieldToUse = "[NCTS DETAIL CONTAINER]."
                        ElseIf InStr(1, DOCUMENTEN, BoxCode) > 0 Then
                            strFieldToUse = "[NCTS DETAIL DOCUMENTEN]."
                        ElseIf InStr(1, BIJZONDERE, BoxCode) > 0 Then
                            strFieldToUse = "[NCTS DETAIL BIJZONDERE]."
                        End If
                    End If
                
                Case 9
                    If InStr(1, COMBINEDHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                        If InStr(1, ZEKERHEID, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS HEADER ZEKERHEID]."
                        End If
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                        
                        If InStr(1, COLLI, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL COLLI]."
                        ElseIf InStr(1, Container, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL CONTAINER]."
                        ElseIf InStr(1, DOCUMENTEN, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL DOCUMENTEN]."
                        ElseIf InStr(1, BIJZONDERE, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL BIJZONDERE]."
                        ElseIf InStr(1, GEVEOLIGE, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL GEVOELIGE]."
                        ElseIf InStr(1, GEODEREN, BoxCode) > 0 Then
                            strFieldToUse = "[COMBINED NCTS DETAIL GOEDEREN]."
                        End If
                    End If
                
                Case 14
                    If InStr(1, PLDAIMPORTHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                        If InStr(1, PLDAIMPORTZEGELS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT HEADER ZEGELS]."
                        ElseIf InStr(1, PLDAIMPORTHEADERHANDELAARS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT HEADER HANDELAARS]."
                        End If
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                        
                        If InStr(1, PLDAIMPORTBIJZONDERE, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL BIJZONDERE]."
                        ElseIf InStr(1, PLDAIMPORTCONTAINER, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL CONTAINER]."
                        ElseIf InStr(1, PLDAIMPORTDOCUMENTEN, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL DOCUMENTEN]."
                        ElseIf InStr(1, PLDAIMPORTZELF, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL ZELF]."
                        ElseIf InStr(1, PLDAIMPORTDETAILHANDELAARS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL HANDELAARS]."
                        ElseIf InStr(1, PLDAIMPORTDETAILBEREKENINGS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA IMPORT DETAIL BEREKENINGS EENHEDEN]."
                        End If
                    End If
                
                Case 18
                    If InStr(1, PLDACOMBINEDHEADER, BoxCode) > 0 Then
                        strFieldToUse = "[" & TableName & " HEADER]."
                        If InStr(1, PLDACOMBINEDZEGELS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED HEADER ZEGELS]."
                        ElseIf InStr(1, PLDACOMBINEDHEADERHANDELAARS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED HEADER HANDELAARS]."
                        End If
                    Else
                        strFieldToUse = "[" & TableName & " DETAIL]."
                        
                        If InStr(1, PLDACOMBINEDBIJZONDERE, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED DETAIL BIJZONDERE]."
                        ElseIf InStr(1, PLDACOMBINEDCONTAINER, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED DETAIL CONTAINER]."
                        ElseIf InStr(1, PLDACOMBINEDDOCUMENTEN, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED DETAIL DOCUMENTEN]."
                        ElseIf InStr(1, PLDACOMBINEDDETAILHANDELAARS, BoxCode) > 0 Then
                            strFieldToUse = "[PLDA COMBINED DETAIL HANDELAARS]."
                        End If
                    End If
                
            End Select
            
            strTableToOpen = Mid(strFieldToUse, 2, Len(strFieldToUse) - 3)
            intDataType = GetDataType(g_conSADBEL, strTableToOpen, BoxCode)
            
            If cboCondition.ListIndex = 0 Then
                If intDataType = 8 Then
                    strTemp = BoxCode & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & BoxCode & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')"
                ElseIf intDataType = 10 Then
                    strTemp = BoxCode & " LIKE " & Chr(39) & "*" & ProcessQuotes(txtValue.Text) & "*" & Chr(39)
                Else
                    strTemp = BoxCode & " LIKE " & txtValue.Text
                End If
            ElseIf cboCondition.ListIndex = 1 Then
                If intDataType = 8 Then
                    'DateCondition = strDateTab & " >= CDate('" + strDate1 + "') and " & strDateTab & " <= CDate('" + strDate2 + "')"
                    strTemp = BoxCode & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & BoxCode & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')"
                Else
                    strValue = FormatValue(txtValue.Text, intDataType)
                    strTemp = BoxCode & " = " & strValue
                End If
                        
            End If
            strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
            '============== END OF BOXCODE + CONDITION + VALUE ===============
            
        End If
    Else
        strTemp = ""
        If icbBox.Text <> "" And txtValue.Text <> "" Then
            'BoxCode = Left(icbBox.SelectedItem.Key, Len(icbBox.SelectedItem.Key) - 1)
            BoxCode = Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)
            
            'Dim bytKey As Byte
            
            'bytKey = Right(icbBox.SelectedItem.Key, 1)
            'If BoxCode = "MRN" Then
            '    bytkey =
            'End If
    
            Select Case Right(icbBox.SelectedItem.Key, 1)
                Case enuGroup.eGroupMain
                    strFieldToUse = "[" & TableName & "]."
                    
                Case enuGroup.eGroupHeader
                    strFieldToUse = "[" & TableName & " HEADER]."
                    
                    Select Case icbType.SelectedItem.Key
                        
                        Case "D" & enuDocType.eDocNCTS, _
                             "D" & enuDocType.edoccombined
                            
                            If InStr(1, ZEKERHEID, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " HEADER ZEKERHEID]."
                            End If
                            
                        Case "D" & enuDocType.eDocPLDAImport
                            
                            If InStr(1, PLDAIMPORTZEGELS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " HEADER ZEGELS]."
                            ElseIf InStr(1, PLDAIMPORTHEADERHANDELAARS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " HEADER HANDELAARS]."
                            End If
                                
                        Case "D" & enuDocType.eDocPLDACombined
                        
                            If InStr(1, PLDACOMBINEDHEADERHANDELAARS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " HEADER HANDELAARS]."
                            ElseIf InStr(1, PLDACOMBINEDZEGELS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " HEADER ZEGELS]."
                            End If
                                
                    End Select
                
                Case enuGroup.eGroupDetail
                    strFieldToUse = "[" & TableName & " DETAIL]."
                    
                    Select Case icbType.SelectedItem.Key
                        
                        Case "D" & enuDocType.eDocNCTS, _
                             enuDocType.edoccombined
    
                            If InStr(1, COLLI, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL COLLI]."
                            ElseIf InStr(1, Container, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL CONTAINER]."
                            ElseIf InStr(1, DOCUMENTEN, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL DOCUMENTEN]."
                            ElseIf InStr(1, BIJZONDERE, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL BIJZONDERE]."
                            ElseIf InStr(1, GEVEOLIGE, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL GEVOELIGE]."
                            ElseIf InStr(1, GEODEREN, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL GOEDEREN]."
                            End If
                            
                        Case "D" & enuDocType.eDocPLDAImport
    
                            If InStr(1, PLDAIMPORTBIJZONDERE, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL BIJZONDERE]."
                            ElseIf InStr(1, PLDAIMPORTCONTAINER, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL CONTAINER]."
                            ElseIf InStr(1, PLDAIMPORTDOCUMENTEN, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL DOCUMENTEN]."
                            ElseIf InStr(1, PLDAIMPORTZELF, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL ZELF]."
                            ElseIf InStr(1, PLDAIMPORTDETAILHANDELAARS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL HANDELAARS]."
                            ElseIf InStr(1, PLDAIMPORTDETAILBEREKENINGS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL BEREKENINGS EENHEDEN]."
                            End If
                        
                        Case "D" & enuDocType.eDocPLDACombined
    
                            If InStr(1, PLDACOMBINEDBIJZONDERE, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL BIJZONDERE]."
                            ElseIf InStr(1, PLDACOMBINEDCONTAINER, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL CONTAINER]."
                            ElseIf InStr(1, PLDACOMBINEDDOCUMENTEN, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL DOCUMENTEN]."
                            ElseIf InStr(1, PLDACOMBINEDDETAILHANDELAARS, BoxCode) > 0 Then
                                strFieldToUse = "[" & TableName & " DETAIL HANDELAARS]."
                            End If
                        
                    End Select
            End Select
            
            'strFieldTemp = "[" & Left(icbBox.SelectedItem.Key, Len(icbBox.SelectedItem.Key) - 1) & "]"
            strFieldTemp = "[" & BoxCode & "]"
            
            intDataType = GetDataType(g_conSADBEL, Mid(strFieldToUse, 2, Len(strFieldToUse) - 3), BoxCode)
            
                strValue = FormatValue(txtValue.Text, intDataType)
            
            Select Case cboCondition.ListIndex
                Case enuCondition.eContains
                    If intDataType = 8 Then
                        'strTemp = strFieldToUse & strFieldTemp & " = " & strValue
                        strTemp = strFieldToUse & strFieldTemp & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & strFieldToUse & strFieldTemp & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')"
                    Else
                        strValue = FormatValue(txtValue.Text, intDataType, eContains)
                        strTemp = strFieldToUse & strFieldTemp & " LIKE " & strValue
                    End If
                    
                Case enuCondition.eIsExactly
                    If intDataType = 8 Then
                        strTemp = strFieldToUse & strFieldTemp & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & strFieldToUse & strFieldTemp & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')"
                    Else
                        strValue = FormatValue(txtValue.Text, intDataType, eIsExactly)
                        strTemp = strFieldToUse & strFieldTemp & " = " & strValue
                    End If
                    
                Case enuCondition.eDoesnotContain
                    strValue = FormatValue(txtValue.Text, intDataType, eDoesnotContain)
                    strTemp = "InStr(1, UCase(" & strFieldToUse & strFieldTemp & "), " & UCase(strValue) & ") = 0"
                    
                Case enuCondition.eIsEmpty
                    'strValue = FormatValue(txtValue.Text, intDataType)
                    strTemp = "len(trim(" & strFieldToUse & strFieldTemp & ")) <= 0 "
                    
                Case enuCondition.eIsNotEmpty
                    'strValue = FormatValue(txtValue.Text, intDataType)
                    strTemp = "len(trim(" & strFieldToUse & strFieldTemp & ")) > 0 "
                
            End Select
            
        End If
        
        strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
        

    End If

    SqlWhere = strSQLWhere
End Function

Private Sub SearchIn(FieldToUse As String, bytDocType As Byte, strYear As String, ParamArray TreeID() As Variant)
    '"IMPORT", 1, "MASTER", "DE"
    Dim strSQl As String
    Dim rstResult As ADODB.Recordset
    Dim intTreeCtr As Integer
    Dim rstBoxDefault As ADODB.Recordset
    Dim strSQLWhere As String
    Dim strPreviousCode As String
    Dim conArchive As ADODB.Connection
    Dim lngCounter As Long
    Dim strPLDAFieldValue As String
    Dim strFieldValue As String 'allan
    Dim conDBToBeSearched As ADODB.Connection       'Dim datDBToBeSearched As dao.Database

    If strYear <> "" Then
        If Dir(cAppPath & "\mdb_history" & Right(strYear, 2) & ".mdb") = "" Then
            Exit Sub
        End If
        '<<< dandan 112306
        '<<< Update with database password
        'Set datArchive = OpenDatabase(cAppPath & "\mdb_history" & Right(strYear, 2) & ".mdb")
        ADOConnectDB conArchive, m_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(strYear, 2)
        'DAOConnectDB datArchive, cAppPath, "mdb_history" & Right(strYear, 2) & ".mdb"
    End If

    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny And txtValue.Text <> "" Then
        If txtValue.Text <> "" Then
            
            Select Case bytDocType
                Case 1
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT IMPORT ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT IMPORT ADMIN] order by [BOX CODE]")
                Case 2
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT EXPORT ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT EXPORT ADMIN] order by [BOX CODE]")
                Case 3
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT TRANSIT ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT TRANSIT ADMIN] order by [BOX CODE]")
                Case 7
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT TRANSIT NCTS ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT TRANSIT NCTS ADMIN] order by [BOX CODE]")
                Case 9
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT COMBINED NCTS ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT COMBINED NCTS ADMIN] order by [BOX CODE]")
                Case 14
                    SetPLDAProperties 14
                    
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT PLDA IMPORT ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT PLDA IMPORT ADMIN] order by [BOX CODE]")
                Case 18
                    SetPLDAProperties 18
                    
                    ADORecordsetOpen "Select [BOX CODE] from [BOX DEFAULT PLDA COMBINED ADMIN] order by [BOX CODE]", _
                                        g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                    'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE] from [BOX DEFAULT PLDA COMBINED ADMIN] order by [BOX CODE]")
            End Select
            
            Do While Not rstBoxDefault.EOF
                DoEvents
                
                If mblnCancel = True Then GoTo Cancelled
                strSQLWhere = SqlWhere(FieldToUse, rstBoxDefault![Box Code], bytDocType)
            
                For intTreeCtr = 0 To UBound(TreeID)
                    StatusBar1.Panels(1).Text = "Searching in " & DocumentType(bytDocType) & " - " & CStr(rstBoxDefault![Box Code]) & " - " & FolderName(CStr(TreeID(intTreeCtr))) & strYear
                    
                    strSQl = GetSQLToUse(FieldToUse, bytDocType, CStr(TreeID(intTreeCtr)), strSQLWhere)
                    
                    If strYear = "" Then
                        
                        ADORecordsetOpen strSQl, g_conSADBEL, rstResult, adOpenKeyset, adLockOptimistic
                        'Set rstResult = datSADBEL.OpenRecordset(strSQl)
                        
                        Set conDBToBeSearched = g_conSADBEL
                    Else    'means from Archive
                        ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                        'Set rstResult = datArchive.OpenRecordset(strSQl)
                        
                        Set conDBToBeSearched = conArchive
                    End If
                    
                    With rstOfflineTemp
                        Do While Not rstResult.EOF
                        
                            If strPreviousCode <> rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value Then
                                .AddNew
                                
                                .Fields("CODE").Value = rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value
                                .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                                .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                                .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                                .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                                .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                                
                                'CSCLP-248
                                If chkAllFields.Value = vbUnchecked Then
                                    For lngCounter = 7 To .Fields.Count - 1
                                        'Only process fields selected by user from Date/Advanced tab
                                        If optFind.Value = True Then    'Is option checked?
                                            If Combo5(1).Text = .Fields(lngCounter).Name Then   'Is selected the current field?
                                                strPLDAFieldValue = Empty
                                                If IsFieldExisting(rstResult, Combo5(1).Text) Then
                                                    strFieldValue = GetFieldValue(rstResult, Combo5(1).Text) 'allan
                                                    .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                                ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                                    strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                                    .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                                End If
                                            End If
                                        End If
                                        
                                        If Len(icbBox.Text) > 0 Then    'Is the box empty?
                                            If Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2) = .Fields(lngCounter).Name Then    'Is selected the current field?
                                                strPLDAFieldValue = Empty
                                                If IsFieldExisting(rstResult, Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)) Then
                                                    strFieldValue = GetFieldValue(rstResult, Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)) 'allan
                                                    .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                                ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                                    strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                                    .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                                End If
                                            End If
                                        End If
                                    Next lngCounter
                                Else
                                    'CSCLP-248 comment
                                    'Original Code
                                    
                                    'lngcounter
                                    For lngCounter = 7 To .Fields.Count - 1
                                        strPLDAFieldValue = ""
                                        If IsFieldExisting(rstResult, .Fields(lngCounter).Name) Then
                                            strFieldValue = GetFieldValue(rstResult, .Fields(lngCounter).Name) 'allan
                                            .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                        ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                            strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                            .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                        End If
                                    Next lngCounter
                                End If
                                
                                .Update
                            
                            End If
                            strPreviousCode = rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value
                            
                            rstResult.MoveNext
                        Loop
                    End With
                    
                Next
        
Cancelled:
            
                ADORecordsetClose rstResult
                
                        
                'Rachelle 051805
                'If the form is closed without stopping the searching process, datSadbel becomes nothing, which leads to an error when there
                'recordsets that depend on it
                If Not g_conSADBEL Is Nothing Then
                    rstBoxDefault.MoveNext
                Else
                    Exit Do
                End If
            Loop
            
        End If
    Else
    
        strSQLWhere = SqlWhere(FieldToUse)
        
        If bytDocType = 14 Or bytDocType = 18 Then
            SetPLDAProperties bytDocType
        End If
        
        For intTreeCtr = 0 To UBound(TreeID)
            StatusBar1.Panels(1).Text = "Searching " & FolderName(CStr(TreeID(intTreeCtr)))
            strSQl = GetSQLToUse(FieldToUse, bytDocType, CStr(TreeID(intTreeCtr)), strSQLWhere)
            
            If strYear = "" Then
                ADORecordsetOpen strSQl, g_conSADBEL, rstResult, adOpenKeyset, adLockOptimistic
                'Set rstResult = datSADBEL.OpenRecordset(strSQl)
                
                Set conDBToBeSearched = g_conSADBEL
            Else    'means from Archive
                ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                'Set rstResult = datArchive.OpenRecordset(strSQl)
                
                Set conDBToBeSearched = conArchive
            End If
            
            With rstOfflineTemp
                Do While Not rstResult.EOF
                    DoEvents
                    If mblnCancel = True Then GoTo Cancelled2
                    If strPreviousCode <> rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value Then
                    
                        .AddNew
                        
                        .Fields("CODE").Value = rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value
                        .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                        .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                        .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                        .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                        .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                        .Fields("ARCHIVE DATE").Value = ""
                        
                        'CSCLP-248
                        If chkAllFields = vbUnchecked Then
                            For lngCounter = 7 To .Fields.Count - 1
                                'Only process fields selected by user from Date/Advanced tab
                                If optFind.Value = True Then    'Is option checked?
                                    If Combo5(1).Text = .Fields(lngCounter).Name Then   'Is selected the current field?
                                        strPLDAFieldValue = Empty
                                        If IsFieldExisting(rstResult, Combo5(1).Text) Then
                                            strFieldValue = GetFieldValue(rstResult, Combo5(1).Text) 'allan
                                            .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                        ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                            strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                            .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                        End If
                                    End If
                                End If
                                
                                If Len(icbBox.Text) > 0 Then    'Is the box empty?
                                    If Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2) = .Fields(lngCounter).Name Then    'Is selected the current field?
                                        strPLDAFieldValue = Empty
                                        If IsFieldExisting(rstResult, Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)) Then
                                            strFieldValue = GetFieldValue(rstResult, Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)) 'allan
                                            .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                        ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                            strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                            .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                        End If
                                    End If
                                End If
                            Next lngCounter
                        Else
                            'CSCLP-248 comment
                            'Original Code
                                    
                            'lngcounter
                            For lngCounter = 7 To .Fields.Count - 1
                                strPLDAFieldValue = ""
                                If IsFieldExisting(rstResult, .Fields(lngCounter).Name) Then
                                    strFieldValue = GetFieldValue(rstResult, .Fields(lngCounter).Name) 'allan
                                    .Fields(lngCounter).Value = IIf(IsNull(strFieldValue), "", IIf(Len(strFieldValue) > 100, Left(strFieldValue, 97) & "...", strFieldValue)) 'allan
                                ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                    strPLDAFieldValue = GetPLDAFieldValue(conDBToBeSearched, bytDocType, .Fields(lngCounter).Name, .Fields("CODE").Value)
                                    .Fields(lngCounter).Value = IIf(IsNull(strPLDAFieldValue), "", IIf(Len(strPLDAFieldValue) > 100, Left(strPLDAFieldValue, 97) & "...", strPLDAFieldValue))
                                End If
                            Next lngCounter
                        End If
                        
                        .Update
                        
                    End If
                    
                    strPreviousCode = rstResult.Fields(IIf(bytDocType <> 14 And bytDocType <> 18, FieldToUse & " HEADER.CODE", "CODE")).Value
                    rstResult.MoveNext
                Loop
            End With
            
        Next
        
Cancelled2:
        If Not mblnCancel Then
            ADORecordsetClose rstResult
        End If
    End If

    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If

    TransferToListView
    'DeleteRecordsInOfflineRst
    CreateOfflineRecordset
End Sub

Private Sub ConnectToDB(ByRef DataSourceProperties As CDataSourceProperties)
    'cAppPath is mdb path...
    Dim lngCounter As Long
    
    '<<< dandan 112306
    '<<< Update with database password
    'Set datSADBEL = OpenDatabase(cAppPath & "\mdb_sadbel.mdb")
    'DAOConnectDB datSADBEL, cAppPath, "mdb_sadbel.mdb"
    ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
                        
    'Set datData = OpenDatabase(cAppPath & "\mdb_data.mdb")
    'DAOConnectDB datData, cAppPath, "mdb_data.mdb"
    ADOConnectDB g_conData, DataSourceProperties, DBInstanceType_DATABASE_DATA
                        
    'Set datEDIFACT = OpenDatabase(cAppPath & "\EDIFACT.mdb")
    'DAOConnectDB datEDIFACT, cAppPath, "EDIFACT.mdb"
    ADOConnectDB g_conEDIFACT, DataSourceProperties, DBInstanceType_DATABASE_EDIFACT
    
    Dim strHistoryYear As String
                        
    File1.Path = cAppPath
    File1.Pattern = "mdb_EDIHistory*.mdb"
    
    If File1.ListCount > 0 Then
        blnEDIHistoryExisting = True
        ReDim datEDIHistory(File1.ListCount - 1)
        
        For lngCounter = 0 To File1.ListCount - 1
            If IsNumeric(Mid(File1.List(lngCounter), 15, 2)) Then
                If Right(File1.List(lngCounter), 3) = "mdb" Then
                    
                    strHistoryYear = Replace(strHistoryYear, ".mdb", "")
                    strHistoryYear = Right$(strHistoryYear, 2)
                                    
                    ADOConnectDB g_conEDIHistory(lngCounter), DataSourceProperties, DBInstanceType_DATABASE_EDI_HISTORY, strHistoryYear
                    
                    '<<< dandan 112306
                    '<<< Update with database password
                    'Set datEDIHistory(lngCounter) = OpenDatabase(cAppPath & "\" & File1.List(lngCounter))
                    'DAOConnectDB datEDIHistory(lngCounter), cAppPath, File1.List(lngCounter)
                
                End If
            End If
        Next lngCounter
    Else
        blnEDIHistoryExisting = False
    End If
    
    ADOConnectDB g_conTemplate, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
    'Set conFind = New ADODB.Connection
    ''conFind.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & cAppPath & "\TemplateCP.mdb"
    'DAOConnectDB conFind, cAppPath, "TemplateCP.mdb"

End Sub

Private Sub UpdateBoxList()
    Dim strList(1) As String
    Dim i As Integer
    Dim strBox As String
    Dim j As Integer
    Dim rstFieldGrouping As ADODB.Recordset
    Dim rstBoxDefault As ADODB.Recordset
    Dim itm As ComboItem
    Dim strCommand As String

    icbBox.Text = ""
    
    Dim strGroup As String
    Dim strTableName As String
    
    Select Case icbType.SelectedItem.Key
        Case "D" & enuDocType.edocimport
            strList(0) = IMPORTHEADER
            strList(1) = IMPORTDETAIL
                    
            strTableName = "IMPORT"
            strGroup = "[GROUP] = 1 or [GROUP] = 4"
            
        Case "D" & enuDocType.eDocExport
            strList(0) = EXPORTHEADER
            strList(1) = EXPORTDETAIL
            
            strTableName = "EXPORT"
            strGroup = "[GROUP] = 1 or [GROUP] = 4"
        
        Case "D" & enuDocType.eDocOTS
            strList(0) = OTSHEADER
            strList(1) = OTSDETAIL
        
            strTableName = "TRANSIT"
            strGroup = "[GROUP] = 1 or [GROUP] = 4"
        
        Case "D" & enuDocType.eDocNCTS
            strList(0) = NCTSHEADER
            strList(1) = NCTSDETAIL
        
            strTableName = "TRANSIT NCTS"
            strGroup = "[GROUP] = 5 or [GROUP] = 8"

        Case "D" & enuDocType.edoccombined
            strList(0) = COMBINEDHEADER
            strList(1) = COMBINEDDETAIL
            
            strTableName = "COMBINED NCTS"
            strGroup = "[GROUP] = 9 or [GROUP] = 12"

        Case "D" & enuDocType.eDocEDIDepartures
            
            strTableName = "EDI NCTS"
            strGroup = "[GROUP] = 9 or [GROUP] = 12"
        
        Case "D" & enuDocType.eDocEDIARRIVALS
        
            strTableName = "EDI NCTS2"
            strGroup = "[GROUP] = 9 or [GROUP] = 12"
        
        Case "D" & enuDocType.eDocPLDAImport
            strList(0) = PLDAIMPORTHEADER
            strList(1) = PLDAIMPORTDETAIL
        
            strTableName = "PLDA IMPORT"
            strGroup = "[GROUP] = 9 or [GROUP] = 12"
        
        Case "D" & enuDocType.eDocPLDACombined
            strList(0) = PLDACOMBINEDHEADER
            strList(1) = PLDACOMBINEDDETAIL
        
            strTableName = "PLDA COMBINED"
            strGroup = "[GROUP] = 9 or [GROUP] = 12"
    End Select

        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[FIELD GROUPING] "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & strGroup & " "
    ADORecordsetOpen strCommand, g_conSADBEL, rstFieldGrouping, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[BOX CODE], "
        strCommand = strCommand & "[DATA TYPE], " < _
        strCommand = strCommand & "[" & cLanguage & " DESCRIPTION] as Description "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[BOX DEFAULT " & Trim$(strTableName) & " ADMIN] "
        strCommand = strCommand & "ORDER BY "
        strCommand = strCommand & "[BOX CODE] "
    ADORecordsetOpen strCommand, g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic

    icbBox.ComboItems.Clear
    
    With rstBoxDefault
        Do While Not .EOF
            If InStr(1, strList(0), ![Box Code]) > 0 Then
                icbBox.ComboItems.Add , CStr("D" & ![Box Code] & enuGroup.eGroupHeader), Trim(Left(![Box Code] & " - " & ![Description], 59))
            Else
                icbBox.ComboItems.Add , CStr("D" & ![Box Code] & enuGroup.eGroupDetail), Trim(Left(![Box Code] & " - " & ![Description], 59))
            End If
        
            rstBoxDefault.MoveNext
        Loop
    End With
    
    Do While Not rstFieldGrouping.EOF

        If UCase(rstFieldGrouping![Import Column]) = "DOC NUMBER" Or UCase(rstFieldGrouping![Import Column]) = "MRN" Then
            icbBox.ComboItems.Add , CStr("D" & rstFieldGrouping![Import Column] & enuGroup.eGroupHeader), rstFieldGrouping![Import Column]
        Else
            icbBox.ComboItems.Add , CStr("D" & rstFieldGrouping![Import Column] & IIf(rstFieldGrouping![Group] = 1 Or rstFieldGrouping![Group] = 5 Or rstFieldGrouping![Group] = 9, enuGroup.eGroupMain, enuGroup.eGroupDetail)), rstFieldGrouping![Import Column]
        End If

        rstFieldGrouping.MoveNext
        
    Loop
    
    ADORecordsetClose rstFieldGrouping
    ADORecordsetClose rstBoxDefault
End Sub

Private Sub icbType_Click()
    Dim blnLVW As Boolean
    
    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny Then
        icbBox.Enabled = False
        UpdateConditionList 0
    Else
        icbBox.Enabled = True
        
        UpdateBoxList
        UpdateConditionList 1
    End If
    
    If strOldDocType <> icbType.SelectedItem Then
        blnLVW = lvwItemsFound.Visible
        If blnLVW = True Then
            lvwItemsFound.Visible = False
            Me.MousePointer = vbHourglass
        End If
        
        SaveSpecs strOldDocType
        strOldDocType = icbType.SelectedItem
        lvwItemsFound.ListItems.Clear
        Me.StatusBar1.Panels(1).Text = ""
        
        LoadListView
        CreateOfflineRecordset
        InitListView ""
        
        If blnLVW = True Then
            lvwItemsFound.Visible = True
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub lvwItemsFound_DblClick()
    Dim strDate() As String
    Dim strRealDate As String
    
    If lvwItemsFound.ListItems.Count > 0 Then
        clsFindForm.SelectedItemTag = lvwItemsFound.SelectedItem.Tag
        clsFindForm.SelectedItemText = lvwItemsFound.SelectedItem.Text
        clsFindForm.ListSubItems = lvwItemsFound.SelectedItem.ListSubItems(2).Tag
        clsFindForm.SubItems = lvwItemsFound.SelectedItem.SubItems(2)
        If (clsFindForm.SubItems = Translate(1379) Or clsFindForm.SubItems = Translate(1074)) And _
            (clsFindForm.ListSubItems = "40ED" Or clsFindForm.ListSubItems = "50ED") Then
            clsFindForm.strYear = lvwItemsFound.SelectedItem.ListSubItems(5).Tag
        Else
            clsFindForm.strYear = ""
        End If
    
        Select Case lvwItemsFound.SelectedItem.ListSubItems(1).Text
    
            Case "Import"
                CallingForm.LoadDocument edocimport, False
                
            Case "Export"
                CallingForm.LoadDocument eDocExport, False
                
            Case "Transit"
                CallingForm.LoadDocument eDocOTS, False
    
            Case "NCTS"
                CallingForm.LoadDocument eDocNCTS, False
    
            Case "Combined NCTS"
                CallingForm.LoadDocument edoccombined, False
        
            Case "EDI Departures"
                CallingForm.LoadDocument eDocEDIDepartures, False
                
            Case "EDI Arrivals"
                CallingForm.LoadDocument eDocEDIARRIVALS, False
                
            Case "PLDA Import"
                CallingForm.LoadDocument eDocPLDAImport, False
            
            Case "PLDA Combined"
                CallingForm.LoadDocument eDocPLDACombined, False
                
        End Select
    End If
End Sub

Private Sub lvwItemsFound_GotFocus()
    cmdFind.Default = False
End Sub

Private Sub lvwItemsFound_KeyDown(KeyCode As Integer, Shift As Integer)
    'Added by BCo 2006-05-03
    'Attempt at context menu
    If KeyCode = 93 Then
        lvwItemsFound_MouseUp 2, 0, 0, 0
    End If
End Sub

Private Sub lvwItemsFound_KeyPress(KeyAscii As Integer)
    'Added by BCo 2006-05-03
    'Added Enter key or Spacebar condition, since it previously to open a document when pressing ESC
    If KeyAscii = 13 Or KeyAscii = 32 Then
        If lvwItemsFound.ListItems.Count > 0 Then
            lvwItemsFound_DblClick
        End If
    End If
End Sub

Private Sub lvwItemsFound_LostFocus()
    cmdFind.Default = True
End Sub

Private Sub lvwItemsFound_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Added by BCo 2006-05-03
    'Licensing prevents access to Open A Copy if selected doc type is unlicensed
    'though, if they find a way to open, the toolbars are disabled anyways. Aesthetic purposes?
    If lvwItemsFound.ListItems.Count > 0 Then
        SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = IsFeatureLicensed(lvwItemsFound.SelectedItem.ListSubItems(1).Text)
    
        If Button = vbRightButton Then
            CheckMenuForGrid
        End If
    End If
End Sub

Private Sub optAll_Click()
    DTPicker1(0).Enabled = False
    DTPicker1(1).Enabled = False
    Text2(0).Enabled = False
    Text2(1).Enabled = False
    Combo5(1).Enabled = False
    UpDown1(0).Enabled = False
    UpDown1(1).Enabled = False
    
    Text2(0).BackColor = vbButtonFace     'RGB(192, 192, 192)
    Text2(1).BackColor = vbButtonFace     'RGB(192, 192, 192)
    Combo5(1).BackColor = vbButtonFace    'RGB(192, 192, 192)
    
    Option2(0).Value = False
    Option2(1).Value = False
    Option2(2).Value = False

    Option2(0).Enabled = False
    Option2(1).Enabled = False
    Option2(2).Enabled = False

End Sub

Private Sub optFind_Click()

    Option2(0).Value = True
    Option2_Click (0)
    
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
    
    Combo5(1).Enabled = True
    Combo5(1).BackColor = vbWindowBackground    'RGB(255, 255, 255)

    If optFind.Value = True Then
        
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    'Option1(1).Value = True
    Combo5(1).Enabled = True
    Combo5(1).BackColor = vbWindowBackground    'RGB(255, 255, 255)
    Select Case Index
        Case 0
            If Option2(0).Value = True Then
                DTPicker1(0).Enabled = True
                DTPicker1(1).Enabled = True
            Else
                DTPicker1(0).Enabled = False
                DTPicker1(1).Enabled = False
            End If
            Text2(0).Enabled = False
            Text2(1).Enabled = False
            Text2(0).BackColor = vbButtonFace    'RGB(192, 192, 192)
            Text2(1).BackColor = vbButtonFace    'RGB(192, 192, 192)
        Case 1
            If Option2(1).Value = True Then
                Text2(0).Enabled = True
                UpDown1(0).Enabled = True
                Text2(0).BackColor = vbWindowBackground    'RGB(255, 255, 255)
            Else
                Text2(0).Enabled = False
                Text2(0).BackColor = vbButtonFace          'RGB(192, 192, 192)
                UpDown1(0).Enabled = False
            End If
            DTPicker1(0).Enabled = False
            DTPicker1(1).Enabled = False
            Text2(1).Enabled = False
            Text2(1).BackColor = vbButtonFace              'RGB(192, 192, 192)
        Case 2
            If Option2(2).Value = True Then
                Text2(1).Enabled = True
                Text2(1).BackColor = vbWindowBackground    'RGB(255, 255, 255)
                UpDown1(1).Enabled = True
            Else
                Text2(1).Enabled = False
                Text2(1).BackColor = vbButtonFace          'RGB(192, 192, 192)
                UpDown1(1).Enabled = False
            End If
            DTPicker1(0).Enabled = False
            DTPicker1(1).Enabled = False
            Text2(0).Enabled = False
            Text2(0).BackColor = vbButtonFace              'RGB(192, 192, 192)
    End Select

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57   ' Allow digits
        Case 8      ' Allow backspace
        Case Else
            KeyAscii = 0: Beep
    End Select
End Sub
Private Sub UpDown1_DownClick(Index As Integer)
    If Index = 0 Then
        If Val(Text2(0).Text) <= 1 Then
            Text2(0).Text = 1
            Beep
        Else
            Text2(0).Text = Val(Text2(0).Text) - 1
        End If
    Else
        If Val(Text2(1).Text) <= 1 Then
            Text2(1).Text = 1
            Beep
        Else
            Text2(1).Text = Val(Text2(1).Text) - 1
        End If
    End If
End Sub

Private Sub UpDown1_UpClick(Index As Integer)
    If Index = 0 Then Text2(0).Text = Text2(0).Text + 1 Else Text2(1).Text = Text2(1).Text + 1
End Sub

Private Function DateCondition(TableToUse As String) As String
    Dim strDateTab As String
    Dim strDate1 As String
    Dim strDate2 As String
    Dim dTempDate As Date
    
    DateCondition = " [Tree ID] <> 'XXX' "
    
    If optFind.Value = True Then
        Select Case Combo5(1).ListIndex
            Case 0
                strDateTab = "[DATE LAST MODIFIED]"
            Case 1
                strDateTab = "[DATE CREATED]"
            Case 2
                strDateTab = "[DATE SEND]"
        End Select
        
        If Option2(0).Value Then    ' between
            strDate1 = CStr(DTPicker1(0).Value + #12:00:00 AM#)
            strDate2 = CStr(DTPicker1(1).Value + Time)
            
            If InStr(1, strDate2, CStr(Date)) = 0 Then
                strDate2 = CStr(DTPicker1(1).Value + #11:59:59 PM#)
            End If
        ElseIf Option2(1).Value Then   ' during the previous n months
            'dTempDate = Now - (CInt(Text2(0).Text) * 30)
            strDate2 = Now
            strDate1 = DateAdd("m", -CInt(Text2(0).Text), strDate2)
        ElseIf Option2(2).Value Then    ' during the previous n days
            'dTempDate = Date - CInt(Text2(1).Text)
            strDate2 = Now
            strDate1 = DateAdd("d", -CInt(Text2(1).Text), strDate2)
        End If
        
        DateCondition = strDateTab & " >= CDate('" & strDate1 & "') AND " & strDateTab & " <= CDate('" & strDate2 & "')"
    End If
End Function

Private Sub SearchInSubfolders(MainDataTable As String, ParamArray bytDocType())
    Dim strCommand As String
    Dim rstTemp As ADODB.Recordset
    Dim i As Integer
    Dim strTreeID As String

    For i = 0 To UBound(bytDocType)
        
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "* "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & MainDataTable & " "
            strCommand = strCommand & "INNER JOIN "
            strCommand = strCommand & "TEMPLATETREELINKS "
            strCommand = strCommand & "ON "
            strCommand = strCommand & MainDataTable & ".[TREE ID] = TEMPLATETREELINKS.[TREE ID] "
            strCommand = strCommand & "WHERE "
            strCommand = strCommand & "TEMPLATETREELINKS.[TREE ID] <> 'TE' "
            strCommand = strCommand & "AND "
            strCommand = strCommand & "DType = " & CStr(bytDocType(i)) & " "
            strCommand = strCommand & "ORDER BY "
            strCommand = strCommand & "TEMPLATETREELINKS.[TREE ID] "
        ADORecordsetOpen strCommand, g_conData, rstTemp, adOpenKeyset, adLockOptimistic
        
        'With rstOfflineTemp
        
            Do While Not rstTemp.EOF
                If mblnCancel = True Then GoTo Cancelled1
                If strTreeID <> rstTemp.Fields("TEMPLATETREELINKS.TREE ID").Value Then
    
                    Select Case bytDocType(i)
                        Case 1
                            SearchIn "Import", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 2
                            SearchIn "Export", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 3
                            SearchIn "Transit", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 7
                            SearchIn "NCTS", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 9
                            SearchIn "Combined NCTS", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 11
                            SearchInEDI CByte(bytDocType(i)), 5, "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 12
                            'SearchIn "Import", bytDocType(i), "", rstTemp![TREE ID]
                            SearchInEDI CByte(bytDocType(i)), 2, "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                            SearchInEDI CByte(bytDocType(i)), 11, "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 14
                            SearchIn "PLDA Import", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                        Case 18
                            SearchIn "PLDA Combined", CByte(bytDocType(i)), "", rstTemp.Fields(MainDataTable & ".TREE ID").Value
                    
                    End Select
                    
                End If
                
                strTreeID = rstTemp.Fields("TEMPLATETREELINKS.TREE ID").Value
                
                rstTemp.MoveNext
            Loop
        '#####
        'End With
        
Cancelled1:
    
        ADORecordsetClose rstTemp
    Next
    
    
End Sub

Private Sub SearchAccdngToDocType(DocType As String, strSingleCharType As String, bytDocType As Byte)
    Dim strYear As String
    Dim lngCounter As Long
    Dim strHistoryDBFile As String

    Select Case icbLookIn.SelectedItem.Key
        Case "D" & enuLookIn.eLookInAll
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "30ED", "31ED", "32ED", "33ED", "34ED", "35ED", "36ED", "37ED", "39ED", "40ED", "41ED", "-1ED"
                    'Added by Rachelle on 06/16/2005
                    
                    If mblnCancel = True Then Exit Sub
                    
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            
                            If mblnCancel = True Then Exit Sub
                            
                            If IsNumeric(Mid(File1.List(lngCounter), 15, 2)) And Right(File1.List(lngCounter), 3) = "mdb" Then
                                SearchInEDI 11, 5, Mid(File1.List(lngCounter), 15, 2), "40ED"
                            End If
                        Next lngCounter
                    End If
                    SearchInSubfolders "MASTEREDINCTS", bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "43ED", "44ED", "45ED", "46ED", "47ED", "48ED", "49ED", "50ED", "51ED", "-2ED"
                    
                    SearchInEDI bytDocType, 11, "", "43ED", "44ED", "45ED", "46ED", "47ED", "48ED", "49ED", "50ED", "51ED", "-2ED"
                   
                    If mblnCancel = True Then Exit Sub
                    
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            
                            If mblnCancel = True Then Exit Sub
                            
                            If IsNumeric(Mid(File1.List(lngCounter), 15, 2)) And Right(File1.List(lngCounter), 3) = "mdb" Then
                                SearchInEDI 12, 2, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                SearchInEDI 12, 11, Mid(File1.List(lngCounter), 15, 2), "50ED"
                            End If
                        Next lngCounter
                    End If
                    SearchInSubfolders "MASTEREDINCTS2", bytDocType
                    SearchInSubfolders "MASTEREDINCTSIE44", bytDocType
                    
                    'csclp-440
                    TransferToListView
                    CreateOfflineRecordset (True)
                    
                Case 1, 2, 3
                    SearchIn DocType, bytDocType, "", "SD" & strSingleCharType & "1", "SD" & strSingleCharType & "2", "DD", "WL1", "WL2", "DE", "TE"
                    
                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")
                    Do Until strHistoryDBFile = ""
                        strHistoryDBFile = Dir()    ' Get next History DB file.
        
                        If strHistoryDBFile <> "" Then
                            strYear = Left(Right(strHistoryDBFile, 6), 2)
                            Select Case bytDocType
                                Case 1
                                    SearchIn "IMPORT", 1, strYear, "HI" & strYear & "I"
                                    If mblnCancel = True Then Exit Sub
                                Case 2
                                    SearchIn "EXPORT", 2, strYear, "HI" & strYear & "E"
                                    If mblnCancel = True Then Exit Sub
                                Case 3
                                    SearchIn "TRANSIT", 3, strYear, "HI" & strYear & "T"
                                    If mblnCancel = True Then Exit Sub
                            End Select
                        End If
                    Loop
                    
                    SearchInSubfolders "MASTER", bytDocType
                Case 7, 9
                    SearchIn DocType, bytDocType, "", "SD" & strSingleCharType & "1", "SD" & strSingleCharType & "2", "DD", "WL1", "WL2", "DE", "TE"
                    
                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")
                    Do Until strHistoryDBFile = ""
                        strHistoryDBFile = Dir()    ' Get next History DB file.
        
                        If strHistoryDBFile <> "" Then
                            strYear = Left(Right(strHistoryDBFile, 6), 2)
                            Select Case bytDocType
                                Case 7
                                    SearchIn "NCTS", 7, strYear, "HI" & strYear & "N"
                                    If mblnCancel = True Then Exit Sub
                                    
                                Case 9
                                    SearchIn "COMBINED NCTS", 9, strYear, "HI" & strYear & "C"
                                    If mblnCancel = True Then Exit Sub
                            End Select
                        End If
                    Loop
                    
                    SearchInSubfolders "MASTERNCTS", bytDocType
            
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "SD" & strSingleCharType & "1", "SD" & strSingleCharType & "2", "SD" & strSingleCharType & "3", "DD", "WL1", "WL2", "DE", "TE", "P" & strSingleCharType & "01"
                    
                    'Added by BCo 2006-08-31
                    'Search in all available history DBs for PLDA
                    Dim arrHistories() As String
                    
                    strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")                  'Get first history filename
                    Do Until Right$(strHistoryDBFile, 1) = ","                              'Process until Dir() is empty
                        strHistoryDBFile = strHistoryDBFile & "," & Dir()                   'CSV history filename
                    Loop
                    strHistoryDBFile = Mid$(strHistoryDBFile, 1, Len(strHistoryDBFile) - 1) 'Erase marker used with Dir()
                    arrHistories = Split(strHistoryDBFile, ",")                             'Partition results
                    
                    For lngCounter = LBound(arrHistories) To UBound(arrHistories)           'Span across array
                        strYear = Mid$(arrHistories(lngCounter), 12, 2)                     'Get 2-digit year from history filename

                        Select Case bytDocType
                            Case 14
                                SearchIn "PLDA IMPORT", 14, strYear, "HI" & strYear & "X"
                                If mblnCancel = True Then Exit Sub

                            Case 18
                                SearchIn "PLDA COMBINED", 18, strYear, "HI" & strYear & "Z"
                                If mblnCancel = True Then Exit Sub
                        End Select
                    Next
                    
                    SearchInSubfolders "MASTERPLDA", bytDocType
            
            End Select
        Case "D" & enuLookIn.eLookInApproved
            If bytDocType = 11 Then
                SearchInEDI bytDocType, 5, "", "34ED", "35ED", "36ED"
            ElseIf bytDocType = 12 Then
                SearchInEDI bytDocType, 2, "", "47ED", "48ED"
                SearchInEDI bytDocType, 11, "", "47ED", "48ED"
            Else
                SearchIn DocType, bytDocType, "", "SD" & strSingleCharType & "1", "SD" & strSingleCharType & "2"
            End If
        
        Case "D" & enuLookIn.eLookInDeleted
            If bytDocType = 11 Then
                SearchInEDI bytDocType, 5, "", "-1ED"
            ElseIf bytDocType = 12 Then
                SearchInEDI bytDocType, 2, "", "-2ED"
                SearchInEDI bytDocType, 11, "", "-2ED"
            Else
        
                SearchIn DocType, bytDocType, "", "DD"
            End If
        
        Case "D" & enuLookIn.eLookInDRAFTS
            If bytDocType = 11 Then
                SearchInEDI bytDocType, 5, "", "30ED"
            ElseIf bytDocType = 12 Then
                SearchInEDI bytDocType, 2, "", "43ED"
                SearchInEDI bytDocType, 11, "", "43ED"
            Else
        
                SearchIn DocType, bytDocType, "", "WL1"
            End If
            
        Case "D" & enuLookIn.eLookInOutbox
            If bytDocType = 11 Then
                SearchInEDI bytDocType, 5, "", "31ED", "32ED"
            ElseIf bytDocType = 12 Then
                SearchInEDI bytDocType, 2, "", "44ED", "45ED"
                SearchInEDI bytDocType, 11, "", "44ED", "45ED"
            Else
        
                SearchIn DocType, bytDocType, "", "WL2"
            End If
            
        Case "D" & enuLookIn.eLookInRejected
            If bytDocType = 11 Then
                SearchInEDI bytDocType, 5, "", "33ED", "37ED"
            ElseIf bytDocType = 12 Then
                SearchInEDI bytDocType, 2, "", "46ED"
                SearchInEDI bytDocType, 11, "", "46ED"
            Else
                SearchIn DocType, bytDocType, "", "DE"
            End If
        
        Case "D" & enuLookIn.eLookInTemplates
            
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "41ED"
                    SearchInSubfolders "MASTEREDINCTS", bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "51ED"
                    SearchInEDI bytDocType, 11, "", "51ED"
                    SearchInSubfolders "MASTEREDINCTS2", bytDocType
                    SearchInSubfolders "MASTEREDINCTSIE44", bytDocType
                Case 1, 2, 3
                    SearchIn DocType, bytDocType, "", "TE"
                    SearchInSubfolders "MASTER", bytDocType
                Case 7, 9
                    SearchIn DocType, bytDocType, "", "TE"
                    SearchInSubfolders "MASTERNCTS", bytDocType
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "TE"
                    SearchInSubfolders "MASTERPLDA", bytDocType
                    
            End Select
            
        'CSCLP-248
        '-------------------------------------------------
        Case "D" & enuLookIn.eLookInToBePrinted                     'SADBEL IET, NCTS, Combined NCTS
            
            Select Case bytDocType
                Case 1, 2, 3, 7, 9
                    SearchIn DocType, bytDocType, "", "SDI2", "SDE2", "SDT2", "SDN2", "SDC2"
            End Select
            
        Case "D" & enuLookIn.eLookInReleased                        'PLDA IE
            
            Select Case bytDocType
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "SDX2", "SDZ2"
            End Select
            
        Case "D" & enuLookIn.eLookInSent                            'Departure, PLDA IE
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "32ED"
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "PX01", "PZ01"
            End Select
            
        Case "D" & enuLookIn.eLookInEmergencyProcedure              'PLDA IE
        
            Select Case bytDocType
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "EP"
            End Select
            
        Case "D" & enuLookIn.eLookInExitEC                          'PLDA E
        
            Select Case bytDocType
                Case 18
                    SearchIn DocType, bytDocType, "", "SDZ4"
            End Select
            
        Case "D" & enuLookIn.eLookInGuarantee                       'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "55ED"
            End Select
            
        Case "D" & enuLookIn.eLookInUnderControl                    'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "35ED"
            End Select
            
        Case "D" & enuLookIn.eLookInReleases                        'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "36ED"
            End Select
            
        Case "D" & enuLookIn.eLookInReleaseRejected                 'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "37ED"
            End Select
            
        Case "D" & enuLookIn.eLookInCancelled                       'PLDA IE, Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "39ED"
                Case 14, 18
                    SearchIn DocType, bytDocType, "", "SDX3", "SDZ3"
            End Select
            
        Case "D" & enuLookIn.eLookInWrittenOff                      'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "40ED"
            End Select
            
        Case "D" & enuLookIn.eLookInAmendmentSent                   'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "52ED"
            End Select
            
        Case "D" & enuLookIn.eLookInAmendmentRejected               'Departure
        
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "53ED"
            End Select
            
        Case "D" & enuLookIn.eLookInAmendmentAccepted               'Departure
         
            Select Case bytDocType
                Case 11
                    SearchInEDI bytDocType, 5, "", "54ED"
            End Select
        
        Case "D" & enuLookIn.eLookInArchives                        'Arrival
        
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "50ED"
                    SearchInEDI bytDocType, 11, "", "50ED"
            End Select
                
        Case "D" & enuLookIn.eLookInArrivalNotificationSent         'Arrival
        
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "45ED"
                    SearchInEDI bytDocType, 11, "", "45ED"
            End Select
        
        Case "D" & enuLookIn.eLookInArrivalNotificationRejected     'Arrival
         
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "46ED"
                    SearchInEDI bytDocType, 11, "", "46ED"
            End Select
            
        Case "D" & enuLookIn.eLookInUnloadingPermitted              'Arrival
         
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "47ED"
                    SearchInEDI bytDocType, 11, "", "47ED"
            End Select
            
        Case "D" & enuLookIn.eLookInUnloadingRemarksSent            'Arrival
        
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "48ED"
                    SearchInEDI bytDocType, 11, "", "48ED"
            End Select
        
        Case "D" & enuLookIn.eLookInUnloadingRemarksRejected        'Arrival
        
            Select Case bytDocType
                Case 12
                    SearchInEDI bytDocType, 2, "", "49ED"
                    SearchInEDI bytDocType, 11, "", "49ED"
            End Select
        
        '-------------------------------------------------
         
         
         Case "D" & enuLookIn.eLookInArchive  'allan nov27
        
            '==== Archive ===='
            'Added by BCo 2006-08-31
            'Search in all available mdb_history DBs
                                
            strHistoryDBFile = Dir(cAppPath & "\mdb_history*.mdb")                  'Get first history filename
            Do
                strHistoryDBFile = strHistoryDBFile & "," & Dir()                   'CSV history filename
            Loop Until Right$(strHistoryDBFile, 1) = ","                            'Process until Dir() is empty
            strHistoryDBFile = Mid$(strHistoryDBFile, 1, Len(strHistoryDBFile) - 1) 'Erase marker used with Dir()
            arrHistories = Split(strHistoryDBFile, ",")                             'Partition results
            
            For lngCounter = LBound(arrHistories) To UBound(arrHistories)           'Span across array
                strYear = Mid$(arrHistories(lngCounter), 12, 2)                     'Get 2-digit year from history filename
                    
                    Select Case DocType
                    Case "IMPORT"
                        SearchIn "IMPORT", 1, strYear, "HI" & strYear & "I"
                        If mblnCancel = True Then Exit Sub
                    Case "EXPORT"
                        SearchIn "EXPORT", 2, strYear, "HI" & strYear & "E"
                        If mblnCancel = True Then Exit Sub
                    Case "TRANSIT"
                        SearchIn "TRANSIT", 3, strYear, "HI" & strYear & "T"
                        If mblnCancel = True Then Exit Sub
                    
                    Case "NCTS"
                        SearchIn "NCTS", 7, strYear, "HI" & strYear & "N"
                        If mblnCancel = True Then Exit Sub
                    
                    Case "EDI DEPARTURES"
 
                        Dim lnCounterNCTS As Long
                        If blnEDIHistoryExisting = True Then
                            For lnCounterNCTS = 0 To UBound(g_conEDIHistory)
                                If IsNumeric(Mid(File1.List(lnCounterNCTS), 15, 2)) And _
                                    Right(File1.List(lnCounterNCTS), 3) = "mdb" Then
                                    If Mid(File1.List(lnCounterNCTS), 15, 2) <> Right(Year(Now), 2) Then
                                        SearchInEDI 11, 5, Mid(File1.List(lnCounterNCTS), 15, 2), "40ED"
                                    End If
                                End If
                            Next lnCounterNCTS
                        End If
                        
                        
                        If mblnCancel = True Then Exit Sub
                    
                    Case "EDI ARRIVALS"
                      
                        If blnEDIHistoryExisting = True Then
                            For lnCounterNCTS = 0 To UBound(g_conEDIHistory)
                                If IsNumeric(Mid(File1.List(lnCounterNCTS), 15, 2)) And _
                                    Right(File1.List(lnCounterNCTS), 3) = "mdb" Then
                                    If Mid(File1.List(lnCounterNCTS), 15, 2) <> Right(Year(Now), 2) Then
                                        SearchInEDI 12, 2, Mid(File1.List(lnCounterNCTS), 15, 2), "50ED"
                                        SearchInEDI 12, 11, Mid(File1.List(lnCounterNCTS), 15, 2), "50ED"
                                    End If
                                End If
                            Next lnCounterNCTS
                        End If
                        
                        
                        If mblnCancel = True Then Exit Sub
                    
                    Case "COMBINED NCTS"
                        SearchIn "COMBINED NCTS", 9, strYear, "HI" & strYear & "C"
                        If mblnCancel = True Then Exit Sub
                    Case "PLDA IMPORT"
                        SearchIn "PLDA IMPORT", 14, strYear, "HI" & strYear & "X"
                        If mblnCancel = True Then Exit Sub
                    Case "PLDA COMBINED"
                        SearchIn "PLDA COMBINED", 18, strYear, "HI" & strYear & "Z"
                        If mblnCancel = True Then Exit Sub
                    End Select
            Next
            
            
        Case Else
            If bytDocType <> 11 And bytDocType <> 12 Then
                strYear = Mid(CStr(icbLookIn.SelectedItem.Key), 4)
                SearchIn DocType, bytDocType, strYear, "HI" & strYear & strSingleCharType
            Else
                'For EDI Archive
                'Added by Rachelle on 06/16/2005
                If blnEDIHistoryExisting = True Then
                    strYear = Mid(CStr(icbLookIn.SelectedItem.Key), 4)
                    For lngCounter = 0 To UBound(g_conEDIHistory)
                        
                        If mblnCancel = True Then Exit Sub
                        
                        If Mid(File1.List(lngCounter), 15, 2) = strYear And Right(File1.List(lngCounter), 3) = "mdb" Then
                            If bytDocType = 11 Then
                                SearchInEDI 11, 5, Mid(File1.List(lngCounter), 15, 2), "40ED"
                            ElseIf bytDocType = 12 Then
                                SearchInEDI 12, 2, Mid(File1.List(lngCounter), 15, 2), "50ED"
                                SearchInEDI 12, 11, Mid(File1.List(lngCounter), 15, 2), "50ED"
                            End If
                        End If
                    Next lngCounter
                End If
            End If
    End Select

End Sub

Private Sub CreateOfflineRecordset(Optional ByVal EnsureNothing As Boolean = True)
    Dim strFieldList() As String
    Dim lngCounter As Long
    Dim strSelectedKey As String
    
    If EnsureNothing = True Then
        If Not rstOfflineTemp Is Nothing Then
            Set rstOfflineTemp = Nothing
        End If
    End If
    
    If blnJustLoaded2 = True Then
        blnJustLoaded2 = False
        strSelectedKey = "D" & enuDocType.eDocAny
    Else
        strSelectedKey = icbType.SelectedItem.Key
    End If
    
    Select Case strSelectedKey
        Case "D" & enuDocType.edocimport: strFieldList = Split(DOC_IMPORT, "**")
            
        Case "D" & enuDocType.eDocExport: strFieldList = Split(DOC_EXPORT, "**")
        
        Case "D" & enuDocType.eDocOTS: strFieldList = Split(DOC_OTS, "**")
        
        Case "D" & enuDocType.eDocNCTS: strFieldList = Split(DOC_SADBELNCTS, "**")
        
        Case "D" & enuDocType.edoccombined: strFieldList = Split(DOC_COMBINEDNCTS, "**")
        
        Case "D" & enuDocType.eDocEDIDepartures: strFieldList = Split(DOC_EDINCTS, "**")
        
        Case "D" & enuDocType.eDocEDIARRIVALS: strFieldList = Split(DOC_EDINCTS2, "**")
        
        Case "D" & enuDocType.eDocPLDAImport: strFieldList = Split(DOC_PLDAIMPORT, "**")
        
        Case "D" & enuDocType.eDocPLDACombined: strFieldList = Split(DOC_PLDACOMBINED, "**")
        
        Case "D" & enuDocType.eDocAny: strFieldList = Split(DOC_ANY, "**")
    
    End Select
    If EnsureNothing = True Then
        Set rstOfflineTemp = New ADODB.Recordset
        
        rstOfflineTemp.Fields.Append "CODE", adVarChar, 25            '--> Unique Code
        rstOfflineTemp.Fields.Append "NAME", adVarChar, 50            '--> Document Name
        rstOfflineTemp.Fields.Append "DOCUMENT", adVarChar, 15        '--> Import, Export, Transit, etc.
        rstOfflineTemp.Fields.Append "IN FOLDER", adVarChar, 100      '--> Approved and Printed, Rejected, Drafts, etc.
        rstOfflineTemp.Fields.Append "USERNAME", adVarChar, 25        '--> Username
        rstOfflineTemp.Fields.Append "DATE MODIFIED", adDate          '--> Date Modified
        rstOfflineTemp.Fields.Append "ARCHIVE DATE", adVarChar, 4        '--> If from Archive, what year
        If Not (Right(strSelectedKey, 1) = enuDocType.eDocAny Or _
            Right(strSelectedKey, 1) = enuDocType.eDocNCTS Or _
            Right(strSelectedKey, 1) = enuDocType.eDocEDIDepartures Or _
            Right(strSelectedKey, 1) = enuDocType.eDocEDIARRIVALS Or _
            Right(strSelectedKey, 1) = enuDocType.eDocPLDAImport Or _
            Right(strSelectedKey, 1) = enuDocType.eDocPLDACombined) Then
            rstOfflineTemp.Fields.Append Translate(437), adVarChar, 50
        End If
        rstOfflineTemp.Fields.Append Translate(713), adVarChar, 50
        rstOfflineTemp.Fields.Append Translate(715), adVarChar, 50
        rstOfflineTemp.Fields.Append Translate(742), adVarChar, 50
        If Right(strSelectedKey, 1) = enuDocType.eDocEDIDepartures Or _
            Right(strSelectedKey, 1) = enuDocType.eDocEDIARRIVALS Then
            rstOfflineTemp.Fields.Append "Date Last Received", adVarChar, 50
        End If
        rstOfflineTemp.Fields.Append "Date Printed", adVarChar, 50
        rstOfflineTemp.Fields.Append "LogID Description", adVarChar, 100
        rstOfflineTemp.Fields.Append "Error String", adVarChar, 100
        rstOfflineTemp.Fields.Append Translate(423), adVarChar, 100
        If Right(strSelectedKey, 1) = enuDocType.eDocNCTS Or _
            Right(strSelectedKey, 1) = enuDocType.edoccombined Or _
            Right(strSelectedKey, 1) = enuDocType.eDocEDIDepartures Or _
            Right(strSelectedKey, 1) = enuDocType.eDocEDIARRIVALS Or _
            Right(strSelectedKey, 1) = enuDocType.eDocPLDAImport Or _
            Right(strSelectedKey, 1) = enuDocType.eDocPLDACombined Then
            rstOfflineTemp.Fields.Append "MRN", adVarChar, 50
        End If
        For lngCounter = 1 To UBound(strFieldList)
            rstOfflineTemp.Fields.Append strFieldList(lngCounter), adVarChar, 100
        Next lngCounter
        
        rstOfflineTemp.Open
    End If
End Sub

Private Sub TransferToListView()
    Dim itmDocumentFound As MSComctlLib.ListItem
    Dim itmDocumentDetails As MSComctlLib.ListSubItem
    Dim lngCounter As Long
    
    Dim lngImageIndex As Long
    
    If Not mblnCancel Then
        With rstOfflineTemp
            If rstOfflineTemp.RecordCount > 0 Then
                .MoveFirst
                
                SSActiveToolBars1.Tools("ID_Open").Enabled = True
                SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = True
            Else
                If lvwItemsFound.ListItems.Count > 0 Then
                    SSActiveToolBars1.Tools("ID_Open").Enabled = True
                    SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = True
                Else
                    SSActiveToolBars1.Tools("ID_Open").Enabled = False
                    SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = False
                End If
                
            End If
            
            Do Until .EOF
                lngImageIndex = GetImageIndex(.Fields("DOCUMENT").Value)
              
                Set itmDocumentFound = lvwItemsFound.ListItems.Add(, , .Fields("NAME").Value, lngImageIndex, lngImageIndex)
                itmDocumentFound.Tag = .Fields("CODE").Value
                
                Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(, , DocumentType(.Fields("DOCUMENT").Value))
                itmDocumentDetails.Tag = .Fields("DOCUMENT").Value
                
                Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(, , FolderName(.Fields("IN FOLDER").Value))
                itmDocumentDetails.Tag = .Fields("IN FOLDER").Value
                
                Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(, , .Fields("USERNAME").Value)
                
                Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(, , .Fields("DATE MODIFIED").Value)
                
                ' Format() pads numeric equivalent of date with zeros both to the left and right of decimal point
                ' Padding is necessary because ListView uses string-type sorting
                Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(, , Format(CDbl(.Fields("DATE MODIFIED").Value), "00000000.0000000000"))
                
                itmDocumentFound.ListSubItems.Item(5).Tag = .Fields("ARCHIVE DATE").Value
                
                For lngCounter = 7 To .Fields.Count - 1
                    If .Fields(lngCounter).Name = Translate(423) Then
                        Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(lngCounter - 1, , Translate(.Fields(.Fields(lngCounter).Name).Value))
                    Else
                        Set itmDocumentDetails = itmDocumentFound.ListSubItems.Add(lngCounter - 1, , .Fields(.Fields(lngCounter).Name).Value)
                    End If
                Next lngCounter
                
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Function FolderName(TreeID As String) As String

    Select Case TreeID
        
        '============= Sadbel ====================
        Case "SDI1", "SDE1", "SDT1", "SDN1", "SDC1", "SDX1", "SDZ1"
            FolderName = Translate(967)
        Case "SDI2", "SDE2", "SDT2", "SDN2", "SDC2"
            FolderName = Translate(902)
        Case "WL1"  'drafts
            FolderName = Translate(969)
        Case "WL2"
            FolderName = Translate(970)
        Case "DD"
            FolderName = Translate(348)
        Case "DE"
            FolderName = Translate(345)
        '=========================================
            
        '============== Departures ===============
        Case "30ED"
            FolderName = Translate(969)
        Case "31ED"
            FolderName = Translate(970)
        Case "32ED", "PX01", "PZ01" 'Sent
            FolderName = Translate(1386)
        Case "33ED"
            FolderName = Translate(345)
        Case "34ED"
            FolderName = Translate(1374)
        Case "35ED"
            FolderName = Translate(1375)
        Case "36ED", "SDX2", "SDZ2" 'Released
            FolderName = Translate(1376)
        Case "37ED"
            FolderName = Translate(1377)
        Case "39ED", "SDX3", "SDZ3" 'Cancelled
            FolderName = Translate(1378)
        Case "40ED"
            FolderName = Translate(1379)
        Case "41ED"
            FolderName = Translate(347)
        Case "-1ED"
            FolderName = Translate(1386)
            
        '==========================================
        
        '============== Arrivals ==================
        Case "43ED"
            FolderName = Translate(969)
        Case "44ED"
            FolderName = Translate(970)
        Case "45ED"
            FolderName = Translate(1381)
        Case "46ED"
            FolderName = Translate(1382)
        Case "47ED"
            FolderName = Translate(1383)
        Case "48ED"
            FolderName = Translate(1384)
        Case "49ED"
            FolderName = Translate(1385)
        Case "50ED"
            FolderName = Translate(757)
        Case "51ED"
            FolderName = Translate(347)
        Case "-2ED"
            FolderName = Translate(1386)

        '==========================================
        
        Case Else
            If Left(TreeID, 2) = "HI" Then
                FolderName = Translate(1074)
            ElseIf IsNumeric(TreeID) Then
                FolderName = TemplateName(TreeID)
            End If
    End Select
    
    'If IsNumeric(TreeID) Or TreeID = "TE" Then  'templates
    If TreeID = "TE" Then   'templates
        FolderName = Translate(347)
    End If
    

End Function


Private Function DocumentType(bytDocType As Byte) As String

    Select Case bytDocType
        Case 1
            DocumentType = "Import"
        Case 2
            DocumentType = "Export"
        Case 3
            DocumentType = "Transit"
        Case 7
            DocumentType = "NCTS"
        Case 9
            DocumentType = "Combined NCTS"
        Case 11
            DocumentType = "EDI Departures"
        Case 12
            DocumentType = "EDI Arrivals"
        Case 14
            DocumentType = "PLDA Import"
        Case 18
            DocumentType = "PLDA Combined"
    End Select
End Function

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Dim nIdx As Integer
    Dim lngCounter As Long
    Dim strToPass As String
    Dim strAppPath As String
    
    If lvwItemsFound.ListItems.Count > 0 Then
        clsFindForm.SelectedItemTag = lvwItemsFound.SelectedItem.Tag
        clsFindForm.SelectedItemText = lvwItemsFound.SelectedItem.Text
        clsFindForm.ListSubItems = lvwItemsFound.SelectedItem.ListSubItems(2).Tag
        clsFindForm.SubItems = lvwItemsFound.SelectedItem.SubItems(2)
    End If
    
    Select Case Tool.ID
        Case "ID_File"
            'Added by BCo 2006-05-03
            'Licensing checking when clicking File menu, before clicking list view
            If Not lvwItemsFound.SelectedItem Is Nothing Then
                SSActiveToolBars1.Tools("ID_OpenACopy").Enabled = IsFeatureLicensed(lvwItemsFound.SelectedItem.ListSubItems(1).Text)
            End If
        Case "ID_SADBELplusHelpTopics"
            strAppPath = GetSetting("ClearingPoint", "Settings", "AppPath")
            App.HelpFile = Trim(strAppPath) & "\" & "CLEARINGPOINT USERS GUIDE.HLP"
            
            SendKeysEx "{F1}"
            
        Case "ID_About"
            g_clsAbout.Show Me, vbModal
            
        Case "ID_Close": Unload Me
        Case "ID_byName": Index_Column Translate(496)
        Case "ID_byDocument": Index_Column Translate(635)
        Case "ID_byFolder": Index_Column Translate(625)
        Case "ID_byOwner": Index_Column Left(Translate(272), Len(Translate(272)) - 1)
        Case "ID_byDate": Index_Column Translate(611)
        
        Case "ID_LargeIcons": ListCheck "ID_LargeIcons"
        Case "ID_SmallIcons": ListCheck "ID_SmallIcons"
        Case "ID_List": ListCheck "ID_List"
        Case "ID_Details": ListCheck "ID_Details"
        
        Case "ID_Import": 'DoOpen "ID_Import", True
            CallingForm.LoadDocument edocimport, True
            
        Case "ID_Export": 'DoOpen "ID_Export", True
            CallingForm.LoadDocument eDocExport, True
            
        Case "ID_Transit": 'DoOpen "ID_Transit", True
            CallingForm.LoadDocument eDocOTS, True
            
        Case "ID_NCTS"
            CallingForm.LoadDocument eDocNCTS, True
            
        Case "ID_Combined"
            CallingForm.LoadDocument edoccombined, True
            
        Case "ID_EDIDepartures"
            CallingForm.LoadDocument eDocEDIDepartures, True
            
        Case "ID_EDIArrivals"
            CallingForm.LoadDocument eDocEDIARRIVALS, True
        
        Case "ID_PLDAImport"
            CallingForm.LoadDocument eDocPLDAImport, True
            
        Case "ID_PLDACombined"
            CallingForm.LoadDocument eDocPLDACombined, True
            
        Case "ID_Open"
            clsFindForm.OpenOnly = True
            lvwItemsFound_DblClick
            clsFindForm.OpenOnly = False
        Case "ID_Delete"
            DeleteItems
        Case "ID_Copy"
        Case "ID_SelectAll"
            'For i = 1 To ListView1.ListItems.Count
            '    ListView1.ListItems(i).Selected = True
            'Next
        Case "ID_Rename"
        
'            Dim rstToUse As DAO.Recordset
'            If UCase(ListView1.SelectedItem.ListSubItems(1).Text) = "IMPORT" Or _
'                UCase(ListView1.SelectedItem.ListSubItems(1).Text) = "EXPORT" Or _
'                UCase(ListView1.SelectedItem.ListSubItems(1).Text) = "TRANSIT" Then
'
'                Set rstToUse = rstMaster(0)
'            Else
'                Set rstToUse = rstMaster(1)
'            End If
'
'
'            rstToUse.Index = "code"
'            rstToUse.Seek "=", Trim(ListView1.SelectedItem.Tag)
'            If rstToUse.NoMatch Then
'            Else
'                Dim nn As Byte, cNewVal As String
'                nn = rstToUse!DType
'                cNewVal = RenameItem(ListView1.SelectedItem.Tag, ListView1.SelectedItem.Text, nn)
'                EditTimeImport
'                If Len(Trim(cNewVal)) > 0 Then ListView1.SelectedItem.Text = cNewVal
'            End If
        
        Case "ID_OpenaCopy"
            OpenACopy
        Case "ID_ShowFields"
            strToPass = ""
            For lngCounter = 1 To lvwItemsFound.ColumnHeaders.Count
                If lvwItemsFound.ColumnHeaders(lngCounter).Width <> 0 Then
                    strToPass = strToPass & "*" & lvwItemsFound.ColumnHeaders(lngCounter).Text
                End If
            Next lngCounter
            ShowFields strToPass
            blnShowFields = True
            RefreshList
    End Select
    
End Sub

Private Function GetImageIndex(bytDocType As Byte) As Long
    On Error Resume Next
    Select Case bytDocType
        Case 1
            GetImageIndex = imgImages.ListImages.Item("Import").Index
        Case 2
            GetImageIndex = imgImages.ListImages.Item("Export").Index
        Case 3
            GetImageIndex = imgImages.ListImages.Item("Transit").Index
        Case 7
            GetImageIndex = imgImages.ListImages.Item("NCTS").Index
        Case 9
            GetImageIndex = imgImages.ListImages.Item("Combined").Index
        Case 11
            GetImageIndex = imgImages.ListImages.Item("EDIDepartures").Index
        Case 12
            GetImageIndex = imgImages.ListImages.Item("EDIArrivals").Index
        Case 14
            GetImageIndex = imgImages.ListImages.Item("PLDA Import").Index
        Case 18
            GetImageIndex = imgImages.ListImages.Item("PLDA Export").Index
    End Select
End Function

Private Function GetSQLToUse(ByVal FieldToUse, ByVal bytDocType As Byte, ByVal strTreeID As String, ByVal strSQLWhere As String) As String
    Select Case bytDocType
        Case 1, 2, 3
            GetSQLToUse = "SELECT * FROM [" & FieldToUse & " HEADER] " & _
                            "INNER JOIN ([" & FieldToUse & "] INNER JOIN [" & FieldToUse & " DETAIL] ON [" & FieldToUse & "].CODE = [" & FieldToUse & " DETAIL].CODE) " & _
                            "ON ([" & FieldToUse & " HEADER].CODE = [" & FieldToUse & " DETAIL].CODE) AND ([" & FieldToUse & " HEADER].HEADER = [" & FieldToUse & " DETAIL].HEADER) " & _
                            "WHERE [" & FieldToUse & "].[TREE ID] = " & Chr(39) & ProcessQuotes(strTreeID) & Chr(39) & " " & strSQLWhere & _
                            " ORDER BY [" & FieldToUse & "].[DATE LAST MODIFIED] DESC"
        Case 7
            
            GetSQLToUse = "SELECT *" & _
                            "FROM ((((((NCTS INNER JOIN [NCTS HEADER] ON NCTS.CODE = [NCTS HEADER].CODE) INNER JOIN [NCTS HEADER ZEKERHEID] ON ([NCTS HEADER].HEADER = [NCTS HEADER ZEKERHEID].HEADER) AND ([NCTS HEADER].CODE = [NCTS HEADER ZEKERHEID].CODE))" & _
                            "INNER JOIN [NCTS DETAIL] ON ([NCTS HEADER].HEADER = [NCTS DETAIL].HEADER) AND ([NCTS HEADER].CODE = [NCTS DETAIL].CODE)) INNER JOIN [NCTS DETAIL DOCUMENTEN] ON ([NCTS DETAIL].DETAIL = [NCTS DETAIL DOCUMENTEN].DETAIL) AND ([NCTS DETAIL].HEADER = [NCTS DETAIL DOCUMENTEN].HEADER) AND ([NCTS DETAIL].CODE = [NCTS DETAIL DOCUMENTEN].CODE)) INNER JOIN [NCTS DETAIL BIJZONDERE] ON ([NCTS DETAIL].DETAIL = [NCTS DETAIL BIJZONDERE].DETAIL) AND ([NCTS DETAIL].HEADER = [NCTS DETAIL BIJZONDERE].HEADER) AND ([NCTS DETAIL].CODE = [NCTS DETAIL BIJZONDERE].CODE)) INNER JOIN [NCTS DETAIL COLLI] ON ([NCTS DETAIL].DETAIL = [NCTS DETAIL COLLI].DETAIL) AND ([NCTS DETAIL].HEADER = [NCTS DETAIL COLLI].HEADER) AND ([NCTS DETAIL].CODE = [NCTS DETAIL COLLI].CODE)) INNER JOIN [NCTS DETAIL CONTAINER]" & _
                            "ON ([NCTS DETAIL].DETAIL = [NCTS DETAIL CONTAINER].DETAIL) AND ([NCTS DETAIL].HEADER = [NCTS DETAIL CONTAINER].HEADER) AND ([NCTS DETAIL].CODE = [NCTS DETAIL CONTAINER].CODE)" & _
                            " WHERE [" & FieldToUse & "].[TREE ID] = " & Chr(39) & ProcessQuotes(strTreeID) & Chr(39) & " " & strSQLWhere & _
                            " ORDER BY [" & FieldToUse & "].[DATE LAST MODIFIED] DESC"
        Case 9
            GetSQLToUse = "SELECT  * FROM (((((((([COMBINED NCTS]" & _
                            " INNER JOIN [COMBINED NCTS HEADER] ON [COMBINED NCTS].CODE = [COMBINED NCTS HEADER].CODE) " & _
                            " INNER JOIN [COMBINED NCTS HEADER ZEKERHEID] ON ([COMBINED NCTS HEADER].HEADER = [COMBINED NCTS HEADER ZEKERHEID].HEADER) AND ([COMBINED NCTS HEADER].CODE = [COMBINED NCTS HEADER ZEKERHEID].CODE))" & _
                            " INNER JOIN [COMBINED NCTS DETAIL] ON ([COMBINED NCTS HEADER ZEKERHEID].HEADER = [COMBINED NCTS DETAIL].HEADER) AND ([COMBINED NCTS HEADER ZEKERHEID].CODE = [COMBINED NCTS DETAIL].CODE)) INNER JOIN [COMBINED NCTS DETAIL BIJZONDERE] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL BIJZONDERE].DETAIL)" & _
                            " AND ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL BIJZONDERE].HEADER) AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL BIJZONDERE].CODE)) INNER JOIN [COMBINED NCTS DETAIL CONTAINER] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL CONTAINER].DETAIL) AND ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL CONTAINER].HEADER) " & _
                            " AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL CONTAINER].CODE)) " & _
                            " INNER JOIN [COMBINED NCTS DETAIL DOCUMENTEN] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL DOCUMENTEN].DETAIL) AND " & _
                            " ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL DOCUMENTEN].HEADER) AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL DOCUMENTEN].CODE)) INNER JOIN [COMBINED NCTS DETAIL GEVOELIGE] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL GEVOELIGE].DETAIL) " & _
                            " AND ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL GEVOELIGE].HEADER) AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL GEVOELIGE].CODE)) INNER JOIN [COMBINED NCTS DETAIL GOEDEREN] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL GOEDEREN].DETAIL) " & _
                            " AND ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL GOEDEREN].HEADER) AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL GOEDEREN].CODE)) INNER JOIN [COMBINED NCTS DETAIL COLLI] ON ([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL COLLI].DETAIL) AND ([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL" & _
                            " COLLI].HEADER) AND ([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL COLLI].CODE)" & _
                            " WHERE [" & FieldToUse & "].[TREE ID] = " & Chr(39) & ProcessQuotes(strTreeID) & Chr(39) & " " & strSQLWhere & _
                            " ORDER BY [" & FieldToUse & "].[DATE LAST MODIFIED] DESC"
        Case 14
            GetSQLToUse = "SELECT [PLDA IMPORT HEADER].Code As [Code], [DOCUMENT NAME], [DTYPE], [TREE ID], [USERNAME], [DATE LAST MODIFIED] FROM (((((((((([PLDA IMPORT]" & _
                            " INNER JOIN [PLDA IMPORT HEADER] ON [PLDA IMPORT].CODE = [PLDA IMPORT HEADER].CODE) " & _
                            " INNER JOIN [PLDA IMPORT HEADER ZEGELS] ON ([PLDA IMPORT HEADER].HEADER = [PLDA IMPORT HEADER ZEGELS].HEADER) AND ([PLDA IMPORT HEADER].CODE = [PLDA IMPORT HEADER ZEGELS].CODE))" & _
                            " INNER JOIN [PLDA IMPORT HEADER HANDELAARS] ON ([PLDA IMPORT HEADER].HEADER = [PLDA IMPORT HEADER HANDELAARS].HEADER) AND ([PLDA IMPORT HEADER].CODE = [PLDA IMPORT HEADER HANDELAARS].CODE))" & _
                            " INNER JOIN [PLDA IMPORT DETAIL] ON ([PLDA IMPORT HEADER ZEGELS].HEADER = [PLDA IMPORT DETAIL].HEADER) AND ([PLDA IMPORT HEADER ZEGELS].CODE = [PLDA IMPORT DETAIL].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL BIJZONDERE] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL BIJZONDERE].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL BIJZONDERE].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL BIJZONDERE].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL CONTAINER] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL CONTAINER].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL CONTAINER].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL CONTAINER].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL DOCUMENTEN] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL DOCUMENTEN].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL DOCUMENTEN].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL DOCUMENTEN].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL ZELF] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL ZELF].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL ZELF].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL ZELF].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL HANDELAARS] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL HANDELAARS].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL HANDELAARS].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL HANDELAARS].CODE)) " & _
                            " INNER JOIN [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].CODE)) " & _
                            " WHERE [" & FieldToUse & "].[TREE ID] = " & Chr(39) & ProcessQuotes(strTreeID) & Chr(39) & " " & strSQLWhere & _
                            " ORDER BY [" & FieldToUse & "].[DATE LAST MODIFIED] DESC"
            'CSCLP-248
            GetSQLToUse = GetSQLToUse & ", [" & FieldToUse & "].[Code] ASC"
    
        Case 18
            GetSQLToUse = "SELECT  [PLDA COMBINED HEADER].CODE as [Code], [DOCUMENT NAME], [DTYPE], [TREE ID], [USERNAME], [DATE LAST MODIFIED] FROM (((((((([PLDA COMBINED]" & _
                            " INNER JOIN [PLDA COMBINED HEADER] ON [PLDA COMBINED].CODE = [PLDA COMBINED HEADER].CODE) " & _
                            " INNER JOIN [PLDA COMBINED HEADER ZEGELS] ON ([PLDA COMBINED HEADER].HEADER = [PLDA COMBINED HEADER ZEGELS].HEADER) AND ([PLDA COMBINED HEADER].CODE = [PLDA COMBINED HEADER ZEGELS].CODE))" & _
                            " INNER JOIN [PLDA COMBINED HEADER HANDELAARS] ON ([PLDA COMBINED HEADER ZEGELS].HEADER = [PLDA COMBINED HEADER HANDELAARS].HEADER) AND ([PLDA COMBINED HEADER HANDELAARS].CODE = [PLDA COMBINED HEADER ZEGELS].CODE))" & _
                            " INNER JOIN [PLDA COMBINED DETAIL] ON ([PLDA COMBINED HEADER HANDELAARS].HEADER = [PLDA COMBINED DETAIL].HEADER) AND ([PLDA COMBINED HEADER HANDELAARS].CODE = [PLDA COMBINED DETAIL].CODE)) " & _
                            " INNER JOIN [PLDA COMBINED DETAIL BIJZONDERE] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL BIJZONDERE].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL BIJZONDERE].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL BIJZONDERE].CODE)) " & _
                            " INNER JOIN [PLDA COMBINED DETAIL CONTAINER] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL CONTAINER].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL CONTAINER].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL CONTAINER].CODE)) " & _
                            " INNER JOIN [PLDA COMBINED DETAIL DOCUMENTEN] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL DOCUMENTEN].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL DOCUMENTEN].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL DOCUMENTEN].CODE)) " & _
                            " INNER JOIN [PLDA COMBINED DETAIL HANDELAARS] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL HANDELAARS].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL HANDELAARS].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL HANDELAARS].CODE)) " & _
                            " WHERE [" & FieldToUse & "].[TREE ID] = " & Chr(39) & ProcessQuotes(strTreeID) & Chr(39) & " " & strSQLWhere & _
                            " ORDER BY [" & FieldToUse & "].[DATE LAST MODIFIED] DESC"
            'CSCLP-248
            GetSQLToUse = GetSQLToUse & ", [" & FieldToUse & "].[Code] ASC"
    
    End Select
End Function

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width <= 6765 Then Me.Width = 6765
    If lvwItemsFound.Visible Then
        If Me.Height <= 6260 Then Me.Height = 6260
        With lvwItemsFound
          .Top = SSTab1.Top + SSTab1.Height + 108
          '.Height = me.h Me.StatusBar1.Top - (Me.SSTab1.Top + Me.SSTab1.Height + 108 + 110)  ' last 100 bottom margin
          .Width = Me.Width - 100
          .Height = Me.Height - (.Top + StatusBar1.Height + 850)
          
        End With
    End If
    
    With StatusBar1
      .Top = Me.Height - 320
      .Panels(1).Width = (Me.Width / 3) * 2
      .Panels(2).Width = (Me.Width / 3)
    End With
    cmdFind.Left = Me.Width - (200 + cmdFind.Width)     '1275
    cmdStop.Left = Me.Width - (200 + cmdFind.Width)
    cmdNewSearch.Left = Me.Width - (200 + cmdFind.Width)
    SSTab1.Width = SSTab1.Width + (cmdNewSearch.Left - (SSTab1.Left + SSTab1.Width) - 100)
    Animation1.Left = cmdNewSearch.Left + 135
End Sub

Private Function FormatValue(ByVal strValue As String, ByVal DataType As ADOX.DataTypeEnum, Optional ByVal eCondition As enuCondition) As String
    Select Case DataType
        Case ADOX.DataTypeEnum.adDate, _
            ADOX.DataTypeEnum.adDBDate        'DATE
            
            FormatValue = "CDate('" & IIf(IsDate(strValue), strValue, Now) & "')"
            
        ' adBSTR = 8                          ' String/TEXT
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
                
            If eCondition = eContains Then
                FormatValue = Chr(39) & "*" & ProcessQuotes(strValue) & "*" & Chr(39)
            Else
                FormatValue = Chr(39) & ProcessQuotes(strValue) & Chr(39)
            End If

        Case Else   'NUMERIC
            If IsNumeric(strValue) Then
                FormatValue = strValue
            Else
                FormatValue = Chr(39) & ProcessQuotes(strValue) & Chr(39)
            End If
    End Select
End Function

Private Sub SearchInEDI(bytDocType As Byte, lngNCTS_IEM_ID As Long, strYear As String, ParamArray TreeID() As Variant)

    
    Dim conArchive As ADODB.Connection
    
    Dim rstResult As ADODB.Recordset
    Dim rstBoxDefault As ADODB.Recordset
    Dim rstBoxSearch As ADODB.Recordset
    Dim RstCount As ADODB.Recordset
    Dim rstMAster As ADODB.Recordset
    
    Dim intTreeCtr As Integer
    Dim strSQl As String
    Dim strSQLWhere As String
    Dim strTableName As String
    Dim strFieldsToAppend As String
    Dim strTempSql As String
    Dim strBoxWhereClause As String
    Dim strBoxDefaultWhere As String
    Dim lngNCTS_IEM_TMS_ID As Long
    Dim blnWithInstances As Boolean
    Dim blnToadd As Boolean
    Dim lngInstance As Long
    Dim blnEDISeparate As Boolean
    Dim lngCounter As Long
    Dim strCommand As String
    
    Const EDISEPARATETABLE_DEP As String = "EO|EJ|S5|SB|Y1|Y5|Z4|W7|T7|AE|AF"
    Const EDISEPARATETABLE_ARR As String = "AI|T7"

    'if from History
    If strYear <> "" Then
        If Dir(cAppPath & "\mdb_EDIhistory" & Right(strYear, 2) & ".mdb") = "" Then
            Exit Sub
        End If
        '<<< dandan 112306
        '<<< Update with database password
        'Set datArchive = OpenDatabase(cAppPath & "\mdb_EDIhistory" & Right(strYear, 2) & ".mdb")
        
        ADOConnectDB conArchive, m_objDataSourceProperties, DBInstanceType_DATABASE_EDI_HISTORY, Right(strYear, 2)
        'DAOConnectDB datArchive, cAppPath, "mdb_EDIhistory" & Right(strYear, 2) & ".mdb"
    End If

    'if selected Type is "Any" and criteria value is not empty
    If icbType.SelectedItem.Key = "D" & enuDocType.eDocAny And txtValue.Text <> "" Then
        If txtValue.Text <> "" Then
            'filter according to NCTS_IEM_ID
            strBoxDefaultWhere = "NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID)
                
                strCommand = vbNullString
                strCommand = strCommand & "SELECT "
                strCommand = strCommand & "[BOX CODE] "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[BOX_SEARCH_MAP] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & strBoxDefaultWhere & " "
                strCommand = strCommand & "ORDER BY "
                strCommand = strCommand & "[BOX CODE]"
            ADORecordsetOpen strCommand, g_conEDIFACT, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datEDIFACT.OpenRecordset(strCommand)
            
            Do While Not rstBoxDefault.EOF
                DoEvents
                
                If mblnCancel = True Then GoTo Cancelled1
                
                    strCommand = vbNullString
                    strCommand = strCommand & "SELECT "
                    strCommand = strCommand & "[BOX CODE] "
                    strCommand = strCommand & "FROM "
                    strCommand = strCommand & "[BOX_SEARCH_MAP] "
                    strCommand = strCommand & "WHERE "
                    strCommand = strCommand & strBoxDefaultWhere & " "
                    strCommand = strCommand & "AND "
                    strCommand = strCommand & "[BOX CODE] = " & Chr(39) & ProcessQuotes(rstBoxDefault![Box Code]) & Chr(39) & " "
                    strCommand = strCommand & "ORDER BY "
                    strCommand = strCommand & "[BOX CODE]"
                ADORecordsetOpen strCommand, g_conEDIFACT, rstBoxSearch, adOpenKeyset, adLockOptimistic
                'Set rstBoxSearch = datEDIFACT.OpenRecordset("Select * from BOX_SEARCH_MAP " & strBoxDefaultWhere & " and [BOX CODE] = " & Chr(39) & ProcessQuotes(rstBoxDefault![Box Code]) & Chr(39) & " ORDER BY [BOX CODE]")
                
                strTableName = rstBoxSearch!BOX_COR_TABLE
                
                If TobeAppended(rstBoxSearch) = True Then
                    rstBoxSearch.MoveFirst
                    Do While Not rstBoxSearch.EOF
                        strTableName = rstBoxSearch!BOX_COR_TABLE
            
                        strFieldsToAppend = strFieldsToAppend & rstBoxSearch!BOX_COR_FIELD & ","
                        
                        lngNCTS_IEM_TMS_ID = rstBoxSearch!NCTS_IEM_TMS_ID
                        lngInstance = IIf(IsNull(rstBoxSearch!NCTS_DATA_INSTANCE), 0, rstBoxSearch!NCTS_DATA_INSTANCE)
                        blnWithInstances = IIf(lngInstance > 0, True, False)
                        strBoxWhereClause = ""
                        
                        rstBoxSearch.MoveNext
                    Loop
        
                    'remove the last comma appended
                    strFieldsToAppend = Left(strFieldsToAppend, Len(strFieldsToAppend) - 1)
                    
                    If cboCondition.ListIndex = 0 Then
                        strBoxWhereClause = "WHERE SearchedField LIKE " & Chr(39) & "*" & ProcessQuotes(txtValue.Text) & "*" & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        
                    ElseIf cboCondition.ListIndex = 1 Then
                    
                        strBoxWhereClause = "WHERE SearchedField = " & Chr(39) & ProcessQuotes(txtValue.Text) & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                    End If
                    
                        strCommand = vbNullString
                        strCommand = strCommand & "SELECT "
                        strCommand = strCommand & "[BOX CODE] "
                        strCommand = strCommand & "FROM "
                        strCommand = strCommand & "[BOX_SEARCH_MAP] "
                        strCommand = strCommand & "WHERE "
                        strCommand = strCommand & "BOX_COR_TABLE = " & Chr(39) & strTableName & Chr(39) & " "
                        strCommand = strCommand & "AND "
                        strCommand = strCommand & "BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " "
                        strCommand = strCommand & "AND "
                        strCommand = strCommand & "NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID & " "
                        strCommand = strCommand & "AND "
                        strCommand = strCommand & "NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " "
                        strCommand = strCommand & "ORDER BY "
                        strCommand = strCommand & "[BOX CODE]"
                    ADORecordsetOpen strCommand, g_conEDIFACT, RstCount, adOpenKeyset, adLockOptimistic
                    'Set RstCount = datEDIFACT.OpenRecordset("Select * from BOX_SEARCH_MAP where BOX_COR_TABLE = " & _
                                Chr(39) & strTableName & Chr(39) & " AND BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " AND NCTS_IEM_TMS_ID = " & _
                                lngNCTS_IEM_TMS_ID & " AND NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " ORDER BY [BOX CODE]")
                    
                    blnEDISeparate = (bytDocType = 11 And InStr(1, EDISEPARATETABLE_DEP, rstBoxDefault![Box Code])) Or _
                        (bytDocType = 12 And InStr(1, EDISEPARATETABLE_ARR, rstBoxDefault![Box Code]))
                    
                    strTempSql = SQLForEDI(strFieldsToAppend, strTableName, strBoxWhereClause, bytDocType, blnEDISeparate, True)
        
                    For intTreeCtr = 0 To UBound(TreeID)
                        StatusBar1.Panels(1).Text = "Searching " & FolderName(CStr(TreeID(intTreeCtr)))
                        
                        'strSql = GetSQLToUseEDI(FieldToUse, bytDocType, CStr(TreeID(intTreeCtr)), strSqlWhere)
                        strSQl = IIf(strTempSql <> "", strTempSql & " AND [TREE ID] = '" & TreeID(intTreeCtr) & "'", "where [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39))
                        
                        If strYear = "" Then
                            
                            ADORecordsetOpen strSQl, g_conEDIFACT, rstResult, adOpenKeyset, adLockOptimistic
                            'Set rstResult = datEDIFACT.OpenRecordset(strSQl)
                        Else    'means from Archive
                            ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                            'Set rstResult = datArchive.OpenRecordset(strSQl)
                        End If
                        
                        If bytDocType = 11 Then
                            If strYear = "" Then
                                
                                ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                            Else
                                ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                            End If
                        Else
                            If lngNCTS_IEM_ID = 2 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            ElseIf lngNCTS_IEM_ID = 11 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            End If
                        End If
                        
                        With rstOfflineTemp
                            Do While Not rstResult.EOF
                                DoEvents
                                
                                If mblnCancel = True Then GoTo Cancelled3
                                
                                blnToadd = IsRecordToBeAdded(rstResult, strTableName, strFieldsToAppend, _
                                        rstBoxDefault![Box Code], blnWithInstances, RstCount.RecordCount, lngInstance, blnEDISeparate)

                                If blnToadd Then
                                    If NotYetInRecordset(rstResult.Fields("Code").Value) Then
                                        'add one here
                                        .AddNew
                                    
                                        .Fields("CODE").Value = rstResult.Fields("CODE").Value
                                        .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                                        .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                                        .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                                        .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                                        .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                                        .Fields("ARCHIVE DATE").Value = strYear
                                        
                                        If Not (rstMAster.EOF And rstMAster.BOF) Then
                                            rstMAster.MoveFirst
                                            rstMAster.Find "Code = '" & rstResult.Fields("CODE").Value & "'", , adSearchForward
                                        End If

                                        If Not (rstMAster.BOF And rstMAster.EOF) Then
                                            If rstMAster.Fields("Code").Value = rstResult.Fields("CODE").Value Then
                                                For lngCounter = 7 To .Fields.Count - 1
                                                    If IsFieldExisting(rstMAster, .Fields(lngCounter).Name) Then
                                                        .Fields(lngCounter).Value = IIf(IsNull(rstMAster.Fields(.Fields(lngCounter).Name).Value), "", IIf(Len(rstMAster.Fields(.Fields(lngCounter).Name).Value) > 100, Left(rstMAster.Fields(.Fields(lngCounter).Name).Value, 97) & "...", rstMAster.Fields(.Fields(lngCounter).Name).Value))
                                                    Else
                                                        .Fields(lngCounter).Value = ""
                                                    End If
                                                Next lngCounter
                                            Else
                                                For lngCounter = 7 To .Fields.Count - 1
                                                    .Fields(lngCounter).Value = ""
                                                Next lngCounter
                                            End If
                                        Else
                                            For lngCounter = 7 To .Fields.Count - 1
                                                .Fields(lngCounter).Value = ""
                                            Next lngCounter
                                        End If
                                        
                                        .Update
                                   
                                    End If
                                End If
                                rstResult.MoveNext
                            Loop
                        End With
                        
                    Next
        
                'ELSE IF ng TobeAppended
                Else
                    
                    rstBoxSearch.MoveFirst
                    
                    'loop in BOXSEARCH
                    Do While Not rstBoxSearch.EOF
                        DoEvents
                        
                        If mblnCancel = True Then GoTo Cancelled2
                        
                        strFieldsToAppend = rstBoxSearch!BOX_COR_FIELD
                        strTableName = rstBoxSearch!BOX_COR_TABLE
                        
                        lngNCTS_IEM_TMS_ID = rstBoxSearch!NCTS_IEM_TMS_ID
                        lngInstance = IIf(IsNull(rstBoxSearch!NCTS_DATA_INSTANCE), 0, rstBoxSearch!NCTS_DATA_INSTANCE)
                        blnWithInstances = IIf(lngInstance > 0, True, False)
                        strBoxWhereClause = ""
                            
                            strCommand = vbNullString
                            strCommand = strCommand & "SELECT "
                            strCommand = strCommand & "* "
                            strCommand = strCommand & "FROM "
                            strCommand = strCommand & "BOX_SEARCH_MAP "
                            strCommand = strCommand & "WHERE "
                            strCommand = strCommand & "BOX_COR_TABLE = " & Chr(39) & ProcessQuotes(strTableName) & Chr(39) & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " "
                            strCommand = strCommand & "ORDER BY "
                            strCommand = strCommand & "[BOX CODE] "
                        ADORecordsetOpen strCommand, g_conEDIFACT, RstCount, adOpenKeyset, adLockOptimistic
                        'Set RstCount = datEDIFACT.OpenRecordset("Select * from BOX_SEARCH_MAP where BOX_COR_TABLE = " & _
                                    Chr(39) & ProcessQuotes(strTableName) & Chr(39) & " AND BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " AND NCTS_IEM_TMS_ID = " & _
                                    lngNCTS_IEM_TMS_ID & " AND NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " ORDER BY [BOX CODE]")
                        
                        blnEDISeparate = (bytDocType = 11 And InStr(1, EDISEPARATETABLE_DEP, rstBoxDefault![Box Code])) Or _
                                        (bytDocType = 12 And InStr(1, EDISEPARATETABLE_ARR, rstBoxDefault![Box Code]))

                        If cboCondition.ListIndex = 0 Then
                            strBoxWhereClause = " " & strFieldsToAppend & " LIKE " & Chr(39) & "*" & ProcessQuotes(txtValue.Text) & "*" & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        ElseIf cboCondition.ListIndex = 1 Then
                            'strBoxWhereClause = " " & strFieldsToAppend & " = " & chr(34) & ProcessQuotes(Replace(txtValue.Text, chr(34), chr(34) & chr(34))) & chr(34) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                            strBoxWhereClause = " " & strFieldsToAppend & " = " & Chr(39) & ProcessQuotes(txtValue.Text) & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        End If

                        strTempSql = SQLForEDI(strFieldsToAppend, strTableName, strBoxWhereClause, bytDocType, blnEDISeparate, True)
                    
                        For intTreeCtr = 0 To UBound(TreeID)
                            
                            StatusBar1.Panels(1).Text = "Searching in " & DocumentType(bytDocType) & " - " & CStr(rstBoxSearch![Box Code]) & " - " & FolderName(CStr(TreeID(intTreeCtr))) & strYear
                            strSQl = IIf(strTempSql <> "", strTempSql & " AND [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39), "where [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39))
                            
                            If strYear = "" Then
                                ADORecordsetOpen strSQl, g_conEDIFACT, rstResult, adOpenKeyset, adLockOptimistic
                                'Set rstResult = datEDIFACT.OpenRecordset(strSQl)
                            Else    'means from Archive
                                ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                                'Set rstResult = datArchive.OpenRecordset(strSQl)
                            End If
                            
                            If bytDocType = 11 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            Else
                                If lngNCTS_IEM_ID = 2 Then
                                    If strYear = "" Then
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    Else
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    End If
                                ElseIf lngNCTS_IEM_ID = 11 Then
                                    If strYear = "" Then
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    Else
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    End If
                                End If
                            End If

                            With rstOfflineTemp
                                Do While Not rstResult.EOF
                                    
                                    DoEvents
                                    
                                    If mblnCancel = True Then GoTo Cancelled4a
                                    
                                    blnToadd = IsRecordToBeAdded(rstResult, strTableName, strFieldsToAppend, _
                                                                rstBoxDefault![Box Code], blnWithInstances, RstCount.RecordCount, lngInstance, blnEDISeparate)
    
                                    If blnToadd Then
                                
                                         If NotYetInRecordset(rstResult.Fields("Code").Value) Then
                                             .AddNew
                                         
                                             .Fields("CODE").Value = rstResult.Fields("CODE").Value
                                             .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                                             .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                                             .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                                             .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                                             .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                                             .Fields("ARCHIVE DATE").Value = strYear
                                            
                                              If Not (rstMAster.BOF And rstMAster.EOF) Then
                                                rstMAster.MoveFirst
                                                rstMAster.Find "Code = '" & rstResult.Fields("CODE").Value & "'", , adSearchForward
                                              End If
                                            
                                              If Not (rstMAster.BOF And rstMAster.EOF) Then
                                                  If rstMAster.Fields("Code").Value = rstResult.Fields("CODE").Value Then
                                                      For lngCounter = 7 To .Fields.Count - 1
                                                          If IsFieldExisting(rstMAster, .Fields(lngCounter).Name) Then
                                                            .Fields(lngCounter).Value = IIf(IsNull(rstMAster.Fields(.Fields(lngCounter).Name).Value), "", IIf(Len(rstMAster.Fields(.Fields(lngCounter).Name).Value) > 100, Left(rstMAster.Fields(.Fields(lngCounter).Name).Value, 97) & "...", rstMAster.Fields(.Fields(lngCounter).Name).Value))
                                                          Else
                                                            .Fields(lngCounter).Value = ""
                                                          End If
                                                      Next lngCounter
                                                  Else
                                                      For lngCounter = 7 To .Fields.Count - 1
                                                          .Fields(lngCounter).Value = ""
                                                      Next lngCounter
                                                  End If
                                              Else
                                                  For lngCounter = 7 To .Fields.Count - 1
                                                      .Fields(lngCounter).Value = ""
                                                  Next lngCounter
                                              End If
                                             
                                             .Update
                                        
                                         End If
                                    End If
                                    rstResult.MoveNext
                                Loop
                            End With
                            
                        Next
                    
                        rstBoxSearch.MoveNext
                    Loop
                End If
                rstBoxSearch.Close
                
                Set rstBoxSearch = Nothing
                
                rstBoxDefault.MoveNext
            Loop

        End If
    
    'if there is a particular document selected
    Else
    
        Dim strBoxCode As String
        Dim blnItemSelected As Boolean
        Dim blnDoSimpleSearch As Boolean
        
        If Not icbBox.SelectedItem Is Nothing Then
            'strBoxCode = Left(icbBox.SelectedItem.Key, Len(icbBox.SelectedItem.Key) - 1)
            strBoxCode = Mid(icbBox.SelectedItem.Key, 2, Len(icbBox.SelectedItem.Key) - 2)
            blnItemSelected = True
        End If
        
        If blnItemSelected Then
            If Right(icbBox.SelectedItem.Key, 1) = "0" Then
                blnDoSimpleSearch = True
            End If
        End If
                
        If blnItemSelected = False Or blnDoSimpleSearch Or strBoxCode = "MRN" Then
            For intTreeCtr = 0 To UBound(TreeID)
                If strYear = "" Then
                    LookInMainTable g_conEDIFACT, strBoxCode, "DATA_NCTS", bytDocType, CStr(TreeID(intTreeCtr)), lngNCTS_IEM_ID
                Else
                    If blnEDIHistoryExisting = True Then
                        For lngCounter = 0 To UBound(g_conEDIHistory)
                            If mblnCancel = True Then Exit For
                            If Not g_conEDIHistory(lngCounter) Is Nothing Then
                                
                                Dim strDBName As String
                                Dim lngPosDBName As Long
                                strDBName = g_conEDIHistory(lngCounter).ConnectionString
                                lngPosDBName = InStr(1, strDBName, ".mdb")
                                If lngPosDBName > 0 Then
                                    strDBName = Mid(strDBName, 1, lngPosDBName - 1)
                                    If Right(strDBName, 2) = strYear Then
                                        LookInMainTable g_conEDIHistory(lngCounter), strBoxCode, "DATA_NCTS", bytDocType, CStr(TreeID(0)), lngNCTS_IEM_ID
                                        Exit For
                                    End If
                                Else
                                    Debug.Assert False
                                End If
                            End If
                        Next lngCounter
                    End If
                End If
            Next
        Else
            ADORecordsetOpen "Select * from BOX_SEARCH_MAP where NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " AND [BOX CODE] = " & Chr(39) & ProcessQuotes(strBoxCode) & Chr(39) & " ORDER BY [BOX CODE]", _
                                g_conEDIFACT, rstBoxSearch, adOpenKeyset, adLockOptimistic
            'Set rstBoxSearch = datEDIFACT.OpenRecordset("Select * from BOX_SEARCH_MAP where NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " AND [BOX CODE] = " & Chr(39) & ProcessQuotes(strBoxCode) & Chr(39) & " ORDER BY [BOX CODE]")
            
            If Not (rstBoxSearch.EOF And rstBoxSearch.BOF) Then
                If TobeAppended(rstBoxSearch) = True Then
                    rstBoxSearch.MoveFirst
                    Do While Not rstBoxSearch.EOF
                        strTableName = rstBoxSearch!BOX_COR_TABLE
            
                        strFieldsToAppend = strFieldsToAppend & rstBoxSearch!BOX_COR_FIELD & ","
                        rstBoxSearch.MoveNext
                    Loop
        
                    'remove the last comma appended
                    strFieldsToAppend = Left(strFieldsToAppend, Len(strFieldsToAppend) - 1)
                    
                    Select Case cboCondition.ListIndex
                        Case enuCondition.eContains
                            strBoxWhereClause = " " & strFieldsToAppend & " Like " & Chr(39) & "*" & ProcessQuotes(txtValue.Text) & "*" & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        Case enuCondition.eIsExactly
                            strBoxWhereClause = " " & strFieldsToAppend & " = " & Chr(39) & ProcessQuotes(txtValue.Text) & Chr(39) & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        Case enuCondition.eDoesnotContain
                            strBoxWhereClause = "InStr(1, UCase(" & strFieldsToAppend & "), " & Chr(39) & ProcessQuotes(UCase(txtValue.Text)) & Chr(39) & ") = 0" & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        Case enuCondition.eIsEmpty
                            strBoxWhereClause = "len(trim(" & strFieldsToAppend & ")) <= 0 OR ISNULL(" & strFieldsToAppend & ") " & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                        Case enuCondition.eIsNotEmpty
                            strBoxWhereClause = "len(trim(" & strFieldsToAppend & ")) > 0 " & IIf(blnEDISeparate = True, "", " AND NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID)
                    End Select

                    blnEDISeparate = (bytDocType = 11 And InStr(1, EDISEPARATETABLE_DEP, rstBoxDefault![Box Code])) Or _
                                        (bytDocType = 12 And InStr(1, EDISEPARATETABLE_ARR, rstBoxDefault![Box Code]))

                    strTempSql = SQLForEDI(strFieldsToAppend, strTableName, strBoxWhereClause, bytDocType, blnEDISeparate, True)
                    
                    For intTreeCtr = 0 To UBound(TreeID)
                    
                        StatusBar1.Panels(1).Text = "Searching " & FolderName(CStr(TreeID(intTreeCtr)))
                        
                        strSQl = IIf(strTempSql <> "", strTempSql & " AND [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39), "where [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39))
                        
                        If strYear = "" Then
                            ADORecordsetOpen strSQl, g_conEDIFACT, rstResult, adOpenKeyset, adLockOptimistic
                            'Set rstResult = datEDIFACT.OpenRecordset(strSQl)
                        Else    'means from Archive
                            ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                            'Set rstResult = datArchive.OpenRecordset(strSQl)
                        End If
                        
                        If bytDocType = 11 Then
                            If strYear = "" Then
                                ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                            Else
                                ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                            End If
                        Else
                            If lngNCTS_IEM_ID = 2 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            ElseIf lngNCTS_IEM_ID = 11 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            End If
                        End If

                        With rstOfflineTemp
                            Do While Not rstResult.EOF
                                DoEvents
                                
                                If mblnCancel = True Then GoTo Cancelled5
                                If NotYetInRecordset(rstResult.Fields("Code").Value) Then
                                    .AddNew
                                
                                    .Fields("CODE").Value = rstResult.Fields("CODE").Value
                                    .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                                    .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                                    .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                                    .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                                    .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                                    .Fields("ARCHIVE DATE").Value = strYear
                                    
                                    If Not (rstMAster.BOF And rstMAster.EOF) Then
                                        rstMAster.MoveFirst
                                        rstMAster.Find "Code = '" & rstResult.Fields("CODE").Value & "'", , adSearchForward
                                    End If
                                    
                                    If Not (rstMAster.BOF And rstMAster.EOF) Then
                                        If rstMAster.Fields("Code").Value = rstResult.Fields("CODE").Value Then
                                            For lngCounter = 7 To .Fields.Count - 1
                                                If IsFieldExisting(rstMAster, .Fields(lngCounter).Name) Then
                                                    .Fields(lngCounter).Value = IIf(IsNull(rstMAster.Fields(.Fields(lngCounter).Name).Value), "", IIf(Len(rstMAster.Fields(.Fields(lngCounter).Name).Value) > 100, Left(rstMAster.Fields(.Fields(lngCounter).Name).Value, 97) & "...", rstMAster.Fields(.Fields(lngCounter).Name).Value))
                                                Else
                                                    .Fields(lngCounter).Value = ""
                                                End If
                                            Next lngCounter
                                        Else
                                            For lngCounter = 7 To .Fields.Count - 1
                                                .Fields(lngCounter).Value = ""
                                            Next lngCounter
                                        End If
                                    Else
                                        For lngCounter = 7 To .Fields.Count - 1
                                            .Fields(lngCounter).Value = ""
                                        Next lngCounter
                                    End If

                                    .Update
                               
                                End If
        
                                
                                rstResult.MoveNext
                            Loop
                        End With
                        
                    Next
        
                Else
                    rstBoxSearch.MoveFirst
                    
                    Do While Not rstBoxSearch.EOF
                    
                        strFieldsToAppend = rstBoxSearch!BOX_COR_FIELD
                        strTableName = rstBoxSearch!BOX_COR_TABLE
                        lngNCTS_IEM_TMS_ID = rstBoxSearch!NCTS_IEM_TMS_ID
                        lngInstance = IIf(IsNull(rstBoxSearch!NCTS_DATA_INSTANCE), 0, rstBoxSearch!NCTS_DATA_INSTANCE)
                        blnWithInstances = IIf(lngInstance > 0, True, False)
                        strBoxWhereClause = ""
                            
                            strCommand = vbNullString
                            strCommand = strCommand & "SELECT "
                            strCommand = strCommand & "* "
                            strCommand = strCommand & "FROM "
                            strCommand = strCommand & "BOX_SEARCH_MAP "
                            strCommand = strCommand & "WHERE "
                            strCommand = strCommand & "BOX_COR_TABLE = " & Chr(39) & ProcessQuotes(strTableName) & Chr(39) & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID & " "
                            strCommand = strCommand & "AND "
                            strCommand = strCommand & "NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " "
                            strCommand = strCommand & "ORDER BY "
                            strCommand = strCommand & "[BOX CODE] "
                        ADORecordsetOpen strCommand, g_conEDIFACT, RstCount, adOpenKeyset, adLockOptimistic
                        'Set RstCount = datEDIFACT.OpenRecordset("Select * from BOX_SEARCH_MAP where BOX_COR_TABLE = " & _
                                    Chr(39) & ProcessQuotes(strTableName) & Chr(39) & " AND BOX_COR_FIELD = " & Chr(39) & ProcessQuotes(strFieldsToAppend) & Chr(39) & " AND NCTS_IEM_TMS_ID = " & _
                                    lngNCTS_IEM_TMS_ID & " AND NCTS_IEM_ID = " & CStr(lngNCTS_IEM_ID) & " ORDER BY [BOX CODE]")
                        
                        blnEDISeparate = (bytDocType = 11 And InStr(1, EDISEPARATETABLE_DEP, rstBoxSearch![Box Code])) Or _
                                        (bytDocType = 12 And InStr(1, EDISEPARATETABLE_ARR, rstBoxSearch![Box Code]))
                        
                        If blnEDISeparate = False Then
                            strBoxWhereClause = "NCTS_IEM_TMS_ID = " & lngNCTS_IEM_TMS_ID
                        End If
                        
                        strTempSql = SQLForEDI(strFieldsToAppend, strTableName, strBoxWhereClause, bytDocType, blnEDISeparate, True)
                    
                    
                        For intTreeCtr = 0 To UBound(TreeID)
                            
                            StatusBar1.Panels(1).Text = "Searching " & FolderName(CStr(TreeID(intTreeCtr)))
                            strSQl = IIf(strTempSql <> "", strTempSql & " AND [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39), "where [TREE ID] = " & Chr(39) & ProcessQuotes(TreeID(intTreeCtr)) & Chr(39))
                            
                            If strYear = "" Then
                                ADORecordsetOpen strSQl, g_conEDIFACT, rstResult, adOpenKeyset, adLockOptimistic
                                'Set rstResult = datEDIFACT.OpenRecordset(strSQl)
                            Else    'means from Archive
                                ADORecordsetOpen strSQl, conArchive, rstResult, adOpenKeyset, adLockOptimistic
                                'Set rstResult = datArchive.OpenRecordset(strSQl)
                            End If
                            
                            If bytDocType = 11 Then
                                If strYear = "" Then
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                Else
                                    ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                    'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                End If
                            Else
                                If lngNCTS_IEM_ID = 2 Then
                                    If strYear = "" Then
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    Else
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    End If
                                ElseIf lngNCTS_IEM_ID = 11 Then
                                    If strYear = "" Then
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    Else
                                        ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'", conArchive, rstMAster, adOpenKeyset, adLockOptimistic
                                        'Set rstMAster = datArchive.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID(intTreeCtr) & "'")
                                    End If
                                End If
                            End If

                            With rstOfflineTemp
                                Do While Not rstResult.EOF

                                    DoEvents
                                    
                                    If mblnCancel = True Then GoTo Cancelled4b
                                
                                    blnToadd = IsRecordToBeAdded(rstResult, strTableName, strFieldsToAppend, _
                                                                strBoxCode, blnWithInstances, RstCount.RecordCount, lngInstance, blnEDISeparate)
                                    
                                    If blnToadd Then
                                         If NotYetInRecordset(rstResult.Fields("Code").Value) Then
                                             .AddNew
                                         
                                             .Fields("CODE").Value = rstResult.Fields("CODE").Value
                                             .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                                             .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                                             .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                                             .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                                             .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                                             .Fields("ARCHIVE DATE").Value = strYear
                                                
                                                If Not (rstMAster.BOF And rstMAster.EOF) Then
                                                    rstMAster.MoveFirst
                                                    rstMAster.Find "Code = '" & rstResult.Fields("CODE").Value & "'", , adSearchForward
                                                End If
                                                
                                                If Not (rstMAster.BOF And rstMAster.EOF) Then
                                                    If rstMAster.Fields("Code").Value = rstResult.Fields("CODE").Value Then
                                                        For lngCounter = 7 To .Fields.Count - 1
                                                            If IsFieldExisting(rstMAster, .Fields(lngCounter).Name) Then
                                                                .Fields(lngCounter).Value = IIf(IsNull(rstMAster.Fields(.Fields(lngCounter).Name).Value), "", IIf(Len(rstMAster.Fields(.Fields(lngCounter).Name).Value) > 100, Left(rstMAster.Fields(.Fields(lngCounter).Name).Value, 97) & "...", rstMAster.Fields(.Fields(lngCounter).Name).Value))
                                                            Else
                                                                .Fields(lngCounter).Value = ""
                                                            End If
                                                        Next lngCounter
                                                    Else
                                                        For lngCounter = 7 To .Fields.Count - 1
                                                            .Fields(lngCounter).Value = ""
                                                        Next lngCounter
                                                    End If
                                                Else
                                                    For lngCounter = 7 To .Fields.Count - 1
                                                        .Fields(lngCounter).Value = ""
                                                    Next lngCounter
                                                End If
    
                                             .Update
                                        
                                         End If
                                    End If
                                    rstResult.MoveNext
                                Loop
                            End With
                            
                        Next
                    
                        rstBoxSearch.MoveNext
                    Loop
                    
                    RstCount.Close
                    Set RstCount = Nothing
                End If
            End If
        End If
        
    End If
    
    ADORecordsetClose rstResult
    ADORecordsetClose rstMAster
    
    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If

    Exit Sub
    
Cancelled1:
    
    ADORecordsetClose rstBoxDefault
    
    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    Exit Sub
    
Cancelled2:
    ADORecordsetClose rstBoxSearch
    ADORecordsetClose rstBoxDefault
    
    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    Exit Sub
    
Cancelled3:
    ADORecordsetClose rstBoxSearch
    ADORecordsetClose rstBoxDefault
    ADORecordsetClose RstCount
    ADORecordsetClose rstResult
    ADORecordsetClose rstMAster
    
    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    Exit Sub
    
Cancelled4a:
    ADORecordsetClose rstBoxSearch
    ADORecordsetClose rstBoxDefault
    ADORecordsetClose RstCount
    ADORecordsetClose rstResult
    ADORecordsetClose rstMAster
    
    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    Exit Sub
    
Cancelled4b:
    ADORecordsetClose rstBoxSearch
    ADORecordsetClose RstCount
    ADORecordsetClose rstResult
    ADORecordsetClose rstMAster

    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    Exit Sub
    
Cancelled5:
    ADORecordsetClose rstBoxSearch
    ADORecordsetClose rstResult
    ADORecordsetClose rstMAster

    If strYear <> "" Then
        ADODisconnectDB conArchive
    End If
    
    '1.)
End Sub

Private Function SqlWhereEDI(TableName As String, bytDocType As Byte) As String
    Dim strTemp As String
    Dim strFieldToUse As String
    Dim strValue As String
    Dim strSQLWhere As String
    Dim strTableToOpen As String
    Dim strFieldTemp As String
    Dim strWildCardCharCheck As String

    strSQLWhere = "1 = 1"
    
    '======== DocType ===================================
    strTemp = ""
    strTemp = "[DTYPE] = " & CStr(bytDocType)
    strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
    '====================================================
    
    '============ for DATE CONDITION ====================
    If optFind.Value = True Then
        strTemp = ""
        strTemp = DateCondition(TableName)
        strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
    End If
    '====================================================

    '============ for DOCUMENT NAME CONDITION ===========
    strTemp = ""
    strWildCardCharCheck = cboName.Text
    If strWildCardCharCheck <> "" Then
        If InStr(1, strWildCardCharCheck, "*") > 0 Or _
            InStr(1, strWildCardCharCheck, "?") > 0 Then
                
            strTemp = "[DOCUMENT NAME] LIKE " & Chr(39) & ProcessQuotes(strWildCardCharCheck) & Chr(39) & " "
        Else
            strTemp = "[DOCUMENT NAME] = " & Chr(39) & ProcessQuotes(strWildCardCharCheck) & Chr(39) & " "
        End If
    End If

    strSQLWhere = IIf(strTemp <> "", strSQLWhere & " AND " & strTemp, strSQLWhere)
    '====================================================
    
    'SqlWhereEDI = IIf(strSqlWhere = "", "Where 1=1", strSqlWhere)
    SqlWhereEDI = strSQLWhere
End Function


Private Function TobeAppended(rstBoxSearch As ADODB.Recordset) As Boolean

    Dim lngNCTS_IEM_TMS_ID As Long

    If rstBoxSearch.RecordCount > 0 Then
        TobeAppended = False
    Else
        TobeAppended = True
        
        rstBoxSearch.MoveFirst
        
        lngNCTS_IEM_TMS_ID = rstBoxSearch![NCTS_IEM_TMS_ID]
        Do While Not rstBoxSearch.EOF
            
            If rstBoxSearch![NCTS_IEM_TMS_ID] <> lngNCTS_IEM_TMS_ID Then
                TobeAppended = False
                Exit Function
                
            End If
            lngNCTS_IEM_TMS_ID = rstBoxSearch![NCTS_IEM_TMS_ID]
            rstBoxSearch.MoveNext
        Loop
    End If
End Function

Private Function SQLForEDI(SelectClause As String, strTableName As String, strBoxWhereClause As String, bytDocType As Byte, blnEDISeparate As Boolean, Optional blnappended As Boolean) As String
    Dim strSQl As String
    Dim strWhere As String
    If blnEDISeparate Then
        
        strSQl = "Select *, DATA_NCTS.CODE as CODE " & IIf(SelectClause <> "", ", " & SelectClause & " AS SearchedField ", "") & _
                 "FROM DATA_NCTS INNER JOIN " & strTableName & " ON DATA_NCTS.CODE = " & strTableName & ".CODE"
    Else
        strSQl = "Select *, DATA_NCTS.CODE as CODE " & IIf(SelectClause <> "", ", " & SelectClause & " AS SearchedField ", "") & _
                "From DATA_NCTS Inner Join (DATA_NCTS_MESSAGES INNER JOIN " & strTableName & " On DATA_NCTS_MESSAGES.DATA_NCTS_MSG_ID = " & _
                strTableName & ".DATA_NCTS_MSG_ID) ON DATA_NCTS.DATA_NCTS_ID = DATA_NCTS_MESSAGES.DATA_NCTS_ID"
    End If
    
    strWhere = SqlWhereEDI(strTableName, bytDocType)
    If Trim(strBoxWhereClause) <> "" And strBoxWhereClause <> "NCTS_IEM_TMS_ID = 0" Then
        
        strSQl = strSQl & " where " & strBoxWhereClause & IIf(strBoxWhereClause <> "" And strWhere <> "", " AND " & strWhere, strWhere)
    Else
        strSQl = strSQl & " where 1=1 and " & strWhere
    End If
    
    SQLForEDI = strSQl
End Function

Private Function NotYetInRecordset(strCode As String) As Boolean

    If rstOfflineTemp.RecordCount = 0 Then
        NotYetInRecordset = True
    Else
        rstOfflineTemp.MoveFirst
        rstOfflineTemp.Find "Code = #" & strCode & "#"
        
        If rstOfflineTemp.EOF Then
            NotYetInRecordset = True
        Else
            NotYetInRecordset = False
        End If
    End If
    
End Function

Private Sub LookInMainTable(ByRef ADOConnection As ADODB.Connection, _
                            ByVal FieldToUse As String, _
                            ByVal TableToUse As String, _
                            ByVal bytDocType As Byte, _
                            ByVal TreeID As String, _
                            ByVal lngNCTS_IEM_ID As Long)
    Dim strSQLWhere As String
    Dim rstResult As ADODB.Recordset
    Dim strSQl As String
    Dim enuADOXDataType As ADOX.DataTypeEnum
    Dim strValue As String
    Dim strTemp As String
    Dim intCtr As Integer
    Dim lngCounter As Long
    Dim rstMAster As ADODB.Recordset
    
    strSQLWhere = SqlWhereEDI(TableToUse, bytDocType)
        
    If FieldToUse <> "" Then
        enuADOXDataType = GetDataType(ADOConnection, TableToUse, FieldToUse)
        'intDataType = GetDataType(datToUse, TableToUse, FieldToUse)
        
        FieldToUse = "[" & FieldToUse & "]"
        
        Select Case cboCondition.ListIndex
            Case enuCondition.eContains
                If enuADOXDataType = adDate Or _
                    enuADOXDataType = adDBDate Then
                    'strTemp = IIf(Trim(strSqlWhere) <> "", " and " & FieldToUse & " = " & strValue, FieldToUse & " LIKE " & strValue)
                    'strTemp = IIf(Trim(strSqlWhere) <> "", " and " & FieldToUse & " LIKE " & strValue, FieldToUse & " LIKE " & strValue)
                    
                    strTemp = IIf(Trim(strSQLWhere) <> "", " and " & FieldToUse & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & FieldToUse & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')", FieldToUse & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & FieldToUse & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')")
                Else
                    strValue = FormatValue(txtValue.Text, enuADOXDataType, eContains)
                    strTemp = IIf(Trim(strSQLWhere) <> "", " and " & FieldToUse & " LIKE " & strValue, FieldToUse & " LIKE " & strValue)
                End If
                
            Case enuCondition.eIsExactly
                If enuADOXDataType = adDate Or _
                    enuADOXDataType = adDBDate Then
                    strTemp = IIf(Trim(strSQLWhere) <> "", " and " & FieldToUse & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & FieldToUse & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')", FieldToUse & " >= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 12:00:00 AM", Now) & " ') and " & FieldToUse & " <= CDate('" & IIf(IsDate(txtValue.Text), txtValue.Text & " 11:59:59 PM", Now) & " ')")
                Else
                    strValue = FormatValue(txtValue.Text, enuADOXDataType, eIsExactly)
                    strTemp = IIf(Trim(strSQLWhere) <> "", " and " & FieldToUse & " = " & strValue, FieldToUse & " = " & strValue)
                End If
            Case enuCondition.eDoesnotContain
                strValue = FormatValue(txtValue.Text, enuADOXDataType, eDoesnotContain)
                strTemp = IIf(Trim(strSQLWhere) <> "", " and " & "InStr(1, UCase(" & FieldToUse & "), " & UCase(strValue) & ") = 0", "InStr(1, UCase(" & FieldToUse & "), " & UCase(strValue) & ") = 0")
                
            Case enuCondition.eIsEmpty
                strValue = FormatValue(txtValue.Text, enuADOXDataType, eIsEmpty)
                strTemp = IIf(Trim(strSQLWhere) <> "", " and " & "len(trim(" & FieldToUse & ")) <= 0 ", "len(trim(" & FieldToUse & ")) <= 0 ")
                
            Case enuCondition.eIsNotEmpty
                strValue = FormatValue(txtValue.Text, enuADOXDataType, eIsNotEmpty)
                strTemp = IIf(Trim(strSQLWhere) <> "", " and " & "len(trim(" & FieldToUse & ")) > 0 ", "len(trim(" & FieldToUse & ")) > 0 ")
            
        End Select
    End If
    
    strSQLWhere = strSQLWhere & strTemp
    
    strSQl = "Select * from " & TableToUse & IIf(Trim(strSQLWhere) <> "", " where " & strSQLWhere, "")
    
    On Error GoTo EarlyExit
    ADORecordsetOpen strSQl & " and [Tree ID] = '" & TreeID & "'", ADOConnection, rstResult, adOpenKeyset, adLockOptimistic
    'Set rstResult = datToUse.OpenRecordset(strSQl & " and [Tree ID] = '" & TreeID & "'")
    
    Dim strDBName As String
    Dim lngPosDBName As Long
    strDBName = ADOConnection.ConnectionString
    lngPosDBName = InStr(1, strDBName, ".mdb")
    If lngPosDBName > 0 Then
        strDBName = Mid(strDBName, 1, lngPosDBName - 1)
    Else
        Debug.Assert False
    End If
    
    If bytDocType = 11 Then
        If IsNumeric(Right(strDBName, 2)) Then
            ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID & "'", ADOConnection, rstMAster, adOpenKeyset, adLockOptimistic
            'Set rstMAster = datToUse.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID & "'")
        Else
            ADORecordsetOpen "Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
            'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS WHERE [Tree ID] = '" & TreeID & "'")
        End If
    Else
        If lngNCTS_IEM_ID = 2 Then
            If IsNumeric(Right(strDBName, 2)) Then
                ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID & "'", ADOConnection, rstMAster, adOpenKeyset, adLockOptimistic
                'Set rstMAster = datToUse.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID & "'")
            Else
                ADORecordsetOpen "Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTS2 WHERE [Tree ID] = '" & TreeID & "'")
            End If
        ElseIf lngNCTS_IEM_ID = 11 Then
            If IsNumeric(Right(strDBName, 2)) Then
                ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID & "'", ADOConnection, rstMAster, adOpenKeyset, adLockOptimistic
                'Set rstMAster = datToUse.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID & "'")
            Else
                ADORecordsetOpen "Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID & "'", g_conData, rstMAster, adOpenKeyset, adLockOptimistic
                'Set rstMAster = datData.OpenRecordset("Select * FROM MASTEREDINCTSIE44 WHERE [Tree ID] = '" & TreeID & "'")
            End If
        End If
    End If
    
    With rstOfflineTemp
        Do While Not rstResult.EOF
            DoEvents
            
            If mblnCancel Then GoTo Cancelled
            
            If NotYetInRecordset(rstResult.Fields("Code").Value) Then

                .AddNew
            
                .Fields("CODE").Value = rstResult.Fields("CODE").Value
                .Fields("NAME").Value = rstResult.Fields("DOCUMENT NAME").Value
                .Fields("DOCUMENT").Value = rstResult.Fields("DTYPE").Value
                .Fields("IN FOLDER").Value = rstResult.Fields("TREE ID").Value
                .Fields("USERNAME").Value = rstResult.Fields("USERNAME").Value
                .Fields("DATE MODIFIED").Value = rstResult.Fields("DATE LAST MODIFIED").Value
                .Fields("ARCHIVE DATE").Value = IIf(IsNumeric(Right(strDBName, 2)), Right(strDBName, 2), "")
                
                'Use masteredincts* to get data
                If Not (rstMAster.BOF And rstMAster.EOF) Then
                    rstMAster.MoveFirst
                    rstMAster.Find "Code = '" & rstResult.Fields("CODE").Value & "'", , adSearchForward
                End If
                
                If Not (rstMAster.BOF And rstMAster.EOF) Then
                    If rstMAster.Fields("Code").Value = rstResult.Fields("CODE").Value Then
                        For lngCounter = 7 To .Fields.Count - 1
                            If IsFieldExisting(rstMAster, .Fields(lngCounter).Name) Then
                                .Fields(lngCounter).Value = IIf(IsNull(rstMAster.Fields(.Fields(lngCounter).Name).Value), "", IIf(Len(rstMAster.Fields(.Fields(lngCounter).Name).Value) > 100, Left(rstMAster.Fields(.Fields(lngCounter).Name).Value, 97) & "...", rstMAster.Fields(.Fields(lngCounter).Name).Value))
                            Else
                                .Fields(lngCounter).Value = ""
                            End If
                        Next lngCounter
                    Else
                        For lngCounter = 7 To .Fields.Count - 1
                            .Fields(lngCounter).Value = ""
                        Next lngCounter
                    End If
                Else
                    For lngCounter = 7 To .Fields.Count - 1
                        .Fields(lngCounter).Value = ""
                    Next lngCounter
                End If
                
                .Update
            End If
            
            rstResult.MoveNext
        Loop
    End With
    
    ADORecordsetClose rstMAster
    ADORecordsetClose rstResult
    
    Exit Sub
    
Cancelled:
    ADORecordsetClose rstMAster
    ADORecordsetClose rstResult
    
EarlyExit:

End Sub

Private Sub Index_Column(ByVal strColumnName As String)
    Dim intColumnHeaderIndex As Integer
    
    With lvwItemsFound
        For intColumnHeaderIndex = 1 To .ColumnHeaders.Count
            If UCase(Trim(.ColumnHeaders(intColumnHeaderIndex).Text)) = UCase(strColumnName) Then
                If intColumnHeaderIndex = 5 Then
                    .SortKey = intColumnHeaderIndex
                Else
                    .SortKey = intColumnHeaderIndex - 1
                End If
                
                If .SortOrder = lvwDescending Then
                    .SortOrder = lvwAscending
                Else
                    .SortOrder = lvwDescending
                End If
                
                .Sorted = True
                
                .Refresh
                
                Exit For
            End If
        Next
        
        .Sorted = False
        
        intActiveSortKey = .SortKey
        intActiveSortOrder = .SortOrder
        
        If Not .SelectedItem Is Nothing Then
            .SelectedItem.EnsureVisible
        End If
    End With
End Sub

Private Sub ListCheck(ByVal cStrID As String)
    strView = cStrID
    Select Case cStrID
        Case "ID_LargeIcons"
            lvwItemsFound.View = lvwIcon
            
            lvwItemsFound.Sorted = True
            lvwItemsFound.SortKey = intActiveSortKey
            lvwItemsFound.SortOrder = intActiveSortOrder
        Case "ID_SmallIcons"
            lvwItemsFound.View = lvwSmallIcon
            
            lvwItemsFound.Sorted = True
            lvwItemsFound.SortKey = intActiveSortKey
            lvwItemsFound.SortOrder = intActiveSortOrder
        Case "ID_List"
            lvwItemsFound.View = lvwList
        Case "ID_Details"
            lvwItemsFound.View = lvwReport
    End Select
    
    lvwItemsFound.Refresh
    
    SSActiveToolBars1.Refresh
End Sub

Private Sub lvwItemsFound_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim intColumnHeaderIndex As Integer
    
    intColumnHeaderIndex = ColumnHeader.Index
    
    If intColumnHeaderIndex = 5 Then
        lvwItemsFound.SortKey = intColumnHeaderIndex
    Else
        lvwItemsFound.SortKey = intColumnHeaderIndex - 1
    End If
    
    If lvwItemsFound.ColumnHeaders(intColumnHeaderIndex).Tag = "A" Then
        lvwItemsFound.SortOrder = lvwDescending
        lvwItemsFound.ColumnHeaders(intColumnHeaderIndex).Tag = "D"
    Else
        lvwItemsFound.SortOrder = lvwAscending
        lvwItemsFound.ColumnHeaders(intColumnHeaderIndex).Tag = "A"
    End If
    
    lvwItemsFound.Sorted = True
    
    lvwItemsFound.Sorted = False
    
    If Not lvwItemsFound.SelectedItem Is Nothing Then
        lvwItemsFound.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub UpdateConditionList(intListIndex As Integer)
    cboCondition.Clear
    
    If intListIndex = 0 Then
        cboCondition.AddItem (Trim(Translate(212)))
        cboCondition.AddItem (Trim(Translate(213)))
    Else
        cboCondition.AddItem (Trim(Translate(212)))
        cboCondition.AddItem (Trim(Translate(213)))
        cboCondition.AddItem (Trim(Translate(214)))
        cboCondition.AddItem (Trim(Translate(215)))
        cboCondition.AddItem (Trim(Translate(216)))
    End If
End Sub

Private Sub DeleteItems()
Dim i As Integer
Dim bytDocType As Byte
Dim strUniqueCode As String
Dim intAnswer As Integer
'change tree id to "DD"
    

    For i = 1 To lvwItemsFound.ListItems.Count
    
        If lvwItemsFound.ListItems(i).Selected = True Then
        
            
            bytDocType = CByte(lvwItemsFound.ListItems(i).ListSubItems(1).Tag)
            strUniqueCode = lvwItemsFound.ListItems(i).Tag
        
            If TheFileIsBeingSent(strUniqueCode, bytDocType) Or TheFileIsOpen(strUniqueCode, bytDocType) Then
                'MsgBox "Cannot delete item"
                MsgBox Translate(2103)
                GoTo NextItem
            End If
            If IsInDeletedFolder(lvwItemsFound.ListItems(i).ListSubItems(2)) Then
                'intAnswer = MsgBox("Permanently delete item?", vbYesNoCancel)
                intAnswer = MsgBox(Translate(2104), vbYesNoCancel)
                If intAnswer = vbYes Then
                    DeleteDocument bytDocType, strUniqueCode
                ElseIf intAnswer = vbNo Then
                    GoTo NextItem
                Else
                    GoTo WasCancelled
                End If
                

            Else
                intAnswer = MsgBox(Translate(2105), vbYesNoCancel)
                If intAnswer = vbYes Then
                    MoveToDeletedFolder bytDocType, strUniqueCode
                ElseIf intAnswer = vbNo Then
                    GoTo NextItem
                Else
                    GoTo WasCancelled
                End If
                
            End If
        End If
NextItem:
    Next

WasCancelled:

End Sub

Private Function TemplateName(strTemplateTreeID As String) As String
Dim rstNodes As ADODB.Recordset
Dim rstTemplateTreeLinks As ADODB.Recordset
Dim strTemplateName As String
Dim lngNodeID2 As Long
Dim strTemp As String
Dim strSplit() As String
Dim i As Integer

    On Error GoTo EarlyExit
    
    ADORecordsetOpen "Select * from TemplateTreeLinks where [TREE ID] = " & Chr(39) & ProcessQuotes(strTemplateTreeID) & Chr(39), _
                        g_conData, rstTemplateTreeLinks, adOpenKeyset, adLockOptimistic
    'Set rstTemplateTreeLinks = datData.OpenRecordset("Select * from TemplateTreeLinks where [TREE ID] = " & Chr(39) & ProcessQuotes(strTemplateTreeID) & Chr(39))
    
    If rstTemplateTreeLinks.RecordCount > 0 Then
        rstTemplateTreeLinks.MoveFirst
        lngNodeID2 = rstTemplateTreeLinks![Node_ID2]
        
        ADORecordsetOpen "Select * from Nodes order by Node_ID", _
                        g_conTemplate, rstNodes, adOpenKeyset, adLockOptimistic
                        
        'ADORecordsetOpen "Select * from Nodes order by Node_ID", CallingForm.G_conConnections.Item("Template").Connection, rstNodes, adOpenStatic, adLockPessimistic
    
        With rstNodes
            .Find "Node_ID = " & lngNodeID2
            'strNodes = NodeKey & "/"
            
            If Not .EOF Then
                strTemp = rstNodes!Node_Text & "/"
                Do While !Node_Level <> 0
                    .Find "Node_ID = " & CStr(!Node_ParentID), , adSearchBackward
                    If .EOF Then
                        Exit Do
                    Else
                        'strNodes = strNodes & CStr(!Node_ID) & "/"
                        strTemp = strTemp & rstNodes!Node_Text & "/"
                    End If
                    '.MoveNext
                Loop
            End If
        End With
        
        strSplit = Split(strTemp, "/")
        For i = UBound(strSplit) - 2 To 0 Step -1
            strTemplateName = strTemplateName & strSplit(i) & "/"
        Next
        Erase strSplit
        
        strTemplateName = Translate(347) & "/" & strTemplateName
        TemplateName = strTemplateName
        
        ADORecordsetClose rstNodes
        
        Set rstNodes = Nothing
    Else
        TemplateName = strTemplateTreeID
        Exit Function
    End If
    
    ADORecordsetClose rstTemplateTreeLinks
    ADORecordsetClose rstNodes

EarlyExit:
End Function

Private Function IsInDeletedFolder(strTreeID As String) As Boolean

    If strTreeID = "-1ED" Or strTreeID = "-2ED" Or strTreeID = "DD" Then
        IsInDeletedFolder = True
    Else
        IsInDeletedFolder = False
    End If
    
End Function

Private Sub MoveToDeletedFolder(bytDocType As Byte, strUniqueCode As String)
    Select Case bytDocType
        Case 1
            
            ExecuteNonQuery g_conSADBEL, "Update IMPORT Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 1"
            ExecuteNonQuery g_conData, "Update MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 1"
        Case 2
            ExecuteNonQuery g_conSADBEL, "Update EXPORT Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 2"
            ExecuteNonQuery g_conData, "Update MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 2"
        
        Case 3
            ExecuteNonQuery g_conSADBEL, "Update TRANSIT Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 3"
            ExecuteNonQuery g_conData, "Update MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 3"
            
        Case 7
            ExecuteNonQuery g_conSADBEL, "Update NCTS Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 7"
            ExecuteNonQuery g_conData, "Update MASTERNCTS Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 7"
            
        Case 9
            ExecuteNonQuery g_conSADBEL, "Update [COMBINED NCTS] Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 9"
            ExecuteNonQuery g_conData, "Update MASTERNCTS Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 9"
            
        Case 11
        Case 12
        
        Case 14
            ExecuteNonQuery g_conSADBEL, "Update [PLDA IMPORT] Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 14"
            ExecuteNonQuery g_conData, "Update MASTERPLDA Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 14"
        
        Case 18
            ExecuteNonQuery g_conSADBEL, "Update [PLDA COMBINED] Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 18"
            ExecuteNonQuery g_conData, "Update MASTERPLDA Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 18"
        
    End Select

End Sub

Private Sub DeleteDocument(bytDocType As Byte, strUniqueCode As String)
    Select Case bytDocType
        Case 1
            ExecuteNonQuery g_conSADBEL, "Delete IMPORT where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 1 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 1 and [Tree ID] = 'DD'"
        Case 2
            ExecuteNonQuery g_conSADBEL, "Delete EXPORT where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 2 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 2 and [Tree ID] = 'DD'"
        
        Case 3
            ExecuteNonQuery g_conSADBEL, "Delete TRANSIT where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 3 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTER Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 3 and [Tree ID] = 'DD'"
            
        Case 7
            ExecuteNonQuery g_conSADBEL, "Delete NCTS where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 7 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTERNCTS Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 7 and [Tree ID] = 'DD'"
            
        Case 9
            ExecuteNonQuery g_conSADBEL, "Delete [COMBINED NCTS] where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 9 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTERNCTS Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 9 and [Tree ID] = 'DD'"
            
        Case 11
        Case 12
        
        Case 14
            ExecuteNonQuery g_conSADBEL, "Delete [PPLDA IMPORT] where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 14 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTERPLDA Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 14 and [Tree ID] = 'DD'"
        
        Case 18
            ExecuteNonQuery g_conSADBEL, "Delete [PLDA COMBINED] where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 18 and [Tree ID] = 'DD'"
            ExecuteNonQuery g_conData, "Delete MASTERPLDA Set [TREE ID] = 'DD' where [CODE] = " & Chr(39) & ProcessQuotes(strUniqueCode) & Chr(39) & " and DType = 18 and [Tree ID] = 'DD'"
    End Select

End Sub


Private Function IsRecordToBeAdded(ByRef rs As ADODB.Recordset, _
                                    strTable As String, _
                                    strField As String, _
                                    strBoxCode As String, _
                                    blnWithInstance As Boolean, _
                                    lngUpperLimit As Long, _
                                    lngInstance As Long, _
                                    blnEDISeparate As Boolean) As Boolean
    Dim strBoxValue As String

    strField = "SEARCHEDFIELD"
    
    If blnEDISeparate Then
        strBoxValue = IIf(IsNull(rs.Fields(strField).Value), "", rs.Fields(strField).Value)
        strBoxValue = UCase(strBoxValue)

        Select Case cboCondition.ListIndex
            Case enuCondition.eContains
                IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) > 0, True, False)
            Case enuCondition.eIsExactly
                IsRecordToBeAdded = IIf(strBoxValue = UCase(txtValue.Text), True, False)
            Case enuCondition.eDoesnotContain
                IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) = 0, True, False)
            Case enuCondition.eIsEmpty
                IsRecordToBeAdded = IIf(Len(strBoxValue) = 0, True, False)
            Case enuCondition.eIsNotEmpty
                IsRecordToBeAdded = IIf(Len(strBoxValue) > 0, True, False)
        End Select
    
    Else
    
        If IsNull(rs.Fields(strTable & "_Instance").Value) Then
            IsRecordToBeAdded = False
        Else
            strBoxValue = IIf(IsNull(rs.Fields(strField).Value), "", rs.Fields(strField).Value)
            strBoxValue = UCase(strBoxValue)
            
            If blnWithInstance Then
                If (rs.Fields(strTable & "_Instance").Value Mod lngUpperLimit) = lngInstance Or (rs.Fields(strTable & "_Instance").Value Mod lngUpperLimit = 0 And lngInstance / rs.Fields(strTable & "_Instance").Value > 1) Then
                    Select Case cboCondition.ListIndex
                        Case enuCondition.eContains
                            IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) > 0, True, False)
                        Case enuCondition.eIsExactly
                            IsRecordToBeAdded = IIf(strBoxValue = UCase(txtValue.Text), True, False)
                        Case enuCondition.eDoesnotContain
                            IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) = 0, True, False)
                        Case enuCondition.eIsEmpty
                            IsRecordToBeAdded = IIf(Len(strBoxValue) = 0, True, False)
                        Case enuCondition.eIsNotEmpty
                            IsRecordToBeAdded = IIf(Len(strBoxValue) > 0, True, False)
                    End Select
            
                Else
                    IsRecordToBeAdded = False
                End If
    
            Else
                Select Case cboCondition.ListIndex
                    Case enuCondition.eContains
                        IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) > 0, True, False)
                    Case enuCondition.eIsExactly
                        IsRecordToBeAdded = IIf(strBoxValue = UCase(txtValue.Text), True, False)
                    Case enuCondition.eDoesnotContain
                        IsRecordToBeAdded = IIf(InStr(1, strBoxValue, UCase(txtValue.Text)) = 0, True, False)
                    Case enuCondition.eIsEmpty
                        IsRecordToBeAdded = IIf(Len(strBoxValue) = 0, True, False)
                    Case enuCondition.eIsNotEmpty
                        IsRecordToBeAdded = IIf(Len(strBoxValue) > 0, True, False)
                End Select
            
            End If
        End If
    End If
End Function

'Rachelle Apr 19,2005
Private Sub CheckMenuForGrid()

    Dim ssTool1 As SSActiveToolBars

    Set ssTool1 = SSActiveToolBars1
    
    ssTool1.PopupMenu ssTool1.Tools("ID_Dummy")
    
    Set ssTool1 = Nothing
    
End Sub

Private Sub OpenACopy()
    If lvwItemsFound.ListItems.Count > 0 Then
        clsFindForm.SelectedItemTag = lvwItemsFound.SelectedItem.Tag
        clsFindForm.SelectedItemText = lvwItemsFound.SelectedItem.Text
        clsFindForm.ListSubItems = lvwItemsFound.SelectedItem.ListSubItems(2).Tag
        clsFindForm.SubItems = lvwItemsFound.SelectedItem.SubItems(2)
        If (clsFindForm.SubItems = Translate(1379) Or clsFindForm.SubItems = Translate(1074)) And _
            (clsFindForm.ListSubItems = "40ED" Or clsFindForm.ListSubItems = "50ED") Then
            clsFindForm.strYear = lvwItemsFound.SelectedItem.ListSubItems(5).Tag
        Else
            clsFindForm.strYear = ""
        End If
    
        Select Case lvwItemsFound.SelectedItem.ListSubItems(1).Text
    
            Case "Import"
                CallingForm.LoadDocument edocimport, False, True
                
            Case "Export"
                CallingForm.LoadDocument eDocExport, False, True
                
            Case "Transit"
                CallingForm.LoadDocument eDocOTS, False, True
    
            Case "NCTS"
                CallingForm.LoadDocument eDocNCTS, False, True
    
            Case "Combined NCTS"
                CallingForm.LoadDocument edoccombined, False, True
        
            Case "EDI Departures"
                CallingForm.LoadDocument eDocEDIDepartures, False, True
                
            Case "EDI Arrivals"
                CallingForm.LoadDocument eDocEDIARRIVALS, False, True
                
            Case "PLDA Import"
                CallingForm.LoadDocument eDocPLDAImport, False, True
                
            Case "PLDA Combined"
                CallingForm.LoadDocument eDocPLDACombined, False, True
                
        End Select
    End If
End Sub

Private Sub SaveSpecs(ByVal strDocType As String, Optional ByVal blnUnload As Boolean)
    Dim strSQl As String
    Dim lngCounter As Long
    Dim strColumnAlignments As String
    Dim strCommand As String
    
    strColumnAlignments = ""
    strSettingWidth = ""
    strPosition = ""
    
    For lngCounter = 1 To lvwItemsFound.ColumnHeaders.Count
        If lngCounter = lvwItemsFound.ColumnHeaders.Count Then
            strSettingWidth = strSettingWidth & lvwItemsFound.ColumnHeaders(lngCounter).Width
            strColumnAlignments = strColumnAlignments & lvwItemsFound.ColumnHeaders(lngCounter).Alignment
            strPosition = strPosition & lvwItemsFound.ColumnHeaders(lngCounter).Position
        Else
            strSettingWidth = strSettingWidth & lvwItemsFound.ColumnHeaders(lngCounter).Width & "*****"
            strColumnAlignments = strColumnAlignments & lvwItemsFound.ColumnHeaders(lngCounter).Alignment & "*****"
            strPosition = strPosition & lvwItemsFound.ColumnHeaders(lngCounter).Position & "*****"
        End If
    Next lngCounter
    
    ADORecordsetClose rstFind
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "FINDVIEWCOLUMNS "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "USER_ID = " & UserID & " "
    ADORecordsetOpen strCommand, g_conTemplate, rstFind, adOpenKeyset, adLockOptimistic

    With rstFind
        If Not (.BOF And .EOF) Then
            .MoveFirst
            .Filter = "FVC_DocumentType = '" & strDocType & "'"
        
            If .BOF And .EOF Then
                .AddNew
            End If
        Else
            .AddNew
        End If
        
        .Fields("User_ID").Value = UserID
        .Fields("FVC_ColumnAlignments") = strPosition
        .Fields("FVC_DocumentType").Value = strDocType
        .Fields("FVC_ColumnWidths").Value = strSettingWidth
        .Fields("FVC_GroupHeaders").Value = strView
        'Sort : sort key--> sort order --> 0-asc 1-desc
        .Fields("FVC_Sort").Value = lvwItemsFound.SortKey & "*****" & lvwItemsFound.SortOrder
        If blnUnload = True Then
            .Fields("FVC_ColumnFormat").Value = frm_Find.Width & "*****" & frm_Find.Height & "*****" & frm_Find.Top & "*****" & frm_Find.Left
            conFind.Execute "UPDATE FINDVIEWCOLUMNS SET FVC_COLUMNFORMAT = '" & frm_Find.Width & "*****" & frm_Find.Height & "*****" & frm_Find.Top & "*****" & frm_Find.Left & _
                            "' WHERE USER_ID = " & UserID
            conFind.Execute "UPDATE FINDVIEWCOLUMNS SET FVC_GroupHeaders = '" & strView & "' WHERE USER_ID = " & UserID
        End If
        
        .Update
        .Filter = adFilterNone
        '.ActiveConnection = Nothing
    End With
End Sub

Private Sub LoadListView()
    Dim blnLoaded As Boolean
    Dim strCommand As String
    
    If blnJustLoaded = True Then
        blnLoaded = True
        
            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "* "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "FINDVIEWCOLUMNS "
            strCommand = strCommand & "WHERE "
            strCommand = strCommand & "USER_ID = " & UserID & " "
        ADORecordsetOpen strCommand, g_conTemplate, rstFind, adOpenKeyset, adLockOptimistic
        blnJustLoaded = False
    Else
        blnLoaded = False
    End If
    
    If Not (rstFind.BOF And rstFind.EOF) Then
        'Load
        rstFind.MoveFirst
        
        If icbType.Text <> "" Then
            rstFind.Find "FVC_DocumentType = '" & icbType.Text & "'"
        Else
            rstFind.Find "FVC_DocumentType = 'Any'"
        End If
        
        If Not (rstFind.EOF) Then
            strSettingWidth = rstFind.Fields("FVC_ColumnWidths").Value
            If blnLoaded = True Then
                strView = rstFind.Fields("FVC_GroupHeaders").Value
            End If
            
            strSort = rstFind.Fields("FVC_Sort").Value
            
            If blnJustLoaded3 = True Then
                blnJustLoaded3 = False
                strFormWidth = rstFind.Fields("FVC_ColumnFormat").Value
            Else
                strFormWidth = ""
            End If
            
            strPosition = rstFind.Fields("FVC_ColumnAlignments")
            
            If blnLoaded = True Then
                If strView <> "" Then
                    SSActiveToolBars1.Tools(strView).State = ssChecked
                    ListCheck strView
                End If
            End If
        Else
            strPosition = ""
            strSettingWidth = ""
            strSort = ""
            If blnLoaded = True Then
                strView = ""
            End If
            strFormWidth = ""
        End If
    Else
        strPosition = ""
        strSettingWidth = ""
        strSort = ""
        If blnLoaded = True Then
            strView = ""
        End If
        strFormWidth = ""
    End If
    
End Sub

Public Sub RefreshList()
    Dim strNewFields() As String
    Dim lngCounter As Long
    Dim lngCounter2 As Long
    Dim strSettingWidth2() As String
    
    lvwItemsFound.Visible = False
    Me.MousePointer = vbHourglass
    
    If blnShowFields = False Then
        lvwItemsFound.Visible = True
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        SetListViewWidths strFields
        blnShowFields = False
        SaveSpecs icbType.Text
    End If
    lvwItemsFound.Visible = True
    Me.MousePointer = vbDefault

End Sub

Private Function IsFieldExisting(ByRef rstRecord As ADODB.Recordset, ByVal strField As String) As Boolean
    Dim lngCounter As Long
    
    For lngCounter = 0 To rstRecord.Fields.Count - 1
        If UCase(rstRecord.Fields(lngCounter).Name) = UCase(strField) Then
            IsFieldExisting = True
            Exit For
        Else
            IsFieldExisting = False
        End If
    Next lngCounter
End Function

Private Function GetPLDAFieldValue(ByRef DBToBeSearched As ADODB.Connection, _
                                    ByVal bytDType As Byte, _
                                    ByVal strFieldName As String, _
                                    ByVal strCode As String)
    Dim rstPLDA As ADODB.Recordset
    Dim strSQl As String
    
    strSQl = "SELECT [" & strFieldName & "]"
    
    'CSCLP-248
    If InStr(1, strPLDAFields, "**" & strFieldName & "**") Or InStr(1, UCase$(strPLDAFields), "**" & UCase$(strFieldName) & "**") Then
    'If InStr(1, strPLDAFields, "**" & strFieldName & "**") Then
        If bytDType = 14 Then
            If (Len(strFieldName) = 2 And (InStr(1, PLDAIMPORTHEADER, strFieldName) Or InStr(1, PLDAIMPORTDETAIL, strFieldName))) Or _
                Len(strFieldName) > 2 Then
                strSQl = strSQl & " FROM (((((((((([PLDA IMPORT]" & _
                                " INNER JOIN [PLDA IMPORT HEADER] ON [PLDA IMPORT].CODE = [PLDA IMPORT HEADER].CODE) " & _
                                " INNER JOIN [PLDA IMPORT HEADER ZEGELS] ON ([PLDA IMPORT HEADER].HEADER = [PLDA IMPORT HEADER ZEGELS].HEADER) AND ([PLDA IMPORT HEADER].CODE = [PLDA IMPORT HEADER ZEGELS].CODE))" & _
                                " INNER JOIN [PLDA IMPORT HEADER HANDELAARS] ON ([PLDA IMPORT HEADER].HEADER = [PLDA IMPORT HEADER HANDELAARS].HEADER) AND ([PLDA IMPORT HEADER].CODE = [PLDA IMPORT HEADER HANDELAARS].CODE))" & _
                                " INNER JOIN [PLDA IMPORT DETAIL] ON ([PLDA IMPORT HEADER ZEGELS].HEADER = [PLDA IMPORT DETAIL].HEADER) AND ([PLDA IMPORT HEADER ZEGELS].CODE = [PLDA IMPORT DETAIL].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL BIJZONDERE] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL BIJZONDERE].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL BIJZONDERE].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL BIJZONDERE].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL CONTAINER] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL CONTAINER].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL CONTAINER].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL CONTAINER].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL DOCUMENTEN] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL DOCUMENTEN].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL DOCUMENTEN].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL DOCUMENTEN].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL ZELF] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL ZELF].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL ZELF].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL ZELF].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL HANDELAARS] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL HANDELAARS].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL HANDELAARS].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL HANDELAARS].CODE)) " & _
                                " INNER JOIN [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] ON ([PLDA IMPORT DETAIL].DETAIL = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].DETAIL) AND ([PLDA IMPORT DETAIL].HEADER = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].HEADER) AND ([PLDA IMPORT DETAIL].CODE = [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN].CODE)) " & _
                                " WHERE [PLDA IMPORT HEADER].[CODE] = " & Chr(39) & strCode & Chr(39) & _
                                " ORDER BY [PLDA IMPORT].[DATE LAST MODIFIED] DESC"
            Else
                GetPLDAFieldValue = ""
                Exit Function
            End If
    
        ElseIf bytDType = 18 Then
            If (Len(strFieldName) = 2 And (InStr(1, PLDACOMBINEDHEADER, strFieldName) Or InStr(1, PLDACOMBINEDDETAIL, strFieldName))) Or _
                Len(strFieldName) > 2 Then
                strSQl = strSQl & " FROM (((((((([PLDA COMBINED]" & _
                                " INNER JOIN [PLDA COMBINED HEADER] ON [PLDA COMBINED].CODE = [PLDA COMBINED HEADER].CODE) " & _
                                " INNER JOIN [PLDA COMBINED HEADER ZEGELS] ON ([PLDA COMBINED HEADER].HEADER = [PLDA COMBINED HEADER ZEGELS].HEADER) AND ([PLDA COMBINED HEADER].CODE = [PLDA COMBINED HEADER ZEGELS].CODE))" & _
                                " INNER JOIN [PLDA COMBINED HEADER HANDELAARS] ON ([PLDA COMBINED HEADER ZEGELS].HEADER = [PLDA COMBINED HEADER HANDELAARS].HEADER) AND ([PLDA COMBINED HEADER HANDELAARS].CODE = [PLDA COMBINED HEADER ZEGELS].CODE))" & _
                                " INNER JOIN [PLDA COMBINED DETAIL] ON ([PLDA COMBINED HEADER HANDELAARS].HEADER = [PLDA COMBINED DETAIL].HEADER) AND ([PLDA COMBINED HEADER HANDELAARS].CODE = [PLDA COMBINED DETAIL].CODE)) " & _
                                " INNER JOIN [PLDA COMBINED DETAIL BIJZONDERE] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL BIJZONDERE].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL BIJZONDERE].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL BIJZONDERE].CODE)) " & _
                                " INNER JOIN [PLDA COMBINED DETAIL CONTAINER] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL CONTAINER].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL CONTAINER].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL CONTAINER].CODE)) " & _
                                " INNER JOIN [PLDA COMBINED DETAIL DOCUMENTEN] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL DOCUMENTEN].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL DOCUMENTEN].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL DOCUMENTEN].CODE)) " & _
                                " INNER JOIN [PLDA COMBINED DETAIL HANDELAARS] ON ([PLDA COMBINED DETAIL].DETAIL = [PLDA COMBINED DETAIL HANDELAARS].DETAIL) AND ([PLDA COMBINED DETAIL].HEADER = [PLDA COMBINED DETAIL HANDELAARS].HEADER) AND ([PLDA COMBINED DETAIL].CODE = [PLDA COMBINED DETAIL HANDELAARS].CODE)) " & _
                                " WHERE [PLDA COMBINED HEADER].[CODE] = " & Chr(39) & strCode & Chr(39) & _
                                " ORDER BY [PLDA COMBINED].[DATE LAST MODIFIED] DESC"
            Else
                GetPLDAFieldValue = ""
                Exit Function
            End If
        End If
        
        ADORecordsetOpen strSQl, DBToBeSearched, rstPLDA, adOpenKeyset, adLockOptimistic
        'Set rstPLDA = DBToBeSearched.OpenRecordset(strSQl)
        
        If rstPLDA.RecordCount <> 0 Then ' allan added june18-------
            If Not (rstPLDA.EOF Or rstPLDA.BOF) Then
                GetPLDAFieldValue = IIf(IsNull(rstPLDA.Fields(strFieldName)), "", rstPLDA.Fields(strFieldName).Value)
            Else
                GetPLDAFieldValue = ""
            End If
        Else
            GetPLDAFieldValue = ""
        End If 'allan end --------------
        
        ADORecordsetClose rstPLDA
    Else
        GetPLDAFieldValue = ""
    End If
End Function

Private Function IsFeatureLicensed(ByVal FeatureText As String) As Boolean
    'Added by BCo 2006-05-03
    'Licensing prevents access to Open A Copy if selected doc type is unlicensed
    'though, if they find a way to open, the toolbars are disabled anyways. Aesthetic purposes?
    Select Case UCase$(FeatureText)
        Case "IMPORT"
            IsFeatureLicensed = clsFindForm.LicSADI
        Case "EXPORT"
            IsFeatureLicensed = clsFindForm.LicSADET
        Case "TRANSIT"
            IsFeatureLicensed = clsFindForm.LicSADET
        Case "NCTS"
            IsFeatureLicensed = clsFindForm.LicSADTC
        Case "COMBINED NCTS"
            IsFeatureLicensed = clsFindForm.LicSADTC
        Case "EDI DEPARTURES"
            IsFeatureLicensed = clsFindForm.LicNCTS
        Case "EDI ARRIVALS"
            IsFeatureLicensed = clsFindForm.LicNCTS
        Case "PLDA IMPORT"
            IsFeatureLicensed = clsFindForm.LicPLDAI
        Case "PLDA EXPORT"
            IsFeatureLicensed = clsFindForm.LicPLDAE
        Case "PLDA COMBINED"
            IsFeatureLicensed = clsFindForm.LicPLDAC
        Case Else
            IsFeatureLicensed = False
    End Select
End Function

Private Sub SetPLDAProperties(ByVal bytDocType As Byte)
    Dim lngCounter As Long
    Dim rstPLDAChecker As ADODB.Recordset
    Dim lngTableCounter As Long
    
    strPLDAFields = "**"
    Erase strSQLFields
        
    If bytDocType = 14 Or bytDocType = 18 Then
        If bytDocType = 14 Then
            ReDim strSQLFields(10)
            strSQLFields(0) = "SELECT TOP 1 * FROM [PLDA IMPORT]"
            strSQLFields(1) = "SELECT TOP 1 * FROM [PLDA IMPORT HEADER]"
            strSQLFields(2) = "SELECT TOP 1 * FROM [PLDA IMPORT HEADER ZEGELS]"
            strSQLFields(3) = "SELECT TOP 1 * FROM [PLDA IMPORT DETAIL]"
            strSQLFields(4) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL BIJZONDERE]"
            strSQLFields(5) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL CONTAINER]"
            strSQLFields(6) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL DOCUMENTEN]"
            strSQLFields(7) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL ZELF]"
            'rachelle 082806
            'for the newly added PLDA IMPORT tables.
            strSQLFields(8) = "SELECT TOP 1 * FROM  [PLDA IMPORT HEADER HANDELAARS]"
            strSQLFields(9) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL HANDELAARS]"
            strSQLFields(10) = "SELECT TOP 1 * FROM  [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN]"
            
        Else
            ReDim strSQLFields(8)
            strSQLFields(0) = "SELECT TOP 1 * FROM [PLDA COMBINED]"
            strSQLFields(1) = "SELECT TOP 1 * FROM [PLDA COMBINED HEADER]"
            strSQLFields(2) = "SELECT TOP 1 * FROM [PLDA COMBINED HEADER ZEGELS]"
            strSQLFields(3) = "SELECT TOP 1 * FROM [PLDA COMBINED HEADER HANDELAARS]"
            strSQLFields(4) = "SELECT TOP 1 * FROM  [PLDA COMBINED DETAIL]"
            strSQLFields(5) = "SELECT TOP 1 * FROM  [PLDA COMBINED DETAIL BIJZONDERE]"
            strSQLFields(6) = "SELECT TOP 1 * FROM  [PLDA COMBINED DETAIL CONTAINER]"
            strSQLFields(7) = "SELECT TOP 1 * FROM  [PLDA COMBINED DETAIL DOCUMENTEN]"
            strSQLFields(8) = "SELECT TOP 1 * FROM  [PLDA COMBINED DETAIL HANDELAARS]"
        End If
        
        For lngTableCounter = 0 To UBound(strSQLFields)
            ADORecordsetOpen strSQLFields(lngTableCounter), g_conSADBEL, rstPLDAChecker, adOpenKeyset, adLockOptimistic
            'Set rstPLDAChecker = datSADBEL.OpenRecordset(strSQLFields(lngTableCounter))
            
            For lngCounter = 0 To rstPLDAChecker.Fields.Count - 1
                strPLDAFields = strPLDAFields & rstPLDAChecker.Fields(lngCounter).Name & "**"
            Next lngCounter
            
            ADORecordsetClose rstPLDAChecker
            
        Next lngTableCounter
    End If

End Sub

Private Function GetFieldValue(ByVal rstResults As ADODB.Recordset, ByVal strFieldName As String)
        
    If rstResults.RecordCount > 0 Then ' allan added june18-------
        If Not (rstResults.EOF Or rstResults.BOF) Then
            GetFieldValue = IIf(IsNull(rstResults.Fields(strFieldName).Value), "", rstResults.Fields(strFieldName).Value)
        Else
            GetFieldValue = ""
        End If
    Else
        GetFieldValue = ""
    End If 'allan end --------------
End Function

Public Property Get SearchInProgress() As Boolean
    SearchInProgress = m_blnSearchInProgress
End Property
