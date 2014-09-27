VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBoxProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "428"
   ClientHeight    =   7110
   ClientLeft      =   2415
   ClientTop       =   2865
   ClientWidth     =   7785
   ClipControls    =   0   'False
   Icon            =   "frmBoxProperties.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "428"
   Begin VB.CommandButton cmdTransact 
      Caption         =   "OK"
      Height          =   345
      Index           =   0
      Left            =   3840
      TabIndex        =   43
      Tag             =   "178"
      Top             =   6705
      Width           =   1200
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "&Apply"
      Height          =   345
      Index           =   2
      Left            =   6435
      TabIndex        =   45
      Tag             =   "180"
      Top             =   6705
      Width           =   1200
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   5137
      TabIndex        =   44
      Tag             =   "179"
      Top             =   6705
      Width           =   1200
   End
   Begin TabDlg.SSTab tabPlatform 
      Height          =   6510
      Left            =   120
      TabIndex        =   46
      Tag             =   "150"
      Top             =   105
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   11483
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "150"
      TabPicture(0)   =   "frmBoxProperties.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEdit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDefinition"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDescription"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraAction"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "429"
      TabPicture(1)   =   "frmBoxProperties.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTabOrder"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "430"
      TabPicture(2)   =   "frmBoxProperties.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSkip"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "141"
      TabPicture(3)   =   "frmBoxProperties.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraPicklist"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "431"
      TabPicture(4)   =   "frmBoxProperties.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblBoxProp(6)"
      Tab(4).Control(1)=   "txtEmptyField"
      Tab(4).Control(2)=   "fraGenDefaultValue"
      Tab(4).Control(3)=   "cboEmptyField"
      Tab(4).Control(4)=   "fraUserDefaultValue"
      Tab(4).ControlCount=   5
      Begin VB.Frame fraUserDefaultValue 
         Caption         =   "User's default values"
         ClipControls    =   0   'False
         Height          =   2280
         Left            =   -74880
         TabIndex        =   64
         Tag             =   "460"
         Top             =   3165
         Width           =   7215
         Begin VB.ComboBox cboDefaultValue 
            Height          =   315
            Index           =   1
            Left            =   2670
            TabIndex        =   39
            Text            =   "cboDefaultValue"
            Top             =   525
            Width           =   3090
         End
         Begin VB.ComboBox cboLogicalID 
            Height          =   315
            Index           =   1
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   525
            Width           =   2490
         End
         Begin MSComctlLib.ListView lvwUserDefaultValue 
            Height          =   1200
            Left            =   120
            TabIndex        =   56
            Top             =   915
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   2117
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4322
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Empty"
            Enabled         =   0   'False
            Height          =   345
            Index           =   9
            Left            =   5940
            TabIndex        =   42
            Tag             =   "402"
            Top             =   1755
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "Re&move"
            Enabled         =   0   'False
            Height          =   345
            Index           =   10
            Left            =   5940
            TabIndex        =   41
            Tag             =   "401"
            Top             =   1350
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   345
            Index           =   11
            Left            =   5925
            TabIndex        =   40
            Tag             =   "251"
            Top             =   945
            Width           =   1110
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Index           =   12
            Left            =   2670
            TabIndex        =   78
            Tag             =   "451"
            Top             =   270
            Width           =   405
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Assigned Logical IDs"
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   77
            Tag             =   "604"
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.ComboBox cboEmptyField 
         Height          =   315
         Left            =   -72800
         TabIndex        =   31
         Text            =   "cboEmptyField"
         Top             =   450
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Frame fraGenDefaultValue 
         Caption         =   "General default value"
         ClipControls    =   0   'False
         Height          =   2280
         Left            =   -74880
         TabIndex        =   63
         Tag             =   "458"
         Top             =   825
         Width           =   7215
         Begin VB.ComboBox cboDefaultValue 
            Height          =   315
            Index           =   0
            Left            =   2685
            TabIndex        =   34
            Text            =   "cboDefaultValue"
            Top             =   540
            Width           =   3090
         End
         Begin VB.ComboBox cboLogicalID 
            Height          =   315
            Index           =   0
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   540
            Width           =   2490
         End
         Begin MSComctlLib.ListView lvwGenDefaultValue 
            Height          =   1200
            Left            =   120
            TabIndex        =   55
            Top             =   915
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   2117
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4322
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Empty"
            Enabled         =   0   'False
            Height          =   345
            Index           =   6
            Left            =   5940
            TabIndex        =   37
            Tag             =   "402"
            Top             =   1770
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "Re&move"
            Enabled         =   0   'False
            Height          =   345
            Index           =   7
            Left            =   5940
            TabIndex        =   36
            Tag             =   "401"
            Top             =   1365
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   345
            Index           =   8
            Left            =   5940
            TabIndex        =   35
            Tag             =   "251"
            Top             =   975
            Width           =   1110
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Index           =   10
            Left            =   2685
            TabIndex        =   76
            Tag             =   "451"
            Top             =   285
            Width           =   405
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Logical ID"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   75
            Tag             =   "717"
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.Frame fraPicklist 
         Caption         =   "Picklist"
         ClipControls    =   0   'False
         Height          =   3300
         Left            =   -74880
         TabIndex        =   62
         Tag             =   "398"
         Top             =   450
         Width           =   7215
         Begin MSComctlLib.ListView lvwPicklist 
            Height          =   2535
            Left            =   135
            TabIndex        =   30
            Top             =   615
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "From"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Validation"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CheckBox chkAutoAdd 
            Caption         =   "Auto add"
            Height          =   195
            Left            =   165
            TabIndex        =   29
            Tag             =   "456"
            Top             =   300
            Width           =   3420
         End
      End
      Begin VB.Frame fraSkip 
         Caption         =   "Skip"
         ClipControls    =   0   'False
         Height          =   3300
         Left            =   -74880
         TabIndex        =   61
         Tag             =   "754"
         Top             =   450
         Width           =   7215
         Begin VB.ComboBox cboSkipBoxCode 
            Height          =   315
            Left            =   135
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   675
            Width           =   765
         End
         Begin VB.ComboBox cboBoxValue 
            Height          =   315
            Index           =   1
            Left            =   985
            TabIndex        =   23
            Top             =   675
            Width           =   2655
         End
         Begin VB.ComboBox cboSkipPosition 
            Height          =   315
            Left            =   3725
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   660
            Width           =   690
         End
         Begin VB.ComboBox cboEmptyBox 
            Height          =   315
            Index           =   1
            Left            =   4500
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   675
            Width           =   1260
         End
         Begin MSComctlLib.ListView lvwSkip 
            Height          =   2100
            Left            =   120
            TabIndex        =   54
            Top             =   1050
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3704
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   1367
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   4762
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2381
            EndProperty
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   345
            Index           =   3
            Left            =   5925
            TabIndex        =   26
            Tag             =   "251"
            Top             =   1080
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "Re&move"
            Enabled         =   0   'False
            Height          =   345
            Index           =   4
            Left            =   5925
            TabIndex        =   27
            Tag             =   "401"
            Top             =   1470
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Empty"
            Enabled         =   0   'False
            Height          =   345
            Index           =   5
            Left            =   5925
            TabIndex        =   28
            Tag             =   "402"
            Top             =   1875
            Width           =   1110
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Box"
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   71
            Tag             =   "329"
            Top             =   405
            Width           =   270
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Index           =   4
            Left            =   985
            TabIndex        =   70
            Tag             =   "451"
            Top             =   405
            Width           =   405
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Position"
            Height          =   195
            Index           =   7
            Left            =   3725
            TabIndex        =   69
            Tag             =   "454"
            Top             =   405
            Width           =   615
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Empty Box"
            Height          =   195
            Index           =   3
            Left            =   4500
            TabIndex        =   68
            Tag             =   "455"
            Top             =   405
            Width           =   1230
         End
      End
      Begin VB.Frame fraTabOrder 
         Caption         =   "Tab Order"
         ClipControls    =   0   'False
         Height          =   3300
         Left            =   -74880
         TabIndex        =   57
         Tag             =   "450"
         Top             =   450
         Width           =   7215
         Begin VB.ComboBox cboBoxValue 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   675
            Width           =   2430
         End
         Begin VB.ComboBox cboGoto 
            Height          =   315
            Left            =   2640
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   675
            Width           =   765
         End
         Begin VB.ComboBox cboEmptyBox 
            Height          =   315
            Index           =   0
            Left            =   3495
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   675
            Width           =   2265
         End
         Begin MSComctlLib.ListView lvwTabOrder 
            Height          =   2100
            Left            =   120
            TabIndex        =   53
            Top             =   1050
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3704
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
               Object.Width           =   4322
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   1535
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   3969
            EndProperty
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Empty"
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   5925
            TabIndex        =   21
            Tag             =   "402"
            Top             =   1875
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "Re&move"
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   5925
            TabIndex        =   20
            Tag             =   "401"
            Top             =   1470
            Width           =   1110
         End
         Begin VB.CommandButton cmdRecordOperation 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            Left            =   5925
            TabIndex        =   19
            Tag             =   "251"
            Top             =   1080
            Width           =   1110
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   74
            Tag             =   "451"
            Top             =   405
            Width           =   405
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Go to"
            Height          =   195
            Index           =   1
            Left            =   2655
            TabIndex        =   73
            Tag             =   "452"
            Top             =   405
            Width           =   390
         End
         Begin VB.Label lblBoxProp 
            AutoSize        =   -1  'True
            Caption         =   "Clear intermediate box"
            Height          =   195
            Index           =   2
            Left            =   3495
            TabIndex        =   72
            Tag             =   "602"
            Top             =   405
            Width           =   1560
         End
      End
      Begin VB.Frame fraAction 
         Caption         =   "&Action"
         ClipControls    =   0   'False
         Height          =   3135
         Left            =   135
         TabIndex        =   58
         Tag             =   "434"
         Top             =   3255
         Width           =   7215
         Begin VB.CheckBox chkAction 
            Caption         =   "Validate value with picklist"
            Height          =   195
            Index           =   10
            Left            =   585
            TabIndex        =   15
            Tag             =   "2144"
            Top             =   2820
            Width           =   6495
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Relate  L1  to  S1"
            Height          =   195
            Index           =   9
            Left            =   585
            TabIndex        =   14
            Tag             =   "449"
            Top             =   2565
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Calculate and propose remaining customs value"
            Height          =   195
            Index           =   8
            Left            =   585
            TabIndex        =   10
            Tag             =   "445"
            Top             =   1530
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Deactivate for sequential tabbing (Default)"
            Height          =   195
            Index           =   0
            Left            =   585
            TabIndex        =   5
            Tag             =   "440"
            Top             =   240
            Value           =   1  'Checked
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Deactivate for sequential tabbing (Active Document)"
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   6
            Tag             =   "441"
            Top             =   495
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Check VAT-number"
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   7
            Tag             =   "442"
            Top             =   765
            Value           =   1  'Checked
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Calculate and propose remaining Net Weight as default value"
            Height          =   195
            Index           =   3
            Left            =   585
            TabIndex        =   8
            Tag             =   "443"
            Top             =   1020
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Calculate and propose remaining Number of Items as default"
            Height          =   195
            Index           =   4
            Left            =   585
            TabIndex        =   9
            Tag             =   "444"
            Top             =   1275
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Copy the content of this box to the next header/detail"
            Height          =   195
            Index           =   5
            Left            =   585
            TabIndex        =   11
            Tag             =   "446"
            Top             =   1785
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Changes are only allowed when in Header 1"
            Height          =   195
            Index           =   6
            Left            =   585
            TabIndex        =   12
            Tag             =   "447"
            Top             =   2055
            Width           =   6255
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Send only when in Header 1"
            Height          =   195
            Index           =   7
            Left            =   585
            TabIndex        =   13
            Tag             =   "448"
            Top             =   2310
            Width           =   6255
         End
      End
      Begin VB.Frame fraDescription 
         Caption         =   "Description"
         ClipControls    =   0   'False
         Height          =   1515
         Left            =   135
         TabIndex        =   59
         Tag             =   "292"
         Top             =   1665
         Width           =   7215
         Begin VB.PictureBox pnlDescEnglish 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   180
            ScaleHeight     =   315
            ScaleWidth      =   6795
            TabIndex        =   85
            Top             =   300
            Width           =   6795
            Begin VB.TextBox txtDescription 
               Height          =   315
               Index           =   0
               Left            =   885
               TabIndex        =   86
               Top             =   0
               Width           =   5820
            End
            Begin VB.Label lblDescription 
               AutoSize        =   -1  'True
               Caption         =   "English"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   87
               Tag             =   "992"
               Top             =   30
               Width           =   870
            End
         End
         Begin VB.PictureBox pnlDescDutch 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   180
            ScaleHeight     =   375
            ScaleWidth      =   6855
            TabIndex        =   82
            Top             =   700
            Width           =   6855
            Begin VB.TextBox txtDescription 
               Height          =   300
               Index           =   1
               Left            =   885
               TabIndex        =   83
               Top             =   0
               Width           =   5820
            End
            Begin VB.Label lblDescription 
               AutoSize        =   -1  'True
               Caption         =   "Dutch"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   84
               Tag             =   "993"
               Top             =   30
               Width           =   915
            End
         End
         Begin VB.PictureBox pnlDescFrench 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   180
            ScaleHeight     =   375
            ScaleWidth      =   6915
            TabIndex        =   79
            Top             =   1080
            Width           =   6915
            Begin VB.TextBox txtDescription 
               Height          =   300
               Index           =   2
               Left            =   890
               TabIndex        =   80
               Top             =   0
               Width           =   5820
            End
            Begin VB.Label lblDescription 
               AutoSize        =   -1  'True
               Caption         =   "French"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   81
               Tag             =   "994"
               Top             =   30
               Width           =   855
            End
         End
      End
      Begin VB.Frame fraDefinition 
         Caption         =   "Definition"
         ClipControls    =   0   'False
         Height          =   1125
         Left            =   135
         TabIndex        =   65
         Tag             =   "432"
         Top             =   450
         Width           =   3570
         Begin VB.TextBox txtDefinition 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   1860
            TabIndex        =   49
            Text            =   "0"
            Top             =   705
            Width           =   525
         End
         Begin VB.TextBox txtDefinition 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2505
            TabIndex        =   48
            Text            =   "0"
            Top             =   330
            Width           =   345
         End
         Begin VB.TextBox txtDefinition 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1860
            TabIndex        =   47
            Text            =   "0"
            Top             =   330
            Width           =   525
         End
         Begin VB.OptionButton optDefinition 
            Caption         =   "Alphanumerical"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   1
            Tag             =   "588"
            Top             =   705
            Width           =   1590
         End
         Begin VB.OptionButton optDefinition 
            Caption         =   "Numerical"
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   195
            TabIndex        =   0
            Tag             =   "587"
            Top             =   330
            Width           =   1350
         End
      End
      Begin VB.Frame fraEdit 
         Caption         =   "Edit"
         ClipControls    =   0   'False
         Height          =   1125
         Left            =   3810
         TabIndex        =   66
         Tag             =   "479"
         Top             =   450
         Width           =   3525
         Begin VB.OptionButton optEdit 
            Caption         =   "Overwrite"
            Height          =   240
            Index           =   2
            Left            =   630
            TabIndex        =   4
            Tag             =   "436"
            Top             =   765
            Width           =   1665
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Default"
            Height          =   240
            Index           =   0
            Left            =   630
            TabIndex        =   2
            Tag             =   "480"
            Top             =   285
            Width           =   1665
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Insert"
            Height          =   240
            Index           =   1
            Left            =   630
            TabIndex        =   3
            Tag             =   "435"
            Top             =   525
            Width           =   1665
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Alignment"
         ClipControls    =   0   'False
         Height          =   1125
         Left            =   5640
         TabIndex        =   67
         Tag             =   "433"
         Top             =   450
         Visible         =   0   'False
         Width           =   1695
         Begin VB.OptionButton Option3 
            Caption         =   "Center"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   52
            Tag             =   "222"
            Top             =   765
            Width           =   1515
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Left"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   50
            Tag             =   "221"
            Top             =   285
            Width           =   1515
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Right"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   51
            Tag             =   "223"
            Top             =   525
            Width           =   1515
         End
      End
      Begin VB.TextBox txtEmptyField 
         Height          =   300
         Left            =   -72780
         TabIndex        =   32
         Text            =   " "
         Top             =   450
         Width           =   2250
      End
      Begin VB.Label lblBoxProp 
         AutoSize        =   -1  'True
         Caption         =   "Empty field value"
         Height          =   195
         Index           =   6
         Left            =   -74730
         TabIndex        =   60
         Tag             =   "457"
         Top             =   480
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmBoxProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***********************************************************************************************************************************************************
' ***********************************************************************************************************************************************************
' ***********************************************************************************************************************************************************
Option Explicit

' constants here
Private Const OPT_NUMERIC = 0
Private Const OPT_ALPHANUMERIC = 1

Private Const TXT_NUMERIC = 0
Private Const TXT_NUMERIC_DECIMAL = 1
Private Const TXT_ALPHANUMERIC = 2

Private Const OPT_DEFAULT = 0
Private Const OPT_INSERT = 1
Private Const OPT_OVERWRITE = 2

Private Const TXT_DESC_ENGLISH = 0
Private Const TXT_DESC_DUTCH = 1
Private Const TXT_DESC_FRENCH = 2

Private Const CHK_DEACTIVATE_SEQ_TAB_DEFAULT = 0
Private Const CHK_DEACTIVATE_SEQ_TAB_ACTIVE = 1
Private Const CHK_CHECK_VAT_NO = 2
Private Const CHK_CALCULATE_NET_WEIGHT = 3
Private Const CHK_CALCULATE_NO_OF_ITEMS = 4
Private Const CHK_COPY_NEXT_H_AND_D = 5
Private Const CHK_CHANGE_WHEN_H_IS_1 = 6
Private Const CHK_SEND_WHEN_H_IS_1 = 7
Private Const CHK_CALCULATE_CUSTOM_VALUE = 8
Private Const CHK_RELATE_L1_TO_S1 = 9
Private Const CHK_VALIDATE_VALUE = 10

Private Const CBO_BOXVALUE_TABORDER = 0
Private Const CBO_BOXVALUE_SKIP = 1

Private Const CBO_LOGID_GENERAL = 0
Private Const CBO_LOGID_USER = 1

Private Const CMD_OK = 0
Private Const CMD_CANCEL = 1
Private Const CMD_APPLY = 2

Private Const CMD_TABORDER_SAVE = 0
Private Const CMD_TABORDER_REMOVE = 1
Private Const CMD_TABORDER_EMPTY = 2

Private Const CMD_SKIP_SAVE = 3
Private Const CMD_SKIP_REMOVE = 4
Private Const CMD_SKIP_EMPTY = 5

Private Const CMD_LOGID_GENERAL_EMPTY = 6
Private Const CMD_LOGID_GENERAL_REMOVE = 7
Private Const CMD_LOGID_GENERAL_SAVE = 8

Private Const CMD_LOGID_USER_EMPTY = 9
Private Const CMD_LOGID_USER_REMOVE = 10
Private Const CMD_LOGID_USER_SAVE = 11


Dim mvarBoxDefaultAdminTable As String      ' 1
Dim mvarBoxDefaultValueTable As String      ' 2
Dim mvarDefaultUserTable As String              ' 3
Dim mvarCodisheetType As cpiCodiSheetTypeEnums ' 4
Dim mvarActiveBoxCode As String ' 5
Dim mvarUserNo As Long ' 6
Dim mvarActiveConnection As ADODB.Connection  ' 7
Dim mvarActiveDocument As String ' 8
Dim mvarActiveType As String ' 9
Dim mvarActiveLanguage As String ' 10
Dim mvarResourceHandle As Long

' 1
Public Property Get BoxDefaultAdminTable() As String
    BoxDefaultAdminTable = mvarBoxDefaultAdminTable
End Property

Public Property Let BoxDefaultAdminTable(ByVal vNewValue As String)
    mvarBoxDefaultAdminTable = vNewValue
End Property

' 2
Public Property Get BoxDefaultValueTable() As String
    BoxDefaultValueTable = mvarBoxDefaultValueTable
End Property

Public Property Let BoxDefaultValueTable(ByVal vNewValue As String)
    mvarBoxDefaultValueTable = vNewValue
End Property

' 3
Public Property Get DefaultUserTable() As String
    DefaultUserTable = mvarDefaultUserTable
End Property

Public Property Let DefaultUserTable(ByVal vNewValue As String)
    mvarDefaultUserTable = vNewValue
End Property

' 4
Public Property Get CodisheetType() As cpiCodiSheetTypeEnums
    CodisheetType = mvarCodisheetType
End Property

Public Property Let CodisheetType(ByVal vNewValue As cpiCodiSheetTypeEnums)
    mvarCodisheetType = vNewValue
    
    Select Case mvarCodisheetType
    
        Case cpiImportCodisheet
        
            mvarActiveType = "I"
            mvarActiveDocument = "Import"
            mvarBoxDefaultAdminTable = "BOX DEFAULT IMPORT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE IMPORT"
            mvarDefaultUserTable = "DEFAULT USER IMPORT"
        
        Case cpiExportCodisheet
        
            mvarActiveType = "E"
            mvarActiveDocument = "Export/Transit"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EXPORT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EXPORT"
            mvarDefaultUserTable = "DEFAULT USER EXPORT"
        
        Case cpiTransitCodisheet
        
            mvarActiveType = "T"
            mvarActiveDocument = "Export/Transit"
            mvarBoxDefaultAdminTable = "BOX DEFAULT TRANSIT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE TRANSIT"
            mvarDefaultUserTable = "DEFAULT USER TRANSIT"

        Case cpiSadbelNCTSCodisheet
        
            mvarActiveType = "N"
            mvarActiveDocument = "Transit NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT TRANSIT NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE TRANSIT NCTS"
            mvarDefaultUserTable = "DEFAULT USER TRANSIT NCTS"
        
        Case cpiCombinedNCTSCodisheet
        
            mvarActiveType = "C"
            mvarActiveDocument = "Combined NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT COMBINED NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE COMBINED NCTS"
            mvarDefaultUserTable = "DEFAULT USER COMBINED NCTS"
        
        Case cpiDepartureIE15Codisheet
        
            mvarActiveType = "D"
            mvarActiveDocument = "EDI NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS"
        
        Case cpiArrivalIE07Codisheet
            
            mvarActiveType = "A"
            mvarActiveDocument = "EDI NCTS2"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS2 ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS2"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS2"
        
        Case cpiArrivalIE44Codisheet
        
            mvarActiveType = "U"
            mvarActiveDocument = "EDI NCTS IE44"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS IE44 ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS IE44"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS IE44"
        
    End Select
    
End Property

' 5
Public Property Get ActiveBoxCode() As String
    ActiveBoxCode = mvarActiveBoxCode
End Property

Public Property Let ActiveBoxCode(ByVal vNewValue As String)
    mvarActiveBoxCode = vNewValue
End Property

' 6
Public Property Get UserNo() As Long
    UserNo = mvarUserNo
End Property

Public Property Let UserNo(ByVal vNewValue As Long)
    mvarUserNo = vNewValue
End Property

' 8
Public Property Get ActiveDocument() As String
    ActiveDocument = mvarActiveDocument
End Property

Public Property Let ActiveDocument(ByVal vNewValue As String)
    mvarActiveDocument = vNewValue
End Property

' 9
Public Property Get ActiveType() As String
    ActiveType = mvarActiveType
End Property

Public Property Let ActiveType(ByVal vNewValue As String)
    mvarActiveType = vNewValue
End Property

' 10
Public Property Get ActiveLanguage() As String
    ActiveLanguage = mvarActiveLanguage
End Property

Public Property Let ActiveLanguage(ByVal vNewValue As String)
    mvarActiveLanguage = vNewValue
End Property



'' functions here
Public Function ShowForm(ByRef OwnerForm As Form, _
                                            ByRef ActiveConnection As ADODB.Connection, _
                                            ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                            ByRef ActiveBoxCode As String, _
                                            ByRef ActiveDocument As String, _
                                            ByRef UserNo As Long, _
                                            ByRef ActiveLanguage As String, _
                                            ByRef ResourceHandle As Long) As Boolean
'
    ' load frmBoxProperties here

    mvarActiveBoxCode = ActiveBoxCode
    mvarActiveDocument = ActiveDocument
    mvarActiveLanguage = UCase$(ActiveLanguage)
    mvarResourceHandle = ResourceHandle
    
    CodisheetType = ActiveCodisheet
    
    mvarUserNo = UserNo
    Set mvarActiveConnection = ActiveConnection

    Set Me.Icon = OwnerForm.Icon
    Screen.MousePointer = vbDefault
    
    Me.Show vbModal

    ' @@@

End Function

Private Sub cboBoxValue_Change(Index As Integer)

    Select Case Index
    
        Case CBO_BOXVALUE_TABORDER
                    
            If ((cboBoxValue(Index).Text = "") Or (cboGoto.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboBoxValue(Index).Text <> "") And (cboGoto.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = True
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If

        Case CBO_BOXVALUE_SKIP
                    
            If ((cboSkipBoxCode.Text = "") Or (cboBoxValue(Index).Text = "") Or _
                (cboSkipPosition.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboSkipBoxCode.Text <> "") And (cboBoxValue(Index).Text <> "") And _
                (cboSkipPosition.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = True
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If


    End Select

End Sub

Private Sub cboBoxValue_Click(Index As Integer)

    Select Case Index
    
        Case CBO_BOXVALUE_TABORDER
                    
            If ((cboBoxValue(Index).Text = "") Or (cboGoto.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboBoxValue(Index).Text <> "") And (cboGoto.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = True
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If

        Case CBO_BOXVALUE_SKIP
                    
            If ((cboBoxValue(Index).Text = "") Or (cboGoto.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboBoxValue(Index).Text <> "") And (cboGoto.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = True
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If

    End Select

End Sub

Private Sub cboDefaultValue_Change(Index As Integer)

    Select Case Index
    
        Case CBO_LOGID_GENERAL
                    
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
        
        Case CBO_LOGID_USER
                
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
    
    End Select

End Sub

Private Sub cboDefaultValue_Click(Index As Integer)

    Select Case Index
    
        Case CBO_LOGID_GENERAL
                    
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
        
        Case CBO_LOGID_USER
                
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
    
    End Select

End Sub

Private Sub cboEmptyBox_Click(Index As Integer)
    
    Select Case Index
    
        Case CBO_BOXVALUE_TABORDER
                    
            If ((cboBoxValue(Index).Text = "") Or (cboGoto.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboBoxValue(Index).Text <> "") And (cboGoto.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = True
                cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If

        Case CBO_BOXVALUE_SKIP
                    
            If ((cboSkipBoxCode.Text = "") Or (cboBoxValue(Index).Text = "") Or _
                (cboSkipPosition.Text = "") Or (cboEmptyBox(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf ((cboSkipBoxCode.Text <> "") And (cboBoxValue(Index).Text <> "") And _
                (cboSkipPosition.Text <> "") And (cboEmptyBox(Index).Text <> "")) Then
            
                cmdRecordOperation(CMD_SKIP_SAVE).Enabled = True
                cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
                cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If

    End Select

End Sub

Private Sub cboGoto_Click()

    If ((cboBoxValue(CBO_BOXVALUE_TABORDER).Text = "") Or (cboGoto.Text = "") Or (cboEmptyBox(CBO_BOXVALUE_TABORDER).Text = "")) Then
    
        cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = False
        cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
        cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = False
    
    ElseIf ((cboBoxValue(CBO_BOXVALUE_TABORDER).Text <> "") And (cboGoto.Text <> "") And (cboEmptyBox(CBO_BOXVALUE_TABORDER).Text <> "")) Then
    
        cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = True
        cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
        cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = True
    
    End If

End Sub

Private Sub cboLogicalID_Click(Index As Integer)

    Select Case Index
    
        Case CBO_LOGID_GENERAL
                    
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
        
        Case CBO_LOGID_USER
                
            If ((cboLogicalID(Index).Text = "") Or (cboDefaultValue(Index).Text = "")) Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = False
            
            ElseIf (cboLogicalID(Index).Text <> "") And (cboDefaultValue(Index).Text <> "") Then
            
                cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = True
                cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
                cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
                cmdTransact(CMD_APPLY).Enabled = True
            
            End If
    
    End Select

End Sub

Private Sub cboSkipBoxCode_Click()

    If ((cboSkipBoxCode.Text = "") Or (cboBoxValue(CBO_BOXVALUE_SKIP).Text = "") Or _
        (cboSkipPosition.Text = "") Or (cboEmptyBox(CBO_BOXVALUE_SKIP).Text = "")) Then
    
        cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = False
    
    ElseIf ((cboSkipBoxCode.Text <> "") And (cboBoxValue(CBO_BOXVALUE_SKIP).Text <> "") And _
        (cboSkipPosition.Text <> "") And (cboEmptyBox(CBO_BOXVALUE_SKIP).Text <> "")) Then
    
        cmdRecordOperation(CMD_SKIP_SAVE).Enabled = True
        cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = True
    
    End If

    Dim strTempBoxCode As String
    
    
    strTempBoxCode = mvarActiveBoxCode
    mvarActiveBoxCode = cboSkipBoxCode.Text
        
    If (mvarActiveBoxCode <> "") Then
        LoadPicklistList cboBoxValue(CBO_BOXVALUE_SKIP)
    
        ' load length here
        Dim intBoxWidth As Integer
        Dim intBoxPosition As Integer
        
        intBoxWidth = GetBoxWidth(mvarActiveBoxCode)
    
        cboSkipPosition.Clear
        For intBoxPosition = 0 To intBoxWidth
            cboSkipPosition.AddItem CStr(intBoxPosition)
        Next intBoxPosition
    
        'cboBoxValue (CBO_BOXVALUE_SKIP).
    
    End If
        
    ' restore original settings
    mvarActiveBoxCode = strTempBoxCode


End Sub

Private Function GetBoxWidth(ByVal ActiveBoxCode As String) As Integer
'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = ActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    GetBoxWidth = clsBoxDefaultAdmin.FIELD_WIDTH
    
    Set clsBoxDefaultAdmin = Nothing
    Set clsBoxDefaultAdmins = Nothing
'
End Function

Private Sub cboSkipPosition_Click()
    
    If ((cboSkipBoxCode.Text = "") Or (cboBoxValue(CBO_BOXVALUE_SKIP).Text = "") Or _
        (cboSkipPosition.Text = "") Or (cboEmptyBox(CBO_BOXVALUE_SKIP).Text = "")) Then
    
        cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = False
    
    ElseIf ((cboSkipBoxCode.Text <> "") And (cboBoxValue(CBO_BOXVALUE_SKIP).Text <> "") And _
        (cboSkipPosition.Text <> "") And (cboEmptyBox(CBO_BOXVALUE_SKIP).Text <> "")) Then
    
        cmdRecordOperation(CMD_SKIP_SAVE).Enabled = True
        cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        cmdTransact(CMD_APPLY).Enabled = True
    
    End If


End Sub

Private Sub cmdRecordOperation_Click(Index As Integer)

    ' check if this record exist in the collection
    Dim intListviewCtr As Integer
    Dim clsListItem As ListItem
            
    Select Case Index
    
        Case CMD_TABORDER_SAVE  ' 0
            SaveTabOrder
        Case CMD_TABORDER_REMOVE    ' 1
            RemoveTabOrder
        Case CMD_TABORDER_EMPTY ' 2
            EmptyTabOrder
        
        Case CMD_SKIP_SAVE '= 3
            SaveSkipTab
        Case CMD_SKIP_REMOVE '= 4
            RemoveSkipTab
        Case CMD_SKIP_EMPTY '= 5
            EmptySkipTab
        
        Case CMD_LOGID_GENERAL_SAVE '= 8
            SaveGeneralLogID
        Case CMD_LOGID_GENERAL_REMOVE ' = 7
            RemoveGeneralLogID
        Case CMD_LOGID_GENERAL_EMPTY ' = 6
            EmptyGeneralLogID
        
        Case CMD_LOGID_USER_SAVE '= 11
            SaveUserLogID
        Case CMD_LOGID_USER_REMOVE '= 10
            RemoveUserLogID
        Case CMD_LOGID_USER_EMPTY '= 9
            EmptyUserLogID
    
    End Select

    cmdTransact(CMD_APPLY).Enabled = True

End Sub

Private Function SaveGeneralLogID() As Boolean
'
    Dim clsListItem As ListItem
    Dim intListviewCtr As Integer

    If (lvwGenDefaultValue.ListItems.Count = 0) Then
    
        Set clsListItem = lvwGenDefaultValue.ListItems.Add(, , cboLogicalID(CBO_LOGID_GENERAL).Text)
        clsListItem.SubItems(1) = cboDefaultValue(CBO_LOGID_GENERAL).Text
        clsListItem.Selected = True
        
        cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
        cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
    
    ElseIf (lvwGenDefaultValue.ListItems.Count <> 0) Then
    
    '                Set clsListItem = lvwGenDefaultValue.ListItems.Add(, , cboLogicalID(CBO_LOGID_GENERAL).Text)
    '                clsListItem.SubItems(1) = cboDefaultValue(CBO_LOGID_GENERAL).Text
    '                cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
    '                cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
        
        Dim blnItemFound As Boolean
        
        For intListviewCtr = 1 To lvwGenDefaultValue.ListItems.Count
        
            If (lvwGenDefaultValue.ListItems(intListviewCtr).Text = cboLogicalID(CBO_LOGID_GENERAL).Text) Then
                
                lvwGenDefaultValue.ListItems(intListviewCtr).SubItems(1) = cboDefaultValue(CBO_LOGID_GENERAL).Text
                lvwGenDefaultValue.ListItems(intListviewCtr).Selected = True
                blnItemFound = True
                Exit For
            
            End If
        
        Next intListviewCtr
    
        If (blnItemFound = False) Then
            '
            Set clsListItem = lvwGenDefaultValue.ListItems.Add(, , cboLogicalID(CBO_LOGID_GENERAL).Text)
            clsListItem.SubItems(1) = cboDefaultValue(CBO_LOGID_GENERAL).Text
            clsListItem.Selected = True
        End If
    
        cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
        cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
    
    
    End If
    
    cmdRecordOperation(CMD_LOGID_GENERAL_SAVE).Enabled = False
    cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = True
    cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = True
    
    On Error Resume Next
    lvwGenDefaultValue.SetFocus
    On Error GoTo 0
'
End Function

Private Function RemoveGeneralLogID() As Boolean
'
    Dim intListviewCtr As Integer
    
    For intListviewCtr = 1 To lvwGenDefaultValue.ListItems.Count
        If (lvwGenDefaultValue.ListItems(intListviewCtr).Selected = True) Then
            lvwGenDefaultValue.ListItems.Remove intListviewCtr
            Exit For
        End If
    Next intListviewCtr

    If (lvwGenDefaultValue.ListItems.Count = 0) Then
        cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
        cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
        On Error Resume Next
        cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
        cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
        cboLogicalID(CBO_LOGID_GENERAL).SetFocus
        On Error GoTo 0
    ElseIf (lvwGenDefaultValue.ListItems.Count > 0) Then
        On Error Resume Next
        cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
        cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
        cboLogicalID(CBO_LOGID_GENERAL).SetFocus
        'lvwGenDefaultValue.lis
        'lvwGenDefaultValue.SetFocus
        On Error GoTo 0
    End If


End Function

Private Function EmptyGeneralLogID() As Boolean
'
    Dim intAns As Integer
    
    'intAns = MsgBox("Are you sure you want to delete all items?", vbYesNo, "Empty List")
    intAns = MsgBox(Trim(Translate_B(405, mvarResourceHandle)), vbYesNo + vbExclamation + vbDefaultButton2, "ClearingPoint")
    If (intAns = vbYes) Then
        
        lvwGenDefaultValue.ListItems.Clear
        cboLogicalID(CBO_LOGID_GENERAL).ListIndex = 0
        cboDefaultValue(CBO_LOGID_GENERAL).Text = ""
        
        If (lvwGenDefaultValue.ListItems.Count = 0) Then
            cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = False
            cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = False
        End If

        On Error Resume Next
        cboLogicalID(CBO_LOGID_GENERAL).SetFocus
        On Error GoTo 0
    
    ElseIf (intAns = vbNo) Then
        ' do nothing
    End If

End Function

Private Function SaveUserLogID() As Boolean
'
    Dim clsListItem As ListItem
    Dim intListviewCtr As Integer

    If (lvwUserDefaultValue.ListItems.Count = 0) Then
    
        Set clsListItem = lvwUserDefaultValue.ListItems.Add(, , cboLogicalID(CBO_LOGID_USER).Text)
        clsListItem.SubItems(1) = cboDefaultValue(CBO_LOGID_USER).Text
        clsListItem.Selected = True
        
        cboLogicalID(CBO_LOGID_USER).ListIndex = 0
        cboDefaultValue(CBO_LOGID_USER).Text = ""
    
    ElseIf (lvwUserDefaultValue.ListItems.Count <> 0) Then
    
        Dim blnItemFound As Boolean
        
        For intListviewCtr = 1 To lvwUserDefaultValue.ListItems.Count
        
            If (lvwUserDefaultValue.ListItems(intListviewCtr).Text = cboLogicalID(CBO_LOGID_USER).Text) Then
                
                lvwUserDefaultValue.ListItems(intListviewCtr).SubItems(1) = cboDefaultValue(CBO_LOGID_USER).Text
                lvwUserDefaultValue.ListItems(intListviewCtr).Selected = True
                blnItemFound = True
                Exit For
            
            End If
        
        Next intListviewCtr
    
        If (blnItemFound = False) Then
            '
            Set clsListItem = lvwUserDefaultValue.ListItems.Add(, , cboLogicalID(CBO_LOGID_USER).Text)
            clsListItem.SubItems(1) = cboDefaultValue(CBO_LOGID_USER).Text
            clsListItem.Selected = True
        End If
    
        cboLogicalID(CBO_LOGID_USER).ListIndex = 0
        cboDefaultValue(CBO_LOGID_USER).Text = ""
    
    
    End If
    
    cmdRecordOperation(CMD_LOGID_USER_SAVE).Enabled = False
    cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = True
    cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = True
    
    On Error Resume Next
    lvwUserDefaultValue.SetFocus
    On Error GoTo 0
'
End Function

Private Function RemoveUserLogID() As Boolean
'
    Dim intListviewCtr As Integer
    
    For intListviewCtr = 1 To lvwUserDefaultValue.ListItems.Count
        If (lvwUserDefaultValue.ListItems(intListviewCtr).Selected = True) Then
            lvwUserDefaultValue.ListItems.Remove intListviewCtr
            Exit For
        End If
    Next intListviewCtr

    If (lvwUserDefaultValue.ListItems.Count = 0) Then
        cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
        cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
        On Error Resume Next
        cboLogicalID(CBO_LOGID_USER).ListIndex = 0
        cboDefaultValue(CBO_LOGID_USER).Text = ""
        cboLogicalID(CBO_LOGID_USER).SetFocus
        On Error GoTo 0
    ElseIf (lvwUserDefaultValue.ListItems.Count > 0) Then
        On Error Resume Next
        cboLogicalID(CBO_LOGID_USER).ListIndex = 0
        cboDefaultValue(CBO_LOGID_USER).Text = ""
        cboLogicalID(CBO_LOGID_USER).SetFocus
        'lvwUserDefaultValue.lis
        'lvwUserDefaultValue.SetFocus
        On Error GoTo 0
    End If

End Function

Private Function EmptyUserLogID() As Boolean
'
    Dim intAns As Integer
    
    'intAns = MsgBox("Are you sure you want to delete all items?", vbYesNo, "Empty List")
    intAns = MsgBox(Trim(Translate_B(405, mvarResourceHandle)), vbYesNo + vbExclamation + vbDefaultButton2, "ClearingPoint")
    
    If (intAns = vbYes) Then
        
        lvwUserDefaultValue.ListItems.Clear
        cboLogicalID(CBO_LOGID_USER).ListIndex = 0
        cboDefaultValue(CBO_LOGID_USER).Text = ""
        
        If (lvwUserDefaultValue.ListItems.Count = 0) Then
            cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = False
            cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = False
        End If

        On Error Resume Next
        cboLogicalID(CBO_LOGID_USER).SetFocus
        On Error GoTo 0
    
    ElseIf (intAns = vbNo) Then
        ' do nothing
    End If

End Function

Private Function SaveTabOrder() As Boolean
'
    Dim clsListItem As ListItem
    Dim intListviewCtr As Integer

    If (lvwTabOrder.ListItems.Count = 0) Then
    
        Set clsListItem = lvwTabOrder.ListItems.Add(, , cboBoxValue(CBO_BOXVALUE_TABORDER).Text)
        clsListItem.SubItems(1) = cboGoto.Text
        clsListItem.SubItems(2) = cboEmptyBox(CBO_BOXVALUE_TABORDER).Text
        clsListItem.Selected = True
        
        cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboGoto.ListIndex = 0
    
    ElseIf (lvwTabOrder.ListItems.Count <> 0) Then
    
        Dim blnItemFound As Boolean
        
        For intListviewCtr = 1 To lvwTabOrder.ListItems.Count
        
            If (lvwTabOrder.ListItems(intListviewCtr).Text = cboBoxValue(CBO_BOXVALUE_TABORDER).Text) Then
                
                lvwTabOrder.ListItems(intListviewCtr).SubItems(1) = cboGoto.Text
                lvwTabOrder.ListItems(intListviewCtr).SubItems(2) = cboEmptyBox(CBO_BOXVALUE_TABORDER).Text
                lvwTabOrder.ListItems(intListviewCtr).Selected = True
                blnItemFound = True
                Exit For
            
            End If
        
        Next intListviewCtr
    
        If (blnItemFound = False) Then
            '
            Set clsListItem = lvwTabOrder.ListItems.Add(, , cboBoxValue(CBO_BOXVALUE_TABORDER).Text)
            clsListItem.SubItems(1) = cboGoto.Text
            clsListItem.SubItems(2) = cboEmptyBox(CBO_BOXVALUE_TABORDER).Text
            clsListItem.Selected = True
        End If
    
        cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboGoto.ListIndex = 0
    
    End If
    
    cmdRecordOperation(CMD_TABORDER_SAVE).Enabled = False
    cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = True
    cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = True
    
    On Error Resume Next
    lvwTabOrder.SetFocus
    On Error GoTo 0
'
End Function

Private Function RemoveTabOrder() As Boolean
'
    Dim intListviewCtr As Integer
    
    For intListviewCtr = 1 To lvwTabOrder.ListItems.Count
        If (lvwTabOrder.ListItems(intListviewCtr).Selected = True) Then
            lvwTabOrder.ListItems.Remove intListviewCtr
            Exit For
        End If
    Next intListviewCtr

    If (lvwTabOrder.ListItems.Count = 0) Then
        cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
        cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
        On Error Resume Next
        cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboGoto.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_TABORDER).SetFocus
        On Error GoTo 0
    ElseIf (lvwTabOrder.ListItems.Count > 0) Then
        On Error Resume Next
        cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboGoto.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_TABORDER).SetFocus
        'lvwTabOrder.lis
        'lvwTabOrder.SetFocus
        On Error GoTo 0
    End If


End Function

Private Function EmptyTabOrder() As Boolean
'
    Dim intAns As Integer
    
    'intAns = MsgBox("Are you sure you want to delete all items?", vbYesNo, "Empty List")
    intAns = MsgBox(Trim(Translate_B(405, mvarResourceHandle)), vbYesNo + vbExclamation + vbDefaultButton2, "ClearingPoint")
    
    If (intAns = vbYes) Then
        
        lvwTabOrder.ListItems.Clear
        cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = 0
        cboGoto.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_TABORDER).ListIndex = 0
        
        If (lvwTabOrder.ListItems.Count = 0) Then
            cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = False
            cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = False
        End If

        On Error Resume Next
        cboBoxValue(CBO_BOXVALUE_TABORDER).SetFocus
        On Error GoTo 0
    
    ElseIf (intAns = vbNo) Then
        ' do nothing
    End If

End Function

Private Function SaveSkipTab() As Boolean
    '
    Dim clsListItem As ListItem
    Dim intListviewCtr As Integer

    If (lvwSkip.ListItems.Count = 0) Then
        
        Set clsListItem = lvwSkip.ListItems.Add(, , cboSkipBoxCode.Text)
        clsListItem.SubItems(1) = cboBoxValue(CBO_BOXVALUE_SKIP).Text
        clsListItem.SubItems(2) = cboSkipPosition.Text
        clsListItem.SubItems(3) = cboEmptyBox(CBO_BOXVALUE_SKIP).Text
        clsListItem.Selected = True
        
        cboSkipBoxCode.ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipPosition.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_SKIP).ListIndex = 0
    
    ElseIf (lvwSkip.ListItems.Count <> 0) Then
    
        Dim blnItemFound As Boolean
        
        For intListviewCtr = 1 To lvwSkip.ListItems.Count
        
            If (lvwSkip.ListItems(intListviewCtr).Text = cboSkipBoxCode.Text) Then
                
                lvwSkip.ListItems(intListviewCtr).SubItems(1) = cboBoxValue(CBO_BOXVALUE_SKIP).Text
                lvwSkip.ListItems(intListviewCtr).SubItems(2) = cboSkipPosition.Text
                lvwSkip.ListItems(intListviewCtr).SubItems(3) = cboEmptyBox(CBO_BOXVALUE_SKIP).Text
                lvwSkip.ListItems(intListviewCtr).Selected = True
                blnItemFound = True
                
                Exit For
            
            End If
        
        Next intListviewCtr
    
        If (blnItemFound = False) Then
            Set clsListItem = lvwSkip.ListItems.Add(, , cboSkipBoxCode.Text)
            clsListItem.SubItems(1) = cboBoxValue(CBO_BOXVALUE_SKIP).Text
            clsListItem.SubItems(2) = cboSkipPosition.Text
            clsListItem.SubItems(3) = cboEmptyBox(CBO_BOXVALUE_SKIP).Text
            clsListItem.Selected = True
        End If
    
        cboSkipBoxCode.ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipPosition.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_SKIP).ListIndex = 0
    
    End If
    
    cmdRecordOperation(CMD_SKIP_SAVE).Enabled = False
    cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = True
    cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = True
    
    On Error Resume Next
    lvwSkip.SetFocus
    On Error GoTo 0
'
End Function

Private Function RemoveSkipTab() As Boolean
'
    Dim intListviewCtr As Integer
    
    For intListviewCtr = 1 To lvwSkip.ListItems.Count
        If (lvwSkip.ListItems(intListviewCtr).Selected = True) Then
            lvwSkip.ListItems.Remove intListviewCtr
            Exit For
        End If
    Next intListviewCtr

    If (lvwSkip.ListItems.Count = 0) Then
        cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
        cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        On Error Resume Next
        cboSkipBoxCode.ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipPosition.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipBoxCode.SetFocus
        On Error GoTo 0
    ElseIf (lvwSkip.ListItems.Count > 0) Then
        On Error Resume Next
        cboSkipBoxCode.ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipPosition.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipBoxCode.SetFocus
        On Error GoTo 0
    End If

End Function

Private Function EmptySkipTab() As Boolean

    Dim intAns As Integer
    
    'intAns = MsgBox("Are you sure you want to delete all items?", vbYesNo, "Empty List")
    intAns = MsgBox(Trim(Translate_B(405, mvarResourceHandle)), vbYesNo + vbExclamation + vbDefaultButton2, "ClearingPoint")
    
    If (intAns = vbYes) Then
        
        lvwSkip.ListItems.Clear
        cboSkipBoxCode.ListIndex = 0
        cboBoxValue(CBO_BOXVALUE_SKIP).ListIndex = 0
        cboSkipPosition.ListIndex = 0
        cboEmptyBox(CBO_BOXVALUE_SKIP).ListIndex = 0
        
        If (lvwSkip.ListItems.Count = 0) Then
            cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = False
            cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = False
        End If

        On Error Resume Next
        cboSkipBoxCode.SetFocus
        On Error GoTo 0
    
    ElseIf (intAns = vbNo) Then
        ' do nothing
    End If

End Function


Private Sub cmdTransact_Click(Index As Integer)

    Select Case Index
    
        Case CMD_OK
        
            SaveValues
            Unload Me
        
        Case CMD_CANCEL
        
            Unload Me
        
        Case CMD_APPLY
        
            SaveValues
            
        
    End Select

End Sub

Private Sub Form_Activate()

    '
    lvwGenDefaultValue.Refresh

End Sub

''  ''  ''  ''  ''  ''  ''  ''  ''
Private Sub Form_Load()

    ' translate caption
    'TranslateCaptions
    LoadResStrings Me, mvarResourceHandle
    
    LoadValues

    ' load rules/restrictions
    LoadRules

'    ' translate caption
'    TranslateCaptions
'
'    ' load general tab
'    Load_General_Tab
'
'    ' load tab order tab
'    Load_TabOrder_Tab
'
'    ' load skip tab
'    Load_Skip_Tab
'
'    ' load picklist tab
'    Load_Picklist_Tab
'    '
'    ' load default values tab
'    Load_DefaultValues_Tab
    '
    tabPlatform.Tab = 0

    ' proper naming
    '
    
End Sub

' load general tab
Private Function Load_General_Tab() As Boolean
'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    ' start mapping here
    ' load check definition frame here
    ' if A the alphanumeric else if N numerical
    If (clsBoxDefaultAdmin.FIELD_DATA_TYPE = "A") Then
        optDefinition(OPT_ALPHANUMERIC).Value = True
        txtDefinition(TXT_ALPHANUMERIC).Text = clsBoxDefaultAdmin.FIELD_WIDTH
    ElseIf (clsBoxDefaultAdmin.FIELD_DATA_TYPE = "N") Then
        optDefinition(OPT_NUMERIC).Value = True
        txtDefinition(TXT_NUMERIC).Text = clsBoxDefaultAdmin.FIELD_WIDTH
        txtDefinition(TXT_NUMERIC_DECIMAL).Text = clsBoxDefaultAdmin.FIELD_DECIMAL
    End If
    
    ' load check edit type here
    If (clsBoxDefaultAdmin.FIELD_INSERT = 0) Then
        optEdit(OPT_DEFAULT).Value = True
    ElseIf (clsBoxDefaultAdmin.FIELD_INSERT = 1) Then
        optEdit(OPT_INSERT).Value = True
    ElseIf (clsBoxDefaultAdmin.FIELD_INSERT = 2) Then
        optEdit(OPT_OVERWRITE).Value = True
    End If
    
    ' load description here
    txtDescription(TXT_DESC_ENGLISH).Text = clsBoxDefaultAdmin.FIELD_ENGLISH_DESCRIPTION
    txtDescription(TXT_DESC_DUTCH).Text = clsBoxDefaultAdmin.FIELD_DUTCH_DESCRIPTION
    txtDescription(TXT_DESC_FRENCH).Text = clsBoxDefaultAdmin.FIELD_FRENCH_DESCRIPTION
                        
    ' load check action here
    chkAction(CHK_DEACTIVATE_SEQ_TAB_DEFAULT).Value = Abs(clsBoxDefaultAdmin.FIELD_SEQUENTIAL_TABBING_DEFAULT)
    chkAction(CHK_DEACTIVATE_SEQ_TAB_ACTIVE).Value = Abs(clsBoxDefaultAdmin.FIELD_SEQUENTIAL_TABBING_ACTIVE)
    chkAction(CHK_CHECK_VAT_NO).Value = Abs(clsBoxDefaultAdmin.FIELD_CHECK_VAT)
    chkAction(CHK_CALCULATE_NET_WEIGHT).Value = Abs(clsBoxDefaultAdmin.FIELD_CALCULATE_NET_WEIGHT)
    chkAction(CHK_CALCULATE_NO_OF_ITEMS).Value = Abs(clsBoxDefaultAdmin.FIELD_CALCULATE_NO_OF_ITEMS)
    chkAction(CHK_COPY_NEXT_H_AND_D).Value = Abs(clsBoxDefaultAdmin.FIELD_COPY_TO_NEXT)
    chkAction(CHK_CHANGE_WHEN_H_IS_1).Value = Abs(clsBoxDefaultAdmin.FIELD_CHANGE_ONLY_IN_H1)
    chkAction(CHK_SEND_WHEN_H_IS_1).Value = Abs(clsBoxDefaultAdmin.FIELD_SEND_ONLY_IN_H1)
    chkAction(CHK_CALCULATE_CUSTOM_VALUE).Value = Abs(clsBoxDefaultAdmin.FIELD_CALCULATE_CUSTOMS_VALUE)
    chkAction(CHK_RELATE_L1_TO_S1).Value = Abs(clsBoxDefaultAdmin.FIELD_RELATE_L1_TO_S1)
    chkAction(CHK_VALIDATE_VALUE).Value = Abs(clsBoxDefaultAdmin.FIELD_VALIDATE_VALUE)
    'CHK_VALIDATE_VALUE
    
    ' enable/disable check boxes
    EnableDisableCheckBoxes
    
    Set clsBoxDefaultAdmin = Nothing
    Set clsBoxDefaultAdmins = Nothing

End Function

Private Function Save_General_Tab() As Boolean
'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
    
'    ' start mapping here
'    ' Save check definition frame here
'    ' if A the alphanumeric else if N numerical
'    If (clsBoxDefaultAdmin.FIELD_DATA_TYPE = "A") Then
'        optDefinition(OPT_ALPHANUMERIC).Value = True
'        txtDefinition(TXT_ALPHANUMERIC).Text = clsBoxDefaultAdmin.FIELD_WIDTH
'    ElseIf (clsBoxDefaultAdmin.FIELD_DATA_TYPE = "N") Then
'        optDefinition(OPT_NUMERIC).Value = True
'        txtDefinition(TXT_NUMERIC).Text = clsBoxDefaultAdmin.FIELD_WIDTH
'        txtDefinition(TXT_NUMERIC_DECIMAL).Text = clsBoxDefaultAdmin.FIELD_DECIMAL
'    End If
    
    ' Save check edit type here
    If (optEdit(OPT_DEFAULT).Value = True) Then
        clsBoxDefaultAdmin.FIELD_INSERT = 0
    ElseIf (optEdit(OPT_INSERT).Value = True) Then
        clsBoxDefaultAdmin.FIELD_INSERT = 1
    ElseIf (optEdit(OPT_OVERWRITE).Value = True) Then
        clsBoxDefaultAdmin.FIELD_INSERT = 2
    End If
    
    ' Save description here
    clsBoxDefaultAdmin.FIELD_ENGLISH_DESCRIPTION = txtDescription(TXT_DESC_ENGLISH).Text
    clsBoxDefaultAdmin.FIELD_DUTCH_DESCRIPTION = txtDescription(TXT_DESC_DUTCH).Text
    clsBoxDefaultAdmin.FIELD_FRENCH_DESCRIPTION = txtDescription(TXT_DESC_FRENCH).Text
                        
    ' Save check action here
    clsBoxDefaultAdmin.FIELD_SEQUENTIAL_TABBING_DEFAULT = CBool(chkAction(CHK_DEACTIVATE_SEQ_TAB_DEFAULT).Value)
    clsBoxDefaultAdmin.FIELD_SEQUENTIAL_TABBING_ACTIVE = CBool(chkAction(CHK_DEACTIVATE_SEQ_TAB_ACTIVE).Value)
    clsBoxDefaultAdmin.FIELD_CHECK_VAT = CBool(chkAction(CHK_CHECK_VAT_NO).Value)
    clsBoxDefaultAdmin.FIELD_CALCULATE_NET_WEIGHT = CBool(chkAction(CHK_CALCULATE_NET_WEIGHT).Value)
    clsBoxDefaultAdmin.FIELD_CALCULATE_NO_OF_ITEMS = CBool(chkAction(CHK_CALCULATE_NO_OF_ITEMS).Value)
    clsBoxDefaultAdmin.FIELD_COPY_TO_NEXT = CBool(chkAction(CHK_COPY_NEXT_H_AND_D).Value)
    clsBoxDefaultAdmin.FIELD_CHANGE_ONLY_IN_H1 = CBool(chkAction(CHK_CHANGE_WHEN_H_IS_1).Value)
    clsBoxDefaultAdmin.FIELD_SEND_ONLY_IN_H1 = CBool(chkAction(CHK_SEND_WHEN_H_IS_1).Value)
    clsBoxDefaultAdmin.FIELD_CALCULATE_CUSTOMS_VALUE = CBool(chkAction(CHK_CALCULATE_CUSTOM_VALUE).Value)
    clsBoxDefaultAdmin.FIELD_RELATE_L1_TO_S1 = CBool(chkAction(CHK_RELATE_L1_TO_S1).Value)
    clsBoxDefaultAdmin.FIELD_VALIDATE_VALUE = CBool(chkAction(CHK_VALIDATE_VALUE).Value)
    
    clsBoxDefaultAdmins.ModifyRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    Set clsBoxDefaultAdmin = Nothing
    Set clsBoxDefaultAdmins = Nothing

End Function

' load tab order tab
Private Function Load_TabOrder_Tab() As Boolean

    'load picklist here
    LoadPicklistList cboBoxValue(CBO_BOXVALUE_TABORDER)
    'cboBoxValue(CBO_BOXVALUE_TABORDER).Clear

    ' load box list tp combo box
    LoadBoxCodeList cboGoto

    ' load cbo Action here
    LoadActionList cboEmptyBox(CBO_BOXVALUE_TABORDER)
    
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).Clear
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem ""
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem "Yes"
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem "No"

    ' load treeview list
    LoadTreeviewList_TabOrder lvwTabOrder

End Function

' Save tab order tab
Private Function Save_TabOrder_Tab() As Boolean

    'Save picklist here
    'SavePicklistList cboBoxValue(CBO_BOXVALUE_TABORDER)
    'cboBoxValue(CBO_BOXVALUE_TABORDER).Clear

    ' Save box list tp combo box
    ' SaveBoxCodeList cboGoto

    ' Save cbo Action here
    'SaveActionList cboEmptyBox(CBO_BOXVALUE_TABORDER)
    
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).Clear
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem ""
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem "Yes"
'    cboEmptyBox(CBO_BOXVALUE_TABORDER).AddItem "No"

    ' Save treeview list
    SaveTreeviewList_TabOrder lvwTabOrder
    

End Function

' load skip tab
Private Function Load_Skip_Tab() As Boolean
'
    ' learn C++
    ' load box list to combo box
    LoadBoxCodeList cboSkipBoxCode
    
    ' box value
    cboBoxValue(CBO_BOXVALUE_SKIP).Clear
    cboBoxValue(CBO_BOXVALUE_SKIP).AddItem ""
    'LoadPicklistList cboBoxValue(CBO_BOXVALUE_SKIP)
    
    ' skip position
    cboSkipPosition.Clear
    cboSkipPosition.AddItem "0"
    
    ' load cbo Action here
    LoadActionList cboEmptyBox(CBO_BOXVALUE_SKIP)
    
    ' load list view
    LoadTreeviewList_Skip lvwSkip
    
End Function


' Save skip tab
Private Function Save_Skip_Tab() As Boolean
    '
    ' learn C++
    ' Save box list to combo box
    'SaveBoxCodeList cboSkipBoxCode
    '
    ' box value
    'cboBoxValue(CBO_BOXVALUE_SKIP).Clear
    '
    ' skip position
    'cboSkipPosition.Clear
    'cboSkipPosition.AddItem "0"
    '
    ' Save cbo Action here
    'SaveActionList cboEmptyBox(CBO_BOXVALUE_SKIP)
    '
    ' Save list view
    SaveTreeviewList_Skip lvwSkip
    
End Function

' load picklist tab
Private Function Load_Picklist_Tab() As Boolean

    ' check auto add here
    LoadAutoAddValue
    
    ' load list view
    LoadTreeviewList_Picklist lvwPicklist
'
End Function

' Save picklist tab
Private Function Save_Picklist_Tab() As Boolean

    ' save auto add here
    SaveAutoAddValue '
    
    ' save list view
    SaveTreeviewList_Picklist lvwPicklist
    
'
End Function

'Save_Picklist_Tab

' load default values tab
Private Function Load_DefaultValues_Tab() As Boolean
'
    LoadEmptyBoxValues
    
    LoadGeneralDefaultValues
    
    LoadTreeviewList_DefaultValue lvwGenDefaultValue, mvarBoxDefaultValueTable

    LoadUserDefaultValues
    
    LoadTreeviewList_DefaultValue lvwUserDefaultValue, mvarDefaultUserTable


    'LoadUserDefaultValues
    '
End Function

' Save default values tab
Private Function Save_DefaultValues_Tab() As Boolean
'
    SaveEmptyBoxValues
    
    SaveTreeviewList_DefaultValue lvwGenDefaultValue, mvarBoxDefaultValueTable

    SaveTreeviewList_DefaultValue lvwUserDefaultValue, mvarDefaultUserTable
    
    '
End Function


Private Function LoadBoxCodeList(ByRef ActiveCombo As ComboBox) As Boolean
'
    Dim clsBOX_DEFAULT_ADMIN As cpiBOX_DEF_ADMIN
    Dim clsBOX_DEFAULT_ADMINs As cpiBOX_DEF_ADMINs
    Dim strSql As String
    Dim rstBoxDefaultAdmin As ADODB.Recordset
    
    Set clsBOX_DEFAULT_ADMINs = New cpiBOX_DEF_ADMINs
    Set clsBOX_DEFAULT_ADMIN = New cpiBOX_DEF_ADMIN
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [" & mvarBoxDefaultAdminTable & "]"
    
    ADORecordsetOpen strSql, mvarActiveConnection, rstBoxDefaultAdmin, adOpenKeyset, adLockOptimistic
    
    Set clsBOX_DEFAULT_ADMINs.Recordset = rstBoxDefaultAdmin
    
    ActiveCombo.Clear
    ActiveCombo.AddItem ""
    
    Do While (clsBOX_DEFAULT_ADMINs.Recordset.EOF = False)
        
        Set clsBOX_DEFAULT_ADMIN = clsBOX_DEFAULT_ADMINs.GetClassRecord(clsBOX_DEFAULT_ADMINs.Recordset)
        
        If (clsBOX_DEFAULT_ADMIN.FIELD_BOX_CODE <> mvarActiveBoxCode) Then
            ActiveCombo.AddItem clsBOX_DEFAULT_ADMIN.FIELD_BOX_CODE
        End If
        
        ' *** *** *** *** *** *** *** *** ***
        '    *** *** *** *** *** *** *** *** ***
        clsBOX_DEFAULT_ADMINs.Recordset.MoveNext
        
    Loop
    
    Set clsBOX_DEFAULT_ADMIN = Nothing
    Set clsBOX_DEFAULT_ADMINs = Nothing
'
End Function

Private Function LoadPicklistList(ByRef ActiveCombo As ComboBox) As Boolean

    Dim strInternalCode As String

    ' load box values here if any
    Dim clsPICKLIST_MAINTENANCE As cpiPICK_MAINT
    Dim clsPICKLIST_MAINTENANCEs As cpiPICK_MAINTs
        
    Set clsPICKLIST_MAINTENANCE = New cpiPICK_MAINT
    Set clsPICKLIST_MAINTENANCEs = New cpiPICK_MAINTs
    
    Dim strSql As String
    Dim rstPicklistMaintenance As ADODB.Recordset
    
    strInternalCode = GetInternalCode(mvarActiveConnection, mvarActiveDocument, mvarActiveBoxCode)
    ActiveCombo.Clear
    
    If (strInternalCode <> "") Then
    
        clsPICKLIST_MAINTENANCEs.SetSqlParameters "[PICKLIST MAINTENANCE " & mvarActiveLanguage & "]", mvarActiveLanguage
        
        strSql = "SELECT * "
        strSql = strSql & " FROM [PICKLIST MAINTENANCE " & mvarActiveLanguage & "] "
        strSql = strSql & " WHERE [INTERNAL CODE]='" & strInternalCode & "'"
            
        ADORecordsetOpen strSql, mvarActiveConnection, rstPicklistMaintenance, adOpenKeyset, adLockOptimistic
        
        Set clsPICKLIST_MAINTENANCEs.Recordset = rstPicklistMaintenance
        
        ActiveCombo.AddItem ""
        
        Do While (clsPICKLIST_MAINTENANCEs.Recordset.EOF = False)
            
            ' CBO_BOXVALUE_TABORDER
            Set clsPICKLIST_MAINTENANCE = clsPICKLIST_MAINTENANCEs.GetClassRecord(clsPICKLIST_MAINTENANCEs.Recordset)
            
            'ActiveCombo.AddItem clsPICKLIST_MAINTENANCE.FIELD_DESCRIPTION
            ActiveCombo.AddItem clsPICKLIST_MAINTENANCE.FIELD_CODE
            
            clsPICKLIST_MAINTENANCEs.Recordset.MoveNext
        Loop
        
    ElseIf (strInternalCode = "") Then

        ActiveCombo.AddItem ""

    End If
    
    Set clsPICKLIST_MAINTENANCE = Nothing
    Set clsPICKLIST_MAINTENANCEs = Nothing

'
End Function

Private Function LoadTreeviewList_TabOrder(ByRef ActiveListView As ListView) As Boolean
    '
    Dim strSql As String
    Dim clsTAB_ORDER As cpiTAB_ORDER
    Dim clsTAB_ORDERs As cpiTAB_ORDERs
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [TAB ORDER] "
    strSql = strSql & "WHERE [REFERENCE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(mvarUserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    Set clsTAB_ORDER = New cpiTAB_ORDER
    Set clsTAB_ORDERs = New cpiTAB_ORDERs
    
    Dim rstTabOrders As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstTabOrders, adOpenKeyset, adLockOptimistic
    
    Set clsTAB_ORDERs.Recordset = rstTabOrders
    
    ActiveListView.ColumnHeaders.Add 1, "k1", "Value", "2430"
    ActiveListView.ColumnHeaders.Add 2, "k2", "Goto", "870"
    ActiveListView.ColumnHeaders.Add 3, "k3", "Clear Intermediate Boxes"
    
    Dim clsListItem As ListItem
    
    ActiveListView.ListItems.Clear
    ActiveListView.View = lvwReport
    
    Do While (clsTAB_ORDERs.Recordset.EOF = False)
        
        Set clsTAB_ORDER = clsTAB_ORDERs.GetClassRecord(clsTAB_ORDERs.Recordset)
        
         Set clsListItem = ActiveListView.ListItems.Add(, , clsTAB_ORDER.FIELD_VALUE)
        ' get values here
        'clsListItem.SubItems(1) = clsTAB_ORDER.FIELD_VALUE
        ' get goto here
        clsListItem.SubItems(1) = clsTAB_ORDER.FIELD_BOX_CODE
        ' get clear intermediate box here
        If (clsTAB_ORDER.FIELD_EMPTY = True) Then
            clsListItem.SubItems(2) = "Yes"
        ElseIf (clsTAB_ORDER.FIELD_EMPTY = False) Then
            clsListItem.SubItems(2) = "No"
        End If
        
        'clsListItem.SubItems(2) = clsTAB_ORDER.FIELD_EMPTY
        clsTAB_ORDERs.Recordset.MoveNext
        
    Loop
    
    Set clsTAB_ORDER = Nothing
    Set clsTAB_ORDERs = Nothing
    '
End Function

' save treeview list tab order
Private Function SaveTreeviewList_TabOrder(ByRef ActiveListView As ListView) As Boolean
'
    Dim strSql As String
    Dim clsTAB_ORDER As cpiTAB_ORDER
    Dim clsTAB_ORDERs As cpiTAB_ORDERs
    
    strSql = "DELETE * "
    strSql = strSql & "FROM [TAB ORDER] "
    strSql = strSql & "WHERE [REFERENCE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & mvarUserNo & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    ' delete old records first
    ExecuteNonQuery mvarActiveConnection, strSql
    'mvarActiveConnection.Execute strSql
    
    Set clsTAB_ORDER = New cpiTAB_ORDER
    Set clsTAB_ORDERs = New cpiTAB_ORDERs
    
    Dim intTabOrderCtr As Integer
    Dim clsListItem As ListItem
    
    For intTabOrderCtr = 1 To ActiveListView.ListItems.Count
    
        clsTAB_ORDER.FIELD_REFERENCE = mvarActiveBoxCode
        clsTAB_ORDER.FIELD_user_no = mvarUserNo
        clsTAB_ORDER.FIELD_TYPE = mvarActiveType
        
        Set clsListItem = ActiveListView.ListItems(intTabOrderCtr)
        
        clsTAB_ORDER.FIELD_VALUE = clsListItem.Text
        clsTAB_ORDER.FIELD_BOX_CODE = clsListItem.SubItems(1)
        
        If (UCase$(clsListItem.SubItems(2)) = "YES") Then
            clsTAB_ORDER.FIELD_EMPTY = True
        ElseIf (UCase$(clsListItem.SubItems(2)) = "NO") Then
            clsTAB_ORDER.FIELD_EMPTY = False
        End If
    
        clsTAB_ORDERs.AddRecord mvarActiveConnection, clsTAB_ORDER
    
    Next intTabOrderCtr
    
    Set clsTAB_ORDER = New cpiTAB_ORDER
    Set clsTAB_ORDERs = New cpiTAB_ORDERs

    ' add records here
    
    
    Set clsTAB_ORDER = Nothing
    Set clsTAB_ORDERs = Nothing
'
End Function


Private Function LoadTreeviewList_Skip(ByRef ActiveListView As ListView) As Boolean
    '
    Dim strSql As String
    Dim clsSKIP As cpiSKIP
    Dim clsSKIPs  As cpiSKIPs
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [SKIP] "
    strSql = strSql & "WHERE [REFERENCE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(mvarUserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    '
    Set clsSKIP = New cpiSKIP
    Set clsSKIPs = New cpiSKIPs
    '
    
    Dim rstSKIPS As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstSKIPS, adOpenKeyset, adLockOptimistic
    
    Set clsSKIPs.Recordset = rstSKIPS
    
    ActiveListView.ColumnHeaders.Add 1, "k1", "Box", 800
    ActiveListView.ColumnHeaders.Add 2, "k2", "Value", 2740
    ActiveListView.ColumnHeaders.Add 3, "k3", "Position", 800
    ActiveListView.ColumnHeaders.Add 4, "k4", "Empty Box"
    
    Dim clsListItem As ListItem
    
    ActiveListView.ListItems.Clear
    ActiveListView.View = lvwReport
    
    Do While (clsSKIPs.Recordset.EOF = False)
        
        Set clsSKIP = clsSKIPs.GetClassRecord(clsSKIPs.Recordset)
        
        Set clsListItem = ActiveListView.ListItems.Add(, , clsSKIP.FIELD_BOX_CODE)
        ' get box code here
        'clsListItem.SubItems(1) = clsSKIP.FIELD_BOX_CODE
        ' get value here
        clsListItem.SubItems(1) = clsSKIP.FIELD_VALUE
        ' get position here
        clsListItem.SubItems(2) = clsSKIP.FIELD_POSITION
        ' get empty box here
        'clsListItem.SubItems(3) = clsSKIP.FIELD_EMPTY
        If (clsSKIP.FIELD_EMPTY = True) Then
            clsListItem.SubItems(3) = "Yes"
        ElseIf (clsSKIP.FIELD_EMPTY = False) Then
            clsListItem.SubItems(3) = "No"
        End If
        
        clsSKIPs.Recordset.MoveNext
        
    Loop
    
    Set clsSKIP = Nothing
    Set clsSKIPs = Nothing
    '
    
End Function

Private Function SaveTreeviewList_Skip(ByRef ActiveListView As ListView) As Boolean
'
    Dim strSql As String
    Dim clsSKIP As cpiSKIP
    Dim clsSKIPs  As cpiSKIPs
    
    Set clsSKIP = New cpiSKIP
    Set clsSKIPs = New cpiSKIPs
    '
    strSql = "DELETE * "
    strSql = strSql & "FROM [SKIP] "
    strSql = strSql & "WHERE [REFERENCE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(mvarUserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    ExecuteNonQuery mvarActiveConnection, strSql
    'mvarActiveConnection.Execute strSql
    
    Dim intSkipCtr As Integer
    Dim clsListItem As ListItem
    
    For intSkipCtr = 1 To ActiveListView.ListItems.Count
    
        clsSKIP.FIELD_REFERENCE = mvarActiveBoxCode
        clsSKIP.FIELD_user_no = mvarUserNo
        clsSKIP.FIELD_TYPE = mvarActiveType
        
        Set clsListItem = ActiveListView.ListItems(intSkipCtr)
        
        clsSKIP.FIELD_BOX_CODE = clsListItem.Text
        clsSKIP.FIELD_VALUE = clsListItem.SubItems(1)
        clsSKIP.FIELD_POSITION = clsListItem.SubItems(2)
        
        If (UCase$(clsListItem.SubItems(3)) = "YES") Then
            clsSKIP.FIELD_EMPTY = True
        ElseIf (UCase$(clsListItem.SubItems(3)) = "NO") Then
            clsSKIP.FIELD_EMPTY = False
        End If
    
        clsSKIPs.AddRecord mvarActiveConnection, clsSKIP
    
    Next intSkipCtr
    
    Set clsSKIP = Nothing
    Set clsSKIPs = Nothing
'
End Function


' LoadTreeviewList_Picklist lvwPicklist
Private Function LoadTreeviewList_Picklist(ByRef ActiveListView As ListView) As Boolean
    '
    Dim strSql As String 'PICKLIST_DEFINITION
    Dim clsPICKLIST_DEFINITION As cpiPICK_DEF 'cpiPICK_DEF
    Dim clsPICKLIST_DEFINITIONs  As cpiPICK_DEFs
    '
    strSql = "SELECT * "
    strSql = strSql & " FROM [PICKLIST DEFINITION] "
    strSql = strSql & " WHERE [DOCUMENT]='" & mvarActiveDocument & "'"
    strSql = strSql & " AND [BOX CODE]='" & mvarActiveBoxCode & "'"
    
    Set clsPICKLIST_DEFINITION = New cpiPICK_DEF
    Set clsPICKLIST_DEFINITIONs = New cpiPICK_DEFs
'
    
    Dim rstPicklistDefinitions As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstPicklistDefinitions, adOpenKeyset, adLockOptimistic
    
    Set clsPICKLIST_DEFINITIONs.Recordset = rstPicklistDefinitions
    
    ActiveListView.ColumnHeaders.Clear
    ActiveListView.ColumnHeaders.Add 1, "k1", "Descripton", 3000
    ActiveListView.ColumnHeaders.Add 2, "k2", "From"
    ActiveListView.ColumnHeaders.Add 3, "k3", "Validation"
    
    Dim clsListItem As ListItem
    
    ActiveListView.ListItems.Clear
    ActiveListView.View = lvwReport
    
    Do While (clsPICKLIST_DEFINITIONs.Recordset.EOF = False)
        
        Set clsPICKLIST_DEFINITION = clsPICKLIST_DEFINITIONs.GetClassRecord(clsPICKLIST_DEFINITIONs.Recordset)
        
        Select Case mvarActiveLanguage
            Case "ENGLISH"
                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH)
                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH
            Case "DUTCH"
                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_DUTCH)
                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_DUTCH
            Case "FRENCH"
                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_FRENCH)
                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_FRENCH
        End Select
        
        ' get box code here '
        'clsListItem.SubItems(1) = clsPICKLIST_DEFINITION.FIELD_BOX_CODE
        'clsListItem.Tag = clsPICKLIST_DEFINITION.FIELD_AUTO_ADD
        ' get picklist decription here'
        'clsListItem.SubItems(1) = clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH
        ' get position here '
        clsListItem.SubItems(1) = clsPICKLIST_DEFINITION.FIELD_FROM
        ' get empty box here '
        clsListItem.SubItems(2) = clsPICKLIST_DEFINITION.FIELD_VALIDS
        '
        clsPICKLIST_DEFINITIONs.Recordset.MoveNext
        
    Loop
    
    If (ActiveListView.ListItems.Count > 0) Then
        ActiveListView.ListItems(1).Selected = True
    ElseIf (ActiveListView.ListItems.Count = 0) Then
        chkAutoAdd.Enabled = False
    End If
    
    Set clsPICKLIST_DEFINITION = Nothing
    Set clsPICKLIST_DEFINITIONs = Nothing
    
End Function

' SaveTreeviewList_Picklist lvwPicklist
Private Function SaveTreeviewList_Picklist(ByRef ActiveListView As ListView) As Boolean
'    '
'    Dim strSql As String        ' PICKLIST_DEFINITION
'    Dim clsPICKLIST_DEFINITION As cpiPICK_DEF            ' cpiPICK_DEF
'    Dim clsPICKLIST_DEFINITIONs  As cpiPICK_DEFs
'    '
'    strSql = "SELECT * "
'    strSql = strSql & " FROM [PICKLIST DEFINITION] "
'    strSql = strSql & " WHERE [DOCUMENT]='" & mvarActiveDocument & "'"
'    strSql = strSql & " AND [BOX CODE]='" & mvarActiveBoxCode & "'"
'
'    Set clsPICKLIST_DEFINITION = New cpiPICK_DEF
'    Set clsPICKLIST_DEFINITIONs = New cpiPICK_DEFs
'    '
'    Set clsPICKLIST_DEFINITIONs.Recordset = mvarActiveConnection.Execute(strSql)
'
'    ActiveListView.ColumnHeaders.Add 1, "k1", "Descripton"
'    ActiveListView.ColumnHeaders.Add 2, "k2", "From"
'    ActiveListView.ColumnHeaders.Add 3, "k3", "Validation"
'
'    Dim clsListItem As ListItem
'
'    ActiveListView.ListItems.Clear
'    ActiveListView.View = lvwReport
'
'    Do While (clsPICKLIST_DEFINITIONs.Recordset.EOF = False)
'
'        Set clsPICKLIST_DEFINITION = clsPICKLIST_DEFINITIONs.GetClassRecord(clsPICKLIST_DEFINITIONs.Recordset)
'
'        Select Case mvarActiveLanguage
'            Case "ENGLISH"
'                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH)
'                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH
'            Case "DUTCH"
'                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_DUTCH)
'                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_DUTCH
'            Case "FRENCH"
'                Set clsListItem = ActiveListView.ListItems.Add(, , clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_FRENCH)
'                'ActiveCombo.AddItem clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_FRENCH
'        End Select
'
'        ' get box code here '
'        'clsListItem.SubItems(1) = clsPICKLIST_DEFINITION.FIELD_BOX_CODE
'        'clsListItem.Tag = clsPICKLIST_DEFINITION.FIELD_AUTO_ADD
'        ' get picklist decription here'
'        'clsListItem.SubItems(1) = clsPICKLIST_DEFINITION.FIELD_PICKLIST_DESCRIPTION_ENGLISH
'        ' get position here '
'        clsListItem.SubItems(2) = clsPICKLIST_DEFINITION.FIELD_FROM
'        ' get empty box here '
'        clsListItem.SubItems(3) = clsPICKLIST_DEFINITION.FIELD_VALIDS
'        '
'        clsPICKLIST_DEFINITIONs.Recordset.MoveNext
'
'    Loop
'
'    If (ActiveListView.ListItems.Count > 0) Then
'        ActiveListView.ListItems(1).Selected = True
'    End If
'
'    Set clsPICKLIST_DEFINITION = Nothing
'    Set clsPICKLIST_DEFINITIONs = Nothing
'
End Function

Private Function LoadTreeviewList_DefaultValue(ByRef ActiveListView As ListView, ByRef ActiveTable As String) As Boolean
'
    ' as cpiBOX_DEF_VAL
    Dim strSql As String
    Dim clsBOX_DEFAULT_VALUE As cpiBOX_DEF_VAL
    Dim clsBOX_DEFAULT_VALUEs As cpiBOX_DEF_VALs
    
    ' @@@
    strSql = "SELECT * "
    strSql = strSql & "FROM [" & ActiveTable & "]"
    strSql = strSql & "WHERE [BOX CODE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(mvarUserNo) & ""
    
    Set clsBOX_DEFAULT_VALUE = New cpiBOX_DEF_VAL
    Set clsBOX_DEFAULT_VALUEs = New cpiBOX_DEF_VALs
'
    Dim rstBoxDefaultValues As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstBoxDefaultValues, adOpenKeyset, adLockOptimistic
    
    Set clsBOX_DEFAULT_VALUEs.Recordset = rstBoxDefaultValues
    
    ActiveListView.ColumnHeaders.Add 1, "k1", "Access Code", "2500"
    ActiveListView.ColumnHeaders.Add 2, "k2", "Value", "2500"
    
    Dim clsListItem As ListItem
    
    ActiveListView.ListItems.Clear
    ActiveListView.View = lvwReport
    
    Do While (clsBOX_DEFAULT_VALUEs.Recordset.EOF = False)
        
        Set clsBOX_DEFAULT_VALUE = clsBOX_DEFAULT_VALUEs.GetClassRecord(clsBOX_DEFAULT_VALUEs.Recordset)
        
         Set clsListItem = ActiveListView.ListItems.Add(, , clsBOX_DEFAULT_VALUE.FIELD_LOGID_DESCRIPTION)
        ' get goto here
        'clsListItem.Width = 5000
        clsListItem.SubItems(1) = clsBOX_DEFAULT_VALUE.FIELD_DEFAULT_VALUE
        '
        clsBOX_DEFAULT_VALUEs.Recordset.MoveNext
        
    Loop
    
    Set clsBOX_DEFAULT_VALUE = Nothing
    Set clsBOX_DEFAULT_VALUEs = Nothing
'
End Function

Private Function SaveTreeviewList_DefaultValue(ByRef ActiveListView As ListView, ByRef ActiveTable As String) As Boolean
    
    ' as cpiBOX_DEF_VAL
    Dim strSql As String
    Dim clsBOX_DEFAULT_VALUE As cpiBOX_DEF_VAL
    Dim clsBOX_DEFAULT_VALUEs As cpiBOX_DEF_VALs
    
    strSql = "DELETE * "
    strSql = strSql & "FROM [" & ActiveTable & "]"
    strSql = strSql & "WHERE [BOX CODE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(mvarUserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    ExecuteNonQuery mvarActiveConnection, strSql
    'mvarActiveConnection.Execute strSql
    
    Set clsBOX_DEFAULT_VALUE = New cpiBOX_DEF_VAL
    Set clsBOX_DEFAULT_VALUEs = New cpiBOX_DEF_VALs
    
    clsBOX_DEFAULT_VALUEs.SetSqlParameters ActiveTable
    
    Dim intBoxDefaultValueCtr As Integer
    Dim clsListItem As ListItem
    
    For intBoxDefaultValueCtr = 1 To ActiveListView.ListItems.Count
    
        clsBOX_DEFAULT_VALUE.FIELD_BOX_CODE = mvarActiveBoxCode
        clsBOX_DEFAULT_VALUE.FIELD_user_no = mvarUserNo
        clsBOX_DEFAULT_VALUE.FIELD_TYPE = mvarActiveType
        
        Set clsListItem = ActiveListView.ListItems(intBoxDefaultValueCtr)
        
        clsBOX_DEFAULT_VALUE.FIELD_LOGID_DESCRIPTION = clsListItem.Text
        clsBOX_DEFAULT_VALUE.FIELD_DEFAULT_VALUE = clsListItem.SubItems(1)
    
        clsBOX_DEFAULT_VALUEs.AddRecord mvarActiveConnection, clsBOX_DEFAULT_VALUE
    
    Next intBoxDefaultValueCtr
    '
    Set clsBOX_DEFAULT_VALUE = Nothing
    Set clsBOX_DEFAULT_VALUEs = Nothing
'
End Function


Private Function LoadTreeviewList_UserValue(ByRef ActiveListView As ListView) As Boolean
'
     'as cpiDEFAULT_USER
    Dim strSql As String
    Dim clsDEFAULT_USER As cpiDEFAULT_USER
    Dim clsDEFAULT_USERs As cpiDEFAULT_USERs
    
    ' @@@
    strSql = "SELECT * "
    strSql = strSql & "FROM [" & mvarDefaultUserTable & "]"
    strSql = strSql & "WHERE [BOX CODE]='" & mvarActiveBoxCode & "'"
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    Set clsDEFAULT_USER = New cpiDEFAULT_USER
    Set clsDEFAULT_USERs = New cpiDEFAULT_USERs
'
    Dim rstDefaultUsers As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstDefaultUsers, adOpenKeyset, adLockOptimistic
    
    Set clsDEFAULT_USERs.Recordset = rstDefaultUsers
    
    ActiveListView.ColumnHeaders.Add 1, "1", "Access Code"
    ActiveListView.ColumnHeaders.Add 2, "2", "Value"
    
    Dim clsListItem As ListItem
    
    ActiveListView.ListItems.Clear
    ActiveListView.View = lvwReport
    
    Do While (clsDEFAULT_USERs.Recordset.EOF = False)
        
        Set clsDEFAULT_USER = clsDEFAULT_USERs.GetClassRecord(clsDEFAULT_USERs.Recordset)
        
         Set clsListItem = ActiveListView.ListItems.Add(, , clsDEFAULT_USER.FIELD_LOGID_DESCRIPTION)
        ' get goto here
        clsListItem.SubItems(1) = clsDEFAULT_USER.FIELD_DEFAULT_VALUE
        '
        clsDEFAULT_USERs.Recordset.MoveNext
        
    Loop
    
    Set clsDEFAULT_USER = Nothing
    Set clsDEFAULT_USERs = Nothing
'
End Function

Private Function LoadActionList(ByRef ActiveComboBox As ComboBox) As Boolean

    ActiveComboBox.Clear
    ActiveComboBox.AddItem ""
    ActiveComboBox.AddItem "Yes"
    ActiveComboBox.AddItem "No"

End Function
    
Private Function LoadEmptyBoxValues() As Boolean

'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    Dim strInternalCode As String
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
        
    ' check if box is date
    strInternalCode = GetInternalCode(mvarActiveConnection, mvarActiveDocument, mvarActiveBoxCode)
    
    If (strInternalCode <> PCK_DATE) Then
        txtEmptyField.Visible = True
        cboEmptyField.Visible = False
        txtEmptyField.Text = clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE
    ElseIf (strInternalCode = PCK_DATE) Then
        
        txtEmptyField.Visible = False
        cboEmptyField.Visible = True
        cboEmptyField.Clear
        cboEmptyField.AddItem ""
        cboEmptyField.AddItem "Yesterday"
        cboEmptyField.AddItem "Today"
        cboEmptyField.AddItem "Tomorrow"
        
        ' For Date fields A-Yesterday, B-Today, C-Tomorrow
        Select Case UCase$(clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE)
            Case "A"
                cboEmptyField.ListIndex = 1
            Case "B"
                cboEmptyField.ListIndex = 2
            Case "C"
                cboEmptyField.ListIndex = 3
        End Select
        
    End If
    
    Set clsBoxDefaultAdmins = Nothing
    Set clsBoxDefaultAdmin = Nothing
    '
End Function

Private Function LoadAutoAddValue() As Boolean
'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    Dim strInternalCode As String
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
        
    chkAutoAdd.Value = Abs(clsBoxDefaultAdmin.FIELD_AUTO_ADD)
    
    Set clsBoxDefaultAdmins = Nothing
    Set clsBoxDefaultAdmin = Nothing
    '
End Function

Private Function SaveAutoAddValue() As Boolean
'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    Dim strInternalCode As String
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
        
    clsBoxDefaultAdmin.FIELD_AUTO_ADD = CBool(chkAutoAdd.Value)
    
    clsBoxDefaultAdmins.ModifyRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    Set clsBoxDefaultAdmins = Nothing
    Set clsBoxDefaultAdmin = Nothing
    '
End Function

Private Function SaveEmptyBoxValues() As Boolean

'
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = mvarActiveBoxCode
    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    Dim strInternalCode As String
    
    ' check if box is date
    strInternalCode = GetInternalCode(mvarActiveConnection, mvarActiveDocument, mvarActiveBoxCode)
    
    If (strInternalCode <> PCK_DATE) Then
        clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE = txtEmptyField.Text
    ElseIf (strInternalCode = PCK_DATE) Then
        
        Select Case UCase$(cboEmptyField.Text)
            
            Case "YESTERDAY"
                clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE = "A"
            Case "TODAY"
                clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE = "B"
            Case "TOMORROW"
                clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE = "C"
            Case Else
                clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE = txtEmptyField.Text
        End Select
        
    End If
    
    clsBoxDefaultAdmins.ModifyRecord mvarActiveConnection, clsBoxDefaultAdmin
    
'    clsBoxDefaultAdmins.GetRecord mvarActiveConnection, clsBoxDefaultAdmin
    
    Set clsBoxDefaultAdmins = Nothing
    Set clsBoxDefaultAdmin = Nothing
'
End Function

Private Function LoadGeneralDefaultValues() As Boolean
'
    ' load logical ids
    LoadLogicalIDList cboLogicalID(CBO_LOGID_GENERAL)
    
    ' load picklist
    LoadPicklistList cboDefaultValue(CBO_LOGID_GENERAL)
    
End Function

Private Function LoadUserDefaultValues() As Boolean
'
    ' load logical ids
    LoadLogicalIDList cboLogicalID(CBO_LOGID_USER)
    
    ' load picklist
    LoadPicklistList cboDefaultValue(CBO_LOGID_USER)
    
End Function

Private Function LoadLogicalIDList(ByRef ActiveComboBox As ComboBox) As Boolean
'
    Dim clsLOGICAL_ID As cpiLOGICAL_ID
    Dim clsLOGICAL_IDs As cpiLOGICAL_IDs
    Dim strSql As String
    
    'Set clsLOGICAL_ID = New cpiLOGICAL_IDs
    Set clsLOGICAL_IDs = New cpiLOGICAL_IDs
    
    strSql = "SELECT * FROM [LOGICAL ID]"
    
    Dim rstLogicalID As ADODB.Recordset
    ADORecordsetOpen strSql, mvarActiveConnection, rstLogicalID, adOpenKeyset, adLockOptimistic
    
    Set clsLOGICAL_IDs.Recordset = rstLogicalID
    
    ActiveComboBox.Clear
    ActiveComboBox.AddItem ""
    
    Do While (clsLOGICAL_IDs.Recordset.EOF = False)
    
        Set clsLOGICAL_ID = clsLOGICAL_IDs.GetClassRecord(clsLOGICAL_IDs.Recordset)
    
        ' save values here
        ActiveComboBox.AddItem clsLOGICAL_ID.FIELD_LOGID_DESCRIPTION
        clsLOGICAL_IDs.Recordset.MoveNext
    
    Loop
        
    Set clsLOGICAL_ID = Nothing
    Set clsLOGICAL_IDs = Nothing
    
End Function

Private Function TranslateCaptions() As Boolean
    '
    Caption = " " & mvarActiveBoxCode & " - " & Translate_B(Me.Caption, mvarResourceHandle)
    tabPlatform.TabCaption(0) = Translate_B(tabPlatform.TabCaption(0), mvarResourceHandle)
    tabPlatform.TabCaption(1) = Translate_B(tabPlatform.TabCaption(1), mvarResourceHandle)
    tabPlatform.TabCaption(2) = Translate_B(tabPlatform.TabCaption(2), mvarResourceHandle)
    tabPlatform.TabCaption(3) = Translate_B(tabPlatform.TabCaption(3), mvarResourceHandle)
    tabPlatform.TabCaption(4) = Translate_B(tabPlatform.TabCaption(4), mvarResourceHandle)
    '
End Function

Private Function EnableDisableCheckBoxes() As Boolean
    '
    Select Case mvarCodisheetType
        Case cpiImportCodisheet
        
        Case cpiExportCodisheet
        
        Case cpiTransitCodisheet

        Case cpiSadbelNCTSCodisheet
        
        Case cpiCombinedNCTSCodisheet
        
        Case cpiDepartureIE15Codisheet
        
                    
        
        Case cpiArrivalIE07Codisheet
        
        Case cpiArrivalIE44Codisheet
        
    End Select

'
End Function

Private Function LoadValues() As Boolean
'
    ' load general tab
    Load_General_Tab
    
    ' load tab order tab
    Load_TabOrder_Tab
    
    ' load skip tab
    Load_Skip_Tab
    
    ' load picklist tab
    Load_Picklist_Tab
    '
    ' load default values tab
    Load_DefaultValues_Tab

'
End Function

Private Function SaveValues() As Boolean
'
    ' save general tab
    Save_General_Tab
    
    ' save tab order tab
    Save_TabOrder_Tab
    
    ' Save skip tab
    Save_Skip_Tab
    
    ' Save picklist tab
    Save_Picklist_Tab
    '
    ' Save default values tab
    Save_DefaultValues_Tab

    cmdTransact(CMD_APPLY).Enabled = False

End Function

Private Function LoadRules() As Boolean
'
    ' enable / disable controls
    'txtEmptyField.Enabled = True

     'chkAutoAdd.Enabled = False
    
    Rule_General_Tab
    
'
End Function

Private Function Rule_General_Tab() As Boolean

    Rule_General_Tab_Action

End Function

Private Function Rule_General_Tab_Action() As Boolean
'
    Select Case mvarCodisheetType
    
        Case cpiImportCodisheet
        
        Case cpiExportCodisheet
        
        Case cpiTransitCodisheet

        Case cpiSadbelNCTSCodisheet
        
        Case cpiCombinedNCTSCodisheet
        
        Case cpiDepartureIE15Codisheet
        
        Case cpiArrivalIE07Codisheet
            
            chkAction(CHK_DEACTIVATE_SEQ_TAB_DEFAULT).Enabled = True
            chkAction(CHK_DEACTIVATE_SEQ_TAB_ACTIVE).Enabled = True
            chkAction(CHK_CHECK_VAT_NO).Enabled = False
            chkAction(CHK_CALCULATE_NET_WEIGHT).Enabled = False
            chkAction(CHK_CALCULATE_NO_OF_ITEMS).Enabled = False
            chkAction(CHK_CALCULATE_CUSTOM_VALUE).Enabled = False
            chkAction(CHK_COPY_NEXT_H_AND_D).Enabled = True
            chkAction(CHK_CHANGE_WHEN_H_IS_1).Enabled = False
            chkAction(CHK_SEND_WHEN_H_IS_1).Enabled = False
            chkAction(CHK_RELATE_L1_TO_S1).Enabled = False
            chkAction(CHK_VALIDATE_VALUE).Enabled = True
                         
                
        Case cpiArrivalIE44Codisheet
        
            chkAction(CHK_DEACTIVATE_SEQ_TAB_DEFAULT).Enabled = True
            chkAction(CHK_DEACTIVATE_SEQ_TAB_ACTIVE).Enabled = True
            chkAction(CHK_CHECK_VAT_NO).Enabled = False
            chkAction(CHK_CALCULATE_NET_WEIGHT).Enabled = False
            chkAction(CHK_CALCULATE_NO_OF_ITEMS).Enabled = False
            chkAction(CHK_CALCULATE_CUSTOM_VALUE).Enabled = False
            chkAction(CHK_COPY_NEXT_H_AND_D).Enabled = True
            chkAction(CHK_CHANGE_WHEN_H_IS_1).Enabled = False
            chkAction(CHK_SEND_WHEN_H_IS_1).Enabled = False
            chkAction(CHK_RELATE_L1_TO_S1).Enabled = False
            chkAction(CHK_VALIDATE_VALUE).Enabled = True
                         
                         
        
        
    End Select
    

'
End Function


Private Sub lvwGenDefaultValue_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim intListCtr As Integer
    
    For intListCtr = 1 To cboLogicalID(CBO_LOGID_GENERAL).ListCount
        If (cboLogicalID(CBO_LOGID_GENERAL).List(intListCtr) = Item.Text) Then
            cboLogicalID(CBO_LOGID_GENERAL).ListIndex = intListCtr
            Exit For
        End If
    Next intListCtr
    
    cboDefaultValue(CBO_LOGID_GENERAL).Text = Item.SubItems(1)

    cmdRecordOperation(CMD_LOGID_GENERAL_REMOVE).Enabled = True
    cmdRecordOperation(CMD_LOGID_GENERAL_EMPTY).Enabled = True

End Sub

Private Sub lvwPicklist_ItemClick(ByVal Item As MSComctlLib.ListItem)

    chkAutoAdd.Value = CBool(Val(Item.Tag))

End Sub

Private Sub lvwSkip_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim intListCtr As Integer
    
    For intListCtr = 1 To cboSkipBoxCode.ListCount
        If (cboSkipBoxCode.List(intListCtr) = Item.Text) Then
            cboSkipBoxCode.ListIndex = intListCtr
            Exit For
        End If
    Next intListCtr
    
    cboBoxValue(CBO_BOXVALUE_SKIP).Text = Item.SubItems(1)
    cboSkipPosition.Text = Item.SubItems(2)
    cboEmptyBox(CBO_BOXVALUE_SKIP).Text = Item.SubItems(3)

    cmdRecordOperation(CMD_SKIP_REMOVE).Enabled = True
    cmdRecordOperation(CMD_SKIP_EMPTY).Enabled = True

End Sub

Private Sub lvwTabOrder_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim intListCtr As Integer
    
    '    For intListCtr = 1 To cboBoxValue(CBO_BOXVALUE_TABORDER).ListCount
    '        If (cboBoxValue(CBO_BOXVALUE_TABORDER).List(intListCtr) = Item.Text) Then
    '            cboBoxValue(CBO_BOXVALUE_TABORDER).ListIndex = intListCtr
    '            Exit For
    '        End If
    '    Next intListCtr
    
    cboBoxValue(CBO_BOXVALUE_TABORDER).Text = Item.Text
    cboGoto.Text = Item.SubItems(1)
    cboEmptyBox(CBO_BOXVALUE_TABORDER).Text = Item.SubItems(2)

    cmdRecordOperation(CMD_TABORDER_REMOVE).Enabled = True
    cmdRecordOperation(CMD_TABORDER_EMPTY).Enabled = True

End Sub

Private Sub lvwUserDefaultValue_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim intListCtr As Integer
    
    For intListCtr = 1 To cboLogicalID(CBO_LOGID_USER).ListCount
        If (cboLogicalID(CBO_LOGID_USER).List(intListCtr) = Item.Text) Then
            cboLogicalID(CBO_LOGID_USER).ListIndex = intListCtr
            Exit For
        End If
    Next intListCtr
    
    cboDefaultValue(CBO_LOGID_USER).Text = Item.SubItems(1)

    cmdRecordOperation(CMD_LOGID_USER_REMOVE).Enabled = True
    cmdRecordOperation(CMD_LOGID_USER_EMPTY).Enabled = True

End Sub

Private Sub optEdit_Click(Index As Integer)
'
    '
    '
'
End Sub
