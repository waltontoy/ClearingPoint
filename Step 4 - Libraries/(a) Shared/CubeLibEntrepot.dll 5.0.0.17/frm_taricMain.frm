VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_taricmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TARIC - Add/Modify Codes"
   ClientHeight    =   8265
   ClientLeft      =   1380
   ClientTop       =   2265
   ClientWidth     =   11910
   Icon            =   "frm_taricMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "857"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   7680
   End
   Begin VB.Frame frmGeneral 
      Caption         =   "General"
      Height          =   1695
      Left            =   120
      TabIndex        =   81
      Tag             =   "269"
      Top             =   120
      Width           =   11655
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdGeneral 
         Caption         =   "&CN-codes..."
         Height          =   300
         Index           =   0
         Left            =   3480
         TabIndex        =   91
         Tag             =   "819"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGeneral 
         Caption         =   "&Kluwer..."
         Height          =   300
         Index           =   1
         Left            =   4920
         TabIndex        =   90
         Tag             =   "830"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGeneral 
         Caption         =   "C&lients..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   6360
         TabIndex        =   89
         Tag             =   "818"
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   1200
         ScaleHeight     =   795
         ScaleWidth      =   10275
         TabIndex        =   82
         Top             =   720
         Width           =   10335
         Begin VB.TextBox txtDutchKey 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            MaxLength       =   20
            TabIndex        =   83
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtDutchDesc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            MaxLength       =   78
            TabIndex        =   84
            Top             =   240
            Width           =   7395
         End
         Begin VB.TextBox txtFrnchKey 
            Appearance      =   0  'Flat
            Height          =   290
            Left            =   0
            MaxLength       =   20
            TabIndex        =   85
            Top             =   520
            Width           =   2895
         End
         Begin VB.TextBox txtFrnchDesc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            MaxLength       =   78
            TabIndex        =   86
            Top             =   520
            Width           =   7395
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Keyword"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   88
            Tag             =   "829"
            Top             =   0
            Width           =   2895
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Description"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   87
            Tag             =   "292"
            Top             =   0
            Width           =   7395
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   94
         Tag             =   "439"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Dutch"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   93
         Tag             =   "751"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "French"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Tag             =   "752"
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame frmSettings 
      Caption         =   "Country Settings"
      Enabled         =   0   'False
      Height          =   4095
      Left            =   120
      TabIndex        =   32
      Tag             =   "823"
      Top             =   3480
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame7 
         Caption         =   "Usage"
         Height          =   855
         Index           =   3
         Left            =   8640
         TabIndex        =   97
         Tag             =   "864"
         Top             =   3000
         Width           =   2895
         Begin VB.CheckBox chkDefault 
            Caption         =   "Default"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Tag             =   "480"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkCommon 
            Caption         =   "Common"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   98
            Tag             =   "821"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frmExport 
         Caption         =   "Export Licence"
         Height          =   855
         Left            =   120
         TabIndex        =   95
         Tag             =   "808"
         Top             =   1320
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CheckBox chkLicenceExp 
            Caption         =   "Required "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Tag             =   "589"
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame frmImport 
         Caption         =   "Import Licence"
         Height          =   855
         Left            =   120
         TabIndex        =   77
         Tag             =   "807"
         Top             =   1320
         Width           =   8415
         Begin VB.CheckBox chkLicenceImp 
            Caption         =   "Required if value exceeds"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Tag             =   "844"
            Top             =   360
            Width           =   2650
         End
         Begin VB.TextBox txtLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3000
            MaxLength       =   14
            TabIndex        =   79
            Top             =   360
            Width           =   2655
         End
         Begin VB.ComboBox cboCurrency 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6000
            TabIndex        =   78
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblCurrency 
            Caption         =   "EUR"
            Height          =   255
            Left            =   6000
            TabIndex        =   100
            Top             =   390
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Country"
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Tag             =   "822"
         Top             =   360
         Width           =   8415
         Begin VB.TextBox txtCtryCode 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   76
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdCountry 
            Caption         =   "..."
            Height          =   300
            Left            =   2040
            TabIndex        =   75
            Top             =   360
            Width           =   280
         End
         Begin VB.TextBox txtCtry 
            Height          =   285
            Left            =   3000
            TabIndex        =   74
            Top             =   360
            Width           =   4815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Special regime"
         Height          =   2535
         Index           =   2
         Left            =   8640
         TabIndex        =   64
         Tag             =   "851"
         Top             =   360
         Width           =   2895
         Begin VB.PictureBox Picture2 
            Height          =   1810
            Left            =   120
            ScaleHeight     =   1755
            ScaleWidth      =   2535
            TabIndex        =   65
            Top             =   360
            Width           =   2600
            Begin VB.TextBox txtRegValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   4
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   56
               Text            =   "0"
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox txtRegValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   3
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   54
               Text            =   "0"
               Top             =   1140
               Width           =   1215
            End
            Begin VB.TextBox txtRegValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   52
               Text            =   "0"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox txtRegValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   1
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   50
               Text            =   "0"
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtRegValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   48
               Text            =   "0"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "..."
               Height          =   295
               Index           =   4
               Left            =   1050
               TabIndex        =   70
               Top             =   1450
               Width           =   290
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "..."
               Height          =   295
               Index           =   3
               Left            =   1050
               TabIndex        =   69
               Top             =   1150
               Width           =   290
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "..."
               Height          =   295
               Index           =   2
               Left            =   1050
               TabIndex        =   68
               Top             =   850
               Width           =   290
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "..."
               Height          =   295
               Index           =   1
               Left            =   1050
               TabIndex        =   67
               Top             =   550
               Width           =   290
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "..."
               Height          =   295
               Index           =   0
               Left            =   1050
               TabIndex        =   66
               Top             =   250
               Width           =   290
            End
            Begin VB.TextBox txtReg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   4
               Left            =   0
               MaxLength       =   2
               TabIndex        =   55
               Text            =   "0"
               Top             =   1440
               Width           =   1040
            End
            Begin VB.TextBox txtReg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   3
               Left            =   0
               MaxLength       =   2
               TabIndex        =   53
               Text            =   "0"
               Top             =   1140
               Width           =   1040
            End
            Begin VB.TextBox txtReg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   0
               MaxLength       =   2
               TabIndex        =   51
               Text            =   "0"
               Top             =   840
               Width           =   1040
            End
            Begin VB.TextBox txtReg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   1
               Left            =   0
               MaxLength       =   2
               TabIndex        =   49
               Text            =   "0"
               Top             =   540
               Width           =   1040
            End
            Begin VB.TextBox txtReg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   0
               MaxLength       =   2
               TabIndex        =   47
               Text            =   "0"
               Top             =   240
               Width           =   1040
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Value"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   1320
               TabIndex        =   72
               Tag             =   "451"
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Reg"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   71
               Tag             =   "843"
               Top             =   0
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Attached documents"
         Height          =   1575
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Tag             =   "812"
         Top             =   2280
         Width           =   8415
         Begin VB.PictureBox Picture3 
            Height          =   1060
            Left            =   120
            ScaleHeight     =   1005
            ScaleWidth      =   7635
            TabIndex        =   34
            Top             =   360
            Width           =   7695
            Begin VB.TextBox txtValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   295
               Index           =   2
               Left            =   5580
               MaxLength       =   12
               TabIndex        =   46
               Text            =   "0"
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txtValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   5580
               MaxLength       =   12
               TabIndex        =   42
               Text            =   "0"
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   5580
               MaxLength       =   12
               TabIndex        =   38
               Text            =   "0"
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox txtDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   295
               Index           =   2
               Left            =   3660
               MaxLength       =   6
               TabIndex        =   45
               Text            =   "0"
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   3660
               MaxLength       =   6
               TabIndex        =   41
               Text            =   "0"
               Top             =   480
               Width           =   1935
            End
            Begin VB.TextBox txtDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   3660
               MaxLength       =   6
               TabIndex        =   37
               Text            =   "0"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtNumber 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   295
               Index           =   2
               Left            =   1860
               MaxLength       =   7
               TabIndex        =   44
               Text            =   "0"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtNumber 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1860
               MaxLength       =   7
               TabIndex        =   40
               Text            =   "0"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtNumber 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1860
               MaxLength       =   7
               TabIndex        =   36
               Text            =   "0"
               Top             =   240
               Width           =   1815
            End
            Begin VB.CommandButton cmdType 
               Caption         =   "..."
               Height          =   250
               Index           =   2
               Left            =   1600
               TabIndex        =   59
               Top             =   760
               Width           =   255
            End
            Begin VB.CommandButton cmdType 
               Caption         =   "..."
               Height          =   250
               Index           =   1
               Left            =   1600
               TabIndex        =   58
               Top             =   510
               Width           =   255
            End
            Begin VB.CommandButton cmdType 
               Caption         =   "..."
               Height          =   250
               Index           =   0
               Left            =   1600
               TabIndex        =   57
               Top             =   250
               Width           =   255
            End
            Begin VB.TextBox txtType 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   0
               MaxLength       =   5
               TabIndex        =   35
               Text            =   "0"
               Top             =   240
               Width           =   1870
            End
            Begin VB.TextBox txtType 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   0
               MaxLength       =   5
               TabIndex        =   39
               Text            =   "0"
               Top             =   495
               Width           =   1870
            End
            Begin VB.TextBox txtType 
               Appearance      =   0  'Flat
               Height          =   295
               Index           =   2
               Left            =   0
               MaxLength       =   5
               TabIndex        =   43
               Text            =   "0"
               Top             =   735
               Width           =   1870
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Value"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   5580
               TabIndex        =   63
               Tag             =   "451"
               Top             =   0
               Width           =   2055
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   3660
               TabIndex        =   62
               Tag             =   "747"
               Top             =   0
               Width           =   1935
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Number"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   1860
               TabIndex        =   61
               Tag             =   "838"
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Type"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   60
               Tag             =   "438"
               Top             =   0
               Width           =   1870
            End
         End
      End
   End
   Begin VB.Frame frmQuantities 
      Caption         =   "Quantities"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Tag             =   "841"
      Top             =   1920
      Width           =   11655
      Begin VB.CheckBox chkGrossCalc 
         Caption         =   "Lock Calculated Weight"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   26
         Tag             =   "831"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkSuppStat 
         Caption         =   "Lock quantity"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   25
         Tag             =   "832"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkSuppCalc 
         Caption         =   "Lock quantity"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   24
         Tag             =   "832"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cboSuppStat 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox cboSuppCalc 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cboGrosCalc 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox cboSuppStatQ 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboSuppCalcQ 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   19
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Supplementary Statistical Unit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Tag             =   "853"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Supplementary Calculation Unit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Tag             =   "852"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Gross Weight Calculation"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Tag             =   "828"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity Handling"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   28
         Tag             =   "842"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity Handling"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5520
         TabIndex        =   27
         Tag             =   "842"
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   11655
      Begin VB.CommandButton cmdExport 
         Caption         =   "&New Country..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   10080
         TabIndex        =   17
         Tag             =   "837"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   10080
         TabIndex        =   16
         Tag             =   "260"
         Top             =   580
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Make a copy..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   10080
         TabIndex        =   15
         Tag             =   "355"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "C&hange settings..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   10080
         TabIndex        =   14
         Tag             =   "356"
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Set as default"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   10080
         TabIndex        =   13
         Tag             =   "848"
         Top             =   1640
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvwExport 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "839"
            Text            =   "Origin"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lic"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "N1"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Q1"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "N2"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Q2"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "N3"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Q3"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "R1"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "R2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "R3"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "R4"
            Object.Width           =   1235
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   11655
      Begin VB.CommandButton cmdImport 
         Caption         =   "&New Country..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   10080
         TabIndex        =   10
         Tag             =   "837"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   10080
         TabIndex        =   9
         Tag             =   "260"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Make a copy..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   10080
         TabIndex        =   8
         Tag             =   "355"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "C&hange settings..."
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   10080
         TabIndex        =   7
         Tag             =   "356"
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Set as default"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   10080
         TabIndex        =   6
         Tag             =   "848"
         Top             =   1640
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvwImport 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "839"
            Text            =   "Origin"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lic"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "N1"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Q1"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "N2"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Q2"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "N3"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Q3"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "R1"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "R2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "R3"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "R4"
            Object.Width           =   1235
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Tag             =   "178"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Tag             =   "180"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Tag             =   "179"
      Top             =   7800
      Width           =   1335
   End
End
Attribute VB_Name = "frm_taricmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strLangOfDesc As String
Private strDocType As String

Private strMainType As String
Private strMainLoaded As String

Public strDetailLoaded As String
Public strDetailType As String

Private m_conTaric As ADODB.Connection    'Private datTaric As DAO.Database
Private m_conSADBEL As ADODB.Connection     'Private datPick As DAO.Database

Dim m_rstMain As ADODB.Recordset '-----> CN CODE when txtcode is changed
Dim m_rstImpGrid As ADODB.Recordset '-----> Import in normal
Dim m_rstExpGrid As ADODB.Recordset '-----> export in normal
Dim m_rstPick As ADODB.Recordset '-----> picklist maintenance in sadbel
Dim m_rstCommon As ADODB.Recordset '----->
Dim m_rstDetail As ADODB.Recordset '-----> import and export in simplified
Dim m_rstSupp As ADODB.Recordset '-----> supp units table used for default vaues in cbo and conversions
Dim m_rstDefault As ADODB.Recordset '-----> import and export in simplified using the default country
Dim m_rstEmptyFieldValues As ADODB.Recordset

Dim blnItemImport As Boolean
Dim blnItemExport As Boolean

'Public strCallType As String
'Dim intSBcode As Integer

Dim blnUncheck As Boolean

'-----> to skip a procedure in picklist
Public blnTaricMain As Boolean

'-----> to verify if ok and apply would be enabled
Dim blnSaveOK As Boolean

'-----> for the setting of gblnformwascancelled
Dim blnApplyClicked As Boolean

'-----> minimum value and currency in import licence
' Dim strLicCurr As String
Dim dblMinVal As Double
Const Display = "##########0.00"

'-----> used if loaded as make a copy
Dim blnCopy As Boolean

'-----> for the default empty fields in the settings
Dim colEmptyFieldValues As Collection

'-----> For third party database use
Dim intThirdPartyDatabase As Integer

Dim blnKluwerClicked As Boolean
Dim intKluwerCountry As Integer
Private mdblTaskID As Double

Private Const SW_SHOWNORMAL = 1    ' Restores window if minimized or maximized
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

'----> for ascii converter
'Dim ACode As String
Dim SBcode As Integer

'-----> to conform with apply command standard
Dim blnStartApply As Boolean

Private strCtryCode() As String
Private colTaricCodes As Collection

Private Sub cboGrosCalc_Change()
    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cboSuppCalc_Change()

    Dim Counter As Integer
    
    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
    
    If Left(cboSuppCalc.Text, 2) = "00" Then
        cboSuppCalcQ.Enabled = False
        cboSuppCalcQ.Text = ""
    Else
        cboSuppCalcQ.Enabled = True
        
    ' ********** Commented September 6, 2001 **********
    ' ********** New Rule:  No assumed SuppCalcQtyCode anymore because it confuses users!
        If cboSuppCalcQ.Text = "" Then
            If strLangOfDesc = "English" Then
                cboSuppCalcQ.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppCalcQ.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppCalcQ.Text = "00-Aucun"
            End If
        End If
    End If

End Sub

Private Sub cboSuppCalc_Click()

    Dim Counter As Integer
    
    If Left(cboSuppCalc.Text, 2) = "00" Then
        cboSuppCalcQ.Enabled = False
        cboSuppCalcQ.Text = ""
    Else
        cboSuppCalcQ.Enabled = True
        
    ' ********** Commented September 6, 2001 **********
    ' ********** New Rule:  No assumed SuppCalcQtyCode anymore because it confuses users!
        If cboSuppCalcQ.Text = "" Then
            If strLangOfDesc = "English" Then
                cboSuppCalcQ.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppCalcQ.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppCalcQ.Text = "00-Aucun"
            End If
        End If
    End If
End Sub

Private Sub cboSuppCalcQ_Change()
    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cboSuppStat_Change()

    Dim Counter As Integer
    
    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
    
    If Left(cboSuppStat.Text, 2) = "00" Then
        cboSuppStatQ.Enabled = False
        cboSuppStatQ.Text = ""
    Else:
        cboSuppStatQ.Enabled = True
        
    ' ********** Commented September 6, 2001 **********
    ' ********** New Rule:  No assumed SuppStatQtyCode anymore because it confuses users!
        If cboSuppStatQ.Text = "" Then
            If strLangOfDesc = "English" Then
                cboSuppStatQ.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppStatQ.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppStatQ.Text = "00-Aucun"
            End If
        End If
    End If

End Sub

Private Sub cboSuppStat_Click()

    Dim Counter As Integer
    
    If Left(cboSuppStat.Text, 2) = "00" Then
        cboSuppStatQ.Enabled = False
        cboSuppStatQ.Text = ""
        
    Else:
        cboSuppStatQ.Enabled = True
        
    ' ********** Commented September 6, 2001 **********
    ' ********** New Rule:  No assumed SuppStatQtyCode anymore because it confuses users!
        If cboSuppStatQ.Text = "" Then
            If strLangOfDesc = "English" Then
                cboSuppStatQ.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppStatQ.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppStatQ.Text = "00-Aucun"
            End If
        End If
    End If

End Sub

Private Sub cboSuppStatQ_Change()

    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkCommon_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkDefault_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
    
    If chkDefault.Value = 1 Then
        chkCommon.Enabled = True
        blnUncheck = False
    Else
        chkCommon.Enabled = False
        chkCommon.Value = 0
        blnUncheck = True
    End If

End Sub

Private Sub chkGrossCalc_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkLicenceExp_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkSuppCalc_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkSuppStat_Click()

    '----> to conform with apply command standard
    If blnStartApply Then
    
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmdApply_Click()
    Dim strTempKey As String
    Dim strTempDesc As String
    Dim Counter As Integer
    
    
    If IsCompletelyFilled Then
       
    
        frm_taricmain.MousePointer = 11
        
        '----> so that the kluwer code which was added would not be deleted
        blnKluwerClicked = False
        
        If strMainType = "Normal" Then
            If strMainLoaded = "Copy" Then
                SaveCopy
            End If
            
            SaveChanges
            '-----> Add to maintenance listview newly added code
            With frm_taricmaintenance.lvwMaintenance
            
                strTempKey = IIf(strLangOfDesc = "Dutch", txtDutchKey.Text, txtFrnchKey.Text)
                strTempDesc = IIf(strLangOfDesc = "Dutch", txtDutchDesc.Text, txtFrnchDesc.Text)
                
        On Error GoTo Add:
        
                If .FindItem(txtCode.Text, , , 0) = 0 Then
                
                    frm_taricmain.MousePointer = 0
                    
                    Exit Sub
Add:
                    .Sorted = False
                    .ListItems.Add , , txtCode.Text
                    .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                    .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                    .Sorted = True
                    For Counter = 1 To .ListItems.Count
                        If .ListItems(Counter).Text = txtCode.Text Then
                            .ListItems(Counter).EnsureVisible
                            .ListItems(Counter).Selected = True
                        End If
                    Next Counter
                Else
                    '-----> If record exist in listview delete it and add the new one
                    .Sorted = False
                    .ListItems.Remove (.FindItem(txtCode.Text, , , 0).Index)
                    .ListItems.Add , , txtCode.Text
                    .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                    .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                    .ListItems(.ListItems.Count).Selected = True
                    .Sorted = True
                    For Counter = 1 To .ListItems.Count
                        If .ListItems(Counter).Text = txtCode.Text Then
                            .ListItems(Counter).EnsureVisible
                            .ListItems(Counter).Selected = True
                        End If
                    Next Counter
                End If
            End With
        Else
            SaveSimplified
            SaveChanges
            blnApplyClicked = True
            gblnFormWasCanceled = False
            '-----> do this only if loaded from picklist
            If Left(Right(gstrTaricMainCallType, Len(gstrTaricMainCallType) - InStr(gstrTaricMainCallType, "/")), 17) = "frm_taricpicklist" Then
                frm_taricpicklist.blnOKWasPressed = True
                '-----> Add to maintenance listview newly added code
                With frm_taricpicklist.lvwPicklist
                    strTempKey = IIf(strLangOfDesc = "Dutch", txtDutchKey.Text, txtFrnchKey.Text)
                    strTempDesc = IIf(strLangOfDesc = "Dutch", txtDutchDesc.Text, txtFrnchDesc.Text)
                    On Error GoTo AddPick:
                    If .FindItem(txtCode.Text, , , 0) = 0 Then
                        frm_taricmain.MousePointer = 0
                        Exit Sub
AddPick:
                        .Sorted = False
                        .ListItems.Add , , txtCode.Text
                        .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                        .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                        .Sorted = True
                        For Counter = 1 To .ListItems.Count
                            If .ListItems(Counter).Text = txtCode.Text Then
                                .ListItems(Counter).EnsureVisible
                                .ListItems(Counter).Selected = True
                            End If
                        Next Counter
                    Else
                        '-----> If record exist in listview delete it and add the new one
                        .Sorted = False
                        .ListItems.Remove (.FindItem(txtCode.Text, , , 0).Index)
                        .ListItems.Add , , txtCode.Text
                        .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                        .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                        .ListItems(.ListItems.Count).Selected = True
                        .Sorted = True
                        For Counter = 1 To .ListItems.Count
                            If .ListItems(Counter).Text = txtCode.Text Then
                                .ListItems(Counter).EnsureVisible
                                .ListItems(Counter).Selected = True
                            End If
                        Next Counter
                    End If
                End With
            End If
            frm_taricmain.MousePointer = 0
        End If
        cmdApply.Enabled = False
        frm_taricmain.MousePointer = 0
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim strCN As String
    
    '-----> delete saved kluwer to database
    If blnKluwerClicked = True Then
        Call DeleteKluwer
        blnKluwerClicked = False
    End If
    
    frm_taricmain.MousePointer = 11
    '-----> check if common is saved
    If strMainType = "Normal" Then
        '-----> check if code exist in cn picklist
            'allanSQL
            strCN = vbNullString
            strCN = strCN & "SELECT "
            strCN = strCN & "* "
            strCN = strCN & "FROM "
            strCN = strCN & "CN "
            strCN = strCN & "WHERE "
            strCN = strCN & "[CN CODE] = " & Chr(39) & ProcessQuotes(Left(txtCode.Text, 8)) & Chr(39) & " "
        ADORecordsetOpen strCN, m_conTaric, m_rstMain, adOpenKeyset, adLockOptimistic
        'Set m_rstMain = m_conTaric.OpenRecordset(strCN, dbOpenForwardOnly)
        With m_rstMain
        
            If Not (.EOF And .BOF) Then
                If Len(txtCode.Text) = 10 Then
                    CancelCheck
                End If
            End If
        End With
        
        ADORecordsetClose m_rstMain
    End If
    
    If blnApplyClicked = False Then
        gblnFormWasCanceled = True
    End If
    
    UnloadControls Me
     
    frm_taricmain.MousePointer = 0
    Set frm_taricmain = Nothing
    
    Unload Me

End Sub

Private Sub cmdCountry_Click()
    Dim strBoxProp As String
    Dim strPickVal As String
    Dim Admin As String
    
    Admin = GetSetting(App.Title, "Settings", "AdminRights")
    '----> Use default string used in codisheet
    If strDocType = "Import" Then
        strBoxProp = "C1#*#H20#*#1#*#" & strLangOfDesc & "#*#" & "Import#*#" & _
        Admin & "#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    Else
        strBoxProp = "C2#*#H21#*#1#*#" & strLangOfDesc & "#*#" & "Import#*#" & _
        Admin & "#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    End If
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    blnTaricMain = True
    g_blnMultiplePick = False
    
    frm_picklist.Show vbModal, Me
    
    blnTaricMain = False
    
    strPickVal = GetSetting(App.Title, "CodiSheet", "Pick_TARIC_PL")
    '-----> get the country code
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then
        txtCtryCode.Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
    End If
    
    DeleteSetting App.Title, "CodiSheet", "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"

End Sub

Private Sub cmdExport_Click(Index As Integer)
    Dim Counter As Integer
    Dim intDelete As Integer
    
    If strMainLoaded = "Copy" Then
        SaveCopy
        strMainLoaded = "Modify"
    End If
    
    strDetailType = "Export"
    
    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then
                cmdApply.Enabled = True
            End If
        End If
    End If

    On Error GoTo Labas:
    Select Case Index
        Case 0 '-----> New Country
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "New Country"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 1 '-----> Delete
                '-----> Confirm to the user if the item is to be deleted
                intDelete = MsgBox(Translate(415) & " " & lvwExport.SelectedItem.Text & "?", vbYesNo + vbCritical + vbApplicationModal, Me.Caption)
                If intDelete = vbNo Then Exit Sub
                
                '-----> If no selected item don't do anything
                'If lvwExport.SelectedItem.Index < 1 Then Exit Sub
                frm_taricmain.MousePointer = 11
                With m_rstExpGrid
                    .MoveFirst
                    Do While Not .EOF
                        '-----> Look for the selected item in the table which is to be deleted
                        If ![CTRY CODE] = Left(lvwExport.SelectedItem.Text, (InStr(lvwExport.SelectedItem.Text, " ") - 1)) Then
                            
                            Dim strCommand As String
                            
                            If ![DEF CODE] = -1 Then '-----> if item to be deleted is the default set the item with lowest count as the default
                                .Delete
                                
                                    strCommand = vbNullString
                                    strCommand = strCommand & "DELETE "
                                    strCommand = strCommand & "* "
                                    strCommand = strCommand & "FROM "
                                    strCommand = strCommand & "[EXPORT] "
                                    strCommand = strCommand & "WHERE "
                                    strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ![CTRY CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ![TARIC CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "![DEF CODE] = -1 "
                                ExecuteNonQuery m_conTaric, strCommand
    
                                ExportAdd
                                lvwExport.ListItems(1).Bold = True
                                For Counter = 1 To 11
                                    lvwExport.ListItems(1).ListSubItems(Counter).Bold = True
                                Next Counter
                                '-----> Save new default country
                                .MoveFirst
                                Do While Not .EOF
                                    If ![CTRY CODE] = Left(lvwExport.ListItems(1).Text, (InStr(lvwExport.ListItems(1).Text, " ") - 1)) Then
                                        '.Edit
                                        ![DEF CODE] = -1
                                        .Update
                                        
                                        UpdateRecordset m_conTaric, m_rstExpGrid, "EXPORT"
                                        
                                        frm_taricmain.MousePointer = 0
                                        Exit Sub
                                    End If
                                .MoveNext
                                Loop
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            Else
                                .Delete
                                    strCommand = vbNullString
                                    strCommand = strCommand & "DELETE "
                                    strCommand = strCommand & "* "
                                    strCommand = strCommand & "FROM "
                                    strCommand = strCommand & "[EXPORT] "
                                    strCommand = strCommand & "WHERE "
                                    strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ![CTRY CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ![TARIC CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "![DEF CODE] <> -1 "
                                ExecuteNonQuery m_conTaric, strCommand
                                
                                ExportAdd
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            End If
                        End If
                    .MoveNext
                    Loop
                End With
                frm_taricmain.MousePointer = 0
        Case 2 '-----> Make a copy
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "Copy"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 3 '-----> Change settings
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "Change"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 4 '-----> Set as default
                frm_taricmain.MousePointer = 11
                
                '-----> Get the country code from the grid
                Dim strCtryCode As String
                
                strCtryCode = Left(lvwExport.SelectedItem.Text, (InStr(lvwExport.SelectedItem.Text, " ") - 1))
                
                With m_rstExpGrid
                    .MoveFirst
                    Do While Not .EOF
                    If ![CTRY CODE] = strCtryCode Then
                        '----->Save the new default country
                        '.Edit
                        ![DEF CODE] = -1
                        .Update
                        '----->Remove previous default
                        .MoveFirst
                        Do While Not .EOF
                            If Not ![CTRY CODE] = strCtryCode Then
                                '.Edit
                                ![DEF CODE] = 0
                                ![COMM CODE] = 0
                                .Update
                            End If
                            .MoveNext
                        Loop
                        frm_taricmain.MousePointer = 0
                        ExportAdd
                        cmdImport(4).Enabled = False
                        Exit Sub
                    End If
                    .MoveNext
                    Loop
                End With
                
                UpdateDefaultCountry strDetailType, strCtryCode
                
    End Select
Labas:
    frm_taricmain.MousePointer = 0
End Sub

Private Sub UpdateDefaultCountry(ByVal TableName As String, _
                                 ByVal NewDefaultCountryCode As String)
    Dim strCommand As String
    
        strCommand = vbNullString
        strCommand = strCommand & "UPDATE "
        strCommand = strCommand & "[" & TableName & "] "
        strCommand = strCommand & "SET "
        strCommand = strCommand & "[DEF CODE] = -1 "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & NewDefaultCountryCode & Chr(39) & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ExecuteNonQuery m_conTaric, strCommand
    
        strCommand = vbNullString
        strCommand = strCommand & "UPDATE "
        strCommand = strCommand & "[" & TableName & "] "
        strCommand = strCommand & "SET "
        strCommand = strCommand & "[DEF CODE] = 0, "
        strCommand = strCommand & "[COMM CODE] = 0 "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "[CTRY CODE] <> " & Chr(39) & NewDefaultCountryCode & Chr(39) & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ExecuteNonQuery m_conTaric, strCommand
    
    frm_taricmain.MousePointer = 0
    ExportAdd
    cmdExport(4).Enabled = False
                
End Sub

Private Sub cmdGeneral_Click(Index As Integer)

' ********** 10/18/02 **********
' Modified various parts of module to
' accomodate possible new version of
' KLUWER. Config.sgm will be replaced
' by Config.xml, the Country name shall
' no longer be used but instead the country
' code will be supplied. The items in the ECO
' return shall not be found in parenthesis anymore.
' **********   END.   **********

    '-----> For TARBEL: Table of Contents HTML page; for Kluwer: config.sgm text file.
    Dim strFileName As String
    '-----> For TARBEL: HTML browser; for Kluwer: Kluwer database application.
    Dim strDefaultBrowser As String
    '-----> For TARBEL: dummy variable; for Kluwer: current directory.
    Dim strDummyDirectory As String
    '-----> For TARBEL: HTML file was opened successfully; for Kluwer: backslash before the Kluwer database application file name was found.
    Dim lngRetVal As Long
    
    '-----> For Kluwer
    Dim intFreeFile As Integer
    Dim intSubscript As Integer
    Dim strTempText() As String
    Dim strClipboard As String
    Dim strPathName As String
    
    Dim udtBrowseInfo As BROWSEINFO
    Dim pIDList As Long
    
    Select Case Index
        Case 0    '-----> CN Codes
            Me.MousePointer = vbHourglass
            gstrTaricCNCallType = "frm_taricmain"
            frm_tariccn.Show vbModal, Me
            Me.MousePointer = vbDefault
        Case 1    '-----> Kluwer
            Select Case intThirdPartyDatabase
                Case 1    '-----> TARBEL
                    strFileName = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchHTML"), GetSetting(App.Title, "Third Party Database", "FrenchHTML"))
                    strDefaultBrowser = Space$(255)
                    
                    '-----> Find the application associated with HTML files.
                    lngRetVal = FindExecutable(strFileName, strDummyDirectory, strDefaultBrowser)
                    strDefaultBrowser = Trim$(Replace(strDefaultBrowser, vbNullChar, " "))
                    
                    '-----> If an application is found, launch it!
                    If lngRetVal <= 32 Or Len(strDefaultBrowser) = 0 Then
                        MsgBox Translate(1040), vbExclamation   '"Could not find associated browser."
                    Else
                        lngRetVal = ShellExecute(Me.hwnd, "open", strFileName, 0&, strDummyDirectory, SW_SHOWNORMAL)
                        
                        If lngRetVal <= 32 Then
                            MsgBox Translate(1041), vbExclamation   '"HTML file cannot be opened."
                        End If
                    End If
                Case 2    '-----> Kluwer
                    If mdblTaskID = 0 Then    ' No instance of Kluwer database application is running.
                    strFileName = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchFile"), GetSetting(App.Title, "Third Party Database", "FrenchFile"))
                    
                    On Error GoTo FileErrHandler
                    
                    ReDim strTempText(12)
                    
                    ' These default values will be used if config.sgm doesn't exist or
                    ' if it does indeed exist but has missing lines or no lines at all.
                    strTempText(1) = "<Config>"
                    strTempText(4) = "<ThirdWindow>0</ThirdWindow>"
                    
                    If strLangOfDesc = "French" Then
                        strTempText(5) = "<Vertical>0</Vertical>"
                    Else
                        strTempText(5) = "<Verticaal>0</Verticaal>"
                    End If
                    
                    strTempText(6) = "<InstellingenTar>0</InstellingenTar>"
                    strTempText(7) = "<IncludeTar>0000000000</IncludeTar>"
                    strTempText(8) = "<StartCX>640</StartCX>"
                    strTempText(9) = "<StartCY>480</StartCY>"
                    strTempText(10) = "<PiramidReport>0</PiramidReport>"
                    strTempText(11) = "<NoStartMessage>0</NoStartMessage>"
                    strTempText(12) = "</Config>"
                    
                    If Len(Dir(strFileName)) Then
                        intFreeFile = FreeFile()
                        
                        Open strFileName For Input As #intFreeFile
                        
                        Do Until EOF(intFreeFile)
                            intSubscript = intSubscript + 1
                            
                            If intSubscript > 12 Then
                                ReDim Preserve strTempText(intSubscript)
                            End If
                            
                            Line Input #intFreeFile, strTempText(intSubscript)
                        Loop
                        
                        Close #intFreeFile
                    End If
                    
                    intFreeFile = FreeFile()
                    
                    Open strFileName For Output As #intFreeFile
                    
                    For intSubscript = 1 To 12
                        If intSubscript = 2 Then
                            If txtCode.Text <> "" Then
                                strClipboard = "<SelectCode>" & txtCode.Text & "</SelectCode>"
                            Else
                                strClipboard = "<SelectCode></SelectCode>"
                            End If
                        ElseIf intSubscript = 3 Then
                            If strLangOfDesc = "French" Then
                                If strMainType = "Simplified" Then
                                    ' ***** 10/18/02 *****
                                    ' If file extension of config is sgm then old version. Use country name.
                                    ' If xml, use country code instead.
                                    ' *****   end.   *****
                                    'strClipboard = "<SelectPays>" & IIf(Len(txtCtry.Text), txtCtry.Text, frm_taricpicklist.strCtryName) & "</SelectPays>"
                                    If Right(Trim(strFileName), 3) = "sgm" Then
                                        strClipboard = "<SelectPays>" & IIf(Len(txtCtry.Text), txtCtry.Text, frm_taricpicklist.strCtryName) & "</SelectPays>"
                                    Else
                                        strClipboard = "<SelectPays>" & IIf(Len(txtCtryCode.Text), txtCtryCode.Text, frm_taricpicklist.strCtryCode) & "</SelectPays>"
                                    End If
                                Else
                                    strClipboard = "<SelectPays></SelectPays>"
                                End If
                            Else
                                If strMainType = "Simplified" Then
                                    ' ***** 10/18/02 *****
                                    ' If file extension of config is sgm then old version. Use country name.
                                    ' If xml, use country code instead.
                                    ' *****   end.   *****
                                    'strClipboard = "<SelectLand>" & IIf(Len(txtCtry.Text), txtCtry.Text, frm_taricpicklist.strCtryName) & "</SelectLand>"
                                    If Right(Trim(strFileName), 3) = "sgm" Then
                                        strClipboard = "<SelectLand>" & IIf(Len(txtCtry.Text), txtCtry.Text, frm_taricpicklist.strCtryName) & "</SelectLand>"
                                    Else
                                        strClipboard = "<SelectLand>" & IIf(Len(txtCtryCode.Text), txtCtryCode.Text, frm_taricpicklist.strCtryCode) & "</SelectLand>"
                                    End If
                                Else
                                    strClipboard = "<SelectLand></SelectLand>"
                                End If
                            End If
                            
                            ' Replace special characters with characters as they appear in config.sgm.
                            strClipboard = Replace(strClipboard, "&", "&amp;")
                            strClipboard = Replace(strClipboard, "'", "&rquo;")
                            strClipboard = Replace(strClipboard, "-", "&ndash;")
                            strClipboard = Replace(strClipboard, "é", "&eacute;")
                            strClipboard = Replace(strClipboard, "ç", "‡")
                            strClipboard = Replace(strClipboard, "ê", "ˆ")
                            strClipboard = Replace(strClipboard, "ë", "‰")
                            strClipboard = Replace(strClipboard, "è", "Š")
                            strClipboard = Replace(strClipboard, "ï", "‹")
                            strClipboard = Replace(strClipboard, "î", "Œ")
                            strClipboard = Replace(strClipboard, "ô", "“")
                            strClipboard = Replace(strClipboard, "ö", "”")
                        Else
                            strClipboard = strTempText(intSubscript)
                        End If
                        
                        Print #intFreeFile, strClipboard
                    Next
                    
                    Close #intFreeFile
                    
                    '-----> Launch Kluwer program
                    strDefaultBrowser = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchCmd"), GetSetting(App.Title, "Third Party Database", "FrenchCmd"))
                    
                    strTempText() = Split(strDefaultBrowser, " -")
                    strTempText(0) = Replace(strTempText(0), Chr(34), "")
                    
                    lngRetVal = InStrRev(strFileName, "\")
                    
                    If lngRetVal Then
                        strPathName = Left(strFileName, lngRetVal - 1)
                    Else
                        strPathName = "C:\Program Files\DBTARN"    ' Default directory.
                    End If
                    
                    strFileName = ""
                    strDummyDirectory = CurDir()
                    Clipboard.SetText "No Data"
                    
                    ChDir strPathName          ' Change current directory.
                    
                    ' Re-used variable; proxies for task ID.
                    mdblTaskID = Shell(strDefaultBrowser, vbNormalFocus)
                    ChDir strDummyDirectory    ' Revert to previous current directory.
                    Timer1.Enabled = True
                    
                    End If
            End Select
        Case 2    '-----> Clients
            Me.MousePointer = vbHourglass
            frm_taricclients.Show vbModal, Me
            Me.MousePointer = vbDefault
    End Select
    
    Exit Sub
    
FileErrHandler:
    
    Select Case Err.Number
        Case 53    ' File not found
            If MsgBox(Err.Description & ":" & vbCrLf & strTempText(0), vbRetryCancel + vbExclamation, Err.Source & " (" & Err.Number & ")") = vbRetry Then
                Resume
            End If
        Case 76    ' Path not found
            If Len(strFileName) Then
                strPathName = strFileName
                lngRetVal = InStrRev(strPathName, "\")
                strFileName = Mid(strPathName, lngRetVal)
            End If
            
            With udtBrowseInfo
                .hWndOwner = Me.hwnd
                .lpszTitle = "The path " & strPathName & " could not be found.  Please select a folder below, then click OK."
                .ulFlags = BIF_RETURNONLYFSDIRS
                .pIDListRoot = 0
            End With
            
            strPathName = String(255, vbNullChar)
            pIDList = SHBrowseForFolder(udtBrowseInfo)
            
            If SHGetPathFromIDList(pIDList, strPathName) Then
                If Len(strFileName) Then
                    strFileName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 1) & strFileName
                Else
                    strPathName = Left(strPathName, InStr(1, strPathName, vbNullChar) - 1)
                End If
                
                Resume
            End If
        Case Else
            If MsgBox(Err.Description, vbRetryCancel + vbExclamation, Err.Source & " (" & Err.Number & ")") = vbRetry Then
                Resume
            End If
    End Select
End Sub


Private Sub cmdImport_Click(Index As Integer)
    Dim Counter As Integer
    Dim intDelete
    If strMainLoaded = "Copy" Then
        SaveCopy
        strMainLoaded = "Modify"
    End If
    
    '----> to conform with apply command standard
    If blnStartApply Then
        If strMainType = "Normal" Then
            cmdApply.Enabled = True
        Else
            If blnSaveOK = True Then cmdApply.Enabled = True
        End If
    End If
    
    
    strDetailType = "Import"
    On Error GoTo Labas:
    Select Case Index
        Case 0 '-----> New Country
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "New Country"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 1 '-----> Delete
                '-----> Confirm to the user if the item is to be deleted
                intDelete = MsgBox(Translate(415) & " " & lvwImport.SelectedItem.Text & "?", vbYesNo + vbCritical + vbApplicationModal, Me.Caption)
                If intDelete = vbNo Then Exit Sub
                
                frm_taricmain.MousePointer = 11
                With m_rstImpGrid
                    .MoveFirst
                    Do While Not .EOF
                        '-----> Look for the selected item in the table which is to be deleted
                        If ![CTRY CODE] = Left(lvwImport.SelectedItem.Text, (InStr(lvwImport.SelectedItem.Text, " ") - 1)) Then
                            Dim strCommand As String
                            
                            If ![DEF CODE] = -1 Then '-----> if item to be deleted is the default set the item with lowest count as the default
                                .Delete

                                    strCommand = vbNullString
                                    strCommand = strCommand & "DELETE "
                                    strCommand = strCommand & "* "
                                    strCommand = strCommand & "FROM "
                                    strCommand = strCommand & "[IMPORT] "
                                    strCommand = strCommand & "WHERE "
                                    strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ![CTRY CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ![TARIC CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "![DEF CODE] = -1 "
                                ExecuteNonQuery m_conTaric, strCommand
                                
                                ImportAdd
                                lvwImport.ListItems(1).Bold = True
                                
                                For Counter = 1 To 11
                                    lvwImport.ListItems(1).ListSubItems(Counter).Bold = True
                                Next Counter
                                '-----> Save new default country
                                .MoveFirst
                                Do While Not .EOF
                                    If ![CTRY CODE] = Left(lvwImport.ListItems(1).Text, (InStr(lvwImport.ListItems(1).Text, " ") - 1)) Then
                                        '.Edit
                                        ![DEF CODE] = -1
                                        .Update
                                        
                                        UpdateRecordset m_conTaric, m_rstImpGrid, "IMPORT"
                                        
                                        frm_taricmain.MousePointer = 0
                                        Exit Sub
                                    End If
                                .MoveNext
                                Loop
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            Else
                                .Delete
                                
                                    strCommand = vbNullString
                                    strCommand = strCommand & "DELETE "
                                    strCommand = strCommand & "* "
                                    strCommand = strCommand & "FROM "
                                    strCommand = strCommand & "[IMPORT] "
                                    strCommand = strCommand & "WHERE "
                                    strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ![CTRY CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ![TARIC CODE] & Chr(39) & " "
                                    strCommand = strCommand & "AND "
                                    strCommand = strCommand & "![DEF CODE] <> -1 "
                                ExecuteNonQuery m_conTaric, strCommand
                                
                                ImportAdd
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            End If
                        End If
                    .MoveNext
                    Loop
                End With
                frm_taricmain.MousePointer = 0
        Case 2 '-----> Make a copy
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "Copy"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 3 '-----> Change settings
                frm_taricmain.MousePointer = 11
                strDetailLoaded = "Change"
                Load frm_taricdetail
                frm_taricdetail.Show vbModal, Me
        Case 4 '-----> Set as default
                frm_taricmain.MousePointer = 11
                
                '-----> Get the country code from the grid
                Dim strCtryCode As String
                strCtryCode = Left(lvwImport.SelectedItem.Text, (InStr(lvwImport.SelectedItem.Text, " ") - 1))

                With m_rstImpGrid
                    .MoveFirst
                    Do While Not .EOF
                    If ![CTRY CODE] = strCtryCode Then
                        '----->Save the new default country
                        '.Edit
                        ![DEF CODE] = -1
                        .Update
                        '----->Remove previous default
                        .MoveFirst
                        Do While Not .EOF
                            If Not ![CTRY CODE] = strCtryCode Then
                                '.Edit
                                ![DEF CODE] = 0
                                ![COMM CODE] = 0
                                .Update
                            End If
                            .MoveNext
                        Loop
                        frm_taricmain.MousePointer = 0
                        ImportAdd
                        cmdImport(4).Enabled = False
                        Exit Sub
                    End If
                    .MoveNext
                    Loop
                End With
                
                UpdateDefaultCountry strDetailType, strCtryCode
                
    End Select
Labas:
    frm_taricmain.MousePointer = 0
End Sub

Private Sub cmdOK_Click()
    Dim strTempKey As String
    Dim strTempDesc As String
    Dim Counter As Integer
    
        If Not IsCompletelyFilled Then
            Exit Sub
        End If
    frm_taricmain.MousePointer = 11
    
    If strMainType = "Normal" Then
        If strMainLoaded = "Copy" Then SaveCopy
        SaveChanges
        '-----> Add to maintenance listview newly added code
        With frm_taricmaintenance.lvwMaintenance
            strTempKey = IIf(strLangOfDesc = "Dutch", txtDutchKey.Text, txtFrnchKey.Text)
            strTempDesc = IIf(strLangOfDesc = "Dutch", txtDutchDesc.Text, txtFrnchDesc.Text)
            On Error GoTo Add:
            If .FindItem(txtCode.Text, , , 0) = 0 Then
                frm_taricmain.MousePointer = 0
                Exit Sub
Add:
                .Sorted = False
                .ListItems.Add , , txtCode.Text
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                .Sorted = True
                For Counter = 1 To .ListItems.Count
                    If .ListItems(Counter).Text = txtCode.Text Then
                        .ListItems(Counter).EnsureVisible
                        .ListItems(Counter).Selected = True
                    End If
                Next Counter
            Else
                '-----> If record exist in listview delete it and add the new one
                .Sorted = False
                .ListItems.Remove (.FindItem(txtCode.Text, , , 0).Index)
                .ListItems.Add , , txtCode.Text
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                .ListItems(.ListItems.Count).Selected = True
                .Sorted = True
                For Counter = 1 To .ListItems.Count
                    If .ListItems(Counter).Text = txtCode.Text Then
                        .ListItems(Counter).EnsureVisible
                        .ListItems(Counter).Selected = True
                    End If
                Next Counter
            End If
        End With
    Else
        SaveSimplified
        SaveChanges
        gblnFormWasCanceled = False
        '-----> do this only if loaded from picklist
        If Left(Right(gstrTaricMainCallType, Len(gstrTaricMainCallType) - InStr(gstrTaricMainCallType, "/")), 17) _
        = "frm_taricpicklist" Then
        frm_taricpicklist.blnOKWasPressed = True
        '-----> Add to maintenance listview newly added code
        With frm_taricpicklist.lvwPicklist
            strTempKey = IIf(strLangOfDesc = "Dutch", txtDutchKey.Text, txtFrnchKey.Text)
            strTempDesc = IIf(strLangOfDesc = "Dutch", txtDutchDesc.Text, txtFrnchDesc.Text)
            On Error GoTo AddPick:
            If .FindItem(txtCode.Text, , , 0) = 0 Then
                frm_taricmain.MousePointer = 0
                Exit Sub
AddPick:
                .Sorted = False
                .ListItems.Add , , txtCode.Text
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                .ListItems(.ListItems.Count).Selected = True
                .Sorted = True
                For Counter = 1 To .ListItems.Count
                    If .ListItems(Counter).Text = txtCode.Text Then
                        .ListItems(Counter).EnsureVisible
                        .ListItems(Counter).Selected = True
                    End If
                Next Counter
            Else
                '-----> If record exist in listview delete it and add the new one
                .Sorted = False
                .ListItems.Remove (.FindItem(txtCode.Text, , , 0).Index)
                .ListItems.Add , , txtCode.Text
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempKey
                .ListItems(.ListItems.Count).ListSubItems.Add , , strTempDesc
                .ListItems(.ListItems.Count).Selected = True
                .Sorted = True
                For Counter = 1 To .ListItems.Count
                    If .ListItems(Counter).Text = txtCode.Text Then
                        .ListItems(Counter).EnsureVisible
                        .ListItems(Counter).Selected = True
                    End If
                Next Counter
            End If
        End With
        End If
    End If
    frm_taricmain.MousePointer = 0
    Screen.MousePointer = 0
    Unload Me

End Sub

Private Sub cmdReg_Click(Index As Integer)
    Dim strBoxProp As String
    Dim strPickVal As String
    Dim Admin As String
    
    Admin = GetSetting(App.Title, "Settings", "AdminRights")
    '----> Use default string used in codisheet
    strBoxProp = "R" & (1 + (Index * 2)) & "#*#D" & (26 + (Index * 2)) & "#*#1#*#" & _
    strLangOfDesc & "#*#" & "Import#*#" & Admin & "#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    
    blnTaricMain = True
    g_blnMultiplePick = False
    frm_picklist.Show vbModal, Me
    blnTaricMain = False
    
    strPickVal = GetSetting(App.Title, "CodiSheet", "Pick_TARIC_PL")
    '-----> get the regime
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then _
        txtReg(Index).Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
        
    DeleteSetting App.Title, "CodiSheet", "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"

End Sub

Private Sub cmdType_Click(Index As Integer)
    Dim strBoxProp As String
    Dim strPickVal As String
    Dim Admin As String
    
    Admin = GetSetting(App.Title, "Settings", "AdminRights")
    '----> Use default string used in codisheet
    strBoxProp = "N" & (1 + Index) & "#*#D" & (14 + (Index * 4)) & "#*#1#*#" & strLangOfDesc & _
    "#*#" & "Import#*#" & Admin & "#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    
    blnTaricMain = True
    g_blnMultiplePick = False
    frm_picklist.Show vbModal, Me
    blnTaricMain = False
    
    strPickVal = GetSetting(App.Title, "CodiSheet", "Pick_TARIC_PL")
    '-----> get the type
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then _
        txtType(Index).Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
        
    DeleteSetting App.Title, "CodiSheet", "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"

End Sub

Private Sub Form_Initialize()
    Set colTaricCodes = New Collection
End Sub

Private Sub Form_Load()

    '-----> To convert captions to default language
    Call LoadResStrings(Me, True)
    
    Dim Counter As Integer
    
    blnCopy = False
    '----> to skip procedure in picklist
    blnTaricMain = False
    
    '-----> initialize kluwer botton
    intThirdPartyDatabase = IIf(GetSetting(App.Title, "Third Party Database", "Usage") = "", 0, GetSetting(App.Title, "Third Party Database", "Usage"))
    Select Case intThirdPartyDatabase
        Case 0
            cmdGeneral(1).Enabled = False
        Case 1
            cmdGeneral(1).Enabled = True
            cmdGeneral(1).Caption = "TARBEL..."
        Case 2
            cmdGeneral(1).Enabled = True
            cmdGeneral(1).Caption = "Kluwer..."
    End Select
    
    '-----> get string values main loaded and main type
    '-----> "mainloaded/maintype"
    strMainLoaded = Left(gstrTaricMainCallType, InStr(gstrTaricMainCallType, "/") - 1)
    If Left(strMainLoaded, 7) = "AddFull" Then strMainLoaded = "AddFull"
    
    '-----> Get string main loaded
    If Right(gstrTaricMainCallType, Len(gstrTaricMainCallType) - InStr(gstrTaricMainCallType, "/")) = "frm_taricpicklist" Then
        strMainType = "Simplified"
        '-----> Language Used
        strLangOfDesc = frm_taricpicklist.strLangOfDesc
        '---->Document type (import or export)
        strDocType = frm_taricpicklist.strDocType
        '-----> if addfull; should be able to determine T or E
        If (strDocType = "Transit" Or strDocType = "Export" Or strDocType = G_CONST_NCTS1_TYPE Or strDocType = G_CONST_NCTS2_TYPE Or strDocType = G_CONST_EDINCTS1_TYPE) And strMainLoaded = "AddFull" Then _
                        gstrTaricMainCallType = gstrTaricMainCallType & Left(strDocType, 1)
        '----> no taricmain for transit defualt transit to export
        If strDocType = "Transit" Then strDocType = "Export"
    ElseIf Right(gstrTaricMainCallType, Len(gstrTaricMainCallType) - InStr(gstrTaricMainCallType, "/")) _
        = "frm_taricmaintenance" Then
         strMainType = "Normal"
        '-----> Language used
        strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
    
    ElseIf Right(gstrTaricMainCallType, Len(gstrTaricMainCallType) - InStr(gstrTaricMainCallType, "/")) _
        = "frm_codsheet" Then
        '----->load from codisheet
        strMainType = "Simplified"
        strDocType = "Import"
        strMainLoaded = "AddFull"
        Select Case newform(Right(Left(gstrTaricMainCallType, 8), 1)).colTaricArgs("A5")
            Case "N", "n"
                strLangOfDesc = "Dutch"
            Case "F", "f"
                strLangOfDesc = "French"
            Case Else
                strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
        End Select
        'strLangofDesc = frm_taricpicklist.strLangofDesc
    Else
        '-----> Load from codexport
        strMainType = "Simplified"
        strDocType = "Export"
        strMainLoaded = "AddFull"
        If Right(gstrTaricMainCallType, 1) = "E" Then
            Select Case newformE(Right(Left(gstrTaricMainCallType, 8), 1)).colTaricArgs("A5")
                Case "N", "n"
                    strLangOfDesc = "Dutch"
                Case "F", "f"
                    strLangOfDesc = "French"
                Case Else
                    strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
            End Select
        Else
            Select Case newformT(Right(Left(gstrTaricMainCallType, 8), 1)).colTaricArgs("A5")
                Case "N", "n"
                    strLangOfDesc = "Dutch"
                Case "F", "f"
                    strLangOfDesc = "French"
                Case Else
                    strLangOfDesc = IIf(cLanguage = "French", "French", "Dutch")
            End Select
        End If
    End If

    ADOConnectDB m_conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    'OpenDAODatabase m_conSADBEL, cAppPath, "mdb_sadbel.mdb"
    
    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTaric, cAppPath, "mdb_taric.mdb"
    
    Dim strSQL As String
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[PICKLIST MAINTENANCE " & strLangOfDesc & "] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[INTERNAL CODE] = '8.29801619052887E+19' "
        strSQL = strSQL & "OR "
        strSQL = strSQL & "[INTERNAL CODE] = '7.67111659049988E+19' "
        strSQL = strSQL & "OR "
        strSQL = strSQL & "[INTERNAL CODE] = '5.35045266151428E+18' "
    ADORecordsetOpen strSQL, m_conSADBEL, m_rstPick, adOpenKeyset, adLockOptimistic
    'Set m_rstPick = m_conSADBEL.OpenRecordset(strSQL)
    
    '-----> recordset for supplementary values
    ADORecordsetOpen "Select * from [SUPP UNITS]", m_conTaric, m_rstSupp, adOpenKeyset, adLockOptimistic
    'Set m_rstSupp = m_conTaric.OpenRecordset("Select * from [SUPP UNITS]")
    
    '-----> fill combo boxes with values
    FillCombo
    
    '-----> input default values for combo boxes
    If strLangOfDesc = "English" Then
        cboSuppStat.Text = "00-None"
        cboSuppCalc.Text = "00-None"
        cboGrosCalc.Text = "00-None"
    ElseIf strLangOfDesc = "Dutch" Then
        cboSuppStat.Text = "00-Geen"
        cboSuppCalc.Text = "00-Geen"
        cboGrosCalc.Text = "00-Geen"
    ElseIf strLangOfDesc = "French" Then
        cboSuppStat.Text = "00-Aucun"
        cboSuppCalc.Text = "00-Aucun"
        cboGrosCalc.Text = "00-Aucun"
    End If
    
    '----->Check if loaded as main or Simplified main
    If strMainType = "Simplified" Then '-----> Simplified
        
        frmSettings.Visible = True
        
        '-----> Check if called from import or export
        '-----> Load default values for either import or export
        If strDocType = "Import" Then
            'frm_taricmain.Caption = "TARIC - Add/Modify Import Codes"
            frmImport.Visible = True
    '        cboCurrency.AddItem "BEF"
    '        cboCurrency.AddItem "LUF"
    '        cboCurrency.AddItem "EUR"
            '-----> get minimum value and currency from database
            Dim rstProperties As ADODB.Recordset
            ADORecordsetOpen "Select * from Properties", m_conTaric, rstProperties, adOpenKeyset, adLockOptimistic
            'Set rstProperties = m_conTaric.OpenRecordset("Select * from Properties")
            With rstProperties
                If Not (.EOF And .BOF) Then
                    .MoveFirst
    '                strLicCurr = ![MIN VALUE CURR]
                    'Mod by BCo
                    'Added Val() to convert string to number, w/ consideration to regional formatting.
                    dblMinVal = Val(![Min Lic Value])
                Else
    '                strLicCurr = "BEF"
    '                dblMinVal = 5000
                    dblMinVal = 5000 / 40.3399
                End If
            End With
            
            ADORecordsetClose rstProperties
            
        ElseIf strDocType = "Export" Or strDocType = G_CONST_NCTS1_TYPE Or strDocType = G_CONST_NCTS2_TYPE Or strDocType = G_CONST_EDINCTS1_TYPE Then
            'frm_taricmain.Caption = "TARIC - Add/MOdify Export Codes"
            frmExport.Visible = True
        End If
    
        '-----> Check to what type; Modify, Add or Copy
        Select Case strMainLoaded
            Case "AddBlank" '----->Leave everything blank
                Call BlankOut
                blnStartApply = True
                Exit Sub
            Case "AddFull" '-----> Get values from the codisheet
                Call FillFromCodi
                blnStartApply = True
                Exit Sub
            Case "Modify" '-----> Get code value form Picklist. Fill all fields
                txtCode.Text = frm_taricpicklist.lvwPicklist.SelectedItem.Text
                txtCtryCode = frm_taricpicklist.strCtryCode
                
                If UCase(gstrTaricMainCallType) = "MODIFY/FRM_TARICPICKLIST" Then
                    LoadValuesFromDB GetTableToUse(strDocType), txtCode.Text
                End If
            Case "Copy" '-----> Except for taric code, fill all fields using the fields of the taric code from picklist.
                txtCode.Text = frm_taricpicklist.lvwPicklist.SelectedItem.Text
                txtCtryCode.Text = frm_taricpicklist.strCtryCode
                blnCopy = True
                txtCode.Text = ""
            Case "View" '-----> No rights to modify and mantain tables
                strMainLoaded = "Modify"
                '-----> get taric code from frm_taricpicklist
                txtCode.Text = frm_taricpicklist.lvwPicklist.SelectedItem.Text
                txtCtryCode.Text = frm_taricpicklist.strCtryCode
                txtCode.Enabled = False
                Dim ctl As Control
                For Each ctl In frm_taricmain.Controls
                    ctl.Enabled = False
                Next
                cmdCancel.Enabled = True
                For Counter = 0 To 13
                    Label1(Counter).Enabled = True
                Next
                For Counter = 0 To 3
                    Frame7(Counter).Enabled = True
                Next
                frmImport.Enabled = True
                frmExport.Enabled = True
                frmGeneral.Enabled = True
                frmQuantities.Enabled = True
                frmSettings.Enabled = True
            End Select
    
    Else '----->Loaded as normal main form
        frmSettings.Visible = False
        'Me.Caption = "TARIC - Add/Modify Codes"
        
        lvwImport.ColumnHeaders(1).Text = Translate(839)
        lvwExport.ColumnHeaders(1).Text = Translate(899)
        '-----> Check to what type; Modify, Add or Copy
        Select Case strMainLoaded
            Case "Add" '----->Leave everything blank
                Exit Sub
            Case "Modify" '-----> Get code value form Picklist. Fill all fields
                txtCode.Text = frm_taricmaintenance.lvwMaintenance.SelectedItem.Text
            Case "Copy" '-----> Except for code, fill all fields using the fields of the code from picklist.
                txtCode.Text = frm_taricmaintenance.lvwMaintenance.SelectedItem.Text
                blnCopy = True
                txtCode.Text = ""
        End Select
    End If
    blnStartApply = True
    Me.MousePointer = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        If blnApplyClicked = False Then
            gblnFormWasCanceled = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set colTaricCodes = Nothing
    
    ADORecordsetClose m_rstMain
    ADORecordsetClose m_rstPick
    ADORecordsetClose m_rstSupp
    
    If strMainType = "Normal" Then
        ADORecordsetClose m_rstImpGrid
        ADORecordsetClose m_rstExpGrid
    ElseIf strMainType = "Simplified" Then
        ADORecordsetClose m_rstDetail
    End If
    
    ADORecordsetClose m_rstCommon
    
    ADODisconnectDB m_conTaric
    ADODisconnectDB m_conSADBEL

    UnloadControls Me
    
    Set frm_taricmain = Nothing

End Sub

Private Sub lvwExport_Click()

    If blnItemExport Then
        blnItemExport = False
    Else
        cmdExport(1).Enabled = False
        cmdExport(2).Enabled = False
        cmdExport(3).Enabled = False
        cmdExport(4).Enabled = False
        Dim Counter As Integer
        For Counter = 1 To lvwExport.ListItems.Count
            lvwExport.ListItems(Counter).Selected = False
        Next Counter
    End If

End Sub

Private Sub lvwExport_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Item.Index <> 0 Then
        cmdExport(1).Enabled = True '-----> Delete
        cmdExport(2).Enabled = True '-----> Make a copy
        cmdExport(3).Enabled = True '-----> Change Settings
        blnItemExport = True
    End If
    
    '-----> Only enabled if selected country is not the default country
    If Item.Bold = False Then
        cmdExport(4).Enabled = True '-----> Set as default
    Else
        cmdExport(4).Enabled = False
    End If

End Sub


Private Sub lvwImport_Click()

    If blnItemImport Then
        blnItemImport = False
    Else:
        cmdImport(1).Enabled = False
        cmdImport(2).Enabled = False
        cmdImport(3).Enabled = False
        cmdImport(4).Enabled = False
        Dim Counter As Integer
        For Counter = 1 To lvwImport.ListItems.Count
        lvwImport.ListItems(Counter).Selected = False
        Next Counter
    End If

End Sub

Private Sub lvwImport_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Item.Index <> 0 Then
        cmdImport(1).Enabled = True '-----> Delete
        cmdImport(2).Enabled = True '-----> Make a Copy
        cmdImport(3).Enabled = True '-----> Change Settings
        blnItemImport = True
    End If
    
    '-----> Only enabled if selected country is not the default country
    If Item.Bold = False Then
        cmdImport(4).Enabled = True '-----> Set as default
    Else
        cmdImport(4).Enabled = False
    End If

End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrHandler
    
    ' Activate Kluwer database application via its task ID.
    AppActivate mdblTaskID, True
    
    Exit Sub
    
ErrHandler:
    
    Select Case Err.Number
        Case 5    ' Invalid procedure call or argument; occurs when Kluwer was already closed.
            If Clipboard.GetFormat(vbCFText) Then
                If Clipboard.GetText <> "No Data" Then
                    Call GetFromKluwer
                End If
            End If
        Case Else
            If MsgBox(Err.Description, vbRetryCancel + vbExclamation, Err.Source & " (" & Err.Number & ")") = vbRetry Then
                Resume
            End If
    End Select
    
    ' Re-initialize mdblTaskID if Kluwer was already closed.
    mdblTaskID = 0
    Timer1.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtCode_Change()

    Dim strKeywordD As String
    Dim strKeywordF As String
    Dim intCopy As Integer
    Dim Counter As Integer
    Dim strCN As String

    frm_taricmain.MousePointer = 11
    
    If Len(Trim(txtCode.Text)) < 10 Then
        '-----> to enable ok and apply at simplified form
        blnSaveOK = False
    
        cmdGeneral(2).Enabled = False '-----> disable command clients
        cmdImport(0).Enabled = False '-----> disable import new country
        cmdExport(0).Enabled = False '-----> disable export new country
        
        cmdApply.Enabled = False '-----> saving should not be allowed
        txtFrnchDesc.Enabled = False
        txtFrnchKey.Enabled = False
        txtDutchDesc.Enabled = False
        txtDutchKey.Enabled = False
        frmQuantities.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        
        '-----> empty all data fields
        '-----> Clear listview import and export if input taric code < 10 digits
        If strMainLoaded = "AddBlank" Or strMainLoaded = "AddFull" Or strMainLoaded = "Modify" Then
            If strMainType = "Normal" Then
                lvwImport.ListItems.Clear
                lvwExport.ListItems.Clear
            ElseIf strMainType = "Simplified" Then
                txtCtryCode.Text = ""
                frmSettings.Enabled = False
            End If
        End If
    
        '-----> delete saved kluwer to database
        If blnKluwerClicked = True Then
            Call DeleteKluwer
            blnKluwerClicked = False
        End If
    End If
    
    ' ************* CHECK CN CODE: 8 CHARACTERS *************
    If Len(Trim(txtCode.Text)) = 8 Or Len(Trim(txtCode.Text)) = 9 Then '-----> compare first 8 digits of cn code to taric code
        
    If strMainLoaded = "Modify" Or strMainLoaded = "AddFull" Or strMainLoaded = "AddBlank" Then
        '-----> set combo boxes to default values
        If strLangOfDesc = "English" Then
            cboSuppStat.Text = "00-None"
            cboSuppCalc.Text = "00-None"
            cboGrosCalc.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppStat.Text = "00-Geen"
                cboSuppCalc.Text = "00-Geen"
                cboGrosCalc.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppStat.Text = "00-Aucun"
                cboSuppCalc.Text = "00-Aucun"
                cboGrosCalc.Text = "00-Aucun"
        End If
        chkSuppStat.Value = 0
        chkSuppCalc.Value = 0
        chkGrossCalc.Value = 0
    End If
        
            '----->source from CN table check if code exist in table
                'allanSQL
                strCN = vbNullString
                strCN = strCN & "SELECT "
                strCN = strCN & "* "
                strCN = strCN & "FROM "
                strCN = strCN & "CN "
                strCN = strCN & "WHERE "
                strCN = strCN & "[CN CODE] = " & Chr(39) & ProcessQuotes(Left(txtCode.Text, 8)) & Chr(39) & " "
            ADORecordsetOpen strCN, m_conTaric, m_rstMain, adOpenKeyset, adLockOptimistic
            'Set m_rstMain = m_conTaric.OpenRecordset(strCN, dbOpenForwardOnly)
            With m_rstMain
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    
                    txtDutchDesc.Text = ![DESC DUTCH]
                    txtFrnchDesc.Text = ![DESC FRENCH]
                    
                    '----->convert supplementary statistical ascii code to sadbel code
                    If ![SUPP STAT UNIT] = "-" Then
                        If strLangOfDesc = "English" Then
                            cboSuppStat.Text = "00-None"
                        ElseIf strLangOfDesc = "Dutch" Then
                            cboSuppStat.Text = "00-Geen"
                        ElseIf strLangOfDesc = "French" Then
                            cboSuppStat.Text = "00-Aucun"
                        End If
                    End If
                    
                    SBcode = 0
                    ASCIIConverter (![SUPP STAT UNIT])
                    
                    If Not (m_rstPick.EOF And m_rstPick.BOF) Then
                        m_rstPick.MoveFirst
                        Do While Not m_rstPick.EOF
                            '----->Supplementary Statistical Unit
                            If m_rstPick![Internal Code] = "7.67111659049988E+19" And _
                                SBcode = m_rstPick![code] Then
                                
                                If strLangOfDesc = "Dutch" Then
                                    cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION DUTCH]
                                ElseIf strLangOfDesc = "French" Then
                                    cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION FRENCH]
                                End If
                            End If
                            
                            m_rstPick.MoveNext
                        Loop
                    End If
                End If
            End With
         
        '-----> Get the first word which is more than three characters
        'capitalized and set as keyword
        
            strKeywordD = txtDutchDesc.Text
Ulit:
            If UCase(Left(strKeywordD, 5)) = "ANDER" Or UCase(Left(strKeywordD, 6)) = "ANDER," Or _
                UCase(Left(strKeywordD, 6)) = "ANDERE" Or UCase(Left(strKeywordD, 7)) = "ANDERE," Or _
                UCase(Left(strKeywordD, 7)) = "NUMBERS" Or UCase(Left(strKeywordD, 8)) = "NUMBERS," Or _
                UCase(Left(strKeywordD, 5)) = "DELEN" Or UCase(Left(strKeywordD, 6)) = "DELEN," Or _
                UCase(Left(strKeywordD, 9)) = "GEBRUIKTE" Or UCase(Left(strKeywordD, 10)) = "GEBRUIKTE," Or _
                UCase(Left(strKeywordD, 7)) = "GETUFTE" Or UCase(Left(strKeywordD, 8)) = "GETUFTE," Or _
                UCase(Left(strKeywordD, 7)) = "GEVULDE" Or UCase(Left(strKeywordD, 8)) = "GEVULDE," Or _
                UCase(Left(strKeywordD, 7)) = "GEZAAGD" Or UCase(Left(strKeywordD, 8)) = "GEZAAGD," Or _
                UCase(Left(strKeywordD, 5)) = "STUKS" Or UCase(Left(strKeywordD, 6)) = "STUKS," Then
                
                strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                    
                If Not InStr(strKeywordD, " ") = 0 Then
                    GoTo Ulit:
                End If
            End If
            
            If Not strKeywordD = "" And Not InStr(strKeywordD, " ") = 0 Then
                If IsNumeric(Left(strKeywordD, InStr(strKeywordD, " ") - 1)) Then
                    strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                    
                    If Not InStr(strKeywordD, " ") = 0 Then
                        GoTo Ulit:
                    End If
                End If
            End If
            
            If InStr(strKeywordD, " ") > 4 Then
                txtDutchKey.Text = UCase(Left(strKeywordD, InStr(strKeywordD, " ") - 1))
                If Right(txtDutchKey.Text, 1) = "," Then
                    txtDutchKey.Text = Left(txtDutchKey.Text, Len(txtDutchKey.Text) - 1)
                End If
            ElseIf InStr(strKeywordD, " ") = 0 Then
                txtDutchKey.Text = UCase(strKeywordD)
                If Right(txtDutchKey.Text, 1) = "," Then
                    txtDutchKey.Text = Left(txtDutchKey.Text, Len(txtDutchKey.Text) - 1)
                End If
            Else
                strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                GoTo Ulit:
            End If
     
                strKeywordF = txtFrnchDesc.Text
Ulit2:
            If UCase(Left(strKeywordF, 5)) = "AUTRE" Or UCase(Left(strKeywordF, 6)) = "AUTRE," Or _
                UCase(Left(strKeywordF, 6)) = "AUTRES" Or UCase(Left(strKeywordF, 7)) = "AUTRES," Or _
                UCase(Left(strKeywordF, 7)) = "NUMBERS" Or UCase(Left(strKeywordF, 8)) = "NUMBERS," Or _
                UCase(Left(strKeywordF, 9)) = "DELEN VAN" Or UCase(Left(strKeywordF, 10)) = "DELEN VAN," Then
                
                strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                
                If Not InStr(strKeywordF, " ") = 0 Then
                    GoTo Ulit2:
                End If
            End If
            
            If Not strKeywordF = "" And Not InStr(strKeywordF, " ") = 0 Then
                If IsNumeric(Left(strKeywordF, InStr(strKeywordF, " ") - 1)) Then
                    strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                    
                    If Not InStr(strKeywordF, " ") = 0 Then
                        GoTo Ulit2:
                    End If
                End If
            End If
            
            If InStr(strKeywordF, " ") > 4 Then
                txtFrnchKey.Text = UCase(Left(strKeywordF, InStr(strKeywordF, " ") - 1))
                If Right(txtFrnchKey.Text, 1) = "," Then
                    txtFrnchKey.Text = Left(txtFrnchKey.Text, Len(txtFrnchKey.Text) - 1)
                End If
            ElseIf InStr(strKeywordF, " ") = 0 Then
                txtFrnchKey.Text = UCase(strKeywordF)
                If Right(txtFrnchKey.Text, 1) = "," Then
                    txtFrnchKey.Text = Left(txtFrnchKey.Text, Len(txtFrnchKey.Text) - 1)
                End If
            Else
                strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                GoTo Ulit2:
            End If
         
    '-----> if input taric code falls below 8 digits clear general data
    ElseIf Len(Trim(txtCode.Text)) = 7 Then
        If strMainLoaded = "AddFull" Or strMainLoaded = "AddBlank" Or strMainLoaded = "Modify" Then
            txtDutchKey.Text = ""
            txtFrnchKey.Text = ""
            txtDutchDesc.Text = ""
            txtFrnchDesc.Text = ""
            If strLangOfDesc = "English" Then
                cboSuppStat.Text = "00-None"
            ElseIf strLangOfDesc = "Dutch" Then
                cboSuppStat.Text = "00-Geen"
            ElseIf strLangOfDesc = "French" Then
                cboSuppStat.Text = "00-Aucun"
            End If
        End If
    End If
    
    
    '************* CHECK TARIC CODE: 10 CHARACTERS **************
    
    If Len(Trim(txtCode.Text)) = 10 Then
                
        cmdGeneral(2).Enabled = True '-----> enable command clients
        cmdImport(0).Enabled = True '-----> enable import new country
        cmdExport(0).Enabled = True '-----> enable export new country
        If strMainType = "Normal" Then
            cmdOK.Enabled = True
        End If
        
        frmQuantities.Enabled = True
        txtFrnchDesc.Enabled = True
        txtFrnchKey.Enabled = True
        txtDutchDesc.Enabled = True
        txtDutchKey.Enabled = True
        frmQuantities.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
    
        '-----> to enable ok and apply at simplified form
        blnSaveOK = True
    
    On Error GoTo NoRecord: '-----> Error occurs when there is no data on the table
        
        '-----> Open recordset rst common with the corresponding txtcode
        Dim strSQL As String
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "COMMON "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
        ADORecordsetOpen strSQL, m_conTaric, m_rstCommon, adOpenKeyset, adLockOptimistic
        'Set m_rstCommon = m_conTaric.OpenRecordset(strSQL)
        
        '-----> fill form with data that matches the taric code
        With m_rstCommon
            If Not (.EOF And .BOF) Then
            
                .MoveFirst
                Do While Not .EOF
                    If Trim(txtCode.Text) = ![TARIC CODE] Then
                        
                        '----> if record already exist warn the user
                        If blnCopy = True And strMainLoaded = "Copy" Then
                            intCopy = MsgBox(Translate(464) & "?", vbYesNoCancel + vbQuestion, Me.Caption)
                            If intCopy = 2 Then '-----> Cancel
                                txtCode.Text = ""
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            ElseIf intCopy = 6 Then '-----> Yes
                                frm_taricmain.MousePointer = 0
                                Exit Sub
                            ElseIf intCopy = 7 Then '-----> No
                                strMainLoaded = "Modify"
                            End If
                        End If
                        
                        '-----> Fill in general data
                        If Not IsNull(![KEY DUTCH]) Then txtDutchKey.Text = ![KEY DUTCH]
                        If Not IsNull(![DESC DUTCH]) Then txtDutchDesc.Text = ![DESC DUTCH]
                        If Not IsNull(![KEY FRENCH]) Then txtFrnchKey.Text = ![KEY FRENCH]
                        If Not IsNull(![DESC FRENCH]) Then txtFrnchDesc.Text = ![DESC FRENCH]
                        
                        '-----> Lock Qualities
                        If ![SUPP STAT LOCK CODE] = -1 Then
                            chkSuppStat.Value = 1
                        Else
                            chkSuppStat.Value = 0
                        End If
                        
                        If ![SUPP CALC LOCK CODE] = -1 Then
                            chkSuppCalc.Value = 1
                        Else
                            chkSuppCalc.Value = 0
                        End If
                        If ![GROSS WT LOCK CODE] = -1 Then
                            chkGrossCalc.Value = 1
                        Else
                            chkGrossCalc.Value = 0
                        End If
                        
                        '-----> Fill in Quantities
                        If Not (m_rstPick.EOF And m_rstPick.BOF) Then
                            m_rstPick.MoveFirst
                            Do While Not m_rstPick.EOF
                                '----->Supplementary Statistical Unit
                                If m_rstPick![Internal Code] = "7.67111659049988E+19" And ![SUPP STAT UNIT] = m_rstPick![code] Then
                                    If strLangOfDesc = "Dutch" Then
                                        cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION FRENCH]
                                    End If
                                End If
                                '----->Supplementary Calculation Unit
                                If m_rstPick![Internal Code] = "5.35045266151428E+18" And ![SUPP CALC UNIT] = m_rstPick![code] Then
                                    If strLangOfDesc = "Dutch" Then
                                        cboSuppCalc.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION DUTCH]
                                    ElseIf strLangOfDesc = "French" Then cboSuppCalc.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION FRENCH]
                                    End If
                                End If
                                m_rstPick.MoveNext
                            Loop
                        End If
                        
                        '-----> Supp Stat is none
                        If Left(cboSuppStat.Text, 2) = "00" Then GoTo SkipSuppStatQ:
                        
                        '-----> Supplementary Statistical Unit Quantity Handling
                        If Not (m_rstSupp.EOF And m_rstSupp.BOF) Then
                            m_rstSupp.MoveFirst
                            For Counter = 1 To 11
                                If m_rstSupp![SUPP QTY CODE] = ![SUPP STAT QTY CODE] Then
                                    If strLangOfDesc = "English" Then
                                        cboSuppStatQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC ENGLISH]
                                    ElseIf strLangOfDesc = "Dutch" Then
                                        cboSuppStatQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        cboSuppStatQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC FRENCH]
                                    End If
                                End If
                                
                                m_rstSupp.MoveNext
                            Next Counter
                        End If
                        
SkipSuppStatQ:
                        '-----> Supp Calc is none
                        If Left(cboSuppCalc.Text, 2) = "00" Then GoTo SkipSuppCalcQ:
                        
                        '-----> Supplementary Calculation Unit Quantity Handling
                        If Not (m_rstSupp.EOF And m_rstSupp.BOF) Then
                            m_rstSupp.MoveFirst
                            For Counter = 1 To 11
                                If m_rstSupp![SUPP QTY CODE] = ![SUPP CALC QTY CODE] Then
                                    If strLangOfDesc = "English" Then
                                        cboSuppCalcQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC ENGLISH]
                                    ElseIf strLangOfDesc = "Dutch" Then
                                        cboSuppCalcQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        cboSuppCalcQ.Text = m_rstSupp![SUPP QTY CODE] & "-" & m_rstSupp![DESC FRENCH]
                                    End If
                                End If
                                m_rstSupp.MoveNext
                            Next Counter
                        End If
SkipSuppCalcQ:
                        
                        '-----> Gross Weight Calculation
                        If Not (m_rstSupp.EOF And m_rstSupp.BOF) Then
                            m_rstSupp.MoveFirst
                            m_rstSupp.Move (12)
                            For Counter = 1 To 8
                                If m_rstSupp![GROSS WT CALC CODE] = ![GROSS WT CALC CODE] Then
                                    If strLangOfDesc = "English" Then
                                        cboGrosCalc.Text = m_rstSupp![GROSS WT CALC CODE] & "-" & m_rstSupp![DESC ENGLISH]
                                    ElseIf strLangOfDesc = "Dutch" Then
                                        cboGrosCalc.Text = m_rstSupp![GROSS WT CALC CODE] & "-" & m_rstSupp![DESC DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        cboGrosCalc.Text = m_rstSupp![GROSS WT CALC CODE] & "-" & m_rstSupp![DESC FRENCH]
                                    End If
                                End If
                                m_rstSupp.MoveNext
                            Next Counter
                        End If
                    End If
                    
                    .MoveNext
                Loop
            End If
        End With
        
NoRecord:
        If blnCopy = False Then GoTo InitialCopy:
        
        If strMainLoaded = "Modify" Or strMainLoaded = "AddFull" Or strMainLoaded = "AddBlank" Then
        '-----> If form is loaded as normal; called from taric picklist
InitialCopy:
            If strMainType = "Normal" Then
                
                '-----> Fill in Import Grid with Data
                '-----> If main is loaded as Normal
                ImportAdd
                ExportAdd
            Else
                frmSettings.Enabled = True
                DetailData
            End If
        End If
    End If
NotIncn:
    ADORecordsetClose m_rstMain
            
    frm_taricmain.MousePointer = 0


End Sub

Public Sub FillCombo()
    '-----> for cbosuppstat and cbosuppcalc (Supplementary Statistical Unit and Supplementary Calculation Unit)
    Dim Counter As Integer
    
    With m_rstPick
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                '-----> internal code for Supplementary Statistical Unit 7.67111659049988E+19
                If ![Internal Code] = "7.67111659049988E+19" Then
                    If strLangOfDesc = "Dutch" Then
                        cboSuppStat.AddItem ![code] & "-" & ![DESCRIPTION DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboSuppStat.AddItem ![code] & "-" & ![DESCRIPTION FRENCH]
                    End If
                End If
                '-----> Imnternal Code for supplementary Calculation Unit 5.35045266151428E+18
                If ![Internal Code] = "5.35045266151428E+18" Then
                    If strLangOfDesc = "Dutch" Then
                        cboSuppCalc.AddItem ![code] & "-" & ![DESCRIPTION DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboSuppCalc.AddItem ![code] & "-" & ![DESCRIPTION FRENCH]
                    End If
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
    If strLangOfDesc = "English" Then
        cboSuppStat.AddItem "00-None"
        cboSuppCalc.AddItem "00-None"
    ElseIf strLangOfDesc = "Dutch" Then
        cboSuppStat.AddItem "00-Geen"
        cboSuppCalc.AddItem "00-Geen"
    ElseIf strLangOfDesc = "French" Then
        cboSuppStat.AddItem "00-Aucun"
        cboSuppCalc.AddItem "00-Aucun"
    End If
    
    '-----> for cbosuppstatq (Supplementary Statistical Unit Quantity Handling)
    '-----> to be modified when desc in dutch and french are inputed to db
    Counter = 0
    With m_rstSupp
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not Counter = 11
                If Not IsNull(![SUPP QTY CODE]) Then
                    If strLangOfDesc = "English" Then
                        cboSuppStatQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC ENGLISH]
                    ElseIf strLangOfDesc = "Dutch" Then
                        cboSuppStatQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboSuppStatQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC FRENCH]
                    End If
                    Counter = Counter + 1
                End If
                
                .MoveNext
            Loop
        End If
    End With

    '-----> for cbosuppcalcq (Supplementary Calculation Unit Quantity Handling)
    '-----> to be modified when desc in dutch and french are inputed to db
    Counter = 0
    With m_rstSupp
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not Counter = 11
                If Not IsNull(![SUPP QTY CODE]) Then
                    If strLangOfDesc = "English" Then
                        cboSuppCalcQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC ENGLISH]
                    ElseIf strLangOfDesc = "Dutch" Then
                        cboSuppCalcQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboSuppCalcQ.AddItem ![SUPP QTY CODE] & "-" & ![DESC FRENCH]
                    End If
                    Counter = Counter + 1
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
    '-----> for cbogrosscalc (Gross Weight Calculation)
    '-----> to be modified when desc in dutch and french are inputed to db
    Counter = 0
    With m_rstSupp
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not Counter = 8
                If Not IsNull(![GROSS WT CALC CODE]) Then
                    If strLangOfDesc = "English" Then
                        cboGrosCalc.AddItem ![GROSS WT CALC CODE] & "-" & ![DESC ENGLISH]
                    ElseIf strLangOfDesc = "Dutch" Then
                        cboGrosCalc.AddItem ![GROSS WT CALC CODE] & "-" & ![DESC DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboGrosCalc.AddItem ![GROSS WT CALC CODE] & "-" & ![DESC FRENCH]
                    End If
                    
                    Counter = Counter + 1
                End If
                
                .MoveNext
            Loop
        End If
    End With

End Sub

Public Sub ImportAdd()

    '-----> Fill in Import Grid
    '-----> For Normal Main
    lvwImport.Sorted = False
    
    Dim strOrigin As String
    Dim strLic As String
    Dim strSQL As String
    Dim colEmptyFieldValues As Collection
    
    '----> Save empty box field values to colEmptyFieldValues
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "[BOX CODE], "
        strSQL = strSQL & "[EMPTY FIELD VALUE] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[BOX DEFAULT IMPORT ADMIN] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[BOX CODE] "
        strSQL = strSQL & "IN "
        strSQL = strSQL & "( "
        strSQL = strSQL & "'N1', 'N2', 'N3', "
        strSQL = strSQL & "'Q1', 'Q2', 'Q3', "
        strSQL = strSQL & "'R1', 'R2', 'R3', "
        strSQL = strSQL & "'R4' "
        strSQL = strSQL & ") "
    ADORecordsetOpen strSQL, m_conSADBEL, m_rstEmptyFieldValues, adOpenKeyset, adLockOptimistic
    'Set m_rstEmptyFieldValues = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
    
    Set colEmptyFieldValues = New Collection
    With m_rstEmptyFieldValues
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                colEmptyFieldValues.Add CStr(![EMPTY FIELD VALUE]), CStr(![BOX CODE])
                
                .MoveNext
            Loop
        End If
    End With
    ADORecordsetClose m_rstEmptyFieldValues
        
    lvwImport.ListItems.Clear
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "IMPORT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, m_rstImpGrid, adOpenKeyset, adLockOptimistic
    'Set m_rstImpGrid = m_conTaric.OpenRecordset(strSQL)
    
    With m_rstImpGrid
        On Error GoTo NoRec:
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not m_rstImpGrid.EOF
                If Trim(txtCode.Text) = ![TARIC CODE] Then
                    '-----> Get Country of Origin
                    strOrigin = ""
                    
                    If Not (m_rstPick.EOF And m_rstPick.BOF) Then
                    
                        m_rstPick.MoveFirst
                        Do Until m_rstPick.EOF
                            If m_rstPick![Internal Code] = "8.29801619052887E+19" Then
                                If ![CTRY CODE] = m_rstPick![code] Then
                                    If strLangOfDesc = "Dutch" Then
                                        strOrigin = ![CTRY CODE] & " - " & m_rstPick![DESCRIPTION DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        strOrigin = ![CTRY CODE] & " - " & m_rstPick![DESCRIPTION FRENCH]
                                    End If
                                End If
                            End If
                            
                            m_rstPick.MoveNext
                        Loop
                    End If
                    
                    If Len(strOrigin) = 0 Then
                        strOrigin = ![CTRY CODE] & " "
                    End If
                    
                    lvwImport.ListItems.Add , , strOrigin
                    
                    If ![LIC REQD] = -1 Then
                        strLic = "Y"
                    Else
                        strLic = "N"
                    End If
                    
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , strLic
                    
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N1]), colEmptyFieldValues("N1"), ![N1])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q1]), colEmptyFieldValues("Q1"), ![Q1])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N2]), colEmptyFieldValues("N2"), ![N2])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q2]), colEmptyFieldValues("Q2"), ![Q2])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N3]), colEmptyFieldValues("N3"), ![N3])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q3]), colEmptyFieldValues("Q3"), ![Q3])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R1]), colEmptyFieldValues("R1"), ![R1])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R2]), colEmptyFieldValues("R2"), ![R2])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R3]), colEmptyFieldValues("R3"), ![R3])
                    lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R4]), colEmptyFieldValues("R4"), ![R4])
                    
                    '-----> Check for the default country and set it as bold
                    If ![DEF CODE] = -1 Then
                        lvwImport.ListItems(lvwImport.ListItems.Count).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(1).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(2).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(3).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(4).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(5).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(6).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(7).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(8).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(9).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(10).Bold = True
                        lvwImport.ListItems(lvwImport.ListItems.Count).ListSubItems(11).Bold = True
                    End If
                End If
                
                m_rstImpGrid.MoveNext
            Loop
        End If
        
        lvwImport.SortKey = lvwAscending
        lvwImport.Sorted = True
    End With

NoRec:
        If lvwImport.ListItems.Count < 1 Then
            cmdImport(1).Enabled = False
            cmdImport(2).Enabled = False
            cmdImport(3).Enabled = False
            cmdImport(4).Enabled = False
        End If
        
    Set colEmptyFieldValues = Nothing

End Sub

Public Sub ExportAdd()

    lvwExport.Sorted = False
    
    '-----> Fill in Export Grid
    '-----> For Normal Main
    
    Dim strOrigin As String
    Dim strLic As String
    Dim strSQL As String
    Dim colEmptyFieldValues As Collection
    
    '----> Save empty box field values to colEmptyFieldValues
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "[BOX CODE], "
        strSQL = strSQL & "[EMPTY FIELD VALUE] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[BOX DEFAULT EXPORT ADMIN] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[BOX CODE] "
        strSQL = strSQL & "IN "
        strSQL = strSQL & "( "
        strSQL = strSQL & "'N1', 'N2', 'N3', "
        strSQL = strSQL & "'Q1', 'Q2', 'Q3', "
        strSQL = strSQL & "'R1', 'R2', 'R3', "
        strSQL = strSQL & "'R4' "
        strSQL = strSQL & ") "
    ADORecordsetOpen strSQL, m_conSADBEL, m_rstEmptyFieldValues, adOpenKeyset, adLockOptimistic
    'Set m_rstEmptyFieldValues = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
    
    Set colEmptyFieldValues = New Collection
    With m_rstEmptyFieldValues
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                colEmptyFieldValues.Add CStr(![EMPTY FIELD VALUE]), CStr(![BOX CODE])
                
                .MoveNext
            Loop
        End If
    End With
    ADORecordsetClose m_rstEmptyFieldValues
    
    lvwExport.ListItems.Clear
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "EXPORT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, m_rstExpGrid, adOpenKeyset, adLockOptimistic
    'Set m_rstExpGrid = m_conTaric.OpenRecordset(strSQL)
    
    With m_rstExpGrid
        On Error GoTo NoRec:
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                If Trim(txtCode.Text) = ![TARIC CODE] Then
                    '-----> Get Country of Origin
                    strOrigin = ""
                    
                    If Not (m_rstPick.EOF And m_rstPick.BOF) Then
                        m_rstPick.MoveFirst
                        
                        Do Until m_rstPick.EOF
                            If m_rstPick![Internal Code] = "8.29801619052887E+19" Then '-----> Country code internal code
                                If ![CTRY CODE] = m_rstPick![code] Then
                                    If strLangOfDesc = "Dutch" Then
                                        strOrigin = ![CTRY CODE] & " - " & m_rstPick![DESCRIPTION DUTCH]
                                    ElseIf strLangOfDesc = "French" Then
                                        strOrigin = ![CTRY CODE] & " - " & m_rstPick![DESCRIPTION FRENCH]
                                    End If
                                End If
                            End If
                            
                            m_rstPick.MoveNext
                        Loop
                    End If
                    
                    If Len(strOrigin) = 0 Then
                        strOrigin = ![CTRY CODE] & " "
                    End If
                    
                    lvwExport.ListItems.Add , , strOrigin
                    
                    If ![LIC REQD] = -1 Then
                        strLic = "Y"
                    Else
                        strLic = "N"
                    End If
                    
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , strLic
                    
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N1]), colEmptyFieldValues("N1"), ![N1])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q1]), colEmptyFieldValues("Q1"), ![Q1])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N2]), colEmptyFieldValues("N2"), ![N2])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q2]), colEmptyFieldValues("Q2"), ![Q2])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![N3]), colEmptyFieldValues("N3"), ![N3])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![Q3]), colEmptyFieldValues("Q3"), ![Q3])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R1]), colEmptyFieldValues("R1"), ![R1])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R2]), colEmptyFieldValues("R2"), ![R2])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R3]), colEmptyFieldValues("R3"), ![R3])
                    lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems.Add , , IIf(IsNull(![R4]), colEmptyFieldValues("R4"), ![R4])
                
                    '-----> Check for the default country and set it as bold
                    If ![DEF CODE] = -1 Then
                        lvwExport.ListItems(lvwExport.ListItems.Count).Bold = True
                        Dim Counter As Integer
                        For Counter = 1 To 11
                            lvwExport.ListItems(lvwExport.ListItems.Count).ListSubItems(Counter).Bold = True
                        Next Counter
                    End If
                End If
                .MoveNext
            Loop
            
            lvwExport.SortKey = lvwAscending
            lvwExport.Sorted = True
        End If
        
    End With
    
NoRec:
        If lvwExport.ListItems.Count < 1 Then
            cmdExport(1).Enabled = False
            cmdExport(2).Enabled = False
            cmdExport(3).Enabled = False
            cmdExport(4).Enabled = False
        End If
        
    Set colEmptyFieldValues = Nothing
    
End Sub


Public Sub ASCIIConverter(ACode As String)
    On Error GoTo NoConvert:
    Dim Counter As Integer
    
    With m_rstSupp
        .MoveFirst
        Do While Not .EOF
            If ACode = ![ASCII UNIT] Then
                SBcode = ![SUPP UNIT]
            End If
            .MoveNext
        Loop
    End With
NoConvert:
End Sub

Public Sub SaveChanges()
    'On Error GoTo FrstRecord
    
    Dim blnAddNew As Boolean
    
    With m_rstCommon
        blnAddNew = True
        blnAddNew = blnAddNew And (.EOF And .BOF)
        
        If Not blnAddNew Then
            .MoveFirst
            .Find "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " ", , adSearchForward
            
            blnAddNew = .EOF
        End If
        
        If blnAddNew Then
            .AddNew
        End If
        
        ![TARIC CODE] = txtCode.Text
        '-----> General
        ![KEY DUTCH] = txtDutchKey.Text
        ![KEY FRENCH] = txtFrnchKey.Text
        ![DESC DUTCH] = txtDutchDesc.Text
        ![DESC FRENCH] = txtFrnchDesc.Text
                    
        '-----> Quantities
        ![SUPP STAT UNIT] = Left(cboSuppStat.Text, 2)
        ![SUPP CALC UNIT] = Left(cboSuppCalc.Text, 2)
        If cboSuppStatQ.Text = "" Then
            ![SUPP STAT QTY CODE] = "00"
        Else
            ![SUPP STAT QTY CODE] = Left(cboSuppStatQ.Text, 2)
        End If
        If cboSuppCalcQ.Text = "" Then
            ![SUPP CALC QTY CODE] = "00"
        Else
            ![SUPP CALC QTY CODE] = Left(cboSuppCalcQ.Text, 2)
        End If
        If cboGrosCalc.Text = "" Then
            ![GROSS WT CALC CODE] = "00"
        Else
            ![GROSS WT CALC CODE] = Left(cboGrosCalc.Text, 2)
        End If
        '-----> lock Codes
        If chkSuppStat.Value = 1 Then
            ![SUPP STAT LOCK CODE] = -1
        Else: ![SUPP STAT LOCK CODE] = 0
        End If
        If chkSuppCalc.Value = 1 Then
            ![SUPP CALC LOCK CODE] = -1
        Else
            ![SUPP CALC LOCK CODE] = 0
        End If
        If chkGrossCalc.Value = 1 Then
            ![GROSS WT LOCK CODE] = -1
        Else
            ![GROSS WT LOCK CODE] = 0
        End If
        
        .Update
        
        If blnAddNew Then
            InsertRecordset m_conTaric, m_rstCommon, "COMMON"
        Else
            UpdateRecordset m_conTaric, m_rstCommon, "COMMON"
        End If
    End With
    
End Sub

Public Sub SaveSimplified()
    Dim blnAddNew As Boolean
    
    With m_rstDetail
    'On Error GoTo FirstRecord:
        blnAddNew = True
        blnAddNew = blnAddNew And (.EOF And .BOF)
        
        '----> Check if Record already exist
        If Not blnAddNew Then
            .MoveFirst
            .Find "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " AND [CTRY CODE] = " & Chr(39) & ProcessQuotes(txtCtryCode.Text) & Chr(39) & " ", , adSearchForward
            
            blnAddNew = .EOF
        End If
        
        SaveDetailChanges blnAddNew
    End With
End Sub

Public Sub CancelCheck()
'-----> Checks if the updated taric record in table Import and Export exist in table common
'-----> If not cancel update

Dim intCancel As Integer
Dim strCancel As String

On Error GoTo NoRecord:
With m_rstCommon
    .MoveFirst
    Do While Not .EOF
        If ![TARIC CODE] = Trim(txtCode.Text) Then
            Exit Sub
        End If
        .MoveNext
    Loop
End With

NoRecord:

'-----> Msg box to update record
strCancel = Translate(867) & " " & txtCode.Text & "?"
intCancel = MsgBox(strCancel, vbYesNo + vbQuestion, Me.Caption)
If intCancel = 6 Then
    SaveChanges
    Exit Sub
End If

'-----> Remove Import
With m_rstImpGrid
    If .RecordCount = 0 Then GoTo NoImport:
    .MoveFirst
    Do While Not .EOF
        If ![TARIC CODE] = Trim(txtCode.Text) Then
            .Delete
        End If
        .MoveNext
    Loop
End With

NoImport:

'-----> Remove Export
With m_rstExpGrid
    If .RecordCount = 0 Then GoTo NoExport:
    .MoveFirst
    Do While Not .EOF
        If ![TARIC CODE] = Trim(txtCode.Text) Then
            .Delete
        End If
        .MoveNext
    Loop
End With

NoExport:
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Dim strCN As String
'-----> if backspace is pressed check if the taric code has been changed
    'allanSQL
    strCN = vbNullString
    strCN = strCN & "SELECT "
    strCN = strCN & "* "
    strCN = strCN & "FROM "
    strCN = strCN & "CN "
    strCN = strCN & "WHERE "
    strCN = strCN & "[CN CODE] = " & Chr(39) & ProcessQuotes(Left(txtCode.Text, 8)) & Chr(39) & " "

ADORecordsetOpen strCN, m_conTaric, m_rstMain, adOpenKeyset, adLockOptimistic
'Set m_rstMain = m_conTaric.OpenRecordset(strCN, dbOpenForwardOnly)
    With m_rstMain
        If .EOF Then Exit Sub
    End With
ADORecordsetClose m_rstMain
'Set m_rstMain = Nothing

If strMainType = "Normal" Then
    If Len(txtCode.Text) = 10 Then
        If Not txtCode.SelStart = 10 Then
            CancelCheck
            Exit Sub
        End If
        If KeyAscii = 8 Then
            CancelCheck
        End If
    End If
End If

End Sub

Public Sub DetailData()

    Dim strSQL As String
            
    'allanSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "* "
    strSQL = strSQL & "FROM "

    If strDocType = "Import" Then
        '-----> Open Import DB set recordset
        strSQL = strSQL & "IMPORT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    Else
        '----->Open export DB set recordest
        strSQL = strSQL & "EXPORT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(frm_taricmain.txtCode.Text) & Chr(39) & " "
    End If
ADORecordsetOpen strSQL, m_conTaric, m_rstDetail, adOpenKeyset, adLockOptimistic
'Set m_rstDetail = m_conTaric.OpenRecordset(strSQL)

End Sub

Private Sub txtCtry_Change()

'----> to conform with apply command standard
If Not blnStartApply Then Exit Sub
If strMainType = "Normal" Then
    cmdApply.Enabled = True
Else
    If blnSaveOK = True Then cmdApply.Enabled = True
End If

End Sub

Private Sub txtCtryCode_Change()
    Dim intBox As Integer
    Dim strSQL As String
    Dim colEmptyFieldValues As Collection
    
    '-----> Saving not allowed if Len(txtCtryCode.Text) <> 3
    'cmdOK.Enabled = False
    cmdApply.Enabled = False
    
    blnUncheck = False
    
    If Len(txtCtryCode.Text) = 3 Then
        With m_rstPick
            .MoveFirst
            
            Do Until .EOF
                If ![Internal Code] = "8.29801619052887E+19" And txtCtryCode.Text = ![code] Then
                    If strLangOfDesc = "Dutch" Then
                        txtCtry.Text = ![DESCRIPTION DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        txtCtry.Text = ![DESCRIPTION FRENCH]
                    End If
                    
                    GoTo OutSide:
                End If
                
                .MoveNext
            Loop
        End With
        
        '-----> If country code does not exist in database,
        '-----> clear txtCtryCode and exit.
        '-----> Disable saving.
        intBox = MsgBox(Translate(866), vbOKOnly + vbExclamation, Me.Caption)
        txtCtryCode.SelStart = 0
        txtCtryCode.SelLength = Len(txtCtryCode.Text)
        
        Exit Sub
        
OutSide:
        
        '-----> Enables OK and Apply in Simplified form
        ' If blnSaveOK Then
            blnSaveOK = True
            cmdOK.Enabled = True
            
            If blnStartApply Then
                cmdApply.Enabled = True    '-----> To conform with the standard
            End If
        ' End If
        
        '----> Save empty box field values to colEmptyFieldValues
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "[BOX CODE], "
            strSQL = strSQL & "[EMPTY FIELD VALUE] "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "[BOX DEFAULT " & strDocType & " ADMIN] "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[BOX CODE] "
            strSQL = strSQL & "IN "
            strSQL = strSQL & "( "
            strSQL = strSQL & "'N1', 'N2', 'N3', "
            strSQL = strSQL & "'O1', 'O2', 'O3', "
            strSQL = strSQL & "'P1', 'P2', 'P3', "
            strSQL = strSQL & "'Q1', 'Q2', 'Q3', "
            strSQL = strSQL & "'R1', 'R2', 'R3', "
            strSQL = strSQL & "'R4', 'R5', 'R6', "
            strSQL = strSQL & "'R7', 'R8', 'R9', "
            strSQL = strSQL & "'RA' "
            strSQL = strSQL & ") "
        ADORecordsetOpen strSQL, m_conSADBEL, m_rstEmptyFieldValues, adOpenKeyset, adLockOptimistic
        'Set m_rstEmptyFieldValues = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
        Set colEmptyFieldValues = New Collection
        
        With m_rstEmptyFieldValues
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do Until .EOF
                    colEmptyFieldValues.Add CStr(![EMPTY FIELD VALUE]), CStr(![BOX CODE])
                    
                    .MoveNext
                Loop
            End If
        End With
        
        ADORecordsetClose m_rstEmptyFieldValues
        
        '-----> Load values from database to settings detail
        With m_rstDetail
            On Error GoTo NoRecord:
            
            .MoveFirst
            
            Do Until .EOF
                If Trim(txtCtryCode.Text) = ![CTRY CODE] And txtCode.Text = ![TARIC CODE] Then
                    '-----> If selected country already exists, enter edit mode
                    '-----> Load required licence
                    If strDocType = "Import" Then
                        If ![LIC REQD] = True Then
                            chkLicenceImp.Value = vbChecked
                            
                            If Not IsNull(![MIN VALUE]) Then txtLimit.Text = ![MIN VALUE]
'                            If Not IsNull(![MIN VALUE CURR]) Then cboCurrency.Text = ![MIN VALUE CURR]
                        Else
                            chkLicenceImp.Value = vbUnchecked
                        End If
                    ElseIf strDocType = "Export" Then
                        If ![LIC REQD] = True Then
                            chkLicenceExp.Value = vbChecked
                        Else
                            chkLicenceExp.Value = vbUnchecked
                        End If
                    End If
                    
                    '-----> Load attached documents
                    txtType(0).Text = IIf(IsNull(![N1]), colEmptyFieldValues("N1"), ![N1])
                    txtType(1).Text = IIf(IsNull(![N2]), colEmptyFieldValues("N2"), ![N2])
                    txtType(2).Text = IIf(IsNull(![N3]), colEmptyFieldValues("N3"), ![N3])
                    txtNumber(0).Text = IIf(IsNull(![O1]), colEmptyFieldValues("O1"), ![O1])
                    txtNumber(1).Text = IIf(IsNull(![O2]), colEmptyFieldValues("O2"), ![O2])
                    txtNumber(2).Text = IIf(IsNull(![O3]), colEmptyFieldValues("O3"), ![O3])
                    txtDate(0).Text = IIf(IsNull(![P1]), colEmptyFieldValues("P1"), ![P1])
                    txtDate(1).Text = IIf(IsNull(![P2]), colEmptyFieldValues("P2"), ![P2])
                    txtDate(2).Text = IIf(IsNull(![P3]), colEmptyFieldValues("P3"), ![P3])
                    txtValue(0).Text = IIf(IsNull(![Q1]), colEmptyFieldValues("Q1"), ![Q1])
                    txtValue(1).Text = IIf(IsNull(![Q2]), colEmptyFieldValues("Q2"), ![Q2])
                    txtValue(2).Text = IIf(IsNull(![Q3]), colEmptyFieldValues("Q3"), ![Q3])
                    
                    '----> Load special regimes
                    txtReg(0).Text = IIf(IsNull(![R1]), colEmptyFieldValues("R1"), ![R1])
                    txtReg(1).Text = IIf(IsNull(![R3]), colEmptyFieldValues("R3"), ![R3])
                    txtReg(2).Text = IIf(IsNull(![R5]), colEmptyFieldValues("R5"), ![R5])
                    txtReg(3).Text = IIf(IsNull(![R7]), colEmptyFieldValues("R7"), ![R7])
                    txtReg(4).Text = IIf(IsNull(![R9]), colEmptyFieldValues("R9"), ![R9])
                    txtRegValue(0).Text = IIf(IsNull(![R2]), colEmptyFieldValues("R2"), ![R2])
                    txtRegValue(1).Text = IIf(IsNull(![R4]), colEmptyFieldValues("R4"), ![R4])
                    txtRegValue(2).Text = IIf(IsNull(![R6]), colEmptyFieldValues("R6"), ![R6])
                    txtRegValue(3).Text = IIf(IsNull(![R8]), colEmptyFieldValues("R8"), ![R8])
                    txtRegValue(4).Text = IIf(IsNull(![RA]), colEmptyFieldValues("RA"), ![RA])
                    
                    Set colEmptyFieldValues = Nothing
                    
                    '-----> Load usage
                    If ![DEF CODE] = True Then
                        chkDefault.Value = vbChecked
                    Else
                        chkDefault.Value = vbUnchecked
                    End If
                    
                    Call DefaultCheck
                    
                    If ![COMM CODE] = True Then
                        chkCommon.Value = vbChecked
                    Else
                        chkCommon.Value = vbUnchecked
                    End If
                    
                    Exit Sub
                End If
                
                .MoveNext
            Loop
        End With
        
NoRecord:
        
        '----> If entered country code doesn't exist in the DB then enter default values
        Call LoadDetailDefault
        Call DefaultCheck
        ' BlankOut
    ElseIf Trim(txtCtryCode.Text) = "" Then
        blnSaveOK = False    ' Prevents the enabling of the Apply button in the Simplified form.
        BlankOut
        txtCtry.Text = ""
    End If
End Sub

Public Sub BlankOut()
Dim Counter As Integer
Dim strSQL As String

'-----> Blank out country setting if there is no existing country code and taric code combination
        
'----> Save empty box field values to colEmptyFieldValues
    'allanSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "[BOX CODE], '"
    strSQL = strSQL & "[EMPTY FIELD VALUE] "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "[BOX DEFAULT " & strDocType & " ADMIN] "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "[BOX CODE] "
    strSQL = strSQL & "IN "
    strSQL = strSQL & "( "
    strSQL = strSQL & "'N1', 'N2', 'N3', "
    strSQL = strSQL & "'O1', 'O2', 'O3', "
    strSQL = strSQL & "'P1', 'P2', 'P3', "
    strSQL = strSQL & "'Q1', 'Q2', 'Q3', "
    strSQL = strSQL & "'R1', 'R2', 'R3', "
    strSQL = strSQL & "'R4', 'R5', 'R6', "
    strSQL = strSQL & "'R7', 'R8', 'R9', "
    strSQL = strSQL & "'RA' "
    strSQL = strSQL & ") "
ADORecordsetOpen strSQL, m_conSADBEL, m_rstEmptyFieldValues, adOpenKeyset, adLockOptimistic
'Set m_rstEmptyFieldValues = m_conSADBEL.OpenRecordset(strSQL, dbOpenForwardOnly)
Set colEmptyFieldValues = New Collection
      
With m_rstEmptyFieldValues
    If Not (.EOF And .BOF) Then
        .MoveFirst
        Do Until .EOF
            colEmptyFieldValues.Add CStr(![EMPTY FIELD VALUE]), CStr(![BOX CODE])
            .MoveNext
        Loop
    End If
End With

ADORecordsetClose m_rstEmptyFieldValues

'---->load if licence required
If strDocType = "Import" Then
    chkLicenceImp.Value = 0
ElseIf strDocType = "Export" Then
    chkLicenceExp.Value = 0
End If
                    
If strDocType <> "Transit NCTS" And strDocType <> G_CONST_EDINCTS1_TYPE Then
    '-----> load attached documents
    For Counter = 1 To 3
        txtType(Counter - 1).Text = colEmptyFieldValues("N" & Counter)
        txtNumber(Counter - 1).Text = colEmptyFieldValues("O" & Counter)
        txtDate(Counter - 1).Text = colEmptyFieldValues("P" & Counter)
        txtValue(Counter - 1).Text = colEmptyFieldValues("Q" & Counter)
    Next Counter
                               
    '----> load special regime
    For Counter = 1 To 5
        txtReg(Counter - 1).Text = colEmptyFieldValues("R" & (Counter + (Counter - 1)))
    Next Counter
    For Counter = 1 To 4
        txtRegValue(Counter - 1).Text = colEmptyFieldValues("R" & Counter * 2)
    Next Counter
    txtRegValue(4).Text = colEmptyFieldValues("RA")
End If

'-----> load usage
chkDefault.Value = 0
chkCommon.Value = 0

Set colEmptyFieldValues = Nothing

End Sub

Private Sub chkLicenceImp_Click()

'----> to conform with apply command standard
If blnStartApply Then
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End If


If chkLicenceImp.Value = 1 Then
    txtLimit.Text = Format(dblMinVal, Display)
    If Right(txtLimit.Text, 3) = ",00" Then
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 3)
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 2) & "." & Right(txtLimit.Text, 2)
    End If
    txtLimit.Enabled = True
'    cboCurrency.Enabled = True
'    cboCurrency.Text = strLicCurr
Else
    txtLimit.Text = ""
    txtLimit.Enabled = False
'    cboCurrency.Enabled = False
'    cboCurrency.Text = ""
End If

End Sub

Public Sub SaveDetailChanges(ByVal AddNew As Boolean)

    Dim strSQL As String
    
    '----> Save to Database
    With m_rstDetail
        If AddNew Then
            .AddNew
        End If
        '-----> Save General
        ![TARIC CODE] = txtCode.Text
        ![CTRY CODE] = txtCtryCode.Text
         
         '----> Save licence
        If strDocType = "Import" Then
             If chkLicenceImp.Value = 1 Then
                 ![LIC REQD] = -1
                 ![MIN VALUE] = txtLimit.Text
                 ![Min Value Curr] = "EUR"    ' cboCurrency.Text
             ElseIf chkLicenceImp.Value = 0 Then ![LIC REQD] = 0
                 ![MIN VALUE] = ""
                 ![Min Value Curr] = ""
             End If
         ElseIf strDocType = "Export" Then
             If chkLicenceExp.Value = 1 Then
                 ![LIC REQD] = -1
             ElseIf chkLicenceExp.Value = 0 Then ![LIC REQD] = 0
             End If
         End If
        
         '-----> Save attached documents
         ![N1] = txtType(0).Text
         ![N2] = txtType(1).Text
         ![N3] = txtType(2).Text
         ![O1] = txtNumber(0).Text
         ![O2] = txtNumber(1).Text
         ![O3] = txtNumber(2).Text
         ![P1] = txtDate(0).Text
         ![P2] = txtDate(1).Text
         ![P3] = txtDate(2).Text
         ![Q1] = txtValue(0).Text
         ![Q2] = txtValue(1).Text
         ![Q3] = txtValue(2).Text
                 
         '----> Save special regime
         ![R1] = txtReg(0).Text
         ![R3] = txtReg(1).Text
         ![R5] = txtReg(2).Text
         ![R7] = txtReg(3).Text
         ![R9] = txtReg(4).Text
         ![R2] = txtRegValue(0).Text
         ![R4] = txtRegValue(1).Text
         ![R6] = txtRegValue(2).Text
         ![R8] = txtRegValue(3).Text
         ![RA] = txtRegValue(4).Text

         '-----> Save usage
         If chkDefault.Value = 1 Then
             ![DEF CODE] = -1
         Else
             If blnUncheck = True Then
                 '-----> make the lowest country as the default country when the default botton is unchecked
                     'allanSQL
                     strSQL = vbNullString
                     strSQL = strSQL & "SELECT "
                     strSQL = strSQL & "* "
                     strSQL = strSQL & "FROM "
                     strSQL = strSQL & "[" & GetTableToUse(strDocType) & "] "
                     strSQL = strSQL & "WHERE "
                     strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
                     strSQL = strSQL & "ORDER BY "
                     strSQL = strSQL & strDocType & ".[CTRY CODE] ASC "
                 ADORecordsetOpen strSQL, m_conTaric, m_rstDefault, adOpenKeyset, adLockOptimistic
                 'Set m_rstDefault = m_conTaric.OpenRecordset(strSQL)
                 If Not (m_rstDefault.EOF And m_rstDefault.BOF) Then
                    m_rstDefault.MoveFirst
                    
                    Do While Not m_rstDefault.EOF
                         If m_rstDefault![DEF CODE] = -1 Then
                             If m_rstDefault.RecordCount = 1 Then
                                 m_rstDefault.MoveNext
                                 GoTo ExitLoop
                             Else
                                 m_rstDefault.MoveFirst
                                 GoTo ExitLoop:
                             End If
                         End If
                         m_rstDefault.MoveNext
                     Loop
ExitLoop:
                    If Not m_rstDefault.EOF Then
                         'm_rstDefault.Edit
                         m_rstDefault![DEF CODE] = -1
                         m_rstDefault.Update
                         
                         UpdateRecordset m_conTaric, m_rstDefault, GetTableToUse(strDocType)
                    End If
                End If
                
                ADORecordsetClose m_rstDefault

                ![DEF CODE] = 0
             Else
                ![DEF CODE] = 0
             End If
             
         End If
         
         If chkCommon.Value = 1 Then
             ![COMM CODE] = -1
         Else
            ![COMM CODE] = 0
         End If
         
         .Update
        
        If AddNew Then
            InsertRecordset m_conTaric, m_rstDetail, strDocType
        Else
            UpdateRecordset m_conTaric, m_rstDetail, strDocType
        End If
        '----->Clear Previous usage default
        If chkDefault.Value = 1 Then
            .MoveFirst
            Do While Not .EOF
                If ![TARIC CODE] = txtCode.Text Then
                    If Not ![CTRY CODE] = txtCtryCode.Text Then
                        '.Edit
                        ![DEF CODE] = 0
                        ![COMM CODE] = 0
                        .Update
                        
                        UpdateRecordset m_conTaric, m_rstDetail, strDocType
                    End If
                End If
                
                .MoveNext
            Loop
        End If
    
        
    End With
End Sub


Public Sub SaveCopy()
    Dim i As Integer
    Dim intArrayCtr As Integer
    Dim j As Integer
    Dim rstCopy As ADODB.Recordset
    Dim blnIncollection As Boolean
    Dim strArray() As String
    Dim blnAdd As Boolean
    Dim strToStore() As String
    'allanSQL****************
    Dim strSQL As String
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "IMPORT"
    ADORecordsetOpen strSQL, m_conTaric, rstCopy, adOpenKeyset, adLockOptimistic
    'Set rstCopy = m_conTaric.OpenRecordset(strSQL)

    With m_rstImpGrid
        
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            blnIncollection = InCollection("I")
            ReDim strToStore(.RecordCount - 1)
            For i = 0 To .RecordCount - 1
                'check if taric code and ctry code combination already exists
                blnAdd = True
                If blnIncollection = True Then
                    'check if [CTRY CODE] is in array
                    strArray() = colTaricCodes(txtCode.Text & "I")
        
                    For j = 0 To UBound(strArray)
                        If strArray(j) = ![CTRY CODE] Then
                            blnAdd = False
                            Exit For
                        End If
                    Next
                    colTaricCodes.Remove txtCode.Text & "I"
                End If
                
                If Not (rstCopy.EOF And rstCopy.BOF) Then
                    rstCopy.MoveFirst
                    rstCopy.Find "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " and [CTRY CODE] = " & Chr(39) & ProcessQuotes(![CTRY CODE]) & Chr(39), , adSearchForward
                
                    If Not rstCopy.EOF Then
                        If blnAdd Then
                            If MsgBox(Translate(894), vbYesNo + vbQuestion + vbApplicationModal, Translate(896)) = vbYes Then
                                blnAdd = False
                            Else
                                GoTo nextrecord1
                            End If
                        End If
                    End If
                End If
                
                strToStore(i) = ![CTRY CODE]
                Call UpdateThisRecordset(blnAdd, rstCopy, m_rstImpGrid, "I")
nextrecord1:
                .MoveNext
                
                'Loop
            Next
            
            colTaricCodes.Add strToStore(), txtCode.Text & "I"
        End If
    End With
    
NoImport:
    
    
    ADORecordsetOpen "Select * from EXPORT", m_conTaric, rstCopy, adOpenKeyset, adLockOptimistic
    'Set rstCopy = m_conTaric.OpenRecordset("Select * from EXPORT")
    On Error GoTo NoRecord:
    With m_rstExpGrid
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            blnIncollection = InCollection("E")
            ReDim strToStore(.RecordCount - 1)
            For i = 0 To .RecordCount - 1
                'check if taric code and ctry code combination already exists
                blnAdd = True
                If blnIncollection = True Then
                    'check if [CTRY CODE] is in array
                    strArray() = colTaricCodes(txtCode.Text & "E")
        
                    For j = 0 To UBound(strArray)
                        If strArray(j) = ![CTRY CODE] Then
                            blnAdd = False
                            Exit For
                        End If
                    Next
                    colTaricCodes.Remove txtCode.Text & "E"
                End If
                
                If Not (rstCopy.EOF And rstCopy.BOF) Then
                    rstCopy.MoveFirst
                    rstCopy.Find "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " and [CTRY CODE] = " & Chr(39) & ProcessQuotes(![CTRY CODE]) & Chr(39), , adSearchForward
                    
                    If Not rstCopy.EOF Then
                        If blnAdd Then
                            If MsgBox(Translate(894), vbYesNo + vbQuestion + vbApplicationModal, Translate(896)) = vbYes Then
                                blnAdd = False
                            Else
                                GoTo nextrecord2
                            End If
                        End If
                    End If
                End If
                
                strToStore(i) = ![CTRY CODE]
                Call UpdateThisRecordset(blnAdd, rstCopy, m_rstExpGrid, "E")
nextrecord2:
                .MoveNext
                
                'Loop
            Next
            
            colTaricCodes.Add strToStore(), txtCode.Text & "E"
            
        End If
    End With
    
NoRecord:

End Sub

Private Sub txtDate_Change(Index As Integer)

    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If

End Sub

Private Sub txtDate_LostFocus(Index As Integer)
    '-----> check if the date inputed is valid
    If txtDate(Index).Text = "" Or txtDate(Index).Text = "0" Then Exit Sub
    If CheckDate(txtDate(Index).Text) = False Then
        txtDate(Index).SetFocus
        txtDate(Index).SelStart = 0
        txtDate(Index).SelLength = Len(txtDate(Index).Text)
    End If

End Sub



Private Sub txtDutchDesc_Change()
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End Sub

Private Sub txtDutchKey_Change()
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
  
End Sub

Private Sub txtDutchKey_KeyPress(KeyAscii As Integer)
    If KeyAscii < 123 And KeyAscii > 96 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub



Private Sub txtFrnchDesc_Change()
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End Sub

Private Sub txtFrnchKey_Change()
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End Sub

Private Sub txtFrnchKey_KeyPress(KeyAscii As Integer)
    If KeyAscii < 123 And KeyAscii > 96 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtLimit_Change()
    If Left(Right(txtLimit.Text, 4), 1) = "." Then
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 1)
        txtLimit.SelStart = Len(txtLimit.Text)
    End If
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If

End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)

    '-----> Allow only backspace and numerical values
    If KeyAscii = 46 Then If InStr(txtLimit.Text, ".") = 0 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtLimit_LostFocus()

    'txtLimit.Text = Format(txtLimit.Text, Display)
    If InStr(txtLimit.Text, ".") = 0 Then txtLimit.Text = txtLimit.Text & ".00"
    If InStr(txtLimit.Text, ".") = (Len(txtLimit.Text) - 1) Then txtLimit.Text = txtLimit.Text & "0"
    If InStr(txtLimit.Text, ".") = Len(txtLimit.Text) Then txtLimit.Text = txtLimit.Text & "00"
    If Right(txtLimit.Text, 3) = ",00" Then
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 3)
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 2) & "." & Right(txtLimit.Text, 2)
    End If
End Sub

Private Sub txtNumber_Change(Index As Integer)
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If

End Sub

Private Sub txtNumber_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtReg_Change(Index As Integer)
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End Sub

Private Sub txtRegValue_Change(Index As Integer)

    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If

End Sub

Private Sub txtRegValue_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtType_Change(Index As Integer)
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If

End Sub

Private Sub txtValue_Change(Index As Integer)
    '----> to conform with apply command standard
    If Not blnStartApply Then Exit Sub
    If strMainType = "Normal" Then
        cmdApply.Enabled = True
    Else
        If blnSaveOK = True Then cmdApply.Enabled = True
    End If
End Sub

Private Sub txtValue_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 46 Then If InStr(txtValue(Index).Text, ".") = 0 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtCtryCode_KeyPress(KeyAscii As Integer)
    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Public Sub FillFromCodi()
    '-----> in cases that the taric code in the codsheet or exsheet does not exist
    '-----> in the taric database and the user wishes to add this code
    '-----> data is retrieved straight from the codisheet
    Dim Counter As Integer
    Dim formTaric As Form
    
    If strDocType = "Import" Then
        Set formTaric = newform(Right(Left(gstrTaricMainCallType, 8), 1))
    Else
        If strDocType = G_CONST_NCTS1_TYPE Then
            'Exit Sub    'exit muna
            Set formTaric = frmNewFormNCTS1(Right(Left(gstrTaricMainCallType, 8), 1))
        ElseIf strDocType = G_CONST_NCTS2_TYPE Then
            'Exit Sub    'exit muna
            Set formTaric = frmNewFormNCTS2(Right(Left(gstrTaricMainCallType, 8), 1))
        ElseIf strDocType = G_CONST_EDINCTS1_TYPE Then
            'Exit Sub    'exit muna
            Set formTaric = frmNewFormEDINCTS1(Right(Left(gstrTaricMainCallType, 8), 1))
        
        Else
            '-----> Load from codexport
            If Right(gstrTaricMainCallType, 1) = "E" Then
                '-----> export
                Set formTaric = newformE(Right(Left(gstrTaricMainCallType, 8), 1))
            Else
                '-----> Transit
                Set formTaric = newformT(Right(Left(gstrTaricMainCallType, 8), 1))
            End If
        End If
    End If
    
    With formTaric
        '-----> to fill in description and keyword
        'txtCode.Text = left(.colTaricArgs("L1"),8)
        txtCode.Text = .colTaricArgs("L1")
        
                        Dim strS1 As String
                        Dim strTempDesc As String
                        Dim intSearchCtr As Integer
                        Dim strKeywordD As String
                        Dim strKeywordF As String
                        Dim strSQL As String
                        Dim rstCN As ADODB.Recordset
                        
                        strS1 = .colTaricArgs("S1")
                        
    ' ********** Modified August 30, 2001 **********
    ' ********** Adding a test for Len(strS1) and changing Right(strS1) to Mid(strS1) remedies
    ' ********** run-time error 5 [Invalid procedure call or argument] that occurs upon the
    ' ********** application of Right() with a negative Length arg.
                        If Len(strS1) Then
                            If .colTaricArgs("A5") = "N" Or .colTaricArgs("A5") = "n" Then
                                If UCase(.colTaricArgs("T1")) Like "MA*" Then        ' Case "MAxxx"
                                    txtDutchDesc.Text = Mid(strS1, 13)
                                ElseIf UCase(.colTaricArgs("T1")) Like "LY*" Then    ' Case "LYxxx"
                                    txtDutchDesc.Text = Mid(strS1, 16)
                                ElseIf .colTaricArgs("T1") = "126" Then
                                    If .colTaricArgs("A4") = "212" Then
                                        If strDocType = "Import" Then
                                            If Not .colTaricArgs("C3") = "CE" Then
                                                txtDutchDesc.Text = Mid(strS1, 16)
                                            End If
                                        Else
                                            txtDutchDesc.Text = Mid(strS1, 16)
                                        End If
                                    End If
                                ElseIf .colTaricArgs("T1") = "126E" Then
                                    If .colTaricArgs("A4") = "101" Then
                                        strTempDesc = strS1
                                        For intSearchCtr = 1 To 3
                                            strTempDesc = Right(strTempDesc, Len(strTempDesc) - InStr(strTempDesc, "*"))
                                        Next intSearchCtr
                                        txtDutchDesc.Text = strTempDesc
                                    End If
                                ElseIf .colTaricArgs("T1") = "126T" Or .colTaricArgs("T1") = "126TS" Then
                                    If strDocType = "Export" Then
                                        txtDutchDesc.Text = Right(strS1, Len(strS1) - InStr(strS1, "*"))
                                    End If
                                End If
                                
                                If Len(txtDutchDesc.Text) = 0 Then
                                    txtDutchDesc.Text = strS1
                                End If
                                
    ' ********** Added September 6, 2001 **********
    ' ********** Strips leading asterisks from the description.
                                txtDutchDesc.Text = StripChars(txtDutchDesc.Text, "*", sbpLeading)
    ' ********** End Add **************************
                            ElseIf .colTaricArgs("A5") = "F" Or .colTaricArgs("A5") = "f" Then
                                If UCase(.colTaricArgs("T1")) Like "MA*" Then        ' Case "MAxxx"
                                    txtFrnchDesc.Text = Mid(strS1, 13)
                                ElseIf UCase(.colTaricArgs("T1")) Like "LY*" Then    ' Case "LYxxx"
                                    txtFrnchDesc.Text = Mid(strS1, 16)
                                ElseIf .colTaricArgs("T1") = "126" Then
                                    If .colTaricArgs("A4") = "212" Then
                                        If strDocType = "Import" Then
                                            If Not .colTaricArgs("C3") = "CE" Then
                                                txtFrnchDesc.Text = Mid(strS1, 16)
                                            End If
                                        Else
                                            txtFrnchDesc.Text = Mid(strS1, 16)
                                        End If
                                    End If
                                ElseIf .colTaricArgs("T1") = "126E" Then
                                    If .colTaricArgs("A4") = "101" Then
                                        strTempDesc = strS1
                                        For intSearchCtr = 1 To 3
                                            strTempDesc = Right(strTempDesc, Len(strTempDesc) - InStr(strTempDesc, "*"))
                                        Next intSearchCtr
                                        txtFrnchDesc.Text = strTempDesc
                                    End If
                                ElseIf .colTaricArgs("T1") = "126T" Or .colTaricArgs("T1") = "126TS" Then
                                    If strDocType = "Export" Then
                                        txtFrnchDesc.Text = Right(strS1, Len(strS1) - InStr(strS1, "*"))
                                    End If
                                End If
                                
                                If Len(txtFrnchDesc.Text) = 0 Then
                                    txtFrnchDesc.Text = strS1
                                End If
                                
    ' ********** Added September 6, 2001 **********
    ' ********** Strips leading asterisks from the description.
                                txtFrnchDesc.Text = StripChars(txtFrnchDesc.Text, "*", sbpLeading)
    ' ********** End Add **************************
                            End If
                        End If
    ' ********** End Modify ************************
                        
    ' ********** Modified August 30, 2001 **********
    ' ********** Optimized Jasper's code by combining the retrieval of the two descriptions
    ' ********** in just one database access.
                        '-----> if descriptions are still null
                        If txtDutchDesc.Text = "" Or txtFrnchDesc.Text = "" Then
                                'allanSQL
                                strSQL = vbNullString
                                strSQL = strSQL & "SELECT "
                                strSQL = strSQL & "[DESC DUTCH], "
                                strSQL = strSQL & "[DESC FRENCH] "
                                strSQL = strSQL & "FROM "
                                strSQL = strSQL & "CN "
                                strSQL = strSQL & "WHERE "
                                strSQL = strSQL & "[CN CODE] = " & Chr(39) & ProcessQuotes(Left$(.colTaricArgs("L1"), 8)) & Chr(39) & " "
                            ADORecordsetOpen strSQL, m_conTaric, rstCN, adOpenKeyset, adLockOptimistic
                            'Set rstCN = m_conTaric.OpenRecordset(strSQL, dbOpenForwardOnly)
                            
                            With rstCN
                                If Not (.EOF And .BOF) Then
                                    .MoveFirst
                                    
                                    If txtDutchDesc.Text = "" Then
                                        txtDutchDesc.Text = Left(![DESC DUTCH], 78)
                                    End If
                                    
                                    If txtFrnchDesc.Text = "" Then
                                        txtFrnchDesc.Text = Left(![DESC FRENCH], 78)
                                    End If
                                End If
                            End With
                            
                            ADORecordsetClose rstCN
                        End If
    
            strKeywordD = txtDutchDesc.Text
Ulit:
            If UCase(Left(strKeywordD, 5)) = "ANDER" Or UCase(Left(strKeywordD, 6)) = "ANDER," Or _
                UCase(Left(strKeywordD, 6)) = "ANDERE" Or UCase(Left(strKeywordD, 7)) = "ANDERE," Or _
                UCase(Left(strKeywordD, 7)) = "NUMBERS" Or UCase(Left(strKeywordD, 8)) = "NUMBERS," Or _
                UCase(Left(strKeywordD, 5)) = "DELEN" Or UCase(Left(strKeywordD, 6)) = "DELEN," Or _
                UCase(Left(strKeywordD, 9)) = "GEBRUIKTE" Or UCase(Left(strKeywordD, 10)) = "GEBRUIKTE," Or _
                UCase(Left(strKeywordD, 7)) = "GETUFTE" Or UCase(Left(strKeywordD, 8)) = "GETUFTE," Or _
                UCase(Left(strKeywordD, 7)) = "GEVULDE" Or UCase(Left(strKeywordD, 8)) = "GEVULDE," Or _
                UCase(Left(strKeywordD, 7)) = "GEZAAGD" Or UCase(Left(strKeywordD, 8)) = "GEZAAGD," Or _
                UCase(Left(strKeywordD, 5)) = "STUKS" Or UCase(Left(strKeywordD, 6)) = "STUKS," Then
                
                strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                
                If Not InStr(strKeywordD, " ") = 0 Then
                    GoTo Ulit:
                End If
            End If
            
            If Not strKeywordD = "" And Not InStr(strKeywordD, " ") = 0 Then
            
                If IsNumeric(Left(strKeywordD, InStr(strKeywordD, " ") - 1)) Then
                
                    strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                    
                    If Not InStr(strKeywordD, " ") = 0 Then
                        GoTo Ulit:
                    End If
                End If
            End If
            If InStr(strKeywordD, " ") > 4 Then
                txtDutchKey.Text = UCase(Left(strKeywordD, InStr(strKeywordD, " ") - 1))
                If Right(txtDutchKey.Text, 1) = "," Then
                    txtDutchKey.Text = Left(txtDutchKey.Text, Len(txtDutchKey.Text) - 1)
                End If
            ElseIf InStr(strKeywordD, " ") = 0 Then
                txtDutchKey.Text = UCase(strKeywordD)
                If Right(txtDutchKey.Text, 1) = "," Then
                    txtDutchKey.Text = Left(txtDutchKey.Text, Len(txtDutchKey.Text) - 1)
                End If
            Else
                strKeywordD = Right(strKeywordD, Len(strKeywordD) - InStr(strKeywordD, " "))
                GoTo Ulit:
            End If
     
            strKeywordF = txtFrnchDesc.Text
Ulit2:
            If UCase(Left(strKeywordF, 5)) = "AUTRE" Or UCase(Left(strKeywordF, 6)) = "AUTRE," Or _
                UCase(Left(strKeywordF, 6)) = "AUTRES" Or UCase(Left(strKeywordF, 7)) = "AUTRES," Or _
                UCase(Left(strKeywordF, 7)) = "NUMBERS" Or UCase(Left(strKeywordF, 8)) = "NUMBERS," Or _
                UCase(Left(strKeywordF, 9)) = "DELEN VAN" Or UCase(Left(strKeywordF, 10)) = "DELEN VAN," Then
                
                strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                
                If Not InStr(strKeywordF, " ") = 0 Then
                    GoTo Ulit2:
                End If
            End If
            
            If Not strKeywordF = "" And Not InStr(strKeywordF, " ") = 0 Then
            
                If IsNumeric(Left(strKeywordF, InStr(strKeywordF, " ") - 1)) Then
                
                    strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                    
                    If Not InStr(strKeywordF, " ") = 0 Then
                        GoTo Ulit2:
                    End If
                End If
            End If
            
            If InStr(strKeywordF, " ") > 4 Then
                txtFrnchKey.Text = UCase(Left(strKeywordF, InStr(strKeywordF, " ") - 1))
                If Right(txtFrnchKey.Text, 1) = "," Then
                    txtFrnchKey.Text = Left(txtFrnchKey.Text, Len(txtFrnchKey.Text) - 1)
                End If
            ElseIf InStr(strKeywordF, " ") = 0 Then
                txtFrnchKey.Text = UCase(strKeywordF)
                If Right(txtFrnchKey.Text, 1) = "," Then
                    txtFrnchKey.Text = Left(txtFrnchKey.Text, Len(txtFrnchKey.Text) - 1)
                End If
            Else
                strKeywordF = Right(strKeywordF, Len(strKeywordF) - InStr(strKeywordF, " "))
                GoTo Ulit2:
            End If
            
                    '-----> Fill in Quantities
                    
            If strDocType = G_CONST_NCTS2_TYPE Then
                If Not (m_rstPick.EOF And m_rstPick.BOF) Then
                    m_rstPick.MoveFirst
                    Do While Not m_rstPick.EOF
                        '----->Supplementary Statistical Unit
                        If m_rstPick![Internal Code] = "7.67111659049988E+19" And .colTaricArgs("M3") = m_rstPick![code] Then
                            If strLangOfDesc = "Dutch" Then
                                cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION DUTCH]
                            ElseIf strLangOfDesc = "French" Then
                                cboSuppStat.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION FRENCH]
                            End If
                        End If
                        
                        If strDocType = "Import" Then
                            '----->Supplementary Calculation Unit
                            If m_rstPick![Internal Code] = "5.35045266151428E+18" And .colTaricArgs("M5") = m_rstPick![code] Then
                                If strLangOfDesc = "Dutch" Then
                                    cboSuppCalc.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION DUTCH]
                                ElseIf strLangOfDesc = "French" Then
                                    cboSuppCalc.Text = m_rstPick![code] & "-" & m_rstPick![DESCRIPTION FRENCH]
                                End If
                            End If
                        End If
                        m_rstPick.MoveNext
                    Loop
                End If
            End If
                    
        '-----> country code
        If strDocType = "Import" Then
            txtCtryCode.Text = .colTaricArgs("C1")
        Else
            txtCtryCode.Text = .colTaricArgs("C2")
        End If
    On Error GoTo NoLicence:
    
        '-----> Licence
        If strDocType = "Import" Then
            If Not .colTaricArgs("M8") = 0 Then
                chkLicenceImp.Value = 1
                If Not gstrMinLicValue = "" Then
                    txtLimit.Text = gstrMinLicValue
                End If
            End If
        ElseIf strDocType = G_CONST_NCTS2_TYPE Then
            If Not (.colTaricArgs("M6")) = 0 Then chkLicenceExp.Value = 1
        End If
NoLicence:
               
        '-----> load attached documents
        
        If strDocType = G_CONST_NCTS2_TYPE Then
            txtType(0).Text = .colTaricArgs("N1")
            txtType(1).Text = .colTaricArgs("N2")
            txtType(2).Text = .colTaricArgs("N3")
            txtNumber(0).Text = .colTaricArgs("O1")
            txtNumber(1).Text = .colTaricArgs("O2")
            txtNumber(2).Text = .colTaricArgs("O3")
            txtDate(0).Text = .colTaricArgs("P1")
            txtDate(1).Text = .colTaricArgs("P2")
            txtDate(2).Text = .colTaricArgs("P3")
            txtValue(0).Text = .colTaricArgs("Q1")
            txtValue(1).Text = .colTaricArgs("Q2")
            txtValue(2).Text = .colTaricArgs("Q3")
                            
            '----> load special regime
            txtReg(0).Text = .colTaricArgs("R1")
            txtReg(1).Text = .colTaricArgs("R3")
            txtReg(2).Text = .colTaricArgs("R5")
            txtReg(3).Text = .colTaricArgs("R7")
            txtReg(4).Text = .colTaricArgs("R9")
            txtRegValue(0).Text = .colTaricArgs("R2")
            txtRegValue(1).Text = .colTaricArgs("R4")
            txtRegValue(2).Text = .colTaricArgs("R6")
            txtRegValue(3).Text = .colTaricArgs("R8")
            txtRegValue(4).Text = .colTaricArgs("RA")
        End If
        
        '-----> default
        chkDefault.Value = 1
    End With
End Sub


Public Sub LoadDetailDefault()

    '----->load simplified form with data from the default country
    
    Dim strSQL As String
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[" & GetTableToUse(strDocType) & "] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[DEF CODE] = -1  "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, m_rstDefault, adOpenKeyset, adLockOptimistic
    'Set m_rstDefault = m_conTaric.OpenRecordset(strSQL, dbOpenForwardOnly)
    With m_rstDefault
        If Not (.EOF And .BOF) Then
            .MoveFirst
                        
            '---->load if licence required
            If strDocType = "Import" Then
                If ![LIC REQD] = -1 Then
                    chkLicenceImp.Value = 1
                    If Not IsNull(![MIN VALUE]) Then
                        txtLimit.Text = ![MIN VALUE]
                    End If
                Else
                    chkLicenceImp.Value = 0
                End If
            ElseIf strDocType = "Export" Then
                If ![LIC REQD] = -1 Then
                    chkLicenceExp.Value = 1
                Else
                    chkLicenceExp.Value = 0
                End If
            End If
            
            '-----> load attached documents
            If Not IsNull(![N1]) Then txtType(0).Text = ![N1]
            If Not IsNull(![N2]) Then txtType(1).Text = ![N2]
            If Not IsNull(![N3]) Then txtType(2).Text = ![N3]
            If Not IsNull(![O1]) Then txtNumber(0).Text = ![O1]
            If Not IsNull(![O2]) Then txtNumber(1).Text = ![O2]
            If Not IsNull(![O3]) Then txtNumber(2).Text = ![O3]
            If Not IsNull(![P1]) Then txtDate(0).Text = ![P1]
            If Not IsNull(![P2]) Then txtDate(1).Text = ![P2]
            If Not IsNull(![P3]) Then txtDate(2).Text = ![P3]
            If Not IsNull(![Q1]) Then txtValue(0).Text = ![Q1]
            If Not IsNull(![Q2]) Then txtValue(1).Text = ![Q2]
            If Not IsNull(![Q3]) Then txtValue(2).Text = ![Q3]
            
            '----> load special regime
            If Not IsNull(![R1]) Then txtReg(0).Text = ![R1]
            If Not IsNull(![R3]) Then txtReg(1).Text = ![R3]
            If Not IsNull(![R5]) Then txtReg(2).Text = ![R5]
            If Not IsNull(![R7]) Then txtReg(3).Text = ![R7]
            If Not IsNull(![R9]) Then txtReg(4).Text = ![R9]
            If Not IsNull(![R2]) Then txtRegValue(0).Text = ![R2]
            If Not IsNull(![R4]) Then txtRegValue(1).Text = ![R4]
            If Not IsNull(![R6]) Then txtRegValue(2).Text = ![R6]
            If Not IsNull(![R8]) Then txtRegValue(3).Text = ![R8]
            If Not IsNull(![RA]) Then txtRegValue(4).Text = ![RA]
            
            '----> Load Default
            chkDefault.Value = 0
        Else
            BlankOut
        End If
    End With
    
    ADORecordsetClose m_rstDefault
    
End Sub

Private Sub GetFromKluwer()
' ********** 10/22/02 **********
' Data saved into clipboard by Kluwer has changed. The country code now comes
' before the ECO number and the ECO number is now space delimited and has no
' more parenthesis. strKluwerOut(4) used to hold the ECO number and strKluwerOut(5)
' holds the country code. Now its strKluwerOut(4) = country code, strKluwerOut(5) = ECO
' However support for old version should still be implemented.
' **********   end.   **********

    Dim intFreeFile As Integer
    Dim intSubscript As Integer
    Dim strTempFile As String
    Dim strKluwerOut() As String
    
    Dim strECONum As String
    Dim intImport As Integer
    Dim intExport As Integer
    Dim fldImport As ADODB.Field
    Dim fldExport As ADODB.Field
    Dim strSQL As String
    
    ' ***** 10/22/02 *****
    Dim strFileName As String
    Dim blnNewVersion As Boolean
    Dim strTempEco As String
    ' *****   end.   *****
    
    '-----> Save data from the clipboard to a temporary text file.
    strTempFile = cAppPath & "\KluwerTemp.txt"
    intFreeFile = FreeFile()
    
    Open strTempFile For Output As #intFreeFile
        Write #intFreeFile, Clipboard.GetText
    Close #intFreeFile
    
    '-----> Save data from temporary text file to variables.
    intFreeFile = FreeFile()
    
    Open strTempFile For Input As #intFreeFile
    
    Do Until EOF(intFreeFile)
        intSubscript = intSubscript + 1
        ReDim Preserve strKluwerOut(intSubscript)
        Line Input #intFreeFile, strKluwerOut(intSubscript)
    Loop
    
    Close #intFreeFile
    
    '-----> Delete temporary text file.
    Kill strTempFile
    
    '-----> Enter first 8 characters first to load from CN table.
    txtCode.Text = Left(Right(strKluwerOut(1), Len(strKluwerOut(1)) - 1), 8)
    
    '-----> Enter whole code.
    txtCode.Text = Right(strKluwerOut(1), Len(strKluwerOut(1)) - 1)
    
    If strLangOfDesc = "French" Then
        txtFrnchDesc.Text = strKluwerOut(2) & " - " & txtFrnchDesc.Text
    Else
        txtDutchDesc.Text = strKluwerOut(2) & " - " & txtDutchDesc.Text
    End If
    
    '-----> Supplementary Statistical Unit
    With m_rstPick
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            Do Until .EOF
                If ![Internal Code] = "7.67111659049988E+19" And strKluwerOut(3) = ![code] Then
                    If strLangOfDesc = "Dutch" Then
                        cboSuppStat.Text = ![code] & "-" & ![DESCRIPTION DUTCH]
                    ElseIf strLangOfDesc = "French" Then
                        cboSuppStat.Text = ![code] & "-" & ![DESCRIPTION FRENCH]
                    End If
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
    ' ***** 10/22/02 *****
    ' determine if new version
    strFileName = IIf(strLangOfDesc = "Dutch", GetSetting(App.Title, "Third Party Database", "DutchFile"), GetSetting(App.Title, "Third Party Database", "FrenchFile"))

    ' create flag for new version
    If Right(Trim(strFileName), 3) = "xml" Then
        blnNewVersion = True
    Else
        blnNewVersion = False
    End If
    ' *****   end.   *****
    
    If strMainType = "Simplified" Then    '-----> If loaded as Simplified, load ECO data to detail fields.
        '-----> Load country.
        'txtCtryCode.Text = strKluwerOut(5)
        ' ***** 10/22/02 *****
        'txtCtryCode.Text = strKluwerOut(5)
        If blnNewVersion Then
            txtCtryCode.Text = strKluwerOut(4)
            strTempEco = strKluwerOut(5)
        Else
            txtCtryCode.Text = strKluwerOut(5)
            strTempEco = strKluwerOut(4)
        End If
        ' *****   end.   *****
        
        '---/--> Load type and number in attached documents.
        Do Until Len(strTempEco) = 0
            '-----> Extract ECO numbers from line 4.
            ' ***** 10/20/02 *****
            ' If this is a new version then it would be tab delimited instead of
            ' parenthesis...
            If blnNewVersion Then
                'strTempEco = Trim(strTempEco)
                strTempEco = Trim(strTempEco)
                If InStr(strTempEco, " ") Then
                    strECONum = Left(strTempEco, InStr(strTempEco, " ") - 1)
                Else
                    strECONum = strTempEco
                End If
            Else
'                strKluwerOut(4) = Right(strKluwerOut(4), Len(strKluwerOut(4)) - 1)
'                strECONum = Left(strKluwerOut(4), InStr(strKluwerOut(4), ")") - 1)
                strTempEco = Right(strTempEco, Len(strTempEco) - 1)
                strECONum = Left(strTempEco, InStr(strTempEco, ")") - 1)
            End If
            
            '-----> If ECO is unique
            If Right(strECONum, 3) = "IMP" Then
                If strDocType = "Import" Then
                    strECONum = Left(strECONum, Len(strECONum) - 3)
                    txtType(intImport).Text = "ECO"
                    
                    txtValue(intImport).Text = Replace(strECONum, ".", "") & ".00"
                    
                    intImport = intImport + 1
                End If
            ElseIf Right(strECONum, 3) = "EXP" Then
                If strDocType = "Export" Then
                    strECONum = Left(strECONum, Len(strECONum) - 3)
                    txtType(intExport).Text = "ECO"
                    
                    txtValue(intExport).Text = Replace(strECONum, ".", "") & ".00"
                    
                    intExport = intExport + 1
                End If
            Else
                txtType(intExport + intImport).Text = "ECO"
                
                txtValue(intImport + intExport).Text = Replace(strECONum, ".", "") & ".00"
                
                If strDocType = "Export" Then
                    intExport = intExport + 1
                Else
                    intImport = intImport + 1
                End If
            End If

            If blnNewVersion Then
                If Len(strTempEco) > 0 Then
                    If InStr(1, strTempEco, " ") Then
                        If (intImport + intExport) > 2 Then
                            strTempEco = ""
                        Else
                            strTempEco = Right(strTempEco, Len(strTempEco) - InStr(strTempEco, " "))
                        End If
                    Else
                        strTempEco = ""
                    End If
                End If
            Else

                If (intImport + intExport) > 2 Then
                    strTempEco = ""
                Else
                    strTempEco = Right(strTempEco, Len(strTempEco) - InStr(strTempEco, ")"))
                End If
            End If
            ' *****   end.   *****
        Loop
        
        If UCase(strKluwerOut(6)) = "YES" Then
            If strDocType = "Import" Then
                txtType(intImport).Text = "DP"
            End If
        End If
    Else    '-----> If loaded as Normal
        blnKluwerClicked = True
        
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "IMPORT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
        ADORecordsetOpen strSQL, m_conTaric, m_rstImpGrid, adOpenKeyset, adLockOptimistic
        'Set m_rstImpGrid = m_conTaric.OpenRecordset(strSQL)
            
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "EXPORT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
        ADORecordsetOpen strSQL, m_conTaric, m_rstExpGrid, adOpenKeyset, adLockOptimistic
        'Set m_rstExpGrid = m_conTaric.OpenRecordset(strSQL)
        
        Dim blnAddNew As Boolean
        
        '-----> Find if a similar country already exists in Import table.
        With m_rstImpGrid
            blnAddNew = True
            blnAddNew = blnAddNew And (.EOF And .BOF)
            
            If Not blnAddNew Then
                If blnNewVersion Then
                    strTempEco = strKluwerOut(4)
                Else
                    strTempEco = strKluwerOut(5)
                End If
                    
                .MoveFirst
                .Find "[CTRY CODE] = " & Chr(39) & ProcessQuotes(strTempEco) & Chr(39) & " ", , adSearchForward
                        
                blnAddNew = .EOF
            End If
            
            If blnAddNew Then
                .AddNew
            End If
        End With
        
ImportCountryExist:
        
        With m_rstExpGrid
            '-----> Find if similar country already exists in Export table.
            blnAddNew = True
            blnAddNew = blnAddNew And (.EOF And .BOF)
            
            If Not blnAddNew Then
            
                If blnNewVersion Then
                    strTempEco = strKluwerOut(4)
                Else
                    strTempEco = strKluwerOut(5)
                End If
                
                .MoveFirst
                .Find "[CTRY CODE] = " & Chr(39) & ProcessQuotes(strTempEco) & Chr(39) & " ", , adSearchForward
                        
                blnAddNew = .EOF
                
            End If
            
            If blnAddNew Then
                .AddNew
            End If
        End With
        
ExportCountryExist:
        
        '-----> Taric code
        m_rstImpGrid![TARIC CODE] = txtCode.Text
        m_rstExpGrid![TARIC CODE] = txtCode.Text
        
        '-----> Country code
        If blnNewVersion Then
            'handle muna
            'm_rstImpGrid![CTRY CODE] = strKluwerOut(4)
            m_rstImpGrid![CTRY CODE] = IIf(strKluwerOut(4) = "", " ", strKluwerOut(4))
            
            'm_rstExpGrid![CTRY CODE] = strKluwerOut(4)
            m_rstExpGrid![CTRY CODE] = IIf(strKluwerOut(4) = "", " ", strKluwerOut(4))
        Else
            
            'm_rstImpGrid![CTRY CODE] = strKluwerOut(5)
            m_rstImpGrid![CTRY CODE] = IIf(strKluwerOut(5) = "", " ", strKluwerOut(5))
            
            'm_rstExpGrid![CTRY CODE] = strKluwerOut(5)
            m_rstExpGrid![CTRY CODE] = IIf(strKluwerOut(5) = "", " ", strKluwerOut(5))
        End If
        '-----> Attached documents
        
        ' ***** 10/22/02 *****
        If blnNewVersion Then
            strTempEco = strKluwerOut(5)
        Else
            strTempEco = strKluwerOut(4)
        End If
        
        'Do Until Len(strKluwerOut(4)) = 0
        Do Until Len(strTempEco) = 0
            If blnNewVersion Then
                strTempEco = Trim(strTempEco)
                If InStr(strTempEco, " ") Then
                    strECONum = Left(strTempEco, InStr(strTempEco, " ") - 1)
                Else
                    strECONum = strTempEco
                End If
                
'                strTempEco = Trim(strTempEco)
'                strECONum = Left(strTempEco, InStr(strTempEco, " ") - 1)
            Else
                '-----> Extract ECO numbers from line 4.
                'strKluwerOut(4) = Right(strKluwerOut(4), Len(strKluwerOut(4)) - 1)
                'strECONum = Left(strKluwerOut(4), InStr(strKluwerOut(4), ")") - 1)
                strTempEco = Right(strTempEco, Len(strTempEco) - 1)
                strECONum = Left(strTempEco, InStr(strTempEco, ")") - 1)
            End If
            
            If Right(strECONum, 3) = "IMP" Then    '-----> ECO for Import only.
                strECONum = Left(strECONum, Len(strECONum) - 3)
                
                For Each fldImport In m_rstImpGrid.Fields
                    If fldImport.Name = "N" & (intImport + 1) Then
                        fldImport.Value = "ECO"
                    End If
                Next
                
                For Each fldImport In m_rstImpGrid.Fields
                    If fldImport.Name = "Q" & (intImport + 1) Then
' ********** Commented September 10, 2001 **********
' ********** The following lines of code could be replaced by just one line.
'                        If InStr(strECONum, ".") = 0 Then
'                            fldImport.Value = strECONum & ".00"
'                        Else
'                            fldImport.Value = Left(strECONum, InStr(strECONum, ".") - 1) & Right(strECONum, Len(strECONum) - InStr(strECONum, ".")) & ".00"
'                        End If
' ********** End Comment ***************************
                        
                        fldImport.Value = Replace(strECONum, ".", "") & ".00"
                    End If
                Next
                
                intImport = intImport + 1
            ElseIf Right(strECONum, 3) = "EXP" Then    '-----> ECO for Export only.
                strECONum = Left(strECONum, Len(strECONum) - 3)
                
                For Each fldExport In m_rstExpGrid.Fields
                    If fldExport.Name = "N" & (intExport + 1) Then
                        fldExport.Value = "ECO"
                    End If
                Next
                
                For Each fldExport In m_rstExpGrid.Fields
                    If fldExport.Name = "Q" & (intExport + 1) Then
                        fldExport.Value = Replace(strECONum, ".", "") & ".00"
                    End If
                Next
                
                intExport = intExport + 1
            Else    '-----> ECO for both Import and Export.
                For Each fldImport In m_rstImpGrid.Fields
                    If fldImport.Name = "N" & (intImport + 1) Then
                        fldImport.Value = "ECO"
                    End If
                Next
                
                For Each fldImport In m_rstImpGrid.Fields
                    If fldImport.Name = "Q" & (intImport + 1) Then
                        fldImport.Value = Replace(strECONum, ".", "") & ".00"
                    End If
                Next
                
                intImport = intImport + 1
                
                For Each fldExport In m_rstExpGrid.Fields
                    If fldExport.Name = "N" & (intExport + 1) Then
                        fldExport.Value = "ECO"
                    End If
                Next
                
                For Each fldExport In m_rstExpGrid.Fields
                    If fldExport.Name = "Q" & (intExport + 1) Then
                        fldExport.Value = Replace(strECONum, ".", "") & ".00"
                    End If
                Next
                
                intExport = intExport + 1
            End If

            If blnNewVersion Then
                If Len(strTempEco) > 0 Then
                    If InStr(1, strTempEco, " ") Then
                        If (intImport + intExport) > 2 Then
                            strTempEco = ""
                        Else
                            strTempEco = Right(strTempEco, Len(strTempEco) - InStr(strTempEco, " "))
                        End If
                    Else
                        strTempEco = ""
                    End If
                End If
            Else
                If (intImport + intExport) > 2 Then
                    strTempEco = ""
                Else
                    strTempEco = Right(strTempEco, Len(strTempEco) - InStr(strTempEco, ")"))
                End If
            End If
        Loop
        
        '-----> Add DP in type of attached documents if yes.
        If UCase(strKluwerOut(6)) = "YES" And strDocType = "Import" Then
            For Each fldImport In m_rstImpGrid.Fields
                If fldImport.Name = "N" & (intImport + 1) Then
                    fldImport.Value = "DP"
                End If
            Next
        End If
        
        m_rstImpGrid.Update
        m_rstExpGrid.Update
        
        If blnNewVersion Then
            intKluwerCountry = strKluwerOut(4)
        Else
            intKluwerCountry = strKluwerOut(5)
        End If
        
        Call ImportAdd
        Call ExportAdd
    End If
End Sub

Private Sub DeleteKluwer()
    Dim strCommand As String

    '-----> Deletes Kluwer from database if the user does not click OK nor Apply.
    With m_rstImpGrid
        If Not (.EOF And .BOF) Then
            .MoveFirst
        
            Do Until .EOF
                If intKluwerCountry = ![CTRY CODE] Then
                    .Delete
                End If
                
                .MoveNext
            Loop
            
                strCommand = vbNullString
                strCommand = strCommand & "DELETE "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[IMPORT] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ProcessQuotes(intKluwerCountry) & Chr(39) & " "
            ExecuteNonQuery m_conTaric, strCommand
        End If
    End With
    
    With m_rstExpGrid
        If Not (.EOF And .BOF) Then
            .MoveFirst
        
            Do Until .EOF
                If intKluwerCountry = ![CTRY CODE] Then
                    .Delete
                End If
                
                .MoveNext
            Loop
            
                strCommand = vbNullString
                strCommand = strCommand & "DELETE "
                strCommand = strCommand & "* "
                strCommand = strCommand & "FROM "
                strCommand = strCommand & "[EXPORT] "
                strCommand = strCommand & "WHERE "
                strCommand = strCommand & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
                strCommand = strCommand & "AND "
                strCommand = strCommand & "[CTRY CODE] = " & Chr(39) & ProcessQuotes(intKluwerCountry) & Chr(39) & " "
            ExecuteNonQuery m_conTaric, strCommand
        End If
    End With
End Sub

Public Sub DefaultCheck()

    '-----> dissable chkdefault if there is only one record or no record exist
    '-----> set this country record to default
    Dim strSQL As String
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[" & GetTableToUse(strDocType) & "] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, m_rstDefault, adOpenKeyset, adLockOptimistic
    'Set m_rstDefault = m_conTaric.OpenRecordset(strSQL)
    With m_rstDefault
        If Not (.EOF And .BOF) Then
            .MoveLast
        End If
        
        Select Case .RecordCount
            Case 0
                chkDefault.Value = 1
                chkDefault.Enabled = False
                chkCommon.Enabled = False
            Case 1
                If ![CTRY CODE] = txtCtryCode.Text Then
                    chkDefault.Value = 1
                    chkDefault.Enabled = False
                    chkCommon.Enabled = False
                Else
                    chkDefault.Enabled = True
                End If
            Case Else
                chkDefault.Enabled = True
        End Select
    End With
    Set m_rstDefault = Nothing

End Sub

Public Function InCollection(IorE As String) As Boolean
    On Error GoTo AddMe
        If Not IsEmpty(colTaricCodes.Item(txtCode.Text & IorE)) Then
            InCollection = True
        End If
        
        Exit Function
AddMe:
        InCollection = False
End Function

Public Sub UpdateThisRecordset(ByVal AddNew As Boolean, ByRef rstCopy As ADODB.Recordset, ByRef rst As ADODB.Recordset, IorE As String)

    If AddNew Then
        rstCopy.AddNew
    'Else
    '    rstCopy.Edit
    End If
    
    Select Case IorE
        Case "I"
            With rst
                rstCopy![TARIC CODE] = txtCode.Text
                
                'handle muna - may 2, 2003
                'rstCopy![CTRY CODE] = ![CTRY CODE]
                rstCopy![CTRY CODE] = IIf(![CTRY CODE] = "", " ", ![CTRY CODE])
                
                '----> Save licence
                rstCopy![LIC REQD] = ![LIC REQD]
                rstCopy![MIN VALUE] = ![MIN VALUE]
                rstCopy![Min Value Curr] = ![Min Value Curr]
                
                '-----> Save attached documents
                rstCopy![N1] = ![N1]
                rstCopy![N2] = ![N2]
                rstCopy![N3] = ![N3]
                rstCopy![O1] = ![O1]
                rstCopy![O2] = ![O2]
                rstCopy![O3] = ![O3]
                rstCopy![P1] = ![P1]
                rstCopy![P2] = ![P2]
                rstCopy![P3] = ![P3]
                rstCopy![Q1] = ![Q1]
                rstCopy![Q2] = ![Q2]
                rstCopy![Q3] = ![Q3]
                            
                '----> Save special regime
                rstCopy![R1] = ![R1]
                rstCopy![R3] = ![R3]
                rstCopy![R5] = ![R5]
                rstCopy![R7] = ![R7]
                rstCopy![R9] = ![R9]
                rstCopy![R2] = ![R2]
                rstCopy![R4] = ![R4]
                rstCopy![R6] = ![R6]
                rstCopy![R8] = ![R8]
                rstCopy![RA] = ![RA]
                
                '-----> save usage
                rstCopy![DEF CODE] = ![DEF CODE]
                rstCopy![COMM CODE] = ![COMM CODE]
                
                rstCopy.Update
            End With
        
        Case "E"
            With rst
        
                rstCopy![TARIC CODE] = txtCode.Text
                
                'handle muna- may 2, 2003
                'rstCopy![CTRY CODE] = ![CTRY CODE]
                rstCopy![CTRY CODE] = IIf(![CTRY CODE] = "", " ", ![CTRY CODE])
                
                '----> Save licence
                rstCopy![LIC REQD] = ![LIC REQD]
                
                '-----> Save attached documents
                rstCopy![N1] = ![N1]
                rstCopy![N2] = ![N2]
                rstCopy![N3] = ![N3]
                rstCopy![O1] = ![O1]
                rstCopy![O2] = ![O2]
                rstCopy![O3] = ![O3]
                rstCopy![P1] = ![P1]
                rstCopy![P2] = ![P2]
                rstCopy![P3] = ![P3]
                rstCopy![Q1] = ![Q1]
                rstCopy![Q2] = ![Q2]
                rstCopy![Q3] = ![Q3]
                            
                '----> Save special regime
                rstCopy![R1] = ![R1]
                rstCopy![R3] = ![R3]
                rstCopy![R5] = ![R5]
                rstCopy![R7] = ![R7]
                rstCopy![R9] = ![R9]
                rstCopy![R2] = ![R2]
                rstCopy![R4] = ![R4]
                rstCopy![R6] = ![R6]
                rstCopy![R8] = ![R8]
                rstCopy![RA] = ![RA]
                
                '-----> save usage
                rstCopy![DEF CODE] = ![DEF CODE]
                rstCopy![COMM CODE] = ![COMM CODE]
                rstCopy.Update
        
        End With
    End Select
    
    If AddNew Then
        InsertRecordset m_conTaric, rstCopy, "IMPORT"
    Else
        UpdateRecordset m_conTaric, rstCopy, "IMPORT"
    End If
End Sub

Private Function IsCompletelyFilled() As Boolean

    IsCompletelyFilled = True
    
    If Len(Trim(txtCode.Text)) < 10 Then
        'MsgBox "Code cannot be less than 10 characters", vbInformation
        MsgBox Translate(2106), vbInformation
        IsCompletelyFilled = False
    
    End If
    
    If txtCtryCode.Text = "" And txtCtryCode.Visible = True Then
        MsgBox Translate(1030), vbInformation
        IsCompletelyFilled = False
    End If
    
End Function

Private Function GetTableToUse(DocType As String) As String
    Select Case DocType
        Case "Import"
            GetTableToUse = DocType
        Case "Export", "Transit", G_CONST_NCTS1_TYPE, G_CONST_NCTS2_TYPE
            GetTableToUse = "Export"
        Case Else
            GetTableToUse = "Export"
    End Select

End Function

Private Sub LoadValuesFromDB(TableName As String, strTaricCode As String)
    Dim rstTaric As ADODB.Recordset
    Dim strSQL As String
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[" & TableName & "] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[Taric Code] = " & Chr(39) & ProcessQuotes(strTaricCode) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, rstTaric, adOpenKeyset, adLockOptimistic
    'Set rstTaric = m_conTaric.OpenRecordset(strSQL)
    
    If Not (rstTaric.EOF And rstTaric.BOF) Then
        rstTaric.MoveFirst
        
        With rstTaric
            
            
            'rstTaric.Seek frm_taricpicklist.strCtryCode & " = [CTRY CODE]", "TARIC_CTRY"
            rstTaric.MoveFirst
            rstTaric.Find "[CTRY CODE] like " & frm_taricpicklist.strCtryCode, , adSearchForward
            
            If rstTaric.EOF Then
                rstTaric.MoveFirst
                rstTaric.Find "[COMM CODE] = true ", , adSearchForward
                
                If rstTaric.EOF Then
                    rstTaric.MoveFirst
                    rstTaric.Find "[DEF CODE] = true ", , adSearchForward
                    
                    If rstTaric.EOF Then
                        rstTaric.MoveFirst
                    End If
                End If
            End If
            
            txtCtryCode.Text = IIf(IsNull(![CTRY CODE]), "0", ![CTRY CODE])
            'attached documents
            'n1
            txtType(0).Text = IIf(IsNull(!N1), "", !N1)
            
            'n2
            txtType(1).Text = IIf(IsNull(!N2), "", !N2)
            
            'n3
            txtType(2).Text = IIf(IsNull(!N3), "", !N3)
            
            'for n1......
            'o1 - number
            txtNumber(0).Text = IIf(IsNull(!O1), "", !O1)
            
            'p1 - date
            txtDate(0).Text = IIf(IsNull(!P1), "", !P1)
            
            'q1 - value
            txtValue(0).Text = IIf(IsNull(!Q1), "", !Q1)
            
            'for n2......
            'o2 - number
            txtNumber(0).Text = IIf(IsNull(!O1), "", !O1)
            
            'p2 - date
            txtDate(0).Text = IIf(IsNull(!P1), "", !P1)
            
            'q2 - value
            txtValue(0).Text = IIf(IsNull(!Q1), "", !Q1)
            
            
            'for n3......
            'o1 - number
            txtNumber(0).Text = IIf(IsNull(!O1), "", !O1)
            
            'p1 - date
            txtDate(0).Text = IIf(IsNull(!P1), "", !P1)
            
            'q1 - value
            txtValue(0).Text = IIf(IsNull(!Q1), "", !Q1)
            
            'special regimes
            'left column
            'r1, r3, r5, r7, r9
            txtReg(0).Text = IIf(IsNull(!R1), "", !R1)
            txtReg(1).Text = IIf(IsNull(!R3), "", !R3)
            txtReg(2).Text = IIf(IsNull(!R5), "", !R5)
            txtReg(3).Text = IIf(IsNull(!R7), "", !R7)
            txtReg(4).Text = IIf(IsNull(!R9), "", !R9)
            
            'right column
            'r2, r4, r6, r8, ra
            txtRegValue(0).Text = IIf(IsNull(!R2), "", !R2)
            txtReg(1).Text = IIf(IsNull(!R3), "", !R3)
            txtReg(2).Text = IIf(IsNull(!R5), "", !R5)
            txtReg(3).Text = IIf(IsNull(!R7), "", !R7)
            txtReg(4).Text = IIf(IsNull(!R9), "", !R9)
    
            'lic reqd
            chkLicenceExp.Value = IIf(![LIC REQD] = True, 1, 0)
            
        End With
    End If

    ADORecordsetClose rstTaric
End Sub
