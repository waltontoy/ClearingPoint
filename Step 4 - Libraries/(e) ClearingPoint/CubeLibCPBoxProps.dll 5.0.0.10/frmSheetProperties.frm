VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSheetProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sheet Properties"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmSheetProperties.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "494"
   Begin VB.CheckBox chkIsItalic 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4770
      Picture         =   "frmSheetProperties.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   490
      Width           =   360
   End
   Begin VB.Frame fraSheetProps 
      Caption         =   "514"
      Height          =   1455
      Index           =   2
      Left            =   123
      TabIndex        =   35
      Tag             =   "514"
      Top             =   3345
      Width           =   5490
      Begin VB.Frame fraSheetProps 
         Caption         =   "184"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   3
         Left            =   150
         TabIndex        =   36
         Tag             =   "184"
         Top             =   240
         Width           =   5190
         Begin VB.TextBox txtBoxSample 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   3
            Left            =   3225
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   47
            TabStop         =   0   'False
            Tag             =   "651"
            Text            =   "Error Active"
            Top             =   540
            Width           =   1305
         End
         Begin VB.TextBox txtBoxSample 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   4
            Left            =   105
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "650"
            Text            =   "Error Inactive"
            Top             =   540
            Width           =   2715
         End
         Begin VB.TextBox txtBoxSample 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   2
            Left            =   2625
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   45
            TabStop         =   0   'False
            Tag             =   "649"
            Text            =   "Disabled"
            Top             =   210
            Width           =   1230
         End
         Begin VB.TextBox txtBoxSample 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   44
            TabStop         =   0   'False
            Tag             =   "638"
            Text            =   "Inactive"
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox txtBoxSample 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   0
            Left            =   105
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   43
            TabStop         =   0   'False
            Tag             =   "637"
            Text            =   "Active"
            Top             =   210
            Width           =   825
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D4"
            Height          =   195
            Index           =   0
            Left            =   4770
            TabIndex        =   48
            Top             =   255
            Width           =   210
         End
         Begin VB.Label lblBoxSample 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   4260
            TabIndex        =   42
            Top             =   210
            Width           =   480
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            Caption         =   "D9"
            Height          =   210
            Index           =   9
            Left            =   4575
            TabIndex        =   41
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            Caption         =   "D8"
            Height          =   210
            Index           =   8
            Left            =   2865
            TabIndex        =   40
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            Caption         =   "D3"
            Height          =   210
            Index           =   7
            Left            =   3900
            TabIndex        =   39
            Top             =   255
            Width           =   210
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            Caption         =   "D2"
            Height          =   210
            Index           =   6
            Left            =   2250
            TabIndex        =   38
            Top             =   255
            Width           =   210
         End
         Begin VB.Label lblBoxSample 
            AutoSize        =   -1  'True
            Caption         =   "D1"
            Height          =   210
            Index           =   5
            Left            =   975
            TabIndex        =   37
            Top             =   255
            Width           =   210
         End
      End
   End
   Begin VB.Frame fraSheetProps 
      Caption         =   "497"
      Height          =   2370
      Index           =   1
      Left            =   123
      TabIndex        =   17
      Tag             =   "497"
      Top             =   960
      Width           =   5490
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   9
         Left            =   4965
         TabIndex        =   11
         Top             =   1905
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   8
         Left            =   4965
         TabIndex        =   9
         Top             =   1545
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   7
         Left            =   4965
         TabIndex        =   7
         Top             =   1215
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   6
         Left            =   4965
         TabIndex        =   5
         Top             =   855
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   5
         Left            =   4965
         TabIndex        =   3
         Top             =   510
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   4
         Left            =   3405
         TabIndex        =   10
         Top             =   1905
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   3
         Left            =   3405
         TabIndex        =   8
         Top             =   1560
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   2
         Left            =   3405
         TabIndex        =   6
         Top             =   1200
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   1
         Left            =   3405
         TabIndex        =   4
         Top             =   855
         Width           =   285
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "..."
         Height          =   240
         Index           =   0
         Left            =   3405
         TabIndex        =   2
         Top             =   495
         Width           =   285
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   2460
         TabIndex        =   34
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   4020
         TabIndex        =   33
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   9
         Left            =   4020
         TabIndex        =   32
         Top             =   1875
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   8
         Left            =   4020
         TabIndex        =   31
         Top             =   1515
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   7
         Left            =   4020
         TabIndex        =   30
         Top             =   1185
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   6
         Left            =   4020
         TabIndex        =   29
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   2460
         TabIndex        =   28
         Top             =   1875
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   2460
         TabIndex        =   27
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H005F5F5F&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   2460
         TabIndex        =   26
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label lblColors 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   2460
         TabIndex        =   25
         Top             =   465
         Width           =   1260
      End
      Begin VB.Label lblInactiveBoxError 
         AutoSize        =   -1  'True
         Caption         =   "Box w/error Inactive"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Tag             =   "504"
         Top             =   1980
         Width           =   1440
      End
      Begin VB.Label lblActiveBoxError 
         AutoSize        =   -1  'True
         Caption         =   "Box w/error Active"
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Tag             =   "503"
         Top             =   1635
         Width           =   1320
      End
      Begin VB.Label lblDisabledBox 
         AutoSize        =   -1  'True
         Caption         =   "Disabled Box"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Tag             =   "502"
         Top             =   1275
         Width           =   930
      End
      Begin VB.Label lblInactiveBox 
         AutoSize        =   -1  'True
         Caption         =   "Inactive Box"
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Tag             =   "501"
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lblActiveBox 
         AutoSize        =   -1  'True
         Caption         =   "Active Box"
         Height          =   195
         Left            =   255
         TabIndex        =   20
         Tag             =   "500"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         Caption         =   "Backcolors"
         Height          =   195
         Left            =   4020
         TabIndex        =   19
         Tag             =   "498"
         Top             =   195
         Width           =   795
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         Caption         =   "Forecolors"
         Height          =   195
         Left            =   2460
         TabIndex        =   18
         Tag             =   "499"
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3123
      TabIndex        =   12
      Tag             =   "178"
      Top             =   4905
      Width           =   1200
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   4413
      TabIndex        =   13
      Tag             =   "179"
      Top             =   4905
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cdgSheetProps 
      Left            =   180
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   -2147483643
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Courier New"
   End
   Begin VB.Frame fraSheetProps 
      Caption         =   "495"
      Height          =   930
      Index           =   0
      Left            =   101
      TabIndex        =   14
      Tag             =   "495"
      Top             =   15
      Width           =   5505
      Begin VB.CheckBox chkIsBold 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         Picture         =   "frmSheetProperties.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   480
         Value           =   1  'Checked
         Width           =   360
      End
      Begin VB.ComboBox cboFontSizeList 
         Height          =   315
         Left            =   2895
         TabIndex        =   1
         Top             =   450
         Width           =   975
      End
      Begin VB.ComboBox cboFontNameList 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   450
         Width           =   2430
      End
      Begin VB.Label lblFontName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Tag             =   "496"
         Top             =   195
         Width           =   465
      End
      Begin VB.Label lblFontSize 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   2895
         TabIndex        =   15
         Tag             =   "409"
         Top             =   195
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmSheetProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TXT_ACTIVE_BOX = 0
Private Const TXT_INACTIVE_BOX = 1
Private Const TXT_DISABLED_BOX = 2
Private Const TXT_ERROR_ACTIVE_BOX = 3
Private Const TXT_ERROR_INACTIVE_BOX = 4

' fore colors
Private Const LBL_ACTIVE_BOX_FORE = 0
Private Const LBL_INACTIVE_BOX_FORE = 1
Private Const LBL_DISABLED_BOX_FORE = 2
Private Const LBL_ERROR_ACTIVE_BOX_FORE = 3
Private Const LBL_ERROR_INACTIVE_BOX_FORE = 4

' back colors
Private Const LBL_ACTIVE_BOX_BACK = 5
Private Const LBL_INACTIVE_BOX_BACK = 6
Private Const LBL_DISABLED_BOX_BACK = 7
Private Const LBL_ERROR_ACTIVE_BOX_BACK = 8
Private Const LBL_ERROR_INACTIVE_BOX_BACK = 9

Private Const CMD_OK = 0
Private Const CMD_CANCEL = 1

Dim mvarUserNo As Long
Dim mvarActiveConnection As ADODB.Connection
Dim mvarResourceHandle As Long
Dim mvarActiveLanguage As String
Dim CodisheetType As cpiCodiSheetTypeEnums
'Dim mvarActiveDocument As String

Private Sub cboFontNameList_Change()
    GetPreview
End Sub

Private Sub cboFontNameList_Click()
    GetPreview
End Sub

Private Sub cboFontSizeList_Change()
    GetPreview
End Sub

Private Sub cboFontSizeList_Click()
    GetPreview
End Sub

Private Sub chkIsBold_Click()

    If (chkIsBold.Value = vbChecked) Then
        txtBoxSample(TXT_ACTIVE_BOX).FontBold = True
        txtBoxSample(TXT_INACTIVE_BOX).FontBold = True
        txtBoxSample(TXT_DISABLED_BOX).FontBold = True
        txtBoxSample(TXT_ERROR_INACTIVE_BOX).FontBold = True
        txtBoxSample(TXT_ERROR_ACTIVE_BOX).FontBold = True
    ElseIf (chkIsBold.Value = vbUnchecked) Then
        txtBoxSample(TXT_ACTIVE_BOX).FontBold = False
        txtBoxSample(TXT_INACTIVE_BOX).FontBold = False
        txtBoxSample(TXT_DISABLED_BOX).FontBold = False
        txtBoxSample(TXT_ERROR_INACTIVE_BOX).FontBold = False
        txtBoxSample(TXT_ERROR_ACTIVE_BOX).FontBold = False
    End If
    

End Sub

Private Sub chkIsItalic_Click()

    If (chkIsItalic.Value = vbChecked) Then
        txtBoxSample(TXT_ACTIVE_BOX).FontItalic = True
        txtBoxSample(TXT_INACTIVE_BOX).FontItalic = True
        txtBoxSample(TXT_DISABLED_BOX).FontItalic = True
        txtBoxSample(TXT_ERROR_INACTIVE_BOX).FontItalic = True
        txtBoxSample(TXT_ERROR_ACTIVE_BOX).FontItalic = True
    ElseIf (chkIsItalic.Value = vbUnchecked) Then
        txtBoxSample(TXT_ACTIVE_BOX).FontItalic = False
        txtBoxSample(TXT_INACTIVE_BOX).FontItalic = False
        txtBoxSample(TXT_DISABLED_BOX).FontItalic = False
        txtBoxSample(TXT_ERROR_INACTIVE_BOX).FontItalic = False
        txtBoxSample(TXT_ERROR_ACTIVE_BOX).FontItalic = False
    End If

End Sub

Private Sub cmdColors_Click(Index As Integer)

    On Error GoTo EarlyExit

    Dim varOriginalColor As Variant

    varOriginalColor = lblColors(Index).BackColor

    cdgSheetProps.ShowColor

    If Index < 5 Then
        If lblColors(Index + 5).BackColor = cdgSheetProps.Color Then
            lblColors(Index).BackColor = varOriginalColor
        Else
            lblColors(Index).BackColor = cdgSheetProps.Color
        End If
    Else
        If lblColors(Index - 5).BackColor = cdgSheetProps.Color Then
            lblColors(Index).BackColor = varOriginalColor
        Else
            lblColors(Index).BackColor = cdgSheetProps.Color
        End If
    End If

    GetPreview

    Exit Sub

EarlyExit:

End Sub

Private Sub cmdTransact_Click(Index As Integer)
    
    If (Index = 0) Then
        SaveData
    End If

    Unload Me

End Sub

Private Sub Form_Load()
    
    Dim intFontCtr As Integer
    
    ' load resources here
    'Call LoadResStrings(Me, True)
    ' translate caption
    TranslateCaptions

    ' load fonts here
    With cboFontNameList
        For intFontCtr = 0 To Screen.FontCount - 1
            .AddItem Screen.Fonts(intFontCtr)
        Next

        ' .ListIndex = 23
    End With

    With cboFontSizeList
        .AddItem 8
        .AddItem 9

        For intFontCtr = 10 To 72 Step 2
            .AddItem intFontCtr
        Next

        ' .ListIndex = 0
    End With

    InitColorsFont
    GetPreview
    
End Sub
'
Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub InitColorsFont()

    Dim intFontCtr As Integer

    Dim clsSHEET_PROPERTIES_Table As cpiSHEET_PROPS_Tbl
    Dim clsSHEET_PROPERTIES_Tables As cpiSHEET_PROPS_Tbls
    Dim blnRecordFound As Boolean

    Set clsSHEET_PROPERTIES_Table = New cpiSHEET_PROPS_Tbl
    Set clsSHEET_PROPERTIES_Tables = New cpiSHEET_PROPS_Tbls

    clsSHEET_PROPERTIES_Table.FIELD_user_no = mvarUserNo
    blnRecordFound = clsSHEET_PROPERTIES_Tables.GetRecord(mvarActiveConnection, clsSHEET_PROPERTIES_Table)

    If (blnRecordFound = False) Then
            '
        ' Forecolors
        lblColors(LBL_ACTIVE_BOX_FORE).BackColor = &H0         '-2147483640
        lblColors(LBL_INACTIVE_BOX_FORE).BackColor = &H0         '-2147483640
        lblColors(LBL_DISABLED_BOX_FORE).BackColor = &H5F5F5F    '6250335
        lblColors(LBL_ERROR_ACTIVE_BOX_FORE).BackColor = &H0         '-2147483640
        lblColors(LBL_ERROR_INACTIVE_BOX_FORE).BackColor = &H0         '-2147483640
        
        ' Backcolors
        lblColors(LBL_ACTIVE_BOX_BACK).BackColor = &H80FFFF    '8454143
        lblColors(LBL_INACTIVE_BOX_BACK).BackColor = &HFFFFFF    '9981440
        lblColors(LBL_DISABLED_BOX_BACK).BackColor = &HC0C0C0    '12632256
        lblColors(LBL_ERROR_ACTIVE_BOX_BACK).BackColor = &HFF80FF    '16744703
        lblColors(LBL_ERROR_INACTIVE_BOX_BACK).BackColor = &H8080FF    '8421631
    
        'chkIsBold.BevelOuter = 1
        'chkIsBold.BackColor = &H80000014
        'chkIsItalic.BevelOuter = 2
        'chkIsItalic.BackColor = &H8000000F
    
        cboFontNameList.ListIndex = 0
        ' cboFontNameList.Text = "Courier New"
    
        For intFontCtr = 0 To cboFontNameList.ListCount - 1
6            If Trim(cboFontNameList.List(intFontCtr)) = "Courier New" Then
                cboFontNameList.ListIndex = intFontCtr
                Exit For
            End If
        Next
    
        cboFontSizeList.ListIndex = 1
    
    
    ElseIf (blnRecordFound = True) Then
        
        lblColors(LBL_ACTIVE_BOX_FORE).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_active_box '![FC active box]
        lblColors(LBL_INACTIVE_BOX_FORE).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_inactive_box ' ![FC inactive box]
        lblColors(LBL_DISABLED_BOX_FORE).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_disabled_box ' ![FC disabled box]
        lblColors(LBL_ERROR_ACTIVE_BOX_FORE).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_werror_active ' ![FC werror active]
        lblColors(LBL_ERROR_INACTIVE_BOX_FORE).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_werror_inactive ' ![FC werror inactive]
        lblColors(LBL_ACTIVE_BOX_BACK).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_active_box ' ![BC active box]
        lblColors(LBL_INACTIVE_BOX_BACK).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_inactive_box ' ![BC inactive box]
        lblColors(LBL_DISABLED_BOX_BACK).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_disabled_box ' ![BC disabled box]
        lblColors(LBL_ERROR_ACTIVE_BOX_BACK).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_werror_active ' ![BC werror active]
        lblColors(LBL_ERROR_INACTIVE_BOX_BACK).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_werror_inactive ' ![BC werror inactive]
    
        If (clsSHEET_PROPERTIES_Table.FIELD_bold = True) Then
            chkIsBold.Value = vbChecked
        ElseIf (clsSHEET_PROPERTIES_Table.FIELD_bold = False) Then
            chkIsBold.Value = vbUnchecked
        End If
    
        If (clsSHEET_PROPERTIES_Table.FIELD_italic = True) Then
            chkIsItalic.Value = vbChecked
        ElseIf (clsSHEET_PROPERTIES_Table.FIELD_italic = False) Then
            chkIsItalic.Value = vbUnchecked
        End If
    
            With cboFontNameList
                For intFontCtr = 0 To .ListCount - 1
                    If Trim(.List(intFontCtr)) = Trim(clsSHEET_PROPERTIES_Table.FIELD_font_name) Then
                        .ListIndex = intFontCtr
                        Exit For
                    End If
                Next intFontCtr
            End With
        
            ' cboFontNameList.Text = ![font name]
        
            With cboFontSizeList
                For intFontCtr = 0 To .ListCount - 1
                    If Trim(.List(intFontCtr)) = Trim(clsSHEET_PROPERTIES_Table.FIELD_size) Then
                        .ListIndex = intFontCtr
                        Exit For
                    End If
                Next intFontCtr
            End With
        
    End If

    Set clsSHEET_PROPERTIES_Table = Nothing
    Set clsSHEET_PROPERTIES_Tables = Nothing

End Sub

' ??????????? to be continued here... 21-Aug-2003
Private Sub LoadData()
    
    'Dim intFontCtr As Integer

    Dim clsSHEET_PROPERTIES_Table As cpiSHEET_PROPS_Tbl
    Dim clsSHEET_PROPERTIES_Tables As cpiSHEET_PROPS_Tbls
    Dim blnRecordFound As Boolean
    
    Set clsSHEET_PROPERTIES_Table = New cpiSHEET_PROPS_Tbl
    Set clsSHEET_PROPERTIES_Tables = New cpiSHEET_PROPS_Tbls
    
    clsSHEET_PROPERTIES_Table.FIELD_user_no = mvarUserNo
    
    blnRecordFound = clsSHEET_PROPERTIES_Tables.GetRecord(mvarActiveConnection, clsSHEET_PROPERTIES_Table)

    
    If (blnRecordFound = True) Then
    
        Dim intFontCtr As Integer
        For intFontCtr = 0 To (cboFontNameList.ListCount - 1)
            If (cboFontNameList.List(intFontCtr) = clsSHEET_PROPERTIES_Table.FIELD_font_name) Then
                cboFontNameList.ListIndex = intFontCtr
            End If
        Next intFontCtr
        
        Dim intFontSizeCtr As Integer
        
        For intFontSizeCtr = 0 To (cboFontSizeList.ListCount - 1)
            If (CInt(cboFontSizeList.List(intFontSizeCtr)) = clsSHEET_PROPERTIES_Table.FIELD_size) Then
                cboFontSizeList.ListIndex = intFontSizeCtr
            End If
        Next intFontSizeCtr
        
        'clsSHEET_PROPERTIES_Table.FIELD_font_name = cboFontNameList.Text
        ' ![Font Name] = cboFontNameList.Text
        'clsSHEET_PROPERTIES_Table.FIELD_size = IIf(Val(cboFontSizeList.Text) = 0, 9, cboFontSizeList.Text)
        ' ![Size] = IIf(Val(cboFontSizeList.Text) = 0, 9, cboFontSizeList.Text)
        
        If (clsSHEET_PROPERTIES_Table.FIELD_bold = True) Then
            chkIsBold.Value = vbChecked
        ElseIf (clsSHEET_PROPERTIES_Table.FIELD_bold = False) Then
            chkIsBold.Value = vbUnchecked
        End If
    
        If (clsSHEET_PROPERTIES_Table.FIELD_italic = True) Then
            chkIsItalic.Value = vbChecked
        ElseIf (clsSHEET_PROPERTIES_Table.FIELD_italic = False) Then
            chkIsItalic.Value = vbUnchecked
        End If
        
        ' save forecolor
        '![FC active box] = Val(lblColors(0).BackColor)
        lblColors(0).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_active_box
        '![FC inactive box] = Val(lblColors(1).BackColor)
        lblColors(1).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_inactive_box
        '![FC disabled box] = Val(lblColors(2).BackColor)
        lblColors(2).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_disabled_box
        '![FC werror active] = Val(lblColors(3).BackColor)
        lblColors(3).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_werror_active
        '![FC werror inactive] = Val(lblColors(4).BackColor)
        lblColors(4).BackColor = clsSHEET_PROPERTIES_Table.FIELD_FC_werror_inactive
        
        ' save back color
        '![BC active box] = Val(lblColors(5).BackColor)
        lblColors(5).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_active_box
        '![BC inactive box] = Val(lblColors(6).BackColor)
        lblColors(6).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_inactive_box
        '![BC disabled box] = Val(lblColors(7).BackColor)
        lblColors(7).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_disabled_box
        '![BC werror active] = Val(lblColors(8).BackColor)
        lblColors(8).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_werror_active
        '![BC werror inactive] = Val(lblColors(9).BackColor)
        lblColors(9).BackColor = clsSHEET_PROPERTIES_Table.FIELD_BC_werror_inactive
        
    End If
        
'        If (blnRecordFound = False) Then
'            clsSHEET_PROPERTIES_Tables.AddRecord mvarActiveConnection, clsSHEET_PROPERTIES_Table
'        ElseIf (blnRecordFound = True) Then
'            clsSHEET_PROPERTIES_Tables.ModifyRecord mvarActiveConnection, clsSHEET_PROPERTIES_Table
'        End If


    Set clsSHEET_PROPERTIES_Table = Nothing
    Set clsSHEET_PROPERTIES_Tables = Nothing

End Sub

'
Private Sub SaveData()
    
    Dim intFontCtr As Integer

    Dim clsSHEET_PROPERTIES_Table As cpiSHEET_PROPS_Tbl
    Dim clsSHEET_PROPERTIES_Tables As cpiSHEET_PROPS_Tbls
    Dim blnRecordFound As Boolean
    
    Set clsSHEET_PROPERTIES_Table = New cpiSHEET_PROPS_Tbl
    Set clsSHEET_PROPERTIES_Tables = New cpiSHEET_PROPS_Tbls
    
    clsSHEET_PROPERTIES_Table.FIELD_user_no = mvarUserNo
    
    blnRecordFound = clsSHEET_PROPERTIES_Tables.GetRecord(mvarActiveConnection, clsSHEET_PROPERTIES_Table)

    clsSHEET_PROPERTIES_Table.FIELD_font_name = cboFontNameList.Text
    ' ![Font Name] = cboFontNameList.Text
    clsSHEET_PROPERTIES_Table.FIELD_size = IIf(Val(cboFontSizeList.Text) = 0, 9, cboFontSizeList.Text)
    ' ![Size] = IIf(Val(cboFontSizeList.Text) = 0, 9, cboFontSizeList.Text)

    If (chkIsBold.Value = vbChecked) Then
        clsSHEET_PROPERTIES_Table.FIELD_bold = True
    ElseIf (chkIsBold.Value = vbUnchecked) Then
        clsSHEET_PROPERTIES_Table.FIELD_bold = False
    End If

    If (chkIsItalic.Value = vbChecked) Then
        clsSHEET_PROPERTIES_Table.FIELD_italic = True
    ElseIf (chkIsItalic.Value = vbUnchecked) Then
        clsSHEET_PROPERTIES_Table.FIELD_italic = False
    End If

    ' save forecolor
    '![FC active box] = Val(lblColors(0).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_FC_active_box = Val(lblColors(0).BackColor)
    '![FC inactive box] = Val(lblColors(1).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_FC_inactive_box = Val(lblColors(1).BackColor)
    '![FC disabled box] = Val(lblColors(2).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_FC_disabled_box = Val(lblColors(2).BackColor)
    '![FC werror active] = Val(lblColors(3).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_FC_werror_active = Val(lblColors(3).BackColor)
    '![FC werror inactive] = Val(lblColors(4).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_FC_werror_inactive = Val(lblColors(4).BackColor)
    
    ' save back color
    '![BC active box] = Val(lblColors(5).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_BC_active_box = Val(lblColors(5).BackColor)
    '![BC inactive box] = Val(lblColors(6).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_BC_inactive_box = Val(lblColors(6).BackColor)
    '![BC disabled box] = Val(lblColors(7).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_BC_disabled_box = Val(lblColors(7).BackColor)
    '![BC werror active] = Val(lblColors(8).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_BC_werror_active = Val(lblColors(8).BackColor)
    '![BC werror inactive] = Val(lblColors(9).BackColor)
    clsSHEET_PROPERTIES_Table.FIELD_BC_werror_inactive = Val(lblColors(9).BackColor)
    
    
    If (blnRecordFound = False) Then
        clsSHEET_PROPERTIES_Tables.AddRecord mvarActiveConnection, clsSHEET_PROPERTIES_Table
    ElseIf (blnRecordFound = True) Then
        clsSHEET_PROPERTIES_Tables.ModifyRecord mvarActiveConnection, clsSHEET_PROPERTIES_Table
    End If

    Set clsSHEET_PROPERTIES_Table = Nothing
    Set clsSHEET_PROPERTIES_Tables = Nothing

End Sub
'
Private Sub GetPreview()

    Dim intBoxSampleCtr As Integer

    For intBoxSampleCtr = 0 To 4
    
        txtBoxSample(intBoxSampleCtr).BackColor = lblColors(intBoxSampleCtr + 5).BackColor
        txtBoxSample(intBoxSampleCtr).ForeColor = lblColors(intBoxSampleCtr).BackColor
        If (cboFontNameList.Text <> "") Then
            txtBoxSample(intBoxSampleCtr).FontName = cboFontNameList.Text  'cboFontNameList.Text = "Courier New"
        ElseIf (cboFontNameList.Text = "") Then
            txtBoxSample(intBoxSampleCtr).FontName = "Courier New"
        End If
        'txtBoxSample(intBoxSampleCtr).FontName = cboFontNameList.Text
        txtBoxSample(intBoxSampleCtr).FontSize = IIf(Val(cboFontSizeList.Text) = 0, 9, Val(cboFontSizeList.Text))

        If (chkIsItalic.Value = vbChecked) Then
            txtBoxSample(intBoxSampleCtr).FontItalic = True
        ElseIf (chkIsItalic.Value = vbUnchecked) Then
            txtBoxSample(intBoxSampleCtr).FontItalic = False
        End If

        If (chkIsBold.Value = vbChecked) Then
            txtBoxSample(intBoxSampleCtr).FontBold = True
        ElseIf (chkIsBold.Value = Unchecked) Then
            txtBoxSample(intBoxSampleCtr).FontBold = False
        End If
        
    Next intBoxSampleCtr

    '???
    lblBoxSample(10).BackColor = lblColors(6).BackColor
    
End Sub

Public Function ShowForm(ByRef OwnerForm As Form, _
                                            ByRef ActiveConnection As ADODB.Connection, _
                                            ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                            ByRef UserNo As Long, _
                                            ByRef ActiveLanguage As String, _
                                            ByRef ResourceHandle As Long) As Boolean
'
    ' load frmBoxProperties here

    'mvarActiveBoxCode = ActiveBoxCode
    'mvarActiveDocument = ActiveDocument
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

Private Function TranslateCaptions() As Boolean

    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If (ctl.Tag <> "") Then
            If (LCase$(TypeName(ctl)) = "textbox") Then
                ctl.Text = Translate_B(ctl.Tag, mvarResourceHandle)
            Else 'If (LCase$(TypeName(ctl)) = "textbox") Then
                ctl.Caption = Translate_B(ctl.Tag, mvarResourceHandle)
            End If
        End If
    Next ' ctl

'    Caption = " " & mvarActiveBoxCode & " - " & Translate_B(Me.Caption, mvarResourceHandle)
'    tabPlatform.TabCaption(0) = Translate_B(tabPlatform.TabCaption(0), mvarResourceHandle)
'    tabPlatform.TabCaption(1) = Translate_B(tabPlatform.TabCaption(1), mvarResourceHandle)
'    tabPlatform.TabCaption(2) = Translate_B(tabPlatform.TabCaption(2), mvarResourceHandle)
'    tabPlatform.TabCaption(3) = Translate_B(tabPlatform.TabCaption(3), mvarResourceHandle)
'    tabPlatform.TabCaption(4) = Translate_B(tabPlatform.TabCaption(4), mvarResourceHandle)
    '
End Function


' :-) ???

' set new

' load new color settings here ->



' save new color settings here ->




' ??? :-)





'
' (-:  testing na!!!! :-)
';

