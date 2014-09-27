VERSION 5.00
Begin VB.Form frm_taricdetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TARIC - Add/Modify Import Detail"
   ClientHeight    =   5070
   ClientLeft      =   4770
   ClientTop       =   3195
   ClientWidth     =   7545
   Icon            =   "frm_taricDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Attached documents"
      Height          =   1575
      Left            =   120
      TabIndex        =   39
      Tag             =   "812"
      Top             =   2160
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Height          =   1060
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   4395
         TabIndex        =   40
         Top             =   360
         Width           =   4455
         Begin VB.TextBox txtType 
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   2
            Left            =   0
            MaxLength       =   5
            TabIndex        =   55
            Text            =   "0"
            Top             =   750
            Width           =   750
         End
         Begin VB.TextBox txtType 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   0
            MaxLength       =   5
            TabIndex        =   54
            Text            =   "0"
            Top             =   510
            Width           =   750
         End
         Begin VB.TextBox txtType 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   53
            Text            =   "0"
            Top             =   240
            Width           =   750
         End
         Begin VB.CommandButton cmdType 
            Caption         =   "..."
            Height          =   250
            Index           =   0
            Left            =   720
            TabIndex        =   52
            Top             =   250
            Width           =   255
         End
         Begin VB.CommandButton cmdType 
            Caption         =   "..."
            Height          =   250
            Index           =   1
            Left            =   720
            TabIndex        =   51
            Top             =   510
            Width           =   255
         End
         Begin VB.CommandButton cmdType 
            Caption         =   "..."
            Height          =   250
            Index           =   2
            Left            =   720
            TabIndex        =   50
            Top             =   760
            Width           =   255
         End
         Begin VB.TextBox txtNumber 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   7
            TabIndex        =   49
            Text            =   "0"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtNumber 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   7
            TabIndex        =   48
            Text            =   "0"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtNumber 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   2
            Left            =   960
            MaxLength       =   7
            TabIndex        =   47
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   46
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   45
            Text            =   "0"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   2
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   44
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   3240
            MaxLength       =   12
            TabIndex        =   43
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3240
            MaxLength       =   12
            TabIndex        =   42
            Text            =   "0"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   2
            Left            =   3240
            MaxLength       =   12
            TabIndex        =   41
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   59
            Tag             =   "439"
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Number"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Tag             =   "838"
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   57
            Tag             =   "747"
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Value"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   56
            Tag             =   "451"
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Special regime"
      Height          =   2175
      Left            =   4920
      TabIndex        =   20
      Tag             =   "851"
      Top             =   2160
      Width           =   2415
      Begin VB.PictureBox Picture2 
         Height          =   1545
         Left            =   120
         ScaleHeight     =   1485
         ScaleWidth      =   1875
         TabIndex        =   21
         Top             =   360
         Width           =   1935
         Begin VB.TextBox txtreg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   0
            MaxLength       =   2
            TabIndex        =   36
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtreg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   0
            MaxLength       =   2
            TabIndex        =   35
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtreg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   0
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtreg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   0
            MaxLength       =   2
            TabIndex        =   33
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtreg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   290
            Index           =   4
            Left            =   0
            MaxLength       =   2
            TabIndex        =   32
            Text            =   "0"
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "..."
            Height          =   275
            Index           =   0
            Left            =   720
            TabIndex        =   31
            Top             =   250
            Width           =   255
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "..."
            Height          =   275
            Index           =   1
            Left            =   720
            TabIndex        =   30
            Top             =   500
            Width           =   255
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "..."
            Height          =   275
            Index           =   2
            Left            =   720
            TabIndex        =   29
            Top             =   730
            Width           =   255
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "..."
            Height          =   275
            Index           =   3
            Left            =   720
            TabIndex        =   28
            Top             =   980
            Width           =   255
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "..."
            Height          =   275
            Index           =   4
            Left            =   720
            TabIndex        =   27
            Top             =   1220
            Width           =   255
         End
         Begin VB.TextBox txtRegValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   6
            TabIndex        =   26
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtRegValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   6
            TabIndex        =   25
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtRegValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   960
            MaxLength       =   6
            TabIndex        =   24
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtRegValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   6
            TabIndex        =   23
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtRegValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   290
            Index           =   4
            Left            =   960
            MaxLength       =   6
            TabIndex        =   22
            Text            =   "0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reg"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Tag             =   "843"
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Value"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            TabIndex        =   37
            Tag             =   "451"
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.Frame frmExport 
      Caption         =   "Export Licence"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Tag             =   "808"
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CheckBox chklicenceExp 
         Caption         =   "Required"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Tag             =   "589"
         Top             =   320
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Tag             =   "179"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Tag             =   "180"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Tag             =   "178"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Usage"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Tag             =   "864"
      Top             =   3720
      Width           =   4695
      Begin VB.CheckBox chkCommon 
         Caption         =   "Common"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Tag             =   "821"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "480"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmImport 
      Caption         =   "Import Licence"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Tag             =   "807"
      Top             =   1320
      Width           =   7215
      Begin VB.ComboBox cboCurrency 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   19
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtLimit 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         MaxLength       =   14
         TabIndex        =   9
         Top             =   270
         Width           =   2175
      End
      Begin VB.CheckBox chkLicenceImp 
         Caption         =   "Required if value exceeds"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "844"
         Top             =   320
         Width           =   2895
      End
      Begin VB.Label lblCurrency 
         Caption         =   "EUR"
         Height          =   255
         Left            =   5520
         TabIndex        =   60
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Tag             =   "269"
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdCountry 
         Caption         =   "..."
         Height          =   280
         Left            =   2230
         TabIndex        =   18
         Top             =   605
         Width           =   320
      End
      Begin VB.TextBox txtCtry 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtCtryCode 
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Country"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Tag             =   "822"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "820"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_taricdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_conSADBEL As ADODB.Connection      'Private datCountry As DAO.Database
Private m_conTaric As ADODB.Connection       'Private datDetail As DAO.Database

Dim m_rstDetail As ADODB.Recordset
Dim m_rstCountry As ADODB.Recordset
Dim m_rstDefault As ADODB.Recordset
Dim m_rstEmptyFieldValues As ADODB.Recordset


Dim strLangOfDesc As String
Dim strDocType As String


Dim strDetailLoaded As String
Dim strDetailType As String

'-----> to picklist
Public blnTaricDetail As Boolean

Dim strCountry As String

Dim blnCopy As Boolean

Dim blnUncheck As Boolean

'----> for getsetting in registry
Dim strRegReturn As String

'-----> for the default empty fields in the settings
Dim colEmptyFieldValues As Collection

'-----> minimum value and currency in import licence
' Dim strLicCurr As String
Dim dblMinVal As Double

Const Display = "##########0.00"

Private Sub chkDefault_Click()
    
    If chkDefault.Value = 1 Then
        chkCommon.Enabled = True
        blnUncheck = False
    Else
        chkCommon.Enabled = False
        chkCommon.Value = 0
        blnUncheck = True
    End If
End Sub

Private Sub chkLicenceImp_Click()
    Dim strLimit As String
    Dim lngCtr As Long
    Dim strDecimal As String

    If chkLicenceImp.Value = 1 Then
    
        'Changes other delimiter types to '.'
        strLimit = dblMinVal
        For lngCtr = 1 To Len(strLimit) + 1
            If InStr("0123456789.", Mid(strLimit, lngCtr, 1)) = 0 Then
                strDecimal = Mid(strLimit, lngCtr, 1)
                Exit For
            End If
        Next
        txtLimit.Text = Replace(strLimit, strDecimal, ".")
        'Appends necessary decimal values since double erased them.
        Call txtLimit_LostFocus
        
        txtLimit.Enabled = True
    Else
        txtLimit.Text = ""
        txtLimit.Enabled = False
    End If
End Sub

Private Sub cmdApply_Click()

    frm_taricdetail.MousePointer = 11
    
    SaveOptions
    
    If strDetailType = "Import" Then
        frm_taricmain.ImportAdd
    ElseIf strDetailType = "Export" Then
        frm_taricmain.ExportAdd
    End If
    
    frm_taricdetail.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()

    Unload frm_taricdetail

End Sub

Private Sub cmdCountry_Click()
    Dim strBoxProp As String
    Dim strPickVal As String
    
    '----> Use default string used in codisheet
    If strDetailType = "Import" Then
        strBoxProp = "C1#*#H20#*#1#*#" & strLangOfDesc & "#*#" & "Import#*#True#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    Else
        strBoxProp = "C2#*#H21#*#1#*#" & strLangOfDesc & "#*#" & "Export#*#True#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    End If
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    
    '-----> to skip a procedure in picklist
    blnTaricDetail = True
    g_blnMultiplePick = False
    
    frm_picklist.Show vbModal, Me
    
    blnTaricDetail = False
    
    strPickVal = GetSetting(App.Title, strRegReturn, "Pick_TARIC_PL")
    
    '-----> get the country code
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then
        txtCtryCode.Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
    End If
    
    DeleteSetting App.Title, strRegReturn, "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"
End Sub

Private Sub cmdOK_Click()

    '-----> Save and Exit
    
    frm_taricdetail.MousePointer = 11
    
    SaveOptions
    
    If strDetailType = "Import" Then
        frm_taricmain.ImportAdd
    ElseIf strDetailType = "Export" Then
        frm_taricmain.ExportAdd
    End If
    
    Unload frm_taricdetail

End Sub

Private Sub cmdReg_Click(Index As Integer)
    Dim strBoxProp As String
    Dim strPickVal As String
    
    '----> Use default string used in codisheet
    strBoxProp = "R" & (1 + (Index * 2)) & "#*#D" & (26 + (Index * 2)) & "#*#1#*#" & _
    strLangOfDesc & "#*#" & strDetailType & "#*#True#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    
    blnTaricDetail = True
    g_blnMultiplePick = False
    
    frm_picklist.Show vbModal, Me
    
    blnTaricDetail = False
    
    strPickVal = GetSetting(App.Title, strRegReturn, "Pick_TARIC_PL")
    '-----> get the regime
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then
        txtReg(Index).Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
    End If
    
    DeleteSetting App.Title, strRegReturn, "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"

End Sub

Private Sub cmdType_Click(Index As Integer)
    Dim strBoxProp As String
    Dim strPickVal As String
    
    
    '----> Use default string used in codisheet
    strBoxProp = "N" & (1 + Index) & "#*#D" & (14 + (Index * 4)) & "#*#1#*#" & strLangOfDesc & _
    "#*#" & strDetailType & "#*#True#*#TARIC#*#PL#*#0#*#0#*#8454143#*#-2147483640#*#"
    
    '----> Save to registry
    SaveSetting App.Title, "Settings", "BoxProperty", strBoxProp
    
    blnTaricDetail = True
    g_blnMultiplePick = False
    
    frm_picklist.Show vbModal, Me
    
    blnTaricDetail = False
    
    strPickVal = GetSetting(App.Title, strRegReturn, "Pick_TARIC_PL")
    '-----> get the type
    If Not Left(strPickVal, InStr(strPickVal, "%") - 1) = 0 Then
        txtType(Index).Text = Left(strPickVal, InStr(strPickVal, "%") - 1)
    End If
    
    DeleteSetting App.Title, strRegReturn, "Pick_TARIC_PL"
    DeleteSetting App.Title, "Settings", "BoxProperty"

End Sub

Private Sub Form_Load()

    '-----> To convert captions to default language
    Call LoadResStrings(Me, True)
    
    strDetailLoaded = frm_taricmain.strDetailLoaded
    strDetailType = frm_taricmain.strDetailType
    strLangOfDesc = frm_taricmain.strLangOfDesc
    
    Dim strSQL As String
    
    blnCopy = False
    blnTaricDetail = False
    
    frm_taricmain.MousePointer = 0
    
    ADOConnectDB m_conTaric, g_objDataSourceProperties, DBInstanceType_DATABASE_TARIC
    'OpenDAODatabase m_conTaric, cAppPath, "mdb_taric.mdb"
                            
    ADOConnectDB m_conSADBEL, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    'OpenDAODatabase m_conSADBEL, cAppPath, "mdb_sadbel.mdb"
                            
    '-----> Check if called from import or export
    '-----> Load default values for either import or export
    
    If strDetailType = "Import" Then
        '-----> Open Import DB set recordset
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "IMPORT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(frm_taricmain.txtCode.Text) & Chr(39) & " "
        ADORecordsetOpen strSQL, m_conTaric, m_rstDetail, adOpenKeyset, adLockOptimistic
        'Set m_rstDetail = m_conTaric.OpenRecordset(strSQL)
        
        frm_taricdetail.Caption = Translate(854)
        frmImport.Visible = True
        '-----> Add items to combo box of currency
    '    cboCurrency.AddItem "BEF"
    '    cboCurrency.AddItem "LUF"
    '    cboCurrency.AddItem "EUR"
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
    '                dblMinVal = 5000#
                    dblMinVal = 5000 / 40.3399
                End If
            End With
            ADORecordsetClose rstProperties

    ElseIf strDetailType = "Export" Then
        '-----> Open Export DB set recordset
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "EXPORT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(frm_taricmain.txtCode.Text) & Chr(39) & " "
        ADORecordsetOpen strSQL, m_conTaric, m_rstDetail, adOpenKeyset, adLockOptimistic
        'Set m_rstDetail = m_conTaric.OpenRecordset(strSQL)
        frmExport.Visible = True
        frm_taricdetail.Caption = Translate(855)
    End If
    
    '-----> Check what type of detail is called from main
    Select Case strDetailLoaded
        Case "New Country" '-----> Form is Blank
            DataFill
            BlankOut
        Case "Copy" '-----> Fill everything with data from main except for Country
            DataFill
            txtCtry.Text = ""
            txtCtryCode.Text = ""
            chkDefault.Value = 0
            blnCopy = True
        Case "Change" '-----> Fill everything with data from main
            DataFill
    End Select
    
    Select Case strDetailType
        Case "Import"
            strRegReturn = "CodiSheet"
        Case "Export"
            strRegReturn = "ExSheet"
    End Select

End Sub

Private Sub DataFill()
    Dim intSpaceLoc As Integer
    
    '-----> Load code and description
    txtCode.Text = frm_taricmain.txtCode.Text
    
    If strLangOfDesc = "Dutch" Then
        txtDesc.Text = frm_taricmain.txtDutchDesc.Text
    ElseIf strLangOfDesc = "French" Then
        txtDesc.Text = frm_taricmain.txtFrnchDesc.Text
    End If
    
    If strDetailLoaded = "New Country" Then Exit Sub
    
    '-----> Load values for Taric Detail country code and country for Import
    If strDetailType = "Import" Then
        strCountry = frm_taricmain.lvwImport.SelectedItem.Text
    Else
        strCountry = frm_taricmain.lvwExport.SelectedItem.Text
    End If
    
    intSpaceLoc = InStr(1, strCountry, " ")
    
    If intSpaceLoc Then
        txtCtryCode.Text = Left(strCountry, intSpaceLoc - 1)
        txtCtry.Text = Mid(strCountry, intSpaceLoc + 3)
    End If
End Sub

Private Sub SaveChanges(ByVal NewRecordAdded As Boolean)
    Dim strSQL As String
    
    '----> Save to Database
    With m_rstDetail
        '-----> Save General
        ![TARIC CODE] = txtCode.Text
        ![CTRY CODE] = txtCtryCode.Text
        
        '----> Save licence
        If strDetailType = "Import" Then
            If chkLicenceImp.Value = 1 Then
                ![LIC REQD] = -1
                ![MIN VALUE] = txtLimit.Text
                ![Min Value Curr] = "EUR"    ' cboCurrency.Text
            ElseIf chkLicenceImp.Value = 0 Then ![LIC REQD] = 0
                ![MIN VALUE] = ""
                ![Min Value Curr] = ""
            End If
        ElseIf strDetailType = "Export" Then
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
                    strSQL = strSQL & "[" & strDetailType & "] "
                    strSQL = strSQL & "WHERE "
                    strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
                    strSQL = strSQL & "ORDER BY "
                    strSQL = strSQL & strDetailType & ".[CTRY CODE] ASC "
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
                         
                         UpdateRecordset m_conTaric, m_rstDefault, strDetailType
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
    End With
    
    If NewRecordAdded Then
        InsertRecordset m_conTaric, m_rstDetail, strDetailType
    Else
        UpdateRecordset m_conTaric, m_rstDetail, strDetailType
    End If
            
    With m_rstDetail
        '----->Clear Previous usage default
        If chkDefault.Value = 1 Then
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While Not .EOF
                    If ![TARIC CODE] = txtCode.Text Then
                        If Not ![CTRY CODE] = txtCtryCode.Text Then
                            '.Edit
                            ![DEF CODE] = 0
                            ![COMM CODE] = 0
                            .Update
                        End If
                    End If
                    
                    .MoveNext
                Loop
                
                ExecuteNonQuery m_conTaric, "UPDATE [" & strDetailType & "] SET [DEF CODE] = 0, [COMM CODE] = 0 WHERE [TARIC CODE] = '" & txtCode.Text & "' AND [CTRY CODE] <> '" & txtCtryCode.Text & "' "
                
            End If
        End If
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ADORecordsetClose m_rstDetail
    ADORecordsetClose m_rstCountry
    
    ADODisconnectDB m_conTaric
    ADODisconnectDB m_conSADBEL
    
    UnloadControls Me
End Sub

Private Sub txtCtryCode_Change()
    Dim strSQL As String
    Dim Counter As Integer
    Dim colEmptyFieldValues As Collection
    
    blnUncheck = False

    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "* "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "[PICKLIST MAINTENANCE " & strLangOfDesc & "] "
    strSQL = strSQL & "WHER "
    strSQL = strSQL & "[INTERNAL CODE] = '8.29801619052887E+19' "
    
    If Len(Trim(txtCtryCode.Text)) = 3 Then
        '-----> if country code is less than 3 digits. Saving is not allowed
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        
        '----> Input Country Name that reflects the country code
        ADORecordsetOpen strSQL, m_conSADBEL, m_rstCountry, adOpenKeyset, adLockOptimistic
        'Set m_rstCountry = m_conSADBEL.OpenRecordset(strSQL)
                                                                
        With m_rstCountry
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While Not .EOF
                    If ![Internal Code] = "8.29801619052887E+19" And _
                        txtCtryCode.Text = ![code] Then
                        
                        If strLangOfDesc = "Dutch" Then
                            txtCtry.Text = m_rstCountry![DESCRIPTION DUTCH]
                        ElseIf strLangOfDesc = "French" Then
                            txtCtry.Text = ![DESCRIPTION FRENCH]
                        End If
                        
                        GoTo OutSide:
                    End If
                    
                    .MoveNext
                Loop
                '-----> If country code does not exist in database
                '-----> clear txtctrycode and exit
                '-----> disable saving
                Dim intBox As Integer
                
                intBox = MsgBox(Translate(866), vbOKOnly + vbExclamation, Me.Caption)
                
                cmdOK.Enabled = False
                cmdApply.Enabled = False
                
                txtCtryCode.SelStart = 0
                txtCtryCode.SelLength = Len(txtCtryCode.Text)
            End If
            
            Exit Sub
        End With
    
OutSide:
        '----> Save empty box field values to colEmptyFieldValues
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "[BOX CODE], "
            strSQL = strSQL & "[EMPTY FIELD VALUE] "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "[BOX DEFAULT " & strDetailType & " ADMIN] "
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
        '-----> save values to collection
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
        
        '-----> load values from database to settings detail
        With m_rstDetail
            On Error GoTo NoRecord:
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While Not .EOF
                    If Trim(txtCtryCode.Text) = ![CTRY CODE] And txtCode.Text = ![TARIC CODE] Then
                        '-----> if selected country already exist enter edit mode
                        '---->load if licence required
                        If strDetailType = "Import" Then
                            If ![LIC REQD] = -1 Then
                                chkLicenceImp.Value = 1
                                If Not IsNull(![MIN VALUE]) Then
                                    txtLimit.Text = ![MIN VALUE]
                                End If
                            Else
                                chkLicenceImp.Value = 0
                            End If
                            
                        ElseIf strDetailType = "Export" Then
                            If ![LIC REQD] = -1 Then
                                chkLicenceExp.Value = 1
                            Else
                                chkLicenceExp.Value = 0
                            End If
                        End If
                         '-----> load attached documents
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
                        '----> load special regime
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
                        If ![DEF CODE] = -1 Then
                            chkDefault.Value = 1
                        Else
                            chkDefault.Value = 0
                        End If
                        
                        Call DefaultCheck
                        
                        If ![COMM CODE] = -1 Then
                            chkCommon.Value = 1
                        Else
                            chkCommon.Value = 0
                        End If
                        
                        Exit Sub
                    End If
                    
                    .MoveNext
                Loop
            End If
        End With
NoRecord:
        '----> if entered country code dosn't exist in the DB then enter default values
        Call BlankOut
        Call DefaultCheck
    Else
        cmdOK.Enabled = False
        cmdApply.Enabled = False
    End If

End Sub

Private Sub BlankOut()
    Dim Counter As Integer
    Dim strSQL As String
    '-----> if selected country dosn't exist enter default values
    If Not blnCopy Then
    
        '----> Save empty box field values to colEmptyFieldValues
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "[BOX CODE], "
            strSQL = strSQL & "[EMPTY FIELD VALUE] "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "[BOX DEFAULT " & strDetailType & " ADMIN] "
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
        
        '-----> load usage
        chkDefault.Value = 0
        chkCommon.Value = 0
        
        Set colEmptyFieldValues = Nothing
    End If
End Sub

Private Sub SaveOptions()

    Dim intSave As Integer
    
    With m_rstDetail
        On Error GoTo FirstRecord:
            '----> Check if Record already exist
            .MoveFirst
            Do While Not .EOF
                If txtCode.Text = ![TARIC CODE] And txtCtryCode.Text = ![CTRY CODE] Then
                    '-----> If Country and Code ecombination already exist
                    If strDetailLoaded = "New Country" Or strDetailLoaded = "Copy" Then
                        '-----> msgbox disabled to avoid confusion
                        'intSave = MsgBox(Translate(464) & "?", vbYesNo + vbQuestion, Me.Caption)
                        'If intSave = 7 Then Exit Sub
                        '.Edit
                        SaveChanges False
                        
                        Exit Sub
                    Else
                        '.Edit
                        SaveChanges False
                        
                        Exit Sub
                    End If
                End If
                
                .MoveNext
            Loop
FirstRecord:

        '----> Add the new Record
        .AddNew
        
        SaveChanges True
    End With
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

Private Sub txtDate_LostFocus(Index As Integer)
    '-----> check if the date inputed is valid
    If txtDate(Index).Text = "" Or txtDate(Index).Text = "0" Then Exit Sub
    
    If CheckDate(txtDate(Index).Text) = False Then
        txtDate(Index).SetFocus
        txtDate(Index).SelStart = 0
        txtDate(Index).SelLength = Len(txtDate(Index).Text)
    End If

End Sub

Private Sub txtLimit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Implemented to check pasted values.
    Dim lngDotGap As Long

    If Button = 2 And Len(txtLimit.Text) > 0 Then
        With txtLimit
            'Check for multiple decimals.
            lngDotGap = (InStrRev(.Text, ".") - InStr(.Text, ".")) - 1
            If lngDotGap > 0 Then
                lngDotGap = (InStr(InStr(.Text, ".") + 1, .Text, ".") - InStr(.Text, ".")) - 1
                If lngDotGap > 2 Then lngDotGap = 2
            End If

            If IsNumericDot(.Text) = True Then
                If lngDotGap > 0 Then
                    .Text = Mid(.Text, 1, (InStr(.Text, ".") + lngDotGap))
                Else
                    .Text = Mid(.Text, 1, InStr(.Text, ".") + 2)
                End If
            Else
                .Text = ""
            End If
        End With
    End If

End Sub

Private Sub txtLimit_Change()
    If Left(Right(txtLimit.Text, 4), 1) = "." Then
        txtLimit.Text = Left(txtLimit.Text, Len(txtLimit.Text) - 1)
        txtLimit.SelStart = Len(txtLimit.Text)
    End If
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKey0 To vbKey9
            ' Do not suppress.
        Case 46     'Ascii code for period
            'Condition prevents multiple periods (as decimal separator)
            If InStr(txtLimit.Text, ".") Then KeyAscii = 0
        Case 8
        Case Else
            KeyAscii = 0
    End Select

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

Private Sub txtNumber_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtRegValue_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub txtValue_KeyPress(Index As Integer, KeyAscii As Integer)

    '-----> only allow backspace and numerical values
    If KeyAscii = 46 Then If InStr(txtValue(Index).Text, ".") = 0 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub DefaultCheck()

    '-----> dissable chkdefault if there is only one record or no record exist
    '-----> set this country record to default
    Dim strSQL As String
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[" & strDetailType & "] "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[TARIC CODE] = " & Chr(39) & ProcessQuotes(txtCode.Text) & Chr(39) & " "
    ADORecordsetOpen strSQL, m_conTaric, m_rstDefault, adOpenKeyset, adLockOptimistic
    'Set m_rstDefault = m_conTaric.OpenRecordset(strSQL)
    
    With m_rstDefault
        If Not .EOF Then .MoveLast
        
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
    
    ADORecordsetClose m_rstDefault

End Sub

Private Function IsNumericDot(SearchText As String) As Boolean
    'Function created for Minimum License Value
    'Returns true if search string only contains numbers and dot(s).
    Dim lngCtr As Long
    
    IsNumericDot = True
    
    For lngCtr = 1 To Len(SearchText) + 1
        If InStr("0123456789.", Mid(SearchText, lngCtr, 1)) = 0 Then
            IsNumericDot = False
            
            Exit Function
        End If
    Next
End Function
