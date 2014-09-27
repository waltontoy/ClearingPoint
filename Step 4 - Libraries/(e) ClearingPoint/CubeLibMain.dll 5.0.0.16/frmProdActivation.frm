VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmProdActivation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Activation"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   6120
      Width           =   1315
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   5055
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "&Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   6120
      Width           =   1315
   End
   Begin VB.Frame fraFeatures 
      Appearance      =   0  'Flat
      Caption         =   "Features"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin GridEX16.GridEX jgxFeatures 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8493
         HeaderStyle     =   2
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnCount     =   2
         CardCaption1    =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Activation Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frmProdActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_conActivation As ADODB.Connection

Public Sub ShowForm(ByRef OwnerForm As Object, ByRef ADOConnection As ADODB.Connection)
    
    Set m_conActivation = ADOConnection
    Set Me.Icon = OwnerForm.Icon
    
    Me.Show vbModal
    
End Sub


Private Sub cmdActivate_Click()
    
    Dim rstKeys As ADODB.Recordset
    Dim rstFeatures As ADODB.Recordset
    Dim lngLicenseValue As Long
    Dim lngLicenseSum As Long
    
    Dim blnInvalidKey As Boolean
    
    ADORecordsetOpen "Select * From LicenseKeys Where LK_Code = '" & Trim(txtCode.Text) & "'", m_conActivation, rstKeys, adOpenKeyset, adLockOptimistic
    'Call RstOpen("Select * From LicenseKeys Where LK_Code = '" & Trim(txtCode.Text) & "'", m_conActivation, rstKeys, adOpenKeyset, adLockOptimistic)
    
    If rstKeys.RecordCount > 0 Then
        MsgBox "The features that you wish to add has been successfully activated.", vbInformation, Me.Caption
        Exit Sub
    End If
    
    lngLicenseValue = GetLicenseValue(txtCode.Text)

    m_conActivation.BeginTrans
    
    If lngLicenseValue > 0 Then
    
        rstKeys.AddNew
        rstKeys!LK_Code = Trim(txtCode.Text)
        rstKeys!LK_Value = lngLicenseValue
        rstKeys.Update
            
        ADORecordsetOpen "Select * From Features Order By Feature_Name", m_conActivation, rstFeatures, adOpenKeyset, adLockOptimistic
        'Call RstOpen("Select * From Features Order By Feature_Name", m_conActivation, rstFeatures, adOpenKeyset, adLockOptimistic)
        
        Do While Not rstFeatures.EOF
            If (lngLicenseValue And rstFeatures!Feature_Code) = rstFeatures!Feature_Code Then
                If rstFeatures!Feature_Activated = True Then
                    blnInvalidKey = True
                    Exit Do
                Else
                    lngLicenseSum = lngLicenseSum + rstFeatures!Feature_Code
                    rstFeatures!Feature_Activated = True
                    rstFeatures.Update
                End If
            End If
            rstFeatures.MoveNext
        Loop
    Else
        blnInvalidKey = True
    End If
    
    If lngLicenseValue = lngLicenseSum And blnInvalidKey = False Then
        m_conActivation.CommitTrans
        Set jgxFeatures.ADORecordset = rstFeatures
        Call FormatGrid
        MsgBox "The features that you wish to add were successfully activated.", vbInformation, Me.Caption
    Else
        m_conActivation.RollbackTrans
        MsgBox "The activation code that you have entered is incorrect.", vbInformation, Me.Caption
        
    End If
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim rstFeatures As ADODB.Recordset
    
    ADORecordsetOpen "Select * From Features Order By Feature_Name", m_conActivation, rstFeatures, adOpenKeyset, adLockOptimistic
    'Call RstOpen("Select * From Features Order By Feature_Name", m_conActivation, rstFeatures, adOpenKeyset, adLockReadOnly)
    
    Set jgxFeatures.ADORecordset = rstFeatures
    
    Call FormatGrid
    
End Sub



Private Sub FormatGrid()

    jgxFeatures.Columns("Feature_ID").Visible = False
    jgxFeatures.Columns("Feature_Code").Visible = False
    
    jgxFeatures.Columns("Feature_Name").Caption = "Name"
    jgxFeatures.Columns("Feature_Name").HeaderAlignment = jgexAlignCenter
    jgxFeatures.Columns("Feature_Name").Width = 3000
    
    jgxFeatures.Columns("Feature_Activated").Caption = "Activated"
    jgxFeatures.Columns("Feature_Activated").HeaderAlignment = jgexAlignCenter

End Sub
