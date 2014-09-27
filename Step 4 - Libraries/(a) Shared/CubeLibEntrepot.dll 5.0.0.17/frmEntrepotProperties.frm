VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEntrepotProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrepot Properties"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmEntrepotProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Tag             =   "2213"
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Tag             =   "178"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Tag             =   "179"
      Top             =   3120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "150"
      TabPicture(0)   =   "frmEntrepotProperties.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSerialNumber"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Repackaging"
      TabPicture(1)   =   "frmEntrepotProperties.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraSerialNumber 
         Caption         =   "Stock Updates"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   6735
         Begin VB.CheckBox chkDisableOutboundUpdates 
            Caption         =   "Disable update of Entrepot stocks for Outbound"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   4095
         End
         Begin VB.CheckBox chkDisableInboundUpdates 
            Caption         =   "Disable update of Entrepot stocks for Inbound"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Settings"
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Tag             =   "948"
         Top             =   360
         Width           =   6735
         Begin VB.Frame Frame2 
            Caption         =   "Gross Weight"
            Height          =   975
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   6135
            Begin VB.TextBox txtPercent 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   1
               Left            =   3795
               MaxLength       =   3
               TabIndex        =   11
               Text            =   "0"
               Top             =   585
               Width           =   390
            End
            Begin VB.CheckBox chkGrossWeight 
               Caption         =   "Disable gross weight checking"
               Height          =   255
               Left            =   360
               TabIndex        =   10
               Top             =   360
               Width           =   3255
            End
            Begin VB.Label Label3 
               Caption         =   "Allow gross weight differences of no more than "
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   13
               Top             =   690
               Width           =   3375
            End
            Begin VB.Label Label2 
               Caption         =   "percent."
               Height          =   255
               Index           =   1
               Left            =   4260
               TabIndex        =   12
               Top             =   675
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Net Weight"
            Height          =   975
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   6135
            Begin VB.TextBox txtPercent 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   0
               Left            =   3795
               MaxLength       =   3
               TabIndex        =   6
               Text            =   "0"
               Top             =   480
               Width           =   390
            End
            Begin VB.CheckBox chkNetWeight 
               Caption         =   "Disable net weight checking"
               Height          =   255
               Left            =   360
               TabIndex        =   5
               Top             =   240
               Width           =   3255
            End
            Begin VB.Label Label3 
               Caption         =   "Allow net weight differences of no more than "
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   8
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label Label2 
               Caption         =   "percent."
               Height          =   255
               Index           =   0
               Left            =   4260
               TabIndex        =   7
               Top             =   570
               Width           =   615
            End
         End
      End
      Begin VB.Label lblPath 
         Caption         =   "lblPath"
         Height          =   285
         Left            =   3555
         TabIndex        =   17
         Top             =   -645
         Visible         =   0   'False
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmEntrepotProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_conConnection As ADODB.Connection
Private m_rstEntrepot As ADODB.Recordset

Private strUserKey As String
Private frmMain As Form
Private blnCancelled As Boolean

Private Sub chkGrossWeight_Click()
    If chkGrossWeight.Value = 1 Then
        Label3(1).Enabled = False
        txtPercent(1).Enabled = False
        Label2(1).Enabled = False
    Else
        Label3(1).Enabled = True
        txtPercent(1).Enabled = True
        Label2(1).Enabled = True
    End If
End Sub

Private Sub chkNetWeight_Click()
    If chkNetWeight.Value = 1 Then
        Label3(0).Enabled = False
        txtPercent(0).Enabled = False
        Label2(0).Enabled = False
    Else
        Label3(0).Enabled = True
        txtPercent(0).Enabled = True
        Label2(0).Enabled = True
    End If
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)
    Dim blnAddNew As Boolean
    
    If Index = 0 Then
        'Check if Value is less or equal to 100%
        If (Len(Trim(txtPercent(0).Text)) <> 0) Then
            If (txtPercent(0).Text > 100) Then
                
                MsgBox "Please enter values ranging from 0 to 100.", vbInformation, "ClearingPoint"
                        
                txtPercent(0).SelStart = 0
                txtPercent(0).SelLength = Len(txtPercent(0).Text)
                On Error Resume Next
                txtPercent(0).SetFocus
                On Error GoTo 0
                
                Exit Sub
            End If
        End If
        
        'Check if Value is less or equal to 100%
        If (Len(Trim(txtPercent(1).Text)) <> 0) Then
            If (txtPercent(1).Text > 100) Then
                
                MsgBox "Please enter values ranging from 0 to 100.", vbInformation, "ClearingPoint"
                        
                txtPercent(1).SelStart = 0
                txtPercent(1).SelLength = Len(txtPercent(1).Text)
                On Error Resume Next
                txtPercent(1).SetFocus
                On Error GoTo 0
                
                Exit Sub
            End If
        End If
            
        blnAddNew = (m_rstEntrepot.EOF And m_rstEntrepot.BOF)
        If blnAddNew Then
            m_rstEntrepot.AddNew
        End If
        
        'Resolved. Changed line from CInt(txtPercent(1).Text)  to CInt(IIf(Len(Trim(txtPercent(1).Text)) > 0, txtPercent(1).Text, 0)) and from CInt(txtPercent(0).Text) to CInt(IIf(Len(Trim(txtPercent(0).Text)) > 0, txtPercent(0).Text, 0))
        'update net % diff and gross % diff
        m_rstEntrepot!Prop_Net_Diff = CInt(IIf(Len(Trim(txtPercent(0).Text)) > 0, txtPercent(0).Text, 0))
        m_rstEntrepot!Prop_Gross_Diff = CInt(IIf(Len(Trim(txtPercent(1).Text)) > 0, txtPercent(1).Text, 0))
         
        m_rstEntrepot!Prop_DisableNetCheck = IIf(chkNetWeight.Value = 1, -1, 0)
        m_rstEntrepot!Prop_DisableGrossCheck = IIf(chkGrossWeight.Value = 1, -1, 0)
        
        m_rstEntrepot!Prop_DisableInboundStocksUpdate = IIf(chkDisableInboundUpdates.Value = 1, -1, 0)
        m_rstEntrepot!Prop_DisableOutboundStocksUpdate = IIf(chkDisableOutboundUpdates.Value = 1, -1, 0)
        
        m_rstEntrepot.Update
        
        If blnAddNew Then
            InsertRecordset m_conConnection, m_rstEntrepot, "EntrepotProperties"
        Else
            UpdateRecordset m_conConnection, m_rstEntrepot, "EntrepotProperties"
        End If
        
        blnCancelled = False
        
    End If
        
    Unload Me
End Sub

Private Sub Form_Load()
    'Mod by BCo 2006-04-17
    'Licensing implemented, removed old activation sequence
    Dim strSQL As String
        
    chkDisableOutboundUpdates.Enabled = True
    chkDisableInboundUpdates.Enabled = True
    
    strSQL = "SELECT Prop_Net_Diff, Prop_Gross_Diff, Prop_DisableNetCheck, Prop_DisableGrossCheck, Prop_DisableOutboundStocksUpdate, Prop_DisableInboundStocksUpdate  " & _
                "FROM EntrepotProperties"

    ADORecordsetOpen strSQL, m_conConnection, m_rstEntrepot, adOpenKeyset, adLockOptimistic
    'm_rstEntrepot.Open strSQL, m_conConnection, adOpenKeyset, adLockOptimistic

    If Not (m_rstEntrepot.EOF And m_rstEntrepot.BOF) Then
        m_rstEntrepot.MoveFirst
        
        txtPercent(0).Text = IIf(IsNull(m_rstEntrepot!Prop_Net_Diff), 0, m_rstEntrepot!Prop_Net_Diff)
        txtPercent(1).Text = IIf(IsNull(m_rstEntrepot!Prop_Gross_Diff), 0, m_rstEntrepot!Prop_Gross_Diff)
        
        'Transfered from old CP- joy 4/30/2006
        chkDisableInboundUpdates.Value = IIf(m_rstEntrepot!Prop_DisableInboundStocksUpdate = -1, 1, 0)
        chkDisableOutboundUpdates.Value = IIf(m_rstEntrepot!Prop_DisableOutboundStocksUpdate = -1, 1, 0)
    
        chkNetWeight.Value = IIf(m_rstEntrepot!Prop_DisableNetCheck = -1, 1, 0)
        chkGrossWeight.Value = IIf(m_rstEntrepot!Prop_DisableGrossCheck = -1, 1, 0)
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ADORecordsetClose m_rstEntrepot
End Sub

Public Function LoadProp(ByVal conSource As ADODB.Connection, ByVal lngResourceHandler As Long) As Boolean
    Set m_conConnection = conSource
    ResourceHandler = lngResourceHandler
    
    Call LoadResStrings(Me, True)
    
    blnCancelled = True

    Call Me.Show(vbModal)
    
    LoadProp = blnCancelled
    Set m_conConnection = Nothing
End Function

'<<< dandan 120307
'Added checking for entries that are pasted
Private Sub txtPercent_Change(Index As Integer)
    If IsNumeric(Trim(txtPercent(Index).Text)) Then
    Else
        MsgBox "Please enter values ranging from 0 to 100.", vbInformation, "ClearingPoint"
        
        txtPercent(Index).Text = 0
        txtPercent(Index).SelStart = 0
        txtPercent(Index).SelLength = Len(txtPercent(Index).Text)
        
    End If
End Sub

Private Sub txtPercent_GotFocus(Index As Integer)
    txtPercent(Index).SelStart = 0
    txtPercent(Index).SelLength = Len(txtPercent(Index).Text)
End Sub

Private Sub txtPercent_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57    ' Allow digits
        Case 8           ' Allow backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPercent_Validate(Index As Integer, Cancel As Boolean)
    If (Len(Trim(txtPercent(Index).Text)) <> 0) Then
        If txtPercent(Index).Text > 100 Then
            MsgBox "Please enter values ranging from 0 to 100.", vbInformation, "ClearingPoint"
                    
            txtPercent(Index).SelStart = 0
            txtPercent(Index).SelLength = Len(txtPercent(Index).Text)
            
            Cancel = True
        End If
    Else
        txtPercent(Index).Text = 0
    End If
End Sub
