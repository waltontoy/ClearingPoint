VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalog Name - Trans"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTrans 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3690
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2340
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1575
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   2778
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFields(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFields(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFields(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFields(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtFields 
         DataField       =   "Product_Name"
         Height          =   315
         Index           =   1
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1020
         Width           =   3555
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Product_Name"
         Height          =   315
         Index           =   0
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IDX_ID = 0
Private Const IDX_NAME = 2

Dim clsRecord As CRecord
Dim blnCancel As Boolean

Public Sub ShowForm(ByRef OwnerForm As Form, _
                                    ByRef Record As CRecord, _
                                    ByRef Button As Long, _
                                    ByRef Cancel As Boolean, _
                                    ByVal FormCaption As String, _
                                    ByVal IDCaption As String, _
                                    ByVal NameCaption As String)
                                        
    blnCancel = True
    Set clsRecord = Record

    Dim oText As TextBox
    
    txtFields(IDX_ID).DataField = Record.RecordSource.Fields(IDX_ID).Name
    txtFields(IDX_NAME - 1).DataField = Record.RecordSource.Fields(IDX_NAME).Name
    
    For Each oText In Me.txtFields
        
        Set oText.DataSource = clsRecord.RecordSource.DataSource
        
    Next   ' oText

    Caption = FormCaption
    lblFields(IDX_ID).Caption = IDCaption
    lblFields(IDX_NAME - 1).Caption = NameCaption

    Set Icon = OwnerForm.Icon
    Me.Show vbModal
    
    Cancel = blnCancel

End Sub


Private Sub cmdTrans_Click(Index As Integer)

    Select Case Index
        Case 0
            blnCancel = CheckDuplicate
            
            ' check if name already exist
            If (blnCancel = False) Then
                Unload Me
            ElseIf (blnCancel = True) Then
                MsgBox "Name already exist.", vbInformation, Caption
            End If
            
        Case 1
            blnCancel = True
            Unload Me
    End Select

End Sub

Private Function CheckDuplicate() As Boolean
'
    Dim varLastFilter As Variant
    Dim varLastPosition As Variant
    Dim strName As String
    
    CheckDuplicate = False
    If ((clsRecord.OldRecordSource Is Nothing) = False) Then
        
        varLastFilter = clsRecord.OldRecordSource.Filter
        varLastPosition = clsRecord.OldRecordSource.Bookmark
        
        clsRecord.OldRecordSource.Filter = ""
        
        If (clsRecord.OldRecordSource.RecordCount > 0) Then
            clsRecord.OldRecordSource.MoveFirst
            
            Do While (clsRecord.OldRecordSource.EOF = False)
                
                strName = IIf(IsNull(clsRecord.OldRecordSource.Fields(2).Value), "", clsRecord.OldRecordSource.Fields(2).Value)
                
                If Trim$(UCase$(strName)) = Trim$(UCase$(txtFields(1).Text)) Then
                    CheckDuplicate = True
                    Exit Do
                End If
                clsRecord.OldRecordSource.MoveNext
            Loop
        End If
        
    
    End If
    
'
End Function
