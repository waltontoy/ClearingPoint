VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInboundDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inbound Document"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmInboundDoc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpDocDate 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   38288
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraCertificate 
      Caption         =   "Certificate"
      Height          =   1455
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtCertificateNum 
         Height          =   315
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCertificateType 
         Height          =   315
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCertificateNum 
         Caption         =   "Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblCertificateType 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame fraDocument 
      Caption         =   "Document"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtDocType 
         Height          =   315
         Left            =   720
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCustomsOfc 
         Caption         =   "..."
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtCustomsOffice 
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   3
         Top             =   960
         Width           =   1155
      End
      Begin VB.TextBox txtDocNumber 
         Height          =   315
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCustomsOffice 
         Caption         =   "Customs Office:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lblDocDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblDocNumber 
         Caption         =   "Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmInboundDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As ADODB.Connection
Dim m_rst As ADODB.Recordset

Dim m_lngLastSeqNum As Long
Dim m_bCancel As Boolean
Dim m_Button As PCubeLibPick.ButtonType
Dim m_clsPicklist As PCubeLibPick.CPicklist

Dim m_strLanguage As String

Public Sub MyLoad(ByVal conn As ADODB.Connection, _
                  ByVal rst As ADODB.Recordset, _
                ByVal Button As PCubeLibPick.ButtonType, _
                ByRef Cancel As Boolean, ByVal MyResourceHandler As Long, _
                ByRef lngLastSeqNum As Long, ByVal clsPicklist As PCubeLibPick.CPicklist, _
                ByVal Language As String)
                   
    Set m_clsPicklist = clsPicklist
    Set m_rst = rst
    Set m_conn = conn
    m_strLanguage = Language
    
    m_lngLastSeqNum = lngLastSeqNum
    
    m_bCancel = True
    m_Button = Button
    
    Select Case Button
        Case cpiAdd
            Me.txtDocType.Text = ""
            Me.txtDocNumber.Text = ""
            Me.dtpDocDate.Value = Now()
            Me.txtCertificateType.Text = ""
            Me.txtCertificateNum.Text = ""
            Me.txtCustomsOffice.Text = ""
        Case cpiCopy, cpiModify
            If Button = cpiModify Then
                txtDocType.Enabled = False
                txtDocNumber.Enabled = False
            End If
            Me.txtDocType.Text = IIf(IsNull(rst.Fields("Doc Type").Value), "", rst.Fields("Doc Type").Value)
            Me.txtDocNumber.Text = IIf(IsNull(rst.Fields("Document Number").Value), "", rst.Fields("Document Number").Value)
            Me.dtpDocDate.Value = IIf(IsNull(rst.Fields("DocDate").Value), Now(), rst.Fields("DocDate").Value)
            Me.txtCertificateType.Text = IIf(IsNull(rst.Fields("Cert_Type").Value), "", rst.Fields("Cert_Type").Value)
            Me.txtCertificateNum.Text = IIf(IsNull(rst.Fields("Cert_Num").Value), "", rst.Fields("Cert_Num").Value)
            Me.txtCustomsOffice.Text = IIf(IsNull(rst.Fields("DocOffice").Value), "", rst.Fields("DocOffice").Value)
    End Select
    
    Me.Show vbModal
    
    Set rst = m_rst
    
    lngLastSeqNum = m_lngLastSeqNum
    Cancel = m_bCancel
End Sub

Private Sub cmdCancel_Click()
    m_bCancel = True
    Unload Me
End Sub

Private Sub cmdCustomsOfc_Click()
    Dim pckCustoms As PCubeLibPick.CPicklist
    Dim gsdPicklist As PCubeLibPick.CGridSeed
    Dim strSQL As String

    Set pckCustoms = New PCubeLibPick.CPicklist
    Set gsdPicklist = pckCustoms.SeedGrid("Code", 1300, "Left", "Description", 2970, "Left")

    'strSql = "SELECT [PICKLIST MAINTENANCE DUTCH].CODE as CODE, [PICKLIST MAINTENANCE DUTCH].[DESCRIPTION DUTCH] AS Description FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE DUTCH] WHERE ([PICKLIST DEFINITION].[BOX CODE]= 'A4') AND ([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND ([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE DUTCH].[internal code]) ORDER BY CODE"
    strSQL = "SELECT [PICKLIST MAINTENANCE " & m_strLanguage & "].CODE as CODE, [PICKLIST MAINTENANCE " & m_strLanguage & "].[DESCRIPTION " & m_strLanguage & "] AS Description FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE " & m_strLanguage & "] WHERE ([PICKLIST DEFINITION].[BOX CODE]= 'A4') AND ([PICKLIST DEFINITION].[DOCUMENT]= 'Import') AND ([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE " & m_strLanguage & "].[internal code]) ORDER BY CODE"

    With pckCustoms
        
        .Search True, "Code", txtCustomsOffice.Text
        .Pick Me, cpiSimplePicklist, m_conn, strSQL, "Code", "Customs Office", vbModal, gsdPicklist, , , True, cpiKeyF2

        If .CancelTrans = False Then

            If Not .SelectedRecord Is Nothing Then
                txtCustomsOffice.Text = .SelectedRecord.RecordSource.Fields("Code").Value
            End If
        Else

        End If
    End With

    Set pckCustoms = Nothing
    Set gsdPicklist = Nothing

End Sub

Private Sub cmdOK_Click()
    Dim varBookMark As Variant
    Dim strOldFilter As String
    Dim blnOldFilterNone As Boolean

    If Len(Trim(txtDocNumber)) = 0 Then
        MsgBox "Please fill the Document Number box.", vbInformation, "Inbound Documents"
        txtDocNumber.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtDocType)) = 0 Then
        MsgBox "Please fill the Document Type box.", vbInformation, "Inbound Documents"
        txtDocType.SetFocus
        Exit Sub
    End If
    
    If Not (m_clsPicklist.GridRecord.BOF And m_clsPicklist.GridRecord.EOF) And (m_Button = cpiAdd Or m_Button = cpiCopy) Then
    
        varBookMark = m_clsPicklist.GridRecord.Bookmark
        
        m_clsPicklist.GridRecord.MoveFirst
        
        blnOldFilterNone = (m_clsPicklist.GridRecord.Filter = adFilterNone)
        strOldFilter = m_clsPicklist.GridRecord.Filter
        m_clsPicklist.GridRecord.Filter = "[Doc Type] = '" & txtDocType.Text & "' AND [Document Number] = '" & txtDocNumber.Text & "'"
        
        If Not m_clsPicklist.GridRecord.EOF Then
            MsgBox "Document Number already exists.", vbInformation, "Inbound Documents"
            txtDocNumber.SetFocus
            txtDocNumber.SelStart = 0
            txtDocNumber.SelLength = Len(txtDocNumber.Text)
            
            ' setting Filter to the string value of adFilterNone (= '0') causes RTE 3001
            If Not blnOldFilterNone Then
                m_clsPicklist.GridRecord.Filter = strOldFilter
            Else
                m_clsPicklist.GridRecord.Filter = adFilterNone
            End If
            
            m_clsPicklist.GridRecord.Bookmark = varBookMark
            Exit Sub
        Else
            ' setting Filter to the string value of adFilterNone (= '0') causes RTE 3001
            If Not blnOldFilterNone Then
                m_clsPicklist.GridRecord.Filter = strOldFilter
            Else
                m_clsPicklist.GridRecord.Filter = adFilterNone
            End If
            m_clsPicklist.GridRecord.Bookmark = varBookMark
        End If
    End If
    
    m_bCancel = False

    m_rst.Fields("Doc Type").Value = Me.txtDocType.Text
    m_rst.Fields("Document Number").Value = Me.txtDocNumber.Text
    m_rst.Fields("DocDate").Value = Me.dtpDocDate.Value
    m_rst.Fields("Cert_Type").Value = Me.txtCertificateType.Text
    m_rst.Fields("Cert_Num").Value = Me.txtCertificateNum.Text
    m_rst.Fields("DocOffice").Value = Me.txtCustomsOffice.Text
    m_rst.Fields("InDoc_Global").Value = True
    
    If m_Button = cpiAdd Or m_Button = cpiCopy Then
        m_lngLastSeqNum = m_lngLastSeqNum + 1
        m_rst.Fields("Sequence Number").Value = m_lngLastSeqNum
    
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_rst = Nothing
End Sub

'Private Function GetAndUpdateLastSeqNum(ByVal strEntrepotNum As String, ByVal conSADBEL As ADODB.Connection) As Long
'    Dim strSQL As String
'    Dim rstEntrepots As ADODB.Recordset
'
'    strSQL = "SELECT Entrepot_LastSeqNum FROM Entrepots INNER JOIN (Products INNER JOIN StockCards" & _
'             " ON Products.Prod_ID = StockCards.Prod_ID)" & _
'             " ON Entrepots.Entrepot_ID = Products.Entrepot_ID" & _
'             " WHERE Entrepot_Type & '-' & Entrepot_Num = '" & strEntrepotNum
'
'    ADORecordsetOpen strSQL, conSADBEL, rstEntrepots, adOpenKeyset, adLockOptimistic
'    With rstEntrepots
'
'        If Not (.BOF And .EOF) Then
'            .MoveFirst
'
'            If IsNull(!Entrepot_LastSeqNum) Then
'
'
'                GetAndUpdateLastSeqNum = 0
'            Else
'
'
'                GetAndUpdateLastSeqNum = !Entrepot_LastSeqNum
'            End If
'
'            .Update
'        Else
'            GetAndUpdateLastSeqNum = 0
'        End If
'
'    End With
'
'    ADORecordsetClose rstEntrepots
'End Function

Private Sub txtCertificateType_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtCustomsOffice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdCustomsOfc_Click
    End If
End Sub

Private Sub txtDocNumber_LostFocus()
    If Len(txtDocNumber.Text) < 7 Then
        'Add zeros to the left until its length becomes 7.
        txtDocNumber.Text = AddZeros(txtDocNumber.Text, 7 - Len(txtDocNumber.Text))
    End If
End Sub

Private Function AddZeros(strNumber As String, bytNumberofZeros As Byte) As String
    Dim bytCounter As Byte
    
    bytCounter = 1
    Do While bytCounter <= bytNumberofZeros
        bytCounter = bytCounter + 1
        strNumber = "0" & strNumber
    Loop
    
    AddZeros = strNumber
End Function

Private Sub txtDocType_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub
