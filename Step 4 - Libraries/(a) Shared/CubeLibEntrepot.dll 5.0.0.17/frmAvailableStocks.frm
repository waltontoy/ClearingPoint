VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAvailableStocks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Stocks"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   Icon            =   "frmAvailableStocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTransact 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   10815
      TabIndex        =   7
      Tag             =   "179"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   9480
      TabIndex        =   6
      Top             =   7560
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "2181"
      TabPicture(0)   =   "frmAvailableStocks.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescription(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProdDescription"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "icbBatch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdProduct"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "jgxAvailableStocks"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "icbProduct"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin MSComctlLib.ImageCombo icbProduct 
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin GridEX16.GridEX jgxAvailableStocks 
         Height          =   5055
         Left            =   165
         TabIndex        =   5
         Top             =   2040
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8916
         TabKeyBehavior  =   1
         CursorLocation  =   3
         HideSelection   =   2
         Options         =   8
         RecordsetType   =   1
         ColumnCount     =   2
         CardCaption1    =   -1  'True
         DataMode        =   1
         ColumnHeaderHeight=   285
      End
      Begin VB.CommandButton cmdProduct 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   225
         Left            =   3780
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Product No."
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Tag             =   "2199"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Batch No."
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Tag             =   "2200"
         Top             =   1035
         Width           =   1575
      End
      Begin MSComctlLib.ImageCombo icbBatch 
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.Label lblProdDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2445
         TabIndex        =   11
         Top             =   1530
         Width           =   9255
      End
      Begin VB.Label Label1 
         Caption         =   "Product description:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "2201"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   6225
         TabIndex        =   9
         Top             =   675
         Visible         =   0   'False
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmAvailableStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conSADBEL As ADODB.Connection
Dim m_conTaric As ADODB.Connection

Dim rstGridOff As ADODB.Recordset
Dim rstGrid2Off As ADODB.Recordset

Private mvarAvailableStocks As cAvailableStocks
Private mvarOldValue As String
Private strBaseSql As String
Private strPreviousFiltering As String
Public Event cmdTransactClick(ByVal Index As Integer)
Private lngProd_ID As Long
Private blnUnload As Boolean
Private lngBookMark As Long
Private blnEnterPressed As Boolean
Private blnTaposNasaBeforeUpdate As Boolean
Private strPreviousProduct As String
Private strPreviousBatch As String


Private Const CB_SHOWDROPDOWN = &H14F

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long


Public Sub MyLoad(ByRef SADBELDB As ADODB.Connection, _
                    ByRef TARICDB As ADODB.Connection, _
                    ByVal strLanguage As String, _
                    ByRef cpiAvailableStocks As cAvailableStocks, _
                    ByVal cpiEDetails As cEntrepotDetails, ByVal strTabCaption As String, ByVal Resource As Long, _
                    ByVal blnIsDIA As Boolean, _
           Optional ByVal strMDBpath As String, _
           Optional ByVal strDocType As String, _
           Optional ByVal strDocNumber As String, _
           Optional ByVal strDocDate As String)
                
    Dim strCommand As String
    Dim lngCtr As Long
                
    Dim rstTemp As ADODB.Recordset

    ' Load Resource Strings
    ResourceHandler = Resource
    modGlobals.LoadResStrings Me, True
    
    Me.MousePointer = vbHourglass
    
    ' Pass Parameters
    Set m_conSADBEL = SADBELDB
    Set m_conTaric = TARICDB
    
    Set mvarAvailableStocks = cpiAvailableStocks
    
    ' Set Active Language to Use
    mvarAvailableStocks.Common.ActiveLanguage = strLanguage
    
    Set rstGridOff = New ADODB.Recordset
    rstGridOff.CursorLocation = adUseClient
    
    ''FillBatchCombo
    
    'txtProductNo.Enabled = False
    icbProduct.Enabled = False
    icbBatch.Enabled = False
    
    ' Glenn 12/7/2006
    ' Check if valid Entrepot, if not, abort displaying available stocks form.
        strCommand = ""
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num as Entrepot "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "Entrepots "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & ProcessQuotes(mvarAvailableStocks.Entrepot_Num) & "' "
    ADORecordsetOpen strCommand, m_conSADBEL, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open strCommand, m_conSADBEL, adOpenKeyset, adLockReadOnly
    
    If rstTemp.EOF And rstTemp.BOF Then
        MsgBox "Invalid Entrepot Number.", vbInformation, "Available Stocks"
        
    Else
        rstTemp.MoveFirst
        
        strCommand = SQLMain(blnIsDIA)
        
        ADORecordsetOpen strCommand, m_conSADBEL, rstGridOff, adOpenKeyset, adLockOptimistic
        'rstGridOff.Open strCommand, m_conSADBEL, adOpenKeyset, adLockPessimistic
        
        FillProductsCombo
        FillCombo
        
        LoadToGrid blnIsDIA, strMDBpath, strDocType, strDocNumber, strDocDate
        
        RecomputeAvailable cpiEDetails, strTabCaption
        
        If mvarAvailableStocks.In_ID <> 0 Then
        
            If Not (rstGrid2Off.BOF And rstGrid2Off.EOF) Then
                rstGrid2Off.MoveFirst
            End If
            
            For lngCtr = 1 To rstGrid2Off.RecordCount
                
                If rstGrid2Off!In_ID = mvarAvailableStocks.In_ID Then
                    'rstGrid2Off.Fields("Qty To Reserve").Value = mvarAvailableStocks.QtyToReserve
                    rstGrid2Off![Qty To Reserve] = mvarAvailableStocks.QtyToReserve
                    
                    'IAN 05-16-2005 => commented to be able to handle excess value in syslink
    '                If Val(rstGrid2Off!Available) >= Val(mvarAvailableStocks.QtyToReserve) Then
                        
                        rstGrid2Off!Available = Replace(Val(rstGrid2Off!Available) - Val(mvarAvailableStocks.QtyToReserve), ",", ".")
                        
    '                End If
                    
                    rstGrid2Off![Job No] = mvarAvailableStocks.JobNumber
                    rstGrid2Off.Update
    
                End If
                
                rstGrid2Off.MoveNext
                
            Next
            
            Set jgxAvailableStocks.ADORecordset = rstGrid2Off
            
            'Glenn - added this to have correct rowcount
            jgxAvailableStocks.MoveLast
            jgxAvailableStocks.MoveFirst
            
            If jgxAvailableStocks.RowCount > 0 Then
                rstGrid2Off.MoveFirst
                cmdTransact(0).Enabled = True
            Else
                cmdTransact(0).Enabled = False
            End If
            
            For lngCtr = 1 To jgxAvailableStocks.RowCount
                
                If jgxAvailableStocks.Value(jgxAvailableStocks.Columns("In_ID").Index) = mvarAvailableStocks.In_ID Then
                    lblProdDescription.Caption = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index))
                    Exit For
                End If
                
                jgxAvailableStocks.MoveNext
                
            Next
            
            If lngCtr > jgxAvailableStocks.Row Then
                jgxAvailableStocks.MoveFirst
                lblProdDescription.Caption = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index))
            End If
            
        Else
            Set jgxAvailableStocks.ADORecordset = rstGrid2Off
    
            If rstGrid2Off.RecordCount > 0 Then
                cmdTransact(0).Enabled = True
                'rstGrid2Off.MoveFirst
                jgxAvailableStocks.MoveFirst
                
                lblProdDescription.Caption = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index))
            Else
                cmdTransact(0).Enabled = False
            End If
        End If
        
       'CSCLP-261
       FilterRst1
       FormatColumns
        
        Me.Show vbModal
    End If
    
    Me.MousePointer = vbNormal
    
    Set rstTemp = Nothing
End Sub


Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0  'Product No.
            If Check1(Index).Value = 1 Then
                icbProduct.Enabled = True
                
                icbProduct_Change

            Else
                icbProduct.Enabled = False
                FilterRst
            End If
        Case 1  'Batch
            If Check1(Index).Value = 1 Then
                icbBatch.Enabled = True
                icbBatch.SetFocus
                icbBatch.Tag = icbBatch.Text
                FilterRst
            Else
                icbBatch.Enabled = False
                FilterRst
            End If
    End Select
End Sub


Private Sub cmdTransact_Click(Index As Integer)
    jgxAvailableStocks.Update
    
    
    If Index = 0 Then
        If IsRowSelected Then
            jgxAvailableStocks.Update
            
            If CheckQTy Then
                    
                    If IsNull(rstGrid2Off.Fields("Qty To Reserve").Value) Then
                        mvarAvailableStocks.QtyToReserve = "0"
                        jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index) = "0"
                    Else
                        If rstGrid2Off.Fields("Qty To Reserve").Value = "" Then
                            mvarAvailableStocks.QtyToReserve = "0"
                            jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index) = "0"
                        Else
                            mvarAvailableStocks.QtyToReserve = rstGrid2Off.Fields("Qty To Reserve").Value
                        End If
                    End If

                    If mvarAvailableStocks.Entrepot_Num = "" Or mvarAvailableStocks.Entrepot_Num = "0" Then
                        mvarAvailableStocks.Entrepot_Num = IIf(IsNull(rstGrid2Off.Fields("Entrepot").Value), "", rstGrid2Off.Fields("Entrepot").Value)
                    End If
                    
                    mvarAvailableStocks.TaricCode = IIf(IsNull(rstGrid2Off.Fields("Taric_Code").Value), "", rstGrid2Off.Fields("Taric_Code").Value)
                    mvarAvailableStocks.BatchNumber = IIf(IsNull(rstGrid2Off.Fields("Batch No").Value), "", rstGrid2Off.Fields("Batch No").Value)
                    mvarAvailableStocks.JobNumber = IIf(IsNull(rstGrid2Off.Fields("Job No").Value), "", rstGrid2Off.Fields("Job No").Value)
                    
                    mvarAvailableStocks.Stock_Num = IIf(IsNull(rstGrid2Off.Fields("Stock Card No").Value), "", rstGrid2Off.Fields("Stock Card No").Value)
                    mvarAvailableStocks.Stock_ID = IIf(IsNull(rstGrid2Off.Fields("Stock_ID").Value), "", rstGrid2Off.Fields("Stock_ID").Value)
                    mvarAvailableStocks.In_ID = IIf(IsNull(rstGrid2Off.Fields("In_ID").Value), "", rstGrid2Off.Fields("In_ID").Value)
        
                    mvarAvailableStocks.ProductNum = IIf(IsNull(rstGrid2Off.Fields("Prod_Num").Value), "", rstGrid2Off.Fields("Prod_Num").Value)
                    mvarAvailableStocks.SelectedRecord = True
                    
                    ComputeQtyorWgt
                    Unload Me
                End If
        Else
            MsgBox Translate(2180), vbOKOnly + vbInformation, Translate(2181)
        
        End If
    Else
        mvarAvailableStocks.SelectedRecord = False
        Unload Me
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)

    ADORecordsetClose rstGrid2Off
    ADORecordsetClose rstGridOff

    Set m_conSADBEL = Nothing
    Set m_conTaric = Nothing

    Set mvarAvailableStocks = Nothing
    
    Set frmAvailableStocks = Nothing
End Sub

Private Sub icbBatch_Change()
    icbBatch.Tag = icbBatch.Text
    FilterRst
End Sub

Private Sub icbBatch_Click()
    If Not icbBatch.SelectedItem Is Nothing Then
    
    
        icbBatch.Tag = icbBatch.SelectedItem.Text
        
    End If
    If icbBatch.Text <> strPreviousProduct Then
        strPreviousBatch = icbBatch.Text
        Debug.Print "Trigger filter icbProduct_Click"
        FilterRst
    End If

End Sub


Private Sub icbBatch_GotFocus()
    icbBatch.SelStart = 0
    icbBatch.SelLength = Len(icbBatch.Text)
End Sub

Private Sub icbBatch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 35 Then
        KeyAscii = 0
    End If
End Sub

Private Sub icbProduct_Change()
    If Len(Trim(icbProduct.Text)) > 0 Then
        icbProduct.Tag = Loop_Products_Contents(icbProduct.Text)
        FilterRst
    Else
        icbProduct.Tag = ""
        FilterRst
    End If
End Sub

Private Function Loop_Products_Contents(ByVal strInputString As String) As String
    Dim lngLoop_Ctr As Long
    Dim strFindString As String
    Dim blnFound As Boolean
    
    blnFound = False

    strFindString = Trim(strInputString)
    
    For lngLoop_Ctr = 1 To icbProduct.ComboItems.Count
        If Trim(icbProduct.ComboItems(lngLoop_Ctr).Text) = strFindString Then
            blnFound = True
            Exit For
        End If
    Next
    
    If blnFound = True Then
      
        Loop_Products_Contents = Left(icbProduct.ComboItems(lngLoop_Ctr).Key, 5)
    Else
        Loop_Products_Contents = "K0"
    End If
End Function

Private Sub icbProduct_Click()
    If Not icbProduct.SelectedItem Is Nothing Then
    
    
        icbProduct.Tag = icbProduct.SelectedItem.Key
        
    End If
    If icbProduct.Text <> strPreviousProduct Then
        strPreviousProduct = icbProduct.Text
        Debug.Print "Trigger filter icbProduct_Click"
        FilterRst
    End If
End Sub

Private Sub icbProduct_GotFocus()
    icbProduct.SelStart = 0
    icbProduct.SelLength = Len(icbProduct.Text)
End Sub

Private Sub icbProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 35 Then
        KeyAscii = 0
    End If
End Sub

Private Sub jgxAvailableStocks_AfterColUpdate(ByVal ColIndex As Integer)
    If jgxAvailableStocks.Columns(ColIndex).Key = "Qty To Reserve" Then
        jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Available").Index) = Replace(Val(IIf(Trim(mvarOldValue) = "", 0, mvarOldValue)) - IIf(Len(Trim(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index))) = 0, 0, Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index))) + Val(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Available").Index)), ",", ".")
        jgxAvailableStocks.Update
    End If
End Sub


Private Sub jgxAvailableStocks_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX16.JSRetBoolean)

    lngBookMark = Row
    With jgxAvailableStocks
        If ColIndex = .Columns("Qty To Reserve").Index Then
            If Val(IIf(Len(.Value(ColIndex)) = 0, "0", .Value(ColIndex))) > .Value(.Columns("Available").Index) + Val(IIf(Len(Trim(OldValue)) = 0, "0", OldValue)) Then

                Cancel = True
                blnUnload = False
                
                MsgBox Translate(2182), vbInformation, Translate(2181)
            Else
                mvarOldValue = OldValue
                blnUnload = True
                Cancel = False
            End If
        End If
    End With

End Sub

Private Sub jgxAvailableStocks_BeforeUpdate(ByVal Cancel As GridEX16.JSRetBoolean)
    If Len(Trim(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index))) = 0 Then
        jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index) = 0
    End If
    
End Sub

Private Sub jgxAvailableStocks_Click()
    If jgxAvailableStocks.Col = jgxAvailableStocks.Columns("Qty To Reserve").Index Then
        jgxAvailableStocks.EditMode = jgexEditModeOn
        jgxAvailableStocks.SelStart = 0
        jgxAvailableStocks.SelLength = Len(jgxAvailableStocks.Value(jgxAvailableStocks.Col))
    End If
    
    If Not jgxAvailableStocks.ADORecordset Is Nothing Then
        If jgxAvailableStocks.RowSelected(jgxAvailableStocks.Row) = True Then
            cmdTransact(0).Enabled = True
            lblProdDescription.Caption = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index))
        Else
            cmdTransact(0).Enabled = False
        End If
    End If
End Sub

Private Sub jgxAvailableStocks_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
    If jgxAvailableStocks.SortKeys.Count > 0 Then
        If jgxAvailableStocks.SortKeys.Item(1).ColIndex = Column.Index Then
            jgxAvailableStocks.SortKeys.Item(1).SortOrder = IIf(jgxAvailableStocks.SortKeys.Item(1).SortOrder = jgexSortAscending, jgexSortDescending, jgexSortAscending)
        Else
            jgxAvailableStocks.SortKeys.Clear
            jgxAvailableStocks.SortKeys.Add Column.Index, jgexSortAscending
        End If
    Else
        jgxAvailableStocks.SortKeys.Add Column.Index, jgexSortAscending
    End If
    
    jgxAvailableStocks.RefreshSort
End Sub

Private Sub jgxAvailableStocks_DblClick()
    cmdTransact_Click (0)
End Sub

Private Sub jgxAvailableStocks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        blnEnterPressed = True
    End If
End Sub

Private Sub jgxAvailableStocks_KeyPress(KeyAscii As Integer)
    If jgxAvailableStocks.Col = jgxAvailableStocks.Columns("Qty To Reserve").Index Then
        Select Case KeyAscii
              Case vbKey0 To vbKey9    ' Allow digits
              Case vbKeyBack           ' Allow backspace
              Case 46                  ' Allow period only once
              
                If jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Handling").Index) = 0 Then
                    KeyAscii = 0
                Else
                  If InStr(1, jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Qty To Reserve").Index), ".") Then
                      KeyAscii = 0
                  End If
                End If
              Case Else
                  KeyAscii = 0
          End Select
    End If
End Sub


Private Sub jgxAvailableStocks_LostFocus()
    jgxAvailableStocks.RowSelected(jgxAvailableStocks.Row) = True
End Sub

Private Sub jgxAvailableStocks_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If jgxAvailableStocks.Row <> lngBookMark Then
        If blnEnterPressed Then
            blnEnterPressed = False
            jgxAvailableStocks.RowSelected(LastRow) = True
            
        End If
    End If
End Sub

Private Function IsRowSelected() As Boolean
    Dim i As Long
    Dim lngQtyCtr As Long
    Dim lngRowPos As Long

    lngRowPos = jgxAvailableStocks.Row
    
    If Not jgxAvailableStocks.ADORecordset.EOF Then
        jgxAvailableStocks.ADORecordset.MoveFirst
    End If

    For i = 1 To jgxAvailableStocks.ADORecordset.RecordCount
    
        If jgxAvailableStocks.ADORecordset.Fields("Qty To Reserve").Value <> "" Then
            If IsNumeric(jgxAvailableStocks.ADORecordset.Fields("Qty To Reserve").Value) Then
                If Val(jgxAvailableStocks.ADORecordset.Fields("Qty To Reserve").Value) > 0 Then
                    lngQtyCtr = lngQtyCtr + 1
                End If
            End If
        End If
    
        If jgxAvailableStocks.RowSelected(i) = True Then
            IsRowSelected = True
        End If

        If Not jgxAvailableStocks.ADORecordset.EOF Then
            jgxAvailableStocks.ADORecordset.MoveNext
        End If
    Next
    
    jgxAvailableStocks.Row = lngRowPos
    
    If lngQtyCtr > 1 Then
        MsgBox Translate(2183), vbInformation, Translate(2181)
    End If
End Function


Public Sub ComputeQtyorWgt()
'Dim dblQtyWgt As Double
'HARD-CODED MUNA YUNG FORMATING!
    With jgxAvailableStocks
        Select Case .Value(.Columns("Prod_Handling").Index)
            Case 0  'quantity
        
                mvarAvailableStocks.PackageQuantity = CStr(Val(.Value(.Columns("Qty To Reserve").Index)))
                mvarAvailableStocks.GrossWeight = CStr((Val(.Value(.Columns("Qty To Reserve").Index)) / Val(.Value(.Columns("In_Orig_Packages_Qty").Index))) * Val(.Value(.Columns("In_Orig_Gross_Weight").Index)))
                mvarAvailableStocks.NetWeight = CStr((Val(.Value(.Columns("Qty To Reserve").Index)) / Val(.Value(.Columns("In_Orig_Packages_Qty").Index))) * Val(.Value(.Columns("In_Orig_Net_Weight").Index)))
            
            Case 1  'gross weight

                mvarAvailableStocks.PackageQuantity = CStr(Val(.Value(.Columns("Qty To Reserve").Index) / Val(.Value(.Columns("In_Orig_Gross_Weight").Index))) * Val(.Value(.Columns("In_Orig_Packages_Qty").Index)))
                mvarAvailableStocks.GrossWeight = CStr(Val(.Value(.Columns("Qty To Reserve").Index)))
                mvarAvailableStocks.NetWeight = CStr((Val(.Value(.Columns("Qty To Reserve").Index)) / Val(.Value(.Columns("In_Orig_Gross_Weight").Index))) * Val(.Value(.Columns("In_Orig_Net_Weight").Index)))

            Case 2  'net weight

                mvarAvailableStocks.PackageQuantity = CStr((Val(.Value(.Columns("Qty To Reserve").Index)) / Val(.Value(.Columns("In_Orig_Net_Weight").Index))) * Val(.Value(.Columns("In_Orig_Packages_Qty").Index)))
                mvarAvailableStocks.GrossWeight = CStr((Val(.Value(.Columns("Qty To Reserve").Index)) / Val(.Value(.Columns("In_Orig_Net_Weight").Index))) * Val(.Value(.Columns("In_Orig_Gross_Weight").Index)))
                mvarAvailableStocks.NetWeight = CStr(Val(.Value(.Columns("Qty To Reserve").Index)))

        End Select

        mvarAvailableStocks.PackageQuantity = Replace(Round(mvarAvailableStocks.PackageQuantity, 0), ",", ".")
        mvarAvailableStocks.GrossWeight = Replace(Round(mvarAvailableStocks.GrossWeight, 2), ",", ".")
        mvarAvailableStocks.NetWeight = Replace(Round(mvarAvailableStocks.NetWeight, 3), ",", ".")
        
    End With
End Sub

Private Sub FormatColumns()

    If mvarAvailableStocks.Entrepot_Num = "" Or mvarAvailableStocks.Entrepot_Num = "0" Then
        jgxAvailableStocks.Columns("Entrepot").Visible = True
    Else
        jgxAvailableStocks.Columns("Entrepot").Visible = False
    End If
    jgxAvailableStocks.Columns("In_ID").Visible = False
    jgxAvailableStocks.Columns("Prod_Desc").Visible = False
    jgxAvailableStocks.Columns("Prod_Num").Visible = False
    jgxAvailableStocks.Columns("Prod_ID").Visible = False
    jgxAvailableStocks.Columns("Taric_Code").Visible = False
    jgxAvailableStocks.Columns("Batch No").Visible = False
    jgxAvailableStocks.Columns("Stock_ID").Visible = False
    jgxAvailableStocks.Columns("Prod_Handling").Visible = False

    jgxAvailableStocks.Columns("In_Orig_Packages_Qty").Visible = False
    jgxAvailableStocks.Columns("In_Orig_Gross_Weight").Visible = False
    jgxAvailableStocks.Columns("In_Orig_Net_Weight").Visible = False
                    
    jgxAvailableStocks.Columns("Entrepot").Selectable = False
    jgxAvailableStocks.Columns("Stock Card No").Selectable = False
    jgxAvailableStocks.Columns("Total Out Qty").Selectable = False
    jgxAvailableStocks.Columns("Document No").Selectable = False
    jgxAvailableStocks.Columns("Quantity/Weight").Selectable = False
    jgxAvailableStocks.Columns("Reserved").Selectable = False
    jgxAvailableStocks.Columns("Available").Selectable = False
    jgxAvailableStocks.Columns("Prod_Handling").Selectable = False
    
    jgxAvailableStocks.Columns("Job No").MaxLength = 50

    jgxAvailableStocks.Columns("Qty To Reserve").TextAlignment = jgexAlignRight
    jgxAvailableStocks.Columns("Qty To Reserve").DefaultValue = 0

    Dim i As Integer
    For i = 1 To 10
        jgxAvailableStocks.Columns(i).Width = 1500
    Next
    
    jgxAvailableStocks.Columns(11).Width = 1605
    jgxAvailableStocks.Columns(12).Width = 1560
    jgxAvailableStocks.Columns(13).Width = 1590
    jgxAvailableStocks.Columns(14).Width = 1500
    jgxAvailableStocks.Columns(15).Width = 1500
    jgxAvailableStocks.Columns(16).Width = 1500
    jgxAvailableStocks.Columns(17).Width = 1305
    jgxAvailableStocks.Columns(18).Width = 1320
    jgxAvailableStocks.Columns(19).Width = 1245
    jgxAvailableStocks.Columns(20).Width = 1410

End Sub

Public Sub LoadToGrid(ByVal blnIsDIA As Boolean, _
                    Optional ByVal strMDBpath As String, _
                    Optional ByVal strDocType As String, _
                    Optional ByVal strDocNumber As String, Optional ByVal strDocDate As String)


    Dim rstDIA As ADODB.Recordset
    Dim strDIASQL As String
    
    Me.MousePointer = vbHourglass

    Set rstGrid2Off = New ADODB.Recordset
    rstGrid2Off.CursorLocation = adUseClient
    
    Dim fld As ADODB.Field
    
    
    'SQL to get DIA records
    strDIASQL = "SELECT Inbounds!In_Batch_Num AS Batch_Num, " & _
        "IIF(ISNULL(Inbounds!Stock_ID),0,Inbounds!Stock_ID) AS Stock_ID, " & _
        "IIF(ISNULL(Inbounds!In_Header),0,Inbounds!In_Header) AS Header, " & _
        "IIF(ISNULL(Inbounds!In_Detail),0,Inbounds!In_Detail) AS Detail " & _
        "FROM Inbounds INNER JOIN (StockCards INNER JOIN Products ON " & _
        "Products.Prod_ID = StockCards.Prod_ID) ON " & _
        "StockCards.Stock_ID = Inbounds.Stock_ID " & _
        "WHERE Inbounds!In_Job_Num='DIA' " & _
        "AND CHOOSE(Products!Prod_Handling + 1, Inbounds!In_Orig_Packages_Qty < 0, Inbounds!In_Orig_Gross_Weight < 0, Inbounds!In_Orig_Net_Weight < 0)"
    ADORecordsetOpen strDIASQL, m_conSADBEL, rstDIA, adOpenKeyset, adLockOptimistic
    'rstDIA.Open strDIASQL, m_conSADBEL, adOpenKeyset, adLockOptimistic
    
    For Each fld In rstGridOff.Fields
        If UCase(fld.Name) <> "DETAIL" And UCase(fld.Name) <> "HEADER" Then
            rstGrid2Off.Fields.Append fld.Name, fld.Type, fld.DefinedSize, adFldIsNullable
        End If
    Next
    
    rstGrid2Off.Fields.Append "Qty To Reserve", adVarChar, 12
    
    rstGrid2Off.Fields.Append "Job No", adVarChar, 50
    
    rstGrid2Off.Open , , adOpenKeyset, adLockOptimistic
    
    If rstGridOff.RecordCount > 0 Then
        rstGridOff.MoveFirst
    End If
    
    Do While Not rstGridOff.EOF
        
        rstDIA.Filter = 0
        rstDIA.Filter = "Batch_Num = '" & Right(rstGridOff![Document No], 7) & "' AND Stock_ID = " & IIf(IsNull(rstGridOff!Stock_ID), 0, rstGridOff!Stock_ID) & " AND Header = " & IIf(IsNull(rstGridOff!Header), 0, rstGridOff!Header) & " AND Detail = " & IIf(IsNull(rstGridOff!Detail), 0, rstGridOff!Detail)
                                
        'IAN 05-20-2005
        'If record count is greater than zero then current record on rstTemp has been
        'cancelled by DIA, hence, don't include in the grid.
        If (rstDIA.EOF And rstDIA.BOF) Then
        
            rstGrid2Off.AddNew
            For Each fld In rstGridOff.Fields
                
                If UCase(fld.Name) <> "DETAIL" And UCase(fld.Name) <> "HEADER" Then
                
                    If UCase(fld.Name) = UCase("Quantity/Weight") Or UCase(fld.Name) = UCase("Total Out Qty") Then
                        'quick solution to display the quantity in period format (".")
                        rstGrid2Off.Fields(fld.Name).Value = Replace(fld.Value, ",", ".")
                    Else
                        rstGrid2Off.Fields(fld.Name).Value = fld.Value
                    
                    End If
                    
                End If
                
            Next
            
            rstGrid2Off.Update
        End If
        
        rstGridOff.MoveNext
    Loop
    
    ADORecordsetClose rstDIA
    
    If blnIsDIA = True Then
        AddEditForDIA strMDBpath, strDocType, strDocNumber, strDocDate
    End If
    
    If rstGrid2Off.RecordCount > 0 Then
        rstGrid2Off.MoveFirst
    End If
    
    Me.MousePointer = vbNormal
End Sub

Private Function SQLMain(ByVal blnIsDIA As Boolean) As String
    Dim strCommand As String
    
    strCommand = ""
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num AS Entrepot, "
    strCommand = strCommand & "Inbounds.In_ID, "
    strCommand = strCommand & "Products.Prod_ID, "
    strCommand = strCommand & "Products.Prod_Num, "
    strCommand = strCommand & "Products.Prod_Desc, "
    strCommand = strCommand & "Products.Prod_Handling, "
    strCommand = strCommand & "Products.Taric_Code, "
    strCommand = strCommand & "Inbounds!In_Header AS Header, "
    strCommand = strCommand & "Inbounds!In_Detail AS Detail, "
    strCommand = strCommand & "In_Batch_Num as [Batch No], "
    strCommand = strCommand & "StockCards.Stock_ID AS [Stock_ID], "
    strCommand = strCommand & "StockCards.Stock_Card_Num AS [Stock Card No], "
    strCommand = strCommand & "InboundDocs.InDoc_Type + '-' + InboundDocs.InDoc_Num AS [Document No], "
    strCommand = strCommand & "CSTR(CHOOSE(Prod_Handling + 1, In_Orig_Packages_Qty, In_Orig_Gross_Weight, In_Orig_Net_Weight)) AS [Quantity/Weight], "
    strCommand = strCommand & "CSTR(In_Orig_Packages_Qty) AS In_Orig_Packages_Qty, "
    strCommand = strCommand & "CSTR(In_Orig_Gross_Weight) as In_Orig_Gross_Weight, "
    strCommand = strCommand & "CSTR(In_Orig_Net_Weight) as In_Orig_Net_Weight , "
    strCommand = strCommand & "CSTR(Inbounds.In_TotalOut_Qty_Wgt) AS [Total Out Qty], "
    strCommand = strCommand & "CSTR(Inbounds.In_Reserved_Qty_Wgt) AS Reserved, "
    strCommand = strCommand & "CSTR(Inbounds.In_Avl_Qty_Wgt) AS [Available] "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "( "
    strCommand = strCommand & "Inbounds "
    strCommand = strCommand & "INNER JOIN "
    strCommand = strCommand & "InboundDocs "
    strCommand = strCommand & "ON "
    strCommand = strCommand & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
    strCommand = strCommand & ") "
    strCommand = strCommand & "INNER JOIN "
    strCommand = strCommand & "( "
    strCommand = strCommand & "StockCards "
    strCommand = strCommand & "INNER JOIN "
        strCommand = strCommand & "( "
        strCommand = strCommand & "Products "
        strCommand = strCommand & "INNER JOIN "
        strCommand = strCommand & "Entrepots "
        strCommand = strCommand & "ON "
        strCommand = strCommand & "Products.Entrepot_ID = Entrepots.Entrepot_ID "
        strCommand = strCommand & ") "
    strCommand = strCommand & "ON "
    strCommand = strCommand & "StockCards.Prod_ID = Products.Prod_ID "
    strCommand = strCommand & ") "
    strCommand = strCommand & "ON "
    strCommand = strCommand & "Inbounds.Stock_ID = StockCards.Stock_ID "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "( "
    strCommand = strCommand & "Inbounds.In_Avl_Qty_Wgt > 0 "
    strCommand = strCommand & "OR "
    strCommand = strCommand & "Inbounds.In_Reserved_Qty_Wgt > 0 "
    strCommand = strCommand & ") "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "IIF(ISNULL(Inbounds!In_Code),'',Inbounds!In_Code NOT LIKE '%<<TEST>>') "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "IIF(ISNULL(Inbounds!In_Code),'',Inbounds!In_Code NOT LIKE '%<<CLOSURE>>') "
    
    If Not ((mvarAvailableStocks.Entrepot_Num = "" Or _
        mvarAvailableStocks.Entrepot_Num = "0") Or _
        blnIsDIA) Then
        
        strCommand = strCommand & "AND "
        strCommand = strCommand & "IIF(ISNULL(Inbounds!In_Code),'',Inbounds!In_Code NOT LIKE '%<<TEST>>') "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "IIF(ISNULL(Inbounds!In_Code),'',Inbounds!In_Code NOT LIKE '%<<CLOSURE>>') "
    End If
        
    ' filtering for country of export will be removed!
    Select Case mvarAvailableStocks.CodiType
        Case eCodiType.eCodi_TransitNCTS, _
            eCodiType.eCodi_EDIDeparture
            
        Case eCodiType.eCodi_PLDACombined, _
            eCodiType.eCodi_PLDAExport, _
            eCodiType.eCodi_PLDAImport
            
            If LenB(Trim$(mvarAvailableStocks.CtryOfOrigin)) > 0 Then
                strCommand = strCommand & "AND "
                strCommand = strCommand & "Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "' "
            End If
            
        Case Else
                
            ' This may be what Paul refers to: filtering for country of export will be removed!
            If (Len(mvarAvailableStocks.CtryOfOrigin) > 0 And mvarAvailableStocks.CtryOfOrigin <> "0") And (Len(mvarAvailableStocks.CtryOfExport) > 0 And mvarAvailableStocks.CtryOfExport <> "0") Then
                'strAdditional = " and Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "' and Prod_Ctry_Export = '" & mvarAvailableStocks.CtryOfExport & "'"
                strCommand = strCommand & "AND "
                strCommand = strCommand & "Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "' "
                
            ElseIf Len(mvarAvailableStocks.CtryOfOrigin) > 0 And mvarAvailableStocks.CtryOfOrigin <> "0" Then
                strCommand = strCommand & "AND "
                strCommand = strCommand & "Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "' "

            End If
    End Select

    SQLMain = strCommand
End Function

Public Function CheckQTy() As Boolean
    If Val(rstGrid2Off.Fields("Available").Value) < 0 Then
        
        CheckQTy = False
        MsgBox Translate(2182), vbInformation, Translate(2181)

    Else
        CheckQTy = True
    End If
End Function



Private Sub FillCombo()
    Dim strSQL As String
    Dim rstBatches As ADODB.Recordset
    
    
    If mvarAvailableStocks.Entrepot_Num = "" Or mvarAvailableStocks.Entrepot_Num = "0" Then
        strSQL = "SELECT DISTINCT In_Batch_Num as [Batch No] " & _
                    " FROM (Inbounds INNER JOIN InboundDocs ON Inbounds.InDoc_ID = InboundDocs.InDoc_ID) INNER JOIN (StockCards INNER JOIN (Products INNER JOIN Entrepots ON Products.Entrepot_ID = Entrepots.Entrepot_ID) ON StockCards.Prod_ID = Products.Prod_ID) ON Inbounds.Stock_ID = StockCards.Stock_ID " & _
                    " WHERE (Inbounds.In_Avl_Qty_Wgt > 0 or Inbounds.In_Reserved_Qty_Wgt > 0) "
    Else
        strSQL = "SELECT DISTINCT In_Batch_Num as [Batch No] " & _
                    " FROM (Inbounds INNER JOIN InboundDocs ON Inbounds.InDoc_ID = InboundDocs.InDoc_ID) INNER JOIN (StockCards INNER JOIN (Products INNER JOIN Entrepots ON Products.Entrepot_ID = Entrepots.Entrepot_ID) ON StockCards.Prod_ID = Products.Prod_ID) ON Inbounds.Stock_ID = StockCards.Stock_ID " & _
                    " WHERE (Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = '" & mvarAvailableStocks.Entrepot_Num & "' and (Inbounds.In_Avl_Qty_Wgt > 0 or Inbounds.In_Reserved_Qty_Wgt > 0) "
    
    End If
        
    ADORecordsetOpen strSQL, m_conSADBEL, rstBatches, adOpenKeyset, adLockPessimistic
    'rstBatches.Open strSQL, m_conSADBEL, adOpenKeyset, adLockPessimistic
    If Not (rstBatches.BOF And rstBatches.EOF) Then
        rstBatches.MoveFirst
    End If
    
    Do While Not rstBatches.EOF
        icbBatch.ComboItems.Add , , IIf(IsNull(rstBatches![Batch No]), "", rstBatches![Batch No])
        
        rstBatches.MoveNext
    Loop

    ADORecordsetClose rstBatches

End Sub



Private Sub FillProductsCombo()
    Dim strBase As String
    Dim strSQL As String
    Dim rstTmp As ADODB.Recordset
    Dim strAdditional As String
    
    If mvarAvailableStocks.Entrepot_Num = "" Or mvarAvailableStocks.Entrepot_Num = "0" Then
        strBase = "Select DISTINCT Prod.Prod_ID, Prod.Prod_Num, Prod.Prod_Desc  from (Products PROD INNER JOIN (StockCards SC INNER JOIN Inbounds on SC.Stock_ID = Inbounds.STock_ID)  on PROD.Prod_ID = SC.prod_ID)  INNER JOIN Entrepots ON Prod.Entrepot_ID = Entrepots.Entrepot_ID " & _
                " WHERE Inbounds.In_Avl_Qty_Wgt > 0 "
    
    
    Else
        strBase = "Select DISTINCT Prod.Prod_ID, Prod.Prod_Num, Prod.Prod_Desc  from (Products PROD INNER JOIN (StockCards SC INNER JOIN Inbounds on SC.Stock_ID = Inbounds.STock_ID)  on PROD.Prod_ID = SC.prod_ID)  INNER JOIN Entrepots ON Prod.Entrepot_ID = Entrepots.Entrepot_ID " & _
                " WHERE (Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = '" & mvarAvailableStocks.Entrepot_Num & "' and Inbounds.In_Avl_Qty_Wgt > 0 "
    End If
        
    If mvarAvailableStocks.CodiType = eCodi_TransitNCTS Or mvarAvailableStocks.CodiType = eCodi_EDIDeparture Then
    
    Else
    
        If (Len(mvarAvailableStocks.CtryOfOrigin) > 0 And mvarAvailableStocks.CtryOfOrigin <> "0") And (Len(mvarAvailableStocks.CtryOfExport) > 0 And mvarAvailableStocks.CtryOfExport <> "0") Then
            strAdditional = " and Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "'"
                
        ElseIf Len(mvarAvailableStocks.CtryOfOrigin) > 0 And mvarAvailableStocks.CtryOfOrigin <> "0" Then
            strAdditional = " and Prod_Ctry_Origin = '" & mvarAvailableStocks.CtryOfOrigin & "'"
        End If
    End If
    
    strSQL = strBase & strAdditional
    
    ADORecordsetOpen strSQL & " order by Prod_Num", m_conSADBEL, rstTmp, adOpenKeyset, adLockPessimistic
    'rstTmp.Open strSQL & " order by Prod_Num", m_conSADBEL, adOpenKeyset, adLockPessimistic
    
    If Not (rstTmp.EOF And rstTmp.BOF) Then
        rstTmp.MoveFirst
        
        Do While Not rstTmp.EOF
            icbProduct.ComboItems.Add , "K" & rstTmp!Prod_ID, IIf(IsNull(rstTmp!Prod_Num), "", rstTmp!Prod_Num) & " - " & IIf(IsNull(rstTmp!Prod_Desc), "", rstTmp!Prod_Desc)
            
            rstTmp.MoveNext
        Loop
    End If
    
    ADORecordsetClose rstTmp

End Sub

Public Function ShowComboBoxContents(Combo As Object) As Boolean

    Dim bAns As Boolean
    Dim lRet As Long
    
    If TypeOf Combo Is ImageCombo Then
        lRet = SendMessage(Combo.hwnd, CB_SHOWDROPDOWN, 1, 0)
        bAns = (lRet > 0)
        
    End If
    
    ShowComboBoxContents = bAns

End Function

'CSCLP-261
Public Sub FilterRst1()
    
    Dim strFiltering As String
    Dim lngIndex As Long
    Dim lngSort As Long
    Dim lngCtr As Long
    
        
    If Len(Trim(mvarAvailableStocks.TaricCode)) <> 0 Then
        strFiltering = "Taric_Code=" & "'" & mvarAvailableStocks.TaricCode & "'"
        Check1(0).Value = 1
        If icbProduct.ComboItems.Count > 0 Then
            
            icbProduct.ComboItems(1).Selected = True
        End If
        strPreviousFiltering = strFiltering
        rstGrid2Off.Filter = adFilterNone
        rstGrid2Off.Filter = strFiltering
        Set jgxAvailableStocks.ADORecordset = rstGrid2Off
        
        Exit Sub
    End If
    
    If mvarAvailableStocks.Entrepot_Num <> "-" Then
        strFiltering = "Entrepot=" & "'" & mvarAvailableStocks.Entrepot_Num & "'"
        
        strPreviousFiltering = strFiltering
        rstGrid2Off.Filter = adFilterNone
        rstGrid2Off.Filter = strFiltering
        
        Set jgxAvailableStocks.ADORecordset = rstGrid2Off
        jgxAvailableStocks.Refresh
    End If
End Sub

Public Sub FilterRst()
    
    Dim strFiltering As String
    Dim lngIndex As Long
    Dim lngSort As Long
    

    If Check1(0).Value = 1 And Check1(1).Value = 1 Then

        If icbProduct.Tag <> "" And icbBatch.Tag <> "" Then
            strFiltering = "Prod_ID = " & Mid(icbProduct.Tag, 2) & " AND [Batch No] = '" & icbBatch.Tag & "'"
        ElseIf icbProduct.Tag <> "" Then
            strFiltering = "Prod_ID = " & Mid(icbProduct.Tag, 2)
        ElseIf icbBatch.Tag <> "" Then
            strFiltering = "[Batch No] = '" & icbBatch.Tag & "'"
        Else
            strFiltering = ""
        End If
    ElseIf Check1(0).Value = 1 Then

        If icbProduct.Tag <> "" Then
            strFiltering = "Prod_ID = " & Mid(icbProduct.Tag, 2)
        Else
            strFiltering = ""
        End If
    ElseIf Check1(1).Value = 1 Then
        If icbBatch.Tag <> "" Then
            strFiltering = "[Batch No] = '" & icbBatch.Tag & "'"
        Else
            strFiltering = ""
        End If
    End If

    If strFiltering <> strPreviousFiltering Then
    
        If jgxAvailableStocks.SortKeys.Count > 0 Then
            lngIndex = jgxAvailableStocks.SortKeys(1).ColIndex
            lngSort = jgxAvailableStocks.SortKeys(1).SortOrder
        End If
        
        strPreviousFiltering = strFiltering
        rstGrid2Off.Filter = adFilterNone
        rstGrid2Off.Filter = strFiltering
        
        Set jgxAvailableStocks.ADORecordset = rstGrid2Off
        FormatColumns
        
        If lngIndex > 0 Then
            jgxAvailableStocks.SortKeys.Add lngIndex, lngSort
            jgxAvailableStocks.RefreshSort
            lngIndex = 0
        End If
        
    End If
    'CSCLP-261
    lblProdDescription.Caption = IIf(IsNull(jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index)), "", jgxAvailableStocks.Value(jgxAvailableStocks.Columns("Prod_Desc").Index))
    
End Sub


Private Sub RecomputeAvailable(ByVal cpiEDetails As cEntrepotDetails, ByVal strTabCaption As String)
                    
    Dim varTotal As Variant
    Dim lngCtr As Long
    Dim i As Long
    
    If Not (rstGrid2Off.BOF And rstGrid2Off.EOF) Then
        rstGrid2Off.MoveFirst
    End If
    
    Do While Not rstGrid2Off.EOF
        varTotal = 0
        
        For i = 1 To cpiEDetails.Count
        
            If cpiEDetails(i).In_ID = rstGrid2Off.Fields("In_ID").Value And cpiEDetails(i).Key <> strTabCaption Then
    
                varTotal = varTotal + Val(cpiEDetails(i).QtyToReserve)
        
            End If
        Next
        
        rstGrid2Off.Fields("Reserved").Value = Replace(CStr(Val(Replace(rstGrid2Off.Fields("Reserved").Value, ",", ".")) + Val(varTotal)), ",", ".")
        
        rstGrid2Off.Fields("In_Orig_Packages_Qty").Value = Replace(rstGrid2Off.Fields("In_Orig_Packages_Qty").Value, ",", ".")
        rstGrid2Off.Fields("In_Orig_Gross_Weight").Value = Replace(rstGrid2Off.Fields("In_Orig_Gross_Weight").Value, ",", ".")
        rstGrid2Off.Fields("In_Orig_Net_Weight").Value = Replace(rstGrid2Off.Fields("In_Orig_Net_Weight").Value, ",", ".")
        rstGrid2Off.Fields("Available").Value = Replace(CStr(Val(Replace(rstGrid2Off.Fields("Available").Value, ",", ".")) - Val(varTotal)), ",", ".")

        rstGrid2Off.Update
        
        rstGrid2Off.MoveNext
    Loop

End Sub

Private Sub AddEditForDIA(Optional ByVal strMDBpath As String, _
                    Optional ByVal strDocType As String, _
                    Optional ByVal strDocNumber As String, Optional ByVal strDocDate As String)

    Dim rstDIA As ADODB.Recordset
    Dim conHistory As ADODB.Connection
    Dim strSQL As String
    Dim rstTmp As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim lngProdHandling As Long

    strDocNumber = PadL(strDocNumber, 7, "0")
    
    If Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(Year(ConvertDDMMYY(strDocDate)), 2) & ".mdb") = "" Then
        Exit Sub
    End If

    '<<< dandan 112306
    '<<< Update database with password
    ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(Year(ConvertDDMMYY(strDocDate)), 2)
    'OpenADODatabase conHistory, strMDBpath, "Mdb_History" & Right(Year(ConvertDDMMYY(strDocDate)), 2) & ".mdb"
                
    strSQL = "Select *, InboundDocs.InDoc_Type+'-'+InboundDocs.InDoc_Num AS [Document No], Inbounds.Stock_ID as Stock_ID, In_Batch_Num as [Batch No] " & _
            " from InboundDocs INNER JOIN  (INBOUNDS INNER JOIN (Outbounds INNER JOIN OutboundDocs ON Outbounds.OutDoc_ID = OutboundDocs.OutDoc_ID) on  Inbounds.In_ID = Outbounds.In_ID)  ON InboundDocs.InDoc_ID = Inbounds.InDoc_ID " & _
             " WHERE OutDoc_Type = '" & strDocType & "' and OutDoc_Num = '" & strDocNumber & "' and CDate(Format(OutDoc_Date, 'mm/dd/yyyy')) = #" & ConvertDDMMYY(strDocDate) & "#"
    
    ADORecordsetOpen strSQL, conHistory, rstDIA, adOpenKeyset, adLockOptimistic
    'rstDIA.Open strSQL, conHistory, adOpenKeyset, adLockReadOnly
    
    If Not (rstDIA.EOF And rstDIA.BOF) Then
        rstDIA.MoveFirst
        Do While Not rstDIA.EOF
            If Not (rstGrid2Off.BOF And rstGrid2Off.EOF) Then
                rstGrid2Off.MoveFirst
            End If
            
            rstGrid2Off.Find "In_ID = " & rstDIA![Inbounds.In_ID]
            
            If rstGrid2Off.EOF Then    'if not found
    
                strSQL = "SELECT Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num as Entrepot, Inbounds.In_ID,  Products.Prod_ID, Products.Prod_Num, Products.Prod_Desc, Products.Prod_Handling, Products.Taric_Code, In_Batch_Num as [Batch No], StockCards.Stock_ID AS [Stock_ID], StockCards.Stock_Card_Num AS [Stock Card No], InboundDocs.InDoc_Type+'-'+InboundDocs.InDoc_Num AS [Document No]  FROM (Inbounds INNER JOIN InboundDocs ON Inbounds.InDoc_ID = InboundDocs.InDoc_ID) INNER JOIN (StockCards INNER JOIN (Products INNER JOIN Entrepots ON Products.Entrepot_ID = Entrepots.Entrepot_ID) ON StockCards.Prod_ID = Products.Prod_ID) ON Inbounds.Stock_ID = StockCards.Stock_ID  " & _
                        " where Inbounds.Stock_ID = " & rstDIA![Stock_ID]
    
                ADORecordsetOpen strSQL, m_conSADBEL, rstTmp, adOpenKeyset, adLockOptimistic
                'rstTmp.Open strSQL, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
                
                If rstTmp.BOF And rstTmp.EOF Then   'if no longer in mdb_sadbel
                    ADORecordsetClose rstTmp
                
                    strSQL = "SELECT Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num as Entrepot,  Products.Prod_ID, Products.Prod_Num, Products.Prod_Desc, Products.Prod_Handling, Products.Taric_Code, StockCards.Stock_ID AS [Stock_ID], StockCards.Stock_Card_Num AS [Stock Card No] FROM Entrepots INNER JOIN (Products INNER JOIN StockCards on Products.Prod_ID = StockCards.Prod_ID)  ON Entrepots.Entrepot_ID = Products.Entrepot_ID  " & _
                            " where StockCards.Stock_ID = " & rstDIA![Stock_ID]
    
                    ADORecordsetOpen strSQL, m_conSADBEL, rstTmp, adOpenKeyset, adLockOptimistic
                    'rstTmp.Open strSQL, m_conSADBEL, adOpenForwardOnly, adLockReadOnly
                    
                    If Not (rstTmp.EOF And rstTmp.BOF) Then
                        rstTmp.MoveFirst
                        
                        rstGrid2Off.AddNew
                        For Each fld In rstGrid2Off.Fields
                        
                            Select Case UCase(fld.Name)
                                Case "ENTREPOT"
                                    rstGrid2Off.Fields("Entrepot").Value = rstTmp.Fields("Entrepot").Value
                                    
                                Case "IN_ID"
                                    rstGrid2Off.Fields("In_ID").Value = rstDIA.Fields("Inbounds.In_ID").Value
                                    
                                Case "PROD_DESC"
                                    rstGrid2Off.Fields("Prod_Desc").Value = rstTmp.Fields("Prod_Desc").Value
                                
                                Case "PROD_NUM"
                                    rstGrid2Off.Fields("Prod_Num").Value = rstTmp.Fields("Prod_Num").Value
                                    
                                Case "PROD_ID"
                                    rstGrid2Off.Fields("Prod_ID").Value = rstTmp.Fields("Prod_ID").Value
                                    
                                Case "TARIC_CODE"
                                    rstGrid2Off.Fields("Taric_Code").Value = rstTmp.Fields("Taric_Code").Value
                                    
                                Case "BATCH NO"
                                    'rstGrid2Off.Fields("Batch No").Value = rstTmp.Fields("Batch No").Value
                                    rstGrid2Off.Fields("Batch No").Value = rstDIA.Fields("Batch No").Value
                                
                                Case "STOCK_ID"
                                    rstGrid2Off.Fields("Stock_ID").Value = rstDIA.Fields("Stock_ID").Value
                                    
                                Case "PROD_HANDLING"
                                    rstGrid2Off.Fields("Prod_Handling").Value = rstTmp.Fields("Prod_Handling").Value
                                    
                                Case "IN_ORIG_PACKAGES_QTY"
                                    rstGrid2Off.Fields("In_Orig_Packages_Qty").Value = rstDIA.Fields("In_Orig_Packages_Qty").Value
                                    
                                Case "IN_ORIG_GROSS_WEIGHT"
                                    rstGrid2Off.Fields("In_Orig_Gross_Weight").Value = rstDIA.Fields("In_Orig_Gross_Weight").Value
                                    
                                Case "IN_ORIG_NET_WEIGHT"
                                    rstGrid2Off.Fields("In_Orig_Net_Weight").Value = rstDIA.Fields("In_Orig_Net_Weight").Value
                                    
                                Case "STOCK CARD NO"
                                    rstGrid2Off.Fields("Stock Card No").Value = rstTmp.Fields("Stock Card No").Value
            
                                Case "TOTAL OUT QTY"
                                    rstGrid2Off.Fields("Total Out Qty").Value = rstDIA.Fields("In_TotalOut_Qty_Wgt").Value - rstDIA.Fields("Out_Packages_Qty_Wgt").Value
                                    
                                Case "DOCUMENT NO"
                                    'rstGrid2Off.Fields("Document No").Value = rstTmp.Fields("Document No").Value
                                    rstGrid2Off.Fields("Document No").Value = rstDIA.Fields("Document No").Value
                                
                                Case "QUANTITY/WEIGHT"
                                    
                                    lngProdHandling = rstTmp.Fields("Prod_Handling").Value
                                    Select Case lngProdHandling
                                        Case 0
                                            rstGrid2Off.Fields("Quantity/Weight").Value = rstDIA.Fields("In_Orig_Packages_Qty").Value
                                        Case 1
                                            rstGrid2Off.Fields("Quantity/Weight").Value = rstDIA.Fields("In_Orig_Gross_Weight").Value
                                        Case 2
                                            rstGrid2Off.Fields("Quantity/Weight").Value = rstDIA.Fields("In_Orig_Net_Weight").Value
                                    End Select
                                    
                                Case "RESERVED"
                                    rstGrid2Off.Fields("Reserved").Value = rstDIA.Fields("In_Reserved_Qty_Wgt").Value
                                    
                                Case "AVAILABLE"
                                    rstGrid2Off.Fields("Available").Value = rstDIA.Fields("In_Avl_Qty_Wgt").Value + rstDIA.Fields("Out_Packages_Qty_Wgt").Value
                                    
                                Case "QTY TO RESERVE"
                                    
                                Case Else
                                
                            End Select
                        
                        Next
                        
                        rstGrid2Off.Update
                    End If
                    
                End If
                
                ADORecordsetClose rstTmp
            Else
                rstGrid2Off.Fields("Total Out Qty").Value = rstGrid2Off.Fields("Total Out Qty").Value - rstDIA.Fields("Out_Packages_Qty_Wgt").Value
                rstGrid2Off.Fields("Available").Value = rstGrid2Off.Fields("Available").Value + rstDIA.Fields("Out_Packages_Qty_Wgt").Value
                rstGrid2Off.Update
            End If
            
            rstDIA.MoveNext
        Loop
    End If

    ADORecordsetClose rstDIA
    
    ADODisconnectDB conHistory
    
End Sub
