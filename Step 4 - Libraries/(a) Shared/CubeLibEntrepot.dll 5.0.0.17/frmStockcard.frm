VERSION 5.00
Begin VB.Form frmStockcard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Card"
   ClientHeight    =   2055
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   4140
   Icon            =   "frmStockcard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStockCard 
      Height          =   1455
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   3930
      Begin VB.CommandButton cmdProductNo 
         Caption         =   "..."
         Height          =   315
         Left            =   3405
         TabIndex        =   6
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox txtStockCardNum 
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   7
         Top             =   960
         Width           =   2040
      End
      Begin VB.TextBox txtProductNo 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblEntrepotNumX 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblEntrepotNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrepot Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblStockCardNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Card No.:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label lblProductNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number:"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   660
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2790
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmStockcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pckProducts As PCubeLibEntrepot.cProducts
Private pckStockCard As PCubeLibEntrepot.cStockCard
Private pckStockProd As PCubeLibEntrepot.cStockProd
Private blnAlienCall As Boolean         'Currently indicates if called by frmStockProdPicklist.

'For E.Type, E.Num, P.Num, and SC.Num uniqueness check
Private m_rstStockCard As ADODB.Recordset

Public strEntrepotType As String
Public strEntrepotNum As String
Public bytSequential As Byte
Public strSeqStart As String
Public lngStockNum As Long
Public lngStock_Id As Long

Private Sub cmdCancel_Click()
    If blnAlienCall = True Then frmStockProdPicklist.blnCancelled = True
    frmStockcardPicklist.blnAddCancel = True
    CleanUpADO
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Checks to make sure all required fields have values.
    If Validation = False Then Exit Sub
    
    frmStockcardPicklist.blnAddCancel = False
    
    If StockcardExistsInDatabase(txtStockCardNum.Text) = True Then
        MsgBox "The stockcard number '" & txtStockCardNum.Text & "' already exists. " & vbCrLf & _
                "Please enter a new stockcard number. ", vbInformation, Translate(2214)
        Exit Sub
    End If
    
    If blnAlienCall = False Then
        'Adds stock card info in Stock Card picklist.
        With pckStockCard.m_rstPass2GridOff
            .AddNew
            .Fields("Stock Card No").Value = txtStockCardNum.Text
            .Fields("Product ID").Value = txtProductNo.Tag
            .Fields("Entrepot ID").Value = lblEntrepotNumX.Tag
            .Fields("New").Value = -1
            'Appends the "9" to the length if it's a ceiling number. E.g. 9, 99, 999, etc.
            If Len(Trim(Replace(txtStockCardNum.Text, "9", ""))) = 0 Then
                .Fields("Length").Value = Len(txtStockCardNum.Text) & "9"
            Else
                .Fields("Length").Value = Len(txtStockCardNum.Text)
            End If
            .Update
            
            'Save the values of the temporary stock card/s.
            With frmStockProdPicklist.m_rstNewStockOff
                .AddNew
                .Fields("Stock Card No").Value = txtStockCardNum.Text
                .Fields("Product ID").Value = txtProductNo.Tag
                .Update
            End With

            CleanUpADO
            Unload Me
        End With
    Else
            
        With m_rstStockCard
            'This is done to prevent the Stock_IDs from jumping into another number.
            If frmStockProdPicklist.m_rstNewStockOff.RecordCount > 0 Then
                lngStock_Id = lngStock_Id + 1
            Else
                'Commits to database so users can select their newly added Stock Card.
                'This also prevents the error of missing Stock ID.
                ADORecordsetOpen "SELECT Stock_ID AS [Stock ID], Stock_Card_Num AS [Stock Card No], Prod_ID AS [Product ID] " & _
                      "FROM StockCards", _
                      pckStockProd.m_conSADBEL, m_rstStockCard, adOpenKeyset, adLockOptimistic
                      
                '.Open "SELECT Stock_ID AS [Stock ID], Stock_Card_Num AS [Stock Card No], Prod_ID AS [Product ID] " & _
                      "FROM StockCards", _
                      pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic
                If Not (.BOF And .EOF) Then
                    .MoveLast
                    lngStock_Id = .Fields("Stock ID").Value + 1
                End If
            End If
        
            With frmStockProdPicklist.m_rstPass2GridOff
                'Adds to Grid's recordset including the Stock ID.
                .AddNew
                .Fields("Stock ID").Value = lngStock_Id
                .Fields("Stock Card No").Value = txtStockCardNum.Text
                .Fields("Product ID").Value = txtProductNo.Tag
                .Fields("Entrepot ID").Value = lblEntrepotNumX.Tag
                'Appends the "9" to the length if it's a ceiling number. E.g. 9, 99, 999, etc.
                If Len(Trim(Replace(txtStockCardNum.Text, "9", ""))) = 0 Then
                    .Fields("Length").Value = Len(txtStockCardNum.Text) & "9"
                Else
                    .Fields("Length").Value = Len(txtStockCardNum.Text)
                End If
                .Update
                
                'Save the values of the temporary stock card/s.
                With frmStockProdPicklist.m_rstNewStockOff
                    .AddNew
                    .Fields("Stock ID").Value = lngStock_Id
                    .Fields("Stock Card No").Value = txtStockCardNum.Text
                    .Fields("Product ID").Value = txtProductNo.Tag
                    .Update
                End With
                
                frmStockProdPicklist.blnCancelled = False
                CleanUpADO
                Unload Me
            End With
            
            ADORecordsetClose m_rstStockCard
        End With
    End If
End Sub

Private Sub cmdProductNo_Click()
    'Loads Products simple picklist.
    'Outputs selected Product number and corresponding Entrepot number.
    With pckStockCard
        pckProducts.ShowProducts 2, Me, .mvarConn_Sadbel, .mvarConn_Taric, .mvarLanguage, .mvarTaricProp, ResourceHandler
    End With
    
    PopGrid
End Sub

Private Sub PopGrid()
    Dim lngCtr As Long
    Dim lngSafeLength As Long
    Dim blnSafe As Boolean
    
    'Becomes True when length and value are ok.
    blnSafe = False
    
    With pckStockCard.m_rstPass2GridOff
        If Not (.BOF Or .EOF) Then
            .Filter = ""
            .MoveFirst
        End If
        
        'Display corresponding Entrepot Num in form.
        Select Case bytSequential
            Case 0
                'Do nothing if product selection was cancelled.
                If Len(lblEntrepotNumX.Tag) <> 0 Then
                    lblEntrepotNumX.Caption = strEntrepotType & "-" & strEntrepotNum
                    
                    'Begin safe length/value search with the length of configured Entrepot_Starting_Num.
                    lngSafeLength = Len(strSeqStart)
                    Do Until blnSafe = True
                        'Retrieve involved records that have all digits as "9".
                        .Filter = "[Entrepot ID] = " & lblEntrepotNumX.Tag & " AND [Length] = " & lngSafeLength & "9"
                        
                        If .BOF Or .EOF Then
                            'Flag safe when highest number for current Stock Card num length is not at ceiling.
                            blnSafe = True
                        Else
                            'Otherwise, move to next Stock Card num length.
                            lngSafeLength = lngSafeLength + 1
                        End If
                    Loop
                    
                    .Filter = "[Entrepot ID] = " & lblEntrepotNumX.Tag & " AND [Length] = " & lngSafeLength
                    
                    'Will not bother performing checking if table is empty.
                    If Not (.BOF Or .EOF) Then
                        'Ascending sort to find highest value in table.
                        .Sort = "[Stock Card No] ASC"
                        .MoveLast
                        
                        'Determines what value to put in Stock Card number after selecting a Stock Card.
                        If .Fields("Stock Card No").Value >= Val(strSeqStart) Then
                            '..the highest Stock Card Num + 1..
                            txtStockCardNum.Text = .Fields("Stock Card No").Value + 1
                            If Len(.Fields("Stock Card No").Value) > Len(txtStockCardNum.Text) Then
                                txtStockCardNum.Text = String$(Len(.Fields("Stock Card No").Value) - Len(txtStockCardNum.Text), "0") & txtStockCardNum.Text
                            End If
                        Else
                            If lngSafeLength = Len(strSeqStart) Then
                                '..the Starting Num according to Entrepot.
                                txtStockCardNum.Text = strSeqStart
                            ElseIf lngSafeLength >= 10 Then
                                'Default maximum length of Stock Card num has been reached.
                                txtStockCardNum.Text = ""
                            Else
                                txtStockCardNum = "1" & String$(lngSafeLength - 1, "0")
                            End If
                        End If
                    Else
                        If lngSafeLength = Len(strSeqStart) Then
                            'Uses Starting Num according to Entrepot.
                            txtStockCardNum.Text = strSeqStart
                        ElseIf lngSafeLength >= 10 Then
                            'Default maximum length of Stock Card num has been reached.
                            txtStockCardNum.Text = ""
                        Else
                            txtStockCardNum.Text = "1" & String$(lngSafeLength - 1, "0")
                        End If
                    End If
                End If
            Case 1
                txtStockCardNum.Text = Empty
        End Select
    End With
End Sub


'Called from Stockcard picklist.
Public Sub Pre_Load(ByRef cpiStockCard As PCubeLibEntrepot.cStockCard, ByVal MyResourceHandler As Long)

    ResourceHandler = MyResourceHandler
    
    'To access rstPass2Grid and required parameters for products picklist loading.
    Set pckProducts = New PCubeLibEntrepot.cProducts
    Set pckStockCard = cpiStockCard
    
    'For verification of entered values.
    ADORecordsetOpen "SELECT E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                      "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StockCard_Numbering AS [Numbering], " & _
                      "E.Entrepot_Starting_Num AS [Starting Num], P.Prod_Num AS [Prod Num] " & _
                      "FROM Entrepots [E] INNER JOIN Products [P] ON E.Entrepot_ID = P.Entrepot_ID " & _
                      "ORDER BY E.Entrepot_ID", _
                      cpiStockCard.mvarConn_Sadbel, m_rstStockCard, adOpenKeyset, adLockOptimistic
                      
    'm_rstStockCard.Open "SELECT E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                      "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StockCard_Numbering AS [Numbering], " & _
                      "E.Entrepot_Starting_Num AS [Starting Num], P.Prod_Num AS [Prod Num] " & _
                      "FROM Entrepots [E] INNER JOIN Products [P] ON E.Entrepot_ID = P.Entrepot_ID " & _
                      "ORDER BY E.Entrepot_ID", _
                      cpiStockCard.mvarConn_Sadbel, adOpenKeyset, adLockOptimistic
    
    'If called by frmStockCard.
    blnAlienCall = False
    
    Me.Show vbModal
End Sub

'Called from Stock/Prod picklist.
Public Sub Pre_Load2(EntrepotID As Long, EntrepotType As String, EntrepotNum As String, ProductID As String, ProductNum As String, _
                     Sequential As Byte, StartingNum As String, StockCardNoHigh As String, _
                     ByRef cpiStockProd As PCubeLibEntrepot.cStockProd, ByVal MyResourceHandler As Long)
                     
    ResourceHandler = MyResourceHandler
    
    Set pckStockProd = cpiStockProd
    
    lblEntrepotNumX.Tag = EntrepotID
    lblEntrepotNumX.Caption = EntrepotType & "-" & EntrepotNum
    txtProductNo.Tag = ProductID
    txtProductNo.Text = ProductNum
    txtProductNo.Enabled = False
    cmdProductNo.Enabled = False
    
    Select Case Sequential
        Case 0
            txtStockCardNum.Text = StockCardNoHigh
        Case 1
            txtStockCardNum.Text = Empty
    End Select
    
    'If not called by frmStockCard.
    blnAlienCall = True
    
    On Error GoTo ErrorHandler
    Me.Show vbModal
    
ErrorHandler:
    If Err.Number = 364 Then
        Err.Clear
        Exit Sub
    End If
    
End Sub

Private Sub CleanUpADO()
    'Clean up.
    If blnAlienCall = False Then
        ADORecordsetClose m_rstStockCard
    End If
    
    Set pckStockCard = Nothing
    Set pckProducts = Nothing
End Sub

Private Function Validation() As Boolean
    'Prevents user from clicking OK when there are blank fields.
    Validation = True
    If Len(txtProductNo.Tag) = 0 Or txtProductNo.Tag = "0" Then
        Validation = False
        MsgBox Translate(2236), vbOKOnly + vbInformation, Translate(2214)
        txtProductNo.SetFocus
        Exit Function
    End If
    If Len(txtProductNo.Text) = 0 Then
        Validation = False
        txtProductNo.SetFocus
        Exit Function
    End If
    If Len(txtStockCardNum.Text) = 0 Then
        Validation = False
        MsgBox Translate(2253), vbOKOnly + vbInformation, Translate(2214)
        txtStockCardNum.SetFocus
        Exit Function
    End If
    
    If blnAlienCall = False Then
    'Called from Stockcard picklist.
        With pckStockCard.m_rstPass2GridOff
            .Filter = "[Entrepot ID] = " & lblEntrepotNumX.Tag & " " & _
                      "AND [Product ID] = " & txtProductNo.Tag & " " & _
                      "AND [Stock Card No] = " & txtStockCardNum.Text
            
            If Not (.RecordCount = 0) Then
                Validation = False
                MsgBox Translate(2289), vbOKOnly + vbInformation, Translate(2214)
                txtStockCardNum.SetFocus
                Exit Function
            End If
        End With
    ElseIf blnAlienCall = True Then
    'Called from StockProd picklist.
        With frmStockProdPicklist.jgxPicklist.ADORecordset
            .Filter = "[Entrepot ID] = " & lblEntrepotNumX.Tag & " " & _
            "AND [Product ID] = " & txtProductNo.Tag & " " & _
            "AND [Stock Card No] = " & txtStockCardNum.Text
            
            If Not (.RecordCount = 0) Then
                Validation = False
                MsgBox Translate(2289), vbOKOnly + vbInformation, Translate(2214)
                txtStockCardNum.SetFocus
                Exit Function
            End If
        End With
    End If
End Function

Private Sub txtProductNo_LostFocus()
    'Only performs this check if user has entered something in textbox.
    If Len(txtProductNo.Text) <> 0 Then
        'Verifies Product validity and existence on lost focus.
        txtProductNo.Tag = GetProd_ID(txtProductNo.Text, pckStockCard.mvarConn_Sadbel, False)
        If txtProductNo.Tag = 0 Then
            If MsgBox(Translate(2287) & vbCrLf & _
                   Translate(2288), vbYesNo + vbInformation, Translate(2214)) = vbYes Then
                cmdProductNo_Click
            End If
        End If
    End If
End Sub

Private Sub txtStockCardNum_KeyPress(KeyAscii As Integer)
    'Prevents user from entering non-numeric characters.
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End Sub

Private Function StockcardExistsInDatabase(StockCardNo As String) As Boolean
    Dim strSQL As String
    Dim rstStockCard As ADODB.Recordset
    
        strSQL = "Select STOCKCARDS.Stock_Card_Num As [Stockcard] FROM STOCKCARDS "
        strSQL = strSQL & " INNER JOIN (Products INNER JOIN Entrepots ON "
        strSQL = strSQL & " Products.Entrepot_ID = Entrepots.Entrepot_ID) ON "
        strSQL = strSQL & " Stockcards.Prod_ID = Products.Prod_ID "
        strSQL = strSQL & " WHERE Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '"
        strSQL = strSQL & lblEntrepotNumX.Caption & "' AND STOCKCARDS.Stock_Card_Num = '" & StockCardNo & "'"
    ADORecordsetOpen strSQL, pckStockProd.m_conSADBEL, rstStockCard, adOpenKeyset, adLockOptimistic
    'rstStockCard.Open strSQL, pckStockProd.m_conSADBEL, adOpenKeyset, adLockOptimistic

    If rstStockCard.EOF And rstStockCard.BOF Then
        StockcardExistsInDatabase = False
    Else
        StockcardExistsInDatabase = True
    End If
    
    ADORecordsetClose rstStockCard
End Function

