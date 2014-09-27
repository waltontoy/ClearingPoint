VERSION 5.00
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product"
   ClientHeight    =   4080
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   5415
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2212"
   Begin VB.Frame fraProducts 
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdEntrepotNum 
         Caption         =   "..."
         Height          =   315
         Left            =   4650
         TabIndex        =   1
         Top             =   250
         Width           =   315
      End
      Begin VB.CommandButton cmdTaricCode 
         Caption         =   "..."
         Height          =   315
         Left            =   4650
         TabIndex        =   4
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4650
         TabIndex        =   9
         Top             =   2520
         Width           =   315
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   4650
         TabIndex        =   7
         Top             =   2160
         Width           =   315
      End
      Begin VB.TextBox txtCountryExportDescription 
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2520
         Width           =   2220
      End
      Begin VB.TextBox txtCountryOriginDescription 
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   2220
      End
      Begin VB.TextBox txtEntrepotNum 
         Height          =   315
         Left            =   1965
         MaxLength       =   19
         TabIndex        =   0
         Top             =   240
         Width           =   2700
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "757"
         Top             =   3285
         Width           =   1275
      End
      Begin VB.ComboBox cboHandling 
         Height          =   315
         ItemData        =   "frmProducts.frx":08CA
         Left            =   1965
         List            =   "frmProducts.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2880
         Width           =   3000
      End
      Begin VB.TextBox txtCountryExport 
         Height          =   315
         Left            =   1965
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtProductNum 
         Height          =   315
         Left            =   1965
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3000
      End
      Begin VB.TextBox txtDescription 
         Height          =   795
         Left            =   1965
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   3000
      End
      Begin VB.TextBox txtTaricCode 
         Height          =   315
         Left            =   1965
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   960
         Width           =   2700
      End
      Begin VB.TextBox txtCountryOrigin 
         Height          =   315
         Left            =   1965
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblEntrepotNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrepot Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Tag             =   "2198"
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label lblHandling 
         BackStyle       =   0  'Transparent
         Caption         =   "Handling:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Tag             =   "2219"
         Top             =   2940
         Width           =   1755
      End
      Begin VB.Label lblCountryExport 
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Export:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Tag             =   "2196"
         Top             =   2580
         Width           =   1755
      End
      Begin VB.Label lblProductNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Tag             =   "2201"
         Top             =   1380
         Width           =   1755
      End
      Begin VB.Label lblTaricCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Taric Code:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Tag             =   "2275"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblCountryOrigin 
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Origin:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Tag             =   "2195"
         Top             =   2220
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Tag             =   "179"
      Top             =   3675
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Tag             =   "178"
      Top             =   3675
      Width           =   1215
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is called in the Products picklist
Option Explicit

Private m_conConnection As ADODB.Connection
Private m_conTARIC As ADODB.Connection

Private m_rstProducts As ADODB.Recordset

Private mblnCancel As Boolean
Private mTaricProperties As Long
Private mLanguage As String
Private ButtonType As PCubeLibPick.ButtonType
Private pckCountry As PCubeLibPick.CPicklist
Private pckEntrepot As PCubeLibEntrepot.cEntrepot
Private pckTaricCodes As PCubeLibPick.CPicklist

Private bytCtryKeys As Byte          '0-No, 1-Yes (for VB bug in GotFocus not working)
Private strBlah As String
Private bytCtryOFound As Byte       '0-No, 1-Yes
Private bytCtryEFound As Byte       '0-No, 1-Yes
Private strCtryOrigin As String
Private strCtryExport As String
Private strStartingNum As String

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Public Sub MyLoad(ByRef rstRecord As ADODB.Recordset, Button As PCubeLibPick.ButtonType, ByRef Cancel As Boolean, ByVal Language As String, ByRef Connection As ADODB.Connection, ByRef ConnectionTaric As ADODB.Connection, TaricProperties As Long, ByVal MyResourceHandler As Long)
    
    ResourceHandler = MyResourceHandler
    mTaricProperties = TaricProperties
    
    'Moved reference here since Form_Load (called after LoadResStrings) checks ButtonType.
    ButtonType = Button
    'Moved here to avoid error in modify button.
    Set m_conConnection = Connection
    Set m_conTARIC = ConnectionTaric
    Set m_rstProducts = rstRecord
    modGlobals.LoadResStrings Me, True
    
    mLanguage = Language
    LoadValues Button
    Me.Show vbModal
    Cancel = mblnCancel
End Sub

Private Sub LoadValues(Button As PCubeLibPick.ButtonType)
    Dim strCountrySQL As String
    Dim rstCountry As ADODB.Recordset
    
    Select Case Button
        Case cpiAdd
            Me.txtEntrepotNum = ""
            Me.txtProductNum = ""
            Me.txtTaricCode = ""
            Me.txtDescription = ""
            Me.txtCountryOrigin = ""
            Me.txtCountryExport = ""
            Me.cboHandling.ListIndex = 0
            Me.chkArchive.Value = False
        Case Else
            
                strCountrySQL = "SELECT Code AS [Key Code], Code as [CODE], [Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Key Description] " & _
                        "FROM [PICKLIST MAINTENANCE " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] INNER JOIN [PICKLIST DEFINITION] ON " & _
                        "[PICKLIST MAINTENANCE " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "].[INTERNAL CODE] = [PICKLIST DEFINITION].[INTERNAL CODE] " & _
                        "WHERE Document = 'Import' and [BOX CODE] = 'C2'"
            ADORecordsetOpen strCountrySQL, m_conConnection, rstCountry, adOpenKeyset, adLockOptimistic
            'rstCountry.Open strCountrySQL, m_conConnection, adOpenForwardOnly, adLockOptimistic
            If Not (rstCountry.EOF And rstCountry.BOF) Then
                rstCountry.MoveFirst
                rstCountry.Find "[Key Code] = '" & m_rstProducts("Origin Code").Value & "'"
            
                If rstCountry.EOF = False Then
                    Me.txtCountryOriginDescription.Text = rstCountry("Key Description").Value
                End If
                
                rstCountry.MoveFirst
                rstCountry.Find "[Key Code] = '" & m_rstProducts("Export Code").Value & "'"
                
                If rstCountry.EOF = False Then
                    Me.txtCountryExportDescription.Text = rstCountry("Key Description").Value
                End If
            End If
            
            Me.txtEntrepotNum.Text = m_rstProducts.Fields("Entrepot Type") & "-" & m_rstProducts.Fields("Entrepot Num")
            Me.txtEntrepotNum.Tag = m_rstProducts("Entrepot ID")
            Me.txtProductNum.Text = IIf(IsNull(m_rstProducts.Fields("Product Num")), "", m_rstProducts.Fields("Product Num"))
            Me.txtTaricCode.Text = IIf(IsNull(m_rstProducts.Fields("Taric Code")), "", m_rstProducts.Fields("Taric Code"))
            Me.txtDescription.Text = IIf(IsNull(m_rstProducts.Fields("Description")), "", m_rstProducts.Fields("Description"))
            Me.txtCountryOrigin = IIf(IsNull(m_rstProducts("Origin Code").Value), "", m_rstProducts("Origin Code").Value)
            Me.txtCountryExport = IIf(IsNull(m_rstProducts("Export Code").Value), "", m_rstProducts("Export Code").Value)
            Me.cboHandling.ListIndex = m_rstProducts!Prod_Handling
            Me.chkArchive.Value = IIf(m_rstProducts!Prod_Archive, 1, 0)
            
            ADORecordsetClose rstCountry

    End Select
End Sub

Private Sub cmdCountry_Click(Index As Integer)
    Dim gsdCountry As PCubeLibPick.CGridSeed
    Dim strCountrySQL As String
    
    Set pckCountry = New CPicklist
    Set gsdCountry = New CGridSeed
    
    Set gsdCountry = pckCountry.SeedGrid("Key Code", 1300, "Left", "Key Description", 2970, "Left")
    
    ' The primary key is mentioned twice to conform to the design of the picklist class.
    strCountrySQL = "SELECT Code AS [Key Code], Code as [CODE], [Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Key Description] " & _
                    "FROM [PICKLIST MAINTENANCE " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] INNER JOIN [PICKLIST DEFINITION] ON " & _
                    "[PICKLIST MAINTENANCE " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "].[INTERNAL CODE] = [PICKLIST DEFINITION].[INTERNAL CODE] " & _
                    "WHERE Document = 'Import' and [BOX CODE] = 'C2'"
    With pckCountry
        Select Case Index
            Case 0
                .Search True, "Key Code", Trim(txtCountryOrigin.Text)
            Case 1
                .Search True, "Key Code", Trim(txtCountryExport.Text)
        End Select
        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick Me, cpiSimplePicklist, m_conConnection, strCountrySQL, "Key Code", "Countries", vbModal, gsdCountry, , , True, cpiKeyF2
        
        If Not .SelectedRecord Is Nothing Then
            Select Case Index
                Case 0
                    txtCountryOrigin.Text = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    txtCountryOriginDescription.Text = .SelectedRecord.RecordSource.Fields("Key Description").Value
                Case 1
                    txtCountryExport.Text = .SelectedRecord.RecordSource.Fields("Key Code").Value
                    txtCountryExportDescription.Text = .SelectedRecord.RecordSource.Fields("Key Description").Value
            End Select
        End If
    End With
    
    Set gsdCountry = Nothing
    Set pckCountry = Nothing
End Sub

Private Sub cmdEntrepotNum_Click()
    'Call Bryan's Entrepot procedure here
    Set pckEntrepot = New PCubeLibEntrepot.cEntrepot
    pckEntrepot.ShowEntrepot Me, m_conConnection, True, mLanguage, ResourceHandler, Me.txtEntrepotNum.Name, Val(txtEntrepotNum.Tag)
    'Added by BCo to fix problem in StockProd having blank Stock Card number when creating
    'a new Stock Card using a newly created Product number.
    If pckEntrepot.Cancelled = False Then strStartingNum = pckEntrepot.StartingNum
    
    Set pckEntrepot = Nothing
End Sub

Private Sub cmdTaricCode_Click()
    frm_taricmaintenance.My_Load Me, IIf((Len(txtTaricCode.Text) = 0), Trim(txtProductNum.Text), Trim(txtTaricCode.Text))
End Sub

Private Sub Form_Load()
    Dim rstTmpProductsWeight As ADODB.Recordset
    Dim strTmp1 As String
    

    If mTaricProperties = 0 Then
        cmdTaricCode.Enabled = False
    End If
    If ButtonType = cpiModify Then
        If Not IsNull(m_rstProducts.Fields("Entrepot ID")) Then
            strTmp1 = "SELECT Products.Prod_ID, SUM(In_Avl_Qty_Wgt) as NATIRANGWEIGHT " & _
                        "FROM (Products INNER JOIN StockCards ON Products.Prod_ID = StockCards.Prod_ID) " & _
                        "INNER JOIN Inbounds ON StockCards.Stock_ID = Inbounds.Stock_ID " & _
                        "Where Products.Entrepot_ID = " & m_rstProducts.Fields("Entrepot ID") & " " & _
                        "GROUP BY Products.Prod_ID"
            ADORecordsetOpen strTmp1, m_conConnection, rstTmpProductsWeight, adOpenKeyset, adLockOptimistic
            'rstTmpProductsWeight.Open strTmp1, m_conConnection, adOpenKeyset, adLockOptimistic
            If Not (rstTmpProductsWeight.EOF And rstTmpProductsWeight.BOF) Then
                rstTmpProductsWeight.MoveFirst
                
                If rstTmpProductsWeight("NATIRANGWEIGHT") > 0 Then
                    cboHandling.Enabled = False
                End If
                
            End If
            ADORecordsetClose rstTmpProductsWeight
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mblnCancel = True
    End If
End Sub

Private Sub cmdOK_Click()
    Dim rstProductVerify As ADODB.Recordset
    
    'Checks to make sure all required fields have values.
    If Validation = False Then Exit Sub
    
    MousePointer = vbHourglass
        
    If ButtonType = cpiAdd Or ButtonType = cpiCopy Then
        If Trim(Me.txtEntrepotNum.Tag) = "" Then
            MousePointer = vbDefault
            'MsgBox "Entrepot Number box not filled.", vbOKOnly + vbInformation, "Products"
            MsgBox Translate(2194), vbOKOnly + vbInformation, Translate(2216)
            Exit Sub
        Else
            ADORecordsetOpen "SELECT P.Prod_Num AS [Product Num], E.Entrepot_Type AS [Entrepot Type], " & _
                                  "E.Entrepot_Num AS [Entrepot Num] " & _
                                  "FROM Entrepots [E] INNER JOIN Products [P] " & _
                                  "ON E.Entrepot_ID = P.Entrepot_ID " & _
                                  "WHERE P.Prod_Num = '" & txtProductNum.Text & "' " & _
                                  "AND E.Entrepot_Type = '" & GetEntrepotType(txtEntrepotNum.Text) & "' " & _
                                  "AND E.Entrepot_Num = '" & GetEntrepotNum(txtEntrepotNum.Text) & "'", _
                                  m_conConnection, rstProductVerify, adOpenKeyset, adLockOptimistic
                                  
            'rstProductVerify.Open "SELECT P.Prod_Num AS [Product Num], E.Entrepot_Type AS [Entrepot Type], " & _
                                  "E.Entrepot_Num AS [Entrepot Num] " & _
                                  "FROM Entrepots [E] INNER JOIN Products [P] " & _
                                  "ON E.Entrepot_ID = P.Entrepot_ID " & _
                                  "WHERE P.Prod_Num = '" & txtProductNum.Text & "' " & _
                                  "AND E.Entrepot_Type = '" & GetEntrepotType(txtEntrepotNum.Text) & "' " & _
                                  "AND E.Entrepot_Num = '" & GetEntrepotNum(txtEntrepotNum.Text) & "'", _
                                  m_conConnection, adOpenKeyset, adLockOptimistic
            
            If Not (rstProductVerify.BOF And rstProductVerify.EOF) Then
                MousePointer = vbDefault
                'MsgBox "The Product Number entered already exists for this Entrepot Number.", vbOKOnly + vbInformation, "Products"
                MsgBox Translate(2228), vbOKOnly + vbInformation, Translate(2216)
                Exit Sub
            End If
            
            ADORecordsetClose rstProductVerify
        End If
    End If
    mblnCancel = False
    
    With m_rstProducts
        .Fields("Entrepot ID").Value = Me.txtEntrepotNum.Tag
        .Fields("Entrepot Type").Value = GetEntrepotType(Me.txtEntrepotNum.Text)
        .Fields("Entrepot Num").Value = GetEntrepotNum(Me.txtEntrepotNum.Text)
        .Fields("Product Num").Value = Me.txtProductNum.Text
        .Fields("Taric Code").Value = Me.txtTaricCode.Text
        .Fields("Description").Value = Me.txtDescription.Text
        .Fields("Origin Code").Value = Me.txtCountryOrigin.Text
        .Fields("Origin Description").Value = Me.txtCountryOriginDescription
        .Fields("Export Description").Value = Me.txtCountryExportDescription
        .Fields("Export Code").Value = Me.txtCountryExport.Text
        .Fields("Prod_Handling").Value = cboHandling.ListIndex
        .Fields("Prod_Archive").Value = Me.chkArchive.Value
        .Fields("Starting Num").Value = strStartingNum
    End With
    
    m_rstProducts.Update
    ' TO DO FOR CP.NET

    Me.MousePointer = vbHourglass
    Me.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub txtCountryOrigin_GotFocus()
    'Sets flag to signify execution (for VB bug in GotFocus not working).
    bytCtryKeys = 1

    strCtryOrigin = txtCountryOrigin.Text
End Sub

Private Sub txtCountryOrigin_KeyDown(KeyCode As Integer, Shift As Integer)
    'First key triggers GotFocus only if GotFocus was not executed earlier (for VB bug in GotFocus not working).
    If bytCtryKeys = 0 Then txtCountryOrigin_GotFocus
    
    If KeyCode = vbKeyF2 Then
        cmdCountry_Click (0)
    End If
End Sub

Private Sub txtCountryOrigin_LostFocus()
    'Resets Country Code textbox workaround counter (for VB bug in GotFocus not working).
    bytCtryKeys = 0
    
    'Performs check and auto description loading for Country.
    If (strCtryOrigin = txtCountryOrigin.Text) Then
        If Len(txtCountryOriginDescription.Text) > 0 And Len(txtCountryOrigin.Text) = 3 Then
            bytCtryOFound = 1
        Else
            bytCtryOFound = 0
        End If
    ElseIf Val(txtCountryOrigin.Text) = 0 And Not (txtCountryOrigin.Text = "000") Then
        txtCountryOriginDescription.Text = Empty
    Else
        If Len(txtCountryOrigin.Text) = 3 Then
            bytCtryOFound = 0
        Else
''The code below blanks the country code/description upon error.
''--------------------------------------------------------------
''            txtCountryOrigin.Text = Empty
''            txtCountryOriginDescription.Text = Empty
''--------------------------------------------------------------
            'Prompts to open picklist.
            If MsgBox(Translate(2229) & vbCrLf & _
                      Translate(2230), vbYesNo + vbInformation, _
                      "Stock Card / Products") = vbYes Then
                cmdCountry_Click (0)
            Else
                'Revert to previous country code.
                txtCountryOrigin.Text = strCtryOrigin
            End If
            
            bytCtryOFound = 1
        End If
    End If
    
    If bytCtryOFound = 0 And Len(txtCountryOrigin.Text) = 3 Then
        strBlah = GetCountryDesc(txtCountryOrigin.Text, m_conConnection, mLanguage)
        'Validates entered code based on description.
        If strBlah = "ALL YOUR BASE ARE BELONG TO US" Then
''The code below blanks the country code/description upon error.
''--------------------------------------------------------------
''            txtCountryOrigin.Text = Empty
''            txtCountryOriginDescription.Text = Empty
''--------------------------------------------------------------
            'Prompts to open picklist.
            If MsgBox(Translate(2229) & vbCrLf & _
                      Translate(2230), vbYesNo + vbInformation, _
                      "Stock Card / Products") = vbYes Then
                cmdCountry_Click (0)
            Else
                'Revert to previous country code.
                txtCountryOrigin.Text = strCtryOrigin
            End If
        Else
            txtCountryOriginDescription.Text = strBlah
        End If
        strBlah = Empty
        bytCtryOFound = 1
    End If
End Sub

Private Sub txtCountryExport_GotFocus()
    'Sets flag to signify execution (for VB bug in GotFocus not working).
    bytCtryKeys = 1
        
    strCtryExport = txtCountryExport.Text
End Sub

Private Sub txtCountryExport_KeyDown(KeyCode As Integer, Shift As Integer)
    'First key triggers GotFocus only if GotFocus was not executed earlier (for VB bug in GotFocus not working).
    If bytCtryKeys = 0 Then txtCountryExport_GotFocus
    
    If KeyCode = vbKeyF2 Then
        cmdCountry_Click (1)
    End If
End Sub

Private Sub txtCountryExport_LostFocus()
    'Resets Country Code textbox workaround counter (for VB bug in GotFocus not working).
    bytCtryKeys = 0
        
    'Performs check and auto description loading for Country.
    If (strCtryExport = txtCountryExport.Text) Then
        If Len(txtCountryExportDescription.Text) > 0 And Len(txtCountryExport.Text) = 3 Then
            bytCtryEFound = 1
        Else
            bytCtryEFound = 0
        End If
    ElseIf Val(txtCountryExport.Text) = 0 And Not (txtCountryExport.Text = "000") Then
        txtCountryExportDescription.Text = Empty
    Else
        If Len(txtCountryExport.Text) = 3 Then
            bytCtryEFound = 0
        Else
''The code below blanks the country code/description upon error.
''--------------------------------------------------------------
''            txtCountryExport.Text = Empty
''            txtCountryExportDescription.Text = Empty
''--------------------------------------------------------------
            'Prompts to open picklist.
            If MsgBox(Translate(2272) & vbCrLf & _
                      Translate(2231), vbYesNo + vbInformation, _
                      "Stock Card / Products") = vbYes Then
            cmdCountry_Click (1)
            Else
                'Revert to previous country code.
                txtCountryExport.Text = strCtryExport
            End If
            
            bytCtryEFound = 1
        End If
    End If
    
    If bytCtryEFound = 0 And Len(txtCountryExport.Text) = 3 Then
        strBlah = GetCountryDesc(txtCountryExport.Text, m_conConnection, mLanguage)
        'Validates entered code based on description.
        If strBlah = "ALL YOUR BASE ARE BELONG TO US" Then
''The code below blanks the country code/description upon error.
''--------------------------------------------------------------
''            txtCountryExport.Text = Empty
''            txtCountryExportDescription.Text = Empty
''--------------------------------------------------------------
            'Prompts to open picklist.
            If MsgBox(Translate(2272) & vbCrLf & _
                       Translate(2231), vbYesNo + vbInformation, _
                       "Stock Card / Products") = vbYes Then
                cmdCountry_Click (1)
            Else
                'Revert to previous country code.
                txtCountryExport.Text = strCtryExport
            End If
        Else
            txtCountryExportDescription.Text = strBlah
        End If
        strBlah = Empty
        bytCtryEFound = 1
    End If
End Sub

Private Sub txtEntrepotNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdEntrepotNum_Click
    End If
End Sub

Private Sub txtEntrepotNum_LostFocus()
    'Only performs this check if user has entered something in textbox.
    If Len(txtEntrepotNum.Text) <> 0 Then
        'Verifies Entrepot validity and existence on lost focus.
        txtEntrepotNum.Tag = GetEntrepot_ID(txtEntrepotNum.Text, m_conConnection, False)
        If txtEntrepotNum.Tag = 0 Then
            If MsgBox(Translate(2233) & vbCrLf & _
                  Translate(2234), vbYesNo + vbInformation, Translate(2212)) = vbYes Then
                cmdEntrepotNum_Click
            Else
                txtEntrepotNum.Text = Empty
            End If
        End If
    End If
End Sub

Private Function Validation() As Boolean
    Validation = True
    If Len(txtEntrepotNum.Tag) = 0 Or txtEntrepotNum.Tag = "0" Then
        Validation = False
        MsgBox Translate(2235), vbOKOnly + vbInformation, Translate(2212)
        txtEntrepotNum.SetFocus
        Exit Function
    End If
    If Len(txtEntrepotNum.Text) = 0 Then
        Validation = False
        txtEntrepotNum.SetFocus
        Exit Function
    End If
    If Len(txtProductNum.Text) = 0 Then
        Validation = False
        'MsgBox "Product Number missing.", vbOKOnly + vbInformation, "Product"
        MsgBox Translate(2232), vbOKOnly + vbInformation, Translate(2212)
        txtProductNum.SetFocus
        Exit Function
    End If
    If Len(txtCountryOrigin.Text) = 0 Then
        Validation = False
        MsgBox "Country of Origin missing.", vbOKCancel + vbInformation, "Product"
        txtCountryOrigin.SetFocus
        Exit Function
    End If
    If Len(txtCountryExport.Text) = 0 Then
        Validation = False
        MsgBox "Country of Export missing.", vbOKCancel + vbInformation, "Product"
        txtCountryExport.SetFocus
        Exit Function
    End If
    
    'Added to ensure correct Country description is used when Enter key is used to trigger OK.
    txtCountryExport_LostFocus
    txtCountryOrigin_LostFocus
End Function

Private Sub txtTaricCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdTaricCode_Click
    End If
End Sub
