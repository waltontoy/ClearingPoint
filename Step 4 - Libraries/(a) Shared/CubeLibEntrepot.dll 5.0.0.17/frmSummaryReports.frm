VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSummaryReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Summary Reports"
   ClientHeight    =   2910
   ClientLeft      =   5685
   ClientTop       =   3390
   ClientWidth     =   6360
   Icon            =   "frmSummaryReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pgbCreateLinkedTables 
      Height          =   135
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.FileListBox flbFilter 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   4560
      Width           =   495
   End
   Begin VB.CheckBox chkShowZero 
      Caption         =   "Show depleted stocks"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   13
      Tag             =   "2217"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Tag             =   "670"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Pre&view"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   1
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   16
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton cmdPicklist 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   3525
         TabIndex        =   15
         Top             =   600
         Width           =   315
      End
      Begin VB.ComboBox cboTypes 
         Height          =   315
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Index           =   0
         Left            =   3960
         MaxLength       =   25
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpPeriod 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Tag             =   "2225"
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38179
      End
      Begin VB.CommandButton cmdPicklist 
         Caption         =   "..."
         Default         =   -1  'True
         Height          =   315
         Index           =   0
         Left            =   5640
         TabIndex        =   6
         Top             =   960
         Width           =   315
      End
      Begin VB.ComboBox cboTypes 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpPeriod 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Tag             =   "2226"
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38179
      End
      Begin VB.Label Label2 
         Caption         =   "Entrepot Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Period End:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Period Start:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Filter Type:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "2224"
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Report Type:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "2223"
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Label lblPreparing 
      Caption         =   "Preparing..."
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmSummaryReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_colConnections As PCubeLibDBReg.CConnections

Private m_lngUserID As Long
Private m_blnAnimationRunOnce As Boolean

Private m_blnPickSummaryButton As Boolean
Private mstrLanguage As String
Private mintTaricUse As Integer
Private mstrAppVersion As String
Private mstrLicenseeName As String
Private mstrLicCompanyName As String
Private mblnLicIsDemo As Boolean

Private mlngStockID As Long
Private mlngStockProductID As Long
Private mlngProductID As Long
Private mlngEntrepotID As Long

Private mstrStockNum As String
Private mstrProductNum As String
Private mstrEntrepotNum As String
'For the "Reduced" StockProdPicklist

Private mobjSummaryReport As DDActiveReports2.ActiveReport

Private blnWithEntrepot As Boolean
Private strPrinterName As String

Private Sub cboTypes_Click(Index As Integer)
    If m_blnAnimationRunOnce Then
        If m_blnAnimationRunOnce Then
            Label1(2).Caption = "Period Start:"
            
            Select Case Index
                Case 0
                    cboTypes(1).Clear
        '            cboTypes(1).Enabled = True
                    Label1(1).Caption = "Filter Type :"
                    chkShowZero.Visible = False
                    
                    If cboTypes(0).ListIndex = 0 Then
                        dtpPeriod(0).Enabled = True
                        
                        cboTypes(1).AddItem "Stock Card Number"
                        cboTypes(1).AddItem "Product Number"
                        cboTypes(1).AddItem "Entrepot Number"
                        
                        cboTypes(1).Enabled = True
                        
                        txtFilter(0).Visible = True
                        cmdPicklist(0).Visible = True
                        txtFilter(0).Enabled = True
                        cmdPicklist(0).Enabled = True
                        cboTypes(1).Visible = True
                        Label1(1).Visible = True
                        HideUnhideEntrepotBox False
                        
                    '== PAUL for IM7 report
                    ElseIf cboTypes(0).ListIndex = 3 Then
                        cboTypes(1).AddItem "IM7 Number"
        
        '                cboTypes(1).Enabled = False
        '                Call GetEntrepots
                        dtpPeriod(0).Value = Now
                        dtpPeriod(0).Value = Now
                        dtpPeriod(0).Enabled = True
                        dtpPeriod(1).Enabled = False
                        Label1(2).Caption = "Doc Date :"
        '                Label1(1).Caption = "Entrepot Name:"
                    '==
                        cboTypes(1).Enabled = True
                        
                        txtFilter(0).Visible = True
                        cmdPicklist(0).Visible = True
                        txtFilter(0).Enabled = True
                        cmdPicklist(0).Enabled = True
                        cboTypes(1).Visible = True
                        Label1(1).Visible = True
                        HideUnhideEntrepotBox False
                    
                    Else
                        cboTypes(1).AddItem "Entrepot Number"
                        
                        If cboTypes(0).ListIndex = 2 Then
                            dtpPeriod(0).Enabled = False
                            chkShowZero.Visible = True
                        'Glenn 3/29/2006
                        ElseIf cboTypes(0).ListIndex = 4 Then
                            dtpPeriod(0).Enabled = True
                            dtpPeriod(1).Enabled = True
                        Else
                            dtpPeriod(0).Enabled = True
                        End If
                        
                        txtFilter(0).Visible = False
                        cmdPicklist(0).Visible = False
                        txtFilter(0).Enabled = False
                        cmdPicklist(0).Enabled = False
                        cboTypes(1).Visible = False
                        Label1(1).Visible = False
                        HideUnhideEntrepotBox True
                    End If
                    
                    txtFilter(0).Text = ""
                    cboTypes(1).ListIndex = 0
                Case 1
                    txtFilter(0).Text = ""
                    
                    If cboTypes(0).ListIndex = 0 Then
                        Select Case cboTypes(1).ListIndex
                            Case 0
                                If mlngStockID <> 0 Then
                                    txtFilter(0).Text = mstrStockNum
                                End If
                                txtFilter(0).Visible = True
                                cmdPicklist(0).Visible = True
                                txtFilter(0).Enabled = True
                                cmdPicklist(0).Enabled = True
                                cboTypes(1).Visible = True
                                Label1(1).Visible = True
                                HideUnhideEntrepotBox False
                                
                            Case 1
                                If mlngProductID <> 0 Then
                                    txtFilter(0).Text = mstrProductNum
                                End If
                                txtFilter(0).Visible = True
                                cmdPicklist(0).Visible = True
                                txtFilter(0).Enabled = True
                                cmdPicklist(0).Enabled = True
                                cboTypes(1).Visible = True
                                Label1(1).Visible = True
                                HideUnhideEntrepotBox False
                                
                            Case 2
                                If mlngEntrepotID <> 0 Then
                                    txtFilter(0).Text = mstrEntrepotNum
                                End If
                                txtFilter(0).Visible = True
                                cmdPicklist(0).Visible = True
                                txtFilter(0).Enabled = False
                                cmdPicklist(0).Enabled = False
                                cboTypes(1).Visible = True
                                Label1(1).Visible = True
                                HideUnhideEntrepotBox False
                                
                        End Select
                    ElseIf cboTypes(0).ListIndex <> 3 Then
                        Select Case cboTypes(1).ListIndex
                            Case 0
                                If mlngEntrepotID <> 0 Then
                                    txtFilter(0).Text = mstrEntrepotNum
                                End If
                        End Select
                    Else
                        txtFilter(0).Text = ""
                    End If
            End Select
        End If
    End If
End Sub

Private Sub GetEntrepots()

    Dim strSQL As String
    Dim rstEntrepots As ADODB.Recordset

        strSQL = "SELECT Entrepot_Type & '-' & Entrepot_Num AS Entrepot_Name FROM Entrepots"
    ADORecordsetOpen strSQL, m_colConnections("TemplateSadbel").Connection, rstEntrepots, adOpenKeyset, adLockOptimistic
    'rstEntrepots.Open strSQL, m_colConnections("TemplateSadbel").Connection, adOpenKeyset, adLockOptimistic
    If Not (rstEntrepots.EOF And rstEntrepots.BOF) Then
        rstEntrepots.MoveFirst
        
        Do While Not rstEntrepots.EOF
            cboTypes(1).AddItem rstEntrepots!Entrepot_Name
            rstEntrepots.MoveNext
        Loop
    End If
    
    If cboTypes(1).ListCount > 0 Then
        cboTypes(1).ListIndex = 0
    End If
    
    ADORecordsetClose rstEntrepots
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Dim objStockCard As rptStockCard
    Dim objSummary71 As rptSummary71
    Dim objCarryOverStock As rptCarryOverStock
    Dim objIM7History_Report As rptIM7History
    Dim objRepackaging As rptRepackaging        'Glenn 3/29/2006
    
    Dim lngPrntCtr As Long
    
    If m_blnAnimationRunOnce Then
        Select Case Index
            Case 0  ' Preview
                If cboTypes(0).ListIndex = 0 Then
                    If mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    ElseIf cboTypes(1).ListIndex = 0 And mlngStockID = 0 Then
                        MsgBox Translate(2293), vbOKOnly + vbInformation
                    ElseIf cboTypes(1).ListIndex = 1 And mlngProductID = 0 Then
                        MsgBox Translate(2294), vbOKOnly + vbInformation
                    ElseIf cboTypes(1).ListIndex = 2 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
                        Set objStockCard = New rptStockCard
                        
                        With objStockCard
                            Select Case cboTypes(1).ListIndex
                                Case 0
                                    .FilterID = mlngStockID
                                    .FilterType = FilterStockID
                                    
                                    .documentName = "Stock Card " & mstrStockNum
                                Case 1
                                    .FilterID = mlngProductID
                                    .FilterType = FilterProductID
                                    
                                    .documentName = "Product " & mstrProductNum
                                Case 2
                                    .FilterID = mlngEntrepotID
                                    .FilterType = FilterEntrepotID
                                    
                                    .documentName = "Entrepot " & mstrEntrepotNum
                            End Select
                            
                            .PeriodFrom = dtpPeriod(0).Value
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .documentName = .documentName & " - Stock Card"
                            .Caption = .documentName & " - Stock Card"
                            .UserID = m_lngUserID
                            .LicCompanyName = mstrLicCompanyName
                            .LicIsDemo = mblnLicIsDemo
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                            
                            '.Show vbModal
                        End With
                        
                        Set mobjSummaryReport = objStockCard
                        Set objStockCard = Nothing
                        
                        Me.Hide
                    End If
                    
                ElseIf cboTypes(0).ListIndex = 1 Then
                    
                    If cboTypes(1).ListIndex = 0 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
                        Set objSummary71 = New rptSummary71
                        
                        With objSummary71
                            Select Case cboTypes(1).ListIndex
                                Case 0
                                    .FilterID = mlngEntrepotID
                                    .FilterType = FilterEntrepotID
                                    
                                    .documentName = "Entrepot " & mstrEntrepotNum
                            End Select
                            
                            .PeriodFrom = dtpPeriod(0).Value
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .documentName = .documentName & " - Summary 71"
                            .Caption = .documentName & " - Summary 71"
                            
                            .LicCompanyName = mstrLicCompanyName
                            .LicIsDemo = mblnLicIsDemo
                            .UserID = m_lngUserID
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                            
                            '.Show vbModal
                        End With
                        
                        Set mobjSummaryReport = objSummary71
                        Set objSummary71 = Nothing
                        
                        Me.Hide
                    End If
                    
                ElseIf cboTypes(0).ListIndex = 2 Then
                    
                    If cboTypes(1).ListIndex = 0 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
            
                        Set objCarryOverStock = New rptCarryOverStock
                        
                        With objCarryOverStock
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .Language = mstrLanguage
                            .MDBPath = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
                            .FilterID = mlngEntrepotID
                            .ShowZeroStocks = IIf(chkShowZero.Value = 0, False, True)
                            .LicCompanyName = mstrLicCompanyName
                            .LicIsDemo = mblnLicIsDemo
                            .UserID = m_lngUserID
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                        
                        End With
                        Set mobjSummaryReport = objCarryOverStock
                        Set objCarryOverStock = Nothing
                        Me.Hide
                    End If
                    
23              ElseIf cboTypes(0).ListIndex = 3 Then
                    
                    Set objIM7History_Report = New rptIM7History
                    
                    With objIM7History_Report
                        .mstrIndoc_Num = txtFilter(0).Text
                        .mstrEntrepot_Type_Num = txtFilter(1).Text
                        .mdatDateGiven = dtpPeriod(0).Value
                        .mstrLicCompanyName = mstrLicCompanyName
                        .mblnLicIsDemo = mblnLicIsDemo
                        .UserID = m_lngUserID
                        Set .mconSADBEL = m_colConnections("TemplateSADBEL").Connection
                    End With
                    
                    Set mobjSummaryReport = objIM7History_Report
                    Set objIM7History_Report = Nothing
                    Me.Hide
                    
                'Glenn 3/29/2006
                ElseIf cboTypes(0).ListIndex = 4 Then
                    
                    Set objRepackaging = New rptRepackaging
                    
                    With objRepackaging
                        .EntrepotID = mlngEntrepotID
                        .PeriodFrom = dtpPeriod(0).Value
                        .PeriodTo = dtpPeriod(1).Value
                        .AppVersion = mstrAppVersion
                        .Language = mstrLanguage
                        .LicCompanyName = mstrLicCompanyName
                        .LicIsDemo = mblnLicIsDemo
                        .UserID = m_lngUserID
                        Set .RepackagingConnection = m_colConnections("TemplateSADBEL").Connection
                    End With
                    
                    Set mobjSummaryReport = objRepackaging
                    Set objRepackaging = Nothing
                    Me.Hide
                    
                End If
                
            Case 1 ' Print
                If cboTypes(0).ListIndex = 0 Then
                    If cboTypes(1).ListIndex = 0 And mlngStockID = 0 Then
                        MsgBox Translate(2293), vbOKOnly + vbInformation
                    ElseIf cboTypes(1).ListIndex = 1 And mlngProductID = 0 Then
                        MsgBox Translate(2294), vbOKOnly + vbInformation
                    ElseIf cboTypes(1).ListIndex = 2 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
                        Set objStockCard = New rptStockCard
                        
                        With objStockCard
                            Select Case cboTypes(1).ListIndex
                                Case 0
                                    .FilterID = mlngStockID
                                    .FilterType = FilterStockID
                                Case 1
                                    .FilterID = mlngProductID
                                    .FilterType = FilterProductID
                                Case 2
                                    .FilterID = mlngEntrepotID
                                    .FilterType = FilterEntrepotID
                            End Select
                            
                            .PeriodFrom = dtpPeriod(0).Value
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .UserID = m_lngUserID
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                            
                            .Run False
                            
                            If .Pages.Count > 0 Then
                                On Error GoTo Printer_Error
                                For lngPrntCtr = 0 To .Printer.NDevices - 1
                                    If UCase(.Printer.Devices(lngPrntCtr)) = UCase(strPrinterName) Then
                                        .Printer.DeviceName = strPrinterName
                                        Exit For
                                    End If
                                Next
                                If ShowPrinters(strPrinterName) Then
                                    .Printer.DeviceName = strPrinterName
                                    .PrintReport False
                                End If
                                On Error GoTo 0
                            End If
                        End With
                        
                        Unload objStockCard
                        
                        Set objStockCard = Nothing
                    End If
                    
                ElseIf cboTypes(0).ListIndex = 1 Then
                    
                    If cboTypes(1).ListIndex = 0 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
                        Set objSummary71 = New rptSummary71
                        
                        With objSummary71
                            Select Case cboTypes(1).ListIndex
                                Case 0
                                    .FilterID = mlngEntrepotID
                                    .FilterType = FilterEntrepotID
                            End Select
                            
                            .PeriodFrom = dtpPeriod(0).Value
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .UserID = m_lngUserID
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                            
                            .Run False
                            
                            If .Pages.Count > 0 Then
                                On Error GoTo Printer_Error
                                For lngPrntCtr = 0 To .Printer.NDevices - 1
                                    If UCase(.Printer.Devices(lngPrntCtr)) = UCase(strPrinterName) Then
                                        .Printer.DeviceName = strPrinterName
                                        Exit For
                                    End If
                                Next
                                If ShowPrinters(strPrinterName) Then
                                    .Printer.DeviceName = strPrinterName
                                    .PrintReport False
                                End If
                                On Error GoTo 0
                            End If
                        
                        End With
                        
                        Unload objSummary71
                        
                        Set objSummary71 = Nothing
                    End If
                    
                ElseIf cboTypes(0).ListIndex = 2 Then
                    
                    If cboTypes(1).ListIndex = 0 And mlngEntrepotID = 0 Then
                        MsgBox Translate(2295), vbOKOnly + vbInformation
                    Else
            
                        Set objCarryOverStock = New rptCarryOverStock
                        
                        With objCarryOverStock
                            .PeriodTo = dtpPeriod(1).Value
                            .AppVersion = mstrAppVersion
                            .LicenseeName = mstrLicenseeName
                            .Language = mstrLanguage
                            .MDBPath = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
                            .FilterID = mlngEntrepotID
                            .UserID = m_lngUserID
                            
                            Set .Connection = m_colConnections("TemplateSADBEL").Connection
                            
                            .Run False
                            
                            If .Pages.Count > 0 Then
                                On Error GoTo Printer_Error
                                For lngPrntCtr = 0 To .Printer.NDevices - 1
                                    If UCase(.Printer.Devices(lngPrntCtr)) = UCase(strPrinterName) Then
                                        .Printer.DeviceName = strPrinterName
                                        Exit For
                                    End If
                                Next
                                If ShowPrinters(strPrinterName) Then
                                    .Printer.DeviceName = strPrinterName
                                    .PrintReport False
                                End If
                                On Error GoTo 0
                            End If
                            
                            
                        End With
                        
                        Unload objCarryOverStock
                        
                        Set objCarryOverStock = Nothing
                        
                    End If
                    
                ElseIf cboTypes(0).ListIndex = 3 Then
                    
                    Set objIM7History_Report = New rptIM7History
                    
                    With objIM7History_Report
                        .mstrIndoc_Num = txtFilter(0).Text
                        .mstrEntrepot_Type_Num = txtFilter(1).Text
                        .mdatDateGiven = dtpPeriod(0).Value
                        .UserID = m_lngUserID
                        Set .mconSADBEL = m_colConnections("TemplateSADBEL").Connection
                    End With
                    
                    objIM7History_Report.Run False
                    
                    If objIM7History_Report.Pages.Count > 0 Then
                        On Error GoTo Printer_Error
                        For lngPrntCtr = 0 To objIM7History_Report.Printer.NDevices - 1
                            If UCase(objIM7History_Report.Printer.Devices(lngPrntCtr)) = UCase(strPrinterName) Then
                                objIM7History_Report.Printer.DeviceName = strPrinterName
                                Exit For
                            End If
                        Next
                        If ShowPrinters(strPrinterName) Then
                            objIM7History_Report.Printer.DeviceName = strPrinterName
                            objIM7History_Report.PrintReport False
                        End If
                        On Error GoTo 0
                    End If
                    
                    Unload objIM7History_Report
                    Set objIM7History_Report = Nothing
                    
                'Glenn 3/29/2006
                ElseIf cboTypes(0).ListIndex = 4 Then
                    
                    Set objRepackaging = New rptRepackaging
                    
                    With objRepackaging
                        .EntrepotID = mlngEntrepotID
                        .PeriodFrom = dtpPeriod(0).Value
                        .PeriodTo = dtpPeriod(1).Value
                        .AppVersion = mstrAppVersion
                        .Language = mstrLanguage
                        .UserID = m_lngUserID
                        Set .RepackagingConnection = m_colConnections("TemplateSADBEL").Connection
                    End With
                    
                    objRepackaging.Run False
                    
                    If objRepackaging.Pages.Count > 0 Then
                        On Error GoTo Printer_Error
                        For lngPrntCtr = 0 To objRepackaging.Printer.NDevices - 1
                            If UCase(objRepackaging.Printer.Devices(lngPrntCtr)) = UCase(strPrinterName) Then
                                objRepackaging.Printer.DeviceName = strPrinterName
                                Exit For
                            End If
                        Next
                        If ShowPrinters(strPrinterName) Then
                            objRepackaging.Printer.DeviceName = strPrinterName
                            objRepackaging.PrintReport False
                        End If
                        On Error GoTo 0
                    End If
                    
                    Unload objRepackaging
                    Set objRepackaging = Nothing
                    
                End If
                
            Case 2 ' Cancel
                Unload Me
                
        End Select
        
        Exit Sub
        
Printer_Error:
        If Err.Number = 482 Or Err.Number = 5707 Or Err.Number = -2147417848 Then
            MsgBox "An error occurred while connecting to your printer.", vbExclamation + vbOKOnly, "ClearingPoint"
            Err.Clear
            Resume Next
        Else
            MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "ClearingPoint"
            Resume Next
        End If
    End If
End Sub

Private Sub CreateLinkedTablesForStockcardReport(ByRef ADOSadbel As ADODB.Connection)
    Dim intYearCtr As Integer
    
    Dim strHistoryDBPath As String
    Dim strHistoryDBYear As String
    
    'Dim strSADBELDBPath As String
    
    Dim astrHistoryDBs() As String
    
    
    'strSADBELDBPath = Replace(ADOSadbel.Properties("Data Source Name").Value, "mdb_sadbel.mdb", vbNullString)
    strHistoryDBPath = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & "mdb_history??.mdb")
    ReDim Preserve astrHistoryDBs(0)
    
    Do Until Len(Trim(strHistoryDBPath)) = 0
        ReDim Preserve astrHistoryDBs(UBound(astrHistoryDBs) + 1)
        astrHistoryDBs(UBound(astrHistoryDBs)) = strHistoryDBPath
        
        strHistoryDBPath = Dir()
    Loop
    
    
    'strSADBELDBPath = ADOSadbel.Properties("Data Source Name").Value
    
    For intYearCtr = 1 To UBound(astrHistoryDBs)
        strHistoryDBYear = Replace(astrHistoryDBs(intYearCtr), "mdb_history", vbNullString)
        strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
        
        'strHistoryDBPath = Replace(strSADBELDBPath, "mdb_sadbel.mdb", vbNullString) & astrHistoryDBs(intYearCtr)
        
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "HistoryInbounds" & strHistoryDBYear, DBInstanceType_DATABASE_HISTORY, "Inbounds", , GetHistoryDBYear(astrHistoryDBs(intYearCtr))
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "HistoryInboundDocs" & strHistoryDBYear, DBInstanceType_DATABASE_HISTORY, "InboundDocs", , GetHistoryDBYear(astrHistoryDBs(intYearCtr))
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "HistoryOutbounds" & strHistoryDBYear, DBInstanceType_DATABASE_HISTORY, "Outbounds", , GetHistoryDBYear(astrHistoryDBs(intYearCtr))
        CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, "HistoryOutboundDocs" & strHistoryDBYear, DBInstanceType_DATABASE_HISTORY, "OutboundDocs", , GetHistoryDBYear(astrHistoryDBs(intYearCtr))
        
        'AddLinkedTableEx "HistoryInbounds" & strHistoryDBYear, strSADBELDBPath, G_Main_Password, "Inbounds", strHistoryDBPath, G_Main_Password
        'AddLinkedTableEx "HistoryInboundDocs" & strHistoryDBYear, strSADBELDBPath, G_Main_Password, "InboundDocs", strHistoryDBPath, G_Main_Password
        'AddLinkedTableEx "HistoryOutbounds" & strHistoryDBYear, strSADBELDBPath, G_Main_Password, "Outbounds", strHistoryDBPath, G_Main_Password
        'AddLinkedTableEx "HistoryOutboundDocs" & strHistoryDBYear, strSADBELDBPath, G_Main_Password, "OutboundDocs", strHistoryDBPath, G_Main_Password
        
        Me.Refresh
        DoEvents
    Next intYearCtr
    
    Erase astrHistoryDBs
    
End Sub

Private Sub DeleteLinkedTablesForStockcardReport(ByRef ADOSadbel As ADODB.Connection)
    Dim intYearCtr As Integer
    
    Dim strHistoryDBPath As String
    Dim strHistoryDBYear As String
    
    'Dim strSADBELDBPath As String
    
    Dim astrHistoryDBs() As String
    
    
    'strSADBELDBPath = Replace(ADOSadbel.Properties("Data Source Name").Value, "mdb_sadbel.mdb", vbNullString)
    strHistoryDBPath = Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & "mdb_history??.mdb")
    
    ReDim Preserve astrHistoryDBs(0)
    
    Do Until Len(Trim(strHistoryDBPath)) = 0
        ReDim Preserve astrHistoryDBs(UBound(astrHistoryDBs) + 1)
        astrHistoryDBs(UBound(astrHistoryDBs)) = strHistoryDBPath
        
        strHistoryDBPath = Dir()
    Loop

    For intYearCtr = 1 To UBound(astrHistoryDBs)
        strHistoryDBYear = Replace(astrHistoryDBs(intYearCtr), "mdb_history", vbNullString)
        strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
        
        strHistoryDBPath = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\" & astrHistoryDBs(intYearCtr)
        
        On Error Resume Next
        ExecuteNonQuery ADOSadbel, "DROP TABLE HistoryInbounds" & strHistoryDBYear
        ExecuteNonQuery ADOSadbel, "DROP TABLE HistoryInboundDocs" & strHistoryDBYear
        ExecuteNonQuery ADOSadbel, "DROP TABLE HistoryOutbounds" & strHistoryDBYear
        ExecuteNonQuery ADOSadbel, "DROP TABLE HistoryOutboundDocs" & strHistoryDBYear
        On Error GoTo 0
    Next intYearCtr
    
    Erase astrHistoryDBs
    
End Sub

Private Sub cmdPicklist_Click(Index As Integer)
    
    Dim clsStockProd As cStockProd
    Dim clsProducts As cProducts
    Dim clsEntrepots As cEntrepot
    
    If m_blnAnimationRunOnce Then
        If Index = 0 Then
            If cboTypes(0).ListIndex = 0 Then
                Select Case cboTypes(1).ListIndex
                    Case 0    ' Load PRODUCT NUMBER - STOCK CARD NUMBER filter picklist
                        Set clsStockProd = New cStockProd
                        
                        ' Used to automatically select previously selected item
                        If mlngStockID <> 0 Then
                            clsStockProd.Stock_ID = mlngStockID
                        End If
                        
                        If mlngStockProductID <> 0 Then
                            clsStockProd.Product_ID = mlngStockProductID
                        End If
                        
                        If Len(txtFilter(0).Text) > 0 Then
                            clsStockProd.StockCardNo = txtFilter(0).Text
                        End If
                        
                        If blnWithEntrepot = True Then
                            clsStockProd.Entrepot_Num = mstrEntrepotNum
                        End If
                        
                        clsStockProd.ShowPicklist Me, m_colConnections("TemplateSADBEL").Connection, m_colConnections("TemplateTARIC").Connection, mstrLanguage, mintTaricUse, ResourceHandler, True, , True, mstrEntrepotNum, True
                        
                        If Len(Trim(clsStockProd.Stock_ID)) Then
                            mlngStockID = clsStockProd.Stock_ID
                            mlngStockProductID = clsStockProd.Product_ID
                            
                            If clsStockProd.Cancel = False Then
                                mstrStockNum = clsStockProd.StockCardNo
                                mstrEntrepotNum = clsStockProd.Entrepot_Num
                                txtFilter(0).Text = clsStockProd.StockCardNo
                                'Determines whether there is an entrepot number saved from frmStockProdPicklist,
                                blnWithEntrepot = True
                            End If
                        Else
                            mlngStockID = 0
                            mlngStockProductID = 0
                        End If
                        
                        Set clsStockProd = Nothing
                    Case 1    ' Load PRODUCT NUMBER filter picklist
                        Set clsProducts = New cProducts
                        
                        If mlngProductID <> 0 Then
                            clsProducts.Product_ID = mlngProductID
                        End If
                        
                        If Len(Trim(txtFilter(1).Text)) > 0 Then
                            clsProducts.Entrepot_Num = Trim(txtFilter(1).Text)
                        End If
                        
                        clsProducts.ShowProducts 1, Me, m_colConnections("TemplateSADBEL").Connection, m_colConnections("TemplateTARIC").Connection, mstrLanguage, mintTaricUse, ResourceHandler, mstrProductNum, True, True
                        
                        If Len(Trim(clsProducts.Product_ID)) Then
                            mlngProductID = clsProducts.Product_ID
                            
                            If clsProducts.Cancelled = False Then
                                mstrProductNum = clsProducts.Product_Num
                                mstrEntrepotNum = clsProducts.Entrepot_Num
                                txtFilter(0).Text = clsProducts.Product_Num
                            End If
                        Else
                            mlngProductID = 0
                        End If
                        
                        Set clsProducts = Nothing
                    Case 2    ' Load ENTREPOT NUMBER filter picklist
    
                End Select
            ElseIf cboTypes(0).ListIndex = 3 Then
                If Not m_blnPickSummaryButton Then
                    m_blnPickSummaryButton = True
                    Call LoadIM7s
                    m_blnPickSummaryButton = False
                End If
            Else
                Select Case cboTypes(1).ListIndex
                    Case 0    ' Load ENTREPOT NUMBER filter picklist
                        Set clsEntrepots = New cEntrepot
                        
                        If mlngEntrepotID <> 0 Then
                            clsEntrepots.Entrepot_ID = mlngEntrepotID
                        End If
                        
                        clsEntrepots.ShowEntrepot Me, m_colConnections("TemplateSADBEL").Connection, True, mstrLanguage, ResourceHandler, , mlngEntrepotID, True
                        
                        If Len(Trim(clsEntrepots.Entrepot_ID)) Then
                            mlngEntrepotID = clsEntrepots.Entrepot_ID
                            
                            If clsEntrepots.Cancelled = False Then
                                mstrEntrepotNum = clsEntrepots.SelectedEntrepot
                                mlngEntrepotID = clsEntrepots.Entrepot_ID
                                txtFilter(0).Text = clsEntrepots.SelectedEntrepot
                                Call txtFilter_LostFocus(0)
                            End If
                        Else
                            mlngEntrepotID = 0
                        End If
                        
                        Set clsEntrepots = Nothing
                End Select
            End If
        Else
            Set clsEntrepots = New cEntrepot
            
            If mlngEntrepotID <> 0 Then
                clsEntrepots.Entrepot_ID = mlngEntrepotID
            End If
            
            clsEntrepots.ShowEntrepot Me, m_colConnections("TemplateSADBEL").Connection, True, mstrLanguage, ResourceHandler, , mlngEntrepotID, True
            
            If Len(Trim(clsEntrepots.Entrepot_ID)) Then
                mlngEntrepotID = clsEntrepots.Entrepot_ID
                
                If clsEntrepots.Cancelled = False Then
                    mstrEntrepotNum = clsEntrepots.SelectedEntrepot
                    mlngEntrepotID = clsEntrepots.Entrepot_ID
                    txtFilter(Index).Text = clsEntrepots.SelectedEntrepot
                    Call txtFilter_LostFocus(Index)
                    If cboTypes(1).Text = "Entrepot Number" Then
                        txtFilter(Index - 1).Text = clsEntrepots.SelectedEntrepot
                    End If
                End If
            Else
                mlngEntrepotID = 0
            End If
            
            Set clsEntrepots = Nothing
            
        End If
    End If
End Sub

Private Sub CreateTemporaryTablesForIM7Report()
    Dim strCommand As String
    Dim strEntrepot As String
    Dim strHistoryDBName As String
    Dim strHistoryDBYear As String
    Dim strHistoryInboundsName As String
    Dim strHistoryInboundDocsName As String
    
    Dim lngHistoryDBCtr As Long
    Dim lngNumberOfTries As Long
    
    Dim sngInterval As Single
    Dim sngEndTime As Single
                
    Dim astrHistoryInboundDocsName() As String
    Dim astrHistoryInboundsName() As String
    
    Dim blnCreationOfLinkedTablesSuccessul As Boolean
    
    Dim conSADBELDB As ADODB.Connection
    Dim rstEntrepots As ADODB.Recordset
                    
    ' Create Linked Tables for the IM7 Document Number Picklist
    flbFilter.Path = NoBackSlash(g_objDataSourceProperties.InitialCatalogPath)
    flbFilter.Pattern = "mdb_history??.mdb"
    flbFilter.Refresh

    ReDim astrHistoryInboundDocsName(1 To flbFilter.ListCount)
    ReDim astrHistoryInboundsName(1 To flbFilter.ListCount)
    
    Dim lngProgressIncrement As Long
    
    lngProgressIncrement = (pgbCreateLinkedTables.Max - 2) / flbFilter.ListCount

    For lngHistoryDBCtr = 1 To flbFilter.ListCount
        
        strHistoryDBName = flbFilter.List(lngHistoryDBCtr - 1)
        strHistoryDBName = Trim$(strHistoryDBName)
        
        ' Remove file extension .mdb
        strHistoryDBYear = Left$(strHistoryDBName, Len(strHistoryDBName) - 4)
        
        ' Get the Year Part of the DB Name
        If IsNumeric(Right$(strHistoryDBYear, 2)) Then
            strHistoryDBYear = Right$(strHistoryDBYear, 2)
        Else
            strHistoryDBYear = vbNullString
        End If
        
        strHistoryDBYear = Trim$(strHistoryDBYear)
        
        strHistoryInboundDocsName = "HistoryInboundDocsForIM7Report" & strHistoryDBYear
        strHistoryInboundsName = "HistoryInboundsForIM7Report" & strHistoryDBYear
        
        blnCreationOfLinkedTablesSuccessul = False
        Do Until blnCreationOfLinkedTablesSuccessul
            ' Create a Link Table for each History Database
            CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, strHistoryInboundDocsName, DBInstanceType_DATABASE_HISTORY, "InboundDocs", , GetHistoryDBYear(strHistoryDBName)
            CreateLinkedTable g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL, strHistoryInboundsName, DBInstanceType_DATABASE_HISTORY, "Inbounds", , GetHistoryDBYear(strHistoryDBName)
            
            'AddLinkedTableEx strHistoryInboundDocsName, NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "InboundDocs", NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\" & strHistoryDBName, G_Main_Password
            'AddLinkedTableEx strHistoryInboundsName, NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb", G_Main_Password, "Inbounds", NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\" & strHistoryDBName, G_Main_Password
            
            ' Store History Linked Table Names
            astrHistoryInboundDocsName(lngHistoryDBCtr) = strHistoryInboundDocsName
            astrHistoryInboundsName(lngHistoryDBCtr) = strHistoryInboundsName
            
            ' Ensure that the Linked Tables are Existing Before Creating
            ' the succeeding Linked Tables
            lngNumberOfTries = 0
            Do While (lngNumberOfTries < 10)
            
                DoEvents
                
                ADOConnectDB conSADBELDB, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
                
                blnCreationOfLinkedTablesSuccessul = (ADOXIsTableExisting(conSADBELDB, strHistoryInboundDocsName) And _
                                                        ADOXIsTableExisting(conSADBELDB, strHistoryInboundsName))
                
                ADODisconnectDB conSADBELDB
                
                If blnCreationOfLinkedTablesSuccessul Then
                    Exit Do
                End If
                
                ' By delaying a short random interval in the range of 0.25 to 1 second before
                ' retrying the operation, the chances of a deadlock are minimized.
                Randomize
                sngInterval = Rnd * 0.75 + 0.25
                sngEndTime = Timer + sngInterval
                
                Do
                    ' Random delay loop
                Loop Until Timer >= sngEndTime
                
                lngNumberOfTries = lngNumberOfTries + 1
            Loop
        Loop
        
        If pgbCreateLinkedTables.Value + lngProgressIncrement > pgbCreateLinkedTables.Max Then
            pgbCreateLinkedTables.Value = pgbCreateLinkedTables.Max
        Else
            pgbCreateLinkedTables.Value = pgbCreateLinkedTables.Value + lngProgressIncrement
        End If
        Me.Refresh
    Next
    
    ' Create Temporary Table for the IM7 Document Number Picklist
    ADOConnectDB conSADBELDB, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    'OpenADODatabase conSADBELDB, NoBackSlash(g_objDataSourceProperties.TracefilePath), "mdb_sadbel.mdb"
                    
        ' Drop Existing Table IM7EntrepotDocumentNumbers and Recreate to Refresh the Data
    On Error Resume Next
        strCommand = vbNullString
        strCommand = strCommand & "DROP TABLE IM7EntrepotDocumentNumbers "
    ExecuteNonQuery conSADBELDB, strCommand
    On Error GoTo 0
    
        ' Create Temporary Table for all Unique Entrepot+DocumentNumber Combinations
        strCommand = vbNullString
        strCommand = strCommand & "SELECT DISTINCT "
        strCommand = strCommand & "Entrepot_Type, "
        strCommand = strCommand & "Entrepot_Num, "
        strCommand = strCommand & "InDoc_Num "
        strCommand = strCommand & "INTO "
        strCommand = strCommand & "IM7EntrepotDocumentNumbers "
        strCommand = strCommand & "FROM "
        
        strCommand = strCommand & "( "
        
        ' Create UNION SQL Command of all HistoryInboundDocs and HistoryInbounds
        For lngHistoryDBCtr = 1 To flbFilter.ListCount
            strHistoryInboundDocsName = astrHistoryInboundDocsName(lngHistoryDBCtr)
            strHistoryInboundsName = astrHistoryInboundsName(lngHistoryDBCtr)
            
            strCommand = strCommand & "" & GetSQLCommandForIM7HistoryInboundLinkedTables(strHistoryInboundDocsName, strHistoryInboundsName) & " "
            
            If lngHistoryDBCtr <> flbFilter.ListCount Then
                ' Use UNION to remove duplicate records
                strCommand = strCommand & " " & "UNION" & " "
            End If
        Next
        strCommand = strCommand & ") "
    
    ExecuteNonQuery conSADBELDB, strCommand
    
    ADODisconnectDB conSADBELDB
End Sub

Private Function GetSQLCommandForIM7HistoryInboundLinkedTables(ByVal HistoryInboundDocsName As String, _
                                                                ByVal HistoryInboundsName As String)
    Dim strCommand As String
    
    strCommand = vbNullString
    strCommand = strCommand & "SELECT DISTINCT "
    strCommand = strCommand & "Entrepot_Type, "
    strCommand = strCommand & "Entrepot_Num, "
    strCommand = strCommand & "InDoc_Num "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "Entrepots "
    strCommand = strCommand & "INNER JOIN ("
        strCommand = strCommand & "Products "
        strCommand = strCommand & "INNER JOIN ("
            strCommand = strCommand & "StockCards "
            strCommand = strCommand & "INNER JOIN ("
            strCommand = strCommand & "" & HistoryInboundDocsName & " "
                strCommand = strCommand & "INNER JOIN "
                strCommand = strCommand & "" & HistoryInboundsName & " "
                strCommand = strCommand & "ON "
                strCommand = strCommand & "" & HistoryInboundDocsName & ".Indoc_ID = " & HistoryInboundsName & ".InDoc_ID) "
            strCommand = strCommand & "ON "
            strCommand = strCommand & "" & HistoryInboundsName & ".Stock_ID = StockCards.Stock_ID) "
        strCommand = strCommand & "ON "
        strCommand = strCommand & "Products.Prod_ID = StockCards.Prod_ID) "
    strCommand = strCommand & "ON "
    strCommand = strCommand & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "NOT ISNULL(" & HistoryInboundDocsName & ".InDoc_Num) "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "TRIM(" & HistoryInboundDocsName & ".InDoc_Num) <> '' "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "IIF(ISNULL(" & HistoryInboundsName & ".In_Code), '', " & HistoryInboundsName & ".In_Code) <> '<<Closure>>' "
    strCommand = strCommand & "AND "
    strCommand = strCommand & "IIF(ISNULL(" & HistoryInboundsName & ".In_Code), '', " & HistoryInboundsName & ".In_Code) NOT LIKE '*<<TEST>>' "
    
    GetSQLCommandForIM7HistoryInboundLinkedTables = strCommand
End Function

Private Sub LoadIM7s()
    Dim clsPicklist As CPicklist
    Dim clsGrid As CGridSeed
    Dim strDB As String
    Dim strCommand As String
    
    '<<< dandan 112706
    '<<< Update with database password
    ADODisconnectDB m_colConnections("TemplateSadbel").Connection
    ADOConnectDB m_colConnections("TemplateSadbel").Connection, g_objDataSourceProperties, DBInstanceType_DATABASE_SADBEL
    'm_colConnections("TemplateSadbel").Connection.Close
    'm_colConnections("TemplateSadbel").Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & NoBackSlash(g_objDataSourceProperties.TracefilePath) & "\mdb_sadbel.mdb" & _
                ";Persist Security Info=False;Jet OLEDB:Database Password=" & G_Main_Password
    
            
    Set clsPicklist = New CPicklist
    Set clsGrid = clsPicklist.SeedGrid("Doc Type", 1300, "Left", "Document Number", 2715, "Left")
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT DISTINCT "
        strCommand = strCommand & "InDoc_Num, "
        strCommand = strCommand & "'IM7' AS [Doc Type], "
        strCommand = strCommand & "InDoc_Num AS [Document Number] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "IM7EntrepotDocumentNumbers "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "IM7EntrepotDocumentNumbers.[Entrepot_Type] & '-' & IM7EntrepotDocumentNumbers.[Entrepot_Num] = '" & txtFilter(1).Text & "' "
    
    clsPicklist.Search True, "Document Number", txtFilter(0).Text
    
    clsPicklist.Pick Me, cpiSimplePicklist, m_colConnections("TemplateSadbel").Connection, strCommand, "Document Number", "IM7", vbModal, clsGrid, , , , cpiKeyF2
    
    If Not clsPicklist.SelectedRecord Is Nothing Then
        txtFilter(0) = clsPicklist.SelectedRecord.RecordSource.Fields("Document Number").Value
    End If
    
    Set clsGrid = Nothing
    Set clsPicklist = Nothing
End Sub

Private Sub dtpPeriod_Change(Index As Integer)
    If m_blnAnimationRunOnce Then
        'This will prevent the dates starting later than end date or ending before the start date.
        If dtpPeriod(0).Value > dtpPeriod(1).Value And cboTypes(0).ListIndex <> 2 Then
            Select Case Index
                Case 0
                    dtpPeriod(0).Value = dtpPeriod(1).Value
                Case 1
                    dtpPeriod(1).Value = dtpPeriod(0).Value
            End Select
        End If
    End If
End Sub

Private Sub Form_Activate()
    txtFilter(1).Enabled = IsEntrepotTextboxEnabled(m_colConnections("TemplateSADBEL").Connection)
    cmdPicklist(1).Enabled = txtFilter(1).Enabled
    txtFilter(1).Text = mstrEntrepotNum
    
    If Not m_blnAnimationRunOnce Then
        pgbCreateLinkedTables.Visible = True
        lblPreparing.Visible = True
        Me.Refresh
                
        pgbCreateLinkedTables.Max = 12
        pgbCreateLinkedTables.Value = 0

            CreateLinkedTablesForStockcardReport m_colConnections("TemplateSADBEL").Connection
            
        pgbCreateLinkedTables.Value = 2
        Me.Refresh
        
            CreateTemporaryTablesForIM7Report
        
        pgbCreateLinkedTables.Visible = False
        lblPreparing.Visible = False
        
        Frame1.Enabled = True
        cmdAction(0).Enabled = True
        cmdAction(1).Enabled = True
        cmdAction(2).Enabled = True
        cmdPicklist(0).Enabled = True
        cmdPicklist(1).Enabled = True
        
        Me.Refresh
        
        m_blnAnimationRunOnce = True
        
        cboTypes_Click (cboTypes(0).ListIndex)
    End If
End Sub

Private Sub Form_Load()
    m_blnAnimationRunOnce = False
    
    ' Report Types
    With cboTypes(0)
        .AddItem "Stock Card"
        .AddItem "Summary 71"
        
        .AddItem "Summary of Carryover Stock"
        .AddItem "IM7 History"
        
        .AddItem "Repackaging"      'Glenn 3/29/2006
        .ListIndex = 0
    End With
    
    m_blnPickSummaryButton = False
    blnWithEntrepot = False
    
    dtpPeriod(0).Value = DateAdd("m", -1, Date)    ' Set to one month prior to current date
    dtpPeriod(1).Value = Date                      ' Set to current date
End Sub

Friend Sub My_Load(ByRef SummaryReport As DDActiveReports2.ActiveReport, _
                   ByVal Connections As PCubeLibDBReg.CConnections, _
                   ByVal Language As String, _
                   ByVal TaricUse As Integer, _
                   ByVal AppVersion As String, _
                   ByVal LicenseeName As String, _
                   ByVal MyResourceHandler As Long, _
                   ByVal strMDBpath As String, _
                   ByVal PrinterName As String, _
                   ByVal LicCompanyName As String, _
                   ByVal LicIsDemo As Boolean, _
          Optional ByVal UserID As String)
    
    'mstrMDBPath = strMDBpath
    ResourceHandler = MyResourceHandler
    mstrLanguage = Language
    m_lngUserID = UserID
    
    modGlobals.LoadResStrings Me, True
    
    Set m_colConnections = Connections
    
    mintTaricUse = TaricUse
    mstrAppVersion = AppVersion
    mstrLicenseeName = LicenseeName
    strPrinterName = PrinterName
    mstrLicCompanyName = LicCompanyName
    mblnLicIsDemo = LicIsDemo
    
    Me.Show vbModal
    
    Set SummaryReport = mobjSummaryReport
    Set mobjSummaryReport = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '<<< dandan 110807
    'added checking to prevent deletion of currently created table links that causes error
    If Not m_blnAnimationRunOnce Then
        Cancel = True
    Else
        DeleteLinkedTablesForStockcardReport m_colConnections("TemplateSADBEL").Connection
        
        Set m_colConnections = Nothing
        
        Set frmSummaryReports = Nothing
    End If

End Sub

Private Sub txtFilter_Change(Index As Integer)
    Dim rstEntrepot As ADODB.Recordset
    Dim strEntrepotDate As String
                
    If m_blnAnimationRunOnce Then
        If Index = 0 Then
            If cboTypes(0).ListIndex <> 3 Then
                If txtFilter(Index).Text <> "" And mstrEntrepotNum <> "" Then
                    strEntrepotDate = GetEntrepotLatestReOpenDate
                    
                    'Glenn 3/29/2006
                    If strEntrepotDate = "" Then
                                            
                        ADORecordsetOpen "SELECT Entrepot_StartDate AS StartDate" & _
                                         " FROM Entrepots" & _
                                         " WHERE Entrepot_Type & '-' & Entrepot_Num = '" & mstrEntrepotNum & "'", _
                                         m_colConnections("TemplateSADBEL").Connection, rstEntrepot, adOpenKeyset, adLockOptimistic
                                         
                        'rstEntrepot.Open "SELECT Entrepot_StartDate AS StartDate" & _
                                         " FROM Entrepots" & _
                                         " WHERE Entrepot_Type & '-' & Entrepot_Num = '" & mstrEntrepotNum & "'", _
                                         m_colConnections("TemplateSADBEL").Connection, adOpenForwardOnly, adLockReadOnly
                        
                        If Not (rstEntrepot.EOF Or rstEntrepot.BOF) Then
                            rstEntrepot.MoveFirst
                            
                            dtpPeriod(0).Value = DateValue(rstEntrepot!StartDate)
                        End If
                        
                        ADORecordsetClose rstEntrepot

                    Else
                        dtpPeriod(0).Value = DateValue(strEntrepotDate)
                    End If
                 Else
                    dtpPeriod(0).Value = DateAdd("m", -1, Date)
                End If
            End If
        ElseIf Index = 1 Then
            If Len(txtFilter(1).Text) = 0 Then
                mlngEntrepotID = 0
                mstrEntrepotNum = ""
            End If
            
            If cboTypes(0).ListIndex = 0 Then
                If Len(txtFilter(0).Text) = 0 Then
                    Select Case cboTypes(1).ListIndex
                        Case 0
                            mstrStockNum = ""
                            mlngStockID = 0
                        Case 1
                            mstrProductNum = ""
                            mlngProductID = 0
                    End Select
                End If
            End If
        End If
    End If
End Sub

Private Sub txtFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If m_blnAnimationRunOnce Then
        If KeyCode = vbKeyF2 Then
            cmdPicklist_Click (Index)
        End If
    End If
End Sub

Private Sub txtFilter_KeyPress(Index As Integer, KeyAscii As Integer)
    If m_blnAnimationRunOnce Then
        If cboTypes(0).ListIndex = 3 And KeyAscii <> vbKeyBack Then
            '<<< dandan 101707
            'Corrected the number of characters you can type to the the filter of MRN/IM7
            If (Len(txtFilter(Index).Text) >= 19) And (txtFilter(Index).SelLength = 0) Then '7
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtFilter_LostFocus(Index As Integer)
    If m_blnAnimationRunOnce Then
        'No input value = No processing
        If Len(Trim$(txtFilter(Index).Text)) > 0 Then
            Dim strSQL As String
            Dim strWhere As String
            Dim rst As ADODB.Recordset
            'Glenn 3/29/2006
            Dim strEntrepotDate As String
            Dim strStartDate As String

                
                'Check first if the entrepot entered in the textbox is existing in the database.
                strStartDate = ""
                If Index = 1 Then
                    ADORecordsetOpen "SELECT * FROM ENTREPOTS WHERE ENTREPOTS.Entrepot_Type & '-' & ENTREPOTS.Entrepot_Num = '" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, rst, adOpenKeyset, adLockOptimistic
                    'rst.Open "SELECT * FROM ENTREPOTS WHERE ENTREPOTS.Entrepot_Type & '-' & ENTREPOTS.Entrepot_Num = '" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, adOpenKeyset, adLockOptimistic
                    
                    If rst.BOF And rst.EOF Then
                        mlngEntrepotID = 0
                        mstrEntrepotNum = ""
                        Call cmdPicklist_Click(Index)
                    Else
                        rst.MoveFirst
                        
                        mlngEntrepotID = rst.Fields("Entrepot_ID").Value
                        mstrEntrepotNum = rst.Fields("Entrepot_Type").Value & "-" & rst.Fields("Entrepot_Num").Value
                    End If
                    
                    ADORecordsetClose rst
                End If
                
                Select Case cboTypes(0).ListIndex
                    'Stock Card
                    Case 0
                        Select Case cboTypes(1).ListIndex
                            'Stock Card number
                            Case 0
                                strSQL = "SELECT E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                                         "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StartDate AS [Entrepot Date], " & _
                                         "P.Prod_ID AS [Product ID], P.Prod_Num AS [Product Num], " & _
                                         "S.Stock_ID AS [Stock ID], S.Stock_Card_Num AS [Stock Num] " & _
                                         "FROM Entrepots [E] INNER JOIN " & _
                                         "(Products [P] INNER JOIN StockCards [S] ON P.Prod_ID = S.Prod_ID) " & _
                                         "ON E.Entrepot_ID = P.Entrepot_ID "
                                strWhere = "WHERE S.Stock_Card_Num"
                                'Open keyset in order to perform Recordcount
                                
                                ADORecordsetOpen strSQL & strWhere & " = '" & txtFilter(0).Text & "'  AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, rst, adOpenKeyset, adLockOptimistic
                                'rst.Open strSQL & strWhere & " = '" & txtFilter(0).Text & "'  AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, adOpenKeyset, adLockReadOnly
                                If Not (rst.BOF And rst.EOF) Then
                                    rst.MoveFirst
                                    
                                    If rst.RecordCount = 1 Then
                                        mlngStockID = rst.Fields("Stock ID").Value
                                        mlngStockProductID = rst.Fields("Product ID").Value
                                        mstrStockNum = rst.Fields("Stock Num").Value
    '                                    mstrEntrepotNum = rst.Fields("Entrepot Type").Value & "-" & rst.Fields("Entrepot Num").Value
                                        txtFilter(0).Text = rst.Fields("Stock Num").Value
    
                                        'Determines whether there is an entrepot number saved from frmStockProdPicklist,
                                        blnWithEntrepot = True
                                        strStartDate = DateValue(rst.Fields("Entrepot Date").Value)
                                    ElseIf rst.RecordCount > 1 Then
                                        Call cmdPicklist_Click(Index)
                                    End If
                                Else
                                    mlngStockID = 0
                                    mlngStockProductID = 0
                                End If
                            'Product number
                            Case 1
                                strSQL = "SELECT E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                                         "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StartDate AS [Entrepot Date], " & _
                                         "P.Prod_ID AS [Product ID], P.Prod_Num AS [Product Num] " & _
                                         "FROM Entrepots [E] INNER JOIN Products [P] " & _
                                         "ON E.Entrepot_ID = P.Entrepot_ID "
                                strWhere = "WHERE P.Prod_Num"
                                
                                'Open keyset in order to perform Recordcount
                                ADORecordsetOpen strSQL & strWhere & " = '" & txtFilter(0).Text & "' AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, rst, adOpenKeyset, adLockOptimistic
                                'rst.Open strSQL & strWhere & " = '" & txtFilter(0).Text & "' AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, adOpenKeyset, adLockReadOnly
                                If Not (rst.BOF And rst.EOF) Then
                                    rst.MoveFirst
                                    If rst.RecordCount = 1 Then
                                        mlngProductID = rst.Fields("Product ID").Value
    '                                    mstrEntrepotNum = rst.Fields("Entrepot Type").Value & "-" & rst.Fields("Entrepot Num").Value
                                        mstrProductNum = rst.Fields("Product Num").Value
                                        strStartDate = DateValue(rst.Fields("Entrepot Date").Value)
                                    ElseIf rst.RecordCount > 1 Then
                                        Call cmdPicklist_Click(Index)
                                    End If
                                Else
                                    mlngProductID = 0
                                End If
                            'Entrepot number
                            Case 2
                                strSQL = "SELECT E.Entrepot_ID AS [Entrepot ID], E.Entrepot_Type AS [Entrepot Type], " & _
                                         "E.Entrepot_Num AS [Entrepot Num], E.Entrepot_StartDate AS [Entrepot Date] " & _
                                         "FROM Entrepots [E] "
                                strWhere = "WHERE E.Entrepot_Type & '-' & E.Entrepot_Num"
                                ADORecordsetOpen strSQL & strWhere & " = '" & txtFilter(1).Text & "' AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, rst, adOpenKeyset, adLockOptimistic
                                'rst.Open strSQL & strWhere & " = '" & txtFilter(1).Text & "' AND E.Entrepot_Type & '-' & E.Entrepot_Num ='" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, adOpenForwardOnly, adLockReadOnly
                                If Not (rst.BOF And rst.EOF) Then
                                    rst.MoveFirst
                                    
                                    mlngEntrepotID = rst.Fields("Entrepot ID").Value
                                    mstrEntrepotNum = rst.Fields("Entrepot Type").Value & "-" & rst.Fields("Entrepot Num").Value
                                    strStartDate = DateValue(rst.Fields("Entrepot Date").Value)
                                    txtFilter(0).Text = mstrEntrepotNum
                                Else
                                    mlngEntrepotID = 0
                                End If
                        End Select
                    'Summary 71 and Summary of Carryover Stock (Same as Stock Card > Entrepot number)
                    Case 1, 2
                        strSQL = "SELECT Entrepot_ID AS [Entrepot ID], Entrepot_Type AS [Entrepot Type], Entrepot_Num AS [Entrepot Num], Entrepot_StartDate AS [Entrepot Date] FROM Entrepots "
                        strWhere = "WHERE Entrepot_Type & '-' & Entrepot_Num"
                        ADORecordsetOpen strSQL & strWhere & " = '" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, rst, adOpenKeyset, adLockOptimistic
                        'rst.Open strSQL & strWhere & " = '" & txtFilter(1).Text & "'", m_colConnections("TemplateSADBEL").Connection, adOpenForwardOnly, adLockReadOnly
                        If Not (rst.BOF And rst.EOF) Then
                            rst.MoveFirst
                            
                            mlngEntrepotID = rst.Fields("Entrepot ID").Value
                            mstrEntrepotNum = rst.Fields("Entrepot Type").Value & "-" & rst.Fields("Entrepot Num").Value
                            strStartDate = DateValue(rst.Fields("Entrepot Date").Value)
                        Else
                            mlngEntrepotID = 0
                        End If
                    Case 3
                        txtFilter(0).Text = Format(txtFilter(0).Text, "0000000")
                        ADORecordsetClose rst
                        Exit Sub
                End Select
                
                'Glenn 3/29/2006
                strEntrepotDate = GetEntrepotLatestReOpenDate
                
                If strEntrepotDate = "" Then
                    'Changes date to Entrepot creation date
                    If strStartDate <> "" And Len(mstrEntrepotNum) > 0 Then
                        dtpPeriod(0).Value = strStartDate
                    Else
                        dtpPeriod(0).Value = DateAdd("m", -1, Date)
                    End If
                Else
                    dtpPeriod(0).Value = DateValue(strEntrepotDate)
                End If
                
                'Only closes recordset if open (1=Open, 0=Close)
                ADORecordsetClose rst
        End If
    End If
End Sub

Private Sub txtFilter_Validate(Index As Integer, Cancel As Boolean)
    If m_blnAnimationRunOnce Then
        If cboTypes(0).ListIndex = 3 Then
            txtFilter(0).Text = Format(txtFilter(0).Text, "0000000")
        End If
    End If
End Sub

Private Function IsEntrepotTextboxEnabled(ByRef conSADBEL As ADODB.Connection) As Boolean
    Dim rstEntrepot As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * FROM ENTREPOTS"
    
    ADORecordsetOpen strSQL, conSADBEL, rstEntrepot, adOpenKeyset, adLockOptimistic
    With rstEntrepot
        '.Open strSQL, conSADBEL, adOpenKeyset, adLockOptimistic
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
            
            If .RecordCount = 1 Then
                IsEntrepotTextboxEnabled = False
                mlngEntrepotID = .Fields("Entrepot_ID").Value
                mstrEntrepotNum = .Fields("Entrepot_Type").Value & "-" & .Fields("Entrepot_Num").Value
            ElseIf .RecordCount > 1 Then
                IsEntrepotTextboxEnabled = True
            End If
        Else
            IsEntrepotTextboxEnabled = False
        End If
        
    End With
    
    ADORecordsetClose rstEntrepot
End Function

Private Sub HideUnhideEntrepotBox(ByVal blnHide As Boolean)
    If blnHide Then
        Label1(2).Top = 990
        dtpPeriod(0).Top = 960
        Label1(3).Top = 1350
        dtpPeriod(1).Top = 1320
        
        Frame1.Height = 1825
        
        chkShowZero.Top = 2050
        cmdAction(0).Top = 2050
        cmdAction(1).Top = 2050
        cmdAction(2).Top = 2050
        
        frmSummaryReports.Height = 3040
    Else
    
        Label1(2).Top = 1350
        dtpPeriod(0).Top = 1320
        Label1(3).Top = 1710
        dtpPeriod(1).Top = 1680
        
        Frame1.Height = 2175
        
        chkShowZero.Top = 2400
        cmdAction(0).Top = 2400
        cmdAction(1).Top = 2400
        cmdAction(2).Top = 2400
        
        frmSummaryReports.Height = 3390
    End If
End Sub

'Added Glenn 3/29/2006
'Added by Rachelle Feb 24, 2006
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetEntrepotLatestReOpenDate() As String
    Dim rstEntrepot As ADODB.Recordset
    Dim strSQL As String

    strSQL = " Select Distinct CDate(InDoc_Date) As [InDoc_Date] FROM (Inbounds Inner Join (Stockcards Inner Join (Products Inner Join Entrepots on " & _
                    " Products.Entrepot_ID = Entrepots.Entrepot_ID) on Stockcards.Prod_ID = Products.Prod_ID) on Inbounds.Stock_ID = Stockcards.Stock_ID) " & _
                    " Inner Join InboundDocs on Inbounds.InDoc_ID = InboundDocs.InDoc_ID Where Trim(Inbounds.In_Code) = '<<Closure>>' And " & _
                    " Entrepots.Entrepot_ID = " & mlngEntrepotID & " Order By CDate(InDoc_Date) Desc"
    ADORecordsetOpen strSQL, m_colConnections("TemplateSadbel").Connection, rstEntrepot, adOpenKeyset, adLockOptimistic
    With rstEntrepot
        '.Open strSQL, m_colConnections("TemplateSadbel").Connection, adOpenKeyset, adLockOptimistic
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
            
            GetEntrepotLatestReOpenDate = .Fields("InDoc_Date").Value
        Else
            GetEntrepotLatestReOpenDate = ""
        End If
    End With
    
    ADORecordsetClose rstEntrepot
End Function
