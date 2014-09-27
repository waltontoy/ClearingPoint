VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLicensee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licensee Information"
   ClientHeight    =   4725
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgBrowse 
      Left            =   2520
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Tag             =   "All"
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Tag             =   "All"
      Top             =   4260
      Width           =   1215
   End
   Begin TabDlg.SSTab tabLicensee 
      Height          =   4035
      Left            =   120
      TabIndex        =   12
      Tag             =   "All"
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7117
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frmLicensee.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Miscellaneous"
      TabPicture(1)   =   "frmLicensee.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Logo"
      TabPicture(2)   =   "frmLicensee.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "cmdClear"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear Picture"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71160
         TabIndex        =   18
         Tag             =   "Logo"
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   3570
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtWebsite 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3120
            Width           =   2895
         End
         Begin VB.TextBox txtCity 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   2710
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1320
            Width           =   1850
         End
         Begin VB.TextBox txtPostalCode 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   2
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtAddress 
            DataSource      =   "rstLicenseeADO"
            Height          =   675
            Left            =   1680
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   2865
         End
         Begin VB.TextBox txtEmail 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   8
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox txtFax 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "+32"
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox txtPhone 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "General"
            Text            =   "+32"
            Top             =   2040
            Width           =   2895
         End
         Begin VB.TextBox txtName 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   40
            TabIndex        =   0
            Top             =   240
            Width           =   2865
         End
         Begin VB.TextBox txtCountry 
            DataSource      =   "rstLicenseeADO"
            Height          =   315
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "Belgium"
            Top             =   1680
            Width           =   2910
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4590
            TabIndex        =   5
            Tag             =   "General"
            Top             =   1680
            Width           =   315
         End
         Begin VB.Label Label5 
            Caption         =   "Postal Code/City :"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1380
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Website :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   3150
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "Country :"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1710
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "E-mail :"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   2790
            Width           =   1140
         End
         Begin VB.Label Label20 
            Caption         =   "Fax :"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2430
            Width           =   1140
         End
         Begin VB.Label Label19 
            Caption         =   "Phone :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   2070
            Width           =   1140
         End
         Begin VB.Label Label18 
            Caption         =   "Address :"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label17 
            Caption         =   "Name :"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   255
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3570
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtLegalInformation 
            Height          =   2475
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   960
            Width           =   2895
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Height          =   315
            Index           =   2
            Left            =   4560
            TabIndex        =   16
            Tag             =   "Miscellaneous"
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   4560
            TabIndex        =   14
            Tag             =   "Miscellaneous"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCurrency 
            Height          =   315
            Left            =   1680
            TabIndex        =   15
            Text            =   "EUR"
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtLanguage 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "Dutch"
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label4 
            Caption         =   "Legal Information :"
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   1005
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Currency :"
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Language :"
            Height          =   225
            Left            =   240
            TabIndex        =   23
            Top             =   285
            Width           =   1275
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3570
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse..."
            Height          =   375
            Left            =   3720
            TabIndex        =   19
            Tag             =   "Logo"
            Top             =   3090
            Width           =   1215
         End
         Begin VB.Image imgLastLogo 
            Height          =   2535
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Image imgLogo 
            Height          =   2535
            Left            =   160
            Top             =   250
            Width           =   615
         End
         Begin VB.Shape shpLogo 
            Height          =   3225
            Left            =   150
            Top             =   240
            Width           =   3435
         End
      End
   End
End
Attribute VB_Name = "frmLicensee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum cpiLicenseeTabConstants
   cpiGeneral = 0
   cpiMiscellaneous = 1
   cpiLogo = 2
End Enum

Private Enum cpiLicenseePicklistConstants
   cpiCountry = 0
   cpiLanguage = 1
   cpiCurrency = 2
End Enum

Public UserConnection As ADODB.Connection
Private m_clsLicensee As CLicensee
Attribute m_clsLicensee.VB_VarHelpID = -1

Dim m_blnAutoSearch As Boolean
Dim strLogoProperties As String
Dim strPictureFile As String


Private Sub cmdbrowse_Click()

    Dim strClass As String
    Dim strRootName() As String
    Dim lngTempWidth As Long
    Dim lngTempHeight As Long
    Dim dblRatio As Double
    Dim strLastPictureFile As String
    
    Set imgLastLogo.Picture = imgLogo.Picture
    
    strLastPictureFile = strPictureFile
    
    cdgBrowse.ShowOpen
    
    imgLogo.Stretch = False
    imgLogo.Visible = False
    
    On Error GoTo ErrHandler
    
    If cdgBrowse.FileTitle <> "" Then
    
        Set imgLogo.Picture = LoadPicture(cdgBrowse.FileName, , vbLPColor)
        
        strPictureFile = cdgBrowse.FileName
        cmdClear.Enabled = True
        
        lngTempWidth = imgLogo.Width
        lngTempHeight = imgLogo.Height
        
        ' check proper size
        If ((lngTempWidth) > shpLogo.Width - 30) Or _
            ((lngTempHeight) > shpLogo.Height - 30) Then
        
            imgLogo.Stretch = True
            
            ' ratio of width/height vice versa
            If lngTempWidth > lngTempHeight Then
            
                dblRatio = lngTempHeight / lngTempWidth
                imgLogo.Width = shpLogo.Width - 30
                imgLogo.Height = imgLogo.Width * dblRatio
            
            ElseIf lngTempWidth < lngTempHeight Then
            
                dblRatio = lngTempWidth / lngTempHeight
                imgLogo.Height = shpLogo.Height - 30
                imgLogo.Width = imgLogo.Height * dblRatio
                
            End If
    
        End If
    
        strLogoProperties = CStr(imgLogo.Width) & "***" & CStr(imgLogo.Height) & "***" & _
                                            CStr(lngTempWidth) & "***" & CStr(lngTempHeight)
        
        imgLogo.Visible = True
    
    Else
        imgLogo.Visible = True
    
    End If
    
    Exit Sub
    
    
ErrHandler:
    
    MsgBox "Invalid picture file.", vbInformation, "Licensee (7027)"
    
    imgLogo.Stretch = True
    imgLogo.Visible = True
    
    Set imgLogo.Picture = imgLastLogo.Picture
    
    strPictureFile = strLastPictureFile
    
End Sub

Private Sub cmdCancel_Click()

   Unload Me
   
End Sub

Private Sub cmdClear_Click()

   strPictureFile = ""
   
   Set imgLogo.Picture = Nothing
   
   cmdClear.Enabled = False
   cmdBrowse.SetFocus

End Sub

Private Sub cmdOK_Click()
        
    Dim strCountry As String
    Dim strCurrency As String
    Dim strLanguage As String
    Dim blnCancel As Boolean
        
   If (CheckData = True) Then
        
        strCountry = txtCountry.Text
        strLanguage = txtLanguage.Text
        strCurrency = txtCurrency.Text
        Call m_clsLicensee.FLicensee_BeforeUpdate(blnCancel, strCountry, strCurrency, strLanguage)
    
    If (blnCancel = False) Then
        MousePointer = vbHourglass
        txtCountry.Text = strCountry
        txtLanguage.Text = strLanguage
        txtCurrency.Text = strCurrency
        Call SaveNewData
        MousePointer = vbDefault
        Unload Me
    End If
      
   
   End If
   
End Sub

Private Sub cmdPicklist_Click(Index As Integer)

    Dim blnCancel As Boolean
    Dim strCountry As String
    Dim strCurrency As String
    Dim strLanguage As String
    Dim strCode As String
    
    

   Select Case Index
   
      Case cpiCountry
      
        If m_blnAutoSearch = True Then
            strCountry = txtCountry.Text
        Else
            strCountry = ""
        End If
        
        Call m_clsLicensee.FLicensee_CountryPicklist(blnCancel, strCode, strCountry)
        
        If (blnCancel = False) And (strCountry <> "") Then
            txtCountry.Text = strCountry
            txtCountry.Tag = strCode
        End If

      Case cpiLanguage
        
        If m_blnAutoSearch = True Then
            strCode = txtLanguage.Text
        Else
            strCode = ""
        End If
        
        Call m_clsLicensee.FLicensee_LanguagePicklist(blnCancel, strCode, strLanguage)
        
        If (blnCancel = False) And (strLanguage <> "") Then
            txtLanguage.Text = strCode
            txtLanguage.Tag = strCode
        End If
     
      
      Case cpiCurrency
        
        If m_blnAutoSearch = True Then
            strCode = txtCurrency.Text
        Else
            strCode = ""
        End If
        
        Call m_clsLicensee.FLicensee_CurrencyPicklist(blnCancel, strCode, strCurrency)
        
        If (blnCancel = False) And (strCurrency <> "") Then
            txtCurrency.Text = strCode
            txtCurrency.Tag = strCode
        End If
         
   End Select

    ' refresh this form
   Refresh
   SendKeysEx "{End}"

End Sub

Private Sub Form_Load()
   
   cdgBrowse.Filter = "Bitmap Files (*.bmp)|*.bmp|" & _
                                       "JPEG File Interchange Format (*.jpg;*.jpeg)|*.jpg;*.jpeg|" & _
                                       "GIF (*.gif)|*.gif|" & _
                                       "All Picture Files|*.bmp;*.jpg;*.jpeg;*.gif|" & _
                                       "All Files|*.*"
                                       
   cdgBrowse.FilterIndex = 4
                                       
    'ResetFormHeight Me
                                       
   Call LoadLicenseeRecord
                                       
                                       
                                       
End Sub




Private Sub Form_Unload(Cancel As Integer)
    
    Call m_clsLicensee.FLicensee_UnloadForm
    
End Sub

Private Sub tabLicensee_Click(PreviousTab As Integer)
    Dim ctlSetter As Control
    Dim intIndex As Integer
    
    Dim strTypeName As String
    Dim strTag As String
    
    ' remove all TabStop
    For Each ctlSetter In Me.Controls
              
        strTypeName = TypeName(ctlSetter)
        
        If strTypeName = "TextBox" Or strTypeName = "CommandButton" Or strTypeName = "SSTab" Or strTypeName = "Checkbox" Then
            ctlSetter.TabStop = False
        End If
        
    Next 'ctlsetter
    
    Select Case tabLicensee.Tab
        
        Case cpiGeneral
            strTag = "General"
            intIndex = 10
            If (txtName.Enabled = True) Then
                txtName.SetFocus
            Else
                txtPhone.SetFocus
            End If
            
        Case cpiMiscellaneous
            strTag = "Miscellaneous"
            intIndex = 18
            txtLanguage.SetFocus
            
        Case cpiLogo
            strTag = "Logo"
            intIndex = 19
            cmdBrowse.SetFocus
            
    End Select
    
    cmdOk.TabIndex = intIndex
    cmdCancel.TabIndex = intIndex + 1
    tabLicensee.TabIndex = intIndex + 2
    
    ' set active group Tabstop
    For Each ctlSetter In Me.Controls
        
        strTypeName = TypeName(ctlSetter)
        
        If strTypeName = "TextBox" Or strTypeName = "CommandButton" Or strTypeName = "SSTab" Then
            
            If ctlSetter.Tag = strTag Or ctlSetter.Tag = "All" Then
                ctlSetter.TabStop = True
            End If
            
        End If
        
    Next 'ctlsetter
   
End Sub

Private Sub txtCountry_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Select Case KeyCode
    
        Case vbKeyReturn
        
            m_blnAutoSearch = True
        
            Call cmdPicklist_Click(cpiCountry)
        
            m_blnAutoSearch = False
            
        Case vbKeyF2
        
            m_blnAutoSearch = False
        
            Call cmdPicklist_Click(cpiCountry)
    
    End Select

End Sub

Private Sub txtCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Select Case KeyCode
      
      Case vbKeyReturn
            
            m_blnAutoSearch = True
            
            Call cmdPicklist_Click(cpiCurrency)
      
            m_blnAutoSearch = False
            
      Case vbKeyF2
            
            m_blnAutoSearch = False
            
            Call cmdPicklist_Click(cpiCurrency)
   
   End Select

End Sub

Private Sub txtLanguage_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Select Case KeyCode
      
      Case vbKeyReturn
            
            m_blnAutoSearch = True
            
            Call cmdPicklist_Click(cpiLanguage)
      
            m_blnAutoSearch = False
            
      Case vbKeyF2
            
            m_blnAutoSearch = False
            
            Call cmdPicklist_Click(cpiLanguage)
   
   End Select

End Sub
Private Sub LoadLicenseeRecord()

   'Dim clsRecord As CRecordset
   Dim rstLicensee As ADODB.Recordset
   Dim rstCurrency As ADODB.Recordset
   Dim rstLanguage As ADODB.Recordset
   Dim strCommandText As String
   
   strCommandText = _
                                "SELECT Licensee.Lic_Name AS [Name], " & _
                                    "Licensee.Lic_Address AS [Address], " & _
                                    "Licensee.Lic_PostalCode AS [Zip Code], " & _
                                    "Licensee.Lic_City AS [City], " & _
                                    "Licensee.Lic_Country AS [Country], " & _
                                    "Licensee.Lic_Phone AS [Phone], " & _
                                    "Licensee.Lic_Fax AS [Fax], " & _
                                    "Licensee.Lic_Email AS [Email_A], " & _
                                    "Licensee.Lic_Website AS [Website_A], " & _
                                    "Licensee.Lic_Language AS [Language], " & _
                                    "Licensee.Lic_Currency AS [Currency], " & _
                                    "Licensee.Lic_LegalInfo AS [Legal Info], " & _
                                    "Licensee.Lic_Logo As [Logo], " & _
                                    "Licensee.Lic_Logosize AS [Logo Size], " & _
                                    "Licensee.Lic_LogoProperties AS [Logo Properties] " & _
                                "FROM Licensee"

    Dim strCode As String
    Dim strDesc As String
    Dim strPictureValue As String
    
    Dim blnCancel As Boolean
    
    'Set clsRecord = New CRecordset
    
    ADORecordsetOpen strCommandText, UserConnection, rstLicensee, adOpenKeyset, adLockOptimistic
    'clsRecord.cpiOpen strCommandText, UserConnection, rstLicensee, adOpenKeyset, adLockOptimistic, , True
    
    strPictureFile = ""
    cmdClear.Enabled = False
   
    If (rstLicensee.RecordCount <> 0) Then
      
'        txtName.Text = IIf(IsNull(rstLicensee.Fields("Name").Value), "", rstLicensee.Fields("Name").Value)
'        txtAddress.Text = IIf(IsNull(rstLicensee.Fields("Address").Value), "", rstLicensee.Fields("Address").Value)
'        txtPostalCode.Text = IIf(IsNull(rstLicensee.Fields("Zip Code").Value), "", rstLicensee.Fields("Zip Code").Value)
'        txtCity.Text = IIf(IsNull(rstLicensee.Fields("City").Value), "", rstLicensee.Fields("City").Value)
'        txtCountry.Text = IIf(IsNull(rstLicensee.Fields("Country").Value), "", rstLicensee.Fields("Country").Value)
'        txtCountry.Tag = ""
        
        txtName.Text = g_typInterface.ILicense.RegCompany
        txtAddress.Text = g_typInterface.ILicense.RegAddress1
        
        If InStr(1, g_typInterface.ILicense.RegAddress2, "|||||") > 0 Then
            txtPostalCode.Text = Left(g_typInterface.ILicense.RegAddress2, InStr(1, g_typInterface.ILicense.RegAddress2, "|||||") - 1)
            txtCity.Text = Mid(g_typInterface.ILicense.RegAddress2, InStr(1, g_typInterface.ILicense.RegAddress2, "|||||") + 5)
        End If
        
        txtCountry.Text = g_typInterface.ILicense.RegAddress3
        txtCountry.Tag = ""
        
        txtPhone.Text = IIf(IsNull(rstLicensee.Fields("Phone").Value), "", rstLicensee.Fields("Phone").Value)
        txtFax.Text = IIf(IsNull(rstLicensee.Fields("Fax").Value), "", rstLicensee.Fields("Fax").Value)
        txtEmail.Text = IIf(IsNull(rstLicensee.Fields("Email_A").Value), "", rstLicensee.Fields("Email_A").Value)
        txtWebsite.Text = IIf(IsNull(rstLicensee.Fields("Website_A").Value), "", rstLicensee.Fields("Website_A").Value)
        txtLanguage.Text = IIf(IsNull(rstLicensee.Fields("Language").Value), "", rstLicensee.Fields("Language").Value)
        txtCurrency.Text = IIf(IsNull(rstLicensee.Fields("Currency").Value), "", rstLicensee.Fields("Currency").Value)
        txtLegalInformation.Text = IIf(IsNull(rstLicensee.Fields("Legal Info").Value), "", rstLicensee.Fields("Legal Info").Value)
        
        strPictureValue = IIf(IsNull(rstLicensee.Fields("Logo Properties").Value), "", rstLicensee.Fields("Logo Properties").Value)
        
        If (strPictureValue <> "") Then
        
           Call Load_Image(rstLicensee)
           
           cmdClear.Enabled = True
           
        End If
        
        '--->Disable Licensee Name field if it's already specified
'        If (Len(Trim$(g_typInterface.ILicense.RegCompany)) > 0) Then
        If g_typInterface.ILicense.ExpireMode = "N" Then
            txtName.Enabled = False
            txtAddress.Enabled = False
            txtPostalCode.Enabled = False
            txtCity.Enabled = False
            txtCountry.Enabled = False
            cmdPicklist(0).Enabled = False
        End If
    End If
    
    ADORecordsetClose rstLicensee
    
    'Set rstLicensee = Nothing
    'Set clsRecord = Nothing
    
End Sub

Public Sub SaveImage(ByRef rstLicensee As ADODB.Recordset)
    
    Dim lngFileSize As Long
    Dim nHandle As Integer
    Dim Chunk() As Byte
    Dim strImagePath As String
            
   strImagePath = cdgBrowse.FileName
   
   If (strImagePath <> "") Then
   
      nHandle = FreeFile
      
      Open strImagePath For Binary Access Read As nHandle
   
      lngFileSize = LOF(nHandle)
      
      If (nHandle = 0) Then
      
          Close nHandle
          
      End If
      
      ReDim Chunk(lngFileSize)
      Get nHandle, , Chunk()
      
      rstLicensee.Fields("Logo").Value = Null
      rstLicensee.Fields("Logo").AppendChunk Chunk()
                     
      rstLicensee.Fields("Logo Size").Value = CStr(lngFileSize)
        
      rstLicensee.Update
      
      UpdateRecordset UserConnection, rstLicensee, "Licensee"
      
      Close nHandle
       
      'rstLicensee.Update
      
   End If
      
End Sub

Public Sub Load_Image(ByRef rstLicensee As ADODB.Recordset)

    Dim lngFileSize As Long
    Dim varChunk() As Byte
    Dim strAppPath As String
    Dim nHandle As Integer
    Dim strFilePathName As String
    Dim arrImageData() As String
    Dim strMdbPath As String
    
    strLogoProperties = IIf(IsNull(rstLicensee.Fields("Logo Properties").Value), _
                                    "", rstLicensee.Fields("Logo Properties").Value)
      
    If (strLogoProperties <> "") Then
    
        imgLogo.Stretch = True
        arrImageData = Split(strLogoProperties, "***")
        imgLogo.Width = CLng(arrImageData(0))
        imgLogo.Height = CLng(arrImageData(1))
    
    End If
    
    strMdbPath = GetMDBPath(UserConnection.ConnectionString)

    nHandle = FreeFile

    strFilePathName = NoBackSlash(GetTemporaryPath) & "\" & "output.bin"
    
    ' check if exist
    If (Dir(strFilePathName) = "output.bin") Then
        SetAttr strFilePathName, vbNormal
        Kill strFilePathName
    End If
    
    Open strFilePathName For Binary Access Write As nHandle
    
    lngFileSize = rstLicensee.Fields("Logo Size").Value
    
    varChunk() = rstLicensee.Fields("Logo").GetChunk(lngFileSize)
    
    Put nHandle, , varChunk()
    
    Close nHandle
    
    Set imgLogo.Picture = LoadPicture(strFilePathName, , vbLPColor)
    
    strPictureFile = strFilePathName
    
    ' Delete the temporary image file
    If (Len(Dir(strFilePathName)) > 0) Then
        On Error Resume Next
        Kill strFilePathName
        On Error GoTo 0
    End If
    
End Sub

Private Function CheckData() As Boolean

    If (Len(Trim$(txtName.Text)) = 0 And txtName.Enabled = True) Then
    
        MsgBox "Enter the Licensee name.", vbInformation, "Field Required (7028)"
        txtName.Text = ""
        tabLicensee.Tab = 0
        txtName.SetFocus
        
    ElseIf (Len(Trim$(txtAddress.Text)) = 0 And txtAddress.Enabled = True) Then
    
        MsgBox "Enter the Address.", vbInformation, "Field Required (7029)"
        txtAddress.Text = ""
        tabLicensee.Tab = 0
        txtAddress.SetFocus
        
    ElseIf (Len(Trim$(txtPostalCode.Text)) = 0 And txtPostalCode.Enabled = True) Then
        MsgBox "Enter the Postal Code.", vbInformation, "Field Required (7030)"
        txtPostalCode.Text = ""
        tabLicensee.Tab = 0
        txtPostalCode.SetFocus
        
    ElseIf (Len(Trim$(txtCity.Text)) = 0 And txtCity.Enabled = True) Then
    
        MsgBox "Enter the City.", vbInformation, "Field Required (7031)"
        txtCity.Text = ""
        tabLicensee.Tab = 0
        txtCity.SetFocus
        
    ElseIf (Len(Trim$(txtCountry.Text)) = 0 And txtCountry.Enabled = True) Then
    
        MsgBox "Pick a Country.", vbInformation, "Field Required (7032)"
        txtCountry.Text = ""
        tabLicensee.Tab = 0
        cmdPicklist(0).SetFocus
         
    ElseIf (Len(Trim$(txtLanguage.Text)) = 0) Then
    
        MsgBox "Pick a Language.", vbInformation, "Field Required (7033)"
        txtLanguage.Text = ""
        tabLicensee.Tab = 1
        cmdPicklist(1).SetFocus
         
    ElseIf (Len(Trim$(txtCurrency.Text)) = 0) Then
    
        MsgBox "Pick a Currency.", vbInformation, "Field Required (7034)"
        txtCurrency.Text = ""
        tabLicensee.Tab = 1
        cmdPicklist(2).SetFocus
         
    Else
    
       CheckData = True
       
     End If

End Function

Private Sub SaveNewData()

     ' save all the field values
    
    'Dim clsRecord As CRecordset
    Dim rstLicensee As ADODB.Recordset
    Dim strCommandText As String
    Dim blnAddNew As Boolean
       
    strCommandText = _
                            "SELECT Licensee.Lic_Name AS [Name], " & _
                                "Licensee.Lic_Address AS [Address], " & _
                                "Licensee.Lic_PostalCode AS [Zip Code], " & _
                                "Licensee.Lic_City AS [City], " & _
                                "Licensee.Lic_Country AS [Country], " & _
                                "Licensee.Lic_Phone AS [Phone], " & _
                                "Licensee.Lic_Fax AS [Fax], " & _
                                "Licensee.Lic_Email AS [Email_A], " & _
                                "Licensee.Lic_Website AS [Websites_A], " & _
                                "Licensee.Lic_Language AS [Language], " & _
                                "Licensee.Lic_Currency AS [Currency], " & _
                                "Licensee.Lic_LegalInfo AS [Legal Info], " & _
                                "Licensee.Lic_Logo As [Logo], " & _
                                "Licensee.Lic_Logosize AS [Logo Size], " & _
                                "Licensee.Lic_LogoProperties AS [Logo Properties] " & _
                            "FROM Licensee"
    
    'Set clsRecord = New CRecordset
    
    ADORecordsetOpen strCommandText, UserConnection, rstLicensee, adOpenKeyset, adLockOptimistic
    'clsRecord.cpiOpen strCommandText, UserConnection, rstLicensee, adOpenKeyset, adLockOptimistic
    
    On Error GoTo ErrorHandler
    
    
    blnAddNew = (rstLicensee.RecordCount = 0)
    
    If blnAddNew Then
    
        rstLicensee.AddNew
    
    End If
    
    rstLicensee.Fields("Name").Value = txtName.Text
    rstLicensee.Fields("Address").Value = txtAddress.Text
    rstLicensee.Fields("Zip Code").Value = txtPostalCode.Text
    rstLicensee.Fields("City").Value = txtCity.Text
    
    rstLicensee.Fields("Country").Value = txtCountry.Text
    If Trim(txtPhone.Text) <> "" Then
        rstLicensee.Fields("Phone").Value = txtPhone.Text
    Else
        rstLicensee.Fields("Phone").Value = ""
    End If
    If Trim(txtFax.Text) <> "" Then
        rstLicensee.Fields("Fax").Value = txtFax.Text
    Else
        rstLicensee.Fields("Fax").Value = ""
    End If
    If Trim(txtEmail.Text) <> "" Then
        rstLicensee.Fields("Email_A").Value = txtEmail.Text
    Else
        rstLicensee.Fields("Email_A").Value = ""
    End If
    If Trim(txtWebsite.Text) <> "" Then
        rstLicensee.Fields("Websites_A").Value = txtWebsite.Text
    Else
        rstLicensee.Fields("Websites_A").Value = ""
    End If
    If Trim(txtLanguage.Text) <> "" Then
        rstLicensee.Fields("Language").Value = txtLanguage.Text
    Else
        rstLicensee.Fields("Language").Value = ""
    End If
    If Trim(txtCurrency.Text) <> "" Then
        rstLicensee.Fields("Currency").Value = txtCurrency.Text
    Else
        rstLicensee.Fields("Currency").Value = ""
    End If
    If Trim(txtLegalInformation.Text) <> "" Then
        rstLicensee.Fields("Legal Info").Value = txtLegalInformation.Text
    Else
        rstLicensee.Fields("Legal Info").Value = ""
    End If
    
    If (Right$(strPictureFile, 3) <> "bin") Then
    
        If (strPictureFile <> "") Then
        
            Call SaveImage(rstLicensee)
            
            rstLicensee.Fields("Logo Properties").Value = strLogoProperties
            
        ElseIf (strPictureFile = "") Then
        
            rstLicensee.Fields("Logo").Value = Null
            
            rstLicensee.Fields("Logo Properties").Value = ""
            rstLicensee.Fields("Logo Size").Value = ""
        
        End If
    
    End If
    
    rstLicensee.Update
    
    If blnAddNew Then
        InsertRecordset UserConnection, rstLicensee, "Licensee"
    Else
        UpdateRecordset UserConnection, rstLicensee, "Licensee"
    End If
    
    ADORecordsetClose rstLicensee
    
    'Set rstLicensee = Nothing
    'Set clsRecord = Nothing
    
    Exit Sub
      
      
ErrorHandler:

   MsgBox Err.Description, vbInformation, "Licensee (7033)"

End Sub

Public Sub ShowForm(ByRef OwnerForm As Object, ByRef ActiveLicensee As CLicensee, ByRef ADOConnection As ADODB.Connection)
    
    Set UserConnection = ADOConnection
    Set m_clsLicensee = ActiveLicensee
    
    Set Me.Icon = OwnerForm.Icon
    Me.Show vbModal
    
    Set ADOConnection = UserConnection
    
End Sub
