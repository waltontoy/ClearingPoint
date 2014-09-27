VERSION 5.00
Begin VB.Form frmAuthorizedParty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authorized Party"
   ClientHeight    =   3960
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   4770
   Icon            =   "frmAuthorizedParty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAuthorized 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Index           =   2
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdCountry 
         Caption         =   "..."
         Height          =   315
         Left            =   4050
         TabIndex        =   5
         Top             =   1635
         Width           =   315
      End
      Begin VB.TextBox txtCity 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "City"
         Top             =   1275
         Width           =   2205
      End
      Begin VB.TextBox txtPostalCode 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "ZIP"
         Top             =   1275
         Width           =   735
      End
      Begin VB.TextBox txtAddress 
         DataSource      =   "rstLicenseeADO"
         Height          =   630
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2925
      End
      Begin VB.TextBox txtEmail 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "e-mail Address"
         Top             =   2715
         Width           =   2925
      End
      Begin VB.TextBox txtFax 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Fax Number"
         Top             =   2355
         Width           =   2925
      End
      Begin VB.TextBox txtPhone 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "Telephone Number"
         Top             =   1995
         Width           =   2925
      End
      Begin VB.TextBox txtName 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "Client"
         Top             =   240
         Width           =   2925
      End
      Begin VB.TextBox txtCountry 
         DataSource      =   "rstLicenseeADO"
         Height          =   315
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "Country"
         Top             =   1635
         Width           =   2625
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Tag             =   "2096"
         Top             =   2775
         Width           =   1140
      End
      Begin VB.Label lblFax 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Tag             =   "964"
         Top             =   2415
         Width           =   1140
      End
      Begin VB.Label lblPhone 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Tag             =   "963"
         Top             =   2055
         Width           =   1140
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Tag             =   "961"
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Tag             =   "1055"
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3465
      TabIndex        =   10
      Tag             =   "179"
      Top             =   3495
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Tag             =   "178"
      Top             =   3495
      Width           =   1215
   End
End
Attribute VB_Name = "frmAuthorizedParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLanguage As String
Private m_conConnection As ADODB.Connection
Private pckCountry As PCubeLibPick.CPicklist
Private rstAuthorizedParty As ADODB.Recordset
Private ButtonType As PCubeLibPick.ButtonType
Private mblnCancel As Boolean

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdCountry_Click()
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
        .Search True, "Key Description", Trim(txtCountry.Text)
        
        ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick Me, cpiSimplePicklist, m_conConnection, strCountrySQL, "Key Code", "Countries", vbModal, gsdCountry, , , True, cpiKeyF2
        
        If Not .SelectedRecord Is Nothing Then
            txtCountry.Text = .SelectedRecord.RecordSource.Fields("Key Description").Value
        End If
    End With
    
    Set gsdCountry = Nothing
    Set pckCountry = Nothing
End Sub

Private Sub cmdOK_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox Translate(2276), vbInformation, Translate(2308)
        Exit Sub
    End If
    MousePointer = vbHourglass
    
    mblnCancel = False
    
    With rstAuthorizedParty
        rstAuthorizedParty("Address") = Me.txtAddress
        rstAuthorizedParty("Auth_City") = Me.txtCity
        rstAuthorizedParty("Auth_Country") = Me.txtCountry
        rstAuthorizedParty("Auth_Email") = Me.txtEmail
        rstAuthorizedParty("Auth_Fax") = Me.txtFax
        rstAuthorizedParty("Name") = Me.txtName
        rstAuthorizedParty("Auth_Phone") = Me.txtPhone
        rstAuthorizedParty("Auth_PostalCode") = Me.txtPostalCode
    End With
    
    rstAuthorizedParty.Update
        
    UpdateRecordset m_conConnection, rstAuthorizedParty, "AuthorizedParties"
    
    Me.MousePointer = vbHourglass
    Me.MousePointer = vbDefault
    
    Unload Me
End Sub

Public Sub MyLoad(ByRef rstRecord As ADODB.Recordset, Button As PCubeLibPick.ButtonType, ByRef Cancel As Boolean, ByVal Language As String, ByRef Connection As ADODB.Connection, ByVal MyResourceHandler As Long)
    ResourceHandler = MyResourceHandler
    LoadResStrings Me, True
    
    Set m_conConnection = Connection
    mLanguage = Language
    ButtonType = Button
    Set rstAuthorizedParty = rstRecord
    LoadValues Button
    Me.Show vbModal
    Cancel = mblnCancel
End Sub

Private Sub LoadValues(Button As PCubeLibPick.ButtonType)
     Me.txtName.Text = IIf(IsNull(rstAuthorizedParty!Name), "", rstAuthorizedParty!Name)
     Me.txtCity.Text = IIf(IsNull(rstAuthorizedParty!Auth_City), "", rstAuthorizedParty!Auth_City)
     Me.txtEmail.Text = IIf(IsNull(rstAuthorizedParty!Auth_Email), "", rstAuthorizedParty!Auth_Email)
     Me.txtFax.Text = IIf(IsNull(rstAuthorizedParty!Auth_Fax), "", rstAuthorizedParty!Auth_Fax)
     Me.txtPostalCode.Text = IIf(IsNull(rstAuthorizedParty!Auth_PostalCode), "", rstAuthorizedParty!Auth_PostalCode)
     Me.txtPhone.Text = IIf(IsNull(rstAuthorizedParty!Auth_Phone), "", rstAuthorizedParty!Auth_Phone)
     Me.txtAddress.Text = IIf(IsNull(rstAuthorizedParty!Address), "", rstAuthorizedParty!Address)
     Me.txtCountry.Text = IIf(IsNull(rstAuthorizedParty!Auth_Country), "", rstAuthorizedParty!Auth_Country)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mblnCancel = True
    End If
End Sub

Private Sub txtCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdCountry_Click
    End If
End Sub
