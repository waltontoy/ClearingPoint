VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "321"
   ClientHeight    =   4665
   ClientLeft      =   3210
   ClientTop       =   2220
   ClientWidth     =   4830
   Icon            =   "frmDocProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "321"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3435
      TabIndex        =   1
      Tag             =   "178"
      Top             =   4200
      Width           =   1260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "269"
      TabPicture(0)   =   "frmDocProperties.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "290"
      TabPicture(1)   =   "frmDocProperties.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(1)=   "Label8(4)"
      Tab(1).Control(2)=   "Label8(3)"
      Tab(1).Control(3)=   "Label8(2)"
      Tab(1).Control(4)=   "Label8(1)"
      Tab(1).Control(5)=   "Label8(0)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "270"
      TabPicture(2)   =   "frmDocProperties.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSComctlLib.ListView ListView1 
         Height          =   1845
         Left            =   -74775
         TabIndex        =   4
         Top             =   2100
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3254
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date/Time Created:"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Tag             =   "273"
         Top             =   1785
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date/Time Send:"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Tag             =   "276"
         Top             =   3405
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date/Time Last Modified:"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Tag             =   "274"
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date/Time Requested:"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Tag             =   "275"
         Top             =   2865
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Left            =   225
         TabIndex        =   15
         Tag             =   "272"
         Top             =   1110
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Name:"
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Tag             =   "271"
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   0
         Left            =   1755
         TabIndex        =   13
         Top             =   645
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   1
         Left            =   1755
         TabIndex        =   12
         Top             =   1110
         Width           =   1470
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   165
         X2              =   4365
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   165
         X2              =   4365
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   11
         Top             =   1785
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   3
         Left            =   2370
         TabIndex        =   10
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   4
         Left            =   2370
         TabIndex        =   9
         Top             =   2865
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "passed as parameter"
         Height          =   195
         Index           =   5
         Left            =   2370
         TabIndex        =   8
         Top             =   3405
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   4
         Left            =   -74775
         TabIndex        =   7
         Tag             =   "292"
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   -74775
         TabIndex        =   6
         Top             =   1485
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   -74775
         TabIndex        =   5
         Top             =   870
         Width           =   4095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes Involved:"
         Height          =   195
         Index           =   1
         Left            =   -74775
         TabIndex        =   3
         Tag             =   "293"
         Top             =   1875
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Error Code:"
         Height          =   195
         Index           =   0
         Left            =   -74775
         TabIndex        =   2
         Tag             =   "291"
         Top             =   645
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmDocProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mvarUserNo As Long
Dim mvarActiveConnection As ADODB.Connection
Dim mvarResourceHandle As Long
Dim mvarActiveLanguage As String
Dim mvarUniqueCode As String
Dim CodisheetType As cpiCodiSheetTypeEnums

Dim clsDataNctsTable As cpiDataNctsTable
Dim blnNotFound As Boolean

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    TranslateCaptions

    If (blnNotFound = False) Then
        Me.Label7(0).Caption = clsDataNctsTable.FIELD_DOCUMENT_NAME ' ![Document Name]
        Me.Label7(1).Caption = clsDataNctsTable.FIELD_USERNAME ' cUser
        Me.Label7(2).Caption = clsDataNctsTable.FIELD_DATE_CREATED ' ![Date Created]
        Me.Label7(3).Caption = clsDataNctsTable.FIELD_DATE_LAST_MODIFIED ' ![Date Last Modified]
        Me.Label7(4).Caption = clsDataNctsTable.FIELD_DATE_REQUESTED ' IIf(Len(Trim(![Date Requested])) > 0, ![Date Requested], " ")
        Me.Label7(5).Caption = clsDataNctsTable.FIELD_DATE_SEND ' IIf(Len(Trim(![Date Send])) > 0, ![Date Send], " ")
    ElseIf (blnNotFound = True) Then
        Me.Label7(0).Caption = clsDataNctsTable.FIELD_DOCUMENT_NAME
        Me.Label7(1).Caption = clsDataNctsTable.FIELD_USERNAME
        Me.Label7(2).Caption = " "
        Me.Label7(3).Caption = " "
        Me.Label7(4).Caption = " "
        Me.Label7(5).Caption = " "
    End If

End Sub

Private Function TranslateCaptions() As Boolean
'  xxx
    '
    Me.Caption = Translate_B(Me.Caption, mvarResourceHandle)
    Label1.Caption = Translate_B(Label1.Tag, mvarResourceHandle)
    Label2.Caption = Translate_B(Label2.Tag, mvarResourceHandle)
    Label3.Caption = Translate_B(Label3.Tag, mvarResourceHandle)
    Label4.Caption = Translate_B(Label4.Tag, mvarResourceHandle)
    Label5.Caption = Translate_B(Label5.Tag, mvarResourceHandle)
    Label6.Caption = Translate_B(Label6.Tag, mvarResourceHandle)
    
    SSTab1.TabCaption(1) = Translate_B(SSTab1.TabCaption(1), mvarResourceHandle)
    SSTab1.TabCaption(2) = Translate_B(SSTab1.TabCaption(2), mvarResourceHandle)
    SSTab1.TabCaption(0) = Translate_B(SSTab1.TabCaption(0), mvarResourceHandle)
    
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
'
End Function

Public Function ShowForm(ByRef OwnerForm As Form, _
                                            ByRef ActiveConnection As ADODB.Connection, _
                                            ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                            ByRef UserNo As Long, _
                                            ByRef ActiveLanguage As String, _
                                            ByRef ResourceHandle As Long, _
                                            ByRef UniqueCode As String) As Boolean
    '
    ' load frmBoxProperties here
    'mvarActiveBoxCode = ActiveBoxCode
    'mvarActiveDocument = ActiveDocument
    
    Dim clsDataNctsTables As cpiDataNctsTables
    Dim strSql As String
    Dim rstDataNCTS As ADODB.Recordset
    
    Set clsDataNctsTable = New cpiDataNctsTable
    Set clsDataNctsTables = New cpiDataNctsTables
    
    strSql = "SELECT * FROM [DATA_NCTS] WHERE [Code]='" & UniqueCode & "'"
    
    ADORecordsetOpen strSql, ActiveConnection, rstDataNCTS, adOpenKeyset, adLockOptimistic
    Set clsDataNctsTables.Recordset = rstDataNCTS
    
    If (clsDataNctsTables.Recordset.EOF = False) Then
        Set clsDataNctsTable = clsDataNctsTables.GetClassRecord(clsDataNctsTables.Recordset)
        blnNotFound = False
    ElseIf (clsDataNctsTables.Recordset.EOF = True) Then
        blnNotFound = True
    End If
    
    mvarActiveLanguage = UCase$(ActiveLanguage)
    mvarResourceHandle = ResourceHandle
    
    CodisheetType = ActiveCodisheet
    
    mvarUserNo = UserNo
    Set mvarActiveConnection = ActiveConnection

    mvarUniqueCode = UniqueCode

    Set Me.Icon = OwnerForm.Icon
    Screen.MousePointer = vbDefault
    
    Set clsDataNctsTables = Nothing
    
    Me.Show vbModal

    ' @@@

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set clsDataNctsTable = Nothing

End Sub
