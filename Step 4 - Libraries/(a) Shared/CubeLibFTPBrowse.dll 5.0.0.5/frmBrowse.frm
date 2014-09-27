VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP"
   ClientHeight    =   6120
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7380
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwTemp 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FileSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CreationDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tag"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLists 
      Left            =   1560
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":058A
            Key             =   "CloseFolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":09DE
            Key             =   "OpenFolder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwBrowse 
      Height          =   5430
      Left            =   3000
      TabIndex        =   2
      Top             =   105
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9578
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   4568
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FileSize"
         Object.Width           =   2848
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvwBrowse 
      Height          =   5430
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9578
      _Version        =   393217
      Indentation     =   617
      Style           =   7
      ImageList       =   "imgLists"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   5430
      Left            =   3015
      TabIndex        =   3
      Top             =   135
      Width           =   4260
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========== NOTES ==================
'AUTHOR: tonio
'Revision History
'   10/24/2003 - omit use of DAO/ADO objects
'===================================

Option Explicit
Private m_strClient As String
Private m_strSubClient As String
Private m_strFolder() As String
Private blnAlreadyLoaded As Boolean
Private m_strUserName As String
Private m_strPassword As String
Private m_strFTPAddress As String
Private blnExit As Boolean
Private oFtp As clsFTP

Private Sub cmdClose_Click()
    blnExit = True

    Unload Me
End Sub

Private Sub Form_Activate()
Dim iRoot As Integer
Dim i As Integer
Dim strRootFolder(1) As String

Dim oFile As clsFile
Dim lvwItem As ListItem

Dim blnOpenSuccess As Boolean

    If blnAlreadyLoaded = True Then Exit Sub
        
    DoEvents
    Label1.ZOrder vbBringToFront
    Label1.Caption = "Please wait..."
    Me.Refresh

    Set oFtp = New clsFTP

    oFtp.Connections.Add m_strFTPAddress, m_strUserName, m_strPassword, "General"
    oFtp.Connections("General").Connect
    
    lvwBrowse.ListItems.Clear
    
    If oFtp.Connections("General").ConnectSuccess Then
        strRootFolder(0) = "TobePrinted"
        strRootFolder(1) = "Archive"
    
        lvwTemp.ListItems.Clear
        
        For iRoot = 0 To 1
            oFtp.Connections("General").BackToRootFolder
            
            Label1.Caption = "Please wait...searching for files under " & strRootFolder(iRoot)
            Me.Refresh
            
            'Glenn - 9/1/2006
            blnOpenSuccess = oFtp.Connections("General").OpenFolder("RemotePrint/" & m_strClient & "/" & m_strSubClient & "/" & strRootFolder(iRoot))
            
            If blnOpenSuccess Then      'Glenn
                For Each oFile In oFtp.Connections("General").Files
                    If blnExit = True Then Exit Sub
                    
                    Set lvwItem = lvwTemp.ListItems.Add(, , oFile.FileName)
                    lvwItem.SubItems(1) = FormatSize(oFile.FileSize)
                    lvwItem.SubItems(2) = oFile.CreationTime
                    lvwItem.SubItems(3) = IIf(iRoot = 0, "T", "A") & TagToUse(m_strFolder, oFile.FileName)
                Next
            End If
        Next
    Else
        Label1.Caption = "Failed to conned to FTP.."
    End If
    lvwBrowse.ListItems.Clear
    
    
    Me.Refresh
    blnAlreadyLoaded = True
    lvwBrowse.Visible = True
    Label1.Visible = False
    
    'COMMENTED: Glenn - disconnecting is called upon form unload.
    'oFtp.Connections("General").Disconnect
    'Set oFtp = Nothing
End Sub

Public Sub MyLoad(strUsername As String, strPassword As String, _
                  strFTPAddress As String, strClient As String, _
                  strSubClient As String, strFolder() As String)

'Dim oFTP As clsFTP
Dim strRootFolder(1) As String
Dim i As Integer
Dim iRoot As Integer
Dim itm As clsFile

    strRootFolder(0) = "TobePrinted"
    strRootFolder(1) = "Archive"

    

    For iRoot = 0 To 1
        frmBrowse.tvwBrowse.Nodes.Add , , IIf(iRoot = 0, "T", "A"), strRootFolder(iRoot), "CloseFolder", "OpenFolder"
        
        For i = 0 To UBound(strFolder)
            Select Case UCase$(Trim$(strFolder(i)))
                Case UCase$(Trim$("PLDA IMPORT"))
                    frmBrowse.tvwBrowse.Nodes.Add IIf(iRoot = 0, "T", "A"), tvwChild, IIf(iRoot = 0, "T", "A") & "PLI", strFolder(i), "CloseFolder", "OpenFolder"
                    
                Case UCase$(Trim$("PLDA EXPORT"))
                    frmBrowse.tvwBrowse.Nodes.Add IIf(iRoot = 0, "T", "A"), tvwChild, IIf(iRoot = 0, "T", "A") & "PLE", strFolder(i), "CloseFolder", "OpenFolder"
                    
                Case Else
                    frmBrowse.tvwBrowse.Nodes.Add IIf(iRoot = 0, "T", "A"), tvwChild, IIf(iRoot = 0, "T", "A") & Left(strFolder(i), 3), strFolder(i), "CloseFolder", "OpenFolder"

            End Select
        Next
    Next
    
    tvwBrowse.Nodes.Item("T").Expanded = True
    tvwBrowse.Nodes.Item("A").Expanded = True
    
    m_strClient = strClient
    m_strSubClient = strSubClient
    m_strFolder() = strFolder()
    m_strUserName = strUsername
    m_strPassword = strPassword
    m_strFTPAddress = strFTPAddress
    
    lvwBrowse.Visible = False
    Me.Show
End Sub

Private Sub Form_Load()
    blnExit = False
    blnAlreadyLoaded = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    For i = 1 To oFtp.Connections.Count
        oFtp.Connections(i).Disconnect
    Next

    Set oFtp = Nothing
End Sub

Private Sub tvwBrowse_Click()
    Dim itmTemp As ListItem
    Dim itm As ListItem

    lvwBrowse.ListItems.Clear

    For Each itmTemp In lvwTemp.ListItems
        If UCase(itmTemp.SubItems(3)) = UCase(tvwBrowse.SelectedItem.Key) Then
            Set itm = lvwBrowse.ListItems.Add(, , itmTemp.Text)
            itm.SubItems(1) = itmTemp.SubItems(1)
            
        End If
    Next
End Sub

Private Function TagToUse(strFolder() As String, strFileName As String) As String
    Dim i As Integer
    
    For i = 0 To UBound(strFolder)
        'Glenn - check if for PLDA
        If UCase(Left(strFileName, 2)) = "PL" Then
            If UCase(Mid(strFileName, 3, 1)) = "I" Then
                TagToUse = "PLI"
                Exit Function
            ElseIf UCase(Mid(strFileName, 3, 1)) = "E" Then
                TagToUse = "PLE"
                Exit Function
            End If
        
        ElseIf UCase(Left(strFileName, 3)) = UCase(Left(strFolder(i), 3)) Then
            TagToUse = Left(strFileName, 3)
            Exit Function
        End If
        
    Next
End Function
