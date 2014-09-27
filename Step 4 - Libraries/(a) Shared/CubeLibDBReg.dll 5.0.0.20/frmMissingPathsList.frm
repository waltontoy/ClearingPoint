VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMissingPathsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Paths"
   ClientHeight    =   3705
   ClientLeft      =   3180
   ClientTop       =   4020
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   1355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   1355
   End
   Begin VB.CommandButton cmdAdvance 
      Caption         =   "&Advanced"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1355
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5440
      Begin VB.CommandButton cmdSetPath 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   720
         Width           =   1355
      End
      Begin MSComctlLib.ListView lvwMissingPaths 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "  "
            Object.Width           =   4674
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "The following resource paths must be set to proceed with using "
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSComctlLib.ImageList imgMissingPaths 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMissingPathsList.frx":0000
            Key             =   "DB"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMissingPathsList.frx":005E
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMissingPathsList.frx":00BC
            Key             =   "Folder"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMissingPathsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mvarOWnerForm As Object
    Dim mvarFormToShow As Object
    Dim mvarApplication As Object
    Dim mvarLocation As CLocations
    
    Dim strMissingPathsStream As String
    
    Dim blnOK As Boolean
    
Private Sub cmdAdvance_Click()
    mvarLocation.TriggerAdvanced
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If (ValidValues = True) Then
        blnOK = True
        
        Call SaveNewValues
        Unload Me
    End If
End Sub

Private Sub cmdSetPath_Click()
    Dim clsBrowse As CBrowse
    Dim blnBrowseResult As Boolean
    Dim strBuildNewTag As String
    Dim strNewValue As String
    Dim lngDetailCtr As Long
    Dim arrTagDetails
    Dim strDefaultPath As String
    
    Set clsBrowse = New CBrowse
    
    strBuildNewTag = ""
    arrTagDetails = Split(CStr(lvwMissingPaths.SelectedItem.Tag), "|")
    
    ' Get Default Path
    If Trim(CStr(arrTagDetails(UBound(arrTagDetails)))) = "" Then
        strDefaultPath = mvarApplication.Path
    Else
        strDefaultPath = CStr(arrTagDetails(UBound(arrTagDetails)))
    End If
    
    Select Case UCase(CStr(arrTagDetails(0)))
        Case "F"
            If clsBrowse.BrowseFile(Me, mvarApplication, strDefaultPath, CStr(arrTagDetails(4))) Then
                strNewValue = AddBackSlashOnPath(clsBrowse.Path) & clsBrowse.FileName
            Else
                Set clsBrowse = Nothing
                Exit Sub
            End If
        Case "P"
            If clsBrowse.BrowseFolder(Me, "Browse for folder to map with " & CStr(arrTagDetails(3)), strDefaultPath) Then
                strNewValue = clsBrowse.Path
            Else
                Set clsBrowse = Nothing
                Exit Sub
            End If
        Case "D"
            If clsBrowse.BrowseFile(Me, mvarApplication, strDefaultPath, CStr(arrTagDetails(3))) Then
                strNewValue = clsBrowse.Path
            Else
                Set clsBrowse = Nothing
                Exit Sub
            End If
    End Select
    
    For lngDetailCtr = 0 To UBound(arrTagDetails) - 1
        strBuildNewTag = strBuildNewTag & "|" & CStr(arrTagDetails(lngDetailCtr))
    Next lngDetailCtr
    
    strBuildNewTag = strBuildNewTag & "|" & Encrypt(strNewValue, KEY_ENCRYPT)
    strBuildNewTag = Mid(strBuildNewTag, 2)
    
    lvwMissingPaths.SelectedItem.Tag = strBuildNewTag
    If strNewValue <> " " Then
        lvwMissingPaths.SelectedItem.Bold = False
        lvwMissingPaths.SelectedItem.ListSubItems(1).Bold = False
        lvwMissingPaths.Refresh
    End If
        
    Set clsBrowse = Nothing
End Sub

Private Sub Form_Load()
    Dim itmListItem As Variant
    Dim itmListSubitem As Variant
    
    Dim arrMissingPathsStream
    Dim arrMissingPathsDetails
    Dim lngMissingPathsCtr As Long
    Dim strLvwIcon As String
    Dim strLvwText As String
    Dim lngFolderLocationIndex As Long
    
    Dim clsRegistry As CRegistry
    
    arrMissingPathsStream = Split(strMissingPathsStream, "|||||")
    
    For lngMissingPathsCtr = 0 To UBound(arrMissingPathsStream)
        arrMissingPathsDetails = Split(CStr(arrMissingPathsStream(lngMissingPathsCtr)), "|")
        ReDim Preserve arrMissingPathsDetails(6)
        
        Select Case UCase(CStr(arrMissingPathsDetails(0)))
            Case "P"
                strLvwIcon = "Folder"
            Case "F"
                strLvwIcon = "File"
            Case "D"
                strLvwIcon = "Database"
        End Select
        strLvwText = CStr(arrMissingPathsDetails(3))
        
        Set itmListItem = lvwMissingPaths.ListItems.Add(, , strLvwIcon)
        If Len(Trim(strLvwText)) > 30 Then
            Set itmListSubitem = lvwMissingPaths.ListItems(lvwMissingPaths.ListItems.Count).ListSubItems.Add(, , Trim(Mid(strLvwText, 1, 30) & "..."))
        Else
            Set itmListSubitem = lvwMissingPaths.ListItems(lvwMissingPaths.ListItems.Count).ListSubItems.Add(, , Trim(strLvwText & Space(30 - Len(strLvwText))))
        End If
                
        ' Append Empty Value to Accomodate Cancelling
        Select Case UCase(CStr(arrMissingPathsDetails(0)))
            Case "P"
                lngFolderLocationIndex = 4
            Case "F"
                lngFolderLocationIndex = 5
            Case "D"
                lngFolderLocationIndex = 6
        End Select
                        
        If Trim(CStr(arrMissingPathsDetails(lngFolderLocationIndex))) = "" Then
            itmListItem.Bold = True
            itmListSubitem.Bold = True
            
            arrMissingPathsStream(lngMissingPathsCtr) = CStr(arrMissingPathsStream(lngMissingPathsCtr)) & "| "
        Else
            arrMissingPathsStream(lngMissingPathsCtr) = CStr(arrMissingPathsStream(lngMissingPathsCtr)) & "|" & CStr(arrMissingPathsDetails(lngFolderLocationIndex))
        End If
                
        '--->Highlight file/folder locations not yet defined
        Set clsRegistry = New CRegistry
        If clsRegistry.GetRegistry(cpiLocalMachine, "FMS", arrMissingPathsDetails(1), arrMissingPathsDetails(2)) Then
            If clsRegistry.RegistryValue = "" Then
                itmListItem.Bold = True
                itmListSubitem.Bold = True
            End If
        End If
        Set clsRegistry = Nothing
        
        
        itmListItem.Tag = CStr(arrMissingPathsStream(lngMissingPathsCtr))
    Next lngMissingPathsCtr
    
    Label1.Caption = Label1.Caption & mvarApplication.ProductName & "."
    
    blnOK = False
End Sub

Public Function ShowMissingPaths(ByRef OwnerForm As Object, ByVal Application As Object, ByVal MissingPathsStream As String, Button As CLocations, ByVal UseAdvanceOption As Boolean) As Boolean
    Set mvarOWnerForm = OwnerForm
    Set mvarApplication = Application
    Set mvarLocation = Button
    
    blnOK = False
    
    strMissingPathsStream = MissingPathsStream
    
    cmdAdvance.Visible = UseAdvanceOption
    
    Set Me.Icon = mvarOWnerForm.Icon
    Me.Show vbModal
    
    ShowMissingPaths = blnOK
End Function

Public Sub SaveNewValues()
    Dim lngListItemCtr As Long
    Dim clsRegistry As CRegistry
    Dim arrListItemDetail
    
    Set clsRegistry = New CRegistry
    
    For lngListItemCtr = 1 To lvwMissingPaths.ListItems.Count
        arrListItemDetail = Split(CStr(lvwMissingPaths.ListItems(lngListItemCtr).Tag), "|")
        
        If Trim(CStr(arrListItemDetail(UBound(arrListItemDetail)))) <> "" Then
            If (Trim(CStr(arrListItemDetail(0))) <> "P") Then
                If (Trim(CStr(arrListItemDetail(0))) = "F") And (Left(Right(Trim(CStr(arrListItemDetail(UBound(arrListItemDetail)))), 4), 1) = ".") Then
                    '--->Tag = "F" (File)
                    clsRegistry.SaveRegistry cpiLocalMachine, mvarApplication.ProductName, CStr(arrListItemDetail(1)), CStr(arrListItemDetail(2)), CStr(arrListItemDetail(UBound(arrListItemDetail)))
                ElseIf (Trim(CStr(arrListItemDetail(0))) = "D") Then
                    '--->Tag = "D" (Database)
                    clsRegistry.SaveRegistry cpiLocalMachine, mvarApplication.ProductName, CStr(arrListItemDetail(1)), CStr(arrListItemDetail(2)), CStr(arrListItemDetail(UBound(arrListItemDetail)))
                End If
            Else
                '--->Tag = "P" (Path/Polder)
                clsRegistry.SaveRegistry cpiLocalMachine, mvarApplication.ProductName, CStr(arrListItemDetail(1)), CStr(arrListItemDetail(2)), CStr(arrListItemDetail(UBound(arrListItemDetail)))
            End If
        End If
    Next lngListItemCtr
    
    Set clsRegistry = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If (blnOK = False) Then
        If (ValidRegistryValues = True) Then
            Set mvarOWnerForm = Nothing
            Set mvarFormToShow = Nothing
            Set mvarApplication = Nothing
            Set mvarLocation = Nothing
        Else
            Set mvarOWnerForm = Nothing
            Set mvarFormToShow = Nothing
            Set mvarApplication = Nothing
            Set mvarLocation = Nothing
            'Cancel = True
        End If
    End If
End Sub

Private Sub lvwMissingPaths_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim varHitListItem As Variant
    Dim arrToolTip
    
    Set varHitListItem = lvwMissingPaths.HitTest(x, y)
    
    If Not varHitListItem Is Nothing Then
        arrToolTip = Split(CStr(varHitListItem.Tag), "|")
        If Trim(CStr(arrToolTip(UBound(arrToolTip)))) = "" Then
            lvwMissingPaths.ToolTipText = CStr(arrToolTip(3))
        Else
            lvwMissingPaths.ToolTipText = CStr(arrToolTip(UBound(arrToolTip)))
        End If
    End If
End Sub

Private Function ValidValues() As Boolean
        Dim lngListItemCtr As Long
    Dim arrListItemDetail
    Dim arrPathDetails
    
    Dim strPathName As String
    Dim strFileName As String
    
    Dim strType As String
    Dim enumDriveType As DriveType
    

    ValidValues = True
    If (lvwMissingPaths.ListItems.Count = 0) Then Exit Function
    
    For lngListItemCtr = 1 To lvwMissingPaths.ListItems.Count
        arrListItemDetail = Split(lvwMissingPaths.ListItems(lngListItemCtr).Tag, "|")
        
        strType = Trim(arrListItemDetail(0))
        strPathName = Trim(arrListItemDetail(UBound(arrListItemDetail)))

        If (strType = "P") Then
                If (GetDriveType(strPathName) <> cpiCDROM) Then

                    '--->Do if drive is fixed
                    If (Right(strPathName, 1) <> "\") Then
                        strPathName = strPathName & "\"
                    End If

                    On Error GoTo Err_DriveHandler
                    If (Dir(Trim(strPathName), vbDirectory) = ".") Then 'checking if with sub folders
                    On Error GoTo 0
                    
                        ValidValues = True
                    Else
                        If (GetDriveType(Trim(strPathName)) = cpiFIXEDDISK) Then
                            ValidValues = True
                        Else
                            MsgBox "The path '" & Trim(strPathName) & "' is an invalid database set folder location. " & vbCrLf & "Please set a new location for the " & _
                                   LCase(Trim(arrListItemDetail(3))) & ".", vbInformation, "ProfitPoint"

                            lvwMissingPaths.ListItems(lngListItemCtr).Selected = True
                            ValidValues = False

                            Exit For
                        End If
                    End If
                Else
                    '---> Do if drive is CDROM
                    MsgBox "The path '" & Trim(strPathName) & "' is an invalid database set folder location. " & vbCrLf & "Please set a new location for the " & _
                           LCase(Trim(arrListItemDetail(3))) & ".", vbInformation, "ProfitPoint"
                    
                    ValidValues = False
                End If

        ElseIf (strType = "F") And (strPathName <> "") And (Left(Right(strPathName, 4), 1) = ".") Then
            '--->Extract filename
            arrPathDetails = Split(strPathName, "\")
            strFileName = arrPathDetails(UBound(arrPathDetails))

            '--->Check existence of file (only if it is for 'Import')
            If (InStr(1, arrListItemDetail(3), "Import") > 0) Then
                If UCase(Dir(strPathName)) = UCase(strFileName) Then
                    ValidValues = True
                Else
                    MsgBox "The folder '" & Trim(strPathName) & "' does not exist. " & vbCrLf & "Please set a new location for the " & _
                           LCase(Trim(arrListItemDetail(3))) & ".", vbInformation, "ProfitPoint"

                    lvwMissingPaths.ListItems(lngListItemCtr).Selected = True
                    ValidValues = False

                    Exit For
                End If
            End If
        ElseIf (strType = "D") And (strPathName <> "") Then
            If (Dir(strPathName & "\TemplateFMS.mdb") = "TemplateFMS.mdb") Then
                ValidValues = True
            End If
        End If
    Next lngListItemCtr

    Exit Function

Err_DriveHandler:

    Select Case Err.Number
        Case 52 'Bad file name
            MsgBox "The path '" & Trim(strPathName) & "' is an invalid database set folder. " & vbCrLf & "Please set a new location for the " & _
                LCase(Trim(arrListItemDetail(3))) & ".", vbInformation & ".", "ProfitPoint"

            ValidValues = False
    End Select
End Function

Private Function ValidRegistryValues() As Boolean
'Checks validity of existing settings, i.e., values that were loaded from registry.
'In cases of non-existent file/directory settings, user is warned to change
'   the value for these settings later.
'FUNCTION IS APPLICABLE TO 'Cancel' BUTTON ONLY.
    
    Dim arrMissingPathsStream
    Dim arrMissingPathsDetails
    Dim arrPathDetails
    
    Dim strType As String
    Dim strFileName As String
    Dim strPathName As String
    
    Dim lngMissingPathsCtr As Long
    
    ValidRegistryValues = True
    
    arrMissingPathsStream = Split(strMissingPathsStream, "|||||")
    
    For lngMissingPathsCtr = 0 To UBound(arrMissingPathsStream)
        arrMissingPathsDetails = Split(arrMissingPathsStream(lngMissingPathsCtr), "|")
        
        strType = Trim(arrMissingPathsDetails(0))
        strPathName = Trim(arrMissingPathsDetails(UBound(arrMissingPathsDetails)))
        
        'Directory setting
        If (strType = "P") Then
            If (Trim(strPathName) = "") Then
                Exit For
            ElseIf (Dir(strPathName & "\", vbDirectory) = ".") Then
                ValidRegistryValues = True
            Else
                MsgBox strPathName & " does not exist." & vbCrLf & "Please set a new location for " & _
                       Trim(arrMissingPathsDetails(3)), vbInformation, "ProfitPoint"
                
                lvwMissingPaths.ListItems(lngMissingPathsCtr + 1).Selected = True
                ValidRegistryValues = False
                Exit For
            End If
        
        'File setting
        ElseIf (Left(Right(strPathName, 4), 1) = ".") Then
            '--->Extract filename
            arrPathDetails = Split(strPathName, "\")
            strFileName = arrPathDetails(UBound(arrPathDetails))
            
            '--->Check existence of file (only if it is for 'import')
            If (InStr(1, arrMissingPathsDetails(3), "Import") > 0) Then
                If (Dir(strPathName) = strFileName) Then
                    ValidRegistryValues = True
                Else
                    MsgBox strPathName & " does not exist." & vbCrLf & "Please do not forget to set a new value for " & _
                           Trim(arrMissingPathsDetails(3)) & " later.", vbInformation, "ProfitPoint"
                    
                    lvwMissingPaths.ListItems(lngMissingPathsCtr + 1).Selected = True
                    ValidRegistryValues = False
                    
                    Exit For
                End If
            End If
        End If
    Next lngMissingPathsCtr
    
End Function
