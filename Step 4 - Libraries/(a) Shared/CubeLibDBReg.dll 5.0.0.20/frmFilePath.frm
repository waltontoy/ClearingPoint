VERSION 5.00
Begin VB.Form frmFilePath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Database"
   ClientHeight    =   2625
   ClientLeft      =   3405
   ClientTop       =   3180
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   6735
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   840
         TabIndex        =   2
         Top             =   630
         Width           =   2610
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   5415
         TabIndex        =   5
         Tag             =   "426"
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   5415
         TabIndex        =   6
         Tag             =   "119"
         Top             =   1065
         Width           =   1200
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Top             =   240
         Width           =   2620
      End
      Begin VB.CommandButton cmdCreate 
         Height          =   315
         Left            =   3510
         Picture         =   "frmFilePath.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox txtDBPath 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   4420
      End
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   3480
         Pattern         =   "*.mdb"
         TabIndex        =   3
         Top             =   630
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "Path :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2175
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Folder :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Drive :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   495
      End
      Begin VB.Image imgDummy 
         Height          =   555
         Left            =   5670
         Top             =   1545
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Preview"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   2550
         Width           =   990
      End
      Begin VB.Image imgFooter 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   2880
         Left            =   1215
         Stretch         =   -1  'True
         Top             =   2550
         Width           =   4032
      End
   End
End
Attribute VB_Name = "frmFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strDefaultPath As String
    Dim strFilter As String
    Dim strTitleCaption As String
    Dim strSelectionOutput As String
    Dim enuImageType As ImageType
    
    Dim NewDBPath As String
    Dim OldDBPath As String
    
    'for image
    Dim fPath As String
    Dim intImgWidth As Single
    Dim intImgHeight As Single
    Dim intTmpW As Single
    Dim intTmpH As Single
    Dim intTmpWRatio As Single
    Dim intTmpHRatio As Single
    Dim intTmpRatio As Single
    Dim strMessage As String
    Dim blnFileSelected As Boolean
    
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub cmdCreate_Click()
    Dim lngPositionBSlash As Long
    
    lngPositionBSlash = InStrRev(Dir1.Path, "\", -1, vbTextCompare)
    Dim strTruncated As String
    strTruncated = Mid(Dir1.Path, 1, lngPositionBSlash - 1)
    If Mid(strTruncated, Len(strTruncated), 1) = ":" Then
        File1.Path = Mid(Dir1.Path, 1, lngPositionBSlash - 1) & "\"
    Else
        File1.Path = Mid(Dir1.Path, 1, lngPositionBSlash - 1)
    End If
    
    Dir1.Path = File1.Path
    txtDBPath.Tag = File1.Path & "\" & txtDBPath.Text
    txtDBPath.Text = Dir1.Path
End Sub

Private Sub cmdCreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub cmdOK_Click()
    If Trim(txtDBPath.Text) = "" Then
        MsgBox "Please choose a file from the list.", vbInformation, "Invalid Request"
        File1.SetFocus
        Exit Sub
    End If
        
    blnFileSelected = True
    
    Unload Me
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    txtDBPath.Tag = File1.Path & "\" & txtDBPath.Text
    txtDBPath.Text = Dir1.Path
End Sub

Private Sub Dir1_Click()
    File1.Path = Dir1.List(Dir1.ListIndex)
    txtDBPath.Tag = File1.Path & "\" & txtDBPath.Text
    txtDBPath.Text = File1.Path
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub Drive1_Change()
    On Error GoTo Error_Handler
    
    Dir1.Path = Drive1.Drive
    txtDBPath.Text = ""
    
Error_Handler:
    Select Case Err.Number
        Case 0  ' No Error
            ' Exit Normally
        Case 68
            MsgBox "Device unavailable.", vbInformation
            Drive1.ListIndex = 1
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Sub

Private Sub Drive1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub File1_Click()
    imgFooter.Visible = False
    imgFooter.Height = 2880
    imgFooter.Width = 4032
    
    
    strMessage = ""
    fPath = ""
    fPath = File1.Path & "\" & File1.FileName
    
    If Not enuImageType = imgFile Then
        imgDummy.Picture = Nothing
        imgDummy.Picture = LoadPicture(fPath)
        
        intTmpW = imgDummy.Width
        intTmpH = imgDummy.Height
        
        If intTmpW > 10080 And intTmpH > 7200 Then 'overdimension on width and height
            '''''strMessage = "The width ang height of the image exceeds the maximum allowable dimension."
            '''''strMessage = strMessage & vbCrLf & "Would you like to reduce the size of the image?"
            '''''If MsgBox(strMessage, vbInformation + vbYesNo) = vbYes Then
                imgFooter.Picture = Nothing
                
                intTmpWRatio = (10080 * 0.75 / (intTmpW))
                intTmpHRatio = (7200 * 0.75 / (intTmpH))
    
                If intTmpWRatio > intTmpHRatio Then
                    imgFooter.Height = imgFooter.Height
                    imgFooter.Width = imgFooter.Width * (1 - (intTmpWRatio - intTmpHRatio))
                ElseIf intTmpWRatio < intTmpHRatio Then
                    imgFooter.Width = imgFooter.Width
                    imgFooter.Height = imgFooter.Height * (1 - (intTmpHRatio - intTmpWRatio))
                ElseIf intTmpWRatio = intTmpHRatio Then
                    imgFooter.Width = imgFooter.Width
                    imgFooter.Height = imgFooter.Height
                End If
                        
                imgFooter.Picture = LoadPicture(fPath)
                imgFooter.Visible = True
                
                '''''intImageWindowW = imgFooter.Width
                '''''intImageWindowH = imgFooter.Height
                '''''blnGotImage = True
                Exit Sub
            '''''Else
            '''''    Exit Sub
            '''''End If
        ElseIf intTmpW > 10080 Then 'overdimension on width
            strMessage = "The width of the image exceeds the maximum allowable width."
            strMessage = strMessage & vbCrLf & "Would you like to reduce the size of the image?"
            If MsgBox(strMessage, vbInformation + vbYesNo) = vbYes Then
                imgFooter.Picture = Nothing
                
                intTmpWRatio = (10080 / intTmpW) * 0.75
                intTmpHRatio = intTmpWRatio * 0.75
                
                imgFooter.Width = imgFooter.Width
                imgFooter.Height = imgFooter.Height * intTmpHRatio
                
                imgFooter.Picture = LoadPicture(fPath)
                imgFooter.Visible = True
                
                '''''intImageWindowW = imgFooter.Width
                '''''intImageWindowH = imgFooter.Height
                '''''blnGotImage = True
                Exit Sub
            Else
                Exit Sub
            End If
        ElseIf intTmpH > 7200 Then 'overdimension on height
            strMessage = "The height of the image exceeds the maximum allowable height."
            strMessage = strMessage & vbCrLf & "Would you like to reduce the size of the image?"
            If MsgBox(strMessage, vbInformation + vbYesNo) = vbYes Then
                imgFooter.Picture = Nothing
                
                intTmpHRatio = (7200 / intTmpH) * 0.75
                intTmpWRatio = intTmpHRatio * 0.75
                            
                imgFooter.Height = imgFooter.Height
                imgFooter.Width = imgFooter.Width * intTmpHRatio
                
                imgFooter.Picture = LoadPicture(fPath)
                imgFooter.Visible = True
                
                '''''intImageWindowW = imgFooter.Width
                '''''intImageWindowH = imgFooter.Height
                '''''blnGotImage = True
                Exit Sub
            Else
                Exit Sub
            End If
        Else
            imgFooter.Picture = Nothing
                
            intTmpWRatio = (intTmpW / 10080) / 0.75
            intTmpHRatio = (intTmpH / 7200) / 0.75
                        
            If intTmpWRatio >= 1 Or intTmpHRatio >= 1 Then
                If intTmpWRatio < intTmpHRatio Then
                    imgFooter.Height = imgFooter.Height
                    imgFooter.Width = imgFooter.Width * (1 - (intTmpHRatio - intTmpWRatio)) '(1 - (intTmpWRatio - intTmpHRatio))
                ElseIf intTmpWRatio > intTmpHRatio Then
                    imgFooter.Width = imgFooter.Width
                    imgFooter.Height = imgFooter.Height * (1 - (intTmpWRatio - intTmpHRatio)) '(1 - (intTmpHRatio - intTmpWRatio))
                ElseIf intTmpWRatio = intTmpHRatio Then
                    imgFooter.Width = imgFooter.Width
                    imgFooter.Height = imgFooter.Height
                End If
            Else
                imgFooter.Width = imgFooter.Width * (intTmpWRatio)
                imgFooter.Height = imgFooter.Height * (intTmpHRatio)
            End If
                
            imgFooter.Picture = LoadPicture(fPath)
            imgFooter.Visible = True
            '''''intImageWindowW = imgFooter.Width
            '''''intImageWindowH = imgFooter.Height
            '''''blnGotImage = True
            
        End If
    End If
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
        
    Me.Caption = strTitleCaption
    
    If Trim(strDefaultPath) = "" Then
        Drive1.ListIndex = 1
        Dir1.Path = Drive1.Drive
    Else
        Drive1.Drive = Mid(strDefaultPath, 1, 2)
        Dir1.Path = strDefaultPath
    End If
    File1.Path = Dir1.List(Dir1.ListIndex)
    
    If enuImageType = imgFile Then
        File1.Pattern = strFilter
        File1.Refresh
                        
        Me.Height = 3005
        Frame1.Height = 2550
    Else
        Select Case enuImageType
            Case imgUnknownImage
                File1.Pattern = strFilter
            Case imgJPG
                File1.Pattern = "*.jpg;*.JPG"
            Case imgIcon
                File1.Pattern = "*.ico;*.ICO"
            Case imgBitmap
                File1.Pattern = "*.bmp;*.BMP"
            Case imgAll
                File1.Pattern = "*.jpg;*.JPG;*.ico;*.ICO;*.bmp;*.BMP"
        End Select
        
        Me.Height = 6050
        Frame1.Height = 5580
    End If
    
    File1.Refresh
    
    txtDBPath.Enabled = True
    txtDBPath.Locked = False
  
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strSelectionOutput = ""
    strSelectionOutput = strSelectionOutput & "*****" & IIf(blnFileSelected, "True", "False")
    strSelectionOutput = strSelectionOutput & "*****" & IIf(blnFileSelected, txtDBPath.Text, "-1")
    strSelectionOutput = strSelectionOutput & "*****" & IIf(blnFileSelected, File1.List(File1.ListIndex), "-1")
    strSelectionOutput = strSelectionOutput & "*****" & IIf(blnFileSelected, CStr(imgDummy.Height), "-1")
    strSelectionOutput = strSelectionOutput & "*****" & IIf(blnFileSelected, CStr(imgDummy.Width), "-1")
    strSelectionOutput = Mid(strSelectionOutput, 6)
End Sub

Private Sub txtDBPath_Change()
    On Error GoTo ErrorHandler
    
    File1.Path = txtDBPath.Text
    File1.ForeColor = &H80000012
    
    
    Exit Sub
    
ErrorHandler:
    File1.ForeColor = &H80000005
End Sub

Private Sub txtDBPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape            '----->Escape
            cmdCancel_Click
            Exit Sub
    End Select
End Sub

Public Function SelectFile(ByVal OwnerForm As Object, ByVal TitleCaption As String, ByVal DefaultPath As String, ByVal Filter As String, Optional ByVal TypeOfImage As ImageType = imgFile)
    strDefaultPath = DefaultPath
    strFilter = Filter
    strTitleCaption = TitleCaption
    blnFileSelected = False
    enuImageType = TypeOfImage
                    
    Set Me.Icon = OwnerForm.Icon
    Me.Show vbModal
    
    SelectFile = strSelectionOutput
End Function
