VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormatColumns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Columns"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4680
      TabIndex        =   4
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   350
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   3735
      Begin VB.OptionButton optRight 
         Caption         =   "Right"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Center"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Alignment:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "A&vailable Fields:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmFormatColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private conFormat As ADODB.Connection
Private clsFormat As CGrid
Private arrFields
Private arrAlignments
Private arrWidths


Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
        
    Dim lngCtr As Long
    
    '>> save column formats
    clsFormat.Widths = ""
    clsFormat.Alignments = ""
    
    'CSCLP-787
    For lngCtr = 0 To UBound(arrFields)
        If lngCtr <= UBound(arrWidths) Then
            clsFormat.Widths = clsFormat.Widths & CDbl(arrWidths(lngCtr)) & "*****"
        Else
            clsFormat.Widths = clsFormat.Widths & CDbl(1200) & "*****"
        End If
        
        If lngCtr <= UBound(arrAlignments) Then
            clsFormat.Alignments = clsFormat.Alignments & Trim(arrAlignments(lngCtr)) & "*****"
        Else
            clsFormat.Widths = clsFormat.Widths & "Left" & "*****"
        End If
    Next
    
    clsFormat.Widths = Mid(clsFormat.Widths, 1, Len(clsFormat.Widths) - 5)
    clsFormat.Alignments = Mid(clsFormat.Alignments, 1, Len(clsFormat.Alignments) - 5)
    
    If frmEditView.Visible = False Then
        clsFormat.UpdateGridSetting conFormat
        clsFormat.DataChanged = True
    End If

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim rstFields As ADODB.Recordset
    
    Dim lngCtr As Long
    Dim strCommandText As String
    
    '>> load column formats
    arrFields = Split(clsFormat.DVCIDs, "*****")
    arrAlignments = Split(clsFormat.Alignments, "*****")
    arrWidths = Split(clsFormat.Widths, "*****")
        
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_ID, "
    strCommandText = strCommandText & "DVC_FieldSource, "
    strCommandText = strCommandText & "DVC_FieldAlias, "
    strCommandText = strCommandText & "DVC_DataType "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "DVC_ID IN (" & Replace(clsFormat.DVCIDs, "*****", ",") & ") "
    
    ADORecordsetOpen strCommandText, conFormat, rstFields, adOpenKeyset, adLockOptimistic
    'Call RstOpen(strCommandText, conFormat, rstFields, adOpenKeyset, adLockReadOnly)
    
    '>> load available fields to listview
    For lngCtr = 0 To UBound(arrFields)
        If rstFields.RecordCount > 0 Then
            rstFields.MoveFirst
            rstFields.Find "DVC_ID = " & Val(arrFields(lngCtr)), , adSearchForward, 0
            If Not rstFields.EOF Then
                lvwFields.ListItems.Add lvwFields.ListItems.Count + 1, , rstFields!DVC_FieldAlias
            End If
        End If
    Next
    
    Call ADORecordsetClose(rstFields)
    
    '>> load format of selected item
    Call lvwFields_ItemClick(lvwFields.SelectedItem)
    
    lvwFields.Refresh
    
End Sub


Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)

    Set conFormat = ADOConnection
    Set clsFormat = GridProps
    
    Set Me.Icon = Window.Icon
    
    Me.Show vbModal
    
    Set GridProps = clsFormat
    Set ADOConnection = conFormat
    
    Set clsFormat = Nothing
    Set conFormat = Nothing
    
End Sub



Private Sub lvwFields_ItemClick(ByVal Item As MSComctlLib.ListItem)

    '>> load column format of selected column
    txtWidth.Text = arrWidths(Item.Index - 1)
    
    Select Case Trim(UCase(arrAlignments(Item.Index - 1)))
        Case "LEFT"
            optLeft.Value = True
            optCenter.Value = False
            optRight.Value = False
            
        Case "CENTER"
            optLeft.Value = False
            optCenter.Value = True
            optRight.Value = False
        
        Case "RIGHT"
            optLeft.Value = False
            optCenter.Value = False
            optRight.Value = True
            
    End Select
    
End Sub


Private Sub optCenter_Click()

    '>> update alignment of selected column
    If optCenter.Value = True Then
        arrAlignments(lvwFields.SelectedItem.Index - 1) = "CENTER"
    End If

End Sub

Private Sub optLeft_Click()

    '>> update alignment of selected column
    If optLeft.Value = True Then
        arrAlignments(lvwFields.SelectedItem.Index - 1) = "LEFT"
    End If
    
End Sub

Private Sub optRight_Click()
    
    '>> update alignment of selected column
    If optRight.Value = True Then
        arrAlignments(lvwFields.SelectedItem.Index - 1) = "RIGHT"
    End If

End Sub


Private Sub txtWidth_Validate(Cancel As Boolean)

    '>> check if width entered is valid
    If IsNumeric(txtWidth.Text) And Val(txtWidth.Text) >= 0 Then
        arrWidths(lvwFields.SelectedItem.Index - 1) = CDbl(txtWidth.Text)
    Else
        MsgBox "Please enter a valid numeric value for column width.", vbInformation, "Cubepoint Library"
        txtWidth.Text = arrWidths(lvwFields.SelectedItem.Index - 1)
        SendKeysEx "{HOME}+{END}"
        Cancel = True
    End If
    
End Sub
