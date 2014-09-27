VERSION 5.00
Begin VB.Form frmEditView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize View"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDummytextBox 
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Current View"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdShowFields 
         Caption         =   "&Fields..."
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdFormatCol 
         Caption         =   "Format &Columns..."
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAutoFormat 
         Caption         =   "&Automatic Formatting..."
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton cmdSettings 
         Caption         =   "&Other Settings..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Fi&lter..."
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Sort..."
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdGroupBy 
         Caption         =   "&Group By..."
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblFields 
         Caption         =   "Fields"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   390
         Width           =   4335
      End
      Begin VB.Label lblFormatCol 
         Caption         =   "Specify the display formats for each field"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   3270
         Width           =   4335
      End
      Begin VB.Label lblAutoFormat 
         Caption         =   "User defined fonts on each record"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   2790
         Width           =   4335
      End
      Begin VB.Label lblSettings 
         Caption         =   "Other Settings"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   2310
         Width           =   4335
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1830
         Width           =   4335
      End
      Begin VB.Label lblSort 
         Caption         =   "Sort"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   1350
         Width           =   4335
      End
      Begin VB.Label lblGroupBy 
         Caption         =   "Group By"
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   870
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmEditView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private conViewSettings As ADODB.Connection
Private clsGrid As CGrid
Private objMain As Object

Private m_blnCancelled As Boolean

Private Sub cmdAutoFormat_Click()

    frmAutoFormat.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub cmdCancel_Click()

    m_blnCancelled = True
    Unload Me
    
End Sub

Private Sub cmdFilter_Click()

    frmFilter.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub cmdFormatCol_Click()

    frmFormatColumns.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub cmdGroupBy_Click()

    frmGroupBy.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub cmdOK_Click()
    
    m_blnCancelled = False
    
    g_blnFromEditViewForm = True
    
    Unload Me
    
End Sub

Private Sub cmdReset_Click()
    
    Dim rstDVC As ADODB.Recordset
    Dim strCommandText As String
    
    '>> clear all view settings
    
    '>> clear groupings
    clsGrid.GroupHeaders = ""
    '>> clear sorting
    clsGrid.Sort = ""
    
    '>> remove all filters
    strCommandText = vbNullString
    strCommandText = strCommandText & " DELETE "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Filter "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " UVC_ID = " & clsGrid.UVC_ID
    
    ExecuteNonQuery conViewSettings, strCommandText
    'conViewSettings.Execute strCommandText
    
    '>> load default view settings
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " DefaultViewColumns "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " TView_ID = " & clsGrid.TView_ID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " DVC_Default = True "
    
    ADORecordsetOpen strCommandText, conViewSettings, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = conViewSettings.Execute(strCommandText)
    
    clsGrid.DVCIDs = ""
    clsGrid.Alignments = ""
    clsGrid.Widths = ""
    clsGrid.RequiredFields = ""
    
    Do While Not rstDVC.EOF
        If rstDVC!DVC_Requirement <> 1 Then
            clsGrid.DVCIDs = clsGrid.DVCIDs & rstDVC!DVC_ID & "*****"
            clsGrid.Alignments = clsGrid.Alignments & rstDVC!DVC_Alignment & "*****"
            clsGrid.Widths = clsGrid.Widths & rstDVC!DVC_Width & "*****"
        Else
            clsGrid.RequiredFields = clsGrid.RequiredFields & rstDVC!DVC_ID & "*****"
        End If
        
        rstDVC.MoveNext
    Loop
    
    If Len(clsGrid.DVCIDs) > 5 Then
        clsGrid.DVCIDs = Mid(clsGrid.DVCIDs, 1, Len(clsGrid.DVCIDs) - 5)
    End If
    If Len(clsGrid.Alignments) > 5 Then
        clsGrid.Alignments = Mid(clsGrid.Alignments, 1, Len(clsGrid.Alignments) - 5)
    End If
    If Len(clsGrid.Widths) > 5 Then
        clsGrid.Widths = Mid(clsGrid.Widths, 1, Len(clsGrid.Widths) - 5)
    End If
    If Len(clsGrid.RequiredFields) > 5 Then
        clsGrid.RequiredFields = Mid(clsGrid.RequiredFields, 1, Len(clsGrid.RequiredFields) - 5)
    End If
    
    '>> clear autformat rules
    strCommandText = vbNullString
    strCommandText = strCommandText & "DELETE "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & clsGrid.UVC_ID
    
    ExecuteNonQuery conViewSettings, strCommandText
    'conViewSettings.Execute strCommandText
    
    cmdReset.Enabled = False
    
    Call ADORecordsetClose(rstDVC) ' by hobbes 10/18/2005
    
End Sub

Private Sub cmdSettings_Click()

    frmSettings.Show vbModal
    
End Sub

Private Sub cmdShowFields_Click()

    frmShowFields.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub cmdSort_Click()

    frmSort.ShowForm objMain, clsGrid, conViewSettings
    
End Sub

Private Sub LoadLabels()
        
    Dim rstDVC As ADODB.Recordset
    Dim rstFilter As ADODB.Recordset
    Dim lngCtr As Long
    Dim blnPutEllipsis As Boolean
    Dim strCommandText As String
    Dim arrID
    Dim arrGroups
    Dim arrSort
    
    '>> load label caption of view options
    arrID = Split(clsGrid.DVCIDs, "*****")
    arrGroups = Split(clsGrid.GroupHeaders, "*****")
    arrSort = Split(clsGrid.Sort, "*****")
    
    lblFields.Caption = ""
    For lngCtr = 0 To UBound(arrID)
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "DefaultViewColumns "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "DVC_ID = " & Val(arrID(lngCtr))
           
        ADORecordsetOpen strCommandText, conViewSettings, rstDVC, adOpenKeyset, adLockOptimistic
        'Set rstDVC = conViewSettings.Execute(strCommandText)
    
        If Not (rstDVC.EOF And rstDVC.BOF) Then
            rstDVC.MoveFirst
            
            lblFields.Caption = lblFields.Caption & rstDVC!DVC_FieldAlias & ", "
            rstDVC.MoveNext
        End If
        
        'Set rstDVC = Nothing
        Call ADORecordsetClose(rstDVC) ' by hobbes 10/18/2005
    Next
    
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "DefaultViewColumns "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "DVC_ID IN (" & Replace(clsGrid.DVCIDs, "*****", ", ") & ")"
        
    ADORecordsetOpen strCommandText, conViewSettings, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = conViewSettings.Execute(strCommandText)

    '>> DVC caption
    If Len(lblFields.Caption) > 2 Then
        lblFields.Caption = Mid(lblFields.Caption, 1, Len(lblFields.Caption) - 2)
    End If
    
    Do While TextWidth(lblFields.Caption) > 4000
        blnPutEllipsis = True
        lblFields.Caption = Mid(lblFields.Caption, 1, Len(lblFields.Caption) - 1)
    Loop
    
    If blnPutEllipsis = True Then
        lblFields.Caption = lblFields.Caption & "..."
        blnPutEllipsis = False
    End If
    
    lblGroupBy.Caption = ""
    
    For lngCtr = 0 To UBound(arrGroups) Step 3
        rstDVC.MoveFirst
        Do While Not rstDVC.EOF
            If rstDVC!DVC_ID = Val(arrGroups(lngCtr)) Then
                If Val(arrGroups(lngCtr + 1)) = jgexSortAscending Then
                    lblGroupBy.Caption = lblGroupBy.Caption & rstDVC!DVC_FieldAlias & " (ascending), "
                Else
                    lblGroupBy.Caption = lblGroupBy.Caption & rstDVC!DVC_FieldAlias & " (descending), "
                End If
            End If
            rstDVC.MoveNext
        Loop
    Next
    
    '>> grouping caption
    If Len(lblGroupBy.Caption) > 2 Then
        lblGroupBy.Caption = Mid(lblGroupBy.Caption, 1, Len(lblGroupBy.Caption) - 2)
    Else
        lblGroupBy.Caption = "None"
    End If
    
    Do While TextWidth(lblGroupBy.Caption) > 4000
        blnPutEllipsis = True
        lblGroupBy.Caption = Mid(lblGroupBy.Caption, 1, Len(lblGroupBy.Caption) - 1)
    Loop
    
    If blnPutEllipsis = True Then
        lblGroupBy.Caption = lblGroupBy.Caption & "..."
        blnPutEllipsis = False
    End If

    lblSort.Caption = ""
    
    '>> sorting caption
    For lngCtr = 0 To UBound(arrSort) Step 2
        If Val(arrSort(lngCtr + 1)) = jgexSortAscending Then
            lblSort.Caption = lblSort.Caption & Trim(arrSort(lngCtr)) & " (ascending), "
        Else
            lblSort.Caption = lblSort.Caption & Trim(arrSort(lngCtr)) & " (descending), "
        End If
    Next
    
    If Len(lblSort.Caption) > 2 Then
        lblSort.Caption = Mid(lblSort.Caption, 1, Len(lblSort.Caption) - 2)
    Else
        lblSort.Caption = "None"
    End If
    
    Do While TextWidth(lblSort.Caption) > 4000
        blnPutEllipsis = True
        lblSort.Caption = Mid(lblSort.Caption, 1, Len(lblSort.Caption) - 1)
    Loop
    
    If blnPutEllipsis = True Then
        lblSort.Caption = lblSort.Caption & "..."
        blnPutEllipsis = False
    End If

    '>> filter caption
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Filter "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & clsGrid.UVC_ID & " "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "Filter_Type, Filter_ID "

    ADORecordsetOpen strCommandText, conViewSettings, rstFilter, adOpenKeyset, adLockOptimistic
    'Set rstFilter = conViewSettings.Execute(strCommandText)
        
    lblFilter.Caption = ""
    
    If Not rstFilter.EOF Then
        Do While Not rstFilter.EOF
            If rstFilter!Filter_Type = 1 Then
                rstDVC.MoveFirst
                Do While Not rstDVC.EOF
                    If rstDVC!DVC_FieldSource = rstFilter!Filter_Field Then
                        lblFilter.Caption = lblFilter.Caption & rstDVC!DVC_FieldAlias & " containing " & rstFilter!Filter_Value & ", "
                        Exit Do
                    End If
                    
                    rstDVC.MoveNext
                Loop
            Else
                rstDVC.MoveFirst
                Do While Not rstDVC.EOF
                    If rstDVC!DVC_FieldSource = rstFilter!Filter_Field Then
                        lblFilter.Caption = lblFilter.Caption & rstDVC!DVC_FieldAlias & " " & _
                                            Condition(rstFilter!Filter_Operator, _
                                            rstFilter!Filter_DataType, rstFilter!Filter_Value) & ", "
                        
                        
                    End If
                    
                    rstDVC.MoveNext
                Loop
            End If
            rstFilter.MoveNext
        Loop
        
        lblFilter.Caption = Mid(lblFilter.Caption, 1, Len(lblFilter.Caption) - 2)
        
        Do While TextWidth(lblFilter.Caption) > 4000
            blnPutEllipsis = True
            lblFilter.Caption = Mid(lblFilter.Caption, 1, Len(lblFilter.Caption) - 1)
        Loop
        
        If blnPutEllipsis = True Then
            lblFilter.Caption = lblFilter.Caption & "..."
            blnPutEllipsis = False
        End If
    Else
        lblFilter.Caption = "Off"
    End If
    
    'Set rstDVC = Nothing
    ' by hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    Call ADORecordsetClose(rstFilter)
    
End Sub

Private Function Condition(OperatorType As Long, DataType As Long, FilterValue As String) As String

    '>> caption based on the selected operator
    Select Case OperatorType
        Case 1
            Condition = "containing " & FilterValue
        Case 2
            Select Case DataType
                Case 1
                    Condition = "is exactly " & FilterValue
                Case 2
                    Condition = "is equal to " & FilterValue
                Case 3
                    Condition = " is on " & FilterValue
                Case 4
                
            End Select
        Case 3
            Select Case DataType
                Case 1
                    Condition = "doesn't contain " & FilterValue
                Case 2
                    Condition = "is less than or equal to " & FilterValue
                Case 3
                    Condition = "is on or before " & FilterValue
                Case 4
                
            End Select
        Case 4
            Select Case DataType
                Case 1
                    Condition = "is empty"
                Case 2
                    Condition = "is greater than or equal to " & FilterValue
                Case 3
                    Condition = "is on or after " & FilterValue
                Case 4
                
            End Select
        Case 5
            Select Case DataType
                Case 1
                    Condition = "is not empty"
                Case 2, 3
                    Condition = "is between " & Replace(Replace(FilterValue, "<", ""), ">", "")
                
                Case 4
                
            End Select
    
    End Select
    
End Function
Public Sub ShowForm(ByRef Window As Object, ByRef GridProps As CGrid, ByRef NavPane As CNavigationPane, _
                    ByRef ADOConnection As ADODB.Connection, ByVal ViewCaption As String, ByRef CustomizedView As Boolean)
    
    Set clsGrid = GridProps
    Set conViewSettings = ADOConnection
    Set objMain = Window

    Me.Caption = "Customize View: " & ViewCaption
    Set Me.Icon = Window.Icon
    
    conViewSettings.BeginTrans
    
    'Disable Automatic Formatting - Edwin - Nov 8, 2007
    Me.cmdAutoFormat.Enabled = False
    
    Me.Show vbModal
    
    CustomizedView = Not m_blnCancelled
    
    If m_blnCancelled = False Then
        conViewSettings.CommitTrans
        Set Window = objMain
        Set GridProps = clsGrid
        Set ADOConnection = conViewSettings
    Else
        conViewSettings.RollbackTrans
        GridProps.SelectGridSetting NavPane, ADOConnection, False
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Activate()

    Call LoadLabels
    
    If clsGrid.CardView = True Then
        cmdGroupBy.Enabled = False
    Else
        cmdGroupBy.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()

    m_blnCancelled = True
    
End Sub
