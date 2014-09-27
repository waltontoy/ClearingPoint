VERSION 5.00
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.70#0"; "CO2FCC~1.OCX"
Begin VB.Form frmPane 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   1455
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBGroundTree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   3855
      TabIndex        =   6
      Top             =   2040
      Width           =   3855
      Begin VB.Frame fraViews 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         TabIndex        =   8
         Top             =   3960
         Width           =   3855
         Begin VB.OptionButton optViews 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblCustomize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Customize Current View..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   600
            TabIndex        =   3
            Top             =   720
            Width           =   1935
         End
      End
      Begin SSActiveTreeView.SSTree tvwMain 
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6588
         _Version        =   65536
         Appearance      =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   3
         Indentation     =   315
         LoadStyleRoot   =   1
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "(None)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.ShortcutCaption sccViews 
         Height          =   255
         Left            =   20000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3855
         _Version        =   589894
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   6
         Caption         =   "Current View"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin SSActiveTreeView.SSTree tvwShortcut 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2143
      _Version        =   65536
      Appearance      =   0
      LabelEdit       =   1
      ScrollStyle     =   1
      Style           =   1
      IndentationStyle=   1
      Indentation     =   315
      LoadStyleRoot   =   1
      PictureBackgroundUseMask=   0   'False
      HasFont         =   -1  'True
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "(None)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption sccCaption 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3855
      _Version        =   589894
      _ExtentX        =   6800
      _ExtentY        =   661
      _StockProps     =   6
      Caption         =   "CAPTION"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption sccShortcut 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   3855
      _Version        =   589894
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   6
      Caption         =   "Favorite Folders"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.ShortcutCaption sccFolders 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3855
      _Version        =   589894
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   6
      Caption         =   "All Folders"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objNodeClicked As SSNode

Public clsNodes As INavigationPane
'Public clsNavigationPane As CNavigationPane

Private m_conTreeview As ADODB.Connection 'from OnDemandPrepare - joy
Private m_rstTreeview As ADODB.Recordset

Private m_enuKeypressed As KeyCodeConstants
Private m_enuMouseButton As MouseButtonConstants

Dim m_blnBeginDrag As Boolean

Private Sub Form_Load()
    
    '>> set scrollbar properties and dimensions
'    SetScrollBar picBGroundTree.hwnd, vbVertical, False

End Sub

Private Sub Form_Resize()

    '>> resize pane controls
    Call PaneResize
    
End Sub

Private Sub fraViews_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Screen.MousePointer = vbDefault
    lblCustomize.FontUnderline = False
    
End Sub



Private Sub lblCustomize_Click()
    
    
    If clsNodes.CustomizeView = True Then
        clsNodes.TriggerNodeClickEvent tvwMain.SelectedItem, True, True
    End If
    
    lblCustomize.FontUnderline = False
    
End Sub

Private Sub lblCustomize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lblCustomize.FontUnderline = True
    
End Sub

Private Sub optViews_Click(Index As Integer)
    
    clsNodes.TriggerViewOptionClick optViews(Index).Caption, Index
    clsNodes.TriggerNodeClickEvent tvwMain.SelectedItem, True, True
    
End Sub


Private Sub optViews_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    lblCustomize.FontUnderline = False
    
End Sub

Private Sub picBGroundTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub picBGroundTree_Resize()
    
    '>> adjust scrollbar dimensions upon resizing the pane
'    AdjustScrollInfo picBGroundTree.hwnd
    
End Sub


Private Sub tvwMain_AfterLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)
    Dim conDBConnection As ADODB.Connection
    
    Dim strTreeTag As String
    Dim blnCancel As Boolean
    
    Dim lngCtr As Long
    
    
    strTreeTag = tvwMain.Tag
    
    Call clsNodes.AfterLabelEdit(blnCancel, NewString, TREEVIEW_MAIN, conDBConnection)
    
    If blnCancel = False Then
        Call EditMainNode(NewString, Val(tvwMain.SelectedItem.Key), conDBConnection)
        
        If (tvwShortcut.Nodes.Count > 0) Then
            'tvwShortcut.Nodes(tvwMain.SelectedItem.key).Text = NewString
            
            For lngCtr = 1 To tvwShortcut.Nodes.Count
                If (tvwShortcut.Nodes(lngCtr).Key = tvwMain.SelectedItem.Key) Then
                    tvwShortcut.Nodes(tvwMain.SelectedItem.Key).Text = NewString
                End If
            Next lngCtr
        End If
    End If
    
    tvwMain.Nodes(tvwMain.SelectedItem.Key).Text = NewString
    
    Cancel = blnCancel
    
    If Cancel = False Then
        tvwMain.Tag = strTreeTag
    End If
    
End Sub

Private Sub tvwMain_BeforeNodeClick(Node As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean)

    Dim blnCancel As Boolean
    
    Call clsNodes.BeforeNodeClick(Node, blnCancel)
    
    Cancel = blnCancel
    
End Sub

Private Sub tvwMain_Click()

    Call clsNodes.Click
    
End Sub

Private Sub tvwMain_Collapse(Node As SSActiveTreeView.SSNode)
    
    If m_blnBeginDrag = True Then
        m_blnBeginDrag = False
        tvwMain.Drag vbEndDrag
    End If

    Call clsNodes.Collapse(Node)
        
End Sub

Private Sub tvwMain_EscapeLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)

    Dim blnCancel As Boolean
    
    Call clsNodes.EscapeLabelEdit(blnCancel, NewString)
    
    Cancel = blnCancel
    
End Sub

Private Sub tvwMain_Expand(Node As SSActiveTreeView.SSNode)

    If m_blnBeginDrag = True Then
        m_blnBeginDrag = False
        tvwMain.Drag vbEndDrag
    End If

    If Not clsNodes Is Nothing Then
        Call clsNodes.Expand(Node)
    End If
    
End Sub

Private Sub tvwMain_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nodParentNode As SSNode
    Dim blnNextNodeFound As Boolean
    
    '>> navigate pane with the use of a keyboard
    If Not tvwMain.SelectedItem Is Nothing Then
        Select Case KeyCode
            Case vbKeyUp
                '>> move focus to the 'Favorites' tree if the selected item in the Folder tree is the top node
                If tvwMain.SelectedItem = tvwMain.SelectedItem.Root And tvwMain.SelectedItem.Previous Is Nothing Then
                    Set tvwMain.DropHighlight = Nothing
                    tvwShortcut.SelectedItem = tvwShortcut.Nodes(tvwShortcut.Nodes.Count)
                    tvwShortcut.SetFocus
                End If
                
            Case vbKeyDown
                '>> move focus to the view options if the selected item is the last node in the folder tree
                
                If Not tvwMain.SelectedItem.Next Is Nothing Then
                    blnNextNodeFound = True
                ElseIf Not tvwMain.SelectedItem.Child Is Nothing Then
                    If tvwMain.Nodes(tvwMain.SelectedItem.Child.Key).Visible = True Then
                        blnNextNodeFound = True
                    End If
                End If
                    
                If blnNextNodeFound = False Then
                    Set nodParentNode = tvwMain.SelectedItem.Parent
                    
                    Do While (Not nodParentNode Is Nothing)
                        If Not nodParentNode.Next Is Nothing Then
                            blnNextNodeFound = True
                            Exit Do
                        Else
                            blnNextNodeFound = False
                            Set nodParentNode = nodParentNode.Parent
                        End If
                    Loop
                End If
                
                If blnNextNodeFound = False Then
                    Set tvwMain.DropHighlight = Nothing
                    On Error Resume Next
                    optViews(1).SetFocus
                    On Error GoTo 0
                End If
            
            Case vbKeyReturn
                
                If Not tvwMain.SelectedItem Is Nothing Then
                    tvwMain.DropHighlight = tvwMain.SelectedItem
                
                '    '>> load view options for the selected node & load grid records and setting
                    clsNodes.TriggerNodeClickEvent tvwMain.SelectedItem
                
                    '>> resize pane based on the number of view options
                    Call PaneResize
                    
                    tvwMain.SetFocus
                    
                    Call clsNodes.NodeClick(tvwMain.SelectedItem)
                End If
                
        End Select
    End If
    
    m_enuKeypressed = KeyCode
    
End Sub

Private Sub tvwMain_KeyUp(KeyCode As Integer, Shift As Integer)

    Call clsNodes.KeyUp(KeyCode, Shift)
    
End Sub

Private Sub tvwMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nodSelected As SSNode
    
    If Button = vbLeftButton Then
        '>> enable dragging of node
        'If Not tvwMain.SelectedItem Is Nothing Then
        '    m_blnBeginDrag = True
        'End If

        Set nodSelected = tvwMain.HitTest(x, y)
        If Not nodSelected Is Nothing Then
            Set m_objNodeClicked = nodSelected
            'm_blnBeginDrag = True
            'tvwMain_NodeClick nodSelected
        Else
            Set m_objNodeClicked = Nothing
        End If

    Else
        Set m_objNodeClicked = Nothing
    End If
    
    Call clsNodes.MouseDown(Button, Shift, x, y)
            
    m_enuMouseButton = Button
End Sub

Private Sub tvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

        Dim nodSelected As SSNode
    
    Screen.MousePointer = vbDefault

    'If Button = vbLeftButton And m_blnBeginDrag = True Then
    If Button = vbLeftButton Then
        '>> show drag icon upon moving
        Set nodSelected = tvwMain.HitTest(x, y)
        If Not nodSelected Is Nothing Then
            If Not m_objNodeClicked Is Nothing Then
                If nodSelected <> m_objNodeClicked Then
                    If Not m_objNodeClicked Is Nothing Then
                        tvwMain.DragIcon = m_objNodeClicked.CreateDragImage
    '                tvwMain.DragIcon = nodSelected.CreateDragImage
    '            Else
    '                tvwMain.DragIcon = tvwMain.SelectedItem.CreateDragImage
                        m_blnBeginDrag = True
                        tvwMain.Drag vbBeginDrag
                    End If
                End If
            End If
'            tvwMain.Drag vbBeginDrag
        Else
            If Not m_objNodeClicked Is Nothing Then
                tvwMain.DragIcon = m_objNodeClicked.CreateDragImage
                m_blnBeginDrag = True
                tvwMain.Drag vbBeginDrag
            End If
        End If
'    Else
'        m_blnBeginDrag = False
'        tvwMain.Drag vbEndDrag
'    End If
    Else
        Set m_objNodeClicked = Nothing
        m_blnBeginDrag = False
    End If

    lblCustomize.FontUnderline = False
    
    Call clsNodes.MouseMove(Button, Shift, x, y)
    
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If m_blnBeginDrag = True Then
        '>> disable dragging of node
        Set m_objNodeClicked = Nothing
        m_blnBeginDrag = False
        tvwMain.Drag vbEndDrag
    End If

    Call clsNodes.MouseUp(Button, Shift, x, y)
    
End Sub

Private Sub tvwMain_NodeClick(Node As SSActiveTreeView.SSNode)
    
    Screen.MousePointer = vbHourglass
    
    If m_enuMouseButton = vbLeftButton Then
        If m_enuKeypressed <> vbKeyDown And m_enuKeypressed <> vbKeyUp Then
            Set tvwShortcut.SelectedItem = Nothing
            Set tvwShortcut.DropHighlight = Nothing
        
            tvwMain.DropHighlight = Node
            tvwMain.SelectedItem = Node

        '    '>> load view options for the selected node & load grid records and setting
            clsNodes.TriggerNodeClickEvent Node

            '>> resize pane based on the number of view options
            Call PaneResize
        
        
        '    tvwMain.SelectedItem.EnsureVisible
            
'            tvwMain.SetFocus
            
            'Call clsNodes.NodeClick(Node)
        Else
            m_enuKeypressed = 0
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub tvwMain_OnDemandFetch(ByVal FetchBuffer As SSActiveTreeView.SSFetchBuffer)

    Dim blnFollowUpLicOk As Boolean
    Dim intIndent As Integer
    Dim lngCtr As Long
    Dim nodTreeview As SSNode
    Dim blnCancelTranslation As Boolean
    
    
    Call clsNodes.OnDemandFetch(FetchBuffer)
    
    ' get parent node's indentation
    If FetchBuffer.ParentNode Is Nothing Then
        intIndent = 0
    Else
        intIndent = FetchBuffer.ParentNode.Level
    End If
    
    ' position to begin fetch record
    If FetchBuffer.StartNode Is Nothing Then
        
        If (m_rstTreeview.EOF = False And m_rstTreeview.BOF = False) Then
            ' move to beginning or end of recordset
            If FetchBuffer.ReadPriorNodes Then
                m_rstTreeview.MoveLast
            Else
                m_rstTreeview.MoveFirst
            End If
        End If
        
    Else
    
        ' move to beginning of recordset
        If ((m_rstTreeview Is Nothing) = True) Then
        
            Exit Sub
            
        ElseIf (m_rstTreeview.RecordCount > 0) Then
        
            m_rstTreeview.MoveFirst
            
        Else
        
            Exit Sub
            
        End If
        
        ' move to next or previous record
        If FetchBuffer.ReadPriorNodes Then
            m_rstTreeview.MovePrevious
        Else
            m_rstTreeview.MoveNext
        End If

    End If
    
    ' fetch specified number of nodes
    For lngCtr = 1 To m_rstTreeview.RecordCount
        
        ' check for BOF or EOF and exit
        If m_rstTreeview.BOF Or m_rstTreeview.EOF Then Exit For
        
        blnFollowUpLicOk = True
        
        Call clsNodes.TriggerBeforeAddFolder(m_conTreeview, m_rstTreeview("Node_ID").Value, blnCancelTranslation)
        
        If m_rstTreeview("Node_Level").Value = 1 And m_rstTreeview("Node_Image").Value = "Repertory Year" Then
            blnCancelTranslation = True
        End If
        If m_rstTreeview("Node_Level").Value = 2 And m_rstTreeview("Node_Image").Value = "Archives Year" Then
            blnCancelTranslation = True
        End If
        
        If blnCancelTranslation Then
            Set nodTreeview = FetchBuffer.Add(Trim$(Str$(m_rstTreeview("Node_ID").Value)), _
                                          m_rstTreeview("Node_Text").Value, CStr(m_rstTreeview!Node_Image), _
                                        CStr(m_rstTreeview!Node_Image), m_rstTreeview("Node_Text").Value)
       
            If m_rstTreeview("Node_Level").Value = 1 And m_rstTreeview("Node_Image").Value = "Repertory Year" Then
                blnCancelTranslation = False
            End If
            If m_rstTreeview("Node_Level").Value = 2 And m_rstTreeview("Node_Image").Value = "Archives Year" Then
                blnCancelTranslation = False
            End If
       
        Else
            'fangs
            'Debug.Print m_rstTreeview("Node_Text").Value
            If m_rstTreeview("Node_Text").Value = "Follow Up Request" Or _
                m_rstTreeview("Node_Text").Value = "Follow Up Request Sent" Or _
                m_rstTreeview("Node_Text").Value = "Follow Up Request Rejected" Then
                
                blnFollowUpLicOk = g_typInterface.ILicense.UserOption(17)
            
            End If
            
            If blnFollowUpLicOk Then
                Set nodTreeview = FetchBuffer.Add(Trim$(Str$(m_rstTreeview("Node_ID").Value)), _
                                              Translate(m_rstTreeview("Node_Text").Value), CStr(m_rstTreeview!Node_Image), _
                                            CStr(m_rstTreeview!Node_Image), Translate(m_rstTreeview("Node_Text").Value))
            End If
        End If
        
        If blnFollowUpLicOk Then 'fangs
            nodTreeview.LoadStyleChildren = ssatLoadStyleChildrenOnDemandKeep
            nodTreeview.Tag = m_rstTreeview("Node_Text").Value
        End If
        
        If FetchBuffer.ReadPriorNodes Then
            m_rstTreeview.MovePrevious
        Else
            m_rstTreeview.MoveNext
        End If
        
    Next

   
End Sub

Private Sub tvwMain_OnDemandPrepare(ParentNode As SSActiveTreeView.SSNode, Result As SSActiveTreeView.SSReturnBoolean)

    
    Dim clsMainSettings As CMainControls
    
    Dim lngParentID As Long
    Dim strSQL As String
    
    Dim enuSortType As SSActiveTreeView.Constants_Sorted
    
    Call clsNodes.OnDemandPrepare(ParentNode, clsMainSettings, m_conTreeview, enuSortType)

    If ParentNode Is Nothing Then
        lngParentID = 0
    Else
        lngParentID = Val(ParentNode.Key)
    End If

    ' construct Sql statement
    strSQL = vbNullString
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & " Nodes "
    strSQL = strSQL & " LEFT OUTER JOIN "
    strSQL = strSQL & " Features "
    strSQL = strSQL & " ON Nodes.Feature_ID = Features.Feature_ID "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " Node_ParentID =  " & lngParentID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " Tree_ID = " & clsMainSettings.TreeID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " (Feature_Activated = True "
    strSQL = strSQL & " OR "
    strSQL = strSQL & " Nodes.Feature_ID = 0 "
    strSQL = strSQL & " OR "
    strSQL = strSQL & " IsNull(Nodes.Feature_ID) = True "
    strSQL = strSQL & " OR "
    strSQL = strSQL & " Trim(Nodes.Feature_ID) = '') "
    strSQL = strSQL & " ORDER BY "
    
    If enuSortType = ssatSortAscending Then
        strSQL = strSQL & " Node_Text ASC "
    Else
        strSQL = strSQL & " Node_Level, "
        strSQL = strSQL & " Node_ID "
    End If
    
    ADORecordsetOpen strSQL, m_conTreeview, m_rstTreeview, adOpenKeyset, adLockOptimistic
    'Call RstOpen(strSQL, m_conTreeview, m_rstTreeview, adOpenKeyset, adLockReadOnly)

    ' check for empty recordset
    If m_rstTreeview.RecordCount = 0 Then
        Result = False
    Else
        Result = True
    End If
    
End Sub

Private Sub tvwMain_TopNodeChange(Node As SSActiveTreeView.SSNode)

    If Me.Visible = True Then
        g_lngTopNode = Val(Node.Key)
        Call clsNodes.UpdateTopNode(g_lngTopNode)
    End If
        
    If Not clsNodes Is Nothing Then
        Call clsNodes.TopNodeChange(Node)
    End If
    
End Sub

Private Sub tvwShortcut_AfterLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)

    Dim conDBConnection As ADODB.Connection
    Dim strTreeTag As String
    Dim blnCancel As Boolean
    
    strTreeTag = tvwMain.Tag
    
    Call clsNodes.AfterLabelEdit(blnCancel, NewString, TREEVIEW_SHORTCUT, conDBConnection)
    
    Cancel = blnCancel
    
    If Cancel = False Then
        tvwShortcut.Tag = strTreeTag
    End If
    
    Call clsNodes.ProcessPopupEvents(Folder_Save)

End Sub

Private Sub tvwShortcut_DragDrop(Source As Control, x As Single, y As Single)

    Dim lngPos As Long
    
    '>> add new shortcut to the 'Favorites' folder
    If Source.Name = "tvwMain" Then
        m_blnBeginDrag = False
        tvwMain.Drag vbEndDrag
        
        If Not tvwShortcut.SelectedItem Is Nothing Then
            If tvwShortcut.SelectedItem.Index > 1 Then
                lngPos = tvwShortcut.SelectedItem.Index
            Else
                lngPos = 1
            End If
        Else
            lngPos = 1
        End If
        
        If clsNodes.AddShortcut(Val(tvwMain.SelectedItem.Key), lngPos) = True Then
            
            
            tvwShortcut.Nodes(tvwMain.SelectedItem.Key).Selected = True
                  
            Call tvwShortcut_NodeClick(tvwShortcut.Nodes(tvwMain.SelectedItem.Key))
        
        End If
    ElseIf Source.Name = "tvwShortcut" Then
        
        
    End If

End Sub

Private Sub tvwShortcut_DragOver(Source As Control, x As Single, y As Single, State As Integer)

    Dim nodHighlighted As SSNode
    
    Set nodHighlighted = tvwShortcut.HitTest(x, y)
    
    If Not nodHighlighted Is Nothing Then
        nodHighlighted.Selected = True
        Set tvwShortcut.DropHighlight = nodHighlighted
    Else
        Set tvwShortcut.DropHighlight = Nothing
    End If
    
    Set nodHighlighted = Nothing
    
End Sub

Private Sub tvwShortcut_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intAns As Integer
    
    
    Select Case KeyCode
        Case vbKeyDelete
            
            If tvwShortcut.SelectedItem.Key <> "NoButtons" Then
                '>> delete shortcut button
                intAns = MsgBox("Are you sure you want to remove '" & tvwShortcut.SelectedItem & "' folder in the list of your 'Favorite' folders?", vbQuestion + vbYesNo, "Remove Confirmation")
                If intAns = vbYes Then
                    clsNodes.DeleteShortcut
                End If
            End If
            
        Case vbKeyDown
            If Not tvwShortcut.SelectedItem Is Nothing Then
                If tvwShortcut.SelectedItem.Index = tvwShortcut.Nodes.Count Then
                    Set tvwShortcut.DropHighlight = Nothing
                    tvwMain.SetFocus
                    tvwMain.SelectedItem = tvwMain.Nodes(1)
                End If
            End If
            
        Case vbKeyReturn
            
            If Not tvwShortcut.SelectedItem Is Nothing Then
                If tvwShortcut.SelectedItem.Key <> "NoButtons" Then
                    tvwShortcut.DropHighlight = tvwShortcut.SelectedItem
                
                    '>> load grid records and settings of the selected shortcut of nodes
                    clsNodes.TriggerNodeClickEvent tvwMain.Nodes(tvwShortcut.SelectedItem.Key)
                
                    Call PaneResize
                    
                    tvwShortcut.SetFocus
                End If
            End If
    End Select
    
    m_enuKeypressed = KeyCode
    
End Sub


Private Sub tvwShortcut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim nodSelected As SSNode
    
    If Button = vbRightButton Then
        Set nodSelected = tvwShortcut.HitTest(x, y)
        If Not nodSelected Is Nothing Then
            Set tvwShortcut.SelectedItem = nodSelected
            tvwShortcut.DropHighlight = tvwShortcut.SelectedItem
            Call clsNodes.ShowPopupMenu(nodSelected)
        End If
        Set nodSelected = Nothing
    ElseIf Button = vbLeftButton Then
        '>> enable dragging of node
        If Not tvwShortcut.SelectedItem Is Nothing Then
            m_blnBeginDrag = True
        End If
    End If

    m_enuMouseButton = Button
    
End Sub

Private Sub tvwShortcut_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim nodSelected As SSNode
    
    If Button = vbLeftButton And m_blnBeginDrag = True Then
        '>> show drag icon upon moving
        Set nodSelected = tvwShortcut.HitTest(x, y)
        If Not nodSelected Is Nothing Then
            If nodSelected <> tvwShortcut.SelectedItem Then
                tvwShortcut.DragIcon = nodSelected.CreateDragImage
            Else
                tvwShortcut.DragIcon = tvwShortcut.SelectedItem.CreateDragImage
            End If
            tvwShortcut.Drag vbBeginDrag
        End If
        
        Set nodSelected = Nothing
        
    Else
        m_blnBeginDrag = False
        tvwShortcut.Drag vbEndDrag
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Function TreeHeight() As Double

    Dim objNode As Object
    
    Dim dblTreeHeight As Long
    Dim lngNodeCtr As Long
    Dim blnNodeIsExpanded As Boolean
    
    '>> compute height of the folder tree based on the number of nodes shown in the tree
    dblTreeHeight = 0
    If tvwMain.Nodes.Count > 0 Then
        For lngNodeCtr = 1 To tvwMain.Nodes.Count
            If tvwMain.Nodes(lngNodeCtr).Level = 1 Then
                dblTreeHeight = dblTreeHeight + 275
            Else
                Set objNode = tvwMain.Nodes(lngNodeCtr)
                Do While Not objNode.Parent Is Nothing
                    If objNode.Parent.Expanded = True Then
                        Set objNode = objNode.Parent
                        blnNodeIsExpanded = True
                    Else
                        Set objNode = Nothing
                        blnNodeIsExpanded = False
                        Exit Do
                    End If
                Loop
                If blnNodeIsExpanded = True Then
                    dblTreeHeight = dblTreeHeight + 275
                End If
            End If
        Next
    End If
    
    TreeHeight = dblTreeHeight
    
End Function

Public Sub PaneResize()

    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim dblViewHeight As Double
    Dim lngCtr As Long
    
    '>> resize pane controls
        
    dblLeft = 0
    dblTop = 0
    dblWidth = Me.Width
    dblHeight = sccCaption.Height
    sccCaption.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = sccCaption.Height
    dblWidth = Me.Width
    dblHeight = sccShortcut.Height
    sccShortcut.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = sccCaption.Height + sccShortcut.Height
    dblWidth = Me.Width
    dblHeight = tvwShortcut.Height
    
    tvwShortcut.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = tvwShortcut.Top + tvwShortcut.Height
    dblWidth = Me.Width
    dblHeight = sccFolders.Height
    
    sccFolders.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = sccFolders.Top + sccFolders.Height
    dblWidth = Me.Width
    dblHeight = Me.Height - sccFolders.Top - sccFolders.Height
    
    picBGroundTree.Move dblLeft, dblTop, dblWidth, IIf(dblHeight < 0, 0, dblHeight)
    
    dblViewHeight = 0
    For lngCtr = 0 To Controls.Count - 1
        If Left(Controls(lngCtr).Name, 7) = "optView" Then
            dblViewHeight = dblViewHeight + Controls(lngCtr).Height
        End If
    Next
        
    If fraViews.Visible = True Then
     
        On Error Resume Next
        If optViews(1).Visible = True Then
            If (Err.Number > 0) Then
                lblCustomize.Move 20000, dblTop, dblWidth, lblCustomize.Height
                sccViews.Left = 20000
                fraViews.Left = 20000
                dblLeft = 0
                dblTop = 0
                dblWidth = picBGroundTree.Width
                dblHeight = picBGroundTree.Height

            Else
                dblLeft = 0
                dblTop = 0
                dblWidth = picBGroundTree.Width
                dblHeight = picBGroundTree.Height - dblViewHeight - sccViews.Height - 500

            End If
        End If
        
        If dblHeight < 1000 Then
            
            tvwMain.Move dblLeft, dblTop, dblWidth, picBGroundTree.Height
            fraViews.Left = 20000
            lblCustomize.Left = 20000
            sccViews.Left = 20000

        Else
            
            Debug.Print "tvwMain dimensions. Left: " & dblLeft & ", Top: " & dblTop & ", Width: " & dblWidth & ", Height: " & dblHeight
            tvwMain.Move dblLeft, dblTop, dblWidth, dblHeight
            
            If (Err.Number = 0) Then
                
                dblLeft = 0
                dblTop = picBGroundTree.Height - dblViewHeight - lblCustomize.Height - 500
                dblWidth = picBGroundTree.Width
                dblHeight = sccViews.Height
            
                sccViews.Move dblLeft, dblTop, dblWidth, dblHeight
                '>> Set the font size because it grows bigger upon resizing
                sccViews.Font.Size = 8
                
                dblLeft = 0
                dblTop = sccViews.Top + sccViews.Height
                dblWidth = picBGroundTree.Width
                
                Debug.Print "fraViews dimensions. Left: " & dblLeft & ", Top: " & dblTop & ", Width: " & dblWidth & ", Height: " & dblViewHeight + lblCustomize.Height + 500
                fraViews.Move dblLeft, dblTop, dblWidth, dblViewHeight + lblCustomize.Height + 500
                
                dblLeft = 150
                dblTop = fraViews.Height - lblCustomize.Height - 500
                dblWidth = picBGroundTree.Width
                dblHeight = lblCustomize.Height
    
                lblCustomize.Move dblLeft, dblTop, dblWidth, dblHeight
            
            End If
            
        End If
        
        On Error GoTo 0

    Else
        dblLeft = 0
        dblTop = 0
        dblWidth = picBGroundTree.Width
        dblHeight = picBGroundTree.Height
        tvwMain.Move dblLeft, dblTop, dblWidth, dblHeight
    
    End If
    
    On Error Resume Next
    Set tvwMain.TopNode = tvwMain.Nodes(CStr(g_lngTopNode))
    On Error GoTo 0
    
err_hand:
    
End Sub

Private Sub ResizeTree()

    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim lngCtr As Long
    
    dblLeft = 0
    dblTop = sccFolders.Top + sccFolders.Height
    dblWidth = Me.Width
    dblHeight = Me.Height - dblTop
    
    If dblHeight < 0 Then
        dblHeight = 0
    End If
    
    picBGroundTree.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = 0
    dblWidth = picBGroundTree.Width
    dblHeight = TreeHeight
    
    tvwMain.Move dblLeft, dblTop, dblWidth, dblHeight
    
    dblLeft = 0
    dblTop = tvwMain.Height
    dblWidth = picBGroundTree.Width
    dblHeight = sccViews.Height
    
    sccViews.Move dblLeft, dblTop, dblWidth, dblHeight
    '>> Set the font size because it grows bigger upon resizing
    sccViews.Font.Size = 8
    
    dblLeft = 0
    dblTop = sccViews.Top + sccViews.Height
    dblWidth = picBGroundTree.Width
    dblHeight = 0
    
    For lngCtr = 0 To Controls.Count - 1
        If Left(Controls(lngCtr).Name, 7) = "optView" Then
            dblHeight = dblHeight + Controls(lngCtr).Height + 75
        End If
    Next
        
    lblCustomize.Move 20000, lblCustomize.Top, lblCustomize.Width, lblCustomize.Height
    sccViews.Move 20000, sccViews.Top, sccViews.Width, sccViews.Height
    
    On Error GoTo Next_Proc
    If optViews(1).Visible = True Then
        sccViews.Left = 0
        lblCustomize.Move 150, dblHeight - 200, lblCustomize.Width, lblCustomize.Height
    End If
    
    
    
Next_Proc:
    On Error GoTo 0
    fraViews.Move dblLeft, dblTop, dblWidth, dblHeight + lblCustomize.Height
    

End Sub

Private Function ForceLoadNode(ByVal NodeKey As Long) As Boolean
    Dim strCheckExists As String
    Dim blnNodeOK As Boolean
    Dim strCommand As String
    Dim strNodeKey As Long
    Dim arrNodesToExpand() As String
    Dim lngNodesToExpandCtr As Long
    Dim rstParentNode As ADODB.Recordset
    
    ' Initialize as with Error; Node is missing
    blnNodeOK = False
    strNodeKey = NodeKey
    lngNodesToExpandCtr = 0
    Erase arrNodesToExpand
    
    Do Until blnNodeOK
        
        lngNodesToExpandCtr = lngNodesToExpandCtr + 1
        
        ReDim Preserve arrNodesToExpand(1 To lngNodesToExpandCtr)
        arrNodesToExpand(lngNodesToExpandCtr) = strNodeKey
                    
        On Error Resume Next
        ' Important to not to use CStr; With CStr Nodes use Key, without it Nodes use Index
        ' Error that occurs are 35601 and 35600 respectively
        strCheckExists = tvwMain.Nodes(CStr(strNodeKey)).Text
        Select Case Err.Number
            Case 35601
                ' Correct Missing Node. Force Node to be loaded into tvwMain
                    strCommand = ""
                    strCommand = strCommand & "SELECT "
                    strCommand = strCommand & "ParentNodes.Node_ID As ParentNodeID "
                    strCommand = strCommand & "FROM "
                    strCommand = strCommand & "Nodes INNER JOIN Nodes As ParentNodes ON Nodes.Node_ParentID = ParentNodes.Node_ID "
                    strCommand = strCommand & "WHERE "
                    strCommand = strCommand & "Nodes.Node_ID = " & strNodeKey & " "
                
                ADORecordsetOpen strCommand, m_conTreeview, rstParentNode, adOpenKeyset, adLockOptimistic
                'RstOpen strCommand, m_conTreeview, rstParentNode, adOpenKeyset, adLockOptimistic, , True
                If rstParentNode.RecordCount > 0 Then
                    rstParentNode.MoveFirst
                    
                    strNodeKey = CStr(rstParentNode.Fields("ParentNodeID").Value)
                Else
                    Debug.Assert False
                End If
                ADORecordsetClose rstParentNode
                
            Case 0
                ' Do Nothing
                blnNodeOK = True
                
            Case Else
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                blnNodeOK = True
        End Select
        On Error GoTo 0
    Loop
    
    If lngNodesToExpandCtr > 1 Then
        ' Don't make the expansion and collapsing visible to the user
        tvwMain.Redraw = False
        
        ' Expand
        ' Topmost level folder must be expanded if not yet expanded
        For lngNodesToExpandCtr = UBound(arrNodesToExpand) To 2 Step -1
            tvwMain.Nodes(arrNodesToExpand(lngNodesToExpandCtr)).Expanded = True
        Next
        
        ' Collapse
        For lngNodesToExpandCtr = UBound(arrNodesToExpand) To 2 Step -1
            tvwMain.Nodes(arrNodesToExpand(lngNodesToExpandCtr)).Expanded = False
        Next

        ' Make succeeding expansion and collapsing visible to the user
        tvwMain.Redraw = True
    End If
End Function

Private Sub tvwShortcut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call clsNodes.FavoritesMouseUp(Button, Shift, x, y)
End Sub

Private Sub tvwShortcut_NodeClick(Node As SSActiveTreeView.SSNode)
    Screen.MousePointer = vbHourglass
    
    If m_enuMouseButton = vbLeftButton Then
        If tvwShortcut.Nodes.Count >= 1 And m_enuKeypressed <> vbKeyDown And m_enuKeypressed <> vbKeyUp Then
            If tvwShortcut.Nodes(1).Key = "NoButtons" Then
                tvwMain.SetFocus
                Exit Sub
            Else
                Set tvwMain.DropHighlight = Nothing
                Set tvwMain.SelectedItem = Nothing
            
                tvwShortcut.DropHighlight = Node
                                
                
                ' Correct Missing Node. Force Node to be loaded into tvwMain
                ForceLoadNode tvwShortcut.SelectedItem.Key
                
                '>> load grid records and settings of the selected shortcut of nodes
                clsNodes.TriggerNodeClickEvent tvwMain.Nodes(tvwShortcut.SelectedItem.Key)
                
                
                Call PaneResize
                
                tvwShortcut.SetFocus
            End If
        Else
            m_enuKeypressed = 0
        End If
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub ResizeFavoriteSection()

    If tvwShortcut.Nodes.Count > 1 Then
        If (tvwShortcut.Nodes.Count * 250) > 1500 Then
            tvwShortcut.Height = 1500
        Else
            tvwShortcut.Height = tvwShortcut.Nodes.Count * 250
        End If
    Else
        tvwShortcut.Height = 250
    End If
    
    Call PaneResize
    
End Sub

Private Sub EditMainNode(NewString As String, Node_ID As Long, ByRef ADOConnection As ADODB.Connection)

    Dim strCommandText As String
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "Nodes "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "Node_Text = '" & NewString & "' "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & Node_ID
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "Buttons "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "Button_Caption = '" & NewString & "' "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & Node_ID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "Button_Default = False "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
End Sub




