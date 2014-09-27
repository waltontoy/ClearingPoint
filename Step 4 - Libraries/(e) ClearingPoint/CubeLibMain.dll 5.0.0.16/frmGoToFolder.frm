VERSION 5.00
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frmGoToFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go to Folder"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4635
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4635
      Width           =   1215
   End
   Begin SSActiveTreeView.SSTree tvwFolders 
      Height          =   4410
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   7779
      _Version        =   65536
      LabelEdit       =   1
      LineStyle       =   1
      ScrollStyle     =   2
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
End
Attribute VB_Name = "frmGoToFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsMainSettings As CMainControls
Private m_clsNavigationPane As INavigationPane
Private m_clsNavPaneProps As CNavigationPane

Private m_conADOConnection As ADODB.Connection
Private m_rstTreeview As ADODB.Recordset

Private m_strApplicationName As String


Public Sub GoToFolder(ByRef OwnerForm As Form, _
                      ByRef Application As App, _
                      ByRef ADOConnection As ADODB.Connection, _
                      ByRef MainSettings As CMainControls, _
                      ByRef NavigationPane As INavigationPane, _
                      ByRef NavPaneProps As CNavigationPane)
    
    Dim lngCtr As Long
    
    
    '--->Set parameters to object local variables
    Set m_clsMainSettings = MainSettings
    Set m_clsNavigationPane = NavigationPane
    Set m_clsNavPaneProps = NavPaneProps
    Set m_conADOConnection = ADOConnection
    
    m_strApplicationName = Application.ProductName
    
    
    '--->Set up treeview
    tvwFolders.UseImageList = True
    Set tvwFolders.ImageList = g_typInterface.ITreeIcons
    tvwFolders.Tag = m_clsMainSettings.TreeID
    tvwFolders.Refresh
    
    
    '--->Collapse nodes with level > 1
    For lngCtr = 1 To tvwFolders.Nodes.Count
        If (tvwFolders.Nodes(lngCtr).Level > 1) Then
            tvwFolders.Nodes(lngCtr).Expanded = False
        Else
            tvwFolders.Nodes(lngCtr).Expanded = True
        End If
    Next lngCtr
    
    
    If (frmPane.tvwMain.SelectedItem Is Nothing = False) Then
        '--->Expand parent of selected node
        Call FindSelectedNode(frmPane.tvwMain.SelectedItem)
        
        '--->Set selected node to currently selected node in the pane window
        Set tvwFolders.SelectedItem = tvwFolders.Nodes(frmPane.tvwMain.SelectedItem.Key)
        tvwFolders.DropHighlight = tvwFolders.SelectedItem
        
        tvwFolders.SelectedItem.EnsureVisible
    End If
    
    
    '--->Display form
    Set Me.Icon = OwnerForm.Icon
    Me.Show vbModal
    
End Sub

Private Function ExpandedNode(ByVal NodeID As Long) As Boolean
    Dim rstTreeSetting As ADODB.Recordset
    Dim lngCharPos As Long
    Dim lngNodeID As Long
    
    
    ExpandedNode = False
    
    If (GetExpandedNodes(rstTreeSetting) = QueryResultSuccessful) Then
        If (rstTreeSetting.EOF = False) Then
            lngCharPos = 1
            
            Do While lngCharPos < Len(Trim(rstTreeSetting!TreeSet_ExpandedNodes))
                lngNodeID = Val(Mid(Trim(rstTreeSetting!TreeSet_ExpandedNodes), lngCharPos, 4))
                
                If (lngNodeID = NodeID) Then
                    ExpandedNode = True
                    
                    Exit Do
                End If
                
                lngCharPos = lngCharPos + 5
            Loop
        End If
    End If
    
    Set rstTreeSetting = Nothing
    
End Function

Private Function GetExpandedNodes(ExpandedNodes As ADODB.Recordset) As QueryResultConstants
    Dim strCommandText As String
    
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "TreeSet_ExpandedNodes "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "TreeSettings "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & m_clsMainSettings.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Tree_ID = " & m_clsMainSettings.TreeID
    
    ADORecordsetOpen strCommandText, m_conADOConnection, ExpandedNodes, adOpenKeyset, adLockOptimistic
    'Set ExpandedNodes = m_conADOConnection.Execute(strCommandText)
    
    If (ExpandedNodes.EOF And ExpandedNodes.BOF) Then
        GetExpandedNodes = QueryResultNoRecord
    Else
        GetExpandedNodes = QueryResultSuccessful
    End If
    
End Function

Private Sub cmdCancel_Click()
    
    '--->Exit form
    Unload frmGoToFolder
    
End Sub

Private Sub cmdOK_Click()
    Dim rstNode As ADODB.Recordset
    
    Dim strCommandText As String
    Dim strNodeKey As String
    Dim arrNodeKeys() As String
    
    Dim lngCtr As Long
    Dim blnNodeIsHidden As Boolean
    
    
    If (tvwFolders.SelectedItem Is Nothing = False) Then
        
        '--->Change current view to selected folder
        
        Set frmPane.tvwMain.DropHighlight = Nothing
        Set frmPane.tvwMain.SelectedItem = Nothing
        
        
        '--->Check if selected folder is currently hidden in the pane window
        strNodeKey = tvwFolders.SelectedItem.Key
        blnNodeIsHidden = True
        For lngCtr = 1 To frmPane.tvwMain.Nodes.Count
            If (frmPane.tvwMain.Nodes(lngCtr).Key = strNodeKey) Then
                blnNodeIsHidden = False
                Exit For
            End If
        Next lngCtr
        
        If (blnNodeIsHidden = True) Then
            ReDim Preserve arrNodeKeys(0)
            arrNodeKeys(0) = strNodeKey
            
            Do While True
                strCommandText = vbNullString
                strCommandText = strCommandText & "SELECT "
                strCommandText = strCommandText & "Node_ParentID "
                strCommandText = strCommandText & "FROM "
                strCommandText = strCommandText & "Nodes "
                strCommandText = strCommandText & "WHERE "
                strCommandText = strCommandText & "Node_ID = " & Val(strNodeKey)
                
                ADORecordsetOpen strCommandText, m_conADOConnection, rstNode, adOpenKeyset, adLockOptimistic
                'Set rstNode = m_conADOConnection.Execute(strCommandText)
                
                If Not (rstNode.EOF And rstNode.BOF) Then
                    rstNode.MoveFirst
                    
                    If (rstNode!Node_ParentID = 0) Then
                        Exit Do
                    Else
                        strNodeKey = rstNode!Node_ParentID
                        ReDim Preserve arrNodeKeys(UBound(arrNodeKeys) + 1)
                        arrNodeKeys(UBound(arrNodeKeys)) = strNodeKey
                    End If
                End If
                
                Call ADORecordsetClose(rstNode)
            Loop
            
            For lngCtr = UBound(arrNodeKeys) To 0 Step -1
                frmPane.tvwMain.Nodes(arrNodeKeys(lngCtr)).Expanded = True
            Next lngCtr
            
        End If
        
        
        frmPane.tvwMain.Nodes(tvwFolders.SelectedItem.Key).Selected = True
        '--->Load grid records and settings of the selected shortcut of nodes
        m_clsNavigationPane.TriggerNodeClickEvent frmPane.tvwMain.Nodes(tvwFolders.SelectedItem.Key)
        frmPane.tvwMain.DropHighlight = frmPane.tvwMain.Nodes(tvwFolders.SelectedItem.Key)
        
        '--->Resize pane window
        Call frmPane.PaneResize
        
        
        '--->Exit form
        Unload frmGoToFolder
        
    Else
        
        MsgBox "Please select a folder to go to.", vbInformation, m_strApplicationName
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call ADORecordsetClose(m_rstTreeview)
    
End Sub

Private Sub tvwFolders_NodeClick(Node As SSActiveTreeView.SSNode)
    
    tvwFolders.DropHighlight = Node
    
End Sub

Private Sub tvwFolders_OnDemandFetch(ByVal FetchBuffer As SSActiveTreeView.SSFetchBuffer)
    Dim intIndent As Integer
    Dim lngCtr As Long
    Dim nodTreeview As SSNode
    
    Dim blnCancelTranslation As Boolean
    
    
    If (FetchBuffer.ParentNode Is Nothing) Then
        intIndent = 0
    Else
        intIndent = FetchBuffer.ParentNode.Level
    End If
    
    
    If (FetchBuffer.StartNode Is Nothing) Then
        
        If (FetchBuffer.ReadPriorNodes = True) Then
            If (m_rstTreeview.EOF = False Or m_rstTreeview.BOF = False) Then
                m_rstTreeview.MoveLast
            End If
        Else
            If (m_rstTreeview.EOF = False Or m_rstTreeview.BOF = False) Then
                m_rstTreeview.MoveFirst
            End If
        End If
        
    Else
        
        If (m_rstTreeview Is Nothing) Then
            Exit Sub
        ElseIf (m_rstTreeview.RecordCount > 0) Then
            m_rstTreeview.MoveFirst
        Else
            Exit Sub
        End If
        
        If (FetchBuffer.ReadPriorNodes = True) Then
            m_rstTreeview.MovePrevious
        Else
            m_rstTreeview.MoveNext
        End If
        
    End If
    
    
    
    For lngCtr = 1 To m_rstTreeview.RecordCount
        
        If (m_rstTreeview.BOF Or m_rstTreeview.EOF) Then
            Exit For
        End If
        
        Call m_clsNavigationPane.TriggerBeforeAddFolder(m_conADOConnection, CLng(m_rstTreeview("Node_ID").Value), blnCancelTranslation)
        Set nodTreeview = FetchBuffer.Add(Trim(Str(m_rstTreeview("Node_ID").Value)), IIf(blnCancelTranslation, m_rstTreeview("Node_Text").Value, Translate(m_rstTreeview("Node_Text").Value)), CStr(m_rstTreeview("Node_Image").Value), CStr(m_rstTreeview("Node_Image").Value), IIf(blnCancelTranslation, m_rstTreeview("Node_Text").Value, Translate(m_rstTreeview("Node_Text").Value)))
        
        nodTreeview.LoadStyleChildren = ssatLoadStyleChildrenOnDemandKeep
        nodTreeview.Tag = m_rstTreeview("Node_Text").Value
        
        If (FetchBuffer.ReadPriorNodes = True) Then
            m_rstTreeview.MovePrevious
        Else
            m_rstTreeview.MoveNext
        End If
        
    Next lngCtr
    
End Sub

Private Sub tvwFolders_OnDemandPrepare(ParentNode As SSActiveTreeView.SSNode, Result As SSActiveTreeView.SSReturnBoolean)
    Dim lngParentID As Long
    Dim strSQL As String
    
    
    If (ParentNode Is Nothing) Then
        lngParentID = 0
    Else
        lngParentID = Val(ParentNode.Key)
    End If
    
    
    strSQL = vbNullString
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & " Nodes "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " Node_ParentID =  " & lngParentID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " Tree_ID = " & m_clsMainSettings.TreeID
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " Node_ID "
    
    ADORecordsetOpen strSQL, m_conADOConnection, m_rstTreeview, adOpenKeyset, adLockOptimistic
    'Call RstOpen(strSQL, m_conADOConnection, m_rstTreeview, adOpenKeyset, adLockOptimistic, , True)
    
    
    '--->Return Result
    If (m_rstTreeview.RecordCount = 0) Then
        Result = False
    Else
        Result = True
    End If
    
End Sub

Private Sub FindSelectedNode(ByRef SelectedNode As SSActiveTreeView.SSNode)
    Dim rstNode As ADODB.Recordset
    
    Dim strCommandText As String
    Dim strNodeKey As String
    Dim arrNodeKeys() As String
    
    Dim lngCtr As Long
    Dim blnNodeIsHidden As Boolean
    
    
    strNodeKey = SelectedNode.Key
    blnNodeIsHidden = True
    For lngCtr = 1 To tvwFolders.Nodes.Count
        If (tvwFolders.Nodes(lngCtr).Key = strNodeKey) Then
            blnNodeIsHidden = False
            Exit For
        End If
    Next lngCtr
    
    
    If (blnNodeIsHidden = True) Then
        ReDim Preserve arrNodeKeys(0)
        arrNodeKeys(0) = strNodeKey
        
        Do While True
            strCommandText = vbNullString
            strCommandText = strCommandText & "SELECT "
            strCommandText = strCommandText & "Node_ParentID "
            strCommandText = strCommandText & "FROM "
            strCommandText = strCommandText & "Nodes "
            strCommandText = strCommandText & "WHERE "
            strCommandText = strCommandText & "Node_ID = " & Val(strNodeKey)
                    
            ADORecordsetOpen strCommandText, m_conADOConnection, rstNode, adOpenKeyset, adLockOptimistic
            'Set rstNode = m_conADOConnection.Execute(strCommandText)
            
            If Not (rstNode.EOF And rstNode.BOF) Then
                rstNode.MoveFirst
                
                If (rstNode!Node_ParentID = 0) Then
                    Exit Do
                Else
                    strNodeKey = rstNode![Node_ParentID]
                    ReDim Preserve arrNodeKeys(UBound(arrNodeKeys) + 1)
                    arrNodeKeys(UBound(arrNodeKeys)) = strNodeKey
                End If
            End If
            
            Call ADORecordsetClose(rstNode)
        Loop
        
        For lngCtr = UBound(arrNodeKeys) To 0 Step -1
            tvwFolders.Nodes(arrNodeKeys(lngCtr)).Expanded = True
        Next lngCtr
        
    End If
    
End Sub
