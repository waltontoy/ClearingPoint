VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INavigationPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent NodeEvents[(arg1, arg2, ... , argn)]
Public Event NodeEvents(ByRef MainApplication As Object, ByRef Window As Object, ByRef MainProps As CMainControls, ByRef NavigationPane As CNavigationPane, ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection, ByRef Cancel As Boolean)
Public Event CustomView(ByRef Window As Object, ByRef GridProps As CGrid, ByRef NavPane As CNavigationPane, ByRef ADOConnection As ADODB.Connection)
Public Event NodeClick(ByVal TreeProps As CNavigationPane)
Public Event PopupMenu(SelectedFavoriteFolder As SSActiveTreeView.SSNode)
Public Event TopNodeChange(ByRef NavigationPane As CNavigationPane)
Public Event tvwShortcutAfterLabelEdit(Cancel As Boolean, NewString As String, ByRef ADOConnection As ADODB.Connection)
Public Event tvwMainAfterLabelEdit(Cancel As Boolean, NewString As String, ByRef ADOConnection As ADODB.Connection)
Public Event tvwMainBeforeNodeClick(Node As SSNode, Cancel As Boolean)
Public Event tvwMainClick()
Public Event tvwMainCollapse(Node As SSNode)
Public Event tvwMainEscapeLabelEdit(Cancel As Boolean, NewString As String)
Public Event tvwMainExpand(Node As SSNode)
Public Event tvwMainKeyUp(KeyCode As Integer, Shift As Integer)
Public Event tvwMainMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event tvwMainMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event tvwMainMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event tvwShortcutMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event tvwMainNodeClick(Node As SSNode)
Public Event tvwMainOnDemandFetch(ByVal FetchBuffer As SSFetchBuffer)
Public Event tvwMainOnDemandPrepare(ParentNode As SSNode, ByRef MainSettings As CMainControls, ByRef ADOConnection As ADODB.Connection, SortType As SSActiveTreeView.Constants_Sorted)
Public Event tvwMainTopNodeChange(Node As SSNode)
Public Event ViewOptionClick(ViewCaption As String, ViewPosition As Integer)
Public Event BeforeAddFolder(ByRef ADOConnection As ADODB.Connection, ByVal NodeID As Long, ByRef CancelTranslation As Boolean)


Public Enum FaveFolderProcess
    Folder_Save = 0
    Folder_Add = 1
    Folder_Modify = 2
    Folder_Delete = 3
End Enum

Public Enum TreeviewControl
    TREEVIEW_MAIN = 1
    TREEVIEW_SHORTCUT = 2
End Enum

Private m_blnDontRefresh As Boolean





Public Function LoadGroupButtons(ByRef Window As Object, _
                        ByVal UserID As Long, _
                        ByRef NavigationPane As CNavigationPane, _
                        ByRef Licensee As CLicensee, _
                        ByRef ADOConnection As ADODB.Connection, _
                        Optional ByVal lngResourceHandler As Long = 0) As Boolean
    
    Dim enuResult As QueryResultConstants

    If lngResourceHandler <> 0 Then
        ResourceHandler = lngResourceHandler
    End If

    If Licensee.CheckLicense(ADOConnection) = False Then
        LoadGroupButtons = False
        Exit Function
    End If

    enuResult = NavigationPane.LoadGroups(Window, ADOConnection, UserID)
    
    Select Case enuResult
        Case QueryResultSuccessful
            LoadGroupButtons = True
            
        Case QueryResultNoRecord
            LoadGroupButtons = False
            MsgBox "No group buttons were found in the database. " & _
                    "Please contact technical support.", vbInformation + vbOKOnly, _
                    "Main Setting Error: LoadGroupButtons Module"
            
        Case QueryResultError
            LoadGroupButtons = False
            MsgBox "Error has occured while trying to access the database. " & _
                    "Please contact technical support.", vbInformation + vbOKOnly, _
                    "Main Setting Error: LoadGroupButtons Module"
            
    End Select
    
    
End Function


Public Sub SelectGroupButton(ByVal Application As Object, _
                        ByRef Window As Object, _
                        ByVal UserID As Long, _
                        ByVal TreeID As Long, _
                        ByRef NavigationPane As CNavigationPane, _
                        ByRef GridProps As CGrid, _
                        ByRef ADOConnection As ADODB.Connection, _
                        Optional TraceFileOn As Boolean = False)

        Dim clsMain As IMainControls
        
        Dim enuResult As QueryResultConstants
        Dim strMdbPath As String
        
        On Error GoTo Error_Handler
        
        If g_clsLog Is Nothing Then
10          Set g_clsLog = New CAppLogFile
        End If
        

        g_blnTraceFileOn = TraceFileOn
        
        g_typInterface.IButtonBar.Enabled = False
        
11      strMdbPath = GetMDBPath(ADOConnection.ConnectionString)
        
20      g_clsLog.StartLogFile Application.ProductName, strMdbPath, UserID, g_blnTraceFileOn
30      g_clsLog.AppendLog "Start SelectGroupButton Method"
        
100     If NavigationPane.TreeID > 0 And NavigationPane.UserID > 0 Then
110         g_clsLog.AppendLog "Start SaveTreeSettings procedure"
120         Call SaveTreeSettings(Application, NavigationPane, GridProps, ADOConnection)
            g_clsLog.AppendLog "End SaveTreeSettings procedure"
        End If
        
130     g_clsLog.AppendLog "Save new selected group's Tree_ID"
140     g_clsLog.AppendLog "Tree ID: " & TreeID
150     NavigationPane.TreeID = TreeID
        
160     g_clsLog.AppendLog "Retrieve selected group info from ButtonGroups table"
170     g_clsLog.AppendLog "User ID: " & UserID
180     enuResult = NavigationPane.SelectGroup(UserID, TreeID, ADOConnection)
        
        Select Case enuResult
        
            Case QueryResultSuccessful
            
                Set frmPane.clsNodes = Me
                
190             g_clsLog.AppendLog "Start LoadTrees procedure"
200             Call LoadTrees(Application, Window, NavigationPane, ADOConnection)
210             g_clsLog.AppendLog "End LoadTrees procedure"
    
220             frmPane.sccCaption.Caption = Replace(NavigationPane.Caption, "&", "&&")
                
'230             Set frmPane.clsNodes = Me
                
                On Error Resume Next
240             frmPane.tvwMain.SelectedItem = frmPane.tvwMain.Nodes(CStr(NavigationPane.NodeID))
                On Error GoTo 0
                
250             frmPane.tvwMain.DropHighlight = frmPane.tvwMain.SelectedItem
260             Set NavigationPane.SelectedNode = frmPane.tvwMain.SelectedItem
    
                
270             g_clsLog.AppendLog "Start LoadViews procedure"
280             If LoadViews(Application, NavigationPane, GridProps, ADOConnection) = QueryResultNoRecord Then
                    
290                 g_clsLog.AppendLog "No views available for selected node"
300                 g_typInterface.IGrid.GroupByBoxVisible = False
            
330                 If g_typInterface.IReadingPane.Visible = True Then
340                     If (g_typInterface.INoView.Width + g_typInterface.INoView.Left) _
                            > g_typInterface.IReadingPane.Left Then
350                             g_typInterface.INoView.Width = g_typInterface.IReadingPane.Left - g_typInterface.INoView.Left
                        Else
360                         g_typInterface.INoView.Width = 2805
                        End If
                    End If
            
370                 Set g_typInterface.IGrid.ADORecordset = Nothing
                    
                    g_typInterface.IGrid.Columns.Clear
                    
380                 g_typInterface.IStatusbar.Panels(1).Text = ""
                    
390                 g_typInterface.IGridHeader.Caption = " " & Replace(NavigationPane.SelectedNode.Text, "&", "&&")
                    
                    g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
                    g_typInterface.INoView.Left = 300
                
400                 g_typInterface.INoView.ZOrder 0
                    g_typInterface.INoView.Caption = "No View"
                
                    Set clsMain = New IMainControls
                    clsMain.LoadArrangeByButtons Window, GridProps
                    Set clsMain = Nothing
                End If
                
410             g_clsLog.AppendLog "Resize navigation pane controls"
420             Call frmPane.PaneResize
                g_clsLog.AppendLog "Finished resizing navigation pane"
                
            Case QueryResultNoRecord
430             g_clsLog.AppendLog "No record for the selected group was found"
                MsgBox "No record for the selected group was found in the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: SelectGroupButton Module"
            
            Case QueryResultError
440             g_clsLog.AppendLog "Error occured in retrieving selected group's record"
                MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: SelectGroupButton Module"
            
        End Select
        
450     g_clsLog.AppendLog "Selected group was loaded successfully"
                
'Reims Nov 22, 2005 - Let mark review creation and destruction of applog classes
'460     g_clsLog.CloseLogFile
'Reims Nov 22, 2005 - Let mark review creation and destruction of applog classes
'        Set g_clsLog = Nothing
        
        g_typInterface.IButtonBar.Enabled = True
        
        Exit Sub
    
Error_Handler:
        g_clsLog.AppendLog "Error occured!"
        g_clsLog.AppendLog "Method: SelectGroupButton"
        g_clsLog.AppendLog "Error Number: " & Err.Number
        g_clsLog.AppendLog "Description: " & Err.Description
        g_clsLog.AppendLog "Line Number: " & Erl
        g_clsLog.CloseLogFile
        
        MsgBox "Error has occured while trying to load the selected button group." & vbCrLf & _
                "Line Number: " & Erl & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Description: " & Err.Description, vbInformation + vbOKOnly, _
                "Main Setting Error: SelectGroupButton Module"
    
    
End Sub

Private Sub LoadTrees(ByVal Application As Object, _
                    ByRef Window As Object, _
                    ByRef NavigationPane As CNavigationPane, _
                    ByRef ADOConnection As ADODB.Connection)
    
        Dim rstButtons As ADODB.Recordset
        Dim rstNodes As ADODB.Recordset
        Dim strExpandedNodes As String
        Dim enuResult As QueryResultConstants
        Dim lngCtr As Long
        Dim strMdbPath As String
        
        Dim blnCancelTranslation As Boolean
        
        On Error GoTo Error_Handler
        
        Screen.MousePointer = vbHourglass
        
10      strMdbPath = GetMDBPath(ADOConnection.ConnectionString)
        
100     If g_clsLog Is Nothing Then
110         Set g_clsLog = New CAppLogFile
120         g_clsLog.StartLogFile Application.ProductName, strMdbPath, NavigationPane.UserID, g_blnTraceFileOn
        End If
        
130     g_clsLog.AppendLog "Clear tree nodes & assign imagelist to treeviews"
140     frmPane.tvwShortcut.Nodes.Clear
150     frmPane.tvwMain.Nodes.Clear
160     frmPane.tvwMain.ImageList = g_typInterface.ITreeIcons
        frmPane.tvwMain.UseImageList = True
        
170     frmPane.tvwShortcut.ImageList = g_typInterface.ITreeIcons

        
Repeat_Process:

370     Set rstButtons = New ADODB.Recordset
            
380     g_clsLog.AppendLog "Retrieve shortcut buttons"

390     enuResult = NavigationPane.SelectShortcutButtons(rstButtons, ADOConnection)
        
        Select Case enuResult
            Case QueryResultSuccessful
400             g_clsLog.AppendLog "Shortcut buttons was retrieved"

                If frmPane.tvwShortcut.Nodes.Count = 1 Then
                    If frmPane.tvwShortcut.Nodes(1).Key = "NoButtons" Then
                        frmPane.tvwShortcut.Nodes.Remove 1
                    End If
                End If

                Do While Not rstButtons.EOF
                    ' If this is to ensure that the shortcut button is a valid folder then
                    ' this is already incorporated in the SQL found in
                    'NavigationPane.SelectShortcutButtons
                    'For lngCtr = 1 To frmPane.tvwMain.Nodes.Count
                    '    If frmPane.tvwMain.Nodes(lngCtr).key = CStr(rstButtons!Node_ID) Then
                            RaiseEvent BeforeAddFolder(ADOConnection, rstButtons!Node_ID, blnCancelTranslation)
410                         frmPane.tvwShortcut.Nodes.Add , , CStr(rstButtons!Node_ID), IIf(blnCancelTranslation, rstButtons!Button_Caption, Translate(rstButtons!Button_Caption)), CStr(rstButtons!Node_Image), CStr(rstButtons!Node_Image)
'                            frmPane.tvwShortcut.Nodes.Add , , CStr(rstButtons!Node_ID), IIf(blnCancelTranslation, Translate(rstButtons!Button_Caption), rstButtons!Button_Caption), frmPane.tvwMain.Nodes(CStr(rstButtons!Node_ID)).Image, frmPane.tvwMain.Nodes(CStr(rstButtons!Node_ID)).Image
                    '        Exit For
                    '    End If
                    'Next
420                 rstButtons.MoveNext
                Loop
                Call frmPane.ResizeFavoriteSection
                
430             g_clsLog.AppendLog "Shortcut buttons were added in the favorites folder"
                
            Case QueryResultNoRecord
440             g_clsLog.AppendLog "No shortcut buttons defined for the selected tree. Add default shortcut buttons"
                enuResult = NavigationPane.LoadDefaultShortcutButtons(ADOConnection)
                Select Case enuResult
                    Case QueryResultSuccessful
441                     g_clsLog.AppendLog "Default shortcut buttons were added"
                        GoTo Repeat_Process
                        
                    Case QueryResultNoRecord
                        frmPane.tvwShortcut.Nodes.Add , , "NoButtons", " Drag favorite folders here..."
442                     g_clsLog.AppendLog "No default shortcut buttons were added"
                        
                    Case QueryResultError
443                     g_clsLog.AppendLog "Error occurs while adding default shortcut buttons"
                        
                End Select
                
                Call frmPane.ResizeFavoriteSection
                
                
            Case QueryResultError
450             g_clsLog.AppendLog "Error occured while retrieving shortcut buttons"
                MsgBox "Error has occured while trying to load 'Favorite Folders'. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: SelectShortcutButtons Module"
                        
        End Select
        
460     Set rstButtons = Nothing
        
        frmPane.Refresh
        
        
'180     Set rstNodes = New ADODB.Recordset
'
'190     g_clsLog.AppendLog "Start retrieving tree nodes from database: LoadTrees Method"
'200     g_clsLog.AppendLog "Tree ID: " & NavigationPane.TreeID
'210     enuResult = NavigationPane.LoadTreeNodes(rstNodes, NavigationPane.TreeID, ADOConnection)
'220     g_clsLog.AppendLog "End LoadTrees method"
        
'        Select Case enuResult
'            Case QueryResultSuccessful
                frmPane.tvwMain.Redraw = False
240             frmPane.tvwMain.UseImageList = True
                frmPane.tvwMain.Tag = NavigationPane.TreeID
                frmPane.tvwMain.Refresh
                
                lngCtr = 0
                Do While lngCtr < frmPane.tvwMain.Nodes.Count
                    lngCtr = lngCtr + 1
                    frmPane.tvwMain.Nodes(lngCtr).Expanded = ExpandedNode(Val(frmPane.tvwMain.Nodes(lngCtr).Key), NavigationPane, ADOConnection)
                Loop
                
                frmPane.tvwMain.Redraw = True
                
'            Case QueryResultNoRecord
'340             g_clsLog.AppendLog "No nodes for the selected tree was retrieved"
'
'            Case QueryResultError
'350             g_clsLog.AppendLog "Error occured while retrieving nodes for the selected tree"
'
'                MsgBox "Error has occured while trying to load 'All Folders'. " & _
'                        "Please contact technical support.", vbInformation + vbOKOnly, _
'                        "Main Setting Error: SelectShortcutButtons Module"
'                Exit Sub
'
'        End Select
'
'360     Set rstNodes = Nothing

        Screen.MousePointer = vbDefault
        
470     g_clsLog.CloseLogFile
        
        Exit Sub
    
Error_Handler:
    
    g_clsLog.AppendLog "Error occured while loading tree nodes"
    g_clsLog.AppendLog "Line Number: " & Erl
    g_clsLog.AppendLog "Error Number: " & Err.Number
    g_clsLog.AppendLog "Description: " & Err.Description
    g_clsLog.CloseLogFile

    MsgBox "Error has occured while trying to load 'All Folders'. " & _
            "Please contact technical support.", vbInformation + vbOKOnly, _
            "Main Setting Error: LoadTrees Module"
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub SaveNavigationPaneSettings(ByVal Application As Object, _
                    ByVal UserID As Long, _
                    ByRef NavigationPane As CNavigationPane, _
                    ByRef GridProps As CGrid, _
                    ByRef ADOConnection As ADODB.Connection)

    Call SaveTreeSettings(Application, NavigationPane, GridProps, ADOConnection)
    
End Sub

Private Sub SaveTreeSettings(ByVal Application As Object, _
                        ByRef NavigationPane As CNavigationPane, _
                        ByRef GridProps As CGrid, _
                        ByRef ADOConnection As ADODB.Connection)
        
        Dim enuResult As QueryResultConstants
        Dim strMdbPath As String
        
        Screen.MousePointer = vbHourglass
        
       On Error GoTo Error_Handler
       
10      strMdbPath = GetMDBPath(ADOConnection.ConnectionString)
        
100     If g_clsLog Is Nothing Then
110         Set g_clsLog = New CAppLogFile
120         g_clsLog.StartLogFile Application.ProductName, strMdbPath, NavigationPane.UserID, g_blnTraceFileOn
        End If
        
130     g_clsLog.AppendLog "Retrieve selected group's record"
140     g_clsLog.AppendLog "User ID: " & NavigationPane.UserID
150     g_clsLog.AppendLog "Tree ID: " & NavigationPane.TreeID
160     enuResult = NavigationPane.SelectGroup(NavigationPane.UserID, _
                                            NavigationPane.TreeID, ADOConnection, True)
        
        Select Case enuResult
            Case QueryResultSuccessful
170             g_clsLog.AppendLog "Selected group record was retrieved"
180             g_clsLog.AppendLog "Update group settings"
190             If NavigationPane.UpdateGroupSettings(ADOConnection) = QueryResultError Then
                    MsgBox "Error has occured while trying to access the database. " & _
                            "Please contact technical support.", vbInformation + vbOKOnly, _
                            "Main Setting Error: UpdateGroupSettings Module"
                End If
                
            Case QueryResultNoRecord
200             g_clsLog.AppendLog "No record for the selected group was found in the database"
210             g_clsLog.AppendLog "Add selected group record in the database"
220             If NavigationPane.AddGroup(ADOConnection) = QueryResultError Then
                    MsgBox "Error has occured while trying to access the database. " & _
                            "Please contact technical support.", vbInformation + vbOKOnly, _
                            "Main Setting Error: AddGroup Module"
                End If
                
            Case QueryResultError
230             g_clsLog.AppendLog "Error occured while retrieving selected group record in the database"
                MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: SelectGroup Module"
                        
        End Select
        
240     g_clsLog.AppendLog "Start saving expanded/collapsed nodes"
250     SaveExpandedNodes NavigationPane, ADOConnection
260     g_clsLog.AppendLog "End saving"
        
270     g_clsLog.AppendLog "Start saving view settings"
280     SaveView NavigationPane, GridProps, ADOConnection
290     g_clsLog.AppendLog "End saving"
        
        Screen.MousePointer = vbDefault
        
300     g_clsLog.CloseLogFile
        
        Exit Sub
        
Error_Handler:

    g_clsLog.AppendLog "Error occured while saving tree settings"
    g_clsLog.AppendLog "Line Number: " & Erl
    g_clsLog.AppendLog "Error Number: " & Err.Number
    g_clsLog.AppendLog "Description: " & Err.Description
    g_clsLog.CloseLogFile
    
End Sub


Friend Sub SaveExpandedNodes(ByRef NavigationPane As CNavigationPane, ByRef ADOConnection As ADODB.Connection)

    Dim strExpandedNodes As String
    Dim enuResult As QueryResultConstants
    
    
    strExpandedNodes = ExpandedNodes
    enuResult = NavigationPane.SelectTreeSettings(ADOConnection)
    
    
    Select Case enuResult
        
        Case QueryResultSuccessful
            If NavigationPane.UpdateTreeSettings(strExpandedNodes, ADOConnection) = QueryResultError Then
                MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: UpdateTreeSettings Module"
            End If
            
        Case QueryResultNoRecord
            If NavigationPane.AddTreeSettings(strExpandedNodes, ADOConnection) = QueryResultError Then
                MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: AddTreeSettings Module"
            End If
        
        Case QueryResultError
            MsgBox "Error has occured while trying to access the database. " & _
                    "Please contact technical support.", vbInformation + vbOKOnly, _
                    "Main Setting Error: SelectTreeSettings Module"
    
    End Select

End Sub

Friend Function ExpandedNode(NodeID As Long, NavigationPane As CNavigationPane, ByRef ADOConnection As ADODB.Connection) As Boolean

    Dim rstTreeSetting As ADODB.Recordset
    Dim lngIndex As Long
    Dim lngCharPos As Long
    Dim lngNodeID As Long
    Dim enuResult As QueryResultConstants
    
Reapeat_Procedure:
    enuResult = NavigationPane.GetExpandedNodes(ADOConnection, rstTreeSetting)
    
    Select Case enuResult
        Case QueryResultSuccessful
            If Not rstTreeSetting.EOF Then
                lngCharPos = 1
                Do While lngCharPos < Len(Trim(rstTreeSetting!TreeSet_ExpandedNodes))
                    lngNodeID = Val(Mid(Trim(rstTreeSetting!TreeSet_ExpandedNodes), lngCharPos, 4))
                    If lngNodeID = NodeID Then
                        ExpandedNode = True
                        Set rstTreeSetting = Nothing
                        Exit Function
                    End If
                    lngCharPos = lngCharPos + 5
                Loop
            End If
            
        Case QueryResultNoRecord
            SaveExpandedNodes NavigationPane, ADOConnection
            GoTo Reapeat_Procedure
        
        Case QueryResultError
            MsgBox "Error has occured while trying to access the database. " & _
                    "Please contact technical support.", vbInformation + vbOKOnly, _
                    "Main Setting Error: GetExpandedNodes Module"
        
    End Select
    
    Set rstTreeSetting = Nothing
    
    ExpandedNode = False
    
End Function


Private Function ExpandedNodes() As String

    Dim lngNodeIndex As Long
    Dim lngNodeLevel As Long
    Dim lngNodeID As Long
    Dim strExpandedNodes As String
    
    On Error GoTo Error_Handler
    
    strExpandedNodes = vbNullString
    
    For lngNodeIndex = 1 To frmPane.tvwMain.Nodes.Count
        'Mod by BCo 2006-05-18
        'Added children node checking, so expanded child-less nodes won't be considered when writing to DB
        If frmPane.tvwMain.Nodes(lngNodeIndex).Expanded = True And frmPane.tvwMain.Nodes(lngNodeIndex).Children > 0 Then
            lngNodeLevel = frmPane.tvwMain.Nodes(lngNodeIndex).Level
            lngNodeID = Val(frmPane.tvwMain.Nodes(lngNodeIndex).Key)
            strExpandedNodes = strExpandedNodes & AddSpaceZero(0, CStr(lngNodeID), 4, True) & lngNodeLevel
        End If
    Next
    
    ExpandedNodes = strExpandedNodes
    
    Exit Function
    
Error_Handler:

    MsgBox "Error has occured while trying to get open folders. " & _
            "Please contact technical support.", vbInformation + vbOKOnly, _
            "Main Setting Error: ExpandedNodes Module"

End Function


Friend Function LoadViews(ByVal Application As Object, _
                ByRef NavigationPane As CNavigationPane, _
                ByRef GridProps As CGrid, _
                ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    
        Dim clsViews As CNavigationPane
        Dim rstViews As ADODB.Recordset
        Dim lngCtr As Long
        Dim lngCount As Long
        Dim lngSelectedView As Long
        Dim strMdbPath As String
        Dim enuResult As QueryResultConstants
        
        On Error GoTo Error_Handler
        Screen.MousePointer = vbHourglass
        
10      strMdbPath = GetMDBPath(ADOConnection.ConnectionString)
        
100     If g_clsLog Is Nothing Then
110         Set g_clsLog = New CAppLogFile
120         g_clsLog.StartLogFile Application.ProductName, strMdbPath, NavigationPane.UserID, g_blnTraceFileOn
        End If
        
130     LoadViews = QueryResultNoRecord
        
140     If (Not frmPane.tvwMain.SelectedItem Is Nothing) Or (Not frmPane.tvwShortcut.SelectedItem Is Nothing) Then
            
150         g_clsLog.AppendLog "Unload all dynamically created view buttons"
160         For lngCtr = 0 To frmPane.Controls.Count - 1
170             If lngCtr < frmPane.Controls.Count Then
180                 If Mid(frmPane.Controls(lngCtr).Name, 1, 8) = "optViews" Then
190                     If frmPane.Controls(lngCtr).Index > 0 Then
200                         Unload frmPane.Controls(lngCtr)
210                         lngCtr = lngCtr - 1
                        End If
                    End If
                End If
            Next
220         g_clsLog.AppendLog "All view options were unloaded"
            
230         g_clsLog.AppendLog "Load all view options for the selected node"
240         g_clsLog.AppendLog "Node ID: " & NavigationPane.NodeID
250         enuResult = NavigationPane.LoadViewOptions(NavigationPane.NodeID, ADOConnection, rstViews)
            
          
            Select Case enuResult
                Case QueryResultSuccessful
                    
260                 g_clsLog.AppendLog "View options were retieved"
270                 g_clsLog.AppendLog "Get selected view"
280                 enuResult = NavigationPane.NodeView(NavigationPane.UserID, _
                                NavigationPane.NodeID, ADOConnection, lngSelectedView)
                    
                    Select Case enuResult
                        Case QueryResultSuccessful
290                         g_clsLog.AppendLog "Start loading view options"
300                         Call CreateViewOptions(GridProps, rstViews, lngSelectedView)
310                         g_clsLog.AppendLog "View options were loaded in the tree"
                            
                        Case QueryResultNoRecord
                            
320                         g_clsLog.AppendLog "No selected view"
330                         g_clsLog.AppendLog "Create default grid settings for the view"
                            
340                         If GridProps.LoadDefaultSettings(NavigationPane, ADOConnection) = QueryResultSuccessful Then
350                             g_clsLog.AppendLog "Default settings for the node were added"
360                             g_clsLog.AppendLog "Start loading view options"
370                             Call CreateViewOptions(GridProps, rstViews, 1)
380                             g_clsLog.AppendLog "View options were loaded in the tree"
390                             NavigationPane.NodeID = GridProps.NodeID
                            Else
400                             GridProps.NodeID = NavigationPane.NodeID
410                             GridProps.ViewID = 0
                            End If
                            
                        Case QueryResultError
                            MsgBox "Error has occured while trying to access the database. " & _
                                    "Please contact technical support.", vbInformation + vbOKOnly, _
                                    "Main Setting Error: NodeView Module"
                                    
                    End Select
                
    '420             Set frmPane.tvwMain.TopNode = frmPane.tvwMain.Nodes(CStr(NavigationPane.TopNode))
                    
                Case QueryResultNoRecord
                    '>> No view options for the selected node were defined in the database
                    GridProps.UVC_ID = 0
                    GridProps.TView_ID = 0
                    GridProps.NodeID = NavigationPane.NodeID
                    
                    GridProps.Alignments = ""
                    GridProps.AutoGroup = False
                    GridProps.CardView = False
                    GridProps.DVCIDs = ""
                    GridProps.ExpandCollapseDefault = 0
                    GridProps.GridLines = False
                    GridProps.GroupHeaderIsVisible = False
                    GridProps.GroupHeaders = ""
                    GridProps.GroupRowFormat = ""
                    GridProps.OddEvenColor = False
                    GridProps.SelectedRows = ""
                    GridProps.ReadingPane = False
                    GridProps.RequiredFields = ""
                    GridProps.Sort = ""
                    GridProps.ViewID = 0
                    GridProps.Widths = ""
                    
                Case QueryResultError
                    MsgBox "Error has occured while trying to access the database. " & _
                            "Please contact technical support.", vbInformation + vbOKOnly, _
                            "Main Setting Error: LoadViewOptions Module"
                            
            End Select
        End If
        
        LoadViews = enuResult
        
        Screen.MousePointer = vbDefault
         
        g_clsLog.CloseLogFile
        
        Exit Function
    
Error_Handler:

    g_clsLog.AppendLog "Error occured while saving tree settings"
    g_clsLog.AppendLog "Line Number: " & Erl
    g_clsLog.AppendLog "Error Number: " & Err.Number
    g_clsLog.AppendLog "Description: " & Err.Description
    g_clsLog.CloseLogFile
    
    MsgBox "Error has occured while trying to load view options. " & _
            "Please contact technical support.", vbInformation + vbOKOnly, _
            "Main Setting Error: LoadViews Module"
    LoadViews = QueryResultError
    
    Screen.MousePointer = vbDefault
End Function

Private Sub CreateViewOptions(ByRef GridProps As CGrid, ByVal ViewRecords As ADODB.Recordset, ByVal SelectedViewPos As Long)

    Dim lngTop As Long
    Dim lngWidth As Long
    
    lngTop = 120
    frmPane.fraViews.Height = 0
    
    Do While Not ViewRecords.EOF
        Load frmPane.optViews(ViewRecords!View_Position)
        frmPane.optViews(ViewRecords!View_Position).Appearance = 0
        frmPane.optViews(ViewRecords!View_Position).Tag = ViewRecords!View_ID
        frmPane.optViews(ViewRecords!View_Position).Caption = Trim(ViewRecords!View_Caption)
        frmPane.optViews(ViewRecords!View_Position).Font = "Arial"
        frmPane.optViews(ViewRecords!View_Position).FontSize = 8
        lngWidth = Len(Trim(ViewRecords!View_Caption)) * 260
        frmPane.optViews(ViewRecords!View_Position).Move 120, lngTop, lngWidth, 300
        
        If SelectedViewPos = ViewRecords!View_Position Then
            GridProps.NodeID = ViewRecords!Node_ID
            GridProps.ViewID = ViewRecords!View_ID
            frmPane.optViews(ViewRecords!View_Position).Value = True
        Else
            frmPane.optViews(ViewRecords!View_Position).Value = False
        End If
        
        On Error Resume Next
        frmPane.optViews(ViewRecords!View_Position).Visible = True
        On Error GoTo 0
        
        lngTop = lngTop + 315
                    
        frmPane.fraViews.Height = frmPane.fraViews.Height + 400
        frmPane.tvwMain.Visible = True
        
        ViewRecords.MoveNext
    Loop

End Sub

'Friend Sub SaveTopNode(ByVal TopNode As SSNode)
'
'    Dim conMain As ADODB.Connection
'    Dim clsNavigation As CNavigationPane
'    Dim clsGrid As CGrid
'    Dim clsMainProps As CMainControls
'    Dim objMain As Object
'    Dim blnCancel As Boolean
'
'    RaiseEvent NodeEvents(objMain, clsMainProps, clsNavigation, clsGrid, conMain, blnCancel)
'
'    If Not blnCancel Then
'        clsNavigation.TopNode = Val(TopNode.Key)
'        SaveTreeSettings clsNavigation, clsGrid, conMain
'    End If
'
'End Sub

Friend Sub TriggerNodeClickEvent(ByRef SelectedNode As SSNode, Optional ChangeViewOnly As Boolean = False, Optional ForceRefresh As Boolean = False)
    
    Dim conMain As ADODB.Connection
    Dim clsNavigation As CNavigationPane
    Dim clsGrid As CGrid
    Dim clsLoadGrid As IGrid
    Dim clsMain As IMainControls
    Dim clsMainProps As CMainControls
    Dim objMain As Object
    Dim objApplication As Object
    Dim lngUserID As Long
    Dim lngCtr As Long
    Dim blnCancel As Boolean
    
    If m_blnDontRefresh Then Exit Sub
    
    Set clsLoadGrid = New IGrid
    
    '>> to retieve selected node
    Set clsNavigation = New CNavigationPane
    
    If SelectedNode Is Nothing = False Then
        
        ' Selected node is in the main treeview
        Set clsNavigation.SelectedNode = SelectedNode
        
    Else
        
        If (clsNavigation.tvwMainControl.SelectedItem Is Nothing = True) Then
            If (clsNavigation.tvwShortcutControl Is Nothing = False) Then
                ' Selected node is in the shortcut/favorites treeview
                
                Set clsNavigation.SelectedNode = clsNavigation.tvwShortcutControl.SelectedItem
                Set SelectedNode = clsNavigation.SelectedNode
            End If
        End If
        
    End If
    '>> End
    
    RaiseEvent NodeEvents(objApplication, objMain, clsMainProps, clsNavigation, clsGrid, conMain, blnCancel)
    
' 2dit Commented namn ung nasa loob ng if condition so no need to perform the If line
'    If Val(SelectedNode.key) <> clsNavigation.NodeID Then
''        objMain.txtSearchKey.Text = ""
''        objMain.cbocolumns.ListIndex = -1
'    End If
    
'    If ForceRefresh = False And Not SelectedNode Is Nothing Then
'        If (clsNavigation.NodeID = Val(SelectedNode.Key)) And (objMain.Visible = True) Then
'            Exit Sub
'        End If
'    End If
    
    '>> Save expanded/collapsed groups
    clsLoadGrid.GroupExpandCollapse clsGrid, objMain
    '>> Save last selected row
    clsLoadGrid.GetLastSelectedRows clsGrid, objMain
    '>> Save grid setting
    clsLoadGrid.SaveGridSetting clsNavigation, clsGrid, conMain
    '>> Save selected view
    SaveView clsNavigation, clsGrid, conMain
        
'    If blnCancel = True Then
'        Set SelectedNode = frmPane.tvwMain.Nodes(CStr(clsNavigation.NodeID)).Root
'        Set frmPane.tvwMain.DropHighlight = SelectedNode
'    End If

    If Not SelectedNode Is Nothing Then
        g_typInterface.IGridHeader.Caption = " " & Replace(SelectedNode.Text, "&", "&&")
'       objMain.imgHeader.Picture = objMain.imgListTree.ListImages(SelectedNode.SelectedImage).Picture
    End If

    If ChangeViewOnly = False Then
        '>> Selected Node was changed
        
        clsNavigation.TopNode = Val(frmPane.tvwMain.TopNode.Key)
        g_lngTopNode = clsNavigation.TopNode
        clsNavigation.NodeID = Val(SelectedNode.Key)
        Set clsNavigation.SelectedNode = SelectedNode
        clsGrid.NodeID = Val(SelectedNode.Key)
        clsNavigation.UpdateGroupSettings conMain
        m_blnDontRefresh = True
        '>> load view options for the selected node
        LoadViews objApplication, clsNavigation, clsGrid, conMain
        m_blnDontRefresh = False
    Else
        
        '>> Get selected view
        clsGrid.ViewID = 0
        For lngCtr = 0 To frmPane.Controls.Count - 1
            If Left(frmPane.Controls(lngCtr).Name, 8) = "optViews" Then
                If frmPane.optViews(frmPane.Controls(lngCtr).Index).Value = True Then
                    clsGrid.ViewID = Val(frmPane.optViews(frmPane.Controls(lngCtr).Index).Tag)
                    Exit For
                End If
            End If
        Next
    End If
    
'    Set clsMain = New IMainControls
'    clsMain.ResizeControls objMain, clsMainProps, clsGrid, conMain
    
    If blnCancel = False Then
        '>> Load grid records
        clsLoadGrid.LoadGrid objMain, clsMainProps, clsNavigation, clsGrid, conMain
    Else
        g_typInterface.IGrid.GroupByBoxVisible = False

        If g_typInterface.IReadingPane.Visible = True Then
            If (g_typInterface.INoView.Width + g_typInterface.INoView.Left) _
                > g_typInterface.IReadingPane.Left Then
                    g_typInterface.INoView.Width = g_typInterface.IReadingPane.Left - g_typInterface.INoView.Left
            Else
                g_typInterface.INoView.Width = 2805
            End If
        End If
        
        g_typInterface.IStatusbar.Panels(1).Text = ""
        
        Set g_typInterface.IGrid.ADORecordset = Nothing

        g_typInterface.IGrid.Columns.Clear
        
        g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
        g_typInterface.INoView.Left = 300
                
        g_typInterface.INoView.ZOrder 0
        g_typInterface.INoView.Caption = "No View"
    End If
    
    Set clsMain = New IMainControls
    clsMain.LoadArrangeByButtons objMain, clsGrid
    Set clsMain = Nothing
    
    
    If g_typInterface.IGrid.ADORecordset Is Nothing Then
        clsGrid.ReadingPane = False
        clsGrid.GroupHeaderIsVisible = False
    End If
    
    RaiseEvent NodeClick(clsNavigation)
    
    ' On Error is provided to handle the case when the application is just starting up and the main window is not yet loaded
    On Error Resume Next
    g_typInterface.IGrid.SetFocus
    On Error GoTo 0
    
    Set clsLoadGrid = Nothing
    Set conMain = Nothing
    Set objMain = Nothing
'    Set clsMain = Nothing
End Sub

Public Function CustomizeView() As Boolean
    
    Dim conView As ADODB.Connection
    Dim clsGrid As CGrid
    Dim clsNavigationPane As CNavigationPane
    Dim objMain As Object
    
    Dim lngControls As Long
    Dim strCaption As String
    
    '>> load settings of the selected view
    For lngControls = 0 To frmPane.Controls.Count - 1
        If Mid(frmPane.Controls(lngControls).Name, 1, 8) = "optViews" Then
            If frmPane.Controls(lngControls).Value = True Then
                strCaption = frmPane.Controls(lngControls).Caption
                Exit For
            End If
        End If
    Next
        
    RaiseEvent CustomView(objMain, clsGrid, clsNavigationPane, conView)
    
    frmEditView.ShowForm objMain, clsGrid, clsNavigationPane, conView, strCaption, CustomizeView
    
    Set objMain = Nothing
    Set clsGrid = Nothing
    Set conView = Nothing
    
End Function

Friend Sub SaveView(ByRef NavigationPane As CNavigationPane, ByRef GridProps As CGrid, _
                    ByRef ADOConnection As ADODB.Connection)
                    
    Dim enuResult As QueryResultConstants
    
    enuResult = NavigationPane.SelectViewSettings(GridProps, ADOConnection)
    
    Select Case enuResult
        Case QueryResultSuccessful
            If NavigationPane.UpdateNodeView(GridProps, ADOConnection) = QueryResultError Then
                MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: SaveNodeView Module"
            End If
            
        Case QueryResultNoRecord
        
        Case QueryResultError
        
    End Select

End Sub

Friend Function AddShortcut(ByVal NodeID As Long, ByVal ShortcutPos As Long) As Boolean

    Dim conMain As ADODB.Connection
    Dim clsNavigation As CNavigationPane
    Dim clsGrid As CGrid
    Dim clsMainProps As CMainControls
    Dim rstShortcut As ADODB.Recordset
    Dim enuResult As QueryResultConstants
    Dim objMain As Object
    Dim objApplication As Object
    Dim blnShortcutExists As Boolean
    Dim blnCancel As Boolean
    
    RaiseEvent NodeEvents(objApplication, objMain, clsMainProps, clsNavigation, clsGrid, conMain, blnCancel)
    
    AddShortcut = Not blnCancel
    
    If AddShortcut Then
        clsNavigation.NodeID = NodeID
        
        blnShortcutExists = False
        
        enuResult = clsNavigation.SelectShortcutButtons(rstShortcut, conMain)
        
        Select Case enuResult
            Case QueryResultSuccessful
                Do While Not rstShortcut.EOF
                    If rstShortcut!Node_ID = NodeID Then
                        blnShortcutExists = True
                        Exit Do
                    End If
                    
                    rstShortcut.MoveNext
                Loop
                
            Case QueryResultNoRecord
                blnShortcutExists = False
                
            Case QueryResultError
            
        End Select
        
            
        If blnShortcutExists = False Then
            
            If frmPane.tvwShortcut.Nodes.Count = 1 Then
                If frmPane.tvwShortcut.Nodes(1).Key = "NoButtons" Then
                    frmPane.tvwShortcut.Nodes.Remove 1
                End If
            End If
            
            Call clsNavigation.AddShortcutButton(conMain, ShortcutPos)
            
            Call LoadShortcutButtons(conMain, clsNavigation)
            
            Call frmPane.ResizeFavoriteSection
            
        End If
    End If
    
    Set rstShortcut = Nothing
    Set conMain = Nothing
    Set objMain = Nothing
    
End Function

Public Sub DeleteSelectedShortcutButton()
    Dim intAns As Integer
    
    If frmPane.tvwShortcut.SelectedItem.Key <> "NoButtons" Then
        '>> delete shortcut button
        intAns = MsgBox("Are you sure you want to remove '" & frmPane.tvwShortcut.SelectedItem & "' folder in the list of your 'Favorite' folders?", vbQuestion + vbYesNo, "Remove Confirmation")
        If intAns = vbYes Then
            DeleteShortcut
        End If
    End If
End Sub

Friend Sub DeleteShortcut()

    Dim conMain As ADODB.Connection
    Dim clsNavigation As CNavigationPane
    Dim clsMainProps As CMainControls
    Dim clsGrid As CGrid
    Dim objMain As Object
    Dim objApplication As Object
    Dim blnCancel As Boolean
    
    RaiseEvent NodeEvents(objApplication, objMain, clsMainProps, clsNavigation, clsGrid, conMain, blnCancel)
    
    If Not blnCancel Then
        
        'clsNavigation.DeleteShortcutButton conMain, frmPane.tvwShortcut.Nodes(CStr(clsNavigation.NodeID)).Index
        '>>>>> p4tric 121907
        clsNavigation.DeleteShortcutButton conMain, frmPane.tvwShortcut.SelectedItem.Index
        
        Call LoadShortcutButtons(conMain, clsNavigation)
        Call frmPane.ResizeFavoriteSection
        
        If frmPane.tvwShortcut.Nodes.Count = 0 Then
            frmPane.tvwShortcut.Nodes.Add , , "NoButtons", " Drag favorite folders here..."
        End If
        
    End If
    
    Set clsNavigation = Nothing
    Set clsMainProps = Nothing
    Set clsGrid = Nothing
    Set objApplication = Nothing
    Set conMain = Nothing
    Set objMain = Nothing
    
End Sub

Friend Sub ShowPopupMenu(SelectedFavoriteFolder As SSActiveTreeView.SSNode)

    RaiseEvent PopupMenu(SelectedFavoriteFolder)
    
End Sub
Public Sub ProcessPopupEvents(ProcessType As FaveFolderProcess)
    
    Dim conMain As ADODB.Connection
    Dim clsNavigation As CNavigationPane
    Dim clsMainProps As CMainControls
    Dim clsGrid As CGrid
    Dim objMain As Object
    Dim objApplication As Object
    Dim blnCancel As Boolean
    Dim intAns As Integer
    
    RaiseEvent NodeEvents(objApplication, objMain, clsMainProps, clsNavigation, clsGrid, conMain, blnCancel)
    
    Select Case ProcessType
        Case Folder_Add
            
        Case Folder_Modify
            frmPane.tvwShortcut.StartLabelEdit
            
        Case Folder_Delete
            If frmPane.tvwShortcut.SelectedItem.Key <> "NoButtons" Then
                clsNavigation.NodeID = frmPane.tvwShortcut.SelectedItem.Key
                intAns = MsgBox("Are you sure you want to remove '" & frmPane.tvwShortcut.SelectedItem & "' folder in the list of your 'Favorite' folders?", vbQuestion + vbYesNo, "Remove Confirmation")
                
                If intAns = vbYes Then
                    
                    Call clsNavigation.DeleteShortcutButton(conMain, frmPane.tvwShortcut.Nodes(CStr(clsNavigation.NodeID)).Index)
                    
                    frmPane.tvwShortcut.Nodes.Remove frmPane.tvwShortcut.Nodes(CStr(clsNavigation.NodeID)).Index
                    
                    Call frmPane.ResizeFavoriteSection
                    
                    If frmPane.tvwShortcut.Nodes.Count = 0 Then
                        frmPane.tvwShortcut.Nodes.Add , , "NoButtons", " Drag favorite folders here..."
                    End If
                    
                End If
            End If
            
        Case Else '>> Save new caption
            Call clsNavigation.RenameShortcutButton(conMain)
    End Select
    
    Set objMain = Nothing

End Sub

Friend Sub UpdateTopNode(ByVal NewTopNodeID As Long)

    Dim clsTopNode As CNavigationPane
    
    RaiseEvent TopNodeChange(clsTopNode)
    
    clsTopNode.TopNode = NewTopNodeID
    
End Sub

Public Sub ChangeSelectedNode(ByVal New_NodeID As Long)

    Set frmPane.tvwMain.SelectedItem = frmPane.tvwMain.Nodes(CStr(New_NodeID))
    frmPane.tvwMain.DropHighlight = frmPane.tvwMain.SelectedItem
    
    Call TriggerNodeClickEvent(frmPane.tvwMain.SelectedItem, False, True)
    
End Sub

Public Sub StartLabelEdit(Treeview As TreeviewControl)

    If Treeview = TREEVIEW_MAIN Then
        frmPane.tvwMain.StartLabelEdit
    ElseIf Treeview = TREEVIEW_SHORTCUT Then
        frmPane.tvwShortcut.StartLabelEdit
    End If
    
End Sub

Public Sub SetTreeviewTag(Treeview As TreeviewControl, Tag As String)

    If Treeview = TREEVIEW_MAIN Then
        frmPane.tvwMain.Tag = Tag
    ElseIf Treeview = TREEVIEW_SHORTCUT Then
        frmPane.tvwShortcut.Tag = Tag
    End If

End Sub

Public Function GetTreeviewTag(Treeview As TreeviewControl) As String

    If Treeview = TREEVIEW_MAIN Then
        GetTreeviewTag = frmPane.tvwMain.Tag
    ElseIf Treeview = TREEVIEW_SHORTCUT Then
        GetTreeviewTag = frmPane.tvwShortcut.Tag
    End If

End Function



Friend Sub AfterLabelEdit(Cancel As Boolean, NewString As String, Treeview As TreeviewControl, ByRef ADOConnection As ADODB.Connection)

    If Treeview = TREEVIEW_MAIN Then
        RaiseEvent tvwMainAfterLabelEdit(Cancel, NewString, ADOConnection)
    Else
        RaiseEvent tvwShortcutAfterLabelEdit(Cancel, NewString, ADOConnection)
    End If
    
End Sub

Friend Sub BeforeNodeClick(Node As SSNode, Cancel As Boolean)

    RaiseEvent tvwMainBeforeNodeClick(Node, Cancel)
    
End Sub
Friend Sub Click()
    
    RaiseEvent tvwMainClick
    
End Sub
Friend Sub Collapse(Node As SSNode)

    RaiseEvent tvwMainCollapse(Node)
    
End Sub
Friend Sub EscapeLabelEdit(Cancel As Boolean, NewString As String)

    RaiseEvent tvwMainEscapeLabelEdit(Cancel, NewString)
    
End Sub
Friend Sub Expand(Node As SSNode)

    RaiseEvent tvwMainExpand(Node)
    
End Sub
Friend Sub KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent tvwMainKeyUp(KeyCode, Shift)
    
End Sub
Friend Sub MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent tvwMainMouseDown(Button, Shift, x, y)
    
End Sub
Friend Sub MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent tvwMainMouseMove(Button, Shift, x, y)
    
End Sub
Friend Sub MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent tvwMainMouseUp(Button, Shift, x, y)
    
End Sub

Friend Sub FavoritesMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent tvwShortcutMouseUp(Button, Shift, x, y)
End Sub

Friend Sub NodeClick(Node As SSNode)

    RaiseEvent tvwMainNodeClick(Node)
    
End Sub
Friend Sub OnDemandFetch(ByVal FetchBuffer As SSFetchBuffer)

    RaiseEvent tvwMainOnDemandFetch(FetchBuffer)
    
End Sub
Friend Sub OnDemandPrepare(ParentNode As SSNode, ByRef MainSettings As CMainControls, ByRef ADOConnection As ADODB.Connection, SortType As SSActiveTreeView.Constants_Sorted)

    RaiseEvent tvwMainOnDemandPrepare(ParentNode, MainSettings, ADOConnection, SortType)
    
End Sub
Friend Sub TopNodeChange(Node As SSNode)

    RaiseEvent tvwMainTopNodeChange(Node)
    
End Sub

Friend Sub TriggerViewOptionClick(Caption As String, Position As Integer)

    RaiseEvent ViewOptionClick(Caption, Position)
    
End Sub

Friend Sub TriggerBeforeAddFolder(ByRef ADOConnection As ADODB.Connection, ByVal NodeID As Long, ByRef CancelTranslation As Boolean)
    
    RaiseEvent BeforeAddFolder(ADOConnection, NodeID, CancelTranslation)
    
End Sub

Friend Sub LoadShortcutButtons(ByRef ADOConnection As ADODB.Connection, ByRef NavigationPane As CNavigationPane)
    Dim lngCtr As Long
    Dim rstShortcut As ADODB.Recordset
    Dim blnCancelTranslation As Boolean
    
    Call NavigationPane.SelectShortcutButtons(rstShortcut, ADOConnection)
    
    frmPane.tvwShortcut.Nodes.Clear
    
    Do While Not rstShortcut.EOF
        ' If this is to ensure that the shortcut button is a valid folder then
        ' this is already incorporated in the SQL found in
        ' NavigationPane.SelectShortcutButtons
        For lngCtr = 1 To frmPane.tvwMain.Nodes.Count
             If frmPane.tvwMain.Nodes(lngCtr).Key = CStr(rstShortcut!Node_ID) Then
                RaiseEvent BeforeAddFolder(ADOConnection, rstShortcut!Node_ID, blnCancelTranslation)
                frmPane.tvwShortcut.Nodes.Add , , CStr(rstShortcut!Node_ID), IIf(blnCancelTranslation, rstShortcut!Button_Caption, Translate(rstShortcut!Button_Caption)), frmPane.tvwMain.Nodes(CStr(rstShortcut!Node_ID)).Image, frmPane.tvwMain.Nodes(CStr(rstShortcut!Node_ID)).Image
                Exit For
            End If
        Next

        rstShortcut.MoveNext
    Loop
    
    Set rstShortcut = Nothing
    
End Sub

