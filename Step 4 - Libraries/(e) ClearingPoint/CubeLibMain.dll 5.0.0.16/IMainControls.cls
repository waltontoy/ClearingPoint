VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMainControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private m_blnAutoRefresh As Boolean
Private m_dblInterval As Double

Const MINIMUM_WIDTH_HEIGHT = 1000
Const MARGIN = 45
Const FIND_BOX_HEIGHT = 420

Public Enum MenuCommandConstant
    [File Exit]
    [Tools Find]
    [View Navigation Pane]
    [View Reading Pane Right]
    [View Reading Pane Bottom]
    [View Reading Pane Off]
    [View Group By Box]
    [View Card View]
    [View List View]
    [View Odd Even Color]
    [View Gridlines]
    [View Collapse All]
    [View Expand All]
    [View Collapse Group]
    [View Expand Group]
    [View Status Bar]
    [View Options]
    [View Customize Current View]
    [Tools Options]
    [Help Licensee]
    [Help About]
    [Context Rename Folder]
    [Context Remove Folder]
    [Product Activation]
    [Showfields]
    [Sort]
    [Group By]
    [Filter]
    [Other Settings]
    [Autoformat]
    [Format Columns]
    [Go Folder]
End Enum

Public Function LoadMainSetting(ByVal Application As Object, _
                            ByRef Window As Object, _
                            ByRef MainSettings As CMainControls, _
                            ByRef NodeProps As CGrid, _
                            ByRef Licensee As CLicensee, _
                            ByRef ADOConnection As ADODB.Connection, _
                            Optional CustomFindDefault As Boolean = False, _
                            Optional ByVal lngResourceHandler As Long = 0) As Boolean
    
    Dim clsRegistry As CRegistry
    Dim enuResult As QueryResultConstants
    Dim blnAutoRefresh As Boolean
    
    LoadMainSetting = True
    
    If lngResourceHandler <> 0 Then
        ResourceHandler = lngResourceHandler
    End If

'August 4, 2006 fix for windows 98 error when Mainsettings table is empty ==
Check_Settings_Again:
'==

    enuResult = MainSettings.LoadSetting(ADOConnection, MainSettings.UserID)
    
    Select Case enuResult
        Case QueryResultSuccessful
            '>> Load Main Settings
            g_typInterface.IButtonBar.ExpandedLinesCount = MainSettings.ExpandedButtons
            
            Call CreateFindBar(Window)
            
            Call ResizeControls(Window, MainSettings, NodeProps, ADOConnection)
            
            If g_typInterface.IButtonBar.FindItem(MainSettings.TreeID) Is Nothing Then
                g_typInterface.IButtonBar.Selected = g_typInterface.IButtonBar.Item(0)
                MainSettings.TreeID = g_typInterface.IButtonBar.Item(0).ID
            Else
                g_typInterface.IButtonBar.Selected = g_typInterface.IButtonBar.FindItem(MainSettings.TreeID)
            End If
            Call LoadArrangeByButtons(Window, NodeProps)
            
        Case QueryResultNoRecord
            
            'August 4, 2006 fix for windows 98 error when Mainsettings table is empty ==
            Call SaveDefaultMainSetting(ADOConnection, MainSettings.UserID, MainSettings.TreeID)
            GoTo Check_Settings_Again
            '==
            
            Call CreateFindBar(Window)
            
            '>> Load default settings
            If MainSettings.DefaultExpandedCount > 0 Then
                MainSettings.ExpandedButtons = MainSettings.DefaultExpandedCount
            Else
                MainSettings.ExpandedButtons = 1
            End If
            If MainSettings.DefaultNavPaneWidth > 0 Then
                MainSettings.NavigationPaneWidth = MainSettings.DefaultNavPaneWidth
            Else
                MainSettings.NavigationPaneWidth = 4000
            End If
            
            Window.Width = IIf(MainSettings.DefaultWindowWidth > 0, MainSettings.DefaultWindowWidth, 10000)
            Window.Height = IIf(MainSettings.DefaultWindowHeight > 0, MainSettings.DefaultWindowHeight, 8000)
            MainSettings.NavigationPane = True
            MainSettings.ViewOptions = False
            MainSettings.FindBox = g_typInterface.IFindBGround.Visible
            
            If MainSettings.UseCustomFind Then
                MainSettings.CustomFind = CustomFindDefault
            Else
                MainSettings.CustomFind = False
            End If
            
            MainSettings.CustomFind = CustomFindDefault
            MainSettings.StatusBar = True
            MainSettings.NavigationPaneWidth = MainSettings.NavigationPaneWidth
            MainSettings.GridWidth = MainSettings.NavigationPaneWidth / 2
            MainSettings.ReadingPanePos = 2
            MainSettings.CbarState = ""
            MainSettings.CBarVersion = 0
            MainSettings.ViewOptions = True
            g_typInterface.IButtonBar.ExpandedLinesCount = MainSettings.ExpandedButtons
            g_typInterface.IButtonBar.Selected = g_typInterface.IButtonBar.FindItem(1)
            
            Call ResizeControls(Window, MainSettings, NodeProps, ADOConnection)
            
        Case QueryResultError
            '>> Error in accessing database
            MsgBox "Error has occured while trying to access the database. Please contact technical support.", vbInformation + vbOKOnly, "Main Setting Error"
            
    End Select
        
    enuResult = NodeProps.LicenseeRecord(ADOConnection)
    
    Select Case enuResult
        Case QueryResultSuccessful
            '>> Licensee record exists
            
        Case QueryResultNoRecord
            '>> No Licensee info
            Call Licensee.ShowForm(Window, ADOConnection)
            If NodeProps.LicenseeRecord(ADOConnection) = QueryResultNoRecord Then
                LoadMainSetting = False
                Exit Function
            End If
            
        Case QueryResultError
            
    End Select
    
    Set clsRegistry = New CRegistry
    
    clsRegistry.GetRegistry cpiCurrentUser, Application.ProductName, "Settings", "RefreshManual"
    blnAutoRefresh = IIf(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) = "False", True, False)
    
    m_blnAutoRefresh = blnAutoRefresh
    g_typInterface.ITimer.Enabled = blnAutoRefresh
    
    clsRegistry.GetRegistry cpiCurrentUser, Application.ProductName, "Settings", "RefreshValue"
    g_typInterface.ITimer.Interval = IIf(IsNumeric(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)), Val(Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)), 10) * 1000
    m_dblInterval = g_typInterface.ITimer.Interval
    Set clsRegistry = Nothing
    
        
End Function

Private Sub SaveDefaultMainSetting(ByRef ADOConnection As ADODB.Connection, ByVal UserID As Long, ByVal TreeID As String)
    Dim strCommand As String
    
    
        strCommand = vbNullString
        strCommand = strCommand & "INSERT INTO "
        strCommand = strCommand & "MainSettings "
        strCommand = strCommand & "("
        strCommand = strCommand & "User_ID, "
        strCommand = strCommand & "Setting_Navigation, "
        strCommand = strCommand & "Setting_FindBox, "
        strCommand = strCommand & "Setting_StatusBar, "
        strCommand = strCommand & "Setting_ReadingPanePos, "
        strCommand = strCommand & "Setting_NavigationWidth, "
        strCommand = strCommand & "Setting_GridWidth, "
        strCommand = strCommand & "Setting_GridHeight, "
        strCommand = strCommand & "Setting_ExpandedButtons, "
        strCommand = strCommand & "Tree_ID, "
        strCommand = strCommand & "Setting_ViewOptions, "
        strCommand = strCommand & "Setting_CommandBar, "
        strCommand = strCommand & "Setting_CommandbarVersion, "
        strCommand = strCommand & "Setting_CustomFind"
        strCommand = strCommand & ") "
        strCommand = strCommand & "VALUES "
        strCommand = strCommand & "("
        strCommand = strCommand & UserID & ", "
        strCommand = strCommand & "True, "
        strCommand = strCommand & "False, "
        strCommand = strCommand & "True, "
        strCommand = strCommand & "2, "
        strCommand = strCommand & "4000, "
        strCommand = strCommand & "2000, "
        strCommand = strCommand & "0, "
        strCommand = strCommand & "1, "
        strCommand = strCommand & TreeID & ", "
        strCommand = strCommand & "True, "
        strCommand = strCommand & "'', "
        strCommand = strCommand & g_typInterface.ICommandBar.GlobalSettings.CommandBarsVersion & ", "
        strCommand = strCommand & "False"
        strCommand = strCommand & ")"
    
    ExecuteNonQuery ADOConnection, strCommand
    'ADOConnection.Execute strCommand
    
End Sub


Public Function SaveMainSettings(ByRef Window As Object, _
                                ByRef MainSettings As CMainControls, _
                                ByRef ADOConnection As ADODB.Connection) _
                                As QueryResultConstants

    Dim enuResult As QueryResultConstants
    Dim strCommandBarsVersion As String
    Dim strDecimalPoint As String
    
    strDecimalPoint = GetInfo(LOCALE_SDECIMAL)
    strCommandBarsVersion = g_typInterface.ICommandBar.GlobalSettings.CommandBarsVersion
    strCommandBarsVersion = Replace(strCommandBarsVersion, ".", strDecimalPoint)
    
    MainSettings.CbarState = g_typInterface.ICommandBar.SaveStateToString(True, True)
    MainSettings.CBarVersion = CDbl(strCommandBarsVersion)
    MainSettings.ExpandedButtons = g_typInterface.IButtonBar.ExpandedLinesCount
    
    enuResult = MainSettings.LoadSetting(ADOConnection, MainSettings.UserID, False)
    
    Select Case enuResult
        Case QueryResultSuccessful
            '>> Update Main setting
            If MainSettings.SaveSetting(ADOConnection) = QueryResultError Then
                GoTo Error_Handler
            End If
        Case QueryResultNoRecord
            '>> New main setting
            If MainSettings.AddSetting(ADOConnection) = QueryResultError Then
                GoTo Error_Handler
            End If
            
        Case QueryResultError
            '>> Error in accessing database
            GoTo Error_Handler
    
    End Select
    
    Exit Function
    
Error_Handler:
    MsgBox "Error has occured while trying to access the database. Please contact technical support.", vbInformation + vbOKOnly, "Main Setting Error"

End Function


Public Sub ResizeControls(ByRef Window As Object, _
                        ByRef MainSettings As CMainControls, _
                        ByRef NodeProps As CGrid, _
                        ByRef ADOConnection As ADODB.Connection)

    Dim lngTlbrLeft As Long
    Dim lngTlbrTop As Long
    Dim lngTlbrWidth As Long
    Dim lngTlbrHeight As Long
    
    Dim dblNavigationWidth As Double
    Dim dblNavigationHeight As Double
    
    Dim dblFindWidth As Double
    
    Dim dblItemlistLeft As Double
    Dim dblItemlistTop As Double
    Dim dblItemlistWidth As Double
    Dim dblItemlistHeight As Double
    
    Dim dblFindRemainingWidth As Double
    Dim dblFindControlWidth As Double
    
    Dim cbrButtonControl As CommandBarButton

    '>> Get toolbar dimension
    g_typInterface.ICommandBar.GetClientRect lngTlbrLeft, lngTlbrTop, lngTlbrWidth, lngTlbrHeight

    '>> Get Navigation Pane's width
    If MainSettings.NavigationPaneWidth > (Window.Width / 2) Then
        dblNavigationWidth = Window.Width / 2
    ElseIf MainSettings.NavigationPaneWidth < Window.Width Then
        dblNavigationWidth = MainSettings.NavigationPaneWidth
    Else
        dblNavigationWidth = Window.Width
    End If
    
    '>> Get Navigation Panes's height
    If (lngTlbrHeight - lngTlbrTop - (2 * MARGIN)) > 0 Then
        dblNavigationHeight = lngTlbrHeight - lngTlbrTop - (2 * MARGIN)
    Else
        dblNavigationHeight = 0
    End If
    
    '>> Set Navigation Pane's new dimension
    g_typInterface.IButtonBar.Move lngTlbrLeft + MARGIN, lngTlbrTop + MARGIN, dblNavigationWidth, dblNavigationHeight
    
    '>> Check visibility of Navigation Pane
    g_typInterface.IButtonBar.Visible = MainSettings.NavigationPane
    
    '>> Set Navigation Pane splitter's new dimension
    g_typInterface.ISplitterImage.Move lngTlbrLeft + dblNavigationWidth + MARGIN, lngTlbrTop + MARGIN, MARGIN, dblNavigationHeight
        
    '>> Check visibility of view options
    frmPane.fraViews.Visible = MainSettings.ViewOptions
    frmPane.sccViews.Visible = MainSettings.ViewOptions
    
    '>> Get find box width
    If MainSettings.NavigationPane = True Then
        If (lngTlbrWidth - lngTlbrLeft - dblNavigationWidth - MARGIN) > 0 Then
            dblFindWidth = lngTlbrWidth - lngTlbrLeft - dblNavigationWidth - MARGIN
        Else
            dblFindWidth = 0
        End If
        g_typInterface.IFindBGround.Move lngTlbrLeft + dblNavigationWidth + (2 * MARGIN), lngTlbrTop + MARGIN, dblFindWidth, FIND_BOX_HEIGHT
    Else
        If (lngTlbrWidth - lngTlbrLeft - (2 * MARGIN)) > 0 Then
            dblFindWidth = lngTlbrWidth - lngTlbrLeft - (2 * MARGIN)
        Else
            dblFindWidth = 0
        End If
        g_typInterface.IFindBGround.Move lngTlbrLeft, lngTlbrTop + MARGIN, dblFindWidth, FIND_BOX_HEIGHT
    End If
    
    
    '>> Find box' visibility
    g_typInterface.IFindBGround.Visible = MainSettings.FindBox
    If MainSettings.UseCustomFind Then
        g_typInterface.ICustomFind.Visible = MainSettings.CustomFind
    Else
        g_typInterface.ICustomFind.Visible = False
    End If
    
    '>> Get Itemlist background's dimensions
    If MainSettings.FindBox = True Then
    
        dblFindRemainingWidth = g_typInterface.IFindBGround.Width - 3500
        
        If dblFindRemainingWidth > 1800 Then
            dblFindControlWidth = (dblFindRemainingWidth / 15) / 2
            
            If dblFindControlWidth > 200 Then
                dblFindControlWidth = 200
            End If
            
            If dblFindControlWidth > 0 Then
                g_typInterface.IFind.ActiveMenuBar.Controls(1).Width = dblFindControlWidth '>> Look For
                g_typInterface.IFind.ActiveMenuBar.Controls(3).Width = dblFindControlWidth - 50 '>> Search In
                g_typInterface.IFind.RecalcLayout
            End If
        End If
        
        dblItemlistTop = g_typInterface.IFindBGround.Top + FIND_BOX_HEIGHT + MARGIN
        If (dblNavigationHeight - FIND_BOX_HEIGHT - MARGIN) > 0 Then
            dblItemlistHeight = dblNavigationHeight - FIND_BOX_HEIGHT - MARGIN
        Else
            dblItemlistHeight = 0
        End If
    Else
        dblItemlistTop = lngTlbrTop + MARGIN
        If (lngTlbrHeight - lngTlbrTop - (2 * MARGIN)) > 0 Then
            dblItemlistHeight = lngTlbrHeight - lngTlbrTop - (2 * MARGIN)
        Else
            dblItemlistHeight = 0
        End If
    End If
    
    If MainSettings.NavigationPane = True Then
        dblItemlistLeft = g_typInterface.IFindBGround.Left
    Else
        dblItemlistLeft = lngTlbrLeft + MARGIN
    End If
        
    If MainSettings.NavigationPane = True Then
        If (lngTlbrWidth - lngTlbrLeft - dblNavigationWidth - MARGIN) > 0 Then
            dblItemlistWidth = lngTlbrWidth - lngTlbrLeft - dblNavigationWidth - (2 * MARGIN)
        Else
            dblItemlistWidth = 0
        End If
    Else
        If (lngTlbrWidth - lngTlbrLeft) > 0 Then
            dblItemlistWidth = lngTlbrWidth - lngTlbrLeft - MARGIN
        Else
            dblItemlistWidth = 0
        End If
    End If
    
    '>> Set itemlist background's new dimension
    g_typInterface.IItemlist.Move dblItemlistLeft, dblItemlistTop, dblItemlistWidth, dblItemlistHeight
    
    '>> Modify Itemlist control's dimensions
    Call AdjustItemlistControls(Window, MainSettings, NodeProps)
    
    '>> Check reading pane's visibility
    g_typInterface.IReadingPane.Visible = NodeProps.ReadingPane
    
    '>> Check status bar's visibility
    Select Case g_typInterface.IStatusbar.Width
        Case Is > 6000
            If g_typInterface.IStatusbar.Panels.Count = 2 Then
                g_typInterface.IStatusbar.Panels(2).Width = 3000
                g_typInterface.IStatusbar.Panels(1).Width = g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(2).Width
            ElseIf g_typInterface.IStatusbar.Panels.Count = 3 Then
                g_typInterface.IStatusbar.Panels(3).Width = g_typInterface.IStatusbar.Width / 3
                g_typInterface.IStatusbar.Panels(1).Width = g_typInterface.IStatusbar.Width / 2 '(g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(2).Width) / 2
                g_typInterface.IStatusbar.Panels(2).Width = g_typInterface.IStatusbar.Width - (g_typInterface.IStatusbar.Panels(3).Width + g_typInterface.IStatusbar.Panels(1).Width)
            End If
'            g_typInterface.IStatusbar.Panels(2).Width = 3000
'            g_typInterface.IStatusbar.Panels(1).Width = g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(2).Width
        Case Is > 3000
            If g_typInterface.IStatusbar.Panels.Count = 2 Then
                g_typInterface.IStatusbar.Panels(2).Width = g_typInterface.IStatusbar.Width / 2
                g_typInterface.IStatusbar.Panels(1).Width = g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(2).Width
            ElseIf g_typInterface.IStatusbar.Panels.Count = 3 Then
                g_typInterface.IStatusbar.Panels(3).Width = g_typInterface.IStatusbar.Width / 3
                g_typInterface.IStatusbar.Panels(1).Width = g_typInterface.IStatusbar.Width / 2 '(g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(2).Width) / 2
                g_typInterface.IStatusbar.Panels(2).Width = g_typInterface.IStatusbar.Width - (g_typInterface.IStatusbar.Panels(3).Width + g_typInterface.IStatusbar.Panels(1).Width)
            End If
            
'            g_typInterface.IStatusbar.Panels(1).Width = 3000
'            g_typInterface.IStatusbar.Panels(2).Width = g_typInterface.IStatusbar.Width - g_typInterface.IStatusbar.Panels(1).Width
    End Select
    
    '>> Set the font size because it grows bigger upon resizing
    g_typInterface.IStatusbar.Font.Size = 8
    g_typInterface.IStatusbar.Visible = MainSettings.StatusBar
    
    If g_typInterface.INoView.Visible = True Then
        g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
    End If
    
End Sub

Private Sub AdjustItemlistControls(ByRef Window As Object, _
                                    ByRef MainSettings As CMainControls, _
                                    ByRef NodeProps As CGrid)

    Dim dblItemlistHeight As Double
    Dim dblItemlistWidth As Double
    
    Dim dblGridHeight As Double
    Dim dblGridWidth As Double
    
    Dim dblReadingPaneLeft As Double
    Dim dblReadingPaneTop As Double
    Dim dblReadingPaneHeight As Double
    Dim dblReadingPaneWidth As Double
    

    dblItemlistHeight = g_typInterface.IItemlist.Height
    dblItemlistWidth = g_typInterface.IItemlist.Width - MARGIN
    
    If NodeProps.ReadingPane = True Then
        '>> modify dimensions of grid and reading pane
        Select Case MainSettings.ReadingPanePos
            Case BotttomPane
                'reading pane is at the bottom
                dblGridWidth = dblItemlistWidth
                dblGridHeight = MainSettings.GridHeight
                
                dblReadingPaneWidth = dblItemlistWidth
                
                Select Case dblGridHeight
                
                    Case Is <= dblItemlistHeight
                        '>> grid height is less than background's height
                        dblReadingPaneHeight = dblItemlistHeight - dblGridHeight - MARGIN
                        If dblReadingPaneHeight < MINIMUM_WIDTH_HEIGHT Then
                            dblReadingPaneHeight = MINIMUM_WIDTH_HEIGHT
                            dblGridHeight = dblItemlistHeight - dblReadingPaneHeight
                        End If
                        '>> grid height is less than its minimum height
                        If dblGridHeight < MINIMUM_WIDTH_HEIGHT Then
                            dblGridHeight = MINIMUM_WIDTH_HEIGHT
                            dblReadingPaneHeight = dblItemlistHeight - dblGridHeight
                        End If
                        
                    Case Is < MINIMUM_WIDTH_HEIGHT
                        '>> grid height is less than its minimum height
                        
                        dblGridHeight = dblItemlistHeight
                        dblReadingPaneHeight = 0
                        
                    Case Is > dblItemlistHeight
                        '>> grid height is greater than the background's height
                        dblReadingPaneHeight = MINIMUM_WIDTH_HEIGHT
                        dblGridHeight = dblItemlistHeight - dblReadingPaneHeight
                        If dblGridHeight < MINIMUM_WIDTH_HEIGHT Then
                            dblGridHeight = MINIMUM_WIDTH_HEIGHT
                            dblReadingPaneHeight = dblItemlistHeight - dblGridHeight
                        End If
                        
                End Select
                                
                '>> get reading pane's dimensions
                dblReadingPaneTop = dblGridHeight + g_typInterface.IGridHeader.Height + (2 * MARGIN)
                dblReadingPaneHeight = IIf(dblReadingPaneHeight <= 0, 0, dblReadingPaneHeight)
                
                '>> set grid splitter's new dimension
                g_typInterface.IGridImage.Move 0, dblReadingPaneTop - MARGIN, dblGridWidth, MARGIN
                
            Case RightPane
                '>> reading pane is at the right
                dblGridHeight = dblItemlistHeight
                dblReadingPaneHeight = dblItemlistHeight
                
                dblGridWidth = MainSettings.GridWidth
                
                Select Case MainSettings.GridWidth
                
                    Case Is <= dblItemlistWidth
                        '>> grid width is less than the background's width
                        dblReadingPaneWidth = dblItemlistWidth - dblGridWidth - MARGIN
                        If dblReadingPaneWidth < MINIMUM_WIDTH_HEIGHT Then
                            dblReadingPaneWidth = MINIMUM_WIDTH_HEIGHT
                            dblGridWidth = dblItemlistWidth - dblReadingPaneWidth - MARGIN
                        End If
                        If dblGridWidth < MINIMUM_WIDTH_HEIGHT Then
                            dblGridWidth = MINIMUM_WIDTH_HEIGHT
                            dblReadingPaneWidth = dblItemlistWidth - dblGridWidth - MARGIN
                        End If
                        
                    Case Is < MINIMUM_WIDTH_HEIGHT
                        '>> grid width is less than its minimum width
                        dblGridWidth = dblItemlistWidth
                        dblReadingPaneWidth = 0
                        
                    Case Is > dblItemlistWidth
                        '>> grid width is greater than the background's width
                        dblReadingPaneWidth = MINIMUM_WIDTH_HEIGHT
                        dblGridWidth = dblItemlistWidth - dblReadingPaneWidth - MARGIN
                        If dblGridWidth < MINIMUM_WIDTH_HEIGHT Then
                            dblGridWidth = MINIMUM_WIDTH_HEIGHT
                            dblReadingPaneWidth = dblItemlistWidth - dblGridWidth - MARGIN
                        End If
                        
                End Select
                            
                '>> get reading pane's dimensions
                If dblReadingPaneWidth > 0 Then
                    dblReadingPaneLeft = dblGridWidth + MARGIN
                    dblReadingPaneWidth = dblReadingPaneWidth
                Else
                    dblReadingPaneLeft = dblGridWidth + MARGIN
                    dblReadingPaneWidth = 0
                End If
                
                '>> set grid splitter's new dimension
                g_typInterface.IGridImage.Move dblGridWidth, 0, MARGIN, dblGridHeight
            Case Else
                '>> Reading Pane is not visible
                If MainSettings.UseCustomFind And _
                    MainSettings.CustomFind = True Then
                    dblGridHeight = dblItemlistHeight - g_typInterface.IGridHeader.Height - g_typInterface.ICustomFind.Height
                Else
                    dblGridHeight = dblItemlistHeight - g_typInterface.IGridHeader.Height
                End If
                dblGridWidth = dblItemlistWidth
                If dblGridHeight < 0 Then
                    dblGridHeight = 0
                End If
        End Select
    Else
        '>> Reading Pane is not visible
        If MainSettings.UseCustomFind And _
            MainSettings.CustomFind = True Then
            dblGridHeight = dblItemlistHeight - g_typInterface.IGridHeader.Height - g_typInterface.ICustomFind.Height - (2 * MARGIN)
        Else
            dblGridHeight = dblItemlistHeight - g_typInterface.IGridHeader.Height - MARGIN
        End If
        dblGridWidth = dblItemlistWidth - MARGIN
        If dblGridWidth < 0 Then
            dblGridWidth = 0
        End If
        If dblGridHeight < 0 Then
            dblGridHeight = 0
        End If
    End If
    
    g_typInterface.IGridHeader.Move 0, 0, dblGridWidth, g_typInterface.IGridHeader.Height
'    Window.imgHeader.Move g_typInterface.IGridHeader.Width - Window.imgHeader.Width - 250, MARGIN
    
    If MainSettings.UseCustomFind And _
        MainSettings.CustomFind = True Then
        g_typInterface.ICustomFind.Move MARGIN, g_typInterface.IGridHeader.Height + MARGIN, dblGridWidth - IIf(dblGridWidth > (3 * MARGIN), (3 * MARGIN), 0), FIND_BOX_HEIGHT
    End If
    
    '>> apply grid's new dimension
    If MainSettings.UseCustomFind And _
        MainSettings.CustomFind = True Then
        'g_typInterface.IGrid.Move 0, g_typInterface.IGridHeader.Height + g_typInterface.ICustomFind.Height + (2 * MARGIN), dblGridWidth - IIf(dblGridWidth > 50, MARGIN, 0), dblGridHeight - IIf(dblGridHeight > 50, MARGIN, 0)
        g_typInterface.IGrid.Move 0, g_typInterface.IGridHeader.Height + g_typInterface.ICustomFind.Height + (2 * MARGIN), dblGridWidth, dblGridHeight - IIf(dblGridHeight > 50, MARGIN, 0)
    Else
        'g_typInterface.IGrid.Move 0, g_typInterface.IGridHeader.Height, dblGridWidth - IIf(dblGridWidth > 50, MARGIN, 0), dblGridHeight - IIf(dblGridHeight > 50, MARGIN, 0)
        g_typInterface.IGrid.Move 0, g_typInterface.IGridHeader.Height, dblGridWidth, dblGridHeight - IIf(dblGridHeight > 50, MARGIN, 0)
    End If
    
    '>> apply reading pane's new dimension
    g_typInterface.IReadingPane.Move dblReadingPaneLeft, dblReadingPaneTop, dblReadingPaneWidth, dblReadingPaneHeight
    
    DoEvents
End Sub

Public Sub ResizeGrid(ByRef Window As Object, _
                    ByRef MainSettings As CMainControls, _
                    ByRef ADOConnection As ADODB.Connection, _
                    ByVal XPosition As Long, ByVal YPosition As Long, _
                    ByVal MouseState As MouseStateConstants)
        
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    
    Select Case MainSettings.ReadingPanePos
        
        Case BotttomPane
            Select Case MouseState
                
                Case MouseDown
                    '>> Show splitter to drag
                    dblLeft = g_typInterface.IGridImage.Left
                    dblTop = g_typInterface.IGridImage.Top
                    dblWidth = g_typInterface.IGridImage.Width
                    dblHeight = g_typInterface.IGridImage.Height
                    
                    g_typInterface.IGridSplitter.Move dblLeft, dblTop, dblWidth, dblHeight
                    g_typInterface.IGridSplitter.Visible = True
                    
                Case MouseMove
                    Screen.MousePointer = vbSizeNS
                    If g_typInterface.IGridSplitter.Visible = True Then
                        '>> move splitter
                        dblLeft = g_typInterface.IGridImage.Left
                        dblTop = g_typInterface.IGridImage.Top + YPosition
                        dblWidth = g_typInterface.IGridImage.Width
                        dblHeight = g_typInterface.IGridImage.Height
                        
                        '>> stop splitter from moving if it reached minimum width/height of the grid/reading pane
                        If dblTop < MINIMUM_WIDTH_HEIGHT Then
                            dblTop = MINIMUM_WIDTH_HEIGHT
                        ElseIf dblTop > (g_typInterface.IItemlist.Height - MINIMUM_WIDTH_HEIGHT) Then
                            dblTop = g_typInterface.IItemlist.Height - MINIMUM_WIDTH_HEIGHT
                        End If
                        
                        g_typInterface.IGridSplitter.Move dblLeft, dblTop, dblWidth, dblHeight
                    End If
                    
                Case MouseUp
                    '>> adjust grid and reading pane dimensions based on the the new position of splitter
                    dblLeft = g_typInterface.IGridImage.Left + XPosition
                    dblTop = g_typInterface.IGridImage.Top + YPosition
                    
                    Call MoveGrid(Window, MainSettings, ADOConnection, dblLeft, dblTop)
                    g_typInterface.IGridSplitter.Visible = False
                    MainSettings.SaveSetting ADOConnection
                    Screen.MousePointer = vbDefault
                    
            End Select
        
        Case RightPane
            Select Case MouseState
                
                Case MouseDown
                    '>> Show splitter to drag
                    dblLeft = g_typInterface.IGridImage.Left
                    dblTop = g_typInterface.IGridImage.Top
                    dblWidth = g_typInterface.IGridImage.Width
                    dblHeight = g_typInterface.IGridImage.Height
                    
                    g_typInterface.IGridSplitter.Move dblLeft, dblTop, dblWidth, dblHeight
                    g_typInterface.IGridSplitter.Visible = True
                    
                Case MouseMove
                    Screen.MousePointer = vbSizeWE
        
                    If g_typInterface.IGridSplitter.Visible = True Then
                        '>> move splitter
                        dblLeft = g_typInterface.IGridImage.Left + XPosition
                        dblTop = g_typInterface.IGridImage.Top
                        dblWidth = g_typInterface.IGridImage.Width
                        dblHeight = g_typInterface.IGridImage.Height
                        
                        '>> stop splitter from moving if it reached minimum width/height of the grid/reading pane
                        If dblLeft < MINIMUM_WIDTH_HEIGHT Then
                            dblLeft = MINIMUM_WIDTH_HEIGHT
                        ElseIf dblLeft > (g_typInterface.IItemlist.Width - MINIMUM_WIDTH_HEIGHT) Then
                            dblLeft = g_typInterface.IItemlist.Width - MINIMUM_WIDTH_HEIGHT
                        End If
                        
                        g_typInterface.IGridSplitter.Move dblLeft, dblTop, dblWidth, dblHeight
                    End If
                    
                Case MouseUp
                    '>> adjust grid and reading pane dimensions based on the the new position of splitter
                    dblLeft = g_typInterface.IGridImage.Left + XPosition
                    dblTop = g_typInterface.IGridImage.Top + YPosition
                    
                    Call MoveGrid(Window, MainSettings, ADOConnection, dblLeft, dblTop)
                    g_typInterface.IGridSplitter.Visible = False
                    Screen.MousePointer = vbDefault
            End Select
        
        Case PaneNotVisible
    
    End Select
    
End Sub

Public Sub MoveGrid(ByRef Window As Object, _
                    ByRef MainSettings As CMainControls, _
                    ADOConnection As ADODB.Connection, _
                    ByVal SplitterLeftPos As Long, _
                    ByVal SplitterTopPos As Long)

    
    Dim dblSplitterLeft As Double
    Dim dblSplitterTop As Double
    Dim dblSplitterWidth As Double
    Dim dblSplitterHeight As Double
    
    Dim dblGridLeft As Double
    Dim dblGridTop As Double
    Dim dblGridWidth As Double
    Dim dblGridHeight As Double

    Dim dblReadPaneLeft As Double
    Dim dblReadPaneTop As Double
    Dim dblReadPaneWidth As Double
    Dim dblReadPaneHeight As Double
    
    Select Case MainSettings.ReadingPanePos
    
        Case BotttomPane
            '>> get splitter's new position
            If SplitterTopPos < MINIMUM_WIDTH_HEIGHT Then
                dblSplitterTop = MINIMUM_WIDTH_HEIGHT
            Else
                dblSplitterTop = SplitterTopPos
            End If
            dblSplitterLeft = g_typInterface.IGridSplitter.Left
            dblSplitterWidth = g_typInterface.IGridSplitter.Width
            dblSplitterHeight = g_typInterface.IGridSplitter.Height
            
            '>> get grid dimensions
            dblGridLeft = 0
            dblGridTop = 0
            dblGridWidth = g_typInterface.IItemlist.Width
            dblGridHeight = dblSplitterTop
            
            '>> get reading pane's height
            dblReadPaneHeight = g_typInterface.IItemlist.Height - dblGridHeight - MARGIN
            
            '>> check if reading pane reaches its minimum width
            If dblReadPaneHeight < MINIMUM_WIDTH_HEIGHT Then
                dblReadPaneHeight = MINIMUM_WIDTH_HEIGHT
                dblGridHeight = g_typInterface.IItemlist.Height - MINIMUM_WIDTH_HEIGHT
                dblSplitterTop = dblGridHeight
            End If
            
            '>> check if grid reaches its minimum width
            If dblGridHeight < MINIMUM_WIDTH_HEIGHT Then
                dblGridHeight = MINIMUM_WIDTH_HEIGHT
                dblReadPaneHeight = g_typInterface.IItemlist.Height - dblGridHeight
                dblSplitterTop = dblGridHeight - MARGIN
            End If
            
            '>> get reading pane dimensions
            dblReadPaneLeft = 0
            dblReadPaneTop = dblSplitterTop + MARGIN
            dblReadPaneWidth = dblGridWidth
            
            
        Case RightPane
            '>> get splitter's new position
            If SplitterLeftPos < MINIMUM_WIDTH_HEIGHT Then
                dblSplitterLeft = MINIMUM_WIDTH_HEIGHT
            Else
                dblSplitterLeft = SplitterLeftPos
            End If
            dblSplitterTop = g_typInterface.IGridSplitter.Top
            dblSplitterWidth = g_typInterface.IGridSplitter.Width
            dblSplitterHeight = g_typInterface.IGridSplitter.Height
            
            
            '>> set grid's width based on the splitter's new position
            dblGridLeft = 0
            dblGridTop = 0
            dblGridWidth = dblSplitterLeft
            dblGridHeight = g_typInterface.IGrid.Height
            
                  
            If (g_typInterface.IItemlist.Width - dblGridWidth) < MINIMUM_WIDTH_HEIGHT Then
                '>> set reading pane to its minimum width if the actual width is less than the minimum width
                dblReadPaneWidth = MINIMUM_WIDTH_HEIGHT
                If (g_typInterface.IItemlist.Width - MINIMUM_WIDTH_HEIGHT) > 0 Then
                    dblSplitterLeft = g_typInterface.IItemlist.Width - MINIMUM_WIDTH_HEIGHT
                Else
                    dblSplitterLeft = 0
                End If
                dblGridWidth = dblSplitterLeft
            Else
                '>> set reading pane's new width
                dblReadPaneWidth = g_typInterface.IItemlist.Width - dblGridWidth - (2 * MARGIN)
            End If
            
            dblReadPaneTop = 0
            dblReadPaneLeft = dblSplitterLeft + dblSplitterWidth
            dblReadPaneHeight = dblGridHeight
            
    
        Case Else
    
    End Select
    
    '>> move controls to its new position and apply new widths/heights
    g_typInterface.IGridImage.Move dblSplitterLeft, dblSplitterTop, dblSplitterWidth, dblSplitterHeight
    g_typInterface.IGrid.Move dblGridLeft, dblGridTop, dblGridWidth, dblGridHeight
    g_typInterface.IReadingPane.Move dblReadPaneLeft, dblReadPaneTop, dblReadPaneWidth, dblReadPaneHeight
    
    '>> save grid's new height and width
    MainSettings.GridHeight = dblGridHeight
    MainSettings.GridWidth = dblGridWidth
    
    
End Sub


Public Sub ResizeNavigationPane(ByRef Window As Object, _
                    ByRef MainSettings As CMainControls, _
                    ByRef NodeProps As CGrid, _
                    ByRef ADOConnection As ADODB.Connection, _
                    ByVal XPosition As Long, _
                    ByVal MouseState As MouseStateConstants)

    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    
    Dim lngTlbrLeft As Long
    Dim lngTlbrTop As Long
    Dim lngTlbrWidth As Long
    Dim lngTlbrHeight As Long

    '>> Get toolbar dimension
    g_typInterface.ICommandBar.GetClientRect lngTlbrLeft, lngTlbrTop, lngTlbrWidth, lngTlbrHeight
    
    Select Case MouseState
        
        Case MouseDown
            '>> Show splitter
            dblLeft = g_typInterface.ISplitterImage.Left
            dblTop = g_typInterface.ISplitterImage.Top
            dblWidth = g_typInterface.ISplitterImage.Width
            dblHeight = g_typInterface.ISplitterImage.Height
            
            g_typInterface.ISplitter.Move dblLeft, dblTop, dblWidth, dblHeight
            g_typInterface.ISplitter.Visible = True
            
        Case MouseMove
            
            Screen.MousePointer = vbSizeWE
            If g_typInterface.ISplitter.Visible = True Then
                
                '>> move splitter
                dblLeft = g_typInterface.ISplitterImage.Left + XPosition
                dblTop = g_typInterface.ISplitterImage.Top
                dblWidth = MARGIN
                dblHeight = g_typInterface.ISplitterImage.Height
                
                '>> stop splitter from moving if it reached minimum width
                If dblLeft < (MINIMUM_WIDTH_HEIGHT + lngTlbrLeft) Then
                    dblLeft = MINIMUM_WIDTH_HEIGHT + lngTlbrLeft
                ElseIf dblLeft > (Window.Width / 2) Then
                    dblLeft = (Window.Width / 2)
                End If
                g_typInterface.ISplitter.ZOrder 0
                g_typInterface.ISplitter.Move dblLeft, dblTop, dblWidth, dblHeight
            End If
            
        Case MouseUp
            '>> adjust navigation pane dimensions based on the the new position of splitter
            dblLeft = g_typInterface.ISplitterImage.Left + XPosition
            dblTop = g_typInterface.IButtonBar.Top
            dblWidth = MARGIN
            dblHeight = g_typInterface.IButtonBar.Height
            
            g_typInterface.ISplitterImage.Move dblLeft, dblTop, dblWidth, dblHeight
            
            dblLeft = 0
            dblTop = g_typInterface.IButtonBar.Top
            dblWidth = g_typInterface.ISplitterImage.Left - lngTlbrLeft - MARGIN
            dblHeight = g_typInterface.IButtonBar.Height
            
            If dblWidth > (Window.Width - MARGIN - lngTlbrLeft) Then
                dblWidth = Window.Width
            ElseIf dblWidth < MINIMUM_WIDTH_HEIGHT Then
                dblWidth = MINIMUM_WIDTH_HEIGHT
            ElseIf dblWidth > (Window.Width / 2) Then
                dblWidth = (Window.Width / 2)
            End If
                        
            g_typInterface.IButtonBar.Move dblLeft, dblTop, dblWidth, dblHeight
            
            MainSettings.NavigationPaneWidth = dblWidth
                        
            Call ResizeControls(Window, MainSettings, NodeProps, ADOConnection)
            
            g_typInterface.ISplitter.Visible = False
            Screen.MousePointer = vbDefault
            
    End Select
            
         
End Sub

Public Sub ToggleControls(ByRef Window As Object, _
                        ByRef Button As MainControlConstants, _
                        ByRef MainSettings As CMainControls, _
                        ByRef NodeProps As CGrid, _
                        ByRef ADOConnection As ADODB.Connection)
                            
    Select Case Button
        Case Navigation
            MainSettings.NavigationPane = Not g_typInterface.IButtonBar.Visible
            
        Case RightReadingPane
            MainSettings.ReadingPanePos = RightPane
            NodeProps.ReadingPane = True
            
        Case BottomReadingPane
            MainSettings.ReadingPanePos = BotttomPane
            NodeProps.ReadingPane = True
            
        Case ReadingPaneOff
            MainSettings.ReadingPanePos = PaneNotVisible
            NodeProps.ReadingPane = False
            
        Case Find
            MainSettings.FindBox = Not g_typInterface.IFindBGround.Visible
                        
        Case Status_Bar
            MainSettings.StatusBar = Not g_typInterface.IStatusbar.Visible
            g_typInterface.IStatusbar.Visible = MainSettings.StatusBar
            g_typInterface.ICommandBar.RecalcLayout
            
    End Select
    
    Call ResizeControls(Window, MainSettings, NodeProps, ADOConnection)
    
End Sub

'>> FOR CP USE - SAME AS TOOLCLICK
Public Function ToolButtonClick(ByRef Window As Object, _
                    ByVal Application As Object, _
                    ByRef Control As XtremeCommandBars.ICommandBarControl, _
                    ByRef MenuConstant As MenuCommandConstant, _
                    ByRef MainSettings As CMainControls, _
                    ByRef NavPaneProps As CNavigationPane, _
                    ByRef NavigationPane As INavigationPane, _
                    ByRef GridProps As CGrid, _
                    ByRef GridSettings As IGrid, _
                    ByRef MachinePaths As CMachinePaths, _
                    ByRef Licensee As CLicensee, _
                    ByRef AboutInfo As CAbout, _
                    ByRef ADOConnection As ADODB.Connection) As Boolean
    
    
    Dim ctlToolbar As CommandBarControl
    Dim ctrlButtons As CommandBarControls
    Dim lngCounter As Long

    ToolButtonClick = True
    
    Select Case MenuConstant
        Case [File Exit]
            ToolButtonClick = False
            Exit Function
            
        Case [Tools Find]
            MainSettings.FindBox = Not g_typInterface.IFindBGround.Visible
            Control.Checked = MainSettings.FindBox
            
        Case [View Navigation Pane]
            MainSettings.NavigationPane = Not g_typInterface.IButtonBar.Visible
            Control.Checked = MainSettings.NavigationPane
            
        Case [View Reading Pane Right]
            MainSettings.ReadingPanePos = RightPane
            GridProps.ReadingPane = True
            Control.Checked = True
            
        Case [View Reading Pane Bottom]
            MainSettings.ReadingPanePos = BotttomPane
            GridProps.ReadingPane = True
            Control.Checked = True
            
        Case [View Reading Pane Off]
            MainSettings.ReadingPanePos = PaneNotVisible
            GridProps.ReadingPane = False
            Control.Checked = True
            
        Case [View Group By Box]
            GridProps.GroupHeaderIsVisible = Not GridProps.GroupHeaderIsVisible
            g_typInterface.IGrid.GroupByBoxVisible = GridProps.GroupHeaderIsVisible
            Control.Checked = GridProps.GroupHeaderIsVisible
            
        Case [View Card View]
            GridProps.CardView = True
            g_typInterface.IGrid.View = GridProps.CardView
            Control.Checked = True
            Set ctlToolbar = g_typInterface.ICommandBar.FindControl(, ID_VIEW_ITEMLIST_LISTVIEW, , True)
            If Not ctlToolbar Is Nothing Then
                ctlToolbar.Checked = False
            End If
            Set ctlToolbar = Nothing
            
        Case [View List View]
            GridProps.CardView = False
            g_typInterface.IGrid.View = GridProps.CardView
            Control.Checked = True
            Set ctlToolbar = g_typInterface.ICommandBar.FindControl(, ID_VIEW_ITEMLIST_CARDVIEW, , True)
            If Not ctlToolbar Is Nothing Then
                ctlToolbar.Checked = False
            End If
            Set ctlToolbar = Nothing
            
        Case [View Odd Even Color]
            GridProps.OddEvenColor = Not GridProps.OddEvenColor
            g_typInterface.IGrid.UseEvenOddColor = GridProps.OddEvenColor
            Control.Checked = GridProps.OddEvenColor
            
        Case [View Gridlines]
            GridProps.GridLines = Not GridProps.GridLines
            g_typInterface.IGrid.GridLines = GridProps.GridLines
            Control.Checked = GridProps.GridLines
            
        Case [View Collapse All]
            g_typInterface.IGrid.CollapseAll
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
            
        Case [View Expand All]
            g_typInterface.IGrid.ExpandAll
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case [View Collapse Group]
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = False
                End If
            End If
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case [View Expand Group]
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True
                End If
            End If
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case [View Status Bar]
            MainSettings.StatusBar = Not g_typInterface.IStatusbar.Visible
            g_typInterface.IStatusbar.Visible = MainSettings.StatusBar
            g_typInterface.ICommandBar.RecalcLayout
            Control.Checked = MainSettings.StatusBar
            
        Case [View Options]
            MainSettings.ViewOptions = Not frmPane.fraViews.Visible
            frmPane.fraViews.Visible = MainSettings.ViewOptions
            frmPane.sccViews.Visible = MainSettings.ViewOptions
            Call frmPane.PaneResize
            Control.Checked = MainSettings.ViewOptions
    
        Case [View Customize Current View]
            If NavigationPane.CustomizeView = True Then
                NavigationPane.TriggerNodeClickEvent frmPane.tvwMain.SelectedItem, True, True
            End If

        Case [Showfields]
            frmShowFields.ShowForm Window, GridProps, ADOConnection
            
        Case [Sort]
            frmSort.ShowForm Window, GridProps, ADOConnection
            
        Case [Group By]
            frmGroupBy.ShowForm Window, GridProps, ADOConnection
            
        Case [Filter]
            frmFilter.ShowForm Window, GridProps, ADOConnection
            
        Case [Other Settings]
        
        Case [Autoformat]
            frmAutoFormat.ShowForm Window, GridProps, ADOConnection
            
        Case [Format Columns]
            frmFormatColumns.ShowForm Window, GridProps, ADOConnection

        Case [Tools Options]
            If frmOptions.ShowOptions(Window, Application, MachinePaths, MainSettings) = "End" Then
                ToolButtonClick = False
                Exit Function
            End If
                    
        Case [Help Licensee]
            Call Licensee.ShowForm(Window, ADOConnection)
            
        Case [Help About]
            AboutInfo.Show Window, vbModal
        
        Case [Context Rename Folder]
            Call NavigationPane.ProcessPopupEvents(Folder_Modify)
            
        Case [Context Remove Folder]
            Call NavigationPane.ProcessPopupEvents(Folder_Delete)
            
'        Case [Product Activation]
'            Call frmActivate.ShowForm(Window, ADOConnection, Licensee, Application)
'            g_typInterface.ILicense.ShowTriggerDlg Window.hwnd, 0, "", 0, 0
        
        Case [Go Folder]
            frmGoToFolder.GoToFolder Window, Application, ADOConnection, MainSettings, NavigationPane, NavPaneProps
            
        Case Is < 0
            
            If Control.Checked = False Then
                
                Set ctrlButtons = Control.Parent.Controls
                For lngCounter = 1 To ctrlButtons.Count
                    ctrlButtons(lngCounter).Checked = False
                Next
                Set ctrlButtons = Nothing
                
                GridSettings.GroupExpandCollapse GridProps, Window
                GridSettings.GridGroupingChange GridProps, Window, ADOConnection
                GridSettings.GetLastSelectedRows GridProps, Window
                GridSettings.GridColumnSort GridProps, Window, g_typInterface.IGrid.Columns(Val(Right(Control.ID, 2))), ADOConnection
                GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
                Control.Checked = True
            End If
            
    End Select
    
    Call ResizeControls(Window, MainSettings, GridProps, ADOConnection)
    
End Function

Public Function ToolClick(ByRef Window As Object, _
                    ByVal Application As Object, _
                    ByRef Control As XtremeCommandBars.ICommandBarControl, _
                    ByRef MainSettings As CMainControls, _
                    ByRef NavPaneProps As CNavigationPane, _
                    ByRef NavigationPane As INavigationPane, _
                    ByRef GridProps As CGrid, _
                    ByRef GridSettings As IGrid, _
                    ByRef MachinePaths As CMachinePaths, _
                    ByRef Licensee As CLicensee, _
                    ByRef AboutInfo As CAbout, _
                    ByRef ADOConnection As ADODB.Connection) As Boolean
    
    Dim ctlToolbar As CommandBarControl
    Dim ctrlButtons As CommandBarControls
    Dim lngCounter As Long

    ToolClick = True
    
    Select Case Control.ID
        
        Case ID_FILE_EXIT
            ToolClick = False
            Exit Function
            
        Case ID_TOOLS_FIND
            MainSettings.FindBox = Not g_typInterface.IFindBGround.Visible
            Control.Checked = MainSettings.FindBox
            
        Case ID_VIEW_NAVIGATIONPANE
            MainSettings.NavigationPane = Not g_typInterface.IButtonBar.Visible
            Control.Checked = MainSettings.NavigationPane
            
        Case ID_VIEW_READINGPANE_RIGHT
            MainSettings.ReadingPanePos = RightPane
            GridProps.ReadingPane = True
            Control.Checked = True
            
        Case ID_VIEW_READINGPANE_BOTTOM
            MainSettings.ReadingPanePos = BotttomPane
            GridProps.ReadingPane = True
            Control.Checked = True
            
        Case ID_VIEW_READINGPANE_OFF
            MainSettings.ReadingPanePos = PaneNotVisible
            GridProps.ReadingPane = False
            Control.Checked = True
            
        Case ID_VIEW_GROUPBYBOX
            GridProps.GroupHeaderIsVisible = Not GridProps.GroupHeaderIsVisible
            g_typInterface.IGrid.GroupByBoxVisible = GridProps.GroupHeaderIsVisible
            Control.Checked = GridProps.GroupHeaderIsVisible
            
        Case ID_VIEW_ITEMLIST_CARDVIEW
            GridProps.CardView = True
            g_typInterface.IGrid.View = GridProps.CardView
            Control.Checked = True
            Set ctlToolbar = g_typInterface.ICommandBar.FindControl(, ID_VIEW_ITEMLIST_LISTVIEW, , True)
            If Not ctlToolbar Is Nothing Then
                ctlToolbar.Checked = False
            End If
            Set ctlToolbar = Nothing
            
        Case ID_VIEW_ITEMLIST_LISTVIEW
            GridProps.CardView = False
            g_typInterface.IGrid.View = GridProps.CardView
            Control.Checked = True
            Set ctlToolbar = g_typInterface.ICommandBar.FindControl(, ID_VIEW_ITEMLIST_CARDVIEW, , True)
            If Not ctlToolbar Is Nothing Then
                ctlToolbar.Checked = False
            End If
            Set ctlToolbar = Nothing
            
        Case ID_VIEW_ITEMLIST_ODDEVENCOLOR
            GridProps.OddEvenColor = Not GridProps.OddEvenColor
            g_typInterface.IGrid.UseEvenOddColor = GridProps.OddEvenColor
            Control.Checked = GridProps.OddEvenColor
            
        Case ID_VIEW_ITEMLIST_GRIDLINES
            GridProps.GridLines = Not GridProps.GridLines
            g_typInterface.IGrid.GridLines = GridProps.GridLines
            Control.Checked = GridProps.GridLines
            
        Case ID_VIEW_COLLAPSEALL
            g_typInterface.IGrid.CollapseAll
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
            
        Case ID_VIEW_EXPANDALL
            g_typInterface.IGrid.ExpandAll
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case ID_VIEW_COLLAPSEGROUP
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = False
                End If
            End If
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case ID_VIEW_EXPANDGROUP
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True
                End If
            End If
            GridSettings.GroupExpandCollapse GridProps, Window
            GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
        
        Case ID_VIEW_STATUSBAR
            MainSettings.StatusBar = Not g_typInterface.IStatusbar.Visible
            g_typInterface.IStatusbar.Visible = MainSettings.StatusBar
            g_typInterface.ICommandBar.RecalcLayout
            Control.Checked = MainSettings.StatusBar
            
        Case ID_VIEW_VIEWOPTIONS
            MainSettings.ViewOptions = Not frmPane.fraViews.Visible
            frmPane.fraViews.Visible = MainSettings.ViewOptions
            frmPane.sccViews.Visible = MainSettings.ViewOptions
            Call frmPane.PaneResize
            Control.Checked = MainSettings.ViewOptions
    
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW
            If NavigationPane.CustomizeView = True Then
                NavigationPane.TriggerNodeClickEvent frmPane.tvwMain.SelectedItem, True, True
            End If
        
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_SHOWFIELDS
            frmShowFields.ShowForm Window, GridProps, ADOConnection
            
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_SORT
            frmSort.ShowForm Window, GridProps, ADOConnection
            
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_GROUPBY
            frmGroupBy.ShowForm Window, GridProps, ADOConnection
            
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_FILTER
            frmFilter.ShowForm Window, GridProps, ADOConnection
            
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_OTHERSETTINGS
        
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_AUTOFORMAT
            frmAutoFormat.ShowForm Window, GridProps, ADOConnection
            
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_FORMATCOLUMNS
            frmFormatColumns.ShowForm Window, GridProps, ADOConnection
        
        Case ID_VIEW_REFRESH
            GridProps.DataChanged = True
            
        Case ID_REPORTS
        
        Case ID_TOOLS
        
        Case ID_TOOLS_OPTIONS
            If frmOptions.ShowOptions(Window, Application, MachinePaths, MainSettings) = "End" Then
                ToolClick = False
                Exit Function
            End If
            
        Case ID_HELP
        
        Case ID_HELP_LICENSEE
            Call Licensee.ShowForm(Window, ADOConnection)
            
        Case ID_HELP_ABOUT
            AboutInfo.Show Window, vbModal
            
'        Case ID_HELP_PRODACTIVATION
'            Call frmActivate.ShowForm(Window, ADOConnection, Licensee, Application)
'            g_typInterface.ILicense.ShowTriggerDlg Window.hwnd, 0, "", 0, 0
        
        Case ID_CONTEXT_RENAMEFAVEFOLDER
            Call NavigationPane.ProcessPopupEvents(Folder_Modify)
            
        Case ID_CONTEXT_REMOVEFAVEFOLDER
            Call NavigationPane.ProcessPopupEvents(Folder_Delete)
            
        Case ID_GO_FOLDER
            frmGoToFolder.GoToFolder Window, Application, ADOConnection, MainSettings, NavigationPane, NavPaneProps
            
        Case Is < 0
            
            
            If Control.Checked = False Then
                
                Set ctrlButtons = Control.Parent.Controls
                For lngCounter = 1 To ctrlButtons.Count
                    ctrlButtons(lngCounter).Checked = False
                Next
                Set ctrlButtons = Nothing
                
                GridSettings.GroupExpandCollapse GridProps, Window
                GridSettings.GridGroupingChange GridProps, Window, ADOConnection
                GridSettings.GetLastSelectedRows GridProps, Window
                GridSettings.GridColumnSort GridProps, Window, g_typInterface.IGrid.Columns(Val(Right(Control.ID, 2))), ADOConnection
                GridSettings.SaveGridSetting NavPaneProps, GridProps, ADOConnection
                Control.Checked = True
            End If
            
    End Select
    
    Call ResizeControls(Window, MainSettings, GridProps, ADOConnection)
    
    
End Function

Friend Sub LoadArrangeByButtons(ByRef Window As Object, _
                                ByRef GridProps As CGrid)

    
    Dim mnuButton As CommandBarPopup
    Dim ctrlButtons As CommandBarControls
    Dim cbrArrangeBy As CommandBarButton
    
    Dim lngCounter As Long
    
    DoEvents
    
    Set mnuButton = g_typInterface.ICommandBar.FindControl(, ID_VIEW_ARRANGEBY, , True)
        
    If Not mnuButton Is Nothing Then
        
        Set ctrlButtons = mnuButton.CommandBar.Controls
        
        '>> Delete all "Arrange By" buttons
        ctrlButtons.DeleteAll
        
        If Not g_typInterface.IGrid.ADORecordset Is Nothing Then
            '>> Create Arrange By Fields
            For lngCounter = 1 To g_typInterface.IGrid.Columns.Count
                If g_typInterface.IGrid.Columns(lngCounter).Visible = True Then
                    ctrlButtons.Add xtpControlButton, -CLng(311 & Format(g_typInterface.IGrid.Columns(lngCounter).Index, "00")), g_typInterface.IGrid.Columns(lngCounter).Caption, -1, False
                End If
            Next
        End If
            
        '>> Add "Show Views In Navigation Pane" button
        If Not ctrlButtons Is Nothing Then
            Set cbrArrangeBy = ctrlButtons.Add(xtpControlButton, ID_VIEW_ARRANGEBY_VIEWOPTIONS, "S&how Views In Navigation Pane", -1, False)
            cbrArrangeBy.BeginGroup = True
            Set cbrArrangeBy = Nothing
        End If
        
        '>> Add "Customize Current View..." button
        If Not ctrlButtons Is Nothing Then
            Set mnuButton = ctrlButtons.Add(xtpControlPopup, ID_VIEW_ARRANGEBY_CURRENTVIEW, "Current &View", -1, False)
            mnuButton.ID = ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW
            mnuButton.CommandBar.Controls.Add xtpControlButton, ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW, "&Customize Current View...", , True
        End If
        
        Set ctrlButtons = Nothing
    
    End If
    
    Set mnuButton = Nothing
    
End Sub

Private Sub CreateFindBar(ByRef Window As Object)
    
    Dim cbrLookFor As CommandBarEdit
    Dim cbrSearchIn As CommandBarComboBox
    
    Dim cbrSearchLabel As CommandBarControl
    Dim cbrTemp As CommandBarControl
    
    Dim cbrFindNow As CommandBarButton
    Dim cbrClear As CommandBarButton
    Dim cbrAdvancedFind As CommandBarButton
    Dim cbrX As CommandBarButton
    
    Dim cbrOptions As CommandBarPopup

    g_typInterface.IFind.EnableCustomization False
    g_typInterface.IFind.ActiveMenuBar.EnableDocking xtpFlagStretched
    g_typInterface.IFind.ActiveMenuBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    g_typInterface.IFind.ActiveMenuBar.Controls.DeleteAll
    
    Set cbrLookFor = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlEdit, ID_FIND_LOOK_FOR, "Look For: ", , True)
    cbrLookFor.ID = ID_FIND_LOOK_FOR
    Set cbrLookFor = Nothing
    
    Set cbrSearchLabel = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlLabel, ID_FIND_SEARCH_IN, "Search In: ", , True)
    cbrSearchLabel.ID = ID_FIND_SEARCH_IN
    Set cbrSearchIn = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlComboBox, ID_FIND_SEARCH_BOX, "", , True)
    cbrSearchIn.ID = ID_FIND_SEARCH_BOX
    Set cbrSearchLabel = Nothing
    Set cbrSearchIn = Nothing
    
    Set cbrFindNow = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlButton, ID_FIND_FIND_NOW, "Find Now", , True)
    cbrFindNow.ID = ID_FIND_FIND_NOW
    Set cbrClear = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlButton, ID_FIND_CLEAR, "Clear", , True)
    cbrClear.ID = ID_FIND_CLEAR
    cbrClear.Enabled = False
    Set cbrFindNow = Nothing
    Set cbrClear = Nothing
        
    
    Set cbrOptions = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlSplitButtonPopup, ID_FIND_OPTION, "Options", , True)
    cbrOptions.ID = ID_FIND_OPTION
    cbrOptions.Style = xtpButtonCaption
    cbrOptions.Flags = xtpFlagRightAlign
    
    Set cbrAdvancedFind = cbrOptions.CommandBar.Controls.Add(xtpControlButton, ID_FIND_OPTION_ADVANCED_FIND, "A&dvanced Find...", , True)
    cbrAdvancedFind.ShortcutText = "Ctrl+Shift+F"
    Set cbrOptions = Nothing
    
    Set cbrX = g_typInterface.IFind.ActiveMenuBar.Controls.Add(xtpControlButton, ID_FIND_X, "", , True)
    cbrX.Flags = xtpFlagRightAlign
    Set cbrX = Nothing
    
End Sub

Public Sub SetObjects(ByRef Application As Object, ByRef Window As Object)

    Dim lngCtrlCtr As Long
    
    Set g_typInterface.IApplication = Application
    
    For lngCtrlCtr = 0 To Window.Controls.Count - 1
        Select Case Window.Controls(lngCtrlCtr).Tag
            Case "Buttonbar"
                Set g_typInterface.IButtonBar = Window.Controls(lngCtrlCtr)
                
            Case "CommandBar"
                Set g_typInterface.ICommandBar = Window.Controls(lngCtrlCtr)
                
            Case "CustomFind"
                Set g_typInterface.ICustomFind = Window.Controls(lngCtrlCtr)
                
            Case "Find"
                Set g_typInterface.IFind = Window.Controls(lngCtrlCtr)
                
            Case "FindBGround"
                Set g_typInterface.IFindBGround = Window.Controls(lngCtrlCtr)
                
            Case "Grid"
                Set g_typInterface.IGrid = Window.Controls(lngCtrlCtr)
                
            Case "GridHeader"
                Set g_typInterface.IGridHeader = Window.Controls(lngCtrlCtr)
                
            Case "GridImage"
                Set g_typInterface.IGridImage = Window.Controls(lngCtrlCtr)
                
            Case "GridSplitter"
                Set g_typInterface.IGridSplitter = Window.Controls(lngCtrlCtr)
                
            Case "Itemlist"
                Set g_typInterface.IItemlist = Window.Controls(lngCtrlCtr)
                
            Case "ReadingPane"
                Set g_typInterface.IReadingPane = Window.Controls(lngCtrlCtr)
                
            Case "Splash"
                Set g_typInterface.ISplash = Window.Controls(lngCtrlCtr)
                
            Case "Splitter"
                Set g_typInterface.ISplitter = Window.Controls(lngCtrlCtr)
                
            Case "SplitterImage"
                Set g_typInterface.ISplitterImage = Window.Controls(lngCtrlCtr)
                
            Case "StatusBar"
                Set g_typInterface.IStatusbar = Window.Controls(lngCtrlCtr)
                
            Case "Timer"
                Set g_typInterface.ITimer = Window.Controls(lngCtrlCtr)
                
            Case "NoView"
                Set g_typInterface.INoView = Window.Controls(lngCtrlCtr)
                
            Case "TreeIcons"
                Set g_typInterface.ITreeIcons = Window.Controls(lngCtrlCtr)
                
            Case "BBarIcons"
                Set g_typInterface.IBBarIcons = Window.Controls(lngCtrlCtr)
                
            Case "License"
                Set g_typInterface.ILicense = Window.Controls(lngCtrlCtr)
                
        End Select
    Next
    
    
End Sub

Public Sub ChangeControlSettings(ByRef Window As Object)

    '>> Change control colors
    If g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE) <> g_typInterface.IGrid.BackColorGBBox Then
        g_typInterface.IGrid.BackColorGBBox = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        g_typInterface.IGrid.BackColorHeader = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
        g_typInterface.IGrid.BackColorInfoText = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        g_typInterface.IGrid.ForeColorInfoText = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_TEXT)
        g_typInterface.IGrid.BackColorRowGroup = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
        g_typInterface.IItemlist.BackColor = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        g_typInterface.IReadingPane.BackColor = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
        g_typInterface.ICustomFind.BackColor = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        g_typInterface.IFindBGround.BackColor = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        Window.BackColor = g_typInterface.ICommandBar.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
    End If
    
    
End Sub

Public Sub LoadButtonState(ByRef Window As Object, _
                            ByVal MainSettings As CMainControls, _
                            ByVal GridProps As CGrid, _
                            ByVal ControlID As Long, _
                            ByVal EnableToolbars)

    Dim blnChecked As Boolean
    Dim lngCtr As Long
    
    Select Case ControlID
        
        Case ID_VIEW_VIEWOPTIONS
            blnChecked = MainSettings.ViewOptions
            Call EnableDisableToolbars(Window, ID_VIEW_VIEWOPTIONS, EnableToolbars, blnChecked)
        
        Case ID_VIEW_NAVIGATIONPANE
            blnChecked = MainSettings.NavigationPane
            Call EnableDisableToolbars(Window, ID_VIEW_NAVIGATIONPANE, True, blnChecked)
        
        Case ID_TOOLS_FIND
            blnChecked = MainSettings.FindBox
            Call EnableDisableToolbars(Window, ID_TOOLS_FIND, True, blnChecked)
        
        Case ID_VIEW_READINGPANE_BOTTOM
            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_BOTTOM, False)
        
        Case ID_VIEW_READINGPANE_OFF
            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_OFF, False, True)
        
        Case ID_VIEW_READINGPANE_RIGHT
            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_RIGHT, False)
        
        Case ID_VIEW_GROUPBYBOX
            blnChecked = GridProps.GroupHeaderIsVisible
            Call EnableDisableToolbars(Window, ID_VIEW_GROUPBYBOX, EnableToolbars, blnChecked)
        
        Case ID_VIEW_COLLAPSEALL
            Call EnableDisableToolbars(Window, ID_VIEW_COLLAPSEALL, EnableToolbars)
        
        Case ID_VIEW_COLLAPSEGROUP
            Call EnableDisableToolbars(Window, ID_VIEW_COLLAPSEGROUP, EnableToolbars)
        
        Case ID_VIEW_EXPANDALL
            Call EnableDisableToolbars(Window, ID_VIEW_EXPANDALL, EnableToolbars)
        
        Case ID_VIEW_EXPANDGROUP
            Call EnableDisableToolbars(Window, ID_VIEW_EXPANDGROUP, EnableToolbars)
        
        Case ID_VIEW_STATUSBAR
            blnChecked = MainSettings.StatusBar
            Call EnableDisableToolbars(Window, ID_VIEW_STATUSBAR, True, blnChecked)
        
        Case ID_VIEW_ITEMLIST_CARDVIEW
            blnChecked = GridProps.CardView
            Call EnableDisableToolbars(Window, ID_VIEW_ITEMLIST_CARDVIEW, EnableToolbars, blnChecked)
        
        Case ID_VIEW_ITEMLIST_LISTVIEW
            blnChecked = GridProps.CardView
            Call EnableDisableToolbars(Window, ID_VIEW_ITEMLIST_LISTVIEW, EnableToolbars, Not blnChecked)
        
        Case ID_VIEW_ITEMLIST_GRIDLINES
            blnChecked = GridProps.GridLines
            Call EnableDisableToolbars(Window, ID_VIEW_ITEMLIST_GRIDLINES, EnableToolbars, blnChecked)
        
        Case ID_VIEW_ITEMLIST_ODDEVENCOLOR
            blnChecked = GridProps.OddEvenColor
            Call EnableDisableToolbars(Window, ID_VIEW_ITEMLIST_ODDEVENCOLOR, EnableToolbars, blnChecked)
        
        Case ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW
            Call EnableDisableToolbars(Window, ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW, EnableToolbars)
        
        Case ID_VIEW_REFRESH
            Call EnableDisableToolbars(Window, ID_VIEW_REFRESH, EnableToolbars)
        
        '>> Button state of "Arrange By" columns
        Case Is < 0
            If g_typInterface.IGrid.Columns(Val(Right(ControlID, 2))).SortOrder <> 0 Then
                Call EnableDisableToolbars(Window, ControlID, EnableToolbars, True)
            Else
                Call EnableDisableToolbars(Window, ControlID, EnableToolbars, False)
            End If
    End Select
    '>> Not yet supported
'    If GridProps.ReadingPane = False Then
'        Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_OFF, EnableToolbars, True)
'        Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_BOTTOM, EnableToolbars)
'        Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_RIGHT, EnableToolbars)
'    Else
'        If MainSettings.ReadingPanePos = RightPane Then
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_OFF, EnableToolbars)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_BOTTOM, EnableToolbars)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_RIGHT, EnableToolbars, True)
'        ElseIf MainSettings.ReadingPanePos = BotttomPane Then
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_OFF, EnsableToolbars)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_BOTTOM, EnableToolbars, True)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_RIGHT, EnableToolbars)
'        Else
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_OFF, EnableToolbars, True)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_BOTTOM, EnableToolbars)
'            Call EnableDisableToolbars(Window, ID_VIEW_READINGPANE_RIGHT, EnableToolbars)
'        End If
'    End If


End Sub

Private Sub EnableDisableToolbars(ByRef Window As Object, _
                                ByVal ControlID As Long, _
                                ByVal Enable As Boolean, _
                                Optional Checked As Boolean = False)

    Dim mnuButton As CommandBarControl
       
    Set mnuButton = g_typInterface.ICommandBar.ActiveMenuBar.FindControl(, ControlID, , True)
    
    On Error Resume Next
    mnuButton.Enabled = Enable
    mnuButton.Checked = Checked
    On Error GoTo 0
    
    Set mnuButton = Nothing
        
End Sub

Public Sub CheckControl(Control As CommandBarControl, _
                        MainSettings As CMainControls, _
                        GridProps As CGrid)
                        
    Dim lngCtr As Long
    
    Select Case Control.ID
        
        Case ID_VIEW_NAVIGATIONPANE
            Control.Checked = MainSettings.NavigationPane
        Case ID_VIEW_READINGPANE_BOTTOM
            If MainSettings.ReadingPanePos <> BotttomPane Then
                Control.Checked = False
            Else
                Control.Checked = True
            End If
            
        Case ID_VIEW_READINGPANE_OFF
            If MainSettings.ReadingPanePos <> PaneNotVisible Then
                Control.Checked = False
            Else
                Control.Checked = True
            End If
            
        Case ID_VIEW_READINGPANE_RIGHT
            If MainSettings.ReadingPanePos <> RightPane Then
                Control.Checked = False
            Else
                Control.Checked = True
            End If
        
        Case ID_VIEW_GROUPBYBOX
            Control.Checked = GridProps.GroupHeaderIsVisible
            Control.Enabled = (Not GridProps.CardView)
            
        Case ID_VIEW_ITEMLIST_CARDVIEW
            Control.Checked = GridProps.CardView
        
        Case ID_VIEW_ITEMLIST_GRIDLINES
            Control.Checked = GridProps.GridLines
            
        Case ID_VIEW_ITEMLIST_LISTVIEW
            Control.Checked = Not (GridProps.CardView)
        
        Case ID_VIEW_ITEMLIST_ODDEVENCOLOR
            Control.Checked = GridProps.OddEvenColor
        
        Case ID_VIEW_STATUSBAR
            Control.Checked = MainSettings.StatusBar
    
        Case Else
            
            If Control.ID > -31200 And Control.ID < -31100 Then
                On Error Resume Next
                If g_typInterface.IGrid.Columns(Val(Mid(Control.ID, 5))).SortOrder <> 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
                On Error GoTo 0
            End If
        
    End Select
                        
End Sub

Public Property Get AutoRefresh() As Boolean
    AutoRefresh = m_blnAutoRefresh
End Property

Public Property Get Interval() As Double
    Interval = m_dblInterval
End Property

