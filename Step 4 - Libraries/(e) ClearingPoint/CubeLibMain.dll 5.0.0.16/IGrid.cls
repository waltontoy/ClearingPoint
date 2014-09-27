VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Enum GridSettings
    CardView = 1
    ListView = 2
    CollapseGroup = 3
    CollapseAll = 4
    ExpandGroup = 5
    ExpandAll = 6
    groupbybox = 7
    GridLines = 8
    OddEvenColor = 9
End Enum

Public Sub LoadGrid(ByRef Window As Object, ByRef MainSettings As CMainControls, ByRef NodeProps As CNavigationPane, _
                            ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection)
    Dim blnEnableToolBars As Boolean
    Dim blnChecked As Boolean
    Dim enuResult As QueryResultConstants
    
'    If Not Screen.ActiveForm Is Nothing Then
'        If Screen.ActiveForm.Name <> Window.Name Then Exit Sub
'    End If
        
    Screen.MousePointer = vbHourglass
    
    g_typInterface.IGrid.Visible = False
    
    enuResult = GridProps.SelectGridSetting(NodeProps, ADOConnection)
    
    Select Case enuResult
        Case QueryResultSuccessful
            blnEnableToolBars = BindRecords(Window, MainSettings, NodeProps, GridProps, ADOConnection)
            
        Case QueryResultNoRecord
            blnEnableToolBars = LoadDefaults(Window, MainSettings, NodeProps, GridProps, ADOConnection)
            
        Case QueryResultError
            MsgBox "Error has occured while trying to access the database. " & _
                        "Please contact technical support.", vbInformation + vbOKOnly, _
                        "Main Setting Error: GridRecords Module"
            Screen.MousePointer = vbDefault
            Exit Sub
    End Select
    
    If Not g_typInterface.IGrid.ADORecordset Is Nothing Then
        If g_typInterface.IGrid.ADORecordset.RecordCount = 0 Then
            g_typInterface.IStatusbar.Panels(1).Text = "0 item found."
        Else
            g_typInterface.IStatusbar.Panels(1).Text = g_typInterface.IGrid.ADORecordset.RecordCount & " item(s) found"
        End If
    Else
        g_typInterface.IStatusbar.Panels(1).Text = ""
    End If
    
    '>> Event to enable users to execute codes right after the grid was updated
    'Call GridProps.TriggerAfterGridUpdate - Edwin Nov09
    
    g_typInterface.IGrid.Visible = True
    
    Screen.MousePointer = vbDefault

End Sub

Private Function BindRecords(ByRef Window As Object, ByRef MainSettings As CMainControls, ByRef NodeProps As CNavigationPane, _
                            ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection) As Boolean
    
    Dim rstGridRecords As ADODB.Recordset
    Dim enuResult As QueryResultConstants
    Dim lngColCtr As Long
    Dim lngLoopTime As Long
    Dim blnAllowMissingDatabase As Boolean
    
    blnAllowMissingDatabase = False
    
    Screen.MousePointer = vbHourglass
    
    Set rstGridRecords = New ADODB.Recordset
    
    enuResult = GridProps.GridRecords(Window, NodeProps, ADOConnection, rstGridRecords, blnAllowMissingDatabase)
    
    Select Case enuResult
        Case QueryResultSuccessful
        
            g_typInterface.INoView.Left = 20000
            Set g_typInterface.IGrid.ADORecordset = rstGridRecords
            
            '>> Loop binding to check if grid already has columns after binding the recordset
            lngLoopTime = Timer
            Do While g_typInterface.IGrid.Columns.Count <> rstGridRecords.Fields.Count
                Set g_typInterface.IGrid.ADORecordset = rstGridRecords
                If (Timer - lngLoopTime) > 5 Then
                    Exit Do
                End If
            Loop
            '>> END
            
            
            GridSettings GridProps, Window, ADOConnection
                        
            BindRecords = True
                                
        Case QueryResultNoRecord
            
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
    
            BindRecords = False
            
        Case QueryResultError
            If Not blnAllowMissingDatabase Then
                MsgBox "Error has occured while trying to access the database. " & _
                            "Please contact technical support.", vbInformation + vbOKOnly, _
                            "Main Setting Error: GridRecords Module"
            Else
                Set g_typInterface.IGrid.ADORecordset = Nothing
                g_typInterface.IGrid.Columns.Clear
                
                g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
                g_typInterface.INoView.Left = 300
                
                g_typInterface.INoView.Caption = "Records are Offline"
                g_typInterface.INoView.ZOrder 0
            End If
            
            BindRecords = False
            
    End Select

    Set rstGridRecords = Nothing
    
    Screen.MousePointer = vbDefault
    
End Function

Private Function LoadDefaults(ByRef Window As Object, ByRef MainSettings As CMainControls, ByRef NodeProps As CNavigationPane, _
                            ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection) As Boolean

    Dim enuResult As QueryResultConstants
    
    Screen.MousePointer = vbHourglass
    
    If GridProps.LoadDefaultSettings(NodeProps, ADOConnection) = QueryResultSuccessful Then
        If GridProps.ViewID = 0 Then
            enuResult = QueryResultNoRecord
        Else
            enuResult = GridProps.AddGridSetting(NodeProps, ADOConnection)
        End If
        
        Select Case enuResult
            Case QueryResultSuccessful
                LoadDefaults = BindRecords(Window, MainSettings, NodeProps, GridProps, ADOConnection)
                
            Case QueryResultNoRecord
                
                g_typInterface.IGrid.GroupByBoxVisible = False
            
                g_typInterface.IStatusbar.Panels(1).Text = ""
                
                Set g_typInterface.IGrid.ADORecordset = Nothing
                
                g_typInterface.IGrid.Columns.Clear
                
                g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
                g_typInterface.INoView.Left = 300
                
                g_typInterface.INoView.ZOrder 0
                g_typInterface.INoView.Caption = "No View"
                
                LoadDefaults = False
                
            Case QueryResultError
                LoadDefaults = False
        
        End Select
    Else
        Set g_typInterface.IGrid.ADORecordset = Nothing
        g_typInterface.IGrid.Columns.Clear
        g_typInterface.IGrid.GroupByBoxVisible = False
        g_typInterface.INoView.Top = g_typInterface.IGrid.Top + 500
        g_typInterface.INoView.Left = 300
        g_typInterface.INoView.ZOrder 0
        g_typInterface.INoView.Caption = "No View"
        g_typInterface.IStatusbar.Panels(1).Text = ""
        
        LoadDefaults = False
    End If

    Screen.MousePointer = vbDefault
    
End Function

Private Sub CreateGridOffline(GridRecords As ADODB.Recordset, ByRef OfflineRecords As ADODB.Recordset)

    Dim lngFieldCtr As Long
    
    Screen.MousePointer = vbHourglass
    
    Set OfflineRecords = New ADODB.Recordset
    OfflineRecords.CursorLocation = adUseClient
    
    For lngFieldCtr = 0 To GridRecords.Fields.Count - 1
        
        OfflineRecords.Fields.Append GridRecords.Fields(lngFieldCtr).Name, _
                                    GridRecords.Fields(lngFieldCtr).Type, _
                                    GridRecords.Fields(lngFieldCtr).DefinedSize, _
                                    adFldIsNullable
    
    Next
    
    OfflineRecords.Open
    
    Screen.MousePointer = vbDefault
    
End Sub
Friend Sub GridSettings(ByRef GridProps As CGrid, ByRef Window As Object, ByRef ADOConnection As ADODB.Connection)
    
    Dim cbrComboBox As CommandBarComboBox
    Dim cbrTextBox As CommandBarEdit
    
    Dim strColumn As String
    Dim strIconPath As String
    
    Dim arrDVCIDs
    Dim arrAlignments
    Dim arrWidths
    Dim arrRequired
    Dim lngColCtr As Long
    Dim lngColIndex As Long
    Dim lngDVCID As Long
    Dim lngIDCtr As Long
    
    Dim blnIDFound As Boolean
    Dim blnColumnIsVisible As Boolean
    
    Screen.MousePointer = vbHourglass
    
    '>> Search In Combo Box
    Set cbrComboBox = g_typInterface.IFind.ActiveMenuBar.Controls(3)
    
    If cbrComboBox Is Nothing Then Exit Sub
    
    strColumn = Trim(cbrComboBox.Text)
    cbrComboBox.Clear
    
    For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
        
        If GridProps.GetDVCID(lngDVCID, g_typInterface.IGrid.Columns(lngColCtr).Caption, ADOConnection) = QueryResultSuccessful Then
            arrDVCIDs = Split(GridProps.DVCIDs, "*****")
            arrRequired = Split(GridProps.RequiredFields, "*****")
            
            blnColumnIsVisible = False
            
            For lngIDCtr = 0 To UBound(arrDVCIDs)
                If Val(arrDVCIDs(lngIDCtr)) = lngDVCID Then
                    blnColumnIsVisible = True
                    blnIDFound = True
                    Exit For
                End If
            Next
            
            If blnColumnIsVisible = False Then
                For lngIDCtr = 0 To UBound(arrRequired)
                    If Val(arrRequired(lngIDCtr)) = lngDVCID Then
                        blnColumnIsVisible = False
                        blnIDFound = True
                        Exit For
                    End If
                Next
            End If
            
            If blnIDFound = True Then
                
                g_typInterface.IGrid.Columns(lngColCtr).Visible = blnColumnIsVisible
                
                If blnColumnIsVisible = True Then
                    If g_typInterface.IGrid.Columns(lngColCtr).Visible = True Then
                        cbrComboBox.AddItem g_typInterface.IGrid.Columns(lngColCtr).Caption
                    End If
                    
                    If g_typInterface.IGrid.Columns(lngColCtr).Caption = strColumn Then
                        lngColIndex = cbrComboBox.ListCount
                    End If
                End If
                
                g_typInterface.IGrid.Columns(lngColCtr).Tag = lngDVCID
            Else
                g_typInterface.IGrid.Columns(lngColCtr).Tag = "Added"
            End If
        Else
            g_typInterface.IGrid.Columns(lngColCtr).Tag = "Added"
        End If
        
        Call GridProps.TriggerSetColumnFormat(g_typInterface.IGrid.Columns(lngColCtr))
        
    Next

    If lngColIndex > 0 Then
        cbrComboBox.ListIndex = lngColIndex
    End If
    
    GridAutoFormat GridProps, Window, ADOConnection
    
    arrDVCIDs = Split(GridProps.DVCIDs, "*****")
    arrAlignments = Split(GridProps.Alignments, "*****")
    arrWidths = Split(GridProps.Widths, "*****")
    
    For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
        If (g_typInterface.IGrid.Columns(lngColCtr).Visible = True) Then
            If (IsNumeric(g_typInterface.IGrid.Columns(lngColCtr).Tag) = True) Then
                Select Case UCase(Trim(arrAlignments(lngColCtr - 1)))
                    Case "LEFT"
                        g_typInterface.IGrid.Columns(lngColCtr).TextAlignment = 0
                    Case "CENTER"
                        g_typInterface.IGrid.Columns(lngColCtr).TextAlignment = 1
                    Case "RIGHT"
                        g_typInterface.IGrid.Columns(lngColCtr).TextAlignment = 2
                End Select
                
                g_typInterface.IGrid.Columns(lngColCtr).Width = CLng(arrWidths(lngColCtr - 1))
                
                strIconPath = Trim(GridProps.HeaderIconPath(ADOConnection, CLng(arrDVCIDs(lngColCtr - 1))))
                
                If (Len(Trim(strIconPath)) > 0) Then
                    
                    Select Case g_typInterface.IGrid.Columns(lngColCtr).ColumnType
                        Case jgexText
                            
                        Case jgexIcon
                            g_typInterface.IGrid.GridImages.Add LoadPicture(strIconPath, , vbLPColor)
                            g_typInterface.IGrid.Columns(lngColCtr).HeaderIcon = g_typInterface.IGrid.GridImages.Count
                            g_typInterface.IGrid.Columns(lngColCtr).Caption = ""
                            
                        Case jgexIconAndText
                            g_typInterface.IGrid.GridImages.Add LoadPicture(strIconPath, , vbLPColor)
                            g_typInterface.IGrid.Columns(lngColCtr).HeaderIcon = g_typInterface.IGrid.GridImages.Count
                            
                    End Select
                    
                    ' Delete the temporary image file
                    If (Len(Dir(strIconPath)) > 0) Then
                        On Error Resume Next
                        Kill strIconPath
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
        
    Next
        
    LoadGridProperties GridProps, Window, ADOConnection
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GridColumnMove(GridProps As CGrid, ByRef Window As Object, _
                            ADOConnection As ADODB.Connection)

    Dim lngDVCID As Long
    Dim lngColCtr As Long
    
    Screen.MousePointer = vbHourglass
    
    GridProps.DVCIDs = ""
    GridProps.Widths = ""
    
    For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
        If g_typInterface.IGrid.Columns(lngColCtr).Visible = True _
            And IsNumeric(g_typInterface.IGrid.Columns(lngColCtr).Tag) = True Then
                ' Reims - Fixed: Cannot use g_typInterface.IGrid.Columns.ItemByPosition(lngColCtr).Caption to
                '                look for corresponding record in DefaultViewColumns because the caption is
                '                an ampty string for iconic columns
                'GridProps.GetDVCID lngDVCID, g_typInterface.IGrid.Columns.ItemByPosition(lngColCtr), ADOConnection
                
                lngDVCID = Val(g_typInterface.IGrid.Columns.ItemByPosition(lngColCtr).Tag)
                Debug.Assert (lngDVCID > 0) ' lngDVCID should be a positive number here
                
                GridProps.DVCIDs = GridProps.DVCIDs & lngDVCID & "*****"
                GridProps.Widths = GridProps.Widths & g_typInterface.IGrid.Columns.ItemByPosition(lngColCtr).Width & "*****"
        End If
    Next
    If Len(GridProps.DVCIDs) > 5 Then
        GridProps.DVCIDs = Mid(GridProps.DVCIDs, 1, Len(GridProps.DVCIDs) - 5)
    End If
    If Len(GridProps.Widths) > 5 Then
        GridProps.Widths = Mid(GridProps.Widths, 1, Len(GridProps.Widths) - 5)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GridColumnResize(GridProps As CGrid, ByRef Window As Object, _
                            ByVal ColIndex As Long, ByVal NewColWidth As Long, _
                            ADOConnection As ADODB.Connection)

    Dim lngDVCID As Long
    Dim lngColCtr As Long
    
    Screen.MousePointer = vbHourglass
    
    GridProps.Widths = ""
    
    For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
        If (g_typInterface.IGrid.Columns(lngColCtr).Visible = True _
            Or g_typInterface.IGrid.Columns(lngColCtr).IsGrouped = True) _
            And IsNumeric(g_typInterface.IGrid.Columns(lngColCtr).Tag) = True Then
                If g_typInterface.IGrid.Columns(ColIndex).Tag = g_typInterface.IGrid.Columns(lngColCtr).Tag Then
                    GridProps.Widths = GridProps.Widths & NewColWidth & "*****"
                Else
                    GridProps.Widths = GridProps.Widths & g_typInterface.IGrid.Columns(lngColCtr).Width & "*****"
                End If
        Else
            '<<< Dandan 6/26/2006
            GridProps.Widths = GridProps.Widths & g_typInterface.IGrid.Columns(lngColCtr).Width & "*****"
        End If
    Next
    
    If Len(GridProps.Widths) > 5 Then
        GridProps.Widths = Mid(GridProps.Widths, 1, Len(GridProps.Widths) - 5)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Friend Sub LoadGridProperties(ByRef GridProps As CGrid, ByRef Window As Object, Optional ByRef ADOConnection As ADODB.Connection)
    
    Dim arrDVC
    Dim arrSorting
    Dim arrGrouping
    Dim arrColFormat
    Dim arrCells
    Dim arrSelected
    Dim lngCtr As Long
    Dim lngIDCtr As Long
    Dim lngGroupCtr As Long
    Dim lngSortCtr As Long
    Dim strAlias As String
    Dim lngTimer As Long
    
    Dim varJSSortKey As JSSortKey
    Dim rstDataType As ADODB.Recordset
    Dim strCommand As String
    
    
    Screen.MousePointer = vbHourglass
    
    g_typInterface.IGrid.Visible = False
    
    '****************************************************************************************
    'g_blnFromEditViewForm was added for this release
    '****************************************************************************************
    'SUBJECT: ADDENDUM: New ClearingPoint Version Released: CP 4.14.14 (RC-1)
    'ITEM   : The optimization of refreshing of the itemlist view has yet to be designed if possible.
    '****************************************************************************************
    If g_blnFromEditViewForm = True Then
        If Trim(GridProps.Sort) <> "" Then
            arrSorting = Split(GridProps.Sort, "*****")
            
            g_typInterface.IGrid.SortKeys.Clear
            For lngSortCtr = 0 To UBound(arrSorting) Step 2
                For lngCtr = 1 To g_typInterface.IGrid.Columns.Count
                    If g_typInterface.IGrid.Columns(lngCtr).Caption = arrSorting(lngSortCtr) Then
                        'g_typInterface.IGrid.SortKeys.Add g_typInterface.IGrid.Columns(lngCtr).Index, Val(arrSorting(lngSortCtr + 1))
                        
                        Set varJSSortKey = g_typInterface.IGrid.SortKeys.Add(g_typInterface.IGrid.Columns(lngCtr).Index, Val(arrSorting(lngSortCtr + 1)))
        
                        strCommand = ""
                        strCommand = strCommand & "SELECT "
                        strCommand = strCommand & "TOP 1 "
                        strCommand = strCommand & "DefaultViewColumns.DVC_DataType AS DVC_DataType "
                        strCommand = strCommand & "FROM "
                        strCommand = strCommand & "DefaultViewColumns "
                        strCommand = strCommand & "WHERE "
                        strCommand = strCommand & "DVC_ID = " & Val(g_typInterface.IGrid.Columns(lngCtr).Tag)
                        
                        ADORecordsetOpen strCommand, ADOConnection, rstDataType, adOpenKeyset, adLockOptimistic
                        'RstOpen strCommand, ADOConnection, rstDataType, adOpenKeyset, adLockOptimistic, , True
                        If rstDataType.RecordCount > 0 Then
                            rstDataType.MoveFirst
                            If Trim$(UCase$(rstDataType.Fields("DVC_DataType").Value)) = "DATE" Then
                                g_typInterface.IGrid.Columns(g_typInterface.IGrid.Columns(lngCtr).Index).SortType = jgexSortTypeDateTime
                            End If
                        End If
                        ADORecordsetClose rstDataType
                        Exit For
                    End If
                Next
            Next
        End If
    End If
    
    g_typInterface.IGrid.UseEvenOddColor = GridProps.OddEvenColor
    g_typInterface.IGrid.GridLines = GridProps.GridLines
    g_typInterface.IGrid.View = IIf(GridProps.CardView, jgexCard, jgexTable)
    If GridProps.CardView = True Then
        GridProps.GroupHeaders = ""
        GridProps.GroupHeaderIsVisible = False
    End If
    g_typInterface.IGrid.GroupByBoxVisible = GridProps.GroupHeaderIsVisible
    
    '>> Groupings
    If Trim(GridProps.GroupHeaders) <> "" Then
        arrDVC = Split(GridProps.DVCIDs, "*****")
        arrGrouping = Split(GridProps.GroupHeaders, "*****")
        g_typInterface.IGrid.Groups.Clear
        
        For lngCtr = 0 To UBound(arrGrouping) Step 3
            For lngIDCtr = 0 To UBound(arrDVC)
                If arrGrouping(lngCtr) = arrDVC(lngIDCtr) Then
                    If Val(arrGrouping(lngCtr + 1)) = 1 Then
                        g_typInterface.IGrid.Groups.Add (lngIDCtr + 1), jgexSortAscending
                    Else
                        g_typInterface.IGrid.Groups.Add (lngIDCtr + 1), jgexSortDescending
                    End If
                    If Val(arrGrouping(lngCtr + 2)) = 1 Then
                        g_typInterface.IGrid.Columns(lngIDCtr + 1).Visible = True
                    Else
                        g_typInterface.IGrid.Columns(lngIDCtr + 1).Visible = False
                    End If
                    Exit For
                End If
            Next
        Next
        
    End If
    
    '>> Column formats
    If Trim(GridProps.ColumnFormats) <> "" Then
        arrColFormat = Split(GridProps.ColumnFormats, "*****")
        
        For lngCtr = 0 To UBound(arrColFormat)
            
            arrCells = Split(arrColFormat(lngCtr), "~")
            
            For lngIDCtr = 1 To g_typInterface.IGrid.Columns.Count
                If Val(arrCells(0)) = Val(g_typInterface.IGrid.Columns(lngIDCtr).Tag) Then
                    g_typInterface.IGrid.Columns(lngIDCtr).Format = Trim(arrCells(1))
                End If
            Next
            
        Next
        
    End If

    lngGroupCtr = 0
    g_typInterface.IGrid.RefreshGroups
    If GridProps.ExpandCollapseDefault = 0 And Trim(GridProps.GroupRowFormat) <> "" Then
        For lngCtr = 1 To g_typInterface.IGrid.RowCount
            If g_typInterface.IGrid.IsGroupItem(lngCtr) = True Then
                lngGroupCtr = lngGroupCtr + 1
                g_typInterface.IGrid.RowExpanded(lngCtr) = CBool(Val(Mid(GridProps.GroupRowFormat, lngGroupCtr, 1)))
            End If
        Next
    ElseIf GridProps.ExpandCollapseDefault = 1 Then
        g_typInterface.IGrid.ExpandAll
    Else
        g_typInterface.IGrid.CollapseAll
    End If
    
    '******************************************************************************************
    'Edwin Nov09
    '******************************************************************************************
    GridProps.TriggerAfterGridUpdate
    
    
    
    If g_blnFromEditViewForm Then
        g_typInterface.IGrid.Visible = True
        g_typInterface.IGrid.Refresh
        
        g_blnFromEditViewForm = False
    End If
    
    If Trim(GridProps.SelectedRows) <> "" Then
        If g_typInterface.IGrid.Row > 0 Then
            g_typInterface.IGrid.RowSelected(g_typInterface.IGrid.Row) = False
        End If
        
        arrSelected = Split(GridProps.SelectedRows, "~")
        
        For lngCtr = 0 To UBound(arrSelected)
            g_typInterface.IGrid.SelectedItems.Add Val(arrSelected(lngCtr))
        Next
        
        If g_typInterface.IGrid.RowCount > 0 Then
            g_typInterface.IGrid.EnsureVisible Val(arrSelected(0))
        End If
    End If
    
'    g_typInterface.IGrid.Visible = True
'
'    g_typInterface.IGrid.Refresh
'    If Trim(GridProps.SelectedRows) <> "" Then
'        If g_typInterface.IGrid.Row > 0 Then
'            g_typInterface.IGrid.RowSelected(g_typInterface.IGrid.Row) = False
'        End If
'        arrSelected = Split(GridProps.SelectedRows, "~")
'        '>> Select last selected rows
'        For lngCtr = 0 To UBound(arrSelected)
'            g_typInterface.IGrid.RowSelected(Val(arrSelected(lngCtr))) = True
'        Next
'        '>> Ensure that first selected row is visible
'        If g_typInterface.IGrid.RowCount > 0 Then
'            g_typInterface.IGrid.EnsureVisible Val(arrSelected(0))
'            g_typInterface.IGrid.EnsureVisible Val(arrSelected(UBound(arrSelected)))
'        End If
'    End If
    '******************************************************************************************
    
    Screen.MousePointer = vbDefault
    
End Sub

Friend Sub GridAutoFormat(ByRef GridProps As CGrid, ByRef Window As Object, ByRef ADOConnection As ADODB.Connection)

    Dim rstFormat As ADODB.Recordset
    Dim lngImageCount As Long
    Dim lngColCtr As Long
    Dim strPicPath As String
    Dim fmtCondition As JSFmtCondition
    Dim enuResult As QueryResultConstants
    Dim blnColumnExists As Boolean
    
    Screen.MousePointer = vbHourglass
    
    enuResult = GridProps.GetFormatConditions(GridProps, rstFormat, ADOConnection)
    
    Select Case enuResult
        Case QueryResultSuccessful
            '>> Clear images and formats
            g_typInterface.IGrid.GridImages.Clear
            g_typInterface.IGrid.FmtConditions.Clear
            For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
                If g_typInterface.IGrid.Columns(lngColCtr).HasValueList = True Then
                    g_typInterface.IGrid.Columns(lngColCtr).ValueList.Clear
                End If
            Next
            '>> END
            
            Do While Not rstFormat.EOF
                blnColumnExists = False
                For lngColCtr = 1 To g_typInterface.IGrid.Columns.Count
                    If UCase$(Trim(g_typInterface.IGrid.Columns(lngColCtr).Caption)) = UCase$(Trim(rstFormat!FC_Field)) Then
                        blnColumnExists = True
                        Exit For
                    End If
                Next
                
                If blnColumnExists Then
                    'g_typInterface.IGrid.FmtConditions.Add g_typInterface.IGrid.Columns(rstFormat!FC_Field).Index, _
                                                     Val(rstFormat!FC_Operator), rstFormat!FC_Value1, _
                                                    rstFormat!FC_Value2, rstFormat!FC_Name
                
                    'g_typInterface.IGrid.FmtConditions(rstFormat!FC_Name).FormatStyle.FontBold = rstFormat!FC_FontBold
                    'g_typInterface.IGrid.FmtConditions(rstFormat!FC_Name).FormatStyle.FontItalic = rstFormat!FC_FontItalic
                    'g_typInterface.IGrid.FmtConditions(rstFormat!FC_Name).FormatStyle.FontUnderline = rstFormat!FC_FontUnderline
                    'g_typInterface.IGrid.FmtConditions(rstFormat!FC_Name).FormatStyle.FontStrikethru = rstFormat!FC_FontStrikeThru
                    'g_typInterface.IGrid.FmtConditions(rstFormat!FC_Name).FormatStyle.ForeColor = rstFormat!FC_ForeColor
                    
                    If rstFormat!FC_ColumnType > 0 Then
                        strPicPath = vbNullString
                        strPicPath = PicPath(rstFormat!FC_ID, rstFormat)
                        If (Len(Trim(strPicPath)) > 0) Then
                            g_typInterface.IGrid.GridImages.Add LoadPicture(strPicPath, , vbLPColor)
                            lngImageCount = g_typInterface.IGrid.GridImages.Count
                            
                            ' Delete the temporary image file
                            If (Len(Dir(strPicPath)) > 0) Then
                                On Error Resume Next
                                Kill strPicPath
                                On Error GoTo 0
                            End If
                        End If
                        
                        g_typInterface.IGrid.Columns(rstFormat!FC_Field).HasValueList = True
                        
                        Select Case rstFormat!FC_ColumnType
                            Case 1
                                g_typInterface.IGrid.Columns(rstFormat!FC_Field).ValueList.Add rstFormat!FC_Value1, rstFormat!FC_ColumnText
                            Case 2
                                g_typInterface.IGrid.Columns(rstFormat!FC_Field).ValueList.Add rstFormat!FC_Value1, "", lngImageCount
                            Case 3
                                g_typInterface.IGrid.Columns(rstFormat!FC_Field).ValueList.Add rstFormat!FC_Value1, rstFormat!FC_ColumnText, lngImageCount
                        End Select
                        g_typInterface.IGrid.Columns(rstFormat!FC_Field).ColumnType = rstFormat!FC_ColumnType - 1
                    End If
                End If
                
                rstFormat.MoveNext
            Loop
            
        Case QueryResultNoRecord
        
        Case QueryResultError
        
    End Select
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstFormat)
    
    Screen.MousePointer = vbDefault
    
End Sub


Public Sub GridColumnSort(ByRef GridProps As CGrid, ByRef Window As Object, _
                            ByRef Column As GridEX16.JSColumn, _
                            ADOConnection As ADODB.Connection)
                            
    Dim lngCtr As Long
    
    Dim intSortType As Integer
    Dim lngDVCID As Long
    Dim varJSSortKey As JSSortKey
    Dim rstDataType As ADODB.Recordset
    Dim strCommand As String
    
    Screen.MousePointer = vbHourglass
    
    GridProps.Sort = ""

    If Column.SortOrder = 0 Then
        intSortType = 1
    Else
        If Column.SortOrder = 1 Then
            intSortType = -1
        Else
            intSortType = 1
        End If
    End If

    g_typInterface.IGrid.SortKeys.Clear
    Set varJSSortKey = g_typInterface.IGrid.SortKeys.Add(Column.Index, intSortType)
    
    strCommand = ""
    strCommand = strCommand & "SELECT "
    strCommand = strCommand & "TOP 1 "
    strCommand = strCommand & "DefaultViewColumns.DVC_DataType AS DVC_DataType "
    strCommand = strCommand & "FROM "
    strCommand = strCommand & "DefaultViewColumns "
    strCommand = strCommand & "WHERE "
    strCommand = strCommand & "DVC_ID = " & Val(Column.Tag)
    ADORecordsetOpen strCommand, ADOConnection, rstDataType, adOpenKeyset, adLockOptimistic
    'RstOpen strCommand, ADOConnection, rstDataType, adOpenKeyset, adLockOptimistic, , True
    If rstDataType.RecordCount > 0 Then
        rstDataType.MoveFirst
        If Trim$(UCase$(rstDataType.Fields("DVC_DataType").Value)) = "DATE" Then
            g_typInterface.IGrid.Columns(Column.Index).SortType = jgexSortTypeDateTime
        End If
    End If
    ADORecordsetClose rstDataType
    
    If IsNumeric(Column.Tag) = True Then
        GridProps.Sort = Column.Caption & "*****" & intSortType
    
        If GridProps.AutoGroup = True Then
            GridProps.GetDVCID lngDVCID, Column.Caption, ADOConnection
            GridProps.GroupHeaders = lngDVCID & "*****" & intSortType & "*****1"
            
        Else
            GridProps.GetDVCID lngDVCID, Column.Caption, ADOConnection
            If intSortType = 1 Then
                GridProps.GroupHeaders = Replace(GridProps.GroupHeaders, lngDVCID & "*****-1", lngDVCID & "*****1")
            Else
                GridProps.GroupHeaders = Replace(GridProps.GroupHeaders, lngDVCID & "*****1", lngDVCID & "*****-1")
            End If
        End If
        
        Me.LoadGridProperties GridProps, Window, ADOConnection
    End If
    
    ' This is needed to make the grid visible again if sorting is done by
    ' clicking the column headers
    g_typInterface.IGrid.Visible = True
    
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub FindItem(ByRef GridProps As CGrid, ByRef Window As Object, ByRef ADOConnection As ADODB.Connection, ByVal SearchKey As Variant, ByVal SearchIn As String)
    
    Dim rstClone As ADODB.Recordset
    Dim cbrClear As CommandBarButton
    
    If g_typInterface.IGrid.ADORecordset Is Nothing Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    g_typInterface.IGrid.Visible = False
    
    If Trim(SearchIn) <> "" And Trim(SearchKey) <> "" Then
        Set rstClone = g_typInterface.IGrid.ADORecordset.Clone
        Select Case rstClone.Fields(SearchIn).Type
            Case adDouble, adNumeric, adInteger, adUnsignedTinyInt
                If IsNumeric(SearchKey) Then
                    rstClone.Filter = "[" & SearchIn & "] = " & SearchKey
                Else
                    rstClone.Filter = "[" & SearchIn & "] = 0 "
                End If
            
            Case adBoolean
                On Error Resume Next
                rstClone.Filter = "[" & SearchIn & "] = " & CBool(SearchKey)
                On Error GoTo 0
                
            Case adDate
                If IsDate(SearchKey) Then
                    rstClone.Filter = "[" & SearchIn & "] = #" & SearchKey & "#"
                End If
                
            Case Else
                If IsNumeric(SearchKey) Then
                    rstClone.Filter = "[" & SearchIn & "] = " & SearchKey
                Else
                    rstClone.Filter = "[" & SearchIn & "] Like '%" & SearchKey & "%'"
                End If
        End Select
        
        g_typInterface.IStatusbar.Panels(1).Text = rstClone.RecordCount & " item(s) found"
        
        Set g_typInterface.IGrid.ADORecordset = rstClone
        
        Set rstClone = Nothing
                
        GridSettings GridProps, Window, ADOConnection
        
        ' Reims - sa iba ba inilagay ni mark ito?
         GridProps.TriggerAfterGridUpdate
        
        Set cbrClear = g_typInterface.IFind.ActiveMenuBar.Controls(5) '>> Clear button
        cbrClear.Enabled = True
        Set cbrClear = Nothing
    End If
    
    g_typInterface.IGrid.Visible = True
    
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub ClearFilter(ByRef GridProps As CGrid, ByRef Window As Object, ByRef ADOConnection As ADODB.Connection)
    Dim rstClone As ADODB.Recordset
    
    Dim cbrClear As CommandBarButton
    Dim cbrSearchIn As CommandBarComboBox
    Dim cbrFindTextBox As CommandBarEdit
    
    If g_typInterface.IGrid.ADORecordset Is Nothing Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    ' Reims - sa iba ba inilagay ni mark ito?
    g_typInterface.IGrid.Visible = False
    
    Set rstClone = g_typInterface.IGrid.ADORecordset.Clone
    rstClone.Filter = adFilterNone
    
    g_typInterface.IStatusbar.Panels(1).Text = rstClone.RecordCount & " item(s) found"
    
    Set g_typInterface.IGrid.ADORecordset = rstClone
    
    Set rstClone = Nothing
    
        
    GridSettings GridProps, Window, ADOConnection
    
    ' Reims - sa iba ba inilagay ni mark ito?
     GridProps.TriggerAfterGridUpdate
            
    ' Reims - sa iba ba inilagay ni mark ito?
    g_typInterface.IGrid.Visible = True
    
    
    ' Disable 'Clear' button
    Set cbrClear = g_typInterface.IFind.ActiveMenuBar.FindControl(xtpControlButton, ID_FIND_CLEAR)
    If (cbrClear Is Nothing = False) Then
        cbrClear.Enabled = False
    End If
    Set cbrClear = Nothing
    
    ' Clear 'Search In' combobox
    Set cbrSearchIn = g_typInterface.IFind.ActiveMenuBar.FindControl(xtpControlComboBox, ID_FIND_SEARCH_BOX)
    If (cbrSearchIn Is Nothing = False) Then
        cbrSearchIn.Text = vbNullString
    End If
    Set cbrSearchIn = Nothing
    
    ' Clear 'Look For' textbox
    Set cbrFindTextBox = g_typInterface.IFind.ActiveMenuBar.FindControl(xtpControlEdit, ID_FIND_LOOK_FOR)
    If (cbrFindTextBox Is Nothing = False) Then
        cbrFindTextBox.Text = vbNullString
    End If
    Set cbrFindTextBox = Nothing
    
    
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub GridGroupingChange(ByRef GridProps As CGrid, ByRef Window As Object, ADOConnection As ADODB.Connection)

    Dim lngGrpCtr As Long
    Dim lngDVCID As Long
    Dim lngIndex As Long
    
    Screen.MousePointer = vbHourglass
    
    GridProps.GroupHeaders = ""
    
    For lngGrpCtr = 1 To g_typInterface.IGrid.Groups.Count
        lngIndex = g_typInterface.IGrid.Groups(lngGrpCtr).ColIndex
        If g_typInterface.IGrid.Columns(lngIndex).IsGrouped = True _
            And IsNumeric(g_typInterface.IGrid.Columns(lngIndex).Tag) = True Then
                GridProps.GetDVCID lngDVCID, g_typInterface.IGrid.Columns(lngIndex).Caption, ADOConnection
                GridProps.GroupHeaders = GridProps.GroupHeaders & lngDVCID & "*****" & _
                                        g_typInterface.IGrid.Columns(lngIndex).SortOrder & _
                                        "*****1*****"
                g_typInterface.IGrid.Columns(lngIndex).Visible = True
        End If
    Next
    
    If Len(GridProps.GroupHeaders) > 5 Then
        GridProps.GroupHeaders = Mid(GridProps.GroupHeaders, 1, Len(GridProps.GroupHeaders) - 5)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GroupExpandCollapse(ByRef GridProps As CGrid, ByRef Window As Object)
    
    Dim lngCtr As Long
        
    Screen.MousePointer = vbHourglass
    
    GridProps.GroupRowFormat = ""

    For lngCtr = 1 To g_typInterface.IGrid.RowCount
        If g_typInterface.IGrid.IsGroupItem(lngCtr) = True Then
            If g_typInterface.IGrid.RowExpanded(lngCtr) = True Then
                GridProps.GroupRowFormat = GridProps.GroupRowFormat & "1"
            Else
                GridProps.GroupRowFormat = GridProps.GroupRowFormat & "0"
            End If
        End If
    Next
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GetLastSelectedRows(GridProps As CGrid, ByRef Window As Object)
    
    Dim lngCtr As Long
    
    Screen.MousePointer = vbHourglass
    
    GridProps.SelectedRows = ""
    
    If g_typInterface.IGrid.SelectedItems.Count > 0 Then
        GridProps.SelectedRows = ""
        For lngCtr = 1 To g_typInterface.IGrid.SelectedItems.Count
            GridProps.SelectedRows = GridProps.SelectedRows & g_typInterface.IGrid.SelectedItems.Item(lngCtr).RowPosition & "~"
        Next
        GridProps.SelectedRows = Mid(GridProps.SelectedRows, 1, Len(GridProps.SelectedRows) - 1)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub


Public Sub ToggleGrid(ByRef GridProps As CGrid, ByRef Window As Object, ToggleSetting As GridSettings)

    Screen.MousePointer = vbHourglass
    
    Select Case ToggleSetting
        Case groupbybox
            GridProps.GroupHeaderIsVisible = Not GridProps.GroupHeaderIsVisible
            g_typInterface.IGrid.GroupByBoxVisible = GridProps.GroupHeaderIsVisible
            
        Case CardView
            GridProps.CardView = True
            g_typInterface.IGrid.View = GridProps.CardView
            
        Case ListView
            GridProps.CardView = False
            g_typInterface.IGrid.View = GridProps.CardView
            
        Case OddEvenColor
            GridProps.OddEvenColor = Not GridProps.OddEvenColor
            g_typInterface.IGrid.UseEvenOddColor = GridProps.OddEvenColor
            
        Case GridLines
            GridProps.GridLines = Not GridProps.GridLines
            g_typInterface.IGrid.GridLines = GridProps.GridLines
            
        Case ExpandGroup
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True
                End If
            End If
        
        Case CollapseGroup
            If g_typInterface.IGrid.SelectedItems.Count > 0 Then
                If g_typInterface.IGrid.IsGroupItem(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = True Then
                    g_typInterface.IGrid.RowExpanded(g_typInterface.IGrid.SelectedItems.Item(1).RowPosition) = False
                End If
            End If
        
        Case ExpandAll
            g_typInterface.IGrid.ExpandAll
            
        Case CollapseAll
            g_typInterface.IGrid.CollapseAll
            
    End Select
    
    Screen.MousePointer = vbDefault
    
End Sub


Public Sub SaveGridSetting(ByRef NodeProps As CNavigationPane, ByRef GridProps As CGrid, _
                            ByRef ADOConnection As ADODB.Connection)

    Dim enuResult As QueryResultConstants
    
    Screen.MousePointer = vbHourglass
    
    enuResult = GridProps.SelectGridSetting(NodeProps, ADOConnection, True)
    
    Select Case enuResult
        Case QueryResultSuccessful
            GridProps.UpdateGridSetting ADOConnection
        Case QueryResultNoRecord
        Case QueryResultError
        
    End Select
    
    Screen.MousePointer = vbDefault
    
End Sub

