VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarDVCIDs As String 'local copy
Private mvarUserID As Long 'local copy
Private mvarNodeID As Long 'local copy
Private mvarTView_ID As Long 'local copy
Private mvarAlignments As String 'local copy
Private mvarWidths As String 'local copy
Private mvarSort As String 'local copy
Private mvarGroupHeaders As String 'local copy
Private mvarGroupHeaderIsVisible As Boolean 'local copy
Private mvarCardView As Boolean 'local copy
Private mvarOddEvenColor As Boolean 'local copy
Private mvarGridLines As Boolean 'local copy
Private mvarRequiredFields As String 'local copy
Private mvarGroupRowFormat As String 'local copy
Private mvarSelectedRows As String 'local copy
Private mvarViewID As Long 'local copy
Private mvarAutoGroup As Boolean 'local copy
Private mvarExpandCollapseDefault As Long 'local copy
Private mvarReadingPane As Boolean 'local copy
Private mvarUVC_ID As Long 'local copy
Private mvarColumnFormats As String 'local copy
Private mvarDataChanged As Boolean 'local copy

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent KeywordValue[(arg1, arg2, ... , argn)]
Public Event KeywordValue(ByVal Keyword As String, ByRef KeyValue As Variant)
Public Event GetObjects(ByRef Window As Object, _
                        ByRef MainSettings As CMainControls, _
                        ByRef GridProps As IGrid, _
                        ByRef NavPane As CNavigationPane, _
                        ByRef ADOConnection As ADODB.Connection)

Public Event AfterGridUpdate()

Public Event BeforeOpenRecordset(ByRef SelectClause As String, _
                                ByRef FromClause As String, _
                                ByRef WhereClause As String, _
                                ByRef GroupByClause As String, _
                                ByRef OrderByClause As String, _
                                ByRef AllowMissingDatabase As Boolean)

Public Event SetColumnFormat(ByRef Column As JSColumn)  '>> To change format/setting of every column shown in the grid
Public Event BeforeShowFields(ByRef FieldDescriptions As ADODB.Recordset)    '>> Retrieve all descriptions of columns

Public Property Let DataChanged(ByVal vData As Boolean)

    Dim objMain As Object
    Dim clsMain As CMainControls
    Dim clsGrid As IGrid
    Dim clsTree As CNavigationPane
    Dim conMain As ADODB.Connection
    
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataChanged = 5
    mvarDataChanged = vData
    
    If mvarDataChanged = True Then
        RaiseEvent GetObjects(objMain, clsMain, clsGrid, clsTree, conMain)
        
        If (clsTree.tvwMainControl.SelectedItem Is Nothing = False) Then
            ' Inserted by Reims - In case the selected node has been change
            ' ex. Go To Another Folder Command from Menu or Toolbar
            If (Len(Trim(clsTree.tvwMainControl.SelectedItem.Key)) > 0) Then
                Set clsTree.SelectedNode = clsTree.tvwMainControl.SelectedItem
                clsTree.NodeID = CLng(clsTree.tvwMainControl.SelectedItem.Key)
            End If
        Else
            ' Inserted by 2dit - error occurs when copying / moving items
            ' from one folder to another when the selected node is in the favorites pane
            If (clsTree.tvwShortcutControl.SelectedItem Is Nothing = False) Then
                Set clsTree.SelectedNode = clsTree.tvwShortcutControl.SelectedItem
                clsTree.NodeID = CLng(clsTree.tvwShortcutControl.SelectedItem.Key)
            End If
        End If
        
        If (clsTree.SelectedNode Is Nothing = False) Then
            Call clsGrid.LoadGrid(objMain, clsMain, clsTree, Me, conMain)
            mvarDataChanged = False
        End If
    End If
    
    Set objMain = Nothing
    Set clsMain = Nothing
    Set clsGrid = Nothing
    Set clsTree = Nothing
    Set conMain = Nothing
    
End Property


Public Property Get DataChanged() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataChanged
    DataChanged = mvarDataChanged
End Property

Friend Property Let ColumnFormats(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnFormats = 5
    mvarColumnFormats = vData
End Property


Friend Property Get ColumnFormats() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnFormats
    ColumnFormats = mvarColumnFormats
End Property



Friend Property Let UVC_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UVC_ID = 5
    mvarUVC_ID = vData
End Property


Friend Property Get UVC_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UVC_ID
    UVC_ID = mvarUVC_ID
End Property



Friend Property Let ReadingPane(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReadingPane = 5
    mvarReadingPane = vData
End Property


Friend Property Get ReadingPane() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadingPane
    ReadingPane = mvarReadingPane
End Property



Friend Property Let ExpandCollapseDefault(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpandCollapseDefault = 5
    mvarExpandCollapseDefault = vData
End Property


Friend Property Get ExpandCollapseDefault() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpandCollapseDefault
    ExpandCollapseDefault = mvarExpandCollapseDefault
End Property



Friend Property Let AutoGroup(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoGroup = 5
    mvarAutoGroup = vData
End Property


Friend Property Get AutoGroup() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoGroup
    AutoGroup = mvarAutoGroup
End Property



Friend Property Let ViewID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ViewID = 5
    mvarViewID = vData
End Property


Friend Property Get ViewID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ViewID
    ViewID = mvarViewID
End Property



Friend Property Let SelectedRows(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedRows = 5
    mvarSelectedRows = vData
End Property


Friend Property Get SelectedRows() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedRows
    SelectedRows = mvarSelectedRows
End Property



Friend Property Let GroupRowFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupRowFormat = 5
    mvarGroupRowFormat = vData
End Property


Friend Property Get GroupRowFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupRowFormat
    GroupRowFormat = mvarGroupRowFormat
End Property

Friend Property Let RequiredFields(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RequiredFields = 5
    mvarRequiredFields = vData
End Property


Friend Property Get RequiredFields() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RequiredFields
    RequiredFields = mvarRequiredFields
End Property



Friend Property Let GridLines(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridLines = 5
    mvarGridLines = vData
End Property


Friend Property Get GridLines() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridLines
    GridLines = mvarGridLines
End Property



Friend Property Let OddEvenColor(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OddEvenColor = 5
    mvarOddEvenColor = vData
End Property


Friend Property Get OddEvenColor() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OddEvenColor
    OddEvenColor = mvarOddEvenColor
End Property



Friend Property Let CardView(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CardView = 5
    mvarCardView = vData
End Property


Friend Property Get CardView() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CardView
    CardView = mvarCardView
End Property



Friend Property Let GroupHeaderIsVisible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupHeaderIsVisible = 5
    mvarGroupHeaderIsVisible = vData
End Property


Friend Property Get GroupHeaderIsVisible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupHeaderIsVisible
    GroupHeaderIsVisible = mvarGroupHeaderIsVisible
End Property



Friend Property Let GroupHeaders(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupHeaders = 5
    mvarGroupHeaders = vData
End Property


Friend Property Get GroupHeaders() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupHeaders
    GroupHeaders = mvarGroupHeaders
End Property



Friend Property Let Sort(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sort = 5
    mvarSort = vData
End Property


Friend Property Get Sort() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sort
    Sort = mvarSort
End Property



Friend Property Let Widths(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Widths = 5
    mvarWidths = vData
End Property


Friend Property Get Widths() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Widths
    Widths = mvarWidths
End Property



Friend Property Let Alignments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Alignments = 5
    mvarAlignments = vData
End Property


Friend Property Get Alignments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Alignments
    Alignments = mvarAlignments
End Property



Public Property Let TView_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TView_ID = 5
    mvarTView_ID = vData
End Property


Public Property Get TView_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TView_ID
    TView_ID = mvarTView_ID
End Property



Friend Property Let NodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NodeID = 5
    mvarNodeID = vData
End Property


Friend Property Get NodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NodeID
    NodeID = mvarNodeID
End Property



Friend Property Let UserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserID = 5
    mvarUserID = vData
End Property


Friend Property Get UserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserID
    UserID = mvarUserID
End Property



Friend Property Let DVCIDs(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DVCIDs = 5
    mvarDVCIDs = vData
End Property


Friend Property Get DVCIDs() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DVCIDs
    DVCIDs = mvarDVCIDs
End Property

Friend Function GridRecords(ByRef Window As Object, _
                            ByVal NodeProps As CNavigationPane, _
                            ByRef ADOConnection As ADODB.Connection, _
                            ByRef GridRecordset As ADODB.Recordset, _
                            ByRef AllowMissingDatabase As Boolean) _
                            As QueryResultConstants
                            
    Dim rstTreeview As ADODB.Recordset
    Dim rstViews As ADODB.Recordset
    Dim rstDVC As ADODB.Recordset

    Dim arrColumns
    Dim arrHiddenColumns
    
    Dim lngColCtr As Long
    Dim lngTimer As Long
    
    Dim strCommandText As String
    Dim strFields As String
    Dim strFilter As String
    Dim strViewSQL As String
    Dim strViewFilter As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrderBy As String
    Dim strGroupBy As String
    
    Dim rstMasterTable As ADODB.Recordset
    Dim strDummyFrom As String
    Dim strFieldCheck As String
    Dim blnFieldExists As Boolean
    
    Dim blnMostLikelyTheDatabaseHasBeenRemoved As Boolean
    
    Dim arrSplit() As String
    Dim lngCtr As Long
    'Allan may26
    Dim rstDummyForRefresh As New ADODB.Recordset
    Dim lngSortOuterLoopCtr As Long
    Dim lngSortOuterLoopCnt As Long
                           
    On Error GoTo Error_Handler
    
    '>> formulate SQl for the selected view
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Nodes "
    strCommandText = strCommandText & "INNER JOIN "
    strCommandText = strCommandText & "TreeViews "
    strCommandText = strCommandText & "ON "
    strCommandText = strCommandText & "Nodes.TView_ID = TreeViews.TView_ID "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Nodes.Node_ID = " & NodeProps.NodeID & " "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstTreeview, adOpenKeyset, adLockOptimistic
    'Set rstTreeview = ADOConnection.Execute(strCommandText)
            
    strDummyFrom = UCase$(rstTreeview!TView_SQLFrom)
    strDummyFrom = KeywordTranslation(strDummyFrom)
    
    
    blnMostLikelyTheDatabaseHasBeenRemoved = False
    
        strCommandText = vbNullString
        strCommandText = strCommandText & " SELECT TOP 1 "
        strCommandText = strCommandText & " * "
        strCommandText = strCommandText & " FROM "
        strCommandText = strCommandText & " " & strDummyFrom & " "
    

    ' CSCLP-583/CSCLP-582
    On Error GoTo 0
    lngTimer = Timer
    Do While True
        Err.Clear
        '>> open grid records
        On Error Resume Next
        
        ADORecordsetOpen strCommandText, ADOConnection, rstMasterTable, adOpenKeyset, adLockOptimistic
        'Set rstMasterTable = ADOConnection.Execute(strCommandText)
        
        If (Timer - lngTimer) > 10 Then
            If Err.Number <> 0 Then
                
                If Err.Number = -2147467259 Then
                    ' Err.Description = '??? is not a valid path. Make sure that the
                    ' path name is spelled correctly and that you are connected
                    ' to the server on which the file resides.
                    On Error GoTo 0
                    blnMostLikelyTheDatabaseHasBeenRemoved = True
                    Exit Do
                ElseIf Err.Number <> -2147217911 Then
                    GoTo Error_Handler
                End If
                Debug.Print ""
            ElseIf Err.Number = 0 Then
                Exit Do
            End If

        ElseIf Err.Number = -2147467259 Then
            ' Err.Description = '??? is not a valid path. Make sure that the
            ' path name is spelled correctly and that you are connected
            ' to the server on which the file resides.
            On Error GoTo 0
            blnMostLikelyTheDatabaseHasBeenRemoved = True
            Exit Do
            
        ElseIf Err.Number = -2147217865 Then
            ' Err.Description = '??? is not a valid Table. This may happen when linked Table no longer exist
            On Error GoTo 0
            blnMostLikelyTheDatabaseHasBeenRemoved = True
            Exit Do
            
        ElseIf Err.Number = 0 Then
            Exit Do
            
        End If
    Loop
    On Error GoTo Error_Handler
    
'''''    On Error Resume Next
'''''    Set rstMasterTable = ADOConnection.Execute(strCommandText)
'''''    If Err.Number = -2147467259 Then
'''''        ' Err.Description = '??? is not a valid path. Make sure that the
'''''        ' path name is spelled correctly and that you are connected
'''''        ' to the server on which the file resides.
'''''        On Error GoTo 0
'''''        blnMostLikelyTheDatabaseHasBeenRemoved = True
'''''    ElseIf Err.Number <> 0 Then
'''''        GoTo Error_Handler
'''''    ElseIf Err.Number = 0 Then
'''''        On Error GoTo 0
'''''    End If
'''''    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "TView_ID = " & Me.TView_ID & " "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = ADOConnection.Execute(strCommandText)
    'Allan may26
    
    ADORecordsetOpen strCommandText, ADOConnection, rstDummyForRefresh, adOpenKeyset, adLockOptimistic
    'Set rstDummyForRefresh = ADOConnection.Execute(strCommandText)
    
    If (rstDVC.EOF And rstDVC.BOF) Then
        GridRecords = QueryResultNoRecord
        ' hobbes 10/18/2005
        Call ADORecordsetClose(rstTreeview)
        Call ADORecordsetClose(rstDVC)
        Exit Function
    End If
    
RepeatProcedure:

    arrColumns = Split(Me.DVCIDs, "*****")
    arrHiddenColumns = Split(Me.RequiredFields, "*****")
    
    strFields = vbNullString
    
    For lngColCtr = 0 To UBound(arrColumns)
        rstDVC.MoveFirst
        Do While Not rstDVC.EOF
            If Val(arrColumns(lngColCtr)) = rstDVC!DVC_ID Then
                Exit Do
            End If
            rstDVC.MoveNext
        Loop
        If Not rstDVC.EOF Then
            blnFieldExists = True

            strFieldCheck = rstDVC!DVC_FieldSource
            If UCase$(strFieldCheck) <> "[TYPE]+[COMM]" Then ' No Checking for Composite Fields
                If Left$(strFieldCheck, 1) = "[" Then
                    strFieldCheck = Right$(strFieldCheck, Len(strFieldCheck) - 1)
                End If
                If Right$(strFieldCheck, 1) = "]" Then
                    strFieldCheck = Left$(strFieldCheck, Len(strFieldCheck) - 1)
                End If
                
                
                ' If No error occurs above, then we could proceed with checking
                ' for the existence of the field in the source table. Otherwise, we
                ' skip this step
                If Not blnMostLikelyTheDatabaseHasBeenRemoved Then
                
                    On Error Resume Next
                    strFieldCheck = rstMasterTable.Fields(strFieldCheck).Name
                    If Err.Number = 3265 Then
                        blnFieldExists = False
                        On Error GoTo 0
                    ElseIf Err.Number = 0 Then
                        blnFieldExists = True
                        On Error GoTo 0
                    Else
                        blnFieldExists = False
                        GoTo Error_Handler
                    End If
                End If
                
                On Error GoTo Error_Handler
            End If

            If blnFieldExists Then
                strFieldCheck = rstDVC!DVC_FieldSource
                If Left$(strFieldCheck, 1) <> "[" Then
                    strFieldCheck = "[" & strFieldCheck
                End If
                If Right$(strFieldCheck, 1) <> "]" Then
                    strFieldCheck = strFieldCheck & "]"
                End If
                If InStr(1, strFields, "[" & rstDVC!DVC_FieldAlias & "]") = 0 Then 'eradicate duplicate aliases
                    strFields = strFields & strFieldCheck & " AS [" & rstDVC!DVC_FieldAlias & "], "
                End If
            End If
        End If
    Next

    For lngColCtr = 0 To UBound(arrHiddenColumns)
        rstDVC.MoveFirst
        Do While Not rstDVC.EOF
            If Val(arrHiddenColumns(lngColCtr)) = rstDVC!DVC_ID Then
                Exit Do
            End If
            rstDVC.MoveNext
        Loop
        If Not rstDVC.EOF Then
            blnFieldExists = True

            strFieldCheck = rstDVC!DVC_FieldSource
            If UCase$(strFieldCheck) <> "[TYPE]+[COMM]" Then ' No Checking for Composite Fields
                If Left$(strFieldCheck, 1) = "[" Then
                    strFieldCheck = Right$(strFieldCheck, Len(strFieldCheck) - 1)
                End If
                If Right$(strFieldCheck, 1) = "]" Then
                    strFieldCheck = Left$(strFieldCheck, Len(strFieldCheck) - 1)
                End If
                
                ' If No error occurs above, then we could proceed with checking
                ' for the existence of the field in the source table. Otherwise, we
                ' skip this step
                If Not blnMostLikelyTheDatabaseHasBeenRemoved Then
                    On Error Resume Next
                    strFieldCheck = rstMasterTable.Fields(strFieldCheck).Name
                    If Err.Number = 3265 Then
                        blnFieldExists = False
                        On Error GoTo 0
                    ElseIf Err.Number = 0 Then
                        blnFieldExists = True
                        On Error GoTo 0
                    Else
                        blnFieldExists = False
                        GoTo Error_Handler
                    End If
                End If

                On Error GoTo Error_Handler
            End If

            If blnFieldExists Then
                strFieldCheck = rstDVC!DVC_FieldSource
                If Left$(strFieldCheck, 1) <> "[" Then
                    strFieldCheck = "[" & strFieldCheck
                End If
                If Right$(strFieldCheck, 1) <> "]" Then
                    strFieldCheck = strFieldCheck & "]"
                End If
                If InStr(1, strFields, rstDVC!DVC_FieldSource) = 0 Then 'eradicate duplicate aliases
                    strFields = strFields & rstDVC!DVC_FieldSource & " AS [" & rstDVC!DVC_FieldAlias & "], "
                End If
            End If
        End If
    Next
    
    If Trim(strFields) = "" Then
        LoadDefaultSettings NodeProps, ADOConnection
        GoTo RepeatProcedure
    Else
        strFields = Mid(strFields, 1, Len(strFields) - 2)
    End If
    
    '>> get filters defined by users
    strFilter = ViewFilters(ADOConnection)
    
    g_typInterface.IGridHeader.Caption = Trim(Replace(g_typInterface.IGridHeader.Caption, " (Filter Applied)", ""))
    
    If Trim(strFilter) <> "" Then
        g_typInterface.IGridHeader.Caption = g_typInterface.IGridHeader.Caption & " (Filter Applied)"
    End If
    
    '>> formulate view SQL
    strViewSQL = vbNullString
    strViewSQL = strViewSQL & "SELECT "
    strViewSQL = strViewSQL & "View_Filter "
    strViewSQL = strViewSQL & "FROM "
    strViewSQL = strViewSQL & "Views "
    strViewSQL = strViewSQL & "INNER JOIN "
    strViewSQL = strViewSQL & "UserViewColumns "
    strViewSQL = strViewSQL & "ON "
    strViewSQL = strViewSQL & "Views.View_ID = UserViewColumns.View_ID "
    strViewSQL = strViewSQL & "WHERE "
    strViewSQL = strViewSQL & "Views.Node_ID = " & NodeProps.NodeID & " "
    strViewSQL = strViewSQL & "AND "
    strViewSQL = strViewSQL & "UserViewColumns.User_ID = " & NodeProps.UserID & " "
    strViewSQL = strViewSQL & "AND "
    strViewSQL = strViewSQL & "UVC_SelectedView = True "
    
    ADORecordsetOpen strViewSQL, ADOConnection, rstViews, adOpenKeyset, adLockOptimistic
    'Set rstViews = ADOConnection.Execute(strViewSQL)
    
    '>> get view filter
    If Not (rstViews.EOF And rstViews.BOF) Then
        rstViews.MoveFirst
        
        strViewFilter = KeywordTranslation(FNullField(rstViews!View_Filter))
    End If
    
    '>> formulate grid SQL
    strFields = "SELECT " & strFields
    strFrom = " FROM " & rstTreeview!TView_SQLFrom
    
    If Trim(rstTreeview!TView_SQLWhere) <> "" Then
        
        strWhere = " WHERE (" & rstTreeview!TView_SQLWhere & ") "
        
        If Trim(strViewFilter) <> "" Then
            strWhere = strWhere & " AND (" & strViewFilter & ") "
        End If
        
        If Trim(strFilter) <> "" Then
            strWhere = strWhere & " AND " & strFilter
        End If
        
    ElseIf Trim(strViewFilter) <> "" Then
        
        strWhere = " WHERE (" & strViewFilter & ") "
        
        If Trim(strFilter) <> "" Then
            strWhere = strWhere & " AND " & strFilter
        End If
        
    ElseIf Trim(strFilter) <> "" Then
    
        strWhere = " WHERE " & strFilter
    
    End If
    
    '***********************************************************************************
    'Sorting - Edwin Nov09
    '***********************************************************************************
    If Trim(Me.Sort) <> "" Then
        
        '*********************** Added to put group by boxes in ORDER BY clause***********************************
        'CSCLP-214
        If Trim(Me.GroupHeaders) <> vbNullString And Not IsNull(Me.GroupHeaders) Then
            lngSortOuterLoopCnt = 1     'If group fields found
        Else
            lngSortOuterLoopCnt = 2     'If no group fields found
        End If
'        lngSortOuterLoopCnt = 1
'        If Trim(Me.GroupHeaders) <> vbNullString And Not IsNull(Me.GroupHeaders) Then
'            lngSortOuterLoopCnt = 2
'        End If
        '***********************************************************************************************************
        
        strOrderBy = vbNullString
        
        '*********************** Added to put group by boxes in ORDER BY clause***********************************
        'CSCLP-214
        'If group fields found performs 1 and 2, while no group fields performs only 2.
        For lngSortOuterLoopCtr = lngSortOuterLoopCnt To 2
'        For lngSortOuterLoopCtr = IIf(lngSortOuterLoopCnt = 2, 1, 2) To lngSortOuterLoopCnt
            Select Case lngSortOuterLoopCtr
                ' Must process Group Headers into Sort Clause First Before the Actual Sort Fields
                Case 1
                    Dim arrSplitGroupHeader() As String
                    Dim lngGroupCtr As Long
                    
                    'Erase arrSplit()
                    ReDim arrSplit(0)
                    lngGroupCtr = 0
                                        
                    arrSplitGroupHeader = Split(Me.GroupHeaders, "*****")
                    
                    For lngGroupCtr = LBound(arrSplitGroupHeader) To UBound(arrSplitGroupHeader) Step 3
                        rstDVC.Filter = adFilterNone
                        rstDVC.Filter = "[DVC_ID] = " & arrSplitGroupHeader(lngGroupCtr)
                        If Not (rstDVC.EOF And rstDVC.BOF) Then
                            
                            ReDim Preserve arrSplit(((lngGroupCtr \ 3) * 2) + 1)
                            
                            ' Alias
                            arrSplit(((lngGroupCtr \ 3) * 2)) = rstDVC!DVC_FieldAlias
                            
                            ' Sort Value
                            arrSplit(((lngGroupCtr \ 3) * 2) + 1) = arrSplitGroupHeader(lngGroupCtr + 1)
                        End If
                        rstDVC.Filter = adFilterNone
                    Next

                Case 2
                    arrSplit = Split(Me.Sort, "*****")
            End Select
        '***********************************************************************************************************
            'added 071108
            If UBound(arrSplit) > 0 Then
            '>>
                For lngCtr = LBound(arrSplit) To UBound(arrSplit) Step 2
                    If lngCtr Mod 2 = 0 Then
                        rstDVC.MoveFirst
                        Do While Not rstDVC.EOF
                            If arrSplit(lngCtr) = rstDVC!DVC_FieldAlias Then
                                Exit Do
                            End If
                            rstDVC.MoveNext
                        Loop
                        
                        If Not rstDVC.EOF Then
                            blnFieldExists = True
        
                            strFieldCheck = rstDVC!DVC_FieldSource
                            If UCase$(strFieldCheck) <> "[TYPE]+[COMM]" Then ' No Checking for Composite Fields
                                If Left$(strFieldCheck, 1) = "[" Then
                                    strFieldCheck = Right$(strFieldCheck, Len(strFieldCheck) - 1)
                                End If
                                If Right$(strFieldCheck, 1) = "]" Then
                                    strFieldCheck = Left$(strFieldCheck, Len(strFieldCheck) - 1)
                                End If
                                
                                ' If No error occurs above, then we could proceed with checking
                                ' for the existence of the field in the source table. Otherwise, we
                                ' skip this step
                                If Not blnMostLikelyTheDatabaseHasBeenRemoved Then
                                    On Error Resume Next
                                    strFieldCheck = rstMasterTable.Fields(strFieldCheck).Name
                                    If Err.Number = 3265 Then
                                        blnFieldExists = False
                                        On Error GoTo 0
                                    ElseIf Err.Number = 0 Then
                                        blnFieldExists = True
                                        On Error GoTo 0
                                    Else
                                        blnFieldExists = False
                                        GoTo Error_Handler
                                    End If
                                End If
                
                                On Error GoTo Error_Handler
                            End If
                
                            If blnFieldExists Then
                                ' jgexSortAscending = 1
                                ' jgexSortNone = 0
                                ' jgexSortDescending = -1
                                
                                ' Do not add anymore if already existing
                                If InStr(1, strOrderBy, rstDVC!DVC_FieldSource & " ") <= 0 Then
                                    strOrderBy = strOrderBy & " " & rstDVC!DVC_FieldSource & " " & IIf(arrSplit(lngCtr + 1) = 1, "ASC", "DESC") & ", "
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
        
                        
        strOrderBy = Trim(strOrderBy)
        If LenB(strOrderBy) > 0 Then
            strOrderBy = Left(strOrderBy, Len(strOrderBy) - 1)
            strOrderBy = " ORDER BY " & strOrderBy & Space(1)
        End If
    Else
        If Trim(rstTreeview!TView_SQLOrderBy) <> "" Then
            strOrderBy = " ORDER BY " & rstTreeview!TView_SQLOrderBy
        End If
    End If
    '***********************************************************************************
    
'    If Trim(rstTreeview!TView_SQLOrderBy) <> "" Then
'        strOrderBy = " ORDER BY " & rstTreeview!TView_SQLOrderBy
'    End If
    
    If Trim(rstTreeview!TView_SQLGroupBy) <> "" Then
        strGroupBy = " GROUP BY " & rstTreeview!TView_SQLGroupBy
    End If
    
    
    RaiseEvent BeforeOpenRecordset(strFields, strFrom, strWhere, strGroupBy, strOrderBy, AllowMissingDatabase)
    
    strCommandText = strFields & strFrom & strWhere & strGroupBy & strOrderBy
    strCommandText = KeywordTranslation(strCommandText)
    
    '>> Loop added to prevent errors caused by the speed of accessing linked tables
    '>> Loop will stop if there are no errors or after 5 seconds
    lngTimer = Timer
    
    Do While True
        Err.Clear
        '>> open grid records
        On Error Resume Next
        ADORecordsetOpen strCommandText, ADOConnection, GridRecordset, adOpenKeyset, adLockOptimistic
        'GridRecordset.Open strCommandText, ADOConnection, adOpenKeyset, adLockReadOnly
'        If (Timer - lngTimer) > 5 And Err.Number = -2147217911 Then
'            GoTo Error_Handler
'        ElseIf Err.Number <> 0 And Err.Number <> -2147217911 Then
'            GoTo Error_Handler
'        ElseIf Err.Number = 0 Then
'            Exit Do
'        End If
        If (Timer - lngTimer) > 10 Then
            If Err.Number <> 0 Then
               
                If Err.Number = -2147217911 Then
                    GoTo Error_Handler
                ElseIf Err.Number <> -2147217911 Then
                   GoTo Error_Handler
                End If
                Debug.Print ""
         
            ElseIf Err.Number <> 0 And Err.Number <> -2147217911 Then
                GoTo Error_Handler
            ElseIf Err.Number = 0 Then
                Exit Do
            End If
 
        ElseIf Err.Number = 0 Then
            Exit Do
        End If
    Loop
            
    GridRecords = QueryResultSuccessful
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    Call ADORecordsetClose(rstTreeview)
    Call ADORecordsetClose(rstViews)
    
    Exit Function
    
Error_Handler:
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    Call ADORecordsetClose(rstTreeview)
    Call ADORecordsetClose(rstViews)
    
    GridRecords = QueryResultError
    
End Function

Private Function ViewFilters(ByRef ADOConnection As ADODB.Connection) As String

    Dim rstFilter As ADODB.Recordset
    Dim lngFilterType As Long
    Dim strCommandText As String
    Dim strFilter1 As String
    Dim strFilter2 As String
    Dim strLastField As String
    Dim strFilterField As String
    Dim arrDateBetween
    Dim strViewFilters As String
    Dim arrFilter2() As String
    Dim lngFilter2Ctr As Long
            
    ReDim arrFilter2(0)
    
    lngFilter2Ctr = 0
    
    '>> open filters defined for the selected view
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Filter "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & Me.UVC_ID & " "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "Filter_Type, Filter_Field "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstFilter, adOpenKeyset, adLockOptimistic
    'RstOpen strCommandText, ADOConnection, rstFilter, adOpenForwardOnly, adLockOptimistic, , True
    
    'Set rstFilter = ADOConnection.Execute(strCommandText)
    
    '>> strFilter1-filter defined in the contents tab
    '>> strFilter2-filter defined in the advanced tab
    
    strFilter2 = vbNullString
    strLastField = vbNullString

    If rstFilter.RecordCount > 0 Then
        rstFilter.MoveFirst
        
        ' Initialize for first record
        strLastField = FNullField(rstFilter!Filter_Field)
        
        Do Until rstFilter.EOF
            strFilterField = FNullField(rstFilter!Filter_Field)
            If Left$(strFilterField, 1) = "[" And Right$(strFilterField, 1) = "]" Then
                strFilterField = Right$(strFilterField, Len(strFilterField) - 1)
                strFilterField = Left$(strFilterField, Len(strFilterField) - 1)
            End If
            
            Select Case FNullField(rstFilter!Filter_Operator)
                Case 1 '>> contains/on
                    Select Case rstFilter!Filter_DataType
                        Case 3 '>> Date
                            If FNullField(rstFilter!Filter_Type) = 1 Then
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter1 = strFilter1 & "( "
                                
                                    strFilter1 = strFilter1 & "DATEPART('m', [" & strFilterField & "]) = " & Month(rstFilter!Filter_Value) & " "
                                    strFilter1 = strFilter1 & "AND "
                                    strFilter1 = strFilter1 & "DATEPART('d', [" & strFilterField & "]) = " & Day(rstFilter!Filter_Value) & " "
                                    strFilter1 = strFilter1 & "AND "
                                    strFilter1 = strFilter1 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(rstFilter!Filter_Value) & " "
                                
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter1 = strFilter1 & ") "
                                
                            Else
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter2 = strFilter2 & "( "
                                
                                    strFilter2 = strFilter2 & "DATEPART('m', [" & strFilterField & "]) = " & Month(rstFilter!Filter_Value) & " "
                                    strFilter2 = strFilter2 & "AND "
                                    strFilter2 = strFilter2 & "DATEPART('d', [" & strFilterField & "]) = " & Day(rstFilter!Filter_Value) & " "
                                    strFilter2 = strFilter2 & "AND "
                                    strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(rstFilter!Filter_Value) & " "
                                
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter2 = strFilter2 & ") "
                            End If
                            
                        Case Else
                            If FNullField(rstFilter!Filter_Type) = 1 Then
                                strFilter1 = strFilter1 & "( "
                                
                                    ' we can have only 1 filter of filter-type-1
                                    strFilter1 = strFilter1 & "[" & strFilterField & "] Like '%" & AQ(rstFilter!Filter_Value) & "%' "
                                
                                strFilter1 = strFilter1 & ") "
                            Else
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter2 = strFilter2 & "( "
                                
                                    strFilter2 = strFilter2 & "[" & strFilterField & "] Like '%" & AQ(rstFilter!Filter_Value) & "%'"
                                    
                                ' Needed for Multiple Filters of Filter-Type-2
                                strFilter2 = strFilter2 & ") "
                            End If
                    End Select
                
                Case 2 '>> is exactly/is equal/on or before
                    Select Case rstFilter!Filter_DataType
                        Case 1 '>> Text
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] = '" & AQ(rstFilter!Filter_Value) & "'"
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 2 '>> Numeric
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] = " & CDbl(rstFilter!Filter_Value)
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            

                        Case 3 '>> Date
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "( "
                                strFilter2 = strFilter2 & "DATEPART('m', [" & strFilterField & "]) = " & Month(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('d', [" & strFilterField & "]) = " & Day(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & ") "
                                strFilter2 = strFilter2 & "OR "
                                strFilter2 = strFilter2 & "( "
                                    strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('m', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('d', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & ") "
                                            strFilter2 = strFilter2 & " < "
                                        strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & Format(Year(rstFilter!Filter_Value), "####") & ""
                                        strFilter2 = strFilter2 & Format(Month(rstFilter!Filter_Value), "0#") & ""
                                        strFilter2 = strFilter2 & Format(Day(rstFilter!Filter_Value), "0#") & ""
                                        strFilter2 = strFilter2 & ") "
                                    strFilter2 = strFilter2 & ") "
                                strFilter2 = strFilter2 & ") "
                                                        
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 4 '>> Boolean
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] = " & CBool(rstFilter!Filter_Value)
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                    End Select
                    
                Case 3 '>> doesn't contain/less than or equal/on or after
                    Select Case rstFilter!Filter_DataType
                        Case 1 '>> Text
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] Not Like '%" & AQ(rstFilter!Filter_Value) & "%'"
                                
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 2 '>> Numeric
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] <= " & CDbl(rstFilter!Filter_Value)
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "

                        Case 3 '>> Date
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "( "
                                strFilter2 = strFilter2 & "DATEPART('m', [" & strFilterField & "]) = " & Month(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('d', [" & strFilterField & "]) = " & Day(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(rstFilter!Filter_Value) & " "
                                strFilter2 = strFilter2 & ") "
                                strFilter2 = strFilter2 & "OR "
                                strFilter2 = strFilter2 & "( "
                                    strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('m', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('d', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & ") "
                                            strFilter2 = strFilter2 & " > "
                                        strFilter2 = strFilter2 & "("
                                        strFilter2 = strFilter2 & Format(Year(rstFilter!Filter_Value), "####") & ""
                                        strFilter2 = strFilter2 & Format(Month(rstFilter!Filter_Value), "0#") & ""
                                        strFilter2 = strFilter2 & Format(Day(rstFilter!Filter_Value), "0#") & ""
                                        strFilter2 = strFilter2 & ") "
                                    strFilter2 = strFilter2 & ") "
                                strFilter2 = strFilter2 & ") "
                                
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 4 '>> Boolean
                    
                    End Select
                            
                Case 4 '>> is empty/greater than or equal/between
                    Select Case rstFilter!Filter_DataType
                        Case 1 '>> Text
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "IsNull([" & strFilterField & "]) = True "
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 2 '>> Numeric
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] >= " & CDbl(rstFilter!Filter_Value)
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                            
                        Case 3 '>> Date
                            arrDateBetween = Split(Trim$(UCase$(rstFilter!Filter_Value)), UCase$(Trim$(" AND ")))
    
                            arrDateBetween(0) = Replace(CStr(arrDateBetween(0)), "<", "")
                            arrDateBetween(0) = Replace(CStr(arrDateBetween(0)), ">", "")
    
                            arrDateBetween(1) = Replace(CStr(arrDateBetween(1)), "<", "")
                            arrDateBetween(1) = Replace(CStr(arrDateBetween(1)), ">", "")
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "( "
                                strFilter2 = strFilter2 & "DATEPART('m', [" & strFilterField & "]) = " & Month(CDate(arrDateBetween(0))) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('d', [" & strFilterField & "]) = " & Day(CDate(arrDateBetween(0))) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(CDate(arrDateBetween(0))) & " "
                                strFilter2 = strFilter2 & ") "
        
                                strFilter2 = strFilter2 & "OR "
        
                                strFilter2 = strFilter2 & "( "
                                strFilter2 = strFilter2 & "DATEPART('m', [" & strFilterField & "]) = " & Month(CDate(arrDateBetween(1))) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('d', [" & strFilterField & "]) = " & Day(CDate(arrDateBetween(1))) & " "
                                strFilter2 = strFilter2 & "AND "
                                strFilter2 = strFilter2 & "DATEPART('yyyy', [" & strFilterField & "]) = " & Year(CDate(arrDateBetween(1))) & " "
                                strFilter2 = strFilter2 & ") "
        
                                strFilter2 = strFilter2 & "OR "
                                
                                strFilter2 = strFilter2 & "( "
                                
                                        strFilter2 = strFilter2 & "( "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('yyyy', [" & strFilterField & "]), '####') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('m', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('d', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & ") "
                                            strFilter2 = strFilter2 & " > "
                                        strFilter2 = strFilter2 & "( "
                                        strFilter2 = strFilter2 & Format(Year(CDate(arrDateBetween(0))), "####") & ""
                                        strFilter2 = strFilter2 & Format(Month(CDate(arrDateBetween(0))), "0#") & ""
                                        strFilter2 = strFilter2 & Format(Day(CDate(arrDateBetween(0))), "0#") & ""
                                        strFilter2 = strFilter2 & ") "
                                                                                
                                        strFilter2 = strFilter2 & "AND "
                                        
                                        strFilter2 = strFilter2 & "( "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('yyyy', [" & strFilterField & "]), '####') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('m', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & "& "
                                        strFilter2 = strFilter2 & "FORMAT(DATEPART('d', [" & strFilterField & "]), '0#') "
                                        strFilter2 = strFilter2 & ") "
                                            strFilter2 = strFilter2 & " < "
                                        strFilter2 = strFilter2 & "( "
                                        strFilter2 = strFilter2 & Format(Year(CDate(arrDateBetween(1))), "####") & ""
                                        strFilter2 = strFilter2 & Format(Month(CDate(arrDateBetween(1))), "0#") & ""
                                        strFilter2 = strFilter2 & Format(Day(CDate(arrDateBetween(1))), "0#") & ""
                                        strFilter2 = strFilter2 & ") "
                                
                                strFilter2 = strFilter2 & ") "
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
    
                        Case 4 '>> Boolean
                    
                    End Select
    
                Case 5 '>> is not empty
                    Select Case rstFilter!Filter_DataType
                        Case 1 '>> Text
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "IsNull([" & strFilterField & "]) = False "
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                        
                        Case 2 '>> Numeric
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & "( "
                            
                                strFilter2 = strFilter2 & "[" & strFilterField & "] Between " & _
                                                    Replace(Replace(rstFilter!Filter_Value, "<", ""), ">", "")
                            
                            ' Needed for Multiple Filters of Filter-Type-2
                            strFilter2 = strFilter2 & ") "
                                        
                        Case 3 '>> Date
                            
                        
                        Case 4 '>> Boolean
                    
                    End Select
                
            End Select
            
            strLastField = FNullField(rstFilter!Filter_Field)
            lngFilterType = FNullField(rstFilter!Filter_Type)
            
            rstFilter.MoveNext
            
            ' Look ahead for the neame of the filter field
            If rstFilter.EOF Then
                strFilterField = ""
            Else
                strFilterField = FNullField(rstFilter!Filter_Field)
            End If
            
            If Trim(strFilter1) <> "" And lngFilterType = 1 Then
                strFilter1 = strFilter1 & " OR "
            End If
            
            If Trim(strFilter2) <> "" And lngFilterType = 2 Then
                If strFilterField = strLastField Then
                    strFilter2 = strFilter2 & " OR "
                Else
                    ' Start Value of lngFilter2Ctr is 1
                    lngFilter2Ctr = lngFilter2Ctr + 1
                    If UBound(arrFilter2) > 0 Then
                        ReDim Preserve arrFilter2(1 To lngFilter2Ctr)
                    Else
                        ReDim arrFilter2(1 To lngFilter2Ctr)
                    End If
                    
                    strFilter2 = Trim$(strFilter2)
                    strFilter2 = "( " & strFilter2 & ") "
                    
                    arrFilter2(lngFilter2Ctr) = strFilter2
                    
                    ' Re-initialize strFilter2 for next Filter if any
                    strFilter2 = vbNullString
                End If
            End If
        Loop
    End If
    
    'Set rstFilter = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstFilter)
                        
    
    ' Remove ' AND '
    strFilter1 = RTrim$(strFilter1)
    If Right$(strFilter1, 3) = " OR" Then
        strFilter1 = Left$(strFilter1, Len(strFilter1) - 3)
    End If
    If LenB(Trim$(strFilter1)) > 0 Then
        If UBound(arrFilter2) > 0 Then
            ReDim Preserve arrFilter2(1 To UBound(arrFilter2) + 1)
        Else
            ReDim arrFilter2(1 To UBound(arrFilter2) + 1)
        End If
        
        strFilter1 = Trim$(strFilter1)
        strFilter1 = "( " & strFilter1 & ") "
        
        arrFilter2(UBound(arrFilter2)) = strFilter1
    End If
    
    ' Reset strFilter2 and rebuild from array
    strFilter2 = vbNullString
    If UBound(arrFilter2) >= 1 Then
        For lngFilter2Ctr = 1 To UBound(arrFilter2)
            If UBound(arrFilter2) > 1 Then
                strFilter2 = strFilter2 & "( " & arrFilter2(lngFilter2Ctr) & ") " & " AND "
            Else
                strFilter2 = arrFilter2(lngFilter2Ctr)
            End If
        Next
    End If
  
    ' Remove ' AND '
    strFilter2 = RTrim$(strFilter2)
    If Right$(strFilter2, 4) = " AND" Then
        strFilter2 = Left$(strFilter2, Len(strFilter2) - 4)
    End If
                    
    If LenB(strFilter2) > 0 Then
        strViewFilters = "( " & strFilter2 & ") "
    End If
    
    strViewFilters = Replace(Replace(strViewFilters, "[[", "["), "]]", "]")
    
    strViewFilters = " " & Trim$(strViewFilters) & " "
        
    ViewFilters = strViewFilters
End Function

Private Function KeywordTranslation(ByVal FilterKey As String) As String
    
    Dim lngCharPos As Long
    Dim lngKeyEnd As Long
    Dim strKeyword As String
    Dim varKeyValue As Variant
    
    lngCharPos = 1
    
    Do While lngCharPos <= Len(Trim(FilterKey))
        If Mid(FilterKey, lngCharPos, 3) = "<<<" Then
            lngKeyEnd = InStr(lngCharPos, FilterKey, ">>>")
            strKeyword = Mid(FilterKey, lngCharPos, lngKeyEnd - lngCharPos + 3)
            
            RaiseEvent KeywordValue(strKeyword, varKeyValue)
            
            FilterKey = Replace(FilterKey, strKeyword, varKeyValue)
            lngCharPos = 1
        Else
            lngCharPos = lngCharPos + 1
        End If
    Loop
    
    KeywordTranslation = FilterKey
    
End Function

Friend Function SelectGridSetting(ByRef NodeProps As CNavigationPane, _
                                    ByRef ADOConnection As ADODB.Connection, _
                                    Optional ValidationOnly As Boolean = False) As QueryResultConstants

    Dim rstView As ADODB.Recordset
    
    Dim strCommandText As String
    Dim blnHasSetting As Boolean
    
    On Error GoTo Error_Handler
        
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "(Nodes "
    strCommandText = strCommandText & "INNER JOIN "
    strCommandText = strCommandText & "Views "
    strCommandText = strCommandText & "ON "
    strCommandText = strCommandText & "Nodes.Node_ID = Views.Node_ID) "
    strCommandText = strCommandText & "INNER JOIN "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "ON "
    strCommandText = strCommandText & "Views.View_ID = UserViewColumns.View_ID "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "[UserViewColumns.User_ID] = " & NodeProps.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "[Nodes.Node_ID] = " & NodeProps.NodeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "[UserViewColumns.UVC_SelectedView] = True "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstView, adOpenKeyset, adLockOptimistic
    'Set rstView = ADOConnection.Execute(strCommandText)
    
    If Not (rstView.EOF And rstView.BOF) Then
        rstView.MoveFirst
        
        If ValidationOnly = False Then
            Me.ViewID = FNullField(rstView![Views.View_ID])
            Me.UserID = FNullField(rstView![User_ID])
            Me.NodeID = FNullField(rstView![Nodes.Node_ID])
            Me.TView_ID = FNullField(rstView![Nodes.TView_ID])
            Me.UVC_ID = FNullField(rstView!UVC_ID)
            Me.DVCIDs = FNullField(rstView!UVC_DVCIDs)
            Me.Alignments = FNullField(rstView!UVC_ColumnAlignments)
            Me.Widths = FNullField(rstView!UVC_ColumnWidths)
            Me.GroupHeaders = FNullField(rstView!UVC_GroupHeaders)
            Me.CardView = FNullField(rstView!UVC_CardsView)
            Me.GridLines = FNullField(rstView!UVC_GridLines)
            Me.GroupHeaderIsVisible = FNullField(rstView!UVC_GroupHeaderVisible)
            Me.OddEvenColor = FNullField(rstView!UVC_OddEvenColor)
            Me.RequiredFields = FNullField(rstView!UVC_RequirementIs1)
            Me.ColumnFormats = FNullField(rstView!UVC_ColumnFormat)
            Me.Sort = FNullField(rstView!UVC_Sort)
            Me.GroupRowFormat = FNullField(rstView!UVC_GroupRowFormat)
            Me.SelectedRows = FNullField(rstView!UVC_SelectedItem)
            Me.AutoGroup = FNullField(rstView!UVC_AutoGroup)
            Me.ExpandCollapseDefault = FNullField(rstView!UVC_ExpandCollapseDefault)
            Me.ReadingPane = FNullField(rstView!UVC_ReadingPane)
        End If
        SelectGridSetting = QueryResultSuccessful
    Else
        SelectGridSetting = QueryResultNoRecord
    End If
    
    'Set rstView = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstView)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstView)
    SelectGridSetting = QueryResultError
    
End Function

Friend Function AddGridSetting(ByRef NodeProps As CNavigationPane, ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    Dim rstUVC As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "( "
    strCommandText = strCommandText & "View_ID, "
    strCommandText = strCommandText & "User_ID, "
    strCommandText = strCommandText & "Node_ID, "
    strCommandText = strCommandText & "TView_ID, "
    strCommandText = strCommandText & "UVC_DVCIDs, "
    strCommandText = strCommandText & "UVC_ColumnAlignments, "
    strCommandText = strCommandText & "UVC_ColumnWidths, "
    strCommandText = strCommandText & "UVC_GroupHeaders, "
    strCommandText = strCommandText & "UVC_CardsView, "
    strCommandText = strCommandText & "UVC_GridLines, "
    strCommandText = strCommandText & "UVC_GroupHeaderVisible, "
    strCommandText = strCommandText & "UVC_OddEvenColor, "
    strCommandText = strCommandText & "UVC_RequirementIs1, "
    strCommandText = strCommandText & "UVC_Sort, "
    strCommandText = strCommandText & "UVC_GroupRowFormat, "
    strCommandText = strCommandText & "UVC_SelectedItem, "
    strCommandText = strCommandText & "UVC_AutoGroup, "
    strCommandText = strCommandText & "UVC_ExpandCollapseDefault, "
    strCommandText = strCommandText & "UVC_ReadingPane, "
    strCommandText = strCommandText & "UVC_SelectedView "
    strCommandText = strCommandText & ") "
    strCommandText = strCommandText & "VALUES "
    strCommandText = strCommandText & "( "
    strCommandText = strCommandText & Me.ViewID & ", "
    strCommandText = strCommandText & Me.UserID & ", "
    strCommandText = strCommandText & Me.NodeID & ", "
    strCommandText = strCommandText & Me.TView_ID & ", "
    strCommandText = strCommandText & "'" & Me.DVCIDs & "', "
    strCommandText = strCommandText & "'" & Me.Alignments & "', "
    strCommandText = strCommandText & "'" & Me.Widths & "', "
    strCommandText = strCommandText & "'" & Me.GroupHeaders & "', "
    strCommandText = strCommandText & IIf(Me.CardView, "True", "False") & ", "
    strCommandText = strCommandText & IIf(Me.GridLines, "True", "False") & ", "
    strCommandText = strCommandText & IIf(Me.GroupHeaderIsVisible, "True", "False") & ", "
    strCommandText = strCommandText & IIf(Me.OddEvenColor, "True", "False") & ", "
    strCommandText = strCommandText & "'" & Me.RequiredFields & "', "
    strCommandText = strCommandText & "'" & Me.Sort & "', "
    strCommandText = strCommandText & "'" & Me.GroupRowFormat & "', "
    strCommandText = strCommandText & "'" & Me.SelectedRows & "', "
    strCommandText = strCommandText & IIf(Me.AutoGroup, "True", "False") & ", "
    strCommandText = strCommandText & Me.ExpandCollapseDefault & ", "
    strCommandText = strCommandText & IIf(Me.ReadingPane, "True", "False") & ", "
    strCommandText = strCommandText & -1
    strCommandText = strCommandText & ") "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
        
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "UVC_ID "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "User_ID = " & Me.UserID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "View_ID = " & Me.ViewID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstUVC, adOpenKeyset, adLockOptimistic
    'Set rstUVC = ADOConnection.Execute(strCommandText)
    
    If Not (rstUVC.EOF And rstUVC.BOF) Then
        rstUVC.MoveFirst
        
        Me.UVC_ID = rstUVC!UVC_ID
    End If
    
    'Set rstUVC = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstUVC)
    
    AddGridSetting = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstUVC)
    AddGridSetting = QueryResultError
    
End Function

Private Function AddDefaultUVCFC(ByVal UVC_ID As Long, ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "FC_Name, "
    strCommandText = strCommandText & "FC_Field, "
    strCommandText = strCommandText & "FC_Operator, "
    strCommandText = strCommandText & "FC_Value1, "
    strCommandText = strCommandText & "FC_Value2, "
    strCommandText = strCommandText & "FC_FontBold, "
    strCommandText = strCommandText & "FC_FontItalic, "
    strCommandText = strCommandText & "FC_FontStrikeThru, "
    strCommandText = strCommandText & "FC_FontUnderline, "
    strCommandText = strCommandText & "FC_ForeColor, "
    strCommandText = strCommandText & "FC_Text, "
    strCommandText = strCommandText & "FC_HasValueList, "
    strCommandText = strCommandText & "FC_ColumnType, "
    strCommandText = strCommandText & "FC_ColumnText, "
    strCommandText = strCommandText & "FC_Icon, "
    strCommandText = strCommandText & "FC_Selected, "
    strCommandText = strCommandText & "FC_Remarks, "
    strCommandText = strCommandText & "FC_Priority, "
    strCommandText = strCommandText & "Node_ID "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "UVC_ID = 0 "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "FC_ID "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "UVC_ID = " & UVC_ID & ", "
    strCommandText = strCommandText & "FC_Default = True "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "(UVC_ID = 0 "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & "IsNull(UVC_ID) = True) "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "FC_Default = False "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
                
    AddDefaultUVCFC = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    AddDefaultUVCFC = QueryResultError
    

End Function
Friend Function UpdateGridSetting(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
  
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "UVC_DVCIDs = '" & Me.DVCIDs & "',"
    strCommandText = strCommandText & "UVC_ColumnAlignments = '" & Me.Alignments & "', "
    strCommandText = strCommandText & "UVC_ColumnWidths = '" & Me.Widths & "', "
    strCommandText = strCommandText & "UVC_GroupHeaders = '" & Me.GroupHeaders & "', "
    strCommandText = strCommandText & "UVC_CardsView = " & IIf(Me.CardView, "True", "False") & ", "
    strCommandText = strCommandText & "UVC_GridLines = " & IIf(Me.GridLines, "True", "False") & ", "
    strCommandText = strCommandText & "UVC_GroupHeaderVisible = " & IIf(Me.GroupHeaderIsVisible, "True", "False") & ", "
    strCommandText = strCommandText & "UVC_OddEvenColor = " & IIf(Me.OddEvenColor, "True", "False") & ", "
    strCommandText = strCommandText & "UVC_RequirementIs1 = '" & Me.RequiredFields & "', "
    strCommandText = strCommandText & "UVC_ColumnFormat = '" & Me.ColumnFormats & "', "
    strCommandText = strCommandText & "UVC_Sort = '" & Me.Sort & "', "
    strCommandText = strCommandText & "UVC_GroupRowFormat = '" & Me.GroupRowFormat & "', "
    strCommandText = strCommandText & "UVC_SelectedItem = '" & Me.SelectedRows & "', "
    strCommandText = strCommandText & "UVC_AutoGroup = " & IIf(Me.AutoGroup, "True", "False") & ", "
    strCommandText = strCommandText & "UVC_ExpandCollapseDefault = " & CStr(Me.ExpandCollapseDefault) & ", "
    strCommandText = strCommandText & "UVC_ReadingPane = " & IIf(Me.ReadingPane, "True", "False") & " "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & Me.UVC_ID & " "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    
    UpdateGridSetting = QueryResultSuccessful
        
    Exit Function
    
Error_Handler:
    
    UpdateGridSetting = QueryResultError




End Function

'Friend Function GetDefaultColumns(ByRef NodeProps As CNavigationPane, ByRef ADOConnection As ADODB.Connection, _
'                                ByRef ColumnRecords As ADODB.Recordset) As QueryResultConstants
'
'
'    Dim rstDVC As ADODB.Recordset
'    Dim strCommandText As String
'
'    On Error GoTo Error_Handler
'
'    strCommandText = vbNullString
'    strCommandText = strCommandText & "SELECT "
'    strCommandText = strCommandText & "* "
'    strCommandText = strCommandText & "FROM "
'    strCommandText = strCommandText & "Nodes "
'    strCommandText = strCommandText & "INNER JOIN "
'    strCommandText = strCommandText & "DefaultViewColumns "
'    strCommandText = strCommandText & "ON "
'    strCommandText = strCommandText & "Nodes.TView_ID = DefaultViewColumns.TView_ID "
'    strCommandText = strCommandText & "WHERE "
'    strCommandText = strCommandText & "Node_ID = " & NodeProps.NodeID & " "
'    strCommandText = strCommandText & "ORDER BY "
'    strCommandText = strCommandText & "DVC_Position"
'
'    Set ColumnRecords = ADOConnection.Execute(strCommandText)
'
'    If ColumnRecords.EOF = False Then
'        GetDefaultColumns = QueryResultSuccessful
'    Else
'        GetDefaultColumns = QueryResultNoRecord
'    End If
'
'    Exit Function
'
'Error_Handler:
'    GetDefaultColumns = QueryResultError
'
'End Function

Friend Function LoadDefaultSettings(ByRef NodeProps As CNavigationPane, _
                ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim rstDVC As ADODB.Recordset
    Dim strCommandText As String

    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Nodes "
    strCommandText = strCommandText & "INNER JOIN "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "ON "
    strCommandText = strCommandText & "Nodes.TView_ID = DefaultViewColumns.TView_ID "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Node_ID = " & NodeProps.NodeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "DVC_Default = True "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "DVC_Position"

    ADORecordsetOpen strCommandText, ADOConnection, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = ADOConnection.Execute(strCommandText)
                    
    If (rstDVC.EOF And rstDVC.BOF) Then
        LoadDefaultSettings = QueryResultNoRecord
    Else
        Me.DVCIDs = ""
        Me.Widths = ""
        Me.Alignments = ""
        Me.RequiredFields = ""
        Me.ColumnFormats = ""
        Me.AutoGroup = False
        Me.CardView = False
        Me.GridLines = False
        Me.GroupHeaderIsVisible = False
        Me.OddEvenColor = False
        Me.ReadingPane = False
        Me.SelectedRows = ""
        Me.Sort = ""
        Me.UserID = NodeProps.UserID
        Me.NodeID = NodeProps.NodeID
        Me.TView_ID = rstDVC![Nodes.TView_ID]
        Me.GroupHeaderIsVisible = False
        
        rstDVC.MoveFirst
        Do While Not rstDVC.EOF
            If rstDVC!DVC_Requirement = 1 Then
                Me.RequiredFields = Me.RequiredFields & rstDVC!DVC_ID & "*****"
            Else
                Me.DVCIDs = Me.DVCIDs & rstDVC!DVC_ID & "*****"
                Me.Widths = Me.Widths & rstDVC!DVC_Width & "*****"
                Me.Alignments = Me.Alignments & UCase(Trim(rstDVC!DVC_Alignment)) & "*****"
                If Trim(FNullField(rstDVC!DVC_Format)) <> "" Then
                    Me.ColumnFormats = Me.ColumnFormats & rstDVC!DVC_ID & "~" & rstDVC!DVC_Format & "*****"
                End If
            End If
            
            rstDVC.MoveNext
        Loop
        
        If Len(Me.DVCIDs) > 5 Then
            Me.DVCIDs = Mid(Me.DVCIDs, 1, Len(Me.DVCIDs) - 5)
        End If
        If Len(Me.Widths) > 5 Then
            Me.Widths = Mid(Me.Widths, 1, Len(Me.Widths) - 5)
        End If
        If Len(Me.Alignments) > 5 Then
            Me.Alignments = Mid(Me.Alignments, 1, Len(Me.Alignments) - 5)
        End If
        If Len(Me.RequiredFields) > 5 Then
            Me.RequiredFields = Mid(Me.RequiredFields, 1, Len(Me.RequiredFields) - 5)
        End If
        If Len(Me.ColumnFormats) > 5 Then
            Me.ColumnFormats = Mid(Me.ColumnFormats, 1, Len(Me.ColumnFormats) - 5)
        End If
        
        LoadDefaultSettings = QueryResultSuccessful
    End If
    
    'Set rstDVC = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    
End Function
Friend Function GetDVCID(ByRef DVCID As Long, ByVal Alias As String, _
                         ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim rstDVC As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_ID "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "TView_ID = " & Me.TView_ID
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "DVC_FieldAlias = '" & Alias & "'"
        
    ADORecordsetOpen strCommandText, ADOConnection, rstDVC, adOpenKeyset, adLockOptimistic
    'Set rstDVC = ADOConnection.Execute(strCommandText)
    
    If Not (rstDVC.EOF And rstDVC.BOF) Then
        rstDVC.MoveFirst
        
        DVCID = rstDVC!DVC_ID
        GetDVCID = QueryResultSuccessful
    Else
        GetDVCID = QueryResultNoRecord
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    
    Exit Function
    
Error_Handler:
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVC)
    
    GetDVCID = QueryResultError
    
End Function


Friend Function GetFormatConditions(ByRef GridProps As CGrid, ByRef FCRecord As ADODB.Recordset, _
                    ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim strCommandText As String
    
    On Error GoTo Error_Handler:
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UVCFormatCondition "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UVC_ID = " & GridProps.UVC_ID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Node_ID = " & GridProps.NodeID & " "
    strCommandText = strCommandText & "ORDER BY "
    strCommandText = strCommandText & "FC_Priority "
    
    ADORecordsetOpen strCommandText, ADOConnection, FCRecord, adOpenKeyset, adLockOptimistic
    'Call RstOpen(strCommandText, ADOConnection, FCRecord, adOpenKeyset, adLockReadOnly)
    
    If FCRecord.RecordCount = 0 Then
        GetFormatConditions = AddDefaultUVCFC(GridProps.UVC_ID, ADOConnection)
        Call ADORecordsetClose(FCRecord)
        
        ADORecordsetOpen strCommandText, ADOConnection, FCRecord, adOpenKeyset, adLockOptimistic
        'Call RstOpen(strCommandText, ADOConnection, FCRecord, adOpenKeyset, adLockReadOnly)
    Else
        GetFormatConditions = QueryResultSuccessful
    End If
    
    Exit Function
    
Error_Handler:
    GetFormatConditions = QueryResultError
    
End Function

Friend Function IsHidden(ByRef ADOConnection As ADODB.Connection, ByVal Caption As String) As Boolean
    
    Dim rstColumn As ADODB.Recordset
    Dim strCommandText As String
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_Requirement "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "DVC_FieldAlias = '" & Caption & "'"
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "TView_ID = " & Me.TView_ID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstColumn, adOpenKeyset, adLockOptimistic
    'Set rstColumn = ADOConnection.Execute(strCommandText)
    
    If Not (rstColumn.EOF And rstColumn.BOF) Then
        rstColumn.MoveFirst
        
        If rstColumn!DVC_Requirement = 1 Then
            IsHidden = True
        Else
            IsHidden = False
        End If
    End If
    
    'Set rstColumn = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstColumn)
    
End Function

Friend Sub TriggerAfterGridUpdate()

    RaiseEvent AfterGridUpdate

End Sub

Friend Function LicenseeRecord(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim rstLicensee As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Licensee "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstLicensee, adOpenKeyset, adLockOptimistic
    'Set rstLicensee = ADOConnection.Execute(strCommandText)
    
    If (rstLicensee.EOF And rstLicensee.BOF) Then
        LicenseeRecord = QueryResultNoRecord
    Else
        LicenseeRecord = QueryResultSuccessful
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstLicensee)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstLicensee)
    
    Err.Clear
    LicenseeRecord = QueryResultError
    
End Function

Friend Function HeaderIconPath(ByRef ADOConnection As ADODB.Connection, ByVal DVCID As Long) As String

    Dim rstDVCIcon As ADODB.Recordset
    
    Dim strCommandText As String
    Dim strImageLoc As String
    
    Dim bytImage() As Byte
    
    Dim lngImgHandle As Long
    
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " DVC_ID, "
    strCommandText = strCommandText & " DVC_HeaderIcon "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " DefaultViewColumns "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " DVC_ID = " & DVCID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstDVCIcon, adOpenKeyset, adLockOptimistic
    'Set rstDVCIcon = ADOConnection.Execute(strCommandText)
    
    If Not (rstDVCIcon.EOF And rstDVCIcon.BOF) Then
        rstDVCIcon.MoveFirst
        
        strImageLoc = WindowsTempPath & "DVC" & Trim(rstDVCIcon!DVC_ID) & ".img"
        
        If Len(Dir$(strImageLoc)) Then
            Kill strImageLoc
        End If
        
        If Not IsNull((rstDVCIcon!DVC_HeaderIcon)) Then
            bytImage = rstDVCIcon.Fields("DVC_HeaderIcon").GetChunk(LenB(rstDVCIcon!DVC_HeaderIcon))
            lngImgHandle = FreeFile()
            Open strImageLoc For Binary As #lngImgHandle
            Put #lngImgHandle, , bytImage()
            Close #lngImgHandle
            
            HeaderIconPath = strImageLoc
        Else
            HeaderIconPath = ""
        End If
        
        
    Else
        HeaderIconPath = ""
    End If
    
    'Set rstDVCIcon = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstDVCIcon)
    
End Function

Friend Sub TriggerSetColumnFormat(ByRef Column As JSColumn)

    RaiseEvent SetColumnFormat(Column)
    
End Sub

Friend Sub TriggerBeforeShowFields(ByRef FieldDescriptions As ADODB.Recordset)

    RaiseEvent BeforeShowFields(FieldDescriptions)
    
End Sub
