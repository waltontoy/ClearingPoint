VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNavigationPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarGroupID As Long 'local copy
Private mvarCaption As String 'local copy
Private mvarUserID As Long 'local copy
Private mvarNodeID As Long 'local copy
Private mvarTreeID As Long 'local copy
Private mvarTopNode As Long 'local copy
Private mvarSelectedNode As SSNode 'local copy

Public tvwMainControl As SSActiveTreeView.SSTree
Public tvwShortcutControl As SSActiveTreeView.SSTree

Private mvarSelectedItemID As Long 'local copy
Private mvarFeatureID As Long 'local copy
Friend Property Let FeatureID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FeatureID = 5
    mvarFeatureID = vData
End Property


Friend Property Get FeatureID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FeatureID
    FeatureID = mvarFeatureID
End Property




Friend Property Let SelectedItemID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarSelectedItemID = vData
End Property


Public Property Get SelectedItemID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    SelectedItemID = mvarSelectedItemID
End Property

Public Property Set SelectedNode(ByVal vData As SSNode)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SelectedNode = Form1
    Set mvarSelectedNode = vData
End Property


Public Property Get SelectedNode() As SSNode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedNode
    Set SelectedNode = mvarSelectedNode
End Property

Friend Property Let TopNode(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopNode = 5
    mvarTopNode = vData
End Property

Friend Property Get TopNode() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopNode
    TopNode = mvarTopNode
End Property

Friend Property Let TreeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreeID = 5
    mvarTreeID = vData
End Property


Friend Property Get TreeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeID
    TreeID = mvarTreeID
End Property


Friend Property Let NodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NodeID = 5
    mvarNodeID = vData
    mvarSelectedItemID = mvarTreeID
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

Friend Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Friend Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Friend Property Let GroupID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarGroupID = vData
End Property


Friend Property Get GroupID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    GroupID = mvarGroupID
End Property





Friend Function LoadGroups(ByRef Window As Object, ByRef ADOConnection As ADODB.Connection, _
                    ByVal UserID As Long) As QueryResultConstants
    
    Dim rstGroups As ADODB.Recordset
    Dim rstBGroup As ADODB.Recordset
    Dim strCommandText As String
    
Repeat_Loading:
    On Error GoTo Error_Handler
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " ButtonGroups "
    strCommandText = strCommandText & " LEFT OUTER JOIN "
    strCommandText = strCommandText & " Features "
    strCommandText = strCommandText & " ON ButtonGroups.Feature_ID = Features.Feature_ID "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " User_ID = " & UserID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " (Feature_Activated = True "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " ButtonGroups.Feature_ID = 0 "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " IsNull(ButtonGroups.Feature_ID) = True "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " Trim(ButtonGroups.Feature_ID) = '') "
    strCommandText = strCommandText & " ORDER BY Tree_ID"
    
    ADORecordsetOpen strCommandText, ADOConnection, rstGroups, adOpenKeyset, adLockOptimistic
    'Set rstGroups = ADOConnection.Execute(strCommandText)
    
    If (rstGroups.EOF And rstGroups.BOF) Then
        strCommandText = vbNullString
        strCommandText = strCommandText & "Select "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "ButtonGroups "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "BGroup_Default = True "
        strCommandText = strCommandText & "AND "
        strCommandText = strCommandText & "User_ID = 0 "
            
        ADORecordsetOpen strCommandText, ADOConnection, rstBGroup, adOpenKeyset, adLockOptimistic
        'Set rstBGroup = ADOConnection.Execute(strCommandText)
        
        If (rstBGroup.EOF And rstBGroup.BOF) Then
            '>> No default groups defined
            LoadGroups = QueryResultNoRecord
            Exit Function
        Else
            rstBGroup.MoveFirst
        End If
        
        Do While Not rstBGroup.EOF
            Me.UserID = UserID
            Me.TreeID = rstBGroup!Tree_ID
            'Me.NodeID = rstBGroup!BGroup_SelectedNodeID
            Me.NodeID = GetDefaultNodeForDefaultGroup(ADOConnection, Me.TreeID)
            Me.Caption = Translate(rstBGroup!BGroup_Caption)
            Me.TopNode = rstBGroup!BGroup_TopNode
            Me.FeatureID = rstBGroup!Feature_ID
            g_lngTopNode = Me.TopNode
            
            AddGroup ADOConnection
            
            rstBGroup.MoveNext
        Loop
        
        GoTo Repeat_Loading
        
    Else
        Load frmPane
        
        g_typInterface.IButtonBar.AddImageList g_typInterface.IBBarIcons
            
        rstGroups.MoveFirst
        Do While rstGroups.EOF = False
            g_typInterface.IButtonBar.AddItem rstGroups!Tree_ID, Replace(Translate(rstGroups!BGroup_Caption), "&", "&&"), frmPane.hwnd
            rstGroups.MoveNext
        Loop
    End If
    
    LoadGroups = QueryResultSuccessful
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstGroups)
    Call ADORecordsetClose(rstBGroup)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstGroups)
    Call ADORecordsetClose(rstBGroup)

    LoadGroups = QueryResultError
    
End Function

Private Function GetDefaultNodeForDefaultGroup(ByRef ADOConnection As ADODB.Connection, _
                                            ByVal TreeID As Long) As Long
    Dim strSQL As String
    Dim rstDefaultNodeForDefaultGroup As ADODB.Recordset
    
    strSQL = vbNullString
    strSQL = strSQL & " SELECT TOP 1 "
    strSQL = strSQL & " Nodes.Node_ID AS Node_ID "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & " Nodes "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " Node_Level =  " & 0 & " "
    strSQL = strSQL & " AND "
    strSQL = strSQL & " Tree_ID = " & TreeID
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " Node_ID "
    ADORecordsetOpen strSQL, ADOConnection, rstDefaultNodeForDefaultGroup, adOpenKeyset, adLockOptimistic
    'Set rstDefaultNodeForDefaultGroup = ADOConnection.Execute(strSQL)
    
    If Not (rstDefaultNodeForDefaultGroup.EOF And rstDefaultNodeForDefaultGroup.BOF) Then
        rstDefaultNodeForDefaultGroup.MoveFirst
        
        GetDefaultNodeForDefaultGroup = rstDefaultNodeForDefaultGroup.Fields("Node_ID").Value
    Else
        Debug.Assert False
        ' This should never be used. But just in case
    End If
    
    ADORecordsetClose rstDefaultNodeForDefaultGroup
End Function

Friend Function SelectShortcutButtons(ByRef ShortcutButtons As ADODB.Recordset, _
                            ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    
    
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " Buttons.Node_ID As Node_ID, "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " ( "
    strCommandText = strCommandText & " Buttons "
    strCommandText = strCommandText & " INNER JOIN "
    strCommandText = strCommandText & " Nodes "
    strCommandText = strCommandText & " ON "
    strCommandText = strCommandText & " Buttons.Node_ID "
    strCommandText = strCommandText & " = "
    strCommandText = strCommandText & " Nodes.Node_ID "
    strCommandText = strCommandText & " ) "
    strCommandText = strCommandText & " LEFT OUTER JOIN "
    strCommandText = strCommandText & " Features "
    strCommandText = strCommandText & " ON Nodes.Feature_ID = Features.Feature_ID "
    
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " BGroup_ID = " & Me.GroupID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Deleted = False "
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " (Feature_Activated = True "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " Nodes.Feature_ID = 0 "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " IsNull(Nodes.Feature_ID) = True "
    strCommandText = strCommandText & " OR "
    strCommandText = strCommandText & " Trim(Nodes.Feature_ID) = '') "
    
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & " Button_Order "
    
    ADORecordsetOpen strCommandText, ADOConnection, ShortcutButtons, adOpenKeyset, adLockOptimistic
    'Set ShortcutButtons = ADOConnection.Execute(strCommandText)
    
    If (ShortcutButtons.EOF And ShortcutButtons.BOF) Then
        SelectShortcutButtons = QueryResultNoRecord
    Else
        SelectShortcutButtons = QueryResultSuccessful
    End If
    
    Exit Function
    
Error_Handler:
    SelectShortcutButtons = QueryResultError
    
End Function

Friend Function AddShortcutButton(ByRef ADOConnection As ADODB.Connection, _
                                    ByVal ShortcutPos As Long) As QueryResultConstants

    Dim rstShortcuts As ADODB.Recordset
    Dim strCommandText As String
    Dim lngRecordsAffected As Long
    Dim strButtonCaption As String
    
    On Error GoTo Error_Handler
    
    If InStr(1, frmPane.tvwMain.SelectedItem.Text, "(") Then
        strButtonCaption = Left(frmPane.tvwMain.SelectedItem.Text, InStr(1, frmPane.tvwMain.SelectedItem.Text, "(") - 1)
    Else
        strButtonCaption = frmPane.tvwMain.SelectedItem.Text
    End If
    
    strCommandText = vbNullString
    strCommandText = strCommandText & " UPDATE "
    strCommandText = strCommandText & " Buttons "
    strCommandText = strCommandText & " SET "
    strCommandText = strCommandText & " BGroup_ID = " & Me.GroupID & ", "
    strCommandText = strCommandText & " Button_Caption = '" & strButtonCaption & "', "
    strCommandText = strCommandText & " Button_Deleted = False, "
    strCommandText = strCommandText & " Button_Order = " & ShortcutPos
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " Tree_ID = " & Me.TreeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Node_ID = " & Me.NodeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Deleted = True "
    
    lngRecordsAffected = ExecuteNonQuery(ADOConnection, strCommandText)
    'ADOConnection.Execute strCommandText, lngRecordsAffected
    
    If lngRecordsAffected = 0 Then
        strCommandText = vbNullString
        strCommandText = strCommandText & "INSERT INTO "
        strCommandText = strCommandText & "Buttons "
        strCommandText = strCommandText & "( "
        strCommandText = strCommandText & "BGroup_ID, "
        strCommandText = strCommandText & "Button_Caption, "
        strCommandText = strCommandText & "Button_IconName, "
        strCommandText = strCommandText & "Tree_ID, "
        strCommandText = strCommandText & "Button_Default, "
        strCommandText = strCommandText & "Node_ID, "
        strCommandText = strCommandText & "Button_Deleted, "
        strCommandText = strCommandText & "Button_Order "
        strCommandText = strCommandText & ") "
        strCommandText = strCommandText & "VALUES "
        strCommandText = strCommandText & "( "
        strCommandText = strCommandText & Me.GroupID & ", "
        strCommandText = strCommandText & "'" & strButtonCaption & "', "
        strCommandText = strCommandText & "' ', "
        strCommandText = strCommandText & Me.TreeID & ", "
        strCommandText = strCommandText & "False, "
        strCommandText = strCommandText & Me.NodeID & ", "
        strCommandText = strCommandText & 0 & ", "
        strCommandText = strCommandText & ShortcutPos
        strCommandText = strCommandText & ") "
        
        ExecuteNonQuery ADOConnection, strCommandText
        'ADOConnection.Execute strCommandText
    
    End If
        
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Buttons "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " BGroup_ID = " & Me.GroupID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Order >= " & ShortcutPos
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Node_ID <> " & Me.NodeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Deleted = False "
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & " Button_Order "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstShortcuts, adOpenKeyset, adLockOptimistic
    'Set rstShortcuts = ADOConnection.Execute(strCommandText)
    
    If Not (rstShortcuts.EOF And rstShortcuts.BOF) Then
        rstShortcuts.MoveFirst
        
        Do While Not rstShortcuts.EOF
            Call ChangeShortcutPos(ADOConnection, Me.GroupID, rstShortcuts!Node_ID, rstShortcuts!Button_Order + 1)
            rstShortcuts.MoveNext
        Loop
    End If
    
    'Set rstShortcuts = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstShortcuts)
    
    AddShortcutButton = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstShortcuts)
    
    AddShortcutButton = QueryResultError
End Function

Friend Function DeleteShortcutButton(ByRef ADOConnection As ADODB.Connection, ByVal ShortcutPos As Long) As QueryResultConstants

    Dim rstShortcuts As ADODB.Recordset
    
    Dim lngAffectedRecords As Long
    Dim strCommandText As String
    
    On Error GoTo Error_Handler

    
    strCommandText = vbNullString
    strCommandText = strCommandText & " UPDATE "
    strCommandText = strCommandText & " Buttons "
    strCommandText = strCommandText & " SET "
    strCommandText = strCommandText & " Button_Deleted = True "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " BGroup_ID = " & Me.GroupID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Tree_ID = " & Me.TreeID
    strCommandText = strCommandText & " AND "
    '>>>>>>>> p4tric 121907
    strCommandText = strCommandText & " Node_ID = " & Me.tvwShortcutControl.SelectedItem.Key 'me.NodeID
    
    lngAffectedRecords = ExecuteNonQuery(ADOConnection, strCommandText)
    'Call ADOConnection.Execute(strCommandText, lngAffectedRecords)
    
    If lngAffectedRecords > 0 Then
        strCommandText = vbNullString
        strCommandText = strCommandText & " SELECT "
        strCommandText = strCommandText & " * "
        strCommandText = strCommandText & " FROM "
        strCommandText = strCommandText & " Buttons "
        strCommandText = strCommandText & " WHERE "
        strCommandText = strCommandText & " Button_Order > " & ShortcutPos
        strCommandText = strCommandText & " AND "
        strCommandText = strCommandText & " BGroup_ID = " & Me.GroupID
        strCommandText = strCommandText & " AND "
        strCommandText = strCommandText & " Tree_ID = " & Me.TreeID
        strCommandText = strCommandText & " AND "
        strCommandText = strCommandText & " Button_Deleted = False "
        strCommandText = strCommandText & " ORDER BY "
        strCommandText = strCommandText & " Button_Order "
           
        ADORecordsetOpen strCommandText, ADOConnection, rstShortcuts, adOpenKeyset, adLockOptimistic
        'Set rstShortcuts = ADOConnection.Execute(strCommandText)
        
        If Not (rstShortcuts.EOF And rstShortcuts.BOF) Then
            rstShortcuts.MoveFirst
            Do While Not rstShortcuts.EOF
                Call ChangeShortcutPos(ADOConnection, Me.GroupID, rstShortcuts!Node_ID, rstShortcuts!Button_Order - 1)
                rstShortcuts.MoveNext
            Loop
        End If
        
        'Set rstShortcuts = Nothing
        ' hobbes 10/18/2005
        Call ADORecordsetClose(rstShortcuts)
        
    End If
    
    DeleteShortcutButton = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstShortcuts)
    
    DeleteShortcutButton = QueryResultError

End Function

Friend Function LoadTreeNodes(ByRef TreeNode As ADODB.Recordset, Tree_ID As Long, _
                    ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    
    
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Nodes "
    strCommandText = strCommandText & " INNER JOIN "
    strCommandText = strCommandText & " Features "
    strCommandText = strCommandText & " ON Nodes.Feature_ID = Features.Feature_ID "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " Tree_ID = " & Tree_ID & " "
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Feature_Activated = True "
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & " Node_Level, "
    strCommandText = strCommandText & " Node_ID "
    
    ADORecordsetOpen strCommandText, ADOConnection, TreeNode, adOpenKeyset, adLockOptimistic
    'Set TreeNode = ADOConnection.Execute(strCommandText)
        
    If (TreeNode.EOF And TreeNode.BOF) Then
        LoadTreeNodes = QueryResultNoRecord
    Else
        LoadTreeNodes = QueryResultSuccessful
    End If
    
    Exit Function
    
Error_Handler:

    LoadTreeNodes = QueryResultError
    
End Function

Friend Function SelectGroup(ByVal UserID As Long, ByVal TreeID As Long, _
                        ByRef ADOConnection As ADODB.Connection, _
                        Optional ForValidationOnly As Boolean = False) As QueryResultConstants
    
    Dim rstGroups As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "ButtonGroups "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Tree_ID = " & TreeID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstGroups, adOpenKeyset, adLockOptimistic
    'Set rstGroups = ADOConnection.Execute(strCommandText)
    
    If Not (rstGroups.EOF And rstGroups.BOF) Then
        rstGroups.MoveFirst
        
        If ForValidationOnly = False Then
            Me.UserID = rstGroups!User_ID
            Me.GroupID = rstGroups!BGroup_ID
            Me.TopNode = rstGroups!BGroup_TopNode
            g_lngTopNode = Me.TopNode
            Me.NodeID = rstGroups!BGroup_SelectedNodeID
            Me.Caption = Translate(rstGroups!BGroup_Caption)
            Me.TreeID = rstGroups!Tree_ID
            Me.FeatureID = rstGroups!Feature_ID
        End If
        SelectGroup = QueryResultSuccessful
    Else
        SelectGroup = QueryResultNoRecord
    End If
    
    'Set rstGroups = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstGroups)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstGroups)
    
    SelectGroup = QueryResultError
    
End Function

Friend Function AddGroup(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants


    Dim strCommandText As String

    On Error GoTo Error_Handler

    strCommandText = vbNullString
    strCommandText = strCommandText & " INSERT INTO "
    strCommandText = strCommandText & " ButtonGroups "
    strCommandText = strCommandText & " ("
    strCommandText = strCommandText & " User_ID, "
    strCommandText = strCommandText & " BGroup_Caption, "
    strCommandText = strCommandText & " BGroup_TopNode, "
    strCommandText = strCommandText & " BGroup_SelectedNodeID, "
    strCommandText = strCommandText & " Tree_ID, "
    strCommandText = strCommandText & " Feature_ID "
    strCommandText = strCommandText & ") "
    strCommandText = strCommandText & " VALUES "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & Me.UserID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "'" & Me.Caption & "'"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.TopNode
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.NodeID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.TreeID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.FeatureID
    strCommandText = strCommandText & ") "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText

    AddGroup = QueryResultSuccessful

    Exit Function

Error_Handler:
    AddGroup = QueryResultError

End Function

Friend Function GetExpandedNodes(ByRef ADOConnection As ADODB.Connection, _
                                ExpandedNodes As ADODB.Recordset) As QueryResultConstants

    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "TreeSet_ExpandedNodes "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "TreeSettings "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
    
    ADORecordsetOpen strCommandText, ADOConnection, ExpandedNodes, adOpenKeyset, adLockOptimistic
    'Set ExpandedNodes = ADOConnection.Execute(strCommandText)
    
    If (ExpandedNodes.EOF And ExpandedNodes.BOF) Then
        GetExpandedNodes = QueryResultNoRecord
    Else
        GetExpandedNodes = QueryResultSuccessful
    End If
    
    Exit Function

Error_Handler:
    GetExpandedNodes = QueryResultError
    
End Function

Friend Function AddTreeSettings(ByVal ExpandedNodes As String, _
                                ByRef ADOConnection As ADODB.Connection) _
                                As QueryResultConstants


    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "TreeSettings "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & "User_ID"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Tree_ID"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "TreeSet_NodeCount"       'Added by BCo 2006-05-18 As instructed by 2dit in case still needed
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "TreeSet_ExpandedNodes"
    strCommandText = strCommandText & ") "
    strCommandText = strCommandText & "VALUES "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & Me.UserID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.TreeID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Len(ExpandedNodes) \ 5    'Added by BCo 2006-05-18 As instructed by 2dit in case still needed
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "'" & ExpandedNodes & "'"
    strCommandText = strCommandText & ") "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    AddTreeSettings = QueryResultSuccessful
        
    Exit Function
    
Error_Handler:
    
    AddTreeSettings = QueryResultError


End Function


Friend Function UpdateTreeSettings(ByVal ExpandedNodes As String, _
                                    ByRef ADOConnection As ADODB.Connection) _
                                    As QueryResultConstants


    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "TreeSettings "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "TreeSet_NodeCount = " & Len(ExpandedNodes) \ 5 & ", "    'Added by BCo 2006-05-18 As instructed by 2dit in case still needed
    strCommandText = strCommandText & "TreeSet_ExpandedNodes = '" & ExpandedNodes & "'"
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID & " "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    
    UpdateTreeSettings = QueryResultSuccessful
        
    Exit Function
    
Error_Handler:
    
    UpdateTreeSettings = QueryResultError


End Function

Friend Function SelectTreeSettings(ByRef ADOConnection As ADODB.Connection) _
                                    As QueryResultConstants

    Dim rstTreeSettings As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "Select "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "TreeSettings "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID & " "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstTreeSettings, adOpenKeyset, adLockOptimistic
    'Set rstTreeSettings = ADOConnection.Execute(strCommandText)
    
    If Not (rstTreeSettings.EOF And rstTreeSettings.BOF) Then
        SelectTreeSettings = QueryResultSuccessful
    Else
        SelectTreeSettings = QueryResultNoRecord
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstTreeSettings)
        
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstTreeSettings)
    
    SelectTreeSettings = QueryResultError


End Function


Friend Function UpdateGroupSettings(ByRef ADOConnection As ADODB.Connection) _
                                    As QueryResultConstants
                                                             
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "ButtonGroups "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "BGroup_TopNode = " & Me.TopNode & ", "
    strCommandText = strCommandText & "BGroup_SelectedNodeID = " & Me.NodeID & " "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "User_ID = " & Me.UserID
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    UpdateGroupSettings = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    UpdateGroupSettings = QueryResultError

End Function

Friend Function LoadViewOptions(ByVal NodeID As Long, _
                                ByRef ADOConnection As ADODB.Connection, _
                                ByRef ViewOptions As ADODB.Recordset) _
                                As QueryResultConstants
    
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Views "
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " Node_ID = " & NodeID
    strCommandText = strCommandText & " ORDER BY "
    strCommandText = strCommandText & " View_Position "
        
    ADORecordsetOpen strCommandText, ADOConnection, ViewOptions, adOpenKeyset, adLockOptimistic
    'Set ViewOptions = ADOConnection.Execute(strCommandText)
    
    If (ViewOptions.EOF And ViewOptions.BOF) Then
        LoadViewOptions = QueryResultNoRecord
    Else
        LoadViewOptions = QueryResultSuccessful
    End If
    
    Exit Function
    
Error_Handler:
    LoadViewOptions = QueryResultError
    
End Function

Friend Function NodeView(ByVal UserID As Long, ByVal NodeID As Long, _
                        ByRef ADOConnection As ADODB.Connection, _
                        ByRef SelectedView As Long) _
                        As QueryResultConstants

    Dim rstView As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "View_Position "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Views "
    strCommandText = strCommandText & "INNER JOIN "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "ON "
    strCommandText = strCommandText & "Views.View_ID = UserViewColumns.View_ID "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "UserViewColumns.Node_ID = " & NodeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "UserViewColumns.User_ID = " & UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "UserViewColumns.UVC_SelectedView = True "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstView, adOpenKeyset, adLockOptimistic
    'Set rstView = ADOConnection.Execute(strCommandText)
    
    If Not (rstView.EOF And rstView.BOF) Then
        SelectedView = rstView!View_Position
        NodeView = QueryResultSuccessful
    Else
        NodeView = QueryResultNoRecord
    End If
    
    'Set rstView = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstView)
    
    Exit Function
    
Error_Handler:
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstView)
    
    NodeView = QueryResultError
End Function
Friend Function SelectViewSettings(ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
    
    Dim rstView As ADODB.Recordset
    Dim strCommandText As String
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "View_ID = " & GridProps.ViewID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstView, adOpenKeyset, adLockOptimistic
    'Set rstView = ADOConnection.Execute(strCommandText)
    
    If (rstView.EOF And rstView.BOF) Then
        SelectViewSettings = QueryResultNoRecord
    Else
        SelectViewSettings = QueryResultSuccessful
    End If
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstView)
    
End Function

Friend Function AddViewSettings(ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim rstViews As ADODB.Recordset
    Dim lngCtr As Long
    Dim strCommandText As String
                
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "UserViewColumns "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "UVC_SelectedView = " & IIf(frmPane.optViews(frmPane.Controls(lngCtr).Index).Value, "True", "False") & " "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID & " "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "View_ID = " & GridProps.ViewID
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    AddViewSettings = QueryResultSuccessful
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstViews)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstViews)
    
    AddViewSettings = QueryResultError


End Function

Friend Function UpdateNodeView(ByRef GridProps As CGrid, ByRef ADOConnection As ADODB.Connection) As QueryResultConstants
                
    Dim rstViews As ADODB.Recordset
    Dim lngCtr As Long
    Dim strCommandText As String
                
    On Error GoTo Error_Handler
    
    For lngCtr = 0 To frmPane.Controls.Count - 1
        If Left(frmPane.Controls(lngCtr).Name, 8) = "optViews" Then
            strCommandText = vbNullString
            strCommandText = strCommandText & "UPDATE "
            strCommandText = strCommandText & "UserViewColumns "
            strCommandText = strCommandText & "SET "
            strCommandText = strCommandText & "UVC_SelectedView = " & IIf(frmPane.optViews(frmPane.Controls(lngCtr).Index).Value, "True", "False") & " "
            strCommandText = strCommandText & "WHERE "
            strCommandText = strCommandText & "User_ID = " & Me.UserID & " "
            strCommandText = strCommandText & "AND "
            strCommandText = strCommandText & "Node_ID = " & Me.NodeID & " "
            strCommandText = strCommandText & "AND "
            strCommandText = strCommandText & "View_ID = " & Val(frmPane.optViews(frmPane.Controls(lngCtr).Index).Tag)
            
            ExecuteNonQuery ADOConnection, strCommandText
            'ADOConnection.Execute strCommandText
        End If
    Next
    
    UpdateNodeView = QueryResultSuccessful
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstViews)
    
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstViews)
    
    UpdateNodeView = QueryResultError
    
End Function

Friend Function RenameShortcutButton(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "Buttons "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "Button_Caption = '" & Me.Caption & "' "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "BGroup_ID = " & Me.GroupID
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Node_ID = " & Me.NodeID
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
    frmPane.tvwShortcut.EndLabelEdit
    
    RenameShortcutButton = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    RenameShortcutButton = QueryResultError


End Function

Friend Function LoadDefaultShortcutButtons(ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim rstButtons As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler

    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "* "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "Buttons "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & "Button_Deleted = True "
    strCommandText = strCommandText & "AND "
    strCommandText = strCommandText & "Button_Default = False "
    
    ADORecordsetOpen strCommandText, ADOConnection, rstButtons, adOpenKeyset, adLockOptimistic
    'Set rstButtons = ADOConnection.Execute(strCommandText)
    
    If (rstButtons.EOF And rstButtons.BOF) Then
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & " * "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "Buttons "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
        strCommandText = strCommandText & " AND "
        strCommandText = strCommandText & "BGroup_ID = 0 "
        strCommandText = strCommandText & " AND "
        strCommandText = strCommandText & "Button_Default = True "
        strCommandText = strCommandText & "ORDER BY "
        strCommandText = strCommandText & "Button_ID "
        
        ADORecordsetOpen strCommandText, ADOConnection, rstButtons, adOpenKeyset, adLockOptimistic
        'Set rstButtons = ADOConnection.Execute(strCommandText)
        
        If Not (rstButtons.EOF And rstButtons.BOF) Then
            strCommandText = vbNullString
            strCommandText = strCommandText & "INSERT INTO "
            strCommandText = strCommandText & "Buttons "
            strCommandText = strCommandText & "SELECT "
            strCommandText = strCommandText & "Button_Caption, "
            strCommandText = strCommandText & "Tree_ID, "
            strCommandText = strCommandText & "Button_IconName, "
            strCommandText = strCommandText & "Node_ID "
            strCommandText = strCommandText & "FROM "
            strCommandText = strCommandText & "Buttons "
            strCommandText = strCommandText & "WHERE "
            strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
            strCommandText = strCommandText & " AND "
            strCommandText = strCommandText & "BGroup_ID = 0 "
            strCommandText = strCommandText & " AND "
            strCommandText = strCommandText & "Button_Default = True "
            strCommandText = strCommandText & "ORDER BY "
            strCommandText = strCommandText & "Button_ID "
            
            ExecuteNonQuery ADOConnection, strCommandText
            'ADOConnection.Execute strCommandText
            
            strCommandText = vbNullString
            strCommandText = strCommandText & "UPDATE "
            strCommandText = strCommandText & "Buttons "
            strCommandText = strCommandText & "SET "
            strCommandText = strCommandText & "BGroup_ID = " & Me.GroupID & ", "
            strCommandText = strCommandText & "Button_Deleted = False, "
            strCommandText = strCommandText & "Button_Default = False "
            strCommandText = strCommandText & "WHERE "
            strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
            strCommandText = strCommandText & " AND "
            strCommandText = strCommandText & "(BGroup_ID = 0 "
            strCommandText = strCommandText & " OR "
            strCommandText = strCommandText & "IsNull(BGroup_ID) = True) "
            strCommandText = strCommandText & "AND "
            strCommandText = strCommandText & "Button_Default = False "
            
            ExecuteNonQuery ADOConnection, strCommandText
            'ADOConnection.Execute strCommandText
            
            LoadDefaultShortcutButtons = QueryResultSuccessful
        Else
            LoadDefaultShortcutButtons = QueryResultNoRecord
        End If
    Else
        LoadDefaultShortcutButtons = QueryResultNoRecord
    End If
    
    'Set rstButtons = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstButtons)
    
    Exit Function
    
Error_Handler:
    
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstButtons)
    
    LoadDefaultShortcutButtons = QueryResultError

End Function

Friend Sub ChangeShortcutPos(ByRef ADOConnection As ADODB.Connection, _
                            ByVal Group_ID As Long, _
                            ByVal Node_ID As Long, _
                            ByVal NewPos As Long)

    Dim strCommandText As String
    
    strCommandText = vbNullString
    strCommandText = strCommandText & " UPDATE "
    strCommandText = strCommandText & " Buttons "
    strCommandText = strCommandText & " SET "
    strCommandText = strCommandText & " Button_Order = " & NewPos
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & " Node_ID = " & Node_ID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " BGroup_ID = " & Group_ID
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Deleted = False "
    strCommandText = strCommandText & " AND "
    strCommandText = strCommandText & " Button_Default = False "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
    
End Sub

Private Sub Class_Initialize()

    Set tvwMainControl = frmPane.tvwMain
    Set tvwShortcutControl = frmPane.tvwShortcut
    
End Sub
