VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMainControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum QueryResultConstants
    QueryResultError = 0
    QueryResultSuccessful = 1
    QueryResultNoRecord = 2
End Enum

Public Enum ReadingPaneLocationConstants
    PaneNotVisible = 0
    BotttomPane = 1
    RightPane = 2
End Enum

Public Enum MouseStateConstants
    MouseDown = 1
    MouseUp = 2
    MouseMove = 3
End Enum

Public Enum MainControlConstants
    Navigation = 1
    RightReadingPane = 2
    BottomReadingPane = 3
    ReadingPaneOff = 4
    Find = 5
    Status_Bar = 6
End Enum
    
'local variable(s) to hold property value(s)
Private mvarUserID As Long
Private mvarNavigationPane As Boolean
Private mvarReadingPane As Boolean
Private mvarFindBox As Boolean
Private mvarNavigationPaneWidth As Double
Private mvarGridWidth As Double
Private mvarReadingPanePos As Integer
Private mvarGridHeight As Double
Private mvarStatusBar As Boolean
Private mvarExpandedButtons As Long
Private mvarTreeID As Long
Private mvarUser_Password As String
Private mvarViewOptions As Boolean
Private mvarCbarState As String
Private mvarCBarVersion As Double
Private mvarCustomFind As Boolean
Private mvarUser_Name As String
Private mvarDefaultNavPaneWidth As Double 'local copy
Private mvarDefaultExpandedCount As Long 'local copy
Private mvarDefaultWindowHeight As Double 'local copy
Private mvarDefaultWindowWidth As Double 'local copy
Private mvarUseCustomFind As Boolean 'local copy


Public Property Let UseCustomFind(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultWindowWidth = 5
    mvarUseCustomFind = vData
End Property

Public Property Get UseCustomFind() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultWindowWidth
    UseCustomFind = mvarUseCustomFind
End Property

Public Property Let DefaultWindowWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultWindowWidth = 5
    mvarDefaultWindowWidth = vData
End Property


Public Property Get DefaultWindowWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultWindowWidth
    DefaultWindowWidth = mvarDefaultWindowWidth
End Property



Public Property Let DefaultWindowHeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultWindowHeight = 5
    mvarDefaultWindowHeight = vData
End Property


Public Property Get DefaultWindowHeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultWindowHeight
    DefaultWindowHeight = mvarDefaultWindowHeight
End Property



Public Property Let DefaultExpandedCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultExpandedCount = 5
    mvarDefaultExpandedCount = vData
End Property


Public Property Get DefaultExpandedCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultExpandedCount
    DefaultExpandedCount = mvarDefaultExpandedCount
End Property



Public Property Let DefaultNavPaneWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultNavPaneWidth = 5
    mvarDefaultNavPaneWidth = vData
End Property


Public Property Get DefaultNavPaneWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultNavPaneWidth
    DefaultNavPaneWidth = mvarDefaultNavPaneWidth
End Property




Public Property Let User_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.User_Name = 5
    mvarUser_Name = vData
End Property

Public Property Get User_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.User_Name
    User_Name = mvarUser_Name
End Property

Public Property Let CustomFind(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomFind = 5
    mvarCustomFind = vData
End Property

Public Property Get CustomFind() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomFind
    CustomFind = mvarCustomFind
End Property

Public Property Let CbarState(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CbarState = 5
    mvarCbarState = vData
End Property

Public Property Get CbarState() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CbarState
    CbarState = mvarCbarState
End Property

Public Property Let CBarVersion(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CBarVersion = 5
    mvarCBarVersion = vData
End Property

Public Property Get CBarVersion() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CBarVersion
    CBarVersion = mvarCBarVersion
End Property

Public Property Let ViewOptions(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ViewOptions = 5
    mvarViewOptions = vData
End Property

Public Property Get ViewOptions() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ViewOptions
    ViewOptions = mvarViewOptions
End Property

Public Property Let User_Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.User_Password = 5
    mvarUser_Password = vData
End Property

Public Property Get User_Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.User_Password
    User_Password = mvarUser_Password
End Property

Public Property Let TreeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreeID = 5
    mvarTreeID = vData
End Property

Public Property Get TreeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeID
    TreeID = mvarTreeID
End Property

Public Property Let ExpandedButtons(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpandedButtons = 5
    mvarExpandedButtons = vData
End Property

Public Property Get ExpandedButtons() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpandedButtons
    ExpandedButtons = mvarExpandedButtons
End Property

Public Property Let StatusBar(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusBar = 5
    mvarStatusBar = vData
End Property

Public Property Get StatusBar() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusBar
    StatusBar = mvarStatusBar
End Property

Public Property Let GridHeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridHeight = 5
    mvarGridHeight = vData
End Property

Public Property Get GridHeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridHeight
    GridHeight = mvarGridHeight
End Property

Public Property Let ReadingPanePos(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReadingPanePos = 5
    mvarReadingPanePos = vData
End Property

Public Property Get ReadingPanePos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadingPanePos
    ReadingPanePos = mvarReadingPanePos
End Property

Public Property Let GridWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridWidth = 5
    mvarGridWidth = vData
End Property

Public Property Get GridWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridWidth
    GridWidth = mvarGridWidth
End Property

Public Property Let NavigationPaneWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NavigationPaneWidth = 5
    mvarNavigationPaneWidth = vData
End Property

Public Property Get NavigationPaneWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NavigationPaneWidth
    NavigationPaneWidth = mvarNavigationPaneWidth
End Property

Public Property Let FindBox(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FindBox = 5
    mvarFindBox = vData
End Property

Public Property Get FindBox() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FindBox
    FindBox = mvarFindBox
End Property

Public Property Let ReadingPane(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReadingPane = 5
    mvarReadingPane = vData
End Property

Public Property Get ReadingPane() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadingPane
    ReadingPane = mvarReadingPane
End Property

Public Property Let NavigationPane(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NavigationPane = 5
    mvarNavigationPane = vData
End Property

Public Property Get NavigationPane() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NavigationPane
    NavigationPane = mvarNavigationPane
End Property

Public Property Let UserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserID = 5
    mvarUserID = vData
End Property

Public Property Get UserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserID
    UserID = mvarUserID
End Property


Public Function LoadSetting(ByRef ADOConnection As ADODB.Connection, ByVal UserID As Long, Optional CopySetting As Boolean = True) As QueryResultConstants

    Dim rstSettings As ADODB.Recordset
    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "MainSettings "
        strCommandText = strCommandText & "WHERE "
        strCommandText = strCommandText & "User_ID = " & UserID & " "
    
    '>> check if setting already exist in the database
    ADORecordsetOpen strCommandText, ADOConnection, rstSettings, adOpenKeyset, adLockOptimistic
    'Set rstSettings = ADOConnection.Execute(strCommandText)
    
    If Not (rstSettings.EOF And rstSettings.BOF) Then
        If CopySetting = True Then
            rstSettings.MoveFirst
            
            '>> save control settings to properties
            Me.UserID = rstSettings![User_ID]
            Me.NavigationPane = rstSettings![Setting_Navigation]
            Me.ViewOptions = rstSettings![Setting_ViewOptions]
            Me.FindBox = FNullField(rstSettings!Setting_FindBox)
            Me.CustomFind = FNullField(rstSettings!Setting_CustomFind)
            Me.StatusBar = FNullField(rstSettings!Setting_StatusBar)
            Me.ReadingPanePos = FNullField(rstSettings!Setting_ReadingPanePos)
            Me.GridWidth = FNullField(rstSettings!Setting_GridWidth)
            Me.GridHeight = FNullField(rstSettings!Setting_GridHeight)
            Me.NavigationPaneWidth = FNullField(rstSettings!Setting_NavigationWidth)
            Me.ExpandedButtons = FNullField(rstSettings!Setting_ExpandedButtons)
            Me.TreeID = IIf(FNullField(rstSettings!Tree_ID) = 0, 1, FNullField(rstSettings!Tree_ID))
            Me.CbarState = FNullField(rstSettings!Setting_CommandBar)
            Me.CBarVersion = FNullField(rstSettings!Setting_CommandBarVersion)
        End If
        LoadSetting = QueryResultSuccessful
    Else
        LoadSetting = QueryResultNoRecord
    End If
    
    'Set rstSettings = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstSettings)
        
    Exit Function
    
Error_Handler:

    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstSettings)
    
    LoadSetting = QueryResultError
    
End Function

Public Function AddSetting(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants

    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "INSERT INTO "
    strCommandText = strCommandText & "MainSettings "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & "User_ID, "
    strCommandText = strCommandText & "Setting_Navigation, "
    strCommandText = strCommandText & "Setting_ViewOptions, "
    strCommandText = strCommandText & "Setting_FindBox, "
    strCommandText = strCommandText & "Setting_CustomFind, "
    strCommandText = strCommandText & "Setting_StatusBar, "
    strCommandText = strCommandText & "Setting_ReadingPanePos, "
    strCommandText = strCommandText & "Setting_NavigationWidth, "
    strCommandText = strCommandText & "Setting_GridWidth, "
    strCommandText = strCommandText & "Setting_GridHeight, "
    strCommandText = strCommandText & "Setting_ExpandedButtons, "
    strCommandText = strCommandText & "Tree_ID"
    strCommandText = strCommandText & ") "
    strCommandText = strCommandText & "VALUES "
    strCommandText = strCommandText & "("
    strCommandText = strCommandText & Me.UserID
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & IIf(Me.NavigationPane, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & IIf(Me.ViewOptions, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & IIf(Me.FindBox, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & IIf(Me.CustomFind, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & IIf(Me.StatusBar, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.ReadingPanePos
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.NavigationPaneWidth
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.GridWidth
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.GridHeight
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.ExpandedButtons
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & Me.TreeID
    strCommandText = strCommandText & ") "
    
    ExecuteNonQuery ADOConnection, strCommandText
    'ADOConnection.Execute strCommandText
        
    AddSetting = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    
    AddSetting = QueryResultError
    

End Function

Public Function SaveSetting(ByRef ADOConnection As ADODB.Connection) As QueryResultConstants


    Dim strCommandText As String
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "UPDATE "
    strCommandText = strCommandText & "MainSettings "
    strCommandText = strCommandText & "SET "
    strCommandText = strCommandText & "Setting_Navigation = " & IIf(Me.NavigationPane, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_ViewOptions = " & IIf(Me.ViewOptions, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_FindBox = " & IIf(Me.FindBox, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_CustomFind = " & IIf(Me.CustomFind, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_StatusBar = " & IIf(Me.StatusBar, "True", "False")
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_ReadingPanePos = " & Me.ReadingPanePos
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_NavigationWidth = CDbl(" & Replace(Me.NavigationPaneWidth, ",", ".") & ")"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_GridWidth = CDbl(" & Replace(Me.GridWidth, ",", ".") & ")"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_GridHeight = CDbl(" & Replace(Me.GridHeight, ",", ".") & ")"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_ExpandedButtons = " & Me.ExpandedButtons
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_CommandBar = '" & Me.CbarState & "'"
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Setting_CommandBarVersion = '" & Me.CBarVersion & "' "
    strCommandText = strCommandText & ", "
    strCommandText = strCommandText & "Tree_ID = " & Me.TreeID
    strCommandText = strCommandText & " WHERE "
    strCommandText = strCommandText & "User_ID = " & UserID & " "
    
    'ADOConnection.Execute strCommandText
    
    ExecuteNonQuery ADOConnection, strCommandText
    
    SaveSetting = QueryResultSuccessful
        
    Exit Function
    
Error_Handler:
    
    SaveSetting = QueryResultError


End Function

Friend Function GetColumn(ColumnID As Long, ADOConnection As ADODB.Connection) As String
    
    Dim rstColumn As ADODB.Recordset
    Dim strCommandText As String
    
    
    '>> formulate SQL
    strCommandText = vbNullString
    strCommandText = strCommandText & "SELECT "
    strCommandText = strCommandText & "DVC_FieldAlias AS Alias "
    strCommandText = strCommandText & "FROM "
    strCommandText = strCommandText & "DefaultViewColumns "
    strCommandText = strCommandText & "WHERE "
    strCommandText = strCommandText & "DVC_ID = " & ColumnID
    
    ADORecordsetOpen strCommandText, ADOConnection, rstColumn, adOpenKeyset, adLockOptimistic
    'Set rstColumn = ADOConnection.Execute(strCommandText)
    
    If Not (rstColumn.EOF And rstColumn.BOF) Then
        rstColumn.MoveFirst
        
        GetColumn = FNullField(rstColumn!Alias)
    End If
    
    'Set rstColumn = Nothing
    ' hobbes 10/18/2005
    Call ADORecordsetClose(rstColumn)
    
End Function

Private Sub Class_Initialize()
    mvarUseCustomFind = True
End Sub
