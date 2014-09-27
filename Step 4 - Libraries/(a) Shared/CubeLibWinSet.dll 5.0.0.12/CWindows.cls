VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable(s) to hold property value(s)
Private m_lngShowCmd As Long
Private m_lngFlags As Long
Private m_lngMinX As Long
Private m_lngMinY As Long

Private mvarWindowHeight As Double 'local copy
Private mvarWindowLeft As Double 'local copy
Private mvarWindowTop As Double 'local copy
Private mvarWindowWidth As Double 'local copy
Private mvarWindowState As Long 'local copy

Public Enum QueryResultConstants
    QueryResultError = 0
    QueryResultSuccessful = 1
    QueryResultNoRecord = 2
End Enum

Public Property Let WindowShowCmd(ByVal NewValue As Long)
    m_lngShowCmd = NewValue
End Property

Public Property Get WindowShowCmd() As Long
    WindowShowCmd = m_lngShowCmd
End Property

Public Property Let WindowFlags(ByVal NewValue As Long)
    m_lngFlags = NewValue
End Property

Public Property Get WindowFlags() As Long
    WindowFlags = m_lngFlags
End Property

Public Property Let WindowMinX(ByVal NewValue As Long)
    m_lngMinX = NewValue
End Property

Public Property Get WindowMinX() As Long
    WindowMinX = m_lngMinX
End Property

Public Property Let WindowMinY(ByVal NewValue As Long)
    m_lngMinY = NewValue
End Property

Public Property Get WindowMinY() As Long
    WindowMinY = m_lngMinY
End Property

Public Property Let WindowState(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowState = 5
    mvarWindowState = vData
End Property

Public Property Get WindowState() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowState
    WindowState = mvarWindowState
End Property

Public Property Let WindowWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowWidth = 5
    mvarWindowWidth = vData
End Property

Public Property Get WindowWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowWidth
    WindowWidth = mvarWindowWidth
End Property

Public Property Let WindowTop(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowTop = 5
    mvarWindowTop = vData
End Property


Public Property Get WindowTop() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowTop
    WindowTop = mvarWindowTop
End Property

Public Property Let WindowLeft(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowLeft = 5
    mvarWindowLeft = vData
End Property


Public Property Get WindowLeft() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowLeft
    WindowLeft = mvarWindowLeft
End Property

Public Property Let WindowHeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowHeight = 5
    mvarWindowHeight = vData
End Property


Public Property Get WindowHeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowHeight
    WindowHeight = mvarWindowHeight
End Property

Public Function AddSetting(ByRef ADOConnection As ADODB.Connection, _
                           ByVal UserID As Long, _
                           ByVal WindowKey As String, _
                           ByVal WindowShowCmd As Long, _
                           ByVal WindowFlags As Long, _
                           ByVal WindowMinX As Long, _
                           ByVal WindowMinY As Long, _
                           ByVal WindowHeight As Long, _
                           ByVal WindowWidth As Long, _
                           ByVal WindowTop As Long, _
                           ByVal WindowLeft As Long, _
                           ByVal WindowState As Long) As QueryResultConstants
    
    Dim strCommand As String
    
    ' MUCP-159 - Start
    Dim lngAffectedRecords As Long
    ' MUCP-159 - End
    
    On Error GoTo Error_Handler
    
    '>> formulate SQL
        strCommand = vbNullString
        strCommand = strCommand & "INSERT INTO "
        strCommand = strCommand & "WindowSettings "
        strCommand = strCommand & "("
        strCommand = strCommand & "User_ID, "
        strCommand = strCommand & "Window_Key, "
        strCommand = strCommand & "Window_ShowCmd, "
        strCommand = strCommand & "Window_Flags, "
        strCommand = strCommand & "Window_MinX, "
        strCommand = strCommand & "Window_MinY, "
        strCommand = strCommand & "Window_Height, "
        strCommand = strCommand & "Window_Width, "
        strCommand = strCommand & "Window_Top, "
        strCommand = strCommand & "Window_Left, "
        strCommand = strCommand & "Window_State"        '<<< Added for window state saving
        strCommand = strCommand & ") "
        strCommand = strCommand & "VALUES "
        strCommand = strCommand & "("
        strCommand = strCommand & UserID
        strCommand = strCommand & ", '"
        strCommand = strCommand & WindowKey
        strCommand = strCommand & "', "
        strCommand = strCommand & WindowShowCmd
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowFlags
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowMinX
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowMinY
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowHeight
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowWidth
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowTop
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowLeft
        strCommand = strCommand & ", "
        strCommand = strCommand & WindowState           '<<< Added for window state saving
        strCommand = strCommand & ") "
        
    ' MUCP-159 - Start
    'ADOConnection.Execute strCommand
    '
    'AddSetting = QueryResultSuccessful
    
    lngAffectedRecords = ExecuteNonQuery(ADOConnection, strCommand)
    
    If lngAffectedRecords = 0 Then
        AddSetting = QueryResultError
    Else
        AddSetting = QueryResultSuccessful
    End If
    ' MUCP-159 - End
    
    Exit Function
    
Error_Handler:
    
    AddSetting = QueryResultError
End Function

Public Function RemoveSetting(ByRef ADOConnection As ADODB.Connection, _
                              ByVal UserID As Long, _
                              ByVal WindowKey As String) As QueryResultConstants
    Dim strCommand As String
    
    ' MUCP-159 - Start
    Dim lngAffectedRecords As Long
    ' MUCP-159 - End
    
    On Error GoTo Error_Handler
    
        strCommand = vbNullString
        strCommand = strCommand & "DELETE "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "WindowSettings "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "User_ID = " & UserID & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "Window_Key = '" & WindowKey & "' "
    
    ' MUCP-159 - Start
    'ADOConnection.Execute strCommand
    '
    'RemoveSetting = QueryResultSuccessful
    
    lngAffectedRecords = ExecuteNonQuery(ADOConnection, strCommand)
    
    If lngAffectedRecords = 0 Then
        RemoveSetting = QueryResultError
    Else
        RemoveSetting = QueryResultSuccessful
    End If
    ' MUCP-159 - End
    
    Exit Function
    
Error_Handler:
    
    RemoveSetting = QueryResultError
End Function

Public Function UpdateSetting(ByRef ADOConnection As ADODB.Connection, _
                              ByVal UserID As Long, _
                              ByVal WindowKey As String, _
                              ByVal WindowShowCmd As Long, _
                              ByVal WindowFlags As Long, _
                              ByVal WindowMinX As Long, _
                              ByVal WindowMinY As Long, _
                              ByVal WindowHeight As Long, _
                              ByVal WindowWidth As Long, _
                              ByVal WindowTop As Long, _
                              ByVal WindowLeft As Long, _
                              ByVal WindowState As Long) As QueryResultConstants
    Dim strCommand As String
    
    ' MUCP-159 - Start
    Dim lngAffectedRecords As Long
    ' MUCP-159 - End
    
    On Error GoTo Error_Handler

        strCommand = vbNullString
        strCommand = strCommand & "UPDATE "
        strCommand = strCommand & "WindowSettings "
        strCommand = strCommand & "SET "
        strCommand = strCommand & "USER_ID = " & UserID
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Key = '" & WindowKey & "' "
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_ShowCmd = " & WindowShowCmd
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Flags= " & WindowFlags
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_MinX = " & WindowMinX
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_MinY = " & WindowMinY
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Height = " & WindowHeight
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Width = " & WindowWidth
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Top = " & WindowTop
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_Left = " & WindowLeft
        strCommand = strCommand & ", "
        strCommand = strCommand & "Window_State = " & WindowState & " "         '<<< Added for window state saving
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "USER_ID = " & UserID & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "Window_Key = '" & WindowKey & "' "
    
    ' MUCP-159 - Start
    'ADOConnection.Execute strCommand
    '
    'UpdateSetting = QueryResultSuccessful
    
    lngAffectedRecords = ExecuteNonQuery(ADOConnection, strCommand)
    
    If lngAffectedRecords = 0 Then
        UpdateSetting = QueryResultError
    Else
        UpdateSetting = QueryResultSuccessful
    End If
    ' MUCP-159 - End
    Exit Function
    
Error_Handler:
    
    UpdateSetting = QueryResultError
End Function

Public Function SelectSetting(ByRef ADOConnection As ADODB.Connection, _
                              ByVal UserID As Long, _
                              ByVal WindowKey As String) As QueryResultConstants

    
    Dim rstSettings As ADODB.Recordset
    Dim strCommand As String
    
    
    On Error GoTo Error_Handler
    
    
    ' Check if window already has a setting in the database
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "* "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "WindowSettings "
        strCommand = strCommand & "WHERE "
        strCommand = strCommand & "User_ID = " & UserID & " "
        strCommand = strCommand & "AND "
        strCommand = strCommand & "Window_Key = '" & WindowKey & "' "
    
    ' MUCP-159 - Start
    'Set rstSettings = ADOConnection.Execute(strCommand)
    ADORecordsetOpen strCommand, ADOConnection, rstSettings, adOpenKeyset, adLockOptimistic
    
    'If rstSettings.EOF = False Then
    If Not (rstSettings.EOF And rstSettings.BOF) Then
    ' MUCP-159 - End
    
        ' Window already has a setting in the database
        Me.WindowShowCmd = FNullField(rstSettings!Window_ShowCmd)
        Me.WindowFlags = FNullField(rstSettings!Window_Flags)
        Me.WindowMinX = FNullField(rstSettings!Window_MinX)
        Me.WindowMinY = FNullField(rstSettings!Window_MinY)
        
        Me.WindowLeft = FNullField(rstSettings!Window_Left)
        Me.WindowTop = FNullField(rstSettings!Window_Top)
        Me.WindowWidth = FNullField(rstSettings!Window_Width)
        Me.WindowHeight = FNullField(rstSettings!Window_Height)
        Me.WindowState = FNullField(rstSettings!Window_State)
        
        SelectSetting = QueryResultSuccessful
        
    Else
        ' Window does not have a setting in the database
        SelectSetting = QueryResultNoRecord
    End If
    
    Set rstSettings = Nothing
    
    Exit Function
    
Error_Handler:
    
    SelectSetting = QueryResultError
End Function

Public Function FNullField(ByRef Data As Variant) As Variant
    Dim strDataType As String
    
    If FIsEmpty(Data) Then
        strDataType = TypeName(Data)
        
        If InStr(1, strDataType, "Byte") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Integer") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Long") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Single") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "String") > 0 Then
            Data = ""
        ElseIf InStr(1, strDataType, "Double") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Currency") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Decimal") > 0 Then
            Data = 0
        ElseIf InStr(1, strDataType, "Date") > 0 Then
        ElseIf Trim(strDataType) = "Null" Then
            Data = ""
        ElseIf InStr(1, strDataType, "Field") > 0 Then
            Select Case Data.Type
                Case 20, 14, 5, 6, 3, 205, 201, 131, 2          'Numeric Field Types
                    Data = 0
                Case 16, 21, 19, 18, 17, 4                      'Numeric Field Types
                    Data = 0
                Case 11                                         ' Boolean
                    Data = False
                Case 129, 204, 200, 202, 203, 130               ' String
                    Data = ""
                Case 7, 133, 134                                ' Date
                    
            End Select
        End If
    End If
    FNullField = Data
End Function

Public Function FIsEmpty(ByVal Data As Variant) As Boolean
    Dim strDummy As String
    
    FIsEmpty = False
    
    If IsObject(Data) And Not TypeName(Data) = "Field" Then ' Check if Variable Passed is an Object Variable
        FIsEmpty = True
        
    ElseIf IsEmpty(Data) Then ' Check if Variable Passed is Not Initializaed
        FIsEmpty = True
        
    ElseIf IsNull(Data) Then ' Check if Variable Passed Contains Invalid Data
        FIsEmpty = True

    Else
        If IsArray(Data) Then
            If UBound(Data) = 0 And (Data(0) = "" Or IsEmpty(Data(0))) Then
                FIsEmpty = True
            End If
        Else
            strDummy = CStr(Data)
            
            If Trim(strDummy) = "" Then
                FIsEmpty = True
            End If
        End If
    End If
End Function



