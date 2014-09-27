VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    ShowCmd As Long
    MinPosition As POINTAPI
    MaxPosition As POINTAPI
    WindowPositionSize As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Sub LoadWindowSettings(ByRef ADOConnection As ADODB.Connection, _
                                ByVal UserID As Long, _
                                ByVal WindowKey As String, _
                                ByRef Window As Object)

    Dim clsSettings As CWindows
    Dim enuSelectSetting As QueryResultConstants
    Dim enuWindowSetting As QueryResultConstants
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim udtWindowPlacement As WINDOWPLACEMENT
    
    Set clsSettings = New CWindows
    
    enuSelectSetting = clsSettings.SelectSetting(ADOConnection, UserID, WindowKey)
    
    ' Check if window setting already exist
    Select Case enuSelectSetting
        Case QueryResultNoRecord
            ' Window setting does not exist, add new setting with default values
            LoadDefaultWindowSettings Window, udtWindowPlacement
            
            enuWindowSetting = clsSettings.AddSetting(ADOConnection, UserID, WindowKey, _
                                                      udtWindowPlacement.ShowCmd, _
                                                      udtWindowPlacement.Flags, _
                                                      udtWindowPlacement.MinPosition.X, _
                                                      udtWindowPlacement.MinPosition.Y, _
                                                      udtWindowPlacement.WindowPositionSize.Bottom, _
                                                      udtWindowPlacement.WindowPositionSize.Right, _
                                                      udtWindowPlacement.WindowPositionSize.Top, _
                                                      udtWindowPlacement.WindowPositionSize.Left, _
                                                      udtWindowPlacement.ShowCmd) '<- ShowCmd is the same with window state property in API
            
            If (enuWindowSetting = QueryResultError) Then
                GoTo Error_Handler
            ElseIf (enuWindowSetting = QueryResultSuccessful) Then
                clsSettings.SelectSetting ADOConnection, UserID, WindowKey
            End If
        
        Case QueryResultError
            GoTo Error_Handler
        
        Case QueryResultSuccessful
            
            '<<< dandan 021307
            'No solution yet in computing the left and top property when the window is outside the screen
            'If Abs(clsSettings.WindowLeft) > Screen.Width Or _
                ((Screen.Height - clsSettings.WindowTop) / Screen.TwipsPerPixelY) < GetSystemMetrics(SM_CYCAPTION / 2) Or _
                Abs(clsSettings.WindowTop / Screen.TwipsPerPixelY) < GetSystemMetrics(SM_CYCAPTION / 2) Then
            '        clsSettings.WindowLeft = (Screen.Width - Window.Width) / 2
            '        clsSettings.WindowTop = (Screen.Height - Window.Height) / 2
            'End If
    End Select
    
    ' Initialize udtWindowPlacement from the registry.
    With udtWindowPlacement
        .Length = Len(udtWindowPlacement)
        
        .Flags = clsSettings.WindowFlags
        .ShowCmd = clsSettings.WindowShowCmd
        
        .MinPosition.X = clsSettings.WindowMinX
        .MinPosition.Y = clsSettings.WindowMinY
        
        .WindowPositionSize.Left = clsSettings.WindowLeft
        .WindowPositionSize.Top = clsSettings.WindowTop
        .WindowPositionSize.Right = clsSettings.WindowWidth
        .WindowPositionSize.Bottom = clsSettings.WindowHeight
    End With
    
    ' Position the form.
    SetWindowPlacement Window.hwnd, udtWindowPlacement
    
    Set clsSettings = Nothing
    
    Exit Sub
    
Error_Handler:
    
    MsgBox "Error has occured while trying to update the database. Please contact technical support.", vbInformation + vbOKOnly, "Window Setting Error"
    
End Sub

Public Sub SaveWindowSettings(ByRef ADOConnection As ADODB.Connection, _
                                ByVal UserID As Long, _
                                ByVal WindowKey As String, _
                                ByVal Window As Object)

    Dim clsSettings As CWindows
    Dim enuWindowSetting As QueryResultConstants
    
    Dim udtWindowPlacement As WINDOWPLACEMENT
    
    Dim lngWindowShowCmd As Long
    Dim lngWindowFlags As Long
    Dim lngWindowMinX As Long
    Dim lngWindowMinY As Long
    
    Dim lngWindowHeight As Long
    Dim lngWindowWidth As Long
    Dim lngWindowLeft As Long
    Dim lngWindowTop As Long
    
    Dim lngWindowState As Long
    
    Set clsSettings = New CWindows
    
    lngWindowState = Window.WindowState
    
    ' Select window setting
    enuWindowSetting = clsSettings.SelectSetting(ADOConnection, UserID, WindowKey)
    
    ' Retreive new window placement properties
    udtWindowPlacement.Length = Len(udtWindowPlacement)
    
    GetWindowPlacement Window.hwnd, udtWindowPlacement
    
    ' Get window properties
    If (lngWindowState <> vbNormal) Then
        Select Case enuWindowSetting
            Case QueryResultSuccessful
                ' Set values retrieved from the database
                lngWindowShowCmd = udtWindowPlacement.ShowCmd
                lngWindowFlags = udtWindowPlacement.Flags
                lngWindowMinX = udtWindowPlacement.MinPosition.X
                lngWindowMinY = udtWindowPlacement.MinPosition.Y
                
                lngWindowHeight = udtWindowPlacement.WindowPositionSize.Bottom    'Window.Height
                lngWindowWidth = udtWindowPlacement.WindowPositionSize.Right      'Window.Width
                lngWindowLeft = udtWindowPlacement.WindowPositionSize.Left        'Window.Left
                lngWindowTop = udtWindowPlacement.WindowPositionSize.Top          'Window.Top
                
            Case QueryResultNoRecord
                ' Set default value
                lngWindowShowCmd = 1
                lngWindowFlags = SW_RESTORE
                lngWindowMinX = 100
                lngWindowMinY = 100
                
                lngWindowHeight = udtWindowPlacement.WindowPositionSize.Bottom    'Window.Height
                lngWindowWidth = udtWindowPlacement.WindowPositionSize.Right      'Window.Width
                lngWindowLeft = udtWindowPlacement.WindowPositionSize.Left        'Window.Left
                lngWindowTop = udtWindowPlacement.WindowPositionSize.Top          'Window.Top
                
            Case QueryResultError
                GoTo Error_Handler
                
        End Select
        
    Else
        
        ' Set values retrieved from the database
        lngWindowShowCmd = udtWindowPlacement.ShowCmd
        lngWindowFlags = udtWindowPlacement.Flags
        lngWindowMinX = udtWindowPlacement.MinPosition.X
        lngWindowMinY = udtWindowPlacement.MinPosition.Y
        
        lngWindowHeight = udtWindowPlacement.WindowPositionSize.Bottom    'Window.Height
        lngWindowWidth = udtWindowPlacement.WindowPositionSize.Right      'Window.Width
        lngWindowLeft = udtWindowPlacement.WindowPositionSize.Left        'Window.Left
        lngWindowTop = udtWindowPlacement.WindowPositionSize.Top          'Window.Top
        
    End If
    
    Select Case enuWindowSetting
        Case QueryResultSuccessful
            ' Window setting exists, update record in the database
            If (clsSettings.UpdateSetting(ADOConnection, UserID, WindowKey, lngWindowShowCmd, lngWindowFlags, lngWindowMinX, lngWindowMinY, lngWindowHeight, lngWindowWidth, lngWindowTop, lngWindowLeft, lngWindowState) = QueryResultError) Then
                GoTo Error_Handler
            End If
        
        Case QueryResultNoRecord
            ' Window setting does not exist, add new window setting
            LoadDefaultWindowSettings Window, udtWindowPlacement
            
            If (clsSettings.AddSetting(ADOConnection, UserID, WindowKey, lngWindowShowCmd, lngWindowFlags, lngWindowMinX, lngWindowMinY, lngWindowHeight, lngWindowWidth, lngWindowTop, lngWindowLeft, lngWindowState) = QueryResultError) Then
                GoTo Error_Handler
            End If
        
        Case QueryResultError
            GoTo Error_Handler
    End Select
        
    Set clsSettings = Nothing
    
    Exit Sub
    
Error_Handler:
    
    MsgBox "Error has occured while trying to update the database. Please contact technical support.", vbInformation, "Window Setting Error"
End Sub

Private Sub LoadDefaultWindowSettings(ByRef Window As Object, ByRef DefaultWindowPlacement As WINDOWPLACEMENT)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    Dim lRet As Long
    Dim apiRECT As RECT

    lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)

    lngLeft = apiRECT.Left
    lngTop = apiRECT.Top
    lngWidth = apiRECT.Right
    lngHeight = apiRECT.Bottom
    
    With Window
        .Left = lngLeft    ' (Screen.Width - Window.Width) / 2
        .Top = lngTop      ' (Screen.Height - Window.Height) / 2
        .WindowState = vbNormal
    End With
    
    
    With DefaultWindowPlacement
        .Length = Len(DefaultWindowPlacement)
        
        .Flags = SW_RESTORE
        .ShowCmd = 1
        .MinPosition.X = 100
        .MinPosition.Y = 100
        
        ' For center width
        If (Window.Width < Screen.Width) Then
            .WindowPositionSize.Left = (lngWidth / 2) - ((Window.Width / 2) / TwipsPerPixelX())
        Else
            .WindowPositionSize.Left = lngLeft
        End If
        
        ' For center height
        If (Window.Height < Screen.Height) Then
            .WindowPositionSize.Top = (lngHeight / 2) - ((Window.Height / 2) / TwipsPerPixelY())
        Else
            .WindowPositionSize.Top = lngTop
        End If
        
        ' For window width
        If (Window.Width > Screen.Width) Then
            .WindowPositionSize.Right = lngWidth
        Else
            .WindowPositionSize.Right = (Window.Width / TwipsPerPixelX()) + .WindowPositionSize.Left
        End If
        
        ' For window height
        If (Window.Height > Screen.Height) Then
            .WindowPositionSize.Bottom = lngHeight
        Else
            .WindowPositionSize.Bottom = (Window.Height / TwipsPerPixelY()) + .WindowPositionSize.Top
        End If
    End With
    
End Sub
