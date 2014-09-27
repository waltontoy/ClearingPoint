Attribute VB_Name = "modGlobal"
Option Explicit

Public Const TBL_DATA_ROOT = "DATA" & "_" & "NCTS"

Public Const TBL_DATA_LEVEL1 = TBL_DATA_ROOT & "_" & "BERICHT"

Public Const TBL_DATA_LEVEL1_1 = TBL_DATA_LEVEL1 & "_" & "HOOFDING"
Public Const TBL_DATA_LEVEL1_2 = TBL_DATA_LEVEL1 & "_" & "VERVOER"
Public Const TBL_DATA_LEVEL1_3 = TBL_DATA_LEVEL1 & "_" & "DOUANEKANTOOR"
Public Const TBL_DATA_LEVEL1_4 = TBL_DATA_LEVEL1 & "_" & "HANDELAAR"
'
Public Const TBL_DATA_LEVEL1_2_1 = TBL_DATA_LEVEL1_2 & "_" & "INCIDENT"
Public Const TBL_DATA_LEVEL1_2_2 = TBL_DATA_LEVEL1_2 & "_" & "OVERLADING"
Public Const TBL_DATA_LEVEL1_2_3 = TBL_DATA_LEVEL1_2 & "_" & "VERZEGELING_INFO"
Public Const TBL_DATA_LEVEL1_2_4 = TBL_DATA_LEVEL1_2 & "_" & "CONTROLE"
'
Public Const TBL_DATA_LEVEL1_2_2_1 = TBL_DATA_LEVEL1_2_2 & "_" & "CONTAINER"
Public Const TBL_DATA_LEVEL1_2_3_1 = TBL_DATA_LEVEL1_2_3 & "_" & "ID"



    Private Declare Function GetThreadLocale Lib "kernel32" () As Long
    Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

    Private Const LOCALE_SDECIMAL = &HE                 ' Decimal separator
    Private Const LOCALE_STHOUSAND = &HF                ' Thousand separator
    Private Const LOCALE_SMONDECIMALSEP = &H16          ' Monetary decimal separator
    Private Const LOCALE_SMONTHOUSANDSEP = &H17         ' Monetary thousand separator
    Private Const LOCALE_IDEFAULTANSICODEPAGE = &H1004&

    Dim mintSystemLanguage As Integer
    Dim mintLastUserLanguage As Integer

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, _
     ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Global gstrMonDecimalSep As String
Global ResourceHandler As Long
Global cLanguage  As String
Public gstrLineValues() As String         ' Added September 25, 2000

'''''Public Function StripNullTerminator(ByVal sCP As String) As String
'''''    Dim posNull As Long
'''''
'''''    posNull = InStr(sCP, Chr$(0))
'''''    StripNullTerminator = Left$(sCP, posNull - 1)
'''''End Function
'''''
'''''Public Function NoBackSlash(ByVal cString As String) As String
'''''    Do While Right(cString, 1) = "\"
'''''        cString = Left(cString, Len(cString) - 1)
'''''    Loop
'''''
'''''    NoBackSlash = cString
'''''End Function

Public Function Translate(ByVal StringToTranslate As Variant) As String
    Dim cTranslated As String * 520
    
    If IsNumeric(StringToTranslate) Then
        LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
        cTranslated = StripNullTerminator(cTranslated)
        Translate = RTrim(cTranslated)
    Else
        Translate = StringToTranslate
    End If
End Function

Public Function G_InitializeMenus(ByRef tbrIE07 As Object, ByVal AppName As String, ByVal AppPath As String)

    ' translate menus, toolbars
    Dim lngToolCtr As Long
    Dim lngToolbarCtr As Long
    Dim lngToolItemCtr As Long
    Dim lngToolItemCtr2 As Long
    Dim strToolName As String
    
    'mvarAppTitle = "ClearingPoint"
    ' InitializeResource AppName, AppPath 'mvarAppTitle
    
    
    ' others
    For lngToolCtr = 1 To tbrIE07.Tools.Count
        tbrIE07.Tools(lngToolCtr).ToolTipText = Replace(Translate(tbrIE07.Tools(lngToolCtr).Name), "&", "")
        tbrIE07.Tools(lngToolCtr).ToolTipText = Replace(tbrIE07.Tools(lngToolCtr).ToolTipText, "...", "")
    Next lngToolCtr
    
    For lngToolbarCtr = 1 To tbrIE07.ToolBars.Count
    'For lngToolbarCtr = 2 To 1 Step -1
    
        'tbrIE07.ToolBars(1).Style
        For lngToolCtr = 1 To tbrIE07.ToolBars(lngToolbarCtr).Tools.Count
        
            strToolName = Translate(tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Name)
            tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Name = strToolName
            'tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).ToolTipText = strToolName
            
            If (tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Type = ssTypeMenu) Then
                For lngToolItemCtr = 1 To tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools.Count
                
                    strToolName = Translate(tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Name)
                    tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Name = strToolName
                    'tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).ToolTipText = strToolName
                    
                    If (tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Type = ssTypeMenu) Then
                        For lngToolItemCtr2 = 1 To tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Menu.Tools.Count
                        
                            strToolName = Translate(tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Menu.Tools(lngToolItemCtr2).Name)
                            tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Menu.Tools(lngToolItemCtr2).Name = strToolName
                            'tbrIE07.ToolBars(lngToolbarCtr).Tools(lngToolCtr).Menu.Tools(lngToolItemCtr).Menu.Tools(lngToolItemCtr2).ToolTipText = strToolName
                        
                        Next lngToolItemCtr2
                    End If
                
                Next lngToolItemCtr
                
                
            End If
            
        Next lngToolCtr
        
    Next lngToolbarCtr

End Function

Private Sub InitializeResource(ByVal strAppTitle As String, ByVal AppPath As String)
    
    Dim strLastUserLanguage As String
    Dim blnFlag As Boolean
    
    strLastUserLanguage = GetSetting(strAppTitle, "Settings", "Last UserLanguage")
    mintSystemLanguage = GetSystemLanguage()
    
    blnFlag = ((strLastUserLanguage = "1") Or (strLastUserLanguage = "2") Or (strLastUserLanguage = "3"))
    
    If (blnFlag = True) Then
        
        mintLastUserLanguage = CInt(strLastUserLanguage)
        SetLanguageSettings mintLastUserLanguage, AppPath
    
    ElseIf (blnFlag = False) Then
        
        mintLastUserLanguage = 0
        SetLanguageSettings mintSystemLanguage, AppPath
    
    End If
    
End Sub

Private Sub SetLanguageSettings(ByVal intLanguageValue As Integer, ByVal AppPath As String)
    
    Dim strDLLFileName As String
    Dim strAppPath As String
    Dim strCodePage As String
    Dim lngLCID As Long
    
    ' Use user language settings only for the resource strings.
    Select Case intLanguageValue
        
        Case 1
            strDLLFileName = "\ResEnglish.dll"
        Case 2
            strDLLFileName = "\ResDutch.dll"
        Case 3
            strDLLFileName = "\ResFrench.dll"
    
    End Select
    
    ' Use system language settings still for the separators.
    lngLCID = GetThreadLocale()
    
    Select Case mintSystemLanguage
        
        Case 1
            SetLocaleInfo lngLCID, LOCALE_SDECIMAL, "." & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_STHOUSAND, "," & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_SMONDECIMALSEP, "." & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_SMONTHOUSANDSEP, "," & Chr$(0)
        
        Case 2
            SetLocaleInfo lngLCID, LOCALE_SDECIMAL, "," & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_STHOUSAND, "." & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_SMONDECIMALSEP, "," & Chr$(0)
            SetLocaleInfo lngLCID, LOCALE_SMONTHOUSANDSEP, "." & Chr$(0)
        
        Case 3
            SetLocaleInfo lngLCID, LOCALE_SDECIMAL, "," & Chr$(0)           ' These SetLocaleInfo calls
            SetLocaleInfo lngLCID, LOCALE_STHOUSAND, "." & Chr$(0)          ' force the default values
            SetLocaleInfo lngLCID, LOCALE_SMONDECIMALSEP, "," & Chr$(0)     ' upon the current locale
            SetLocaleInfo lngLCID, LOCALE_SMONTHOUSANDSEP, "." & Chr$(0)    ' in case of user override.  [Andrei]
    
    End Select
    
    strCodePage = Space$(16)
    
    ' ********** Added August 10, 2000 **********
    GetLocaleInfo lngLCID, LOCALE_SMONDECIMALSEP, strCodePage, Len(strCodePage)
    strCodePage = StripNullTerminator(strCodePage)
    gstrMonDecimalSep = strCodePage
    ' ********** End Add ************************
    
    strAppPath = NoBackSlash(AppPath)
    
    Dim blnFlag As Boolean
    
    Do While True ' infinite loop
        
        blnFlag = Len(Dir(strAppPath & strDLLFileName))
        
        If (blnFlag = True) Then
            
            ResourceHandler = LoadLibrary(strAppPath & strDLLFileName)
            Exit Do
        
        ElseIf (blnFlag = False) Then
            
            strAppPath = InputBox("Cannot find the file " & Mid$(strDLLFileName, 2) & " in " & strAppPath, "Missing DLL file.", strAppPath)
            'If strAppPath = "" Then End
        
        End If
        
    Loop
    
    cLanguage = Translate(746)

End Sub


Private Function GetSystemLanguage() As Integer

    ' ********** Repositioned February 28, 2001 **********
    ' ********** Language setting must be determined before performing Syslink via SendKeys because
    ' ********** menu keystrokes differ among the three languages.
    Dim lngLCID As Long
    Dim strLCID As String

    lngLCID = GetThreadLocale()             ' Get current locale
    strLCID = Hex$(Trim$(CStr(lngLCID)))    ' Convert to hexadecimal

    Select Case UCase$(strLCID)

        ' English
        Case "409", "809", "C09", "1009", "1409", "1809", "1C09", "2009", "2409"

            GetSystemLanguage = 1

        ' French
        Case "40C", "80C", "C0C", "100C", "140C"

            GetSystemLanguage = 3

        ' Dutch
        Case Else

            GetSystemLanguage = 2

    End Select

End Function

Public Sub LoadResStrings(ByRef frmFormToLoad As Form, Optional ByVal blnUseTag As Boolean)
    Dim ctlControlToLoad As Control
    Dim Tool As SSTool
    Dim Tool2 As SSTool
    Dim Tool3 As SSTool
    Dim Tool4 As SSTool
    
    Dim strTypeName As String
    Dim strToolTipText As String
    
    Dim intCtrlCount As Integer
    Dim intToolCount2 As Integer
    Dim intToolCount3 As Integer
    Dim intToolCount4 As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim nCtr As Integer
        
    On Error Resume Next
    
    If blnUseTag Then
        If frmFormToLoad.Tag <> "" Then
            frmFormToLoad.Caption = Translate(frmFormToLoad.Tag)
        End If
    Else
        frmFormToLoad.Caption = Translate(frmFormToLoad.Caption)
    End If
    
    For Each ctlControlToLoad In frmFormToLoad.Controls
        strTypeName = LCase(TypeName(ctlControlToLoad))
        If TypeOf ctlControlToLoad Is SSActiveToolBars Then
        End If
        
        Select Case strTypeName
            Case "ssactivetoolbars", "tlbrMain"
                For nCtr = 1 To ctlControlToLoad.ToolBars.Count  '2
                
                    intCtrlCount = ctlControlToLoad.ToolBars(nCtr).Tools.Count
                    
                    For i = 1 To intCtrlCount
                        Set Tool = ctlControlToLoad.ToolBars(nCtr).Tools(i)
                        Tool.Name = Translate(Tool.Name)
                        
                        strToolTipText = ""
                        strToolTipText = Translate(Tool.ToolTipText)    ' Translate(ctlControlToLoad.Tools(i).ToolTipText)
                        strToolTipText = NoAmpersandEllipse(strToolTipText)
                        If UCase(Tool.ID) <> "SEPARATOR" Then
                            ctlControlToLoad.Tools(Tool.ID).ToolTipText = strToolTipText
                        End If
                        
                        'On Error Resume Next
                        If Tool.Type = ssTypeMenu Then
                            intToolCount2 = Tool.Menu.Tools.Count
                        Else
                            intToolCount2 = 0
                        End If
                        'If Err.Number <> 40006 Then
                        '    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                        'End If
                        'Err.Clear
                        'On Error GoTo 0
                        
                        
                        
                        For j = 1 To intToolCount2
                            Set Tool2 = Tool.Menu.Tools(j)
                            Tool2.Name = Translate(Tool2.Name)
                            
                            strToolTipText = ""
                            strToolTipText = Translate(Tool2.ToolTipText)
                            strToolTipText = NoAmpersandEllipse(strToolTipText)
                            
                            If Len(strToolTipText) Then
                                ctlControlToLoad.Tools(Tool2.ID).ToolTipText = strToolTipText
                            End If
                            
                            If Tool2.Type = ssTypeMenu Then
                                intToolCount3 = Tool2.Menu.Tools.Count
                            Else
                                intToolCount3 = 0
                            End If
                            
                            For k = 1 To intToolCount3
                                Set Tool3 = Tool2.Menu.Tools(k)
                                Tool3.Name = Translate(Tool3.Name)
                                
                                strToolTipText = ""
                                strToolTipText = Translate(Tool3.ToolTipText)
                                strToolTipText = NoAmpersandEllipse(strToolTipText)
                                
                                If Len(strToolTipText) Then
                                    ctlControlToLoad.Tools(Tool3.ID).ToolTipText = strToolTipText
                                End If
                                
                                If Tool3.Type = ssTypeMenu Then
                                    intToolCount4 = Tool3.Menu.Tools.Count
                                Else
                                    intToolCount4 = 0
                                End If

                            
                                For l = 1 To intToolCount4
                                    Set Tool4 = Tool3.Menu.Tools(l)
                                    Tool4.Name = Translate(Tool4.Name)
                                    
                                    strToolTipText = ""
                                    strToolTipText = Translate(Tool4.ToolTipText)
                                    strToolTipText = NoAmpersandEllipse(strToolTipText)
                                    
                                    If Len(strToolTipText) Then
                                        ctlControlToLoad.Tools(Tool4.ID).ToolTipText = strToolTipText
                                    End If
                                        
                                Next
                            
                            Next
                        Next
                    Next
                Next
            Case "sstab"
                intCtrlCount = ctlControlToLoad.Tabs
                                
                For i = 0 To intCtrlCount - 1
                    ctlControlToLoad.TabCaption(i) = Translate(ctlControlToLoad.TabCaption(i))
                Next
            Case "tabstrip"
                intCtrlCount = ctlControlToLoad.Tabs.Count
                'edited by alg
                'For i = 0 To intCtrlCount  --> why i=0???...by alg
                For i = 1 To intCtrlCount
                    ctlControlToLoad.Tabs(i).Caption = Translate(ctlControlToLoad.Tabs(i).Caption)
                Next
            Case "label", "optionbutton", "frame", "commandbutton", "sscommand", "sspanel", "checkbox"
                If blnUseTag Then
                    If ctlControlToLoad.Tag <> "" Then
                        ctlControlToLoad.Caption = Translate(ctlControlToLoad.Tag)
                    End If
                Else
                    ctlControlToLoad.Caption = Translate(ctlControlToLoad.Caption)
                End If
        End Select
    Next
End Sub

Public Function NoAmpersandEllipse(ByVal cText As String) As String
    Dim i As Integer
    
    i = InStr(1, cText, "&")
    
    If i > 0 Then
        cText = Mid(cText, 1, i - 1) + Mid(cText, i + 1)
    End If
    
    If Right(cText, 3) = "..." Then
        cText = Left(cText, Len(cText) - 3)
    End If
    
    NoAmpersandEllipse = cText
End Function
