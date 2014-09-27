Attribute VB_Name = "modFunctions"
Option Explicit

Public Const LOCALE_USER_DEFAULT = &H400

Public Const LOCALE_SDECIMAL = &HE                 ' Decimal separator
Public Const LOCALE_STHOUSAND = &HF                ' Thousand separator
Public Const LOCALE_SMONDECIMALSEP = &H16          ' Monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17         ' Monetary thousand separator
Public Const LOCALE_NOUSEROVERRIDE = &H80000000    ' Do not use user overrides!
Public Const LOCALE_SENGCOUNTRY = &H1002 '  English name of country
Public Const LOCALE_SENGLANGUAGE = &H1001  '  English name of language
Public Const LOCALE_SNATIVELANGNAME = &H4  '  native name of language
Public Const LOCALE_SNATIVECTRYNAME = &H8  '  native name of country

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public g_blnFromEditViewForm As Boolean 'Edwin Nov09

'  Used in GetFileLastAccessedDate function -joy
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

'>> Get UNC path
Const VER_PLATFORM_WIN32s = 0          'Win32s on Windows 3.1
Const VER_PLATFORM_WIN32_WINDOWS = 1   'Win32 on Windows 95
Const VER_PLATFORM_WIN32_NT = 2        'Win32 on Windows NT

Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
  
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
   Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey _
   As String, ByVal ulOptions As Long, ByVal samDesired _
   As Long, phkResult As Long) As Long
   
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias _
  "RegQueryValueA" (ByVal HKey As Long, ByVal lpSubKey As _
   String, ByVal lpValue As String, lpcbValue As Long) As Long

' Note that if you declare lpData as String, then it is
' necessary to pass it with ByVal
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
  Alias "RegQueryValueExA" (ByVal HKey As Long, _
   ByVal lpValueName As String, ByVal lpReserved As Long, _
   lpType As Long, lpData As Any, lpcbData As Long) As Long
  
   Private Declare Function RegEnumKey Lib "advapi32.dll" _
   Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex _
   As Long, ByVal lpName As String, ByVal cbName As Long) _
   As Long

'*************************************************************************************************************

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, _
    ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long


Public Type CInterface
    IApplication As Object
    IGrid As Object             'jsgxMain           Janus grid control
    IButtonBar As Object        'bbrMain            Shortcut bar control
    ICommandBar As Object       'cbrMain            Command bar control
    IFind As Object             'cbrFind            Comamnd bar control
    IFindBGround As Object      'picFind            Picture box control for Default find
    ICustomFind As Object       'picCustomFind      Picture box control for Custom find
    IItemlist As Object         'picItemlist        Picture box control for the itemlist controls
    ISplitter As Object         'picNavSplitter
    ISplitterImage As Object    'imgNavigation
    IGridHeader As Object       'sccHeader          Shortcut caption control - Grid header where caption of selected node is displayed
    IReadingPane As Object      'picReadingPane     Picture box control for the reading pane
    IGridSplitter As Object     'picGridSplitter    Picture box control - splitter between grid and reading pane
    ISplash As Object           'imgSplash          Image control for the Splash screen image
    IStatusbar As Object        'sbrMain            Status bar control
    ITimer As Object            'tmrCubelib2003RefreshTimer
    IGridImage As Object        'imgGrid            Image control for the selected node's image
    INoView As Object           'sspnlNoView
    ITreeIcons As Object        'ImageList2
    IBBarIcons As Object        'imgListButtonBar
    ILicense As Object          'lfpLicensing
End Type

Public g_typInterface As CInterface


Public Enum AQChar
    Apostrophe = 1
    Quotation = 2
    BothAQ = 3
End Enum

Public Const ID_FILE = 1
Public Const ID_FILE_EXIT = 101

Public Const ID_EDIT = 2

Public Const ID_VIEW = 3

Public Const ID_VIEW_ARRANGEBY = 311
Public Const ID_VIEW_ARRANGEBY_VIEWOPTIONS = 309
Public Const ID_VIEW_ARRANGEBY_CURRENTVIEW = 31102
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW = 308
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_SHOWFIELDS = 30801
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_GROUPBY = 30802
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_SORT = 30803
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_FILTER = 30204
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_OTHERSETTINGS = 30805
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_AUTOFORMAT = 30806
Public Const ID_VIEW_ARRANGEBY_CUSTOMIZECURRENTVIEW_FORMATCOLUMNS = 30807

Public Const ID_VIEW_NAVIGATIONPANE = 302

Public Const ID_VIEW_READINGPANE = 303
Public Const ID_VIEW_READINGPANE_BOTTOM = 30301
Public Const ID_VIEW_READINGPANE_RIGHT = 30302
Public Const ID_VIEW_READINGPANE_OFF = 30303

Public Const ID_VIEW_GROUPBYBOX = 304

Public Const ID_VIEW_COLLAPSEEXPAND = 305
Public Const ID_VIEW_COLLAPSEGROUP = 30501
Public Const ID_VIEW_EXPANDGROUP = 30502
Public Const ID_VIEW_COLLAPSEALL = 30503
Public Const ID_VIEW_EXPANDALL = 30504

Public Const ID_VIEW_ITEMLIST = 306
Public Const ID_VIEW_ITEMLIST_CARDVIEW = 306001
Public Const ID_VIEW_ITEMLIST_LISTVIEW = 306002
Public Const ID_VIEW_ITEMLIST_GRIDLINES = 306003
Public Const ID_VIEW_ITEMLIST_ODDEVENCOLOR = 306004

Public Const ID_VIEW_STATUSBAR = 307

Public Const ID_VIEW_VIEWOPTIONS = 309
Public Const ID_VIEW_REFRESH = 310


Public Const ID_REPORTS = 4

Public Const ID_TOOLS = 5
Public Const ID_TOOLS_OPTIONS = 501
Public Const ID_TOOLS_FINDTOOLS = 502
Public Const ID_TOOLS_FIND = 50201
Public Const ID_TOOLS_ADVANCEDFIND = 50202

Public Const ID_HELP = 6
Public Const ID_HELP_LICENSEE = 601
Public Const ID_HELP_ABOUT = 602
'Public Const ID_HELP_PRODACTIVATION = 603

Public Const ID_CONTEXT_ADDNEWFAVEFOLDER = 951
Public Const ID_CONTEXT_RENAMEFAVEFOLDER = 952
Public Const ID_CONTEXT_REMOVEFAVEFOLDER = 953
Public Const KEY_ENCRYPT = ""

Public Const ID_FIND_LOOK_FOR = -90101
Public Const ID_FIND_SEARCH_IN = -90102
Public Const ID_FIND_SEARCH_BOX = -90103
Public Const ID_FIND_FIND_NOW = -90104
Public Const ID_FIND_CLEAR = -90105
Public Const ID_FIND_SPACE = -90106
Public Const ID_FIND_OPTION = -90107
Public Const ID_FIND_OPTION_ADVANCED_FIND = -9010701
Public Const ID_FIND_X = -90108

Public Const ID_GO_FOLDER = 80182

Public g_clsLog As CAppLogFile
Public g_clsLogError As CAppLogFile
Public g_blnTraceFileOn As Boolean
Public g_lngTopNode As Long

Public g_strSessionCode As String

Public Function AddSpaceZero(CharToAdd As String, FieldValue As String, FixedLength As Integer, blnBefore As Boolean)
    '----> to add extra characters if the field requires fixed length
    Do While Not Len(FieldValue) >= FixedLength
        If blnBefore Then
            FieldValue = CharToAdd & FieldValue
        Else
            FieldValue = FieldValue & CharToAdd
        End If
    Loop
    
    AddSpaceZero = FieldValue

End Function


Public Function WindowsTempPath() As String
'--->Returns the Windows temporary folder

    Dim strTempPath As String
    
    '--->Create a buffer
    strTempPath = String(200, Chr(0))
    
    '--->Get the temporary path
    Call GetTempPath(200, strTempPath)
    
    '--->Strip the rest of the buffer
    strTempPath = Left(strTempPath, InStr(strTempPath, Chr(0)) - 1)
    
    
    WindowsTempPath = strTempPath
End Function

Public Function PicPath(ByVal ID As String, ByVal ADORecordset As ADODB.Recordset) As String
    Dim rstIcon As ADODB.Recordset
    Dim lngImgHandle As Long
    Dim strImageLoc As String
    Dim bytImage() As Byte
    
    
    Set rstIcon = ADORecordset.Clone
    
    If (rstIcon.RecordCount > 0) Then
        rstIcon.MoveFirst
        rstIcon.Find "FC_ID = " & ID, , adSearchForward, 0
        
        If (rstIcon.EOF = False) Then
            strImageLoc = WindowsTempPath & "FC" & Trim(rstIcon![FC_ID]) & ".img"
            
            If (Len(Dir$(strImageLoc)) > 0) Then
                Kill strImageLoc
            End If
            
            
            If (IsNull(rstIcon![FC_Icon]) = False) Then
                
                ' Temporarily save the image to a variable
                bytImage = rstIcon.Fields("FC_Icon").GetChunk(LenB(rstIcon!FC_Icon))
                
                ' Create image file
                lngImgHandle = FreeFile()
                Open strImageLoc For Binary As #lngImgHandle
                Put #lngImgHandle, , bytImage()
                Close #lngImgHandle
                
                PicPath = strImageLoc
                
            Else
                
                PicPath = ""
                
            End If
        End If
    End If
    
    ADORecordsetClose rstIcon
    
End Function

Public Function AQ(ByVal Field As String, Optional CharToFormat As AQChar = BothAQ) As String
    Dim intFieldPosition As Integer
    Dim blnDone As Boolean
    Dim intLength As Integer
    Dim arrLink
    Dim intLinkCtr As Integer
    Dim strChain As String
    Dim intLoopCtr As Integer
    
    ReDim arrLink(0)
    blnDone = False
    intLinkCtr = 0
    strChain = ""
    
    Do While blnDone = False
        intLength = Len(Field)
        
        If InStr(1, Field, Chr(39)) > 0 And CharToFormat <> Quotation Then   '-----> Apostrophe
            intFieldPosition = InStr(1, Field, Chr(39))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            arrLink(intLinkCtr - 1) = Left(Field, intFieldPosition) & Chr(39)
            If intFieldPosition <> Len(Field) Then
                Field = Mid(Field, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        ElseIf InStr(1, Field, Chr(34)) > 0 And CharToFormat <> Apostrophe Then  '-----> Quote
            intFieldPosition = InStr(1, Field, Chr(34))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            arrLink(intLinkCtr - 1) = Left(Field, intFieldPosition) & Chr(34)
            If intFieldPosition <> Len(Field) Then
                Field = Mid(Field, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        Else
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            arrLink(intLinkCtr - 1) = Field
            blnDone = True
        End If
    Loop
                        
    For intLoopCtr = 0 To (UBound(arrLink) - 1)
        If Trim(strChain) = "" Then
            strChain = CStr(arrLink(intLoopCtr))
        Else
            strChain = strChain & CStr(arrLink(intLoopCtr))
        End If
        
    Next
    AQ = Replace(strChain, " ", Chr(32))
    
    
End Function


Public Function GetMDBPath(ConnectionString As String) As String

    Dim lngStart As Long
    Dim lngDataSourceEnd As Long
    Dim lngPathEnd As Long
    Dim lngEnd As Long
    
    lngStart = InStr(1, ConnectionString, "Data Source=") + 12
    lngDataSourceEnd = InStr(lngStart, ConnectionString, ";")
    
    lngPathEnd = lngStart
    
    Do While lngPathEnd <= lngDataSourceEnd
        lngEnd = lngPathEnd
        lngPathEnd = InStr(lngPathEnd + 1, ConnectionString, "\")
        If lngPathEnd = 0 Or lngEnd >= lngDataSourceEnd Then
            Exit Do
        End If
    Loop
    
    If lngStart > 0 And lngEnd > 0 Then
        GetMDBPath = Mid(ConnectionString, lngStart, lngEnd - lngStart)
    Else
        GetMDBPath = App.Path
    End If

End Function

Public Function Translate(ByVal StringToTranslate As Variant, _
                            Optional ByVal ReturnStringToTranslate As Boolean = True) _
                            As String

    Dim cTranslated As String * 520
    
    If (IsNumeric(StringToTranslate) = True) And ResourceHandler <> 0 Then
        LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
        cTranslated = StripNullTerminator(cTranslated)
        Translate = RTrim$(cTranslated)
        
        If LenB(Translate) = 0 Then
            If ReturnStringToTranslate Then
                Translate = StringToTranslate
            End If
        End If
        
    Else
        Translate = StringToTranslate
    End If
    
End Function

Public Function GetLicenseValue(ActivationCode As String) As Long

    Dim lngArrCtr As Long
    Dim lngCtr As Long
    Dim lngValue As Long
    Dim lngTotalValue As Long
    Dim arrCodeGrp
    
    Dim strCode As String
    Dim strFeatureCode As String
    Dim strNumericCode As String
    
    On Error GoTo Error_Handler
    
    arrCodeGrp = Split(ActivationCode, "-")
    
    strFeatureCode = Trim(arrCodeGrp(0))
    
    For lngArrCtr = 2 To UBound(arrCodeGrp)
        
        If lngArrCtr = UBound(arrCodeGrp) Then
            lngTotalValue = lngTotalValue + Val(arrCodeGrp(lngArrCtr))
            Exit For
        End If
        
        lngValue = 1
        For lngCtr = 1 To Len(Trim(arrCodeGrp(lngArrCtr)))
            If IsNumeric(Mid(Trim(arrCodeGrp(lngArrCtr)), lngCtr, 1)) = True Then
                lngValue = lngValue * Val(Mid(Trim(arrCodeGrp(lngArrCtr)), lngCtr, 2))
                lngCtr = lngCtr + 1
            Else
                lngValue = lngValue * Asc(Mid(Trim(arrCodeGrp(lngArrCtr)), lngCtr, 1))
            End If
        Next
        lngTotalValue = lngTotalValue + lngValue
        
    Next
    
    For lngCtr = 1 To Len(Trim(arrCodeGrp(1)))
        If IsNumeric(Mid(Trim(arrCodeGrp(1)), lngCtr, 1)) = True Then
            lngTotalValue = lngTotalValue * Val(Mid(Trim(arrCodeGrp(1)), lngCtr, 1))
        Else
            lngTotalValue = lngTotalValue * ConvertBase(Mid(Trim(arrCodeGrp(1)), lngCtr, 1), 16, 10)
        End If
    Next
    
    GetLicenseValue = lngTotalValue
    
    Exit Function
    
Error_Handler:
    
    Select Case Err.Number
        Case 9
            GetLicenseValue = 0
        
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
            
    End Select
    
End Function


' This procedure is created to get the date of the computer where the database is located-joy 04/25/2006
Public Function GetFileLastAccessedDate(ByVal FilePath As String, ByRef AccessedDate As Date) As Boolean
    Dim lngFileHandle As Long
    Dim udtDateCreated As FILETIME
    Dim udtLastDateAccessed As FILETIME
    Dim udtLastDateWrite As FILETIME
    Dim udtLocalFileTime As FILETIME
    Dim udtSysTime As SYSTEMTIME
    
    GetFileLastAccessedDate = False
    
    If Len(Trim(FilePath)) > 0 Then
        If Len(Dir(FilePath)) > 0 Then
            'Get file handle
             lngFileHandle = CreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
             If lngFileHandle > 0 Then
                 'Get the fil's time
                 GetFileTime lngFileHandle, udtDateCreated, udtLastDateAccessed, udtLastDateWrite
                 'Convert the file time to the local file time
                 FileTimeToLocalFileTime udtLastDateAccessed, udtLocalFileTime
                 'Convert the file time to system file time
                 FileTimeToSystemTime udtLocalFileTime, udtSysTime
                 
                 AccessedDate = Format(udtSysTime.wMonth & "/" & udtSysTime.wDay & "/" & udtSysTime.wYear, g_typInterface.ILicense.DateFormat)
                 
                 'Close the file
                 CloseHandle lngFileHandle
                 
                 GetFileLastAccessedDate = True
             End If
        Else
            AccessedDate = Date
        End If
    Else
        AccessedDate = Date
    End If
End Function

Public Function ISLicenseExpiredOrClockTurnedBack(ByVal DateofServerPC As Date) As Boolean
    Dim dteLastUsedDate As Date
    Dim dteDateofServer As Date
    Dim dteExpireDate As Date
    
    dteLastUsedDate = Format(CDate(g_typInterface.ILicense.LastUsedDate), g_typInterface.ILicense.DateFormat)
    dteDateofServer = Format(DateofServerPC, vbUseSystem)
    
    If Not (g_typInterface.ILicense.ExpireMode = "N" And g_typInterface.ILicense.ExpireDateSoft = "0/0/0") Then
        dteExpireDate = Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat)
    Else
        ISLicenseExpiredOrClockTurnedBack = False
        Exit Function
    End If
    
    ISLicenseExpiredOrClockTurnedBack = False
    
    If dteExpireDate < dteDateofServer Then
        ISLicenseExpiredOrClockTurnedBack = True
    
    'rachelle 101706 for type N license with expiration date
'    ElseIf (dteLastUsedDate > dteDateofServer) And g_typInterface.ILicense.ExpireMode <> "N" Then
    ElseIf dteLastUsedDate > dteDateofServer Then
        ISLicenseExpiredOrClockTurnedBack = True

    End If
    
End Function

Public Function GetInfo(ByVal lInfo As Long) As String
    Dim Buffer As String, Ret As String
    Buffer = String$(256, 0)
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function
