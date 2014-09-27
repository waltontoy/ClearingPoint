VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'References:
'       1. CubeLibDBReg - Use for registry calls
'**********************************************************************************************************
'Referenced By:
'       1. ClearingPoint
'       2. ClearingPoint Task Scheduler
'**********************************************************************************************************
Option Explicit

'Rooth Path of Keys to check for OpLock Setting
Private Const REG_OPLOCKSDISABLED_ROOTPATH = "SYSTEM\CurrentControlSet\Services\MRXSmb\Parameters"
Private Const REG_ENABLEOPLOCKS_ROOTPATH_SERVER = "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
Private Const REG_ENABLEOPLOCKS_ROOTPATH_WORKSTATION = "SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters"

'Keys to check for OpLock Setting
Private Const REG_OPLOCKSDISABLED_REGKEY = "OplocksDisabled"
Private Const REG_ENABLEOPLOCKS_REGKEY = "EnableOplocks"
Private Const REG_SMB2_REGKEY = "SMB2"

'Private Const OS_WINDOWS_95 As String = "WINDOWS 95"
'Private Const OS_WINDOWS_98 As String = "WINDOWS 98"
'Private Const OS_WINDOWS_MILLENIUM As String = "WINDOWS MILLENIUM"
'Private Const OS_WINDOWS_NT_3_51 As String = "WINDOWS NT 3.51"
'Private Const OS_WINDOWS_NT_4_0 As String = "WINDOWS NT 4.0"
'Private Const OS_WINDOWS_2000 As String = "WINDOWS 2000"
'Private Const OS_WINDOWS_XP As String = "WINDOWS XP"
'Private Const OS_WINDOWS_SERVER_2003 As String = "WINDOWS SERVER 2003"
'Private Const OS_WINDOWS_VISTA As String = "WINDOWS VISTA"
'Private Const OS_WINDOWS_7 As String = "WINDOWS 7"                          'CSCLP-768
'Private Const OS_WINDOWS_SERVER_2008 As String = "WINDOWS SERVER 2008"      'CSCLP-768
'Private Const OS_WINDOWS_UNKNOWN As String = "WINDOWS PLATFORM UNKNOWN"

Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128 '// Service Pack
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128 '// Service Pack
    wServicePackMajor           As Integer
    wServicePackMinor           As Integer
    wSuiteMask                  As Integer
    wProductType                As Byte
    wReserved                   As Byte
End Type

Private Type SYSTEM_INFO
    '//dwOemID                  As Long
        wProcessorArchitecture  As Integer
        wReserved               As Integer
    dwPageSize                  As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask       As Long
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    wProcessorLevel             As Integer
    wProcessorRevision          As Integer
End Type

'**********
'
Public Enum WindowsVersion
       WIN_CE = -1
       UnKnown = 0
       WIN_31 = 1
       WIN_95 = 2
       WIN_98 = 3
       WIN_ME = 4
       WIN_NT_3x = 5
       WIN_NT_4x = 6
       WIN_2000 = 7
       WIN_2000Serv = 8
       WIN_XP = 9
       WIN_XP_HOME = 10
       WIN_2003 = 11
       WIN_VISTA = 12
       WIN_2008Serv = 13
       WIN_Win7 = 14
       WIN_2008ServR2 = 15
End Enum
'
Public Type WINVER
   OSVersion   As String
   OSVersShort As WindowsVersion
   BuildNr     As String
   ServicePack As String
   FullInfo    As String
End Type

Public OSInfos As String

'**********
'// API: min. Windows 95
Private Declare Function GetVersionEx1 Lib "kernel32.dll" Alias _
    "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO _
    ) As Long

'// API: min. Windows 2000
Private Declare Function GetVersionEx2 Lib "kernel32.dll" Alias _
    "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX _
    ) As Long

'// API: min. Windows XP
Private Declare Function IsWow64Process Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByRef Wow64Process As Long _
    ) As Long

'// API: min Windows 95
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'// API: min. Windows 2000
Private Declare Sub GetSystemInfo Lib "kernel32.dll" ( _
    ByRef lpSystemInfo As SYSTEM_INFO _
    )

'// API: min. Windows 2003 (WOW64)
Private Declare Sub GetNativeSystemInfo Lib "kernel32.dll" ( _
    ByRef lpSystemInfo As SYSTEM_INFO _
    )

'// API: min. Windows 2000
Private Declare Function GetSystemMetrics Lib "user32.dll" ( _
    ByVal nIndex As Long _
    ) As Long

'// API: only Windows Vista or Windows Server 2008
Private Declare Function GetProductInfo Lib "kernel32.dll" ( _
    ByVal dwOSMajorVersion As Long, _
    ByVal dwOSMinorVersion As Long, _
    ByVal dwSpMajorVersion As Long, _
    ByVal dwSpMinorVersion As Long, _
    ByRef pdwReturnedProductType As Long _
    ) As Boolean
    
'// Const: GetSystemMetrics
Private Const SM_SERVERR2                               As Long = 89&       '// Windows Server 2003 R2
Private Const SM_MEDIACENTER                            As Long = 87&       '// Windows XP Media Center Edition
Private Const SM_STARTER                                As Long = 88&       '// Windows XP Starter Edition
Private Const SM_TABLETPC                               As Long = 86&       '// Windows XP Tablet PC Edition

'// Const: GetVersionEx.wProcessorArchitecture
Private Const PROCESSOR_ARCHITECTURE_AMD64              As Long = &H9&      '// x64 (AMD Or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA64               As Long = &H6&      '// Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_INTEL              As Long = &H0&      '// x86
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN            As Long = &HFFFF&   '// Unknown architecture.

'// Const: GetVersionEx.wProductType
Private Const VER_NT_DOMAIN_CONTROLLER                  As Long = &H2&      '// The system is a domain controller and the operating system is Windows Server 2008, Windows Server 2003, or Windows 2000 Server.
Private Const VER_NT_SERVER                             As Long = &H3&      '// The operating system is Windows Server 2008, Windows Server 2003, or Windows 2000 Server. Note that a server that is also a domain controller is reported as VER_NT_DOMAIN_CONTROLLER, not VER_NT_SERVER.
Private Const VER_NT_WORKSTATION                        As Long = &H1&      '// The operating system is Windows Vista, Windows XP Professional, Windows XP Home Edition, or Windows 2000 Professional.

'// Const: GetVersionEx.wSuiteMask
Private Const VER_SUITE_BACKOFFICE                      As Long = &H4&      '// Microsoft BackOffice components are installed.
Private Const VER_SUITE_BLADE                           As Long = &H400&    '// Windows Server 2003, Web Edition is installed.
Private Const VER_SUITE_COMPUTE_SERVER                  As Long = &H4000&   '// Windows Server 2003, Compute Cluster Edition is installed.
Private Const VER_SUITE_DATACENTER                      As Long = &H80&     '// Windows Server 2008 Datacenter, Windows Server 2003, Datacenter Edition, or Windows 2000 Datacenter Server is installed.
Private Const VER_SUITE_ENTERPRISE                      As Long = &H2&      '// Windows Server 2008 Enterprise, Windows Server 2003, Enterprise Edition, or Windows 2000 Advanced Server is installed. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_EMBEDDEDNT                      As Long = &H40&     '// Windows XP Embedded is installed.
Private Const VER_SUITE_PERSONAL                        As Long = &H200&    '// Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition is installed.
Private Const VER_SUITE_SINGLEUSERTS                    As Long = &H100&    '// Remote Desktop is supported, but only one interactive session is supported. This value is set unless the system is running in application server mode.
Private Const VER_SUITE_SMALLBUSINESS                   As Long = &H1&      '// Microsoft Small Business Server was once installed on the system, but may have been upgraded to another version of Windows. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED        As Long = &H20&     '// Microsoft Small Business Server is installed with the restrictive client license in force. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_STORAGE_SERVER                  As Long = &H2000&   '// Windows Storage Server 2003 R2 or Windows Storage Server 2003is installed.
Private Const VER_SUITE_TERMINAL                        As Long = &H10&     '// Terminal Services is installed. This value is always set. If VER_SUITE_TERMINAL is set but VER_SUITE_SINGLEUSERTS is not set, the system is running in application server mode.
Private Const VER_SUITE_WH_SERVER                       As Long = &H8000&   '// Windows Home Server is installed.

'// Const: GetVersionEx.dwPlatformId
Private Const VER_PLATFORM_WIN32s                       As Long = &H0&      '// Specifies the Windows 3.1 OS.
Private Const VER_PLATFORM_WIN32_WINDOWS                As Long = &H1&      '// Specifies the Windows 95 or Windows 98 OS.
Private Const VER_PLATFORM_WIN32_NT                     As Long = &H2&      '// Specifies the Windows NT OS.
Private Const VER_PLATFORM_WIN32_CE                     As Long = &H3&      '// Specifies the Windows CE OS.

'// Const: GetProductInfo(pdwReturnedProductType)
Private Const PRODUCT_BUSINESS                          As Long = &H6&      '// Business Edition
Private Const PRODUCT_BUSINESS_N                        As Long = &H10&     '// Business Edition
Private Const PRODUCT_CLUSTER_SERVER                    As Long = &H12&     '// Cluster Server Edition
Private Const PRODUCT_DATACENTER_SERVER                 As Long = &H8&      '// Server Datacenter Edition (full installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE            As Long = &HC&      '// Server Datacenter Edition (core installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE_V          As Long = &H27&     '// Server Datacenter Edition without Hyper-V (core installation)
Private Const PRODUCT_DATACENTER_SERVER_V               As Long = &H25&     '// Server Datacenter Edition without Hyper-V (full installation)
Private Const PRODUCT_ENTERPRISE                        As Long = &H4&      '// Enterprise Edition
Private Const PRODUCT_ENTERPRISE_N                      As Long = &H1B&     '// Enterprise Edition
Private Const PRODUCT_ENTERPRISE_SERVER                 As Long = &HA&      '// Server Enterprise Edition (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE            As Long = &HE&      '// Server Enterprise Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V          As Long = &H29&     '// Server Enterprise Edition without Hyper-V (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_IA64            As Long = &HF&      '// Server Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_ENTERPRISE_SERVER_V               As Long = &H26&     '// Server Enterprise Edition without Hyper-V (full installation)
Private Const PRODUCT_HOME_BASIC                        As Long = &H2&      '// Home Basic Edition
Private Const PRODUCT_HOME_BASIC_N                      As Long = &H5&      '// Home Basic Edition
Private Const PRODUCT_HOME_PREMIUM                      As Long = &H3&      '// Home Premium Edition
Private Const PRODUCT_HOME_PREMIUM_N                    As Long = &H1A&     '// Home Premium Edition
Private Const PRODUCT_HOME_SERVER                       As Long = &H13&     '// Home Server Edition
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT  As Long = &H1E&     '// Windows Essential Business Server Management Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING   As Long = &H20&     '// Windows Essential Business Server Messaging Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY    As Long = &H1F&     '// Windows Essential Business Server Security Server
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS          As Long = &H18&     '// Server for Small Business Edition
Private Const PRODUCT_SMALLBUSINESS_SERVER              As Long = &H9&      '// Small Business Server
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM      As Long = &H19&     '// Small Business Server Premium Edition
Private Const PRODUCT_STANDARD_SERVER                   As Long = &H7&      '// Server Standard Edition (full installation)
Private Const PRODUCT_STANDARD_SERVER_CORE              As Long = &HD&      '// Server Standard Edition (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE_V            As Long = &H28&     '// Server Standard Edition without Hyper-V (core installation)
Private Const PRODUCT_STANDARD_SERVER_V                 As Long = &H24&     '// Server Standard Edition without Hyper-V (full installation)
Private Const PRODUCT_STARTER                           As Long = &HB&      '// Starter Edition
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER         As Long = &H17&     '// Storage Server Enterprise Edition
Private Const PRODUCT_STORAGE_EXPRESS_SERVER            As Long = &H14&     '// Storage Server Express Edition
Private Const PRODUCT_STORAGE_STANDARD_SERVER           As Long = &H15&     '// Storage Server Standard Edition
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER          As Long = &H16&     '// Storage Server Workgroup Edition
Private Const PRODUCT_UNDEFINED                         As Long = &H0&      '// An unknown product
Private Const PRODUCT_ULTIMATE                          As Long = &H1&      '// Ultimate Edition
Private Const PRODUCT_ULTIMATE_N                        As Long = &H1C&     '// Ultimate Edition
Private Const PRODUCT_WEB_SERVER                        As Long = &H11&     '// Web Server Edition (full installation)
Private Const PRODUCT_WEB_SERVER_CORE                   As Long = &H1D&     '// Web Server Edition (core installation)


Public Function GetOSVersion() As String
    
    GetOSVersion = GetOSVersionEx.OSVersion
    
End Function

Private Function GetOSVersionEx() As WINVER 'String
'
Dim OsVersInfoEx    As OSVERSIONINFOEX
Dim OsVersInfo      As OSVERSIONINFO
Dim OsSystemInfo    As SYSTEM_INFO
Dim OSVersion       As String
Dim Ret             As Long
Dim WinE            As WindowsVersion

    OsVersInfo.dwOSVersionInfoSize = Len(OsVersInfo)
    If GetVersionEx1(OsVersInfo) = 0 Then Err.Raise -1

    Select Case OsVersInfo.dwPlatformId
    Case VER_PLATFORM_WIN32s        '//Specifies the Windows 3.1 OS.
        OSVersion = "Windows 3.1"
        WinE = WIN_31
        
    Case VER_PLATFORM_WIN32_WINDOWS '// Specifies the Windows 95 or Windows 98 or Windows ME OS.
        Select Case OsVersInfo.dwMinorVersion
        Case 0:
            If OsVersInfo.dwBuildNumber = 950 Then
                OSVersion = "Windows 95"
            ElseIf (OsVersInfo.dwBuildNumber > 950) And (OsVersInfo.dwBuildNumber <= 1080) Then
                OSVersion = "Windows 95 SP1"
            ElseIf OsVersInfo.dwBuildNumber > 1080 Then
                OSVersion = "Windows 95 OSR2"
            End If
            '
            WinE = WIN_95
            
        Case 10:
            If OsVersInfo.dwBuildNumber = 1998 Then
                OSVersion = "Windows 98"
            ElseIf (OsVersInfo.dwBuildNumber > 1998) And (OsVersInfo.dwBuildNumber < 2183) Then
                OSVersion = "Windows 98 SP1"
            ElseIf OsVersInfo.dwBuildNumber >= 2183 Then
                OSVersion = "Windows 98 SE"
            End If
            '
            WinE = WIN_98
            
        Case 90:
            OSVersion = "Windows ME"
            WinE = WIN_ME
            
        End Select
        
    Case VER_PLATFORM_WIN32_NT      '// Specifies the Windows NT OS.
        Select Case OsVersInfo.dwMajorVersion
        Case 3:     OSVersion = "Windows NT 3.x"
                    WinE = WIN_NT_3x
        Case 4:     OSVersion = "Windows NT 4.x"
                    WinE = WIN_NT_4x
        Case 5:                     '// The operating system is Windows Server 2003 R2, _
                                        Windows Server 2003, Windows XP, or Windows 2000.
            OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
            If GetVersionEx2(OsVersInfoEx) = 0 Then Err.Raise -1
                
            Select Case OsVersInfoEx.dwMinorVersion
            Case 0:                 '// The operating system is Windows 2000
                If OsVersInfoEx.wProductType = VER_NT_WORKSTATION Then
                    OSVersion = "Windows 2000 Professional"
                    WinE = WIN_2000
                Else
                    If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                        OSVersion = "Windows 2000 Datacenter Server"
                    ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                        OSVersion = "Windows 2000 Advanced Server"
                    Else
                        OSVersion = "Windows 2000 Server"
                    End If
                    '
                    WinE = WIN_2000Serv
                End If
                '
            Case 1:                 '// The operating system is Windows XP.
                If GetSystemMetrics(SM_MEDIACENTER) Then
                    OSVersion = "Windows XP Media Center Edition"
                    WinE = WIN_XP
                ElseIf GetSystemMetrics(SM_STARTER) Then
                    OSVersion = "Windows XP Starter Edition"
                    WinE = WIN_XP_HOME
                ElseIf GetSystemMetrics(SM_TABLETPC) Then
                    OSVersion = "Windows XP Tablet PC Edition"
                    WinE = WIN_XP
                ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_PERSONAL Then
                    OSVersion = "Windows XP Home Edition"
                    WinE = WIN_XP_HOME
                Else
                    OSVersion = "Windows XP Professional"
                    WinE = WIN_XP
                End If
                        
            Case 2:                 '// The operating system is Windows Server 2003 R2, _
                                        Windows Server 2003, or Windows XP Professional x64 Edition.
                If IsWow64Process(GetCurrentProcess(), Ret) = 0 Then Err.Raise -1
                If Ret <> 0 Then
                    Call GetNativeSystemInfo(OsSystemInfo)
                Else
                    Call GetSystemInfo(OsSystemInfo)
                End If
    
                If GetSystemMetrics(SM_SERVERR2) Then
                    OSVersion = "Windows Server 2003 R2, "
                    WinE = WIN_2003
                ElseIf OsVersInfoEx.wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                    OSVersion = "Windows Storage Server 2003"
                    WinE = WIN_2003
                ElseIf (OsVersInfoEx.wProductType = VER_NT_WORKSTATION) And (OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                    OSVersion = "Windows XP Professional x64 Edition"
                    WinE = WIN_XP
                Else
                    OSVersion = "Windows Server 2003, "
                    WinE = WIN_2003
                End If
                
                '// Test for the server type.
                If OsVersInfoEx.wProductType <> VER_NT_WORKSTATION Then
                    If OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64 Then
                        If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                           OSVersion = OSVersion & "Datacenter Edition for Itanium-based Systems"
                        ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                           OSVersion = OSVersion & "Enterprise Edition for Itanium-based Systems"
                        End If
                    End If

                ElseIf OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                    If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                        OSVersion = OSVersion & "Datacenter x64 Edition"
                    ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                        OSVersion = OSVersion & "Enterprise x64 Edition"
                    Else
                        OSVersion = OSVersion & "Standard x64 Edition"
                    End If
                
                Else
                    If OsVersInfoEx.wSuiteMask And VER_SUITE_COMPUTE_SERVER Then
                        OSVersion = OSVersion & "Compute Cluster Edition"
                    ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                        OSVersion = OSVersion & "Datacenter Edition"
                    ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                        OSVersion = OSVersion & "Enterprise Edition"
                    ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_BLADE Then
                        OSVersion = OSVersion & "Web Edition"
                    Else
                        OSVersion = OSVersion & "Standard Edition"
                    End If
                End If
                
            End Select
            
        Case 6:                     '// The operating system is Windows Vista or _
                                        Windows Server 2008.
            OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
            If GetVersionEx2(OsVersInfoEx) = 0 Then Err.Raise -1
            
            If OsVersInfoEx.wProductType = VER_NT_WORKSTATION Then
               Select Case OsVersInfoEx.dwMinorVersion
               Case 0:
                  OSVersion = "Windows Vista "
                  WinE = WIN_VISTA
               Case 1:
                  OSVersion = "Windows 7 "
                  WinE = WIN_Win7
               End Select
            Else
               Select Case OsVersInfoEx.dwMinorVersion
               Case 0:
                  OSVersion = "Windows Server 2008 "
                  WinE = WIN_2008Serv
               Case 1:
                  OSVersion = "Windows Server 2008 R2 "
                  WinE = WIN_2008ServR2
               End Select
            End If
            '
            '
            Dim dwType As Long
            Call GetProductInfo(6, 0, 0, 0, dwType)

            Select Case dwType
            Case PRODUCT_ULTIMATE:                      OSVersion = OSVersion & "Ultimate Edition"
            Case PRODUCT_HOME_PREMIUM:                  OSVersion = OSVersion & "Home Premium Edition"
            Case PRODUCT_HOME_BASIC:                    OSVersion = OSVersion & "Home Basic Edition"
            Case PRODUCT_ENTERPRISE:                    OSVersion = OSVersion & "Enterprise Edition"
            Case PRODUCT_BUSINESS:                      OSVersion = OSVersion & "Business Edition"
            Case PRODUCT_STARTER:                       OSVersion = OSVersion & "Starter Edition"
            Case PRODUCT_CLUSTER_SERVER:                OSVersion = OSVersion & "Cluster Server Edition"
            Case PRODUCT_DATACENTER_SERVER:             OSVersion = OSVersion & "Datacenter Edition"
            Case PRODUCT_DATACENTER_SERVER_CORE:        OSVersion = OSVersion & "Datacenter Edition (core installation)"
            Case PRODUCT_ENTERPRISE_SERVER:             OSVersion = OSVersion & "Enterprise Edition"
            Case PRODUCT_ENTERPRISE_SERVER_CORE:        OSVersion = OSVersion & "Enterprise Edition (core installation)"
            Case PRODUCT_ENTERPRISE_SERVER_IA64:        OSVersion = OSVersion & "Enterprise Edition for Itanium-based Systems"
            Case PRODUCT_SMALLBUSINESS_SERVER:          OSVersion = OSVersion & "Small Business Server"
            Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM:  OSVersion = OSVersion & "Small Business Server Premium Edition"
            Case PRODUCT_STANDARD_SERVER:               OSVersion = OSVersion & "Standard Edition"
            Case PRODUCT_STANDARD_SERVER_CORE:          OSVersion = OSVersion & "Standard Edition (core installation)"
            Case PRODUCT_WEB_SERVER:                    OSVersion = OSVersion & "Web Server Edition"
            End Select

            If IsWow64Process(GetCurrentProcess(), Ret) = 0 Then Err.Raise -1
            If Ret <> 0 Then
                Call GetNativeSystemInfo(OsSystemInfo)
            Else
                Call GetSystemInfo(OsSystemInfo)
            End If
            
            If OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                OSVersion = OSVersion & ", 64-bit"
            ElseIf OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL Then
                OSVersion = OSVersion & ", 32-bit"
            End If
            
        End Select
        
    Case VER_PLATFORM_WIN32_CE:     '// Specifies the Windows CE OS.
        OSVersion = "Windows CE"
        WinE = WIN_CE
    End Select
    
    '// Include service pack.
    Dim strSP   As String
    Dim strFI(1)   As String
    Dim lSp     As Long
    '
    strSP = OsVersInfo.szCSDVersion
    lSp = InStr(1, strSP, Chr$(0))
    '
    If lSp Then strSP = Mid$(strSP, 1, lSp - 1)
    'If Len(strSp) Then OSVersion = OSVersion & " (" & strSp & ")"
        
    With GetOSVersionEx
      .OSVersion = OSVersion
      .ServicePack = strSP
      .OSVersShort = WinE
      .BuildNr = OsVersInfo.dwBuildNumber
      .FullInfo = OSVersion
         If Len(CStr(.BuildNr)) Then strFI(0) = "Build " & CStr(.BuildNr)
         If Len(strSP) Then strFI(1) = strSP
         strFI(0) = Join(strFI(), ", ")
         If Len(strFI(0)) Then .FullInfo = .FullInfo & " (" & strFI(0) & ")"
    End With
    '
End Function

'Output determines whether the user wants to continue or not
Public Function IsOpLockEnabled(ByVal ApplicationName As String, _
                                ByVal MMain As Object) As Boolean
    
    Dim clsRegistry As New CRegistry
    
    Dim arrSubkeys As Variant
    Dim arrValues As Variant
    
    Dim lngSubkeyCtr As Long
    Dim lngValuesCtr As Long
    
    Dim blnOpsLockKeyFound As Boolean
    Dim blnOpslockIsEnabled As Boolean
    Dim blnOSIsSMB2 As Boolean
    Dim blnSMB2IsEnabled As Boolean
    Dim blnSMB2KeyFound As Boolean
    Dim blnNetworkPath As Boolean
    
    Dim strOSVersion As String
    Dim strMessage As String
    Dim strOplocksPathFound As String
    Dim strMDBPath As String
    
    Dim objVersion As WINVER
    
    On Error GoTo ErrorHandler
    
    blnOpsLockKeyFound = False
    blnOpslockIsEnabled = True
    blnOSIsSMB2 = False
    blnSMB2IsEnabled = True
    blnSMB2KeyFound = False
    
    MMain.OpLockCheckAddToTrace "Start of Opportunistic Lock Checking..."
    MMain.OpLockCheckAddToTrace "Start of OS Version Checking..."
    
    strMDBPath = GetSetting("ClearingPoint", "Settings", "MDBPath")
    blnNetworkPath = IsMDBOnNetworkPath(strMDBPath)
    
    MMain.OpLockCheckAddToTrace "MDB Path = " & strMDBPath
    
    objVersion = GetOSVersionEx
    strOSVersion = objVersion.FullInfo
    
    '*********************************************************************************************************************************************
    'IF Operating System is Win95/Win98/Win9x/WinME Bypass Opslock checking.
    'IF Operating System is Unknown Bypass Opslock Chaecking.
    '*********************************************************************************************************************************************
    Select Case objVersion.OSVersShort
        Case 0, -1
            MMain.OpLockCheckAddToTrace "Platform: " & strOSVersion
            GoTo EndChecking3
            
        Case 2, 3, 4
            MMain.OpLockCheckAddToTrace "Platform: " & strOSVersion
            GoTo EndChecking4
        
        Case 5, 6, 7, 8, 9, 10, 11
            MMain.OpLockCheckAddToTrace "Platform: " & strOSVersion
            
        Case 12, 13, 14, 15
            MMain.OpLockCheckAddToTrace "Platform: " & strOSVersion
            blnOSIsSMB2 = True
    End Select
    '*********************************************************************************************************************************************
    
    '*********************************************************************************************************************************************
    'For Windows Vista and later OS version checking for SMB2 is necessary
    '*********************************************************************************************************************************************
    If blnOSIsSMB2 = True Then
        MMain.OpLockCheckAddToTrace "Check for SMB2 for Windows Vista and later OS versions."
        
        ' Search for SMB2 Key - if set to 0 Oplock is Disabled ( Server )
        If clsRegistry.GetRegistrySettings(cpiLocalMachine, "", REG_ENABLEOPLOCKS_ROOTPATH_SERVER, cpiUserDefined, True) = True Then
            
            MMain.OpLockCheckAddToTrace "Registry Path: " & REG_ENABLEOPLOCKS_ROOTPATH_SERVER
            
            If Trim(clsRegistry.RegistryValue) <> "" Then
                arrSubkeys = Split(clsRegistry.RegistryValue, "|||||")
                For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                    arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                    For lngValuesCtr = 0 To UBound(arrValues)
                        
                        MMain.OpLockCheckAddToTrace "Registry Key: " & CStr(arrValues(lngValuesCtr)) & " = " & CStr(arrValues(lngValuesCtr + 1))
                        Debug.Print "Registry Key: " & CStr(arrValues(lngValuesCtr)) & " = " & CStr(arrValues(lngValuesCtr + 1))
                        
                        If UCase(CStr(arrValues(lngValuesCtr))) = UCase(REG_SMB2_REGKEY) Then
                            If CLng(arrValues(lngValuesCtr + 1)) = 0 Then
                                blnSMB2IsEnabled = False
                            Else
                                GoTo EndChecking2
                            End If
                            
                            blnSMB2KeyFound = True
                        End If
                        
                        Exit For
                    Next
                Next
                
                If blnSMB2KeyFound = False Then GoTo EndChecking2
            End If
        End If
    End If
    '*********************************************************************************************************************************************
    
    '*********************************************************************************************************************************************
    'Search for EnableOplocks Key - if set to 1 Oplock is Enabled ( Server )
    '*********************************************************************************************************************************************
    If (clsRegistry.GetRegistrySettings(cpiLocalMachine, "", REG_ENABLEOPLOCKS_ROOTPATH_SERVER, cpiUserDefined, True) = True) And Not blnNetworkPath Then
        
        MMain.OpLockCheckAddToTrace "Registry Path: " & REG_ENABLEOPLOCKS_ROOTPATH_SERVER
        
        If Trim(clsRegistry.RegistryValue) <> "" Then
            arrSubkeys = Split(clsRegistry.RegistryValue, "|||||")
            For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                For lngValuesCtr = 0 To UBound(arrValues)
                    
                    MMain.OpLockCheckAddToTrace "Registry Key: " & CStr(arrValues(lngValuesCtr)) & " = " & CStr(arrValues(lngValuesCtr + 1))
                    
                    If UCase(CStr(arrValues(lngValuesCtr))) = UCase(REG_ENABLEOPLOCKS_REGKEY) Then
                        If CLng(arrValues(lngValuesCtr + 1)) = 0 Then
                            blnOpslockIsEnabled = False
                        End If
                        
                        blnOpsLockKeyFound = True
                        
                        strOplocksPathFound = REG_ENABLEOPLOCKS_ROOTPATH_SERVER & "\" & CStr(arrValues(lngValuesCtr))
                        
                        GoTo EndChecking
                    End If
                    
                    Exit For
                Next
            Next
        End If
    End If
    '*********************************************************************************************************************************************
    
    '*********************************************************************************************************************************************
    'Search for EnableOplocks Key - if set to 1 Oplock is Enabled ( Workstation )
    '*********************************************************************************************************************************************
    If (clsRegistry.GetRegistrySettings(cpiLocalMachine, "", REG_ENABLEOPLOCKS_ROOTPATH_WORKSTATION, cpiUserDefined, True) = True) And blnNetworkPath Then
        
        MMain.OpLockCheckAddToTrace "Registry Path: " & REG_ENABLEOPLOCKS_ROOTPATH_WORKSTATION
        
        If Trim(clsRegistry.RegistryValue) <> "" Then
            arrSubkeys = Split(clsRegistry.RegistryValue, "|||||")
            For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                For lngValuesCtr = 0 To UBound(arrValues)
                    
                    MMain.OpLockCheckAddToTrace "Registry Key: " & CStr(arrValues(lngValuesCtr)) & " = " & CStr(arrValues(lngValuesCtr + 1))
                    
                    If UCase(CStr(arrValues(lngValuesCtr))) = UCase(REG_ENABLEOPLOCKS_REGKEY) Then
                        If CLng(arrValues(lngValuesCtr + 1)) = 0 Then
                            blnOpslockIsEnabled = False
                        End If
                        
                        blnOpsLockKeyFound = True
                        
                        strOplocksPathFound = REG_ENABLEOPLOCKS_ROOTPATH_WORKSTATION & "\" & CStr(arrValues(lngValuesCtr))
                        
                        GoTo EndChecking
                    End If
                    
                    Exit For
                Next
            Next
        End If
    End If
    '*********************************************************************************************************************************************
    
    '*********************************************************************************************************************************************
    'Search for OplocksDisabled Key - if set to 0 Oplock is Enabled
    '*********************************************************************************************************************************************
    If clsRegistry.GetRegistrySettings(cpiLocalMachine, "", REG_OPLOCKSDISABLED_ROOTPATH, cpiUserDefined, True) = True Then
        
        MMain.OpLockCheckAddToTrace "Registry Path: " & REG_OPLOCKSDISABLED_ROOTPATH
        
        If Trim(clsRegistry.RegistryValue) <> "" Then
            arrSubkeys = Split(clsRegistry.RegistryValue, "|||||")
            For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                For lngValuesCtr = 0 To UBound(arrValues)
                    
                    MMain.OpLockCheckAddToTrace "Registry Key: " & CStr(arrValues(lngValuesCtr)) & " = " & CStr(arrValues(lngValuesCtr + 1))
                    
                    If UCase(CStr(arrValues(lngValuesCtr))) = UCase(REG_OPLOCKSDISABLED_REGKEY) Then
                        If CLng(arrValues(lngValuesCtr + 1)) = 1 Then
                            blnOpslockIsEnabled = False
                        End If
                        
                        blnOpsLockKeyFound = True
                        
                        strOplocksPathFound = REG_OPLOCKSDISABLED_ROOTPATH & "\" & CStr(arrValues(lngValuesCtr))
                        
                        GoTo EndChecking
                    End If
                    
                    Exit For
                Next
            Next
        End If
    End If
    '*********************************************************************************************************************************************
    
EndChecking:
    
    '*********************************************************************************************************************************************
    'START OF <OS Supports SMB2 which by default has opslock enabled ( need to disable SMB2 before disabling opslock ).>
    '*********************************************************************************************************************************************
    If blnOSIsSMB2 = True Then
        'OS uses OpsLock and Registry Key for opslock was not found ( According to MSDN Opslock is enabled by default ) and supports SMB2
        If blnOpsLockKeyFound = False Then
            strMessage = vbNullString
            strMessage = strMessage & "Platform: " & strOSVersion
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "SMB2 is disabled on this machine but opslock key is not found ( According to MSDN Opslock is enabled by default )."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "The risk of corrupting the database is high."
            
            MMain.OpLockCheckAddToTrace strMessage
            IsOpLockEnabled = False
            
        'Registry Key for opslock was found and Opslock is enabled based on RegKey
        ElseIf blnOpsLockKeyFound = True And blnOpslockIsEnabled = True Then
            strMessage = vbNullString
            strMessage = strMessage & "Platform: " & strOSVersion
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "SMB2 is disabled on this machine but opslock key is enabled."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "The risk of corrupting the database is high."
            
            MMain.OpLockCheckAddToTrace strMessage
            IsOpLockEnabled = False
        
        'Registry Key for opslock was found and Opslock is disabled based on RegKey but SMB2 is enabled, SMB2 overrides opslock setting
        ElseIf blnOpsLockKeyFound = True And blnOpslockIsEnabled = False Then
            MMain.OpLockCheckAddToTrace "Oplocks is disabled at " & strOplocksPathFound & "."
            IsOpLockEnabled = False
            
        End If
    '*********************************************************************************************************************************************
    'END OF <OS Supports SMB2 which by default has opslock enabled ( need to disable SMB2 before disabling opslock ).>
    '*********************************************************************************************************************************************
        
    '*********************************************************************************************************************************************
    'START OF <OS DOES NOT Supports SMB2 (OS is older than Windows VISTA).>
    '*********************************************************************************************************************************************
    Else
        'Registry Key for opslock was not found ( According to MSDN Opslock is enabled by default ).
        If blnOpsLockKeyFound = False Then
            strMessage = vbNullString
            strMessage = strMessage & "Platform: " & strOSVersion
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "Registry Key for opslock was not found ( According to MSDN Opslock is enabled by default )."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "The risk of corrupting the database is high."
              
            MMain.OpLockCheckAddToTrace strMessage
            IsOpLockEnabled = False
        
        'Registry Key for opslock was found and Opslock is enabled based on RegKey.
        ElseIf blnOpsLockKeyFound = True And blnOpslockIsEnabled = True Then
            strMessage = vbNullString
            strMessage = strMessage & "Platform: " & strOSVersion
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "Registry Key for opslock was found and Opslock is enabled based on RegKey."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "The risk of corrupting the database is high."
               
            MMain.OpLockCheckAddToTrace strMessage
            IsOpLockEnabled = False
        
        'Opslock is Disabled.
        ElseIf blnOpsLockKeyFound = True And blnOpslockIsEnabled = False Then
            MMain.OpLockCheckAddToTrace "Oplocks is disabled at " & strOplocksPathFound & "."
            IsOpLockEnabled = False
        
        End If
        '*********************************************************************************************************************************************
    End If
    '*********************************************************************************************************************************************
    'END OF <OS DOES NOT Supports SMB2 (OS is older than Windows VISTA).>
    '*********************************************************************************************************************************************
    
    Exit Function
    
EndChecking2:
    '*********************************************************************************************************************************************
    'When SMB2 Key is found and is enabled / SMB2 key is not found ( therefore via MSDN its by default enabled )
    '*********************************************************************************************************************************************
    strMessage = vbNullString
    strMessage = strMessage & "Platform: " & strOSVersion
    strMessage = strMessage & vbCrLf
    strMessage = strMessage & "SMB2 is enabled on this machine ( According to MSDN Opslock cannot be disabled when SMB2 is enabled )."
    strMessage = strMessage & vbCrLf
    strMessage = strMessage & "The risk of corrupting the database is high."
     
    MMain.OpLockCheckAddToTrace strMessage
    IsOpLockEnabled = False
    
    Exit Function
    '*********************************************************************************************************************************************
    
EndChecking3:
    '*********************************************************************************************************************************************
    'OS is unknown ( Ask user if he/she wants to continue running application ).
    '*********************************************************************************************************************************************
    strMessage = vbNullString
    strMessage = strMessage & "Platform: UNKNOWN"
    strMessage = strMessage & vbCrLf
    strMessage = strMessage & "Opportunistic locking cannot be determined."
    strMessage = strMessage & vbCrLf
    strMessage = strMessage & "The risk of corrupting the database is high."
    
    MMain.OpLockCheckAddToTrace strMessage
    IsOpLockEnabled = False
    
    Exit Function
    
    '*********************************************************************************************************************************************

EndChecking4:
    '*********************************************************************************************************************************************
    'OS does not support Opslock.
    '*********************************************************************************************************************************************
    MMain.OpLockCheckAddToTrace "OS does not support Oplocks."
    IsOpLockEnabled = False
    Exit Function
    '*********************************************************************************************************************************************
    
ErrorHandler:
    
    Select Case Err.Number
        Case 0
            'Do Nothing
            
        Case Else
            strMessage = vbNullString
            strMessage = strMessage & "Error occured while checking opportunistic lock setting."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "Error Message: " & Err.Description & "( " & Err.Number & " )."
            strMessage = strMessage & vbCrLf
            strMessage = strMessage & "The risk of corrupting the database is unknown. Do you want to continue running " & ApplicationName & "?"
            
            MMain.OpLockCheckAddToTrace "Error occured while checking opportunistic lock setting."
            MMain.OpLockCheckAddToTrace "Error Message: " & Err.Description & "( " & Err.Number & " )."
            
            MMain.OpLockCheckAddToTrace strMessage
            IsOpLockEnabled = False
            
    End Select
        
    MMain.OpLockCheckAddToTrace "End of Opportunistic Lock Checking..." & vbNewLine
        
    Set clsRegistry = Nothing
    Set arrSubkeys = Nothing
    Set arrValues = Nothing
    
End Function


Private Function IsMDBOnNetworkPath(ByVal DBPath As String) As Boolean
    
    Dim fsoPath As New FileSystemObject
    Dim folPath As Folder
    
    If Left(Trim$(DBPath), 2) = "\\" Then
        IsMDBOnNetworkPath = True
    Else
        If Len(DBPath) > 3 Then
            Set folPath = fsoPath.GetFolder(Left(Trim$(DBPath), 3))
            
            Select Case folPath.Type
                Case "Network Drive"
                    IsMDBOnNetworkPath = True
                Case Else
                    IsMDBOnNetworkPath = False
                
            End Select
        End If
    End If
    
End Function


