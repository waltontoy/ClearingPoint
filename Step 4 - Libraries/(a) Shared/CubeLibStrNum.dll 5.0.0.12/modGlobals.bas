Attribute VB_Name = "modGlobals"
Option Explicit
    
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Type NUMBERFMT
    NumDigits As Long                 '  number of decimal digits
    LeadingZero As Long             '  if leading zero in decimal fields
    Grouping As Long                   '  group size left of decimal
    lpDecimalSep As String         '  ptr to decimal separator string
    lpThousandSep As String      '  ptr to thousand separator string
    NegativeOrder As Long          '  negative number ordering
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long

Public Declare Function GetDateFormat Lib "kernel32" _
                                            Alias "GetDateFormatA" ( _
                                            ByVal Locale As Long, _
                                            ByVal dwFlags As Long, _
                                            lpDate As SYSTEMTIME, _
                                            ByVal lpFormat As String, _
                                            ByVal lpDateStr As String, _
                                            ByVal cchDate As Long _
                                            ) As Long
                                            
Public Declare Function GetNumberFormat Lib "kernel32" _
                                            Alias "GetNumberFormatA" ( _
                                            ByVal Locale As Long, _
                                            ByVal dwFlags As Long, _
                                            ByVal lpValue As String, _
                                            lpFormat As NUMBERFMT, _
                                            ByVal lpNumberStr As String, _
                                            ByVal cchNumber As Long _
                                            ) As Long
                                            
Public Declare Function GetCurrencyFormat Lib "kernel32" _
                                            Alias "GetCurrencyFormatA" ( _
                                            ByVal Locale As Long, _
                                            ByVal dwFlags As Long, _
                                            ByVal lpValue As String, _
                                            lpFormat As Any, _
                                            ByVal lpCurrencyStr As String, _
                                            ByVal cchCurrency As Long _
                                            ) As Long

Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" ( _
                                            ByVal lpPath As String) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" _
                                            Alias "GetSystemDirectoryA" ( _
                                            ByVal lpBuffer As String, _
                                            ByVal nSize As Long _
                                            ) As Long

Public Declare Function GetTempPath Lib "kernel32" _
                                            Alias "GetTempPathA" ( _
                                            ByVal nBufferLength As Long, _
                                            ByVal lpBuffer As String _
                                            ) As Long

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Declare Sub MoveMemory Lib "kernel32" _
                                        Alias "RtlMoveMemory" ( _
                                        dest As Any, _
                                        ByVal Source As Long, _
                                        ByVal Length As Long)

Public Declare Function VerQueryValue Lib "Version.dll" _
                                        Alias "VerQueryValueA" ( _
                                        pBlock As Any, _
                                        ByVal lpSubBlock As String, _
                                        lplpstrFileInfoString As Any, _
                                        puLen As Long _
                                        ) As Long

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" _
                                        Alias "GetFileVersionInfoSizeA" ( _
                                        ByVal lptstrFilename As String, _
                                        lpdwHandle As Long _
                                        ) As Long

Public Declare Function GetShortPathName Lib "kernel32" _
                                        Alias "GetShortPathNameA" ( _
                                        ByVal lpszLongPath As String, _
                                        ByVal lpszShortPath As String, _
                                        ByVal lBuffer As Long _
                                        ) As Long

Public Declare Function GetFileVersionInfo Lib "Version.dll" _
                                        Alias "GetFileVersionInfoA" ( _
                                        ByVal lptstrFilename As String, _
                                        ByVal dwhandle As Long, _
                                        ByVal dwlen As Long, _
                                        lpData As Any _
                                        ) As Long

Public Declare Function lstrcpy Lib "kernel32" _
                                        Alias "lstrcpyA" ( _
                                        ByVal lpString1 As String, _
                                        ByVal lpString2 As Long _
                                        ) As Long

Public Declare Function CreateProcessA Lib "kernel32" ( _
                                        ByVal lpApplicationName As Long, _
                                        ByVal lpCommandLine As String, _
                                        ByVal lpProcessAttributes As Long, _
                                        ByVal lpThreadAttributes As Long, _
                                        ByVal bInheritHandles As Long, _
                                        ByVal dwCreationFlags As Long, _
                                        ByVal lpEnvironment As Long, _
                                        ByVal lpCurrentDirectory As Long, _
                                        lpStartupInfo As STARTUPINFO, _
                                        lpProcessInformation As PROCESS_INFORMATION _
                                        ) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
                                        ByVal hProcess As Long, _
                                        lpExitCode As Long _
                                        ) As Long

Public Declare Function CloseHandle Lib "kernel32" ( _
                                        ByVal hObject As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" ( _
                                        ByVal hHandle As Long, _
                                        ByVal dwMilliseconds As Long _
                                        ) As Long



'>> To retrieve regional settings

Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Declare Function GetThreadLocale Lib "kernel32" () As Long
