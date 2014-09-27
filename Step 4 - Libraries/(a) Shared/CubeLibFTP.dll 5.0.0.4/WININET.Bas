Attribute VB_Name = "modWinInet"
Option Explicit

Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000


Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long

' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const INTERNET_FLAG_PASSIVE = &H8000000

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFilename As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#
Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))

Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Public hOpen As Long



Public Sub GetFilesAndFolders(sFilter As String, ByRef oConnection As clsConnection, hConnection As Long)
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim dError As Long
    Dim ptr As Long
    Dim pData As WIN32_FIND_DATA
    Dim sFilename As String
        
    'Set mDirCol = Nothing
    DoEvents
    Set oConnection.Files = Nothing
    Set oConnection.Folders = Nothing
    
    
    Set oConnection.Files = New clsFiles
    Set oConnection.Folders = New clsFolders
    
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, sFilter, pData, 0, 0)
    DoEvents
    'nLastError = Err.LastDllError
    DoEvents
    If hFind = 0 Then
        'If (nLastError <> ERROR_NO_MORE_FILES) Then
        '    ErrorOut Err.LastDllError, "FtpFindFirstFile"
        'End If
        Exit Sub
    End If
    
    dError = NO_ERROR
    Dim bRet As Boolean

    sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    'mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
    
    If (pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY) Then
        oConnection.Folders.Add sFilename, sFilename
    Else
        oConnection.Files.Add sFilename, pData.nFileSizeLow, Win32ToVbTime(pData.ftCreationTime), sFilename
    End If
    
    Do
        DoEvents
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        DoEvents
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
                'ErrorOut Err.LastDllError, "InternetFindNextFile"
                InternetCloseHandle (hFind)
                'Exit Function
            End If
            DoEvents
        Else
            sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            'mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
            If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
                oConnection.Folders.Add sFilename, sFilename
            Else
                oConnection.Files.Add sFilename, pData.nFileSizeLow, Win32ToVbTime(pData.ftCreationTime), sFilename
            End If
            
        End If
        DoEvents
    Loop
       'DoEvents
    'Set GetDirectoryListing = mDirCol
    'DoEvents
    InternetCloseHandle (hFind)
End Sub

Function Win32ToVbTime(ft As Currency) As Date
    Dim ftl As Currency
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        MsgBox Err.LastDllError
    End If
End Function

