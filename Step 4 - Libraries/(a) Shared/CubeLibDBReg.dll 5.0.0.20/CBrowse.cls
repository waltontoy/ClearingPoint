VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Public Enum BrowseMode
        enuFolder = 1
        enuFile = 2
    End Enum
    
    Private Type BrowseInfo
        hwndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long
        lpszTitle As Long
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type

    Public Enum ImageType
        imgFile = 0
        imgIcon = 1
        imgBitmap = 2
        imgJPG = 3
        imgAll = 4
        imgUnknownImage = 5
    End Enum
    
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type


    Private mvarPath As String
    Private mvarFileName As String
    Private mvarImageTrueHeight As Double
    Private mvarImageTrueWidth As Double

    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
    Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
    Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const MAX_PATH = 260
        
    Private Const LPTR = (&H0 Or &H40)

Private Property Let ImageTrueWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImageTrueWidth = 5
    mvarImageTrueWidth = vData
End Property


Public Property Get ImageTrueWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImageTrueWidth
    ImageTrueWidth = mvarImageTrueWidth
End Property



Private Property Let ImageTrueHeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImageTrueHeight = 5
    mvarImageTrueHeight = vData
End Property


Public Property Get ImageTrueHeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImageTrueHeight
    ImageTrueHeight = mvarImageTrueHeight
End Property

Public Function Browse(ByVal OwnerForm As Object, ByVal SelectionTitle As String, ByVal DefaultPath As String, ByVal Filter As String, Optional ByVal FilterType As ImageType) As Boolean
    Dim arrResult

    mvarPath = DefaultPath
    mvarFileName = ""

    arrResult = Split(frmFilePath.SelectFile(OwnerForm, SelectionTitle, DefaultPath, Filter, FilterType), "*****")

    mvarPath = CStr(arrResult(1))
    mvarFileName = CStr(arrResult(2))
    mvarImageTrueHeight = CStr(arrResult(3))
    mvarImageTrueWidth = CStr(arrResult(4))

    Browse = IIf(CStr(arrResult(0)) = "True", True, False)
End Function

Private Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



Private Property Let Path(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Path = 5
    mvarPath = vData
End Property


Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property

Public Function BrowseFolder(ByRef OwnerForm As Object, ByVal DialogInstruction As String, ByVal DefaultFolder As String) As Boolean
    Dim iNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim lpPathPointer As Long
    
    Dim udtBI As BrowseInfo

    With udtBI
        ' Set the Owner Window
        .hwndOwner = OwnerForm.hWnd
        
        ' lstrcat Appends the two strings and returns the Memory Address
        .lpszTitle = lstrcat(DialogInstruction, "")
        
        ' Dialog CallBack Function that preselects the Folder Specified
        .lpfnCallback = FunctionPointer(AddressOf BrowseCallbackProcStr)
        
        ' Allocate a String and get the Pointer to the allocation
        lpPathPointer = LocalAlloc(LPTR, Len(DefaultFolder) + 1)
        
        ' Copy the Path to the Allocated String
        CopyMemory ByVal lpPathPointer, ByVal DefaultFolder, Len(DefaultFolder) + 1
        
        ' The pointer to the Folder Path to Pre-Select for the CallBack Function
        .lParam = lpPathPointer
    
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
        '.ulFlags = 2
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    LocalFree lpPathPointer
    
    mvarPath = sPath
    mvarFileName = ""
    mvarImageTrueHeight = 0
    mvarImageTrueWidth = 0
    
    BrowseFolder = IIf(Trim(sPath) = "", False, True)
End Function

Public Function BrowseFile(ByRef OwnerForm As Object, ByRef Application As Object, ByVal DefaultPath As String, ByVal Filter As String) As Boolean
    Dim OFName As OPENFILENAME
    Dim strOpenPath As String
    Dim strOpenFile As String
    Dim strOpenPathFile As String
    
    OFName.lStructSize = Len(OFName)
    
    'Set the parent window
    OFName.hwndOwner = OwnerForm.hWnd
    
    'Set the application's instance
    OFName.hInstance = Application.hInstance
    
    'Select a filter
    OFName.lpstrFilter = "Files (" & Filter & ")" + Chr$(0) + Filter + Chr$(0)
    
    'create a buffer for the file
    OFName.lpstrFile = Space$(500)
    
    'set the maximum length of a returned file
    OFName.nMaxFile = 501
    
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(500)
    
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 501
    
    'Set the initial directory
    OFName.lpstrInitialDir = DefaultPath
    
    'Set the title
    OFName.lpstrTitle = "Open File - " & Filter
    
    ' No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    BrowseFile = GetOpenFileName(OFName)
            
    If BrowseFile Then
        strOpenPathFile = StripNullTerminator(OFName.lpstrFile)
        
        strOpenPath = Left(strOpenPathFile, InStrRev(strOpenPathFile, "\") - 1)
        strOpenFile = Mid(strOpenPathFile, InStrRev(strOpenPathFile, "\") + 1)
        
        mvarPath = strOpenPath
        mvarFileName = strOpenFile
        mvarImageTrueHeight = 0
        mvarImageTrueWidth = 0
    Else
        mvarPath = ""
        mvarFileName = ""
        mvarImageTrueHeight = 0
        mvarImageTrueWidth = 0
    End If


End Function
