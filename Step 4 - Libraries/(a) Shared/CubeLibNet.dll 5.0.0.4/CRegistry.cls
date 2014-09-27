VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Public Enum DriveType
    cpiCDROM = 5
    cpiREMOVABLEDISK = 2
    cpiFIXEDDISK = 3
End Enum

Private Const FILE_DSN_NAME = "DBCFS.dsn"
       
Private Type QueryStringResult
    StringResult As Boolean
    StringValue As String
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum RegistryPathType
    cpiStandard = 1
    cpiUserDefined = 2
End Enum

Public Enum RegistryOperation
    cpiVoid = 0
    cpiGetSetting = 1
    cpiSaveSetting = 2
    cpiDeleteSetting = 3
    cpiDeleteKey = 4
    cpiEnumSettings = 5
    cpiEnumSubkeys = 6
End Enum

Public Enum TargetRegistry
    cpiCurrentUser = 1
    cpiLocalMachine = 2
    cpiCurrentConfig = 3
    cpiClassesRoot = 4
End Enum

Public Enum cpiRegistryDataTypeConstants

   cpiStringReg = 1
   cpiBinaryReg = 3
   cpiDWordReg = 4

End Enum

Private strFileDSNPath As String
Private strConnectionString As String

Private Const ERROR_SUCCESS = 0
Private Const REG_DWORD_BIG_ENDIAN = 5

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_CONFIG = &H80000005

Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4

Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)


Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpData As String, ByVal cbData As Long) As Long

Private mvarRegistryValue As String
Private mvarRegistryOperation As RegistryOperation
Private mvarChangeOperation As Boolean
    
Public Property Get RegistryOperation() As RegistryOperation
    
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RegistryOperation
    RegistryOperation = mvarRegistryOperation

End Property

Public Property Get RegistryValue() As String
    
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RegistryValue
    RegistryValue = mvarRegistryValue

End Property


Function GetRegistry(ByVal Target As TargetRegistry, _
                    ByVal ApplicationName As String, _
                    ByVal RegistryKey As String, _
                    ByVal RegistrySetting As String, _
                    Optional PathType As RegistryPathType = cpiStandard, _
                    Optional ByVal DefaultValue As String = "", _
                    Optional AllowNoAdminRights As Boolean = True) _
                    As Boolean

    Dim QueryResult As QueryStringResult
    Dim lngReturnHandle As Long
    Dim strRegistryPath As String
    Dim hKey As Long
    Dim hAccess As Long
    
    GetRegistry = True
    mvarRegistryOperation = cpiGetSetting
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If
                
    hAccess = IIf(AllowNoAdminRights, KEY_READ, KEY_ALL_ACCESS)
    
    ' Open the Key
    RegOpenKeyEx hKey, strRegistryPath, 0, hAccess, lngReturnHandle
    
    ' Get the Key's Result
    QueryResult = RegQueryStringValue(lngReturnHandle, RegistrySetting)
    If Not QueryResult.StringResult Then
        GetRegistry = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = QueryResult.StringValue
        
        If mvarRegistryValue = "" And DefaultValue <> "" Then
            mvarChangeOperation = False
            If Not SaveRegistry(Target, ApplicationName, RegistryKey, RegistrySetting, DefaultValue) Then
                GetRegistry = False
                Exit Function
            End If
        End If
    End If
    
    ' Close the Key
    If RegCloseKey(lngReturnHandle) <> 0 And lngReturnHandle <> 0 Then
        GetRegistry = False
    End If

End Function

Public Function SaveRegistry(ByVal Target As TargetRegistry, _
                            ByVal ApplicationName As String, _
                            ByVal RegistryKey As String, _
                            ByVal RegistrySetting As String, _
                            ByVal NewValue As String, _
                            Optional PathType As RegistryPathType = cpiStandard) _
                            As Boolean

    Dim lngKeyHandle As Long
    Dim strRegistryPath As String
    Dim strRegistrySetting As String
    
    Dim hKey As Long
                
    
    SaveRegistry = True
    If mvarChangeOperation Then
        mvarRegistryValue = ""
        mvarRegistryOperation = cpiSaveSetting
    End If
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If
            
    If RegCreateKey(hKey, strRegistryPath, lngKeyHandle) <> 0 Then
        SaveRegistry = True
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    End If
    
    If RegSetValueEx(lngKeyHandle, RegistrySetting, 0, REG_SZ, ByVal NewValue, Len(NewValue)) <> 0 Then
        SaveRegistry = False
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    End If
    
    If RegCloseKey(lngKeyHandle) <> 0 Then
        SaveRegistry = False
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    Else
        mvarRegistryValue = NewValue
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
    End If
    
End Function

Private Function RegQueryStringValue(ByVal hKey As Long, _
                                    ByVal strValueName As String) As QueryStringResult

    Dim lngLengthMinusNUllTerminator As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    Dim lngResult As Long
    
    Dim intData As Integer
    
    Dim strBuffer As String
    
    
    RegQueryStringValue.StringResult = True
    RegQueryStringValue.StringValue = ""
    
    ' Retrieve information about the key
    lngResult = RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal 0, lngDataBufferSize)
        
    If lngValueType = REG_SZ Then
        ' Create a buffer
        strBuffer = String(lngDataBufferSize, Chr$(0))
        
        ' Retrieve the key's content
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal strBuffer, lngDataBufferSize) <> 0 Then
            RegQueryStringValue.StringResult = False
        Else
            lngLengthMinusNUllTerminator = InStr(1, strBuffer, Chr$(0)) - 1
            If lngLengthMinusNUllTerminator > 0 Then
                RegQueryStringValue.StringValue = Left$(strBuffer, lngLengthMinusNUllTerminator)
            Else
                RegQueryStringValue.StringValue = ""
            End If
            ' RegQueryStringValue.StringValue = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
        End If
    ElseIf lngValueType = REG_BINARY Or lngValueType = REG_DWORD Then
        
        ' Retrieve the key's value
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, intData, lngDataBufferSize) <> 0 Then
            RegQueryStringValue.StringResult = False
        Else
            RegQueryStringValue.StringValue = intData
        End If
    End If
    
End Function

Public Function DeleteRegistry(ByVal Target As TargetRegistry, _
                                ByVal ApplicationName As String, _
                                ByVal RegistryKey As String, _
                                ByVal RegistrySetting As String, _
                                Optional PathType As RegistryPathType = cpiStandard) _
                                As Boolean

    Dim lngReturnHandle As Long
    Dim strRegistryPath As String
    
    Dim hKey As Long
        
        
    mvarRegistryOperation = cpiDeleteSetting
    DeleteRegistry = True
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
            
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If

    ' Open the Key
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    ' Delete the key's value
    If RegDeleteValue(lngReturnHandle, RegistrySetting) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    'close the key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strRegistryPath & "|||||" & RegistrySetting
    End If
    
End Function

Public Function DeleteRegistryKey(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RegistryKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard) _
                                    As Boolean
    Dim lngReturnHandle As Long
    Dim strRegistryPath As String
    
    Dim hKey As Long
    
    
    mvarRegistryOperation = cpiDeleteKey
    DeleteRegistryKey = True
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName
    Else
        strRegistryPath = RegistryKey
    End If
            
    ' Open the Key
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    ' Delete the Key
    If RegDeleteKey(lngReturnHandle, RegistryKey) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    ' Close the Key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
    Else
        If PathType = cpiStandard Then
            mvarRegistryValue = strRegistryPath & "\" & RegistryKey
        Else
            mvarRegistryValue = RegistryKey
        End If
    End If
    
End Function

Public Function GetRegistrySubkeys(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RootKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard) _
                                    As Boolean

    Dim lngReturnHandle As Long
    
    Dim intSubkeyCtr As Integer
    Dim hKey As Long
    
    Dim strRegistryPath As String
    Dim strSubkeyStream As String
    Dim strBufferSave As String
    
        
    mvarRegistryOperation = cpiEnumSubkeys
    
    GetRegistrySubkeys = True
    strSubkeyStream = ""
    intSubkeyCtr = 0
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RootKey
    Else
        strRegistryPath = RootKey
    End If
        
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        GetRegistrySubkeys = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    Do While True
        ' Create a buffer
        strBufferSave = String(255, 0)
        
        ' Enumerate the keys
        If RegEnumKeyEx(lngReturnHandle, intSubkeyCtr, strBufferSave, 255, 0, vbNullString, ByVal 0&, ByVal 0&) <> 0 Then
            Exit Do
        End If
        
        ' Concatenate Results
        strSubkeyStream = strSubkeyStream & "|||||" & StripTerminator(strBufferSave)
        
        intSubkeyCtr = intSubkeyCtr + 1
    Loop
    
    strSubkeyStream = IIf(Trim(strSubkeyStream) <> "", Mid(strSubkeyStream, 6), strSubkeyStream)
    
    ' Close the registry key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        GetRegistrySubkeys = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strSubkeyStream
    End If
    
End Function
Public Function GetRegistrySettings(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RegistryKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard, _
                                    Optional AllowNoAdminRights As Boolean = True) As Boolean

    Dim QueryResult As QueryStringResult
    Dim lngReturnHandle As Long
    
    Dim intSubkeyCtr As Integer
    Dim hKey As Long
    Dim hAccess As Long
    
    Dim strBufferSaveValue As String
    Dim strRegistryPath As String
    Dim strSubkeyStream As String
    Dim strBufferSave As String
           
    mvarRegistryOperation = cpiEnumSettings
        
    GetRegistrySettings = True
    strSubkeyStream = ""
    intSubkeyCtr = 0
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If
    
    hAccess = IIf(AllowNoAdminRights, KEY_READ, KEY_ALL_ACCESS)
    
    If RegOpenKeyEx(hKey, strRegistryPath, 0, hAccess, lngReturnHandle) <> 0 Then
        GetRegistrySettings = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    Do While True
        ' Create a buffer
        strBufferSave = String(255, 0)
        
        ' enumerate the values
        If RegEnumValue(lngReturnHandle, intSubkeyCtr, strBufferSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then
            Exit Do
        End If
        
        QueryResult = RegQueryStringValue(lngReturnHandle, StripTerminator(strBufferSave))
        If Not QueryResult.StringResult Then
            GetRegistrySettings = False
            mvarRegistryValue = ""
            Exit Do
        Else
            strBufferSaveValue = QueryResult.StringValue
        
            ' Concatenate Results
            strSubkeyStream = strSubkeyStream & "|||||" & StripTerminator(strBufferSave) & "*****" & IIf(Trim(strBufferSaveValue) = "", "~~~~~", strBufferSaveValue)
            
            intSubkeyCtr = intSubkeyCtr + 1
        End If
    Loop
    
    strSubkeyStream = IIf(Trim(strSubkeyStream) <> "", Mid(strSubkeyStream, 6), strSubkeyStream)
    
    ' Close the Registry
    If RegCloseKey(lngReturnHandle) <> 0 Then
        GetRegistrySettings = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strSubkeyStream
    End If
    
End Function

Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    
    'Search the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
    
End Function

Private Sub Class_Initialize()

    mvarRegistryValue = ""
    mvarChangeOperation = True
    mvarRegistryOperation = cpiVoid
    
End Sub

Public Function RegistryKeyExists(ByVal Target As TargetRegistry, _
                                ByVal ApplicationName As String, _
                                ByVal RootKey As String, _
                                ByVal SearchKey As String, _
                                Optional PathType As RegistryPathType = cpiStandard, _
                                Optional GetSubKeys As Boolean = False) _
                                As Boolean
    
    Dim strSubkeyStream As String
    Dim strTempKey As String
    Dim strRootKey As String
    Dim strSubKey As String
    
    Dim arrSubkeys
    Dim arrValues
    
    Dim lngSubkeyCtr As Long
    Dim lngValuesCtr As Long
    Dim intCount As Long
    
    Dim blnParseSettings As Boolean
    
    RegistryKeyExists = False
    strTempKey = ""
                    
    ' Get SubKeys
    If GetSubKeys = True Then
        blnParseSettings = GetRegistrySubkeys(Target, ApplicationName, RootKey, PathType)
    Else
        blnParseSettings = GetRegistrySettings(Target, ApplicationName, RootKey, PathType)
    End If
    
    If blnParseSettings Then
        If Trim(mvarRegistryValue) <> "" Then
            arrSubkeys = Split(mvarRegistryValue, "|||||")
            For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                For lngValuesCtr = 0 To UBound(arrValues)
                    If UCase(CStr(arrValues(lngValuesCtr))) = UCase(SearchKey) Then
                        RegistryKeyExists = True
                        Exit For
                    End If
                Next
            Next
        Else
            If UCase(CStr(Trim(mvarRegistryValue))) = UCase(Trim(SearchKey)) Then
                RegistryKeyExists = True
            End If
        End If
    End If
    
    ' Make Sure it is Empty
    mvarRegistryValue = ""
    
End Function

Public Function IsAllowedAdminRights() As Boolean
    Dim QueryResult As QueryStringResult
    
    Dim lngReturnHandle As Long
    Dim hAccess As Long
    Dim hKey As Long
    
    Dim strRegistryPath As String
    
    
    IsAllowedAdminRights = True
        
    hKey = HKEY_LOCAL_MACHINE
    
    strRegistryPath = "Software\Microsoft\Windows\CurrentVersion"
    
    ' Open the Key
    RegOpenKeyEx hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle
    
    ' Get the Key's Result
    QueryResult = RegQueryStringValue(lngReturnHandle, "ProgramFilesDir")
    If Not QueryResult.StringResult Then
        IsAllowedAdminRights = False
    End If
    
    ' Close the Key
    If RegCloseKey(lngReturnHandle) <> 0 And lngReturnHandle <> 0 Then
        IsAllowedAdminRights = False
    End If
    
End Function

Public Function CreateRegistryKey(ByVal Target As TargetRegistry, _
                            ByVal RegistryKey As String) As Boolean

    Dim hKey As Long
    Dim lngKeyHandle As Long
    Dim strRegistryPath As String
    Dim blnFlag As Boolean
                
    CreateRegistryKey = False
    
    hKey = GetTrueHkey(Target)
    
    strRegistryPath = RegistryKey
    If (RegCreateKey(hKey, strRegistryPath, lngKeyHandle) = 0) Then
        
        CreateRegistryKey = True
        
    End If
    
    
End Function
Public Function CreateRegistryValue(ByVal Target As TargetRegistry _
                                                                              , ByVal RootKeyPath As String _
                                                                              , ByVal NewValueName As String _
                                                                              , ByVal DataType As cpiRegistryDataTypeConstants) As Boolean
                                                                              
    Dim lngHkey As Long
    
    lngHkey = GetTrueHkey(Target)
                                                                              
    If CreateValueName(lngHkey, RootKeyPath, NewValueName, DataType) Then
        CreateRegistryValue = True
        'SaveSetting
    Else
        CreateRegistryValue = False
        'MsgBox "Can't create valuename '" & NewValueName & "'", vbInformation
    End If
   
    'r.SetValue(eHKEY_LOCAL_MACHINE, subkey, "RegisteredOwner", n)
   
End Function


Private Function CreateValueName(PredefinedKey As Long, KeyName As String, ValueName As String, _
                                                                  DataType As cpiRegistryDataTypeConstants) As Boolean
    Dim rc As Long
    Dim hKey As Long
    Dim lpType As Long
    Dim lpcbData As Long
    Dim lpData As String
    Dim lngTemp As Long
    'Dim Value As String
    
    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    
    'Open the registry
    rc = RegOpenKeyEx(PredefinedKey, _
            KeyName, _
            0, _
            KEY_ALL_ACCESS, _
            hKey)

   'lngTemp = 0

    If rc = ERROR_SUCCESS Then
        'Create a buffer so we can retrieve the data type of
        'the key. We'll need this to determine which API
        'we should call.
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
         Select Case DataType
            Case cpiStringReg
                 'Use a string data type
                 rc = RegSetValueExString(hKey, _
                     ValueName, _
                     0, _
                     REG_SZ, _
                      "", _
                      1)
            
            Case cpiBinaryReg
               lpcbData = 0
               lngTemp = 0
                 rc = RegSetValueEx(hKey, _
                      ValueName, _
                     0, _
                     REG_BINARY, _
                       lngTemp, _
                     lpcbData)
               
            Case cpiDWordReg
                 'Use a DWORD data type
                 ' error here
               lpcbData = 0
               lngTemp = 0
                 
                 rc = RegSetValueEx(hKey, _
                     ValueName, _
                     0, _
                     REG_DWORD_BIG_ENDIAN, _
                     lngTemp, _
                     lpcbData)
'                 rc = RegSetValueEx(hKey, _
                     ValueName, _
                     0, _
                     REG_DWORD, _
                     lngTemp, _
                     lpcbData)
                     
            End Select
            
            'Return the result code
            'SetValue = True
         
        
        'Close the registry
        RegCloseKey hKey
    End If
    
'    'Return the result code
    CreateValueName = True
    
    'Bypass the error handler
    Exit Function

handler:
    CreateValueName = False
End Function


Private Function GetTrueHkey(ByRef Target As TargetRegistry) As Long
'
    Dim hKey As Long
    Select Case Target
        
        Case cpiCurrentConfig
            
            hKey = HKEY_CURRENT_CONFIG
        
        Case cpiCurrentUser
            
            hKey = HKEY_CURRENT_USER
        
        Case cpiLocalMachine
            
            hKey = HKEY_LOCAL_MACHINE
        
        Case cpiClassesRoot
            
            hKey = HKEY_CLASSES_ROOT
    
    End Select

    GetTrueHkey = hKey
    
End Function

Private Function GetValueNameCount(ByVal HKEYPath As Long _
                                                      , ByVal RootKeyPath As String _
                                                      ) As Long
    
    Dim strRootKey As String
    Dim strTempKey As String
    Dim lngKeyIndex As Long
    Dim blnValuename As Boolean

    strRootKey = RootKeyPath

   On Error GoTo ErrorMe

   Do While (ValueNameExist(HKEYPath, strRootKey, lngKeyIndex))
        lngKeyIndex = lngKeyIndex + 1
        
    Loop
    
    GetValueNameCount = lngKeyIndex + 1
    
    Exit Function

ErrorMe:

   MsgBox "Error"
    
End Function

Private Function ListValueName(PredefinedKey As Long _
                                                            , KeyName As String _
                                                            , Index As Long) As String
    Dim rc As Long
    Dim hKey As Long
    Dim dwIndex As Long
    'Dim lpValueName As String
    Dim lpValueName As String
    Dim lpcbValueName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lpType As Long
    Dim lpData As Byte
    Dim lpcbData As Long
    
    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    
    'Attempt to open the registry
    rc = RegOpenKeyEx(PredefinedKey, KeyName, _
            0, KEY_ALL_ACCESS, hKey)
            
    If rc = ERROR_SUCCESS Then
        'Allocate buffers for lpName
        'lpcbName = Len(ValueName) + 1
        'lpValueName = String$(lpcbName, Chr(0))
                    
        lpcbData = 255
        'lpData = String(lpcbData, Chr(0))
        
        lpcbValueName = 255
        lpValueName = String(lpcbValueName, Chr(0))
        
        'Get the value type first.
        'It will be returned via lpType argument
        'rc = RegQueryValueEx(hKey, _
            ValueName, _
            0, lpType, _
            ByVal lpData, _
            lpcbData)
                    
        'Get the Valuename
        rc = RegEnumValue(hKey, Index, lpValueName, _
                lpcbValueName, lpReserved, lpType, ByVal lpData, lpcbData)
                
        If rc = ERROR_SUCCESS Then
            'Return the result
            ListValueName = Left$(CStr(lpValueName), lpcbValueName)
            'ListValueName = lpValueName
        Else
            ListValueName = ""
        End If
         
         'ListValueName = Trim(ListValueName)
        'Close the registry
        RegCloseKey hKey
    End If
    
    'Bypass the error handler
    Exit Function

handler:
   ListValueName = ""
End Function

Public Function RegistryValueExists(ByVal Target As TargetRegistry, _
                                ByVal RootKey As String, _
                                ByVal SearchValue As String) _
                                As Boolean
    
'Public Function RegistryValueExists(ByVal HKEYPath As cpiHKEYPathConstants _
                                                      , ByVal RootKeyPath As String _
                                                      , ByVal SearchValueName As String) As Boolean
    
    Dim strRootKey As String
    Dim strSubKey As String
    Dim strTempKey As String
    Dim intCount As Long
    Dim lngValueIndex As Long
    Dim lngValueCount As Long
    
    Dim lngHkey As Long
    
    lngHkey = GetTrueHkey(Target)
    
    'This is where we find the reg data...
    strRootKey = RootKey
    
    lngValueCount = GetValueNameCount(lngHkey, strRootKey)
    
    For lngValueIndex = 0 To lngValueCount - 1
    
        strTempKey = ListValueName(lngHkey, strRootKey, lngValueIndex)
                           
         If strTempKey = SearchValue Then
            RegistryValueExists = True
            Exit For
         End If
            
      Next lngValueIndex
    
End Function


Public Function SetRegistryKeyValue(ByVal Target As TargetRegistry _
                                                            , ByVal RootKeyPath As String _
                                                            , ByVal ValueName As String _
                                                            , ByVal Value As Variant) As Boolean

    Dim lngHkey As Long

    lngHkey = GetTrueHkey(Target)

   If SetValue(lngHkey, RootKeyPath, ValueName, Value) Then
      SetRegistryKeyValue = True
   End If

End Function

Private Function SetValue(PredefinedKey As Long, KeyName As String, ValueName As String, _
        Value As Variant) As Boolean
    Dim rc As Long
    Dim hKey As Long
    Dim lpType As Long
    Dim lpcbData As Long
    Dim lpData As String
    
    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    
    'Open the registry
    rc = RegOpenKeyEx(PredefinedKey, _
            KeyName, _
            0, _
            KEY_ALL_ACCESS, _
            hKey)

    If rc = ERROR_SUCCESS Then
        'Create a buffer so we can retrieve the data type of
        'the key. We'll need this to determine which API
        'we should call.
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
        'Get the value type first.
        'It will be returned via lpType argument
        rc = RegQueryValueEx(hKey, _
            ValueName, _
            0, lpType, _
            ByVal lpData, _
            lpcbData)
        
      '  If rc = ERROR_SUCCESS Then
            Select Case lpType
                Case REG_SZ
                    'Use a string data type
                    rc = RegSetValueExString(hKey, _
                        ValueName, _
                        0, _
                        REG_SZ, _
                        CStr(Value), _
                        Len(Value) + 1)
               
               Case REG_BINARY
                    'Use a DWORD data type
                    lpcbData = (Value \ 256) + 2
                    rc = RegSetValueEx(hKey, _
                        ValueName, _
                        0, _
                        REG_BINARY, _
                        CLng(Value), _
                        lpcbData)
                Case REG_DWORD
                    'Use a DWORD data type
                    rc = RegSetValueEx(hKey, _
                        ValueName, _
                        0, _
                        REG_DWORD, _
                        CLng(Value), _
                        lpcbData)
            End Select
            
            'Return the result code
            'SetValue = True
         
      '  Else
            
            'rc = regsetvalueex(hkey,valuename,0,
      '  End If
        
        'Close the registry
        RegCloseKey hKey
    End If
    
'    'Return the result code
    SetValue = True
    
    'Bypass the error handler
    Exit Function

handler:
    SetValue = False
End Function

' private subs
Private Function ValueNameExist(PredefinedKey As Long _
                                                            , KeyName As String _
                                                            , Index As Long) As Boolean
    Dim rc As Long
    Dim hKey As Long
    Dim dwIndex As Long
    'Dim lpValueName As String
    Dim lpValueName As String
    Dim lpcbValueName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lpType As Long
    Dim lpData As Byte
    Dim lpcbData As Long
    
    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    
    'Attempt to open the registry
    rc = RegOpenKeyEx(PredefinedKey, KeyName, _
            0, KEY_ALL_ACCESS, hKey)
            
    If rc = ERROR_SUCCESS Then
        'Allocate buffers for lpName
        'lpcbName = Len(ValueName) + 1
        'lpValueName = String$(lpcbName, Chr(0))
                    
        lpcbData = 255
        'lpData = String(lpcbData, Chr(0))
        
        lpcbValueName = 255
        lpValueName = String(lpcbValueName, Chr(0))
        
        'Get the value type first.
        'It will be returned via lpType argument
        'rc = RegQueryValueEx(hKey, _
            ValueName, _
            0, lpType, _
            ByVal lpData, _
            lpcbData)
                    
        'Get the Valuename
        rc = RegEnumValue(hKey, Index, lpValueName, _
                lpcbValueName, lpReserved, lpType, ByVal lpData, lpcbData)
                
        If rc = ERROR_SUCCESS Then
            'Return the result
            ValueNameExist = True
            'ListValueName = lpValueName
        Else
            'ListValueName = ""
        End If
         
         'ListValueName = Trim(ListValueName)
        'Close the registry
        RegCloseKey hKey
    End If
    
    'Bypass the error handler
    Exit Function

handler:
   'ListValueName = ""
End Function



