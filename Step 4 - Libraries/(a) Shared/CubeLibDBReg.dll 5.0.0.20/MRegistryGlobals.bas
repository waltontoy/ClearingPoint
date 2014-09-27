Attribute VB_Name = "MRegistryGlobals"
Option Explicit

Public Const ERROR_SUCCESS = 0
Public Const REG_DWORD_BIG_ENDIAN = 5

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000

Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)


Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))


    
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'FROM CLinkedTable for UNCT --> Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
        '--- lpData as String <DIFFERENCE>

Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Function HasRegistryAdminRights_F(ByVal RegistryGroup As TargetRegistry, _
                                          ByVal RegistryPath As String, _
                                          ByVal RegistryKey As String, _
                                          ByVal RegistrySetting As String, _
                                          HasAdminRights As Boolean) As String
    
    Dim strReturnValue As String
    Dim lngOpenedKeyHandle As Long
    Dim lngRegOpenKeyExReturnValue As Long
    Dim QueryResult As QueryStringResult
    
    On Error GoTo ErrHandler
    
    HasRegistryAdminRights_F = ""
    
    '----->  Open key for reading and writing
    Select Case RegistryGroup
        Case TargetRegistry.cpiLocalMachine
            lngRegOpenKeyExReturnValue = RegOpenKeyEx(HKEY_LOCAL_MACHINE, RegistryPath & "\" & RegistryKey, 0, KEY_ALL_ACCESS, lngOpenedKeyHandle)
        Case TargetRegistry.cpiCurrentUser
            lngRegOpenKeyExReturnValue = RegOpenKeyEx(HKEY_CURRENT_USER, RegistryPath & "\" & RegistryKey, 0, KEY_ALL_ACCESS, lngOpenedKeyHandle)
    End Select
    
    '----->  User has admin rights?
    If lngRegOpenKeyExReturnValue = 0 Then
        HasAdminRights = True

        '----->  Get key value
        QueryResult = RegQueryStringValue_F(lngOpenedKeyHandle, StripNullTerminator(RegistrySetting))
        If Not QueryResult.StringResult Then
            strReturnValue = ""
        Else
            strReturnValue = QueryResult.StringValue
        End If
    Else
        HasAdminRights = False
        '----->  Open for reading only
    End If
    
    '----->  Close key
    Call RegCloseKey(lngOpenedKeyHandle)
    
    HasRegistryAdminRights_F = strReturnValue
    
ErrHandler:
    
    Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
        Case 0                ' No error; normal exit.
            ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
        Case Else
            HasAdminRights = False
    End Select
End Function

Public Function RegQueryStringValue_F(ByVal hKey As Long, _
                                    ByVal strValueName As String) As QueryStringResult

    Dim lngLengthMinusNUllTerminator As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    Dim lngResult As Long
    
    Dim intData As Integer
    
    Dim strBuffer As String
    
    
    RegQueryStringValue_F.StringResult = True
    RegQueryStringValue_F.StringValue = ""
    
    ' Retrieve information about the key
    lngResult = RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal 0, lngDataBufferSize)
        
    If lngValueType = REG_SZ Then
        ' Create a buffer
        strBuffer = String(lngDataBufferSize, Chr$(0))
        
        ' Retrieve the key's content
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal strBuffer, lngDataBufferSize) <> 0 Then
            RegQueryStringValue_F.StringResult = False
        Else
            lngLengthMinusNUllTerminator = InStr(1, strBuffer, Chr$(0)) - 1
            If lngLengthMinusNUllTerminator > 0 Then
                RegQueryStringValue_F.StringValue = Left$(strBuffer, lngLengthMinusNUllTerminator)
            Else
                RegQueryStringValue_F.StringValue = ""
            End If
            ' RegQueryStringValue_F.StringValue = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
        End If
    ElseIf lngValueType = REG_BINARY Or lngValueType = REG_DWORD Then
        
        ' Retrieve the key's value
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, intData, lngDataBufferSize) <> 0 Then
            RegQueryStringValue_F.StringResult = False
        Else
            RegQueryStringValue_F.StringValue = intData
        End If
    End If
    
End Function


