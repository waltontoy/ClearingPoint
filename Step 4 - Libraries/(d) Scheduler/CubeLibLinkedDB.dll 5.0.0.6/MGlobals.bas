Attribute VB_Name = "MGlobals"
'''''Option Explicit
'''''
'''''Const HKEY_LOCAL_MACHINE = &H80000002
'''''
'''''Private Declare Function RegEnumValue Lib "advapi32.dll" _
'''''                Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex _
'''''                As Long, ByVal lpValueName As String, lpcbValueName _
'''''                As Long, ByVal lpReserved As Long, lpType As Long, _
'''''                ByVal lpData As String, lpcbData As Long) As Long
'''''
'''''Private Declare Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long
'''''
'''''Private Declare Function RegOpenKey Lib "advapi32.dll" _
'''''                Alias "RegOpenKeyA" (ByVal HKey As Long, _
'''''                ByVal lpSubKey As String, phkResult As Long) As Long
'''''
'''''Private Declare Function GetComputerName Lib "kernel32" _
'''''                Alias "GetComputerNameA" (ByVal lpBuffer As String, _
'''''                nSize As Long) As Long
'''''
'''''Private Declare Function WNetGetConnection Lib _
'''''                "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName _
'''''                As String, ByVal lpszRemoteName As String, _
'''''                cbRemoteName As Long) As Long
'''''
'''''Public Const KEY_ENCRYPT = ""
'''''
Public Enum QueryResultConstants
    QueryResultError = 0
    QueryResultSuccessful = 1
    QueryResultNoRecord = 2
End Enum
'''''
'''''Public Function FNullField(ByRef Data As Variant) As Variant
'''''    Dim strDataType As String
'''''
'''''    If FIsEmpty(Data) Then
'''''        strDataType = TypeName(Data)
'''''
'''''        If InStr(1, strDataType, "Byte") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Integer") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Long") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Single") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "String") > 0 Then
'''''            Data = ""
'''''        ElseIf InStr(1, strDataType, "Double") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Currency") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Decimal") > 0 Then
'''''            Data = 0
'''''        ElseIf InStr(1, strDataType, "Date") > 0 Then
'''''        ElseIf Trim(strDataType) = "Null" Then
'''''            Data = ""
'''''        ElseIf InStr(1, strDataType, "Field") > 0 Then
'''''            Select Case Data.Type
'''''                Case 20, 14, 5, 6, 3, 205, 201, 131, 2          'Numeric Field Types
'''''                    Data = 0
'''''                Case 16, 21, 19, 18, 17, 4                      'Numeric Field Types
'''''                    Data = 0
'''''                Case 11                                         ' Boolean
'''''                    Data = False
'''''                Case 129, 204, 200, 202, 203, 130               ' String
'''''                    Data = ""
'''''                Case 7, 133, 134                                ' Date
'''''
'''''            End Select
'''''        End If
'''''    End If
'''''    FNullField = Data
'''''End Function
'''''
'''''Private Function FIsEmpty(ByVal Data As Variant) As Boolean
'''''    Dim strDummy As String
'''''
'''''    FIsEmpty = False
'''''
'''''    If IsObject(Data) And Not TypeName(Data) = "Field" Then ' Check if Variable Passed is an Object Variable
'''''        FIsEmpty = True
'''''
'''''    ElseIf IsEmpty(Data) Then ' Check if Variable Passed is Not Initializaed
'''''        FIsEmpty = True
'''''
'''''    ElseIf IsNull(Data) Then ' Check if Variable Passed Contains Invalid Data
'''''        FIsEmpty = True
'''''
'''''    Else
'''''        If IsArray(Data) Then
'''''            If UBound(Data) = 0 And (Data(0) = "" Or IsEmpty(Data(0))) Then
'''''                FIsEmpty = True
'''''            End If
'''''        Else
'''''            strDummy = CStr(Data)
'''''
'''''            If Trim(strDummy) = "" Then
'''''                FIsEmpty = True
'''''            End If
'''''        End If
'''''    End If
'''''End Function
'''''
'''''Public Function GetUNCNameNT(pathName As String) As String
'''''
'''''Dim HKey As Long
'''''Dim hKey2 As Long
'''''Dim exitFlag As Boolean
'''''Dim i As Double
'''''Dim ErrCode As Long
'''''Dim rootKey As String
'''''Dim Key As String
'''''Dim computerName As String
'''''Dim lComputerName As Long
'''''Dim stPath As String
'''''Dim firstLoop As Boolean
'''''Dim Ret As Boolean
'''''
'''''' first, verify whether the disk is connected to the network
'''''If Mid(pathName, 2, 1) = ":" Then
'''''   Dim UNCName As String
'''''   Dim lenUNC As Long
'''''
'''''   UNCName = String$(520, 0)
'''''   lenUNC = 520
'''''   ErrCode = WNetGetConnection(Left(pathName, 2), UNCName, lenUNC)
'''''
'''''   If ErrCode = 0 Then
'''''      UNCName = Trim(Left$(UNCName, InStr(UNCName, _
'''''        vbNullChar) - 1))
'''''      GetUNCNameNT = UNCName & Mid(pathName, 3)
'''''      Exit Function
'''''   End If
'''''End If
'''''
'''''' else, scan the registry looking for shared resources
''''''(NT version)
'''''computerName = String$(255, 0)
'''''lComputerName = Len(computerName)
'''''ErrCode = GetComputerName(computerName, lComputerName)
'''''If ErrCode <> 1 Then
'''''   GetUNCNameNT = pathName
'''''   Exit Function
'''''End If
'''''
'''''computerName = Trim(Left$(computerName, InStr(computerName, _
'''''   vbNullChar) - 1))
'''''rootKey = "SYSTEM\CurrentControlSet\Services\LanmanServer\Shares"
'''''ErrCode = RegOpenKey(HKEY_LOCAL_MACHINE, rootKey, HKey)
'''''
'''''If ErrCode <> 0 Then
'''''   GetUNCNameNT = pathName
'''''   Exit Function
'''''End If
'''''
'''''firstLoop = True
'''''
'''''Do Until exitFlag
'''''   Dim szValue As String
'''''   Dim szValueName As String
'''''   Dim cchValueName As Long
'''''   Dim dwValueType As Long
'''''   Dim dwValueSize As Long
'''''
'''''   szValueName = String(1024, 0)
'''''   cchValueName = Len(szValueName)
'''''   szValue = String$(500, 0)
'''''   dwValueSize = Len(szValue)
'''''
'''''   ' loop on "i" to access all shared DLLs
'''''   ' szValueName will receive the key that identifies an element
'''''   ErrCode = RegEnumValue(HKey, i#, szValueName, _
'''''       cchValueName, 0, dwValueType, szValue, dwValueSize)
'''''
'''''   If ErrCode <> 0 Then
'''''      If Not firstLoop Then
'''''         exitFlag = True
'''''      Else
'''''         i = -1
'''''         firstLoop = False
'''''      End If
'''''   Else
'''''      stPath = GetPath(szValue)
'''''      If firstLoop Then
'''''         Ret = (UCase(stPath) = UCase(pathName))
'''''         stPath = ""
'''''      Else
'''''         Ret = (UCase(stPath) = UCase(Left$(pathName, _
'''''        Len(stPath))))
'''''         stPath = Mid$(pathName, Len(stPath) + 1)
'''''      End If
'''''      If Ret Then
'''''         exitFlag = True
'''''         szValueName = Left$(szValueName, cchValueName)
'''''         GetUNCNameNT = "\\" & computerName & "\" & _
'''''            szValueName & stPath
'''''      End If
'''''   End If
'''''   i = i + 1
'''''Loop
'''''
'''''RegCloseKey HKey
'''''If GetUNCNameNT = "" Then GetUNCNameNT = pathName
'''''
'''''End Function
'''''
'''''Private Function GetPath(st As String) As String
'''''   Dim pos1 As Long, pos2 As Long, pos3 As Long
'''''   Dim stPath As String
'''''
'''''   pos1 = InStr(st, "Path")
'''''   If pos1 > 0 Then
'''''      pos2 = InStr(pos1, st, vbNullChar)
'''''      stPath = Mid$(st, pos1, pos2 - pos1)
'''''      pos3 = InStr(stPath, "=")
'''''      If pos3 > 0 Then
'''''         stPath = Mid$(stPath, pos3 + 1)
'''''         GetPath = stPath
'''''      End If
'''''   End If
'''''End Function
