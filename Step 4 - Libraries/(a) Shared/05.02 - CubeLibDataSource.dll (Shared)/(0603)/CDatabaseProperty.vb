Imports Microsoft.Win32

Public Class CDatabaseProperty
    Private Const PROPERTY_FILE As String = "persistence.txt"
    Private Const REGKEY_CLEARINGPOINT_SETTINGS As String = "Software\Wow6432Node\Cubepoint\Clearingpoint\Settings"
    Private Const REGKEY_CLEARINGPOINT_SETTINGS_XP As String = "Software\Cubepoint\ClearingPoint\Settings"
    Private m_objProp As CProperty

    Public Enum DatabaseType
        SQLSERVER
        ACCESS97
        ACCESS2003
        ORACLE
        MYSQL
    End Enum

    Public Sub New(ByVal filePath As String)
        MyBase.New()

        If filePath = vbNullString Then
            Throw New ClearingPointException("Error in CDatabaseProperty - Persistence path is empty or null string.")
        End If

        Try
            m_objProp = New CProperty(filePath, PROPERTY_FILE)

        Catch ex As Exception
            Throw New ClearingPointException("Error in CDatabaseProperty - " & ex.Message)
        End Try
    End Sub

    Public Function getOutputFilePath() As String
        Return m_objProp.getPropertyKey("OutputFilePath")
    End Function

    Public Function getDatabaseType() As DatabaseType
        Dim dbType As String = m_objProp.getPropertyKey("database")
        Return DirectCast([Enum].Parse(GetType(DatabaseType), dbType), DatabaseType)
    End Function

    Public Function getServerName() As String
        Return m_objProp.getPropertyKey("servername")
    End Function

    Public Function getUserName() As String
        Return m_objProp.getPropertyKey("username")
    End Function

    Public Function getPassword() As String
        Return m_objProp.getPropertyKey("password")
    End Function

    Public Function getDatabasePathFromRegistry() As String
        Dim strDBPath As String
        Dim regKey As RegistryKey

        regKey = Registry.LocalMachine.OpenSubKey(REGKEY_CLEARINGPOINT_SETTINGS, False)

        If Not regKey Is Nothing Then
            strDBPath = regKey.GetValue("MdbPath")
        Else
            strDBPath = vbNullString
        End If

        Return strDBPath
    End Function

    Public Function getDatabasePathFromPersistence() As String
        Return m_objProp.getPropertyKey("MdbPath")
    End Function

    Public Function printDebugTrace() As Boolean
        If m_objProp.getPropertyKey("debug").ToUpper = "TRUE" Then
            Return True
        Else
            Return False
        End If
    End Function

    'TODO: Need to add a registry source for SQL UserName and SQL Data Source
    Public Function GetRegistryKey(ByVal Key As String) As String
        Dim strDBPath As String
        Dim regKey As RegistryKey

        regKey = Registry.LocalMachine.OpenSubKey(REGKEY_CLEARINGPOINT_SETTINGS, False)
        strDBPath = regKey.GetValue(Key)
        AddToTrace("RegKey: " & REGKEY_CLEARINGPOINT_SETTINGS & " DBPath: " & strDBPath)
        If strDBPath Is Nothing AndAlso Len(strDBPath) < 0 Then
            regKey = Registry.LocalMachine.OpenSubKey(REGKEY_CLEARINGPOINT_SETTINGS_XP, False)
            strDBPath = regKey.GetValue(Key)
            AddToTrace("RegKey: " & REGKEY_CLEARINGPOINT_SETTINGS_XP & " DBPath: " & strDBPath)
        End If

        Return strDBPath
    End Function
End Class
