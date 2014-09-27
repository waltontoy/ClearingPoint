Imports System.Data.Common

Public Class DNetConnections

    Public Const DATABASE_SADBEL As String = "SADBEL"
    Public Const DATABASE_DATA = "DATA"
    Public Const DATABASE_EDIFACT = "EDIFACT"
    Public Const DATABASE_SCHEDULER = "SCHEDULER"
    Public Const DATABASE_TEMPLATE = "TEMPLATE"
    Public Const DATABASE_TARIC = "TARIC"
    Public Const DATABASE_HISTORY = "HISTORY"
    Public Const DATABASE_REPERTORY = "REPERTORY"
    Public Const DATABASE_EDI_HISTORY = "EDI_HISTORY"

    Private m_colDatabase As New Dictionary(Of String, DbConnection)

    Public Sub addDatabase(ByVal key As String, ByRef database As DbConnection, Optional ByVal removeExisting As Boolean = False)
        If m_colDatabase.ContainsKey(key) Then
            If (removeExisting) Then
                deleteDatabase(key)
                m_colDatabase.Add(key, database)
            End If
        Else
            m_colDatabase.Add(key, database)
        End If
    End Sub

    Public Sub deleteDatabase(ByVal key As String)
        m_colDatabase.Remove(key)
    End Sub

    Public Function getDatabase(ByVal key As String) As DbConnection
        If m_colDatabase.ContainsKey(key) Then
            Return m_colDatabase.Item(key)
        End If

        Return Nothing
    End Function
End Class
