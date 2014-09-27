Public Class CProperty

    Private Const PROPERTY_COMMENT As String = "#"
    Private Const KEY_VALUE_SEPARATOR As String = "="

    Private m_arrProperties() As String

    Public Sub New(ByVal filePath As String, ByVal propFileName As String)
        MyBase.New()
        Dim objReader As System.IO.StreamReader

        filePath = IIf(filePath.EndsWith("\"), filePath + propFileName, filePath + "\" + propFileName)
        Try
            objReader = New System.IO.StreamReader(filePath)
            m_arrProperties = objReader.ReadToEnd.Split(vbCrLf)
        Catch ex As Exception
            m_arrProperties = My.Resources.persistence.Split(vbCrLf)
        End Try
    End Sub

    Public Function getPropertyKey(ByVal propertyPath As String) As String

        For Each prop As String In m_arrProperties
            If ((prop <> vbNullString) AndAlso Not prop.StartsWith(PROPERTY_COMMENT)) Then
                Dim arrKeyValue() As String = prop.Split(KEY_VALUE_SEPARATOR)
                If arrKeyValue.Length = 2 Then
                    If String.Equals(arrKeyValue(0).Trim.ToUpper, propertyPath.Trim.ToUpper) Then
                        Return arrKeyValue(1)
                        Exit Function
                    End If
                End If
            End If
        Next
        getPropertyKey = vbNullString
    End Function
End Class
