Imports System.IO

Module MTrace

    Private Const CONST_PREFIX As String = "CubelibDataSource Error Log: "

    Public Sub AddToTrace(ByVal Message As String, Optional ByVal DebugOnly As Boolean = False)
        Dim fileReName As String = G_ObjProp.getDatabasePathFromPersistence & "\DatasourceTracefile_" & Format(Now, "yyyyMMdd_hhmmss") & ".log"
        Dim fileName As String = G_ObjProp.getDatabasePathFromPersistence & "\DatasourceTracefile.log"
        Dim info As New FileInfo(fileName)
        Dim sw As StreamWriter = Nothing

        Try
            If (info.Exists) Then
                If info.Length < 360000 Then
                    sw = info.AppendText()
                Else
                    info.CopyTo(fileReName)
                    info.Delete()

                    sw = info.CreateText()
                End If

            Else
                sw = info.CreateText()
            End If

            If DebugOnly Then
                If G_ObjProp.printDebugTrace Then
                    sw.WriteLine(CONST_PREFIX & Format(Now, "yyyy-MM-dd hh:mm:ss") & " : " & Message)
                End If
            Else
                sw.WriteLine(CONST_PREFIX & Format(Now, "yyyy-MM-dd hh:mm:ss") & " : " & Message)
            End If

        Catch e As Exception
            Err.Clear()
        Finally
            If Not sw Is Nothing Then
                sw.Close()
            End If
        End Try

    End Sub

End Module
