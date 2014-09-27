Attribute VB_Name = "modDBFunction"
Option Explicit

Public Function SearchRecord(ByRef ActiveConnection As ADODB.Connection, _
                                                    ByVal TableName, _
                                                    ByVal SearchField, _
                                                    ByVal SearchValue) As Boolean
'
    Dim strSql As String
    Dim lngRecordsAffected As Long
    Dim rstDB As ADODB.Recordset
    
    On Error GoTo ERROR_SEARCH
    SearchField = Trim$(SearchField)
    If ((SearchField <> "") And (SearchValue <> "")) Then
        If (Len(SearchField) > 2) Then
            If ((Left$(SearchField, 1) <> "[") And (Right$(SearchField, 1) <> "]")) Then
                SearchField = "[" & SearchField & "]"
            End If
        End If
    
        strSql = "SELECT TOP 1 " & SearchField & " FROM " & TableName & " WHERE " & SearchField & "=" & Trim$(SearchValue)
        
        ' MUCP-159 - Start
        ADORecordsetOpen strSql, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic

        SearchValue = Not (rstDB.EOF And rstDB.BOF)

'''''        Set rstDB = ActiveConnection.Execute(strSql, lngRecordsAffected)
'''''
'''''        If (rstDB.EOF = False) Then
'''''            SearchValue = True
'''''        ElseIf (rstDB.EOF = True) Then
'''''            SearchValue = False
'''''        End If
        ' MUCP-159 - End
    End If
    
    Exit Function
    
ERROR_SEARCH:
    SearchValue = False
End Function

'''''Public Function G_FNullField(ByRef ActiveField As ADODB.Field) As Variant
''''''
'''''    Dim varNullValue As Variant
'''''    Dim blnNull As Boolean
'''''
'''''    Select Case ActiveField.Type
'''''
'''''        Case adVarWChar, adChar, adBSTR
'''''
'''''            varNullValue = ""
'''''
'''''        Case adInteger, adUnsignedTinyInt, adSingle, adDouble, adCurrency, adNumeric, adDecimal, adLongVarBinary, adIDispatch, adLongVarWChar, _
'''''                    adPropVariant, adChapter, adBoolean, adBinary, adBigInt, adSmallInt
'''''
'''''            varNullValue = 0
'''''
'''''        Case adDate, adDBDate
'''''
'''''            varNullValue = #1/1/1900#
'''''
'''''        Case adDBTime, adFileTime, adDBFileTime, adDBTimeStamp
'''''
'''''            varNullValue = #12:00:00 AM#
'''''
'''''        Case Else
'''''
'''''            varNullValue = ""
'''''
'''''    End Select
'''''
'''''    blnNull = IsNull(ActiveField.Value)
'''''
'''''    If (blnNull = False) Then
'''''        G_CheckNull = ActiveField.Value
'''''    ElseIf (blnNull = True) Then
'''''        G_CheckNull = varNullValue
'''''    End If
'''''
'''''End Function


