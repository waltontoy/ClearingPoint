Attribute VB_Name = "MTemplateGlobal"
Option Explicit

' Function #4
' Function Name: G_GetChunkData
' Description: get next 100 bytes of current  binary values from database field
' Syntax: [blnFlag] = G_GetChunkData([fldObject],[varData])
' Scope: Public
' Fan-In: <none>
Public Function G_GetChunkData(ByRef ActiveField As ADODB.Field, _
                               ByRef varData As Variant) As Boolean

    Const CHUNK_SIZE = 100

    Dim lngOffset As Long
    Dim lngObjectSize As Long
    Dim varChunk As Variant

    On Error GoTo ERROR_CHUNK
    varData = Empty
    lngObjectSize = ActiveField.ActualSize

    Do While (lngOffset < lngObjectSize)
        varChunk = ActiveField.GetChunk(CHUNK_SIZE)
        varData = varData & varChunk
        lngOffset = lngOffset + CHUNK_SIZE
    Loop

    G_GetChunkData = True

    Exit Function

ERROR_CHUNK:
    G_GetChunkData = False
End Function

' Function #5
' Function Name: G_GetOLEFromDB
' Description: get recordset with binary data (not yet implemented)
' Syntax: [blnFlag] = G_GetOLEFromDB([conObject],[strSQL],[varData])
' Scope: Public
' Fan-In: <none>
Public Function G_GetOLEFromDB(ByRef ActiveConnection As ADODB.Connection, _
                               ByRef CommandText As String, _
                               ByRef varData As Variant) As Boolean

    Dim rstOLE As ADODB.Recordset

    Set rstOLE = New ADODB.Recordset

    On Error GoTo ERROR_QUERY
    rstOLE.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic

    If (rstOLE.EOF = False) Then
        ' not yet implemented
    End If

    Set rstOLE = Nothing
    G_GetOLEFromDB = True
    Exit Function

ERROR_QUERY:
    Set rstOLE = Nothing
    G_GetOLEFromDB = False
End Function

' Function #6
' Function Name: G_SetEnclosedChr
' Description: set proper enclosed character for each data type
' Syntax: [varData] = G_SetEnclosedChr([varDataType])
' Scope: Public
' Fan-In: <none>
Public Function G_SetEnclosedChr(ByRef varValue As Variant) As Variant

    Dim varReturn As Variant
    Dim strTypeName As String
    
    On Error GoTo ERROR_VALUE
    strTypeName = UCase$(TypeName(varValue))
    varValue = Replace(varValue, "'", "''", , , vbTextCompare)

    Select Case strTypeName
        Case "STRING"
            varReturn = "'" & varValue & "'"
        
        Case "DATE"
            varReturn = "CDate('" & varValue & "')"
        
        Case Else
            varReturn = varValue
    End Select

    G_SetEnclosedChr = varReturn
    Exit Function

ERROR_VALUE:
    G_SetEnclosedChr = varValue
End Function

' Function #7
' Function Name: Translate
' Description: get corresponding language value based on ID passed
' Syntax: [varData] = Translate([varDataType])
' Scope: Public
' Fan-In: StripNullTerminator
Public Function Translate(ByVal variToTranslate As Variant, _
                          ByVal lngiResourceHandler As Long) As String
    
    Dim strTranslated As String * 520
    
    If (((variToTranslate Like "*[A-Z]*") Or (variToTranslate Like "*[a-z]*")) = True) Then
        Translate = variToTranslate
    
    ElseIf (((variToTranslate Like "*[A-Z]*") Or (variToTranslate Like "*[a-z]*")) = False) Then
        If (Trim(variToTranslate) <> "") And (Trim(variToTranslate) <> "...") Then
            LoadString lngiResourceHandler, CLng(variToTranslate), strTranslated, 520
        End If
        
        strTranslated = StripNullTerminator(strTranslated)
        Translate = RTrim$(strTranslated)
    End If

End Function


