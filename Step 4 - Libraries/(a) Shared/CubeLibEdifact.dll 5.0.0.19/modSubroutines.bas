Attribute VB_Name = "modSubroutines"
Option Explicit

Private mblnStartTraceFile As Boolean
Public lSavingTrace As Boolean
Public G_strMdbPath As String
Public G_blnTraceOn As Boolean

Public Function IsEDISeparator(ByVal Character As String) As Boolean
    Dim blnReturnValue As Boolean
    
    blnReturnValue = (Character = EDI_SEP_SEGMENT) Or _
                     (Character = EDI_SEP_COMPOSITE_DATA_ELEMENT) Or _
                     (Character = EDI_SEP_DATA_ELEMENT) Or _
                     (Character = EDI_SEP_RELEASE_CHARACTER)
    IsEDISeparator = blnReturnValue
    
    
End Function

Public Function IsInRecordset(ByRef SourceRecordset As ADODB.Recordset, Criteria As String) As Boolean
          Dim blnReturnValue As Boolean
10        On Error GoTo ErrHandler
20        If SourceRecordset.RecordCount > 0 Then SourceRecordset.MoveFirst
30        SourceRecordset.Find Criteria, , adSearchForward
40        If SourceRecordset.EOF Then
50            SourceRecordset.Find Criteria, , adSearchBackward
60        End If
          
70        blnReturnValue = Not (SourceRecordset.BOF Or SourceRecordset.EOF)
80        IsInRecordset = blnReturnValue
          Exit Function
90 ErrHandler:
          
100       AddToTrace "Error In CubeLibEdifact.modSubroutines.IsInRecordset (" & Erl & "," & Err.Number & ") - " & Err.Description
          
End Function

'''''Public Sub DAOOpenTable(ByRef EDIDatabase As DAO.Database, _
'''''                        ByRef EDIRecords As DAO.Recordset, _
'''''                        ByVal TableName As String)
'''''
'''''    Set EDIRecords = EDIDatabase.OpenRecordset(TableName, dbOpenTable)
'''''
'''''End Sub

Public Function GetWholeNumber(ByVal strNumber As String) As String
Dim intDecimalPos As Integer
    intDecimalPos = InStr(1, strNumber, ".")
    
    If intDecimalPos > 0 Then
        GetWholeNumber = Left(strNumber, intDecimalPos - 1)
    Else
        GetWholeNumber = strNumber
    End If
    
End Function

Public Function GetDecimalNumber(ByVal strNumber As String) As String
Dim intDecimalPos As Integer
    intDecimalPos = InStr(1, strNumber, ".")
    
    If intDecimalPos > 0 Then
        GetDecimalNumber = Mid(strNumber, intDecimalPos)
    Else
        GetDecimalNumber = ""
    End If

End Function

Public Sub AddToTrace(ByVal strTraceString As String)
    Dim intFreeFile As Integer
    
    On Error GoTo ErrHandler

    G_strMdbPath = NoBackSlash(GetSetting("ClearingPoint", "Settings", "MdbPath", "C:\Program Files\Cubepoint\ClearingPoint"))
    
    If G_strMdbPath = vbNullString And G_blnTraceOn = False Then Exit Sub
    
20  lSavingTrace = True
    
25  intFreeFile = FreeFile()
    
30  If Len(Dir(G_strMdbPath & "\TraceFile.txt")) Then
35      If FileLen(G_strMdbPath & "\TraceFile.txt") >= 360000 Then
'40          Name G_strMdbPath & "\TraceFile.txt" As G_strMdbPath & "\TraceFile" & Format(Now, "ddMMyyyyhhmm") & ".txt"
            Name G_strMdbPath & "\TraceFile.txt" As RandomTraceFileName
            
45          Open G_strMdbPath & "\TraceFile.txt" For Output As #intFreeFile
            
50          mblnStartTraceFile = True
        Else
55          Open G_strMdbPath & "\TraceFile.txt" For Append As #intFreeFile
        End If
    Else
60      Open G_strMdbPath & "\TraceFile.txt" For Output As #intFreeFile
        
65      mblnStartTraceFile = True
    End If
    
'70  If mblnStartTraceFile Then
'75      Print #intFreeFile, AdditionalTracefileInfo
'
'80      mblnStartTraceFile = False
'    End If
    
85  'If blnSend Then
90  '    Print #intFreeFile, Now & ". Send   : " & strTraceString
    'Else
95      Print #intFreeFile, Now & ". Receive: " & strTraceString
    'End If
    
100 Close #intFreeFile
    
ErrHandler:
    
    lSavingTrace = False
    
    Select Case Err.Number    ' Used Select Case control structure for easy maintenance in case of new reported errors [Andrei]
        Case 0                ' No error; normal exit.
            ' Do nothing; included just to prevent Case Else from handling Err.Number = 0.
        Case Else
            
    End Select
End Sub

Public Function RandomTraceFileName() As String
    Dim lngCtr As Long
    Dim strRandomNumber As String
    Dim strRandomFileName As String
    
    
    strRandomFileName = G_strMdbPath & "\TraceFile" & Format(Now, "ddMMyyyyhhmm")
    
Repeat_Generator:
    Randomize
    
    For lngCtr = 1 To 4
        strRandomNumber = strRandomNumber & Int((9 * Rnd) + 0)
    Next lngCtr
    
    strRandomFileName = strRandomFileName & "_" & strRandomNumber & ".txt"
    
    If (Len(Dir(strRandomFileName)) > 0) Then
        GoTo Repeat_Generator
    End If
    
    RandomTraceFileName = strRandomFileName
End Function
'''''
'''''Public Function NoBackSlash(ByVal cString As String) As String
'''''    Do While Right(cString, 1) = "\"
'''''        cString = Left(cString, Len(cString) - 1)
'''''    Loop
'''''
'''''    NoBackSlash = cString
'''''End Function
