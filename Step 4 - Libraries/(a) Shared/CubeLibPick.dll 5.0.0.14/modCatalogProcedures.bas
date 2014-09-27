Attribute VB_Name = "modCatalogProcedures"
Option Explicit

    Public blnOperator As Boolean 'allan pick

    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type
    
    Public CatalogForm As frmCatalog
    
    Private Declare Function SetSystemTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME) As Long
    Private Declare Function GetSystemTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME) As Long
    

      Private Const GWL_WNDPROC = -4
      Private Const WM_GETMINMAXINFO = &H24

      Private Type POINTAPI
          x As Long
          y As Long
      End Type

      Private Type MINMAXINFO
          ptReserved As POINTAPI
          ptMaxSize As POINTAPI
          ptMaxPosition As POINTAPI
          ptMinTrackSize As POINTAPI
          ptMaxTrackSize As POINTAPI
      End Type

      Global lpPrevWndProc As Long
      Global gHW As Long

      Private Declare Function DefWindowProc Lib "user32" Alias _
         "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function CallWindowProc Lib "user32" Alias _
         "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
          ByVal hwnd As Long, ByVal Msg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function SetWindowLong Lib "user32" Alias _
         "SetWindowLongA" (ByVal hwnd As Long, _
          ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
      Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias _
         "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
          ByVal cbCopy As Long)
      Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias _
         "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
          ByVal cbCopy As Long)
          
    'This was declared locally rather than globally to support multiple instances of
    'picklist. => IAN 09-15-04
    
    Private g_FormWid
    Private g_FormHgt
    
    Public Const G_MAIN_PASSWORD = "wack2"

Public Function GetSQLFromTable(ByVal SQLCommand As String) As String
    Dim strSQLCommand As String
    Dim lngPos As Long
    
    strSQLCommand = UCase$(Trim$(SQLCommand))
    
    strSQLCommand = Mid(strSQLCommand, InStr(1, strSQLCommand, "FROM"))
    
    lngPos = 1
    Do While Mid(strSQLCommand, lngPos, 1) <> " "
        
        lngPos = lngPos + 1
    Loop
    
    strSQLCommand = Mid(strSQLCommand, lngPos + 1)
    
    If Left$(strSQLCommand, 1) = "[" Then
        strSQLCommand = Mid(strSQLCommand, 2)
        
        strSQLCommand = Left$(strSQLCommand, InStr(1, strSQLCommand, "]") - 1)
    Else
        strSQLCommand = Left$(strSQLCommand, InStr(1, strSQLCommand, " ") - 1)
    End If
    
    GetSQLFromTable = strSQLCommand
End Function


Public Function RegenerateSQL(ByVal SQL As String, Optional ByVal blnNoTop As Boolean) As String
    Dim strParserDummyLeft As String
    Dim strParserDummyMidle As String
    Dim strParserDummyRight As String
    
    Dim blnSQLHasTop As Boolean
    Dim strShiftString As String
    Dim blnFirstCharacter As Boolean
    
    Dim strSelect() As String
    
    blnSQLHasTop = False
    strParserDummyRight = SQL
    strShiftString = ""
        
     
        
    strParserDummyLeft = Left(strParserDummyRight, InStr(1, strParserDummyRight, " "))
    strParserDummyRight = Mid(strParserDummyRight, InStr(1, strParserDummyRight, " ") + 1)
  
    Do While Left(strParserDummyRight, 1) = " "
        strParserDummyLeft = strParserDummyLeft & " "
        strParserDummyRight = Mid(strParserDummyRight, 2)
    Loop
            
'     Debug.Print strParserDummyLeft
'     Debug.Print strParserDummyRight
    ' INSERT TOP 1
    strShiftString = Left(strParserDummyRight, 1)
    blnFirstCharacter = True
      'Debug.Print SQL
    
    Do While True
    
        strParserDummyMidle = Left(strParserDummyRight, InStr(1, strParserDummyRight, " ") - 1)
        
        Select Case UCase(strShiftString)
            Case " "
                strParserDummyLeft = strParserDummyLeft & " "
                strParserDummyRight = Mid(strParserDummyRight, 2)
                                            
                strShiftString = Left(strParserDummyRight, 1)
                
            Case "ALL", "DISTINCT", "DISTINCTROW", "TOP"
                
                If UCase(strShiftString) = "TOP" Then
                    blnSQLHasTop = True
                    
                    strParserDummyLeft = strParserDummyLeft & strShiftString & " 1 "
                    strParserDummyRight = Trim(strParserDummyRight)
                    strParserDummyRight = Mid(strParserDummyRight, InStr(1, strParserDummyRight, " ") + 1)
                    strParserDummyRight = Trim(strParserDummyRight)
                Else
                    strParserDummyLeft = strParserDummyLeft & strShiftString & " "
                    strParserDummyRight = Trim(strParserDummyRight) 'Mid(strParserDummyRight, InStr(1, strParserDummyRight, " ") + 1)
                End If
                
                strShiftString = Left(strParserDummyRight, 1)
                strParserDummyRight = Mid(strParserDummyRight, 2)
            Case Else
                'Case Else
                If Not blnFirstCharacter Then
                    strShiftString = strShiftString & Left(strParserDummyRight, 1)
                Else
                    blnFirstCharacter = False
                End If
                strParserDummyRight = Mid(strParserDummyRight, 2)
                
                If Len(strShiftString) = 11 Or UCase(strShiftString) = "FROM" Then
                    strParserDummyRight = strShiftString & strParserDummyRight
                    
                    Exit Do
                End If
                'Exit Do
         
        End Select
    Loop
    
    If blnSQLHasTop Then
        RegenerateSQL = strParserDummyLeft & " " & strParserDummyRight
    Else
      ' changes HERE->
      If blnNoTop Then
        RegenerateSQL = strParserDummyLeft & "  " & strParserDummyRight
      Else
        'RegenerateSQL = strParserDummyLeft & " TOP 1 " & strParserDummyRight
        If (InStr(1, UCase$(strParserDummyRight), "SELECT ", vbTextCompare) > 0) Then
            
            strSelect = Split(strParserDummyRight, "SELECT ", , vbTextCompare)
            RegenerateSQL = "SELECT " & strSelect(0) & " TOP 1 " & strSelect(1)
        ElseIf (InStr(1, UCase$(strParserDummyRight), "SELECT ", vbTextCompare) = 0) Then
        
            RegenerateSQL = strParserDummyLeft & " TOP 1 " & strParserDummyRight
        
        End If
      End If
    End If
            
    ' Add all other fields into SQL if not all fields are selected in the original SQL
    strParserDummyRight = RegenerateSQL
    strParserDummyLeft = Left(strParserDummyRight, InStr(1, UCase(strParserDummyRight), " FROM ") - 1)
    strParserDummyRight = Mid(strParserDummyRight, InStr(1, UCase(strParserDummyRight), " FROM "))
    'Debug.Print SQL
    If InStr(1, strParserDummyLeft, "*") <= 0 Then
        RegenerateSQL = strParserDummyLeft & ", *" & strParserDummyRight
    End If
End Function

Public Function GetRecordCriteria(ByVal PKBaseFieldName As String, ByVal PKFieldValue As String, ByVal PKBaseFieldDataType As Long, Optional BaseTable As String = "") As String
    Dim clsSQLQuotes As CRecordset
    
    Dim strBaseTable As String
    Dim strPKCondition As String
     
    strPKCondition = " "
    
    If Trim(BaseTable) <> "" Then
        strBaseTable = Trim(BaseTable) & "."
    Else
        strBaseTable = ""
    End If
    
    Select Case PKBaseFieldDataType
        ' adSmallInt = 2
        ' adInteger = 3
        ' adSingle = 4
        ' adDouble = 5
        ' adCurrency = 6
        ' adIDispatch = 9
        ' adDecimal = 14
        ' adTinyInt= 16
        ' adUnsignedTinyInt = 17
        ' adUnsignedSmallInt = 18
        ' adUnsignedInt = 19
        ' adBigInt = 20
        ' adUnsignedBigInt = 21
        ' adBinary = 128
        ' adNumeric = 131
        ' adChapter = 136
        ' adPropVariant= 138
        ' adLongVarBinary = 205
        Case adUnsignedInt, adUnsignedSmallInt, adInteger, adUnsignedTinyInt, _
                adSingle, adDouble, adCurrency, adNumeric, adDecimal, _
                adLongVarBinary, adIDispatch, adPropVariant, adChapter, _
                adBinary, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt
                
            strPKCondition = strBaseTable & "[" & PKBaseFieldName & "]" & " = " & PKFieldValue & " "
            
        ' adBSTR = 8                                            ' String
            ' dbText = 10
            ' dbMemo = 12
            ' adChar = 129
            ' adWChar = 130
            ' adVarChar = 200
            ' adLongVarChar = 201
            ' adVarWChar = 202
            ' adLongVarWChar= 203
            ' adVarBinary = 204
            Case adChar, adLongVarChar, adVarBinary, adVarChar, _
                    adVarWChar, adWChar, 10, 12, adBSTR
                        
            strPKCondition = strBaseTable & "[" & PKBaseFieldName & "]" & " = '" & ProcessQuotes(PKFieldValue) & "' "
                        
        Case 205    ' OLE Object
            ' Do nothing
        Case Else
            ' Do nothing
    End Select
    
    GetRecordCriteria = strPKCondition
End Function

Public Function PrintTimeNow(Optional Caption As String = "") As String
    Dim TimeNow As SYSTEMTIME
    Static StaticCount As Long
    
    StaticCount = StaticCount + 1
    
    Call GetSystemTime(TimeNow)
    With TimeNow
        PrintTimeNow = Space(40 - Len(Caption)) & Caption & ":: " & StaticCount & "::" & .wMinute & ":" & .wSecond & ":" & .wMilliseconds
        Debug.Print PrintTimeNow
    End With
End Function

'''''
'''''Public Function RstCopy(ByRef Source As ADODB.Recordset _
'''''                                             , ByRef Disconnected As Boolean _
'''''                                             , ByRef RecordStart As Long _
'''''                                             , ByRef RecordEnd As Long _
'''''                                             , ByRef AbsolutePosition As Long _
'''''                                             , Optional ByVal FieldOnly As Boolean)
''''''                                    , ByRef Destination As ADODB.Recordset
'''''
'''''   Dim rstTemp As ADODB.Recordset
'''''   Dim intIndex As Long
'''''   Dim intPropertyIdx As Long
'''''   Dim fldTemp As ADODB.Field
'''''   Dim lngSourcePos As Long
'''''
'''''  If Not Disconnected Then
'''''   Set RstCopy = Source
'''''   Exit Function
'''''  End If
'''''
'''''   'if rstpatch'
'''''   'If ErrorPatch(Source) Then Exit Function
'''''
'''''   Set rstTemp = New ADODB.Recordset
'''''   ' STEP -1 copy all the fields properties
'''''   For intIndex = 0 To Source.Fields.Count - 1
'''''      rstTemp.Fields.Append Source.Fields(intIndex).Name, Source.Fields(intIndex).Type _
'''''                                                   , Source.Fields(intIndex).DefinedSize, Source.Fields(intIndex).Attributes
'''''   Next intIndex
'''''
'''''   rstTemp.Open
'''''
'''''   If FieldOnly Then
'''''      rstTemp.AddNew
'''''      rstTemp.Update
'''''      GoTo rstEnd
'''''   End If
'''''
'''''If Source.RecordCount = 0 Or Source.RecordCount = -1 Then GoTo rstEnd
'''''   lngSourcePos = Source.AbsolutePosition
'''''   Source.MoveFirst
'''''   Source.Move RecordStart '+ 1
'''''
'''''   ' copy all the record values of the source
'''''   For intIndex = RecordStart To RecordEnd
'''''   rstTemp.AddNew
'''''      For Each fldTemp In Source.Fields
'''''         rstTemp.Fields(fldTemp.Name) = fldTemp.Value
'''''      Next
'''''      Source.MoveNext
'''''   Next intIndex
'''''
'''''   rstTemp.Update
'''''
'''''   rstTemp.AbsolutePosition = AbsolutePosition
'''''   If Source.EOF Then Source.MoveLast
'''''   Source.AbsolutePosition = lngSourcePos
'''''
'''''rstEnd:
'''''   Set RstCopy = rstTemp
'''''   Set rstTemp = Nothing
'''''End Function


Public Function ErrorPatch(ByRef rstiRecord As ADODB.Recordset) As Boolean
   ' patches
   If rstiRecord.RecordCount = 0 Then
      If rstiRecord.EOF Or rstiRecord.BOF Or rstiRecord.AbsolutePosition = adPosUnknown Then
         ErrorPatch = True
      End If
   End If

End Function

Public Function rstDBCopy(ByRef coniConnection As ADODB.Connection _
                                                , ByRef RunSQL As String _
                                                , ByRef CursorType As CursorTypeEnum _
                                                , ByRef LockType As LockTypeEnum _
                                                , ByRef Disconnected As Boolean) As ADODB.Recordset
   
   Dim rstTemp As ADODB.Recordset
   Dim strOrderBy As String
   Dim intOrderLoc As Integer
   Dim intOrderEndLoc As Integer
   
   Set rstTemp = New ADODB.Recordset
   
   On Error GoTo ERROR_MSG
   
   'mark 12052002, temporary only...for auto-positioning & sorting
    intOrderLoc = InStr(1, UCase(RunSQL), "ORDER BY")
    If intOrderLoc > 0 Then
        intOrderEndLoc = InStr(intOrderLoc + 9, RunSQL, " ")
        If intOrderEndLoc = 0 Then
            intOrderEndLoc = Len(RunSQL) + 1
        End If
        strOrderBy = Mid(RunSQL, intOrderLoc, intOrderEndLoc - intOrderLoc)
        RunSQL = Replace(RunSQL, Trim(strOrderBy), " ")
        RunSQL = RunSQL & " " & strOrderBy & " "
    End If
   '-------------------------------------------------------------------
   
   ADORecordsetOpen RunSQL, coniConnection, rstTemp, CursorType, LockType
   'rstTemp.Open RunSQL, coniConnection, CursorType, LockType
   
   If Disconnected Then
      Set rstDBCopy = RstCopy(rstTemp, True, 0, rstTemp.RecordCount - 1, 1)
   Else
      Set rstDBCopy = rstTemp
   End If
    
   Set rstTemp = Nothing
   
   Exit Function
   
ERROR_MSG:

   MsgBox Err.Description, vbInformation, "Catalog"
   Resume Next

End Function

'   For intIndex = 0 To Source.Fields.Count
'      For intPropertyIdx = 0 To Source.Fields(intIndex).Properties.Count - 1
'         ' copy all field properties values
'         rstTemp.Fields(intIndex).Properties(intPropertyIdx).Value _
'            = Source.Fields(intIndex).Properties(intPropertyIdx).Value
'      Next intPropertyIdx
'   Next intIndex
         
'   If Not (IsMissing(RecordStart)) Then lngRecordStart = RecordStart
'   If Not (IsMissing(RecordEnd)) Then lngRecordEnd = RecordEnd

'Public Function OpenRecordsForGrid(Source As String, conToUse As ADODB.Connection, rstToOpen As ADODB.Recordset, CursorType As CursorTypeEnum, LockType As LockTypeEnum, ByRef GridSeed As CGridSeed, Optional lngCacheSize As Long = 1, Optional ByVal MakeOffline As Boolean = False, Optional ByVal AddTag As Boolean = False) As String
''Public Function OpenRecordsForGrid(Source As String, conToUse As ADODB.Connection, rstToOpen As ADODB.Recordset, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Optional lngCacheSize As Long = 1, Optional ByVal MakeOffline As Boolean = False, Optional ByVal AddTag As Boolean = False) As ADODB.Recordset
'    Dim rstDummy As ADODB.Recordset
'    Dim fldDummy As ADODB.Field
'    Dim rstOutput As ADODB.Recordset
'
'    Dim strSourceLeftOfFrom As String
'    Dim strSourceRightOfFrom As String
'    Dim strOriginalSource As String
'
'    Dim strFinalSource As String
'
'    On Error GoTo ERROR_HANDLER_BOOKMARK
'
'    If Not rstToOpen Is Nothing Then
'        If rstToOpen.State = adStateOpen Then
'            rstToOpen.Close
'        End If
'        Set rstToOpen = Nothing
'    End If
'    Set rstToOpen = New ADODB.Recordset
'
'    If Not rstOutput Is Nothing Then
'        If rstOutput.State = adStateOpen Then
'            rstOutput.Close
'        End If
'        Set rstOutput = Nothing
'    End If
'    Set rstOutput = New ADODB.Recordset
'
'    rstToOpen.CacheSize = lngCacheSize
'
'    If MakeOffline = True Then
'        'rstToOpen.Open Source, conToUse, CursorType, LockType
'
'        'Set rstToOpen.ActiveConnection = Nothing
'        rstOutput.CursorLocation = adUseClient
'        rstToOpen.CursorLocation = adUseClient
'
'        If AddTag Then
'            strOriginalSource = Source
'            Source = RegenerateSQL(Source)
'
'            Set rstDummy = New ADODB.Recordset
'            rstDummy.Open Source, conToUse, CursorType, LockType
'
'            With rstToOpen
'
'                'For Each fldDummy In rstDummy.Fields
'                '    rstOutput.Fields.Append fldDummy.Name, fldDummy.Type, fldDummy.DefinedSize, fldDummy.Attributes
'                'Next
'                'rstOutput.Open
'
'                'Set OpenRecordsForGrid = rstOutput
'
'                ' Create Recordset based on SQL with additional field "Tag"
'                rstToOpen.Source = strOriginalSource
'                rstToOpen.Source = Source
'                For Each fldDummy In rstDummy.Fields
'                    rstToOpen.Fields.Append fldDummy.Name, fldDummy.Type, fldDummy.DefinedSize, fldDummy.Attributes
'
'                    ' Needed to add WHERE clause in retrieveing specific record based on PK field
'                    ' Error handler is placed to handle fields which do not belong in the Grid recordset
'                    On Error Resume Next
'                    GridSeed.GridColumns(fldDummy.Name).ColumnBaseFieldName = fldDummy.Properties(0).Value
'                    On Error GoTo 0
'                Next
'
'                rstToOpen.Fields.Append "Tag", adVarWChar, 1, adFldIsNullable
'                rstToOpen.Open
'
'                If Not rstDummy Is Nothing Then
'                    If rstDummy.State = adStateOpen Then
'                        rstDummy.Close
'                    End If
'                    Set rstDummy = Nothing
'                End If
'                Set rstDummy = New ADODB.Recordset
'                rstDummy.Open strOriginalSource, conToUse, CursorType, LockType
'
'                ' Populate Recordset for Grid with data
'                If rstDummy.RecordCount > 0 Then
'                    rstDummy.MoveFirst
'
'                    Do Until rstDummy.EOF
'                        .AddNew
'
'                        For Each fldDummy In rstDummy.Fields
'                            .Fields(fldDummy.Name).Value = rstDummy.Fields(fldDummy.Name).Value
'                        Next
'
'                        !Tag = "O"
'                        .Update
'
'                        rstDummy.MoveNext
'                    Loop
'                End If
'            End With
'
'            rstDummy.Close
'            Set rstDummy = Nothing
'
'        Else
'            ' Create Recordset based on SQL
'            rstToOpen.Open Source, conToUse, CursorType, LockType
'
'            Set rstToOpen.ActiveConnection = Nothing
'
'            rstToOpen.Source = Source
'            For Each fldDummy In rstToOpen.Fields
'                rstOutput.Fields.Append fldDummy.Name, fldDummy.Type, fldDummy.DefinedSize, fldDummy.Attributes
'            Next
'            rstOutput.Open
'
'            'Set OpenRecordsForGrid = rstOutput
'        End If
'    Else
'        rstToOpen.Open Source, conToUse, CursorType, LockType
'
'        For Each fldDummy In rstToOpen.Fields
'            rstOutput.Fields.Append fldDummy.Name, fldDummy.Type, fldDummy.DefinedSize, fldDummy.Attributes
'        Next
'        rstOutput.Open
'    End If
'
'    Set rstOutput = Nothing
'
'    On Error GoTo 0
'
'    Exit Function
'
'ERROR_HANDLER_BOOKMARK:
'    Select Case Err.Number
'        Case -2147467259
'            Resume
'        Case Else
'            Err.Raise Err.Number, , Err.Description
'    End Select
'End Function


'The WIDTH and HEIGHT were just passed to this procedure instead of declaring a global
'variable in purpose of multiple instances of picklist  => IAN 09-15-04

Public Sub Hook(FormWid, FormHgt)
          
    g_FormWid = FormWid
    g_FormHgt = FormHgt
    
    'Start subclassing.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
       AddressOf WindowProc)
       
End Sub

Public Sub Unhook()
    Dim temp As Long

    'Cease subclassing.
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO

    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = g_FormWid
        MinMax.ptMinTrackSize.y = g_FormHgt

'              'Specify new maximum size for window.
        'MinMax.ptMaxTrackSize.x = g_FormWid * 3
        'MinMax.ptMaxTrackSize.y = g_FormHgt * 3

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
           wParam, lParam)
    End If
End Function

Public Function IsCompiled() As Boolean
    On Error GoTo RUNNING_IN_IDE
    Debug.Print 1 / 0
    
    IsCompiled = True
    
RUNNING_IN_IDE:
    Exit Function
End Function

