Attribute VB_Name = "mod_UserSetUp"
Option Explicit


    Public g_conSADBEL As ADODB.Connection          'Public datSADBEL As dao.Database
    Public g_conEDIHistory() As ADODB.Connection    'Private datEDIHistory() As dao.Database
    Public g_conEDIFACT As ADODB.Connection         'Private datEDIFACT As dao.Database
    Public g_conData As ADODB.Connection            'Private datData As dao.Database
    Public g_conTemplate As ADODB.Connection        'Private conFind As ADODB.Connection

Public cAppPath As String

Public cLanguage As String

Public UserID As String

Public ResourceHandler As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, _
     ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Public CallingForm As Form
Public clsFindForm As cpiFind
Public AppTitle As String
Public strFields As String
Public blnEDIHistoryExisting As Boolean

Public g_clsAbout As Object
Public Const G_Main_Password = "wack2"

Public Sub LoadResStrings(ByRef frmFormToLoad As Form, Optional ByVal blnUseTag As Boolean)
    Dim ctlControlToLoad As Control
    Dim Tool As SSTool
    Dim Tool2 As SSTool
    Dim Tool3 As SSTool
    Dim Tool4 As SSTool
    
    Dim strTypeName As String
    Dim strToolTipText As String
    
    Dim intCtrlCount As Integer
    Dim intToolCount2 As Integer
    Dim intToolCount3 As Integer
    Dim intToolCount4 As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim nCtr As Integer
        
    On Error Resume Next
    
    If blnUseTag Then
        If frmFormToLoad.Tag <> "" Then
            frmFormToLoad.Caption = Translate(frmFormToLoad.Tag)
        End If
    Else
        frmFormToLoad.Caption = Translate(frmFormToLoad.Caption)
    End If
    
    For Each ctlControlToLoad In frmFormToLoad.Controls
        strTypeName = LCase(TypeName(ctlControlToLoad))
        If TypeOf ctlControlToLoad Is SSActiveToolBars Then
        End If
        
        Select Case strTypeName
            Case "ssactivetoolbars", "tlbrMain"
                For nCtr = 1 To ctlControlToLoad.ToolBars.Count  '2
                
                    intCtrlCount = ctlControlToLoad.ToolBars(nCtr).Tools.Count
                    
                    For i = 1 To intCtrlCount
                        Set Tool = ctlControlToLoad.ToolBars(nCtr).Tools(i)
                        
                        If (Tool.Name <> "") Then
                            Tool.Name = Translate(Tool.Name)
                        End If
                        
                        strToolTipText = ""
                        
                        If (Tool.ToolTipText <> "") Then
                            strToolTipText = Translate(Tool.ToolTipText)    ' Translate(ctlControlToLoad.Tools(i).ToolTipText)
                        End If
                        
                        strToolTipText = NoAmpersandEllipse(strToolTipText)
                        If UCase(Tool.ID) <> "SEPARATOR" Then
                            ctlControlToLoad.Tools(Tool.ID).ToolTipText = strToolTipText
                        End If
                        
                        'On Error Resume Next
                        If Tool.Type = ssTypeMenu Then
                            intToolCount2 = Tool.Menu.Tools.Count
                        Else
                            intToolCount2 = 0
                        End If
                        'If Err.Number <> 40006 Then
                        '    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                        'End If
                        'Err.Clear
                        'On Error GoTo 0
                        
                        
                        
                        For j = 1 To intToolCount2
                            Set Tool2 = Tool.Menu.Tools(j)
                            
                            If (Tool2.Name <> "") Then
                                Tool2.Name = Translate(Tool2.Name)
                            End If
                            
                            strToolTipText = ""
                            
                            If (Tool2.ToolTipText <> "") Then
                                strToolTipText = Translate(Tool2.ToolTipText)
                            End If
                            
                            strToolTipText = NoAmpersandEllipse(strToolTipText)
                            
                            If Len(strToolTipText) Then
                                ctlControlToLoad.Tools(Tool2.ID).ToolTipText = strToolTipText
                            End If
                            
                            If Tool2.Type = ssTypeMenu Then
                                intToolCount3 = Tool2.Menu.Tools.Count
                            Else
                                intToolCount3 = 0
                            End If
                            
                            For k = 1 To intToolCount3
                                Set Tool3 = Tool2.Menu.Tools(k)
                                
                                If (Tool3.Name <> "") Then
                                    Tool3.Name = Translate(Tool3.Name)
                                End If
                                
                                strToolTipText = ""
                                
                                If (Tool3.ToolTipText <> "") Then
                                    strToolTipText = Translate(Tool3.ToolTipText)
                                End If
                                
                                strToolTipText = NoAmpersandEllipse(strToolTipText)
                                
                                If Len(strToolTipText) Then
                                    ctlControlToLoad.Tools(Tool3.ID).ToolTipText = strToolTipText
                                End If
                                
                                If Tool3.Type = ssTypeMenu Then
                                    intToolCount4 = Tool3.Menu.Tools.Count
                                Else
                                    intToolCount4 = 0
                                End If

                            
                                For l = 1 To intToolCount4
                                    Set Tool4 = Tool3.Menu.Tools(l)
                                    Tool4.Name = Translate(Tool4.Name)
                                    
                                    strToolTipText = ""
                                    
                                    If (Tool4.ToolTipText <> "") Then
                                        strToolTipText = Translate(Tool4.ToolTipText)
                                    End If
                                    
                                    strToolTipText = NoAmpersandEllipse(strToolTipText)
                                    
                                    If Len(strToolTipText) Then
                                        ctlControlToLoad.Tools(Tool4.ID).ToolTipText = strToolTipText
                                    End If
                                        
                                Next
                            
                            Next
                        Next
                    Next
                Next
            Case "sstab"
                intCtrlCount = ctlControlToLoad.Tabs
                                
                For i = 0 To intCtrlCount - 1
                    ctlControlToLoad.TabCaption(i) = Translate(ctlControlToLoad.TabCaption(i))
                Next
            Case "tabstrip"
                intCtrlCount = ctlControlToLoad.Tabs.Count
                'edited by alg
                'For i = 0 To intCtrlCount  --> why i=0???...by alg
                For i = 1 To intCtrlCount
                    If (ctlControlToLoad.Tabs(i).Caption <> "") Then
                        ctlControlToLoad.Tabs(i).Caption = Translate(ctlControlToLoad.Tabs(i).Caption)
                    End If
                Next
            Case "label", "optionbutton", "frame", "commandbutton", "sscommand", "sspanel", "checkbox"
                If blnUseTag Then
                    If Trim$(ctlControlToLoad.Tag) <> "" Then
                        ctlControlToLoad.Caption = Translate(ctlControlToLoad.Tag)
                    End If
                Else
                    If (Trim$(ctlControlToLoad.Caption) <> "") Then
                        ctlControlToLoad.Caption = Translate(ctlControlToLoad.Caption)
                    End If
                End If
        End Select
    Next
End Sub

Public Function Translate(ByVal StringToTranslate As Variant) As String
    Dim cTranslated As String * 520
    
    If StringToTranslate Like "*[A-Z]*" Or StringToTranslate Like "*[a-z]*" Then
        Translate = StringToTranslate
    Else
        If (IsNumeric(StringToTranslate) = True) Then
            LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
        End If
        cTranslated = StripNullTerminator(cTranslated)
        Translate = RTrim(cTranslated)
    End If
    
End Function

'''Public Function StripNullTerminator(ByVal sCP As String) As String
'''    Dim posNull As Long
'''
'''    posNull = InStr(sCP, Chr$(0))
'''    StripNullTerminator = Left$(sCP, posNull - 1)
'''End Function

Public Function NoAmpersandEllipse(ByVal cText As String) As String
    Dim i As Integer
    
    i = InStr(1, cText, "&")
    
    If i > 0 Then
        cText = Mid(cText, 1, i - 1) + Mid(cText, i + 1)
    End If
    
    If Right(cText, 3) = "..." Then
        cText = Left(cText, Len(cText) - 3)
    End If
    
    NoAmpersandEllipse = cText
End Function

Public Function TheFileIsBeingSent(ByVal strUniqueCode As String, ByVal bytDocType As Byte) As Boolean
    Dim strSendFileName As String
    
    Dim bytDocumentType As Byte
    
    Select Case bytDocType
        Case 1, 2, 3
            bytDocumentType = bytDocType
        Case 4, 7
            bytDocumentType = 4
        Case 5, 9
            bytDocumentType = 5
        Case 6, 11
            bytDocumentType = 6
        Case 14
            bytDocumentType = 7
        Case 18
            bytDocumentType = 8
    End Select

    'REIMS - Modified to add checking for NCTS being sent
    'strSendFileName = cAppPath & "\" & Trim(strUniqueCode) & Choose(bytDocType, ".sdi", ".sde", ".sdt")
        
    strSendFileName = cAppPath & "\" & Trim(strUniqueCode) & Choose(bytDocumentType, ".sdi", ".sde", ".sdt", ".sdn", ".sdc", ".sdd", ".sdx", ".sdz")
    
    On Error Resume Next
    If (Dir(strSendFileName) <> "") Then
        Kill strSendFileName
    End If
    On Error GoTo 0
    
    If Len(Dir(strSendFileName)) Then
        TheFileIsBeingSent = True
    End If
End Function

Public Function TheFileIsOpen(ByVal strUniqueCode As String, ByVal bytDocType As Byte) As Boolean
    Dim intCtr As Long
    
    With CallingForm.File1
        .Path = cAppPath
' ********** Modified August 25, 2000 **********
' ********** Previously, this did not provide for distinction between Export/Transit.
        Select Case bytDocType
            Case 1, 4
                .Pattern = "*.csi"
            Case 2, 5
                .Pattern = "*.cse"
            Case 3, 6
                .Pattern = "*.cst"
            Case 7
                .Pattern = "*.csn"
            Case 9, 10
                .Pattern = "*.csc"
            Case 11
                .Pattern = "*.csd"
            Case 12
                .Pattern = "*.csa"
            Case 14
                .Pattern = "*.csx"
            Case 18
                .Pattern = "*.csz"
            Case Else
                Exit Function
        End Select
' ********** End Modify ************************
        .Refresh
        
        On Error Resume Next
        
        For intCtr = 0 To .ListCount - 1
            
            Kill cAppPath & "\" & .List(intCtr)
        Next
        
        On Error GoTo 0
        
        .Refresh
        
        For intCtr = 0 To .ListCount - 1
            If strUniqueCode = Mid(.List(intCtr), 1, Len(.List(intCtr)) - 4) Then    ' Disregard file extension
                TheFileIsOpen = True
                Exit For
            End If
        Next
    End With
End Function

Public Function CountChr(ByVal StringSearch As String, ByVal FindWhat As String) As Integer
    Dim nCtr As Integer
    Dim i As Integer
    Dim nFindwhat As Integer
    
    nCtr = 0
    nFindwhat = Len(Trim(FindWhat))
    
    For i = 1 To Len(Trim(StringSearch))
        If Mid(StringSearch, i, nFindwhat) = FindWhat Then
            nCtr = nCtr + 1
        End If
    Next
    
    CountChr = nCtr
End Function

Public Function GetDataType(ByRef ADOConnection As ADODB.Connection, _
                            ByVal TableName As String, _
                            ByVal FieldName As String) As ADOX.DataTypeEnum
                            
'Public Function GetDataType(datToUse As dao.Database, TableToOpen As String, FieldName As String) As Variant
    'Dim tbl As dao.TableDef
    
    Dim catWhere   As ADOX.Catalog
    Dim tblWhere   As ADOX.Table
    Dim fldWhere   As ADOX.Column
    
    catWhere.ActiveConnection = ADOConnection

    For Each tblWhere In catWhere.Tables  ' Loop through ADOConnection catalog tables
        If UCase$(Trim$(tblWhere.Type)) = "TABLE" And _
            UCase$(Trim$(tblWhere.Name)) = UCase$(Trim$(TableName)) Then
            For Each fldWhere In tblWhere.Columns  ' Loop through table fields
                            
                Debug.Assert False
                'GetDataType = tbl.Fields(Replace(Replace(FieldName, "[", ""), "]", "")).Type
                If UCase$(Trim$(fldWhere.Name)) = UCase$(Trim$(FieldName)) Then
                    GetDataType = fldWhere.Type
                    Exit For
                End If
            Next
        End If
    Next
    
    Set catWhere = Nothing
    Set tblWhere = Nothing
    Set fldWhere = Nothing
End Function

Public Sub ShowFields(ByVal strFields As String)
    Dim strFields2() As String
    
    strFields2 = Split(strFields, "*")
    frmShowFields.PreLoad strFields2
    
    Set frmShowFields = Nothing
End Sub

