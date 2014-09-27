Attribute VB_Name = "modGlobals"
Option Explicit

Global g_objDataSourceProperties As CDataSourceProperties

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_ORIENTATION = &H1&
Const DM_DUPLEX = &H1000&

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
    
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Enum ReportFilterType
    FilterStockID = 0
    FilterProductID = 1
    FilterEntrepotID = 2
    FilterAuthorizedPartyID = 3
End Enum
Public ResourceHandler As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, _
     ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

'Public g_conSadbel As ADODB.Connection    'Public datSADBEL As DAO.Database 'for Database Creation 1/5/05 rac
'Public g_conData As ADODB.Connection      'Public datData As DAO.Database 'for Database Creation 1/5/05 rac

Public Const G_Main_Password = "wack2"

Public Function GetDeleteCommandFromSelect(ByVal Command As String, _
                                           ByVal ParentTable As String) As String
    Dim strCommand As String
    Dim lngPos As Long
    
    strCommand = UCase$(Trim$(Command))
    
    lngPos = InStr(1, strCommand, " FROM ")
    
    If lngPos > 0 Then
        strCommand = "DELETE [" & ParentTable & "].* " & Mid(strCommand, lngPos)
    Else
        strCommand = vbNullString
        Debug.Assert False
    End If
    
    GetDeleteCommandFromSelect = strCommand
End Function

Public Function ConvertToDbl(ByVal Expression As Variant) As Double
' if passed expression is empty or not numeric then return 0

    If Expression = "" Or Not IsNumeric(Expression) Then
        ConvertToDbl = CDbl("0")
    Else
        ConvertToDbl = CDbl(Expression)
    End If

End Function

Public Function GetEntrepotType(ByVal strEntrepot As String) As String
    Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    GetEntrepotType = cpiEntrepotFunc.GetEntrepotType(strEntrepot)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function ConvertDDMMYY(ByVal DDMMYY As String) As String
Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    ConvertDDMMYY = cpiEntrepotFunc.ConvertDDMMYY(DDMMYY)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function GetProd_HandlingUsingStockID(ByVal StockID As Long, conSADBEL As ADODB.Connection) As Long
Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    GetProd_HandlingUsingStockID = cpiEntrepotFunc.GetProd_HandlingUsingStockID(StockID, conSADBEL)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function GetEntrepotNum(ByVal strEntrepot As String) As String
Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    GetEntrepotNum = cpiEntrepotFunc.GetEntrepotNum(strEntrepot)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function GetProd_Handling(ByVal lngIn_ID As Long, ByVal connSadbel As ADODB.Connection) As Long
Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    GetProd_Handling = cpiEntrepotFunc.GetProd_Handling(lngIn_ID, connSadbel)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function GetEntrepot_ID(ByVal strEntrepotNum As String, ByVal SADBELDB As ADODB.Connection, _
                                Optional ByVal blnShowMessage As Boolean, Optional ByVal strMessage As String, _
                                Optional ByVal strTitle As String) As Long
    Dim strSQL As String
    Dim rst As ADODB.Recordset
    
    ADORecordsetOpen "Select * from Entrepots where (Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = '" & strEntrepotNum & "'", SADBELDB, rst, adOpenKeyset, adLockOptimistic
    'rst.Open "Select * from Entrepots where (Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = '" & strEntrepotNum & "'", SADBELDB, adOpenForwardOnly, adLockReadOnly
    
    If rst.BOF And rst.EOF Then
        If blnShowMessage Then
            MsgBox strMessage, vbInformation, strTitle
        End If
    Else
        rst.MoveFirst
        GetEntrepot_ID = rst!Entrepot_ID
        
    End If
    
    ADORecordsetClose rst
End Function

Public Function GetProd_ID(ByVal strProdNum As String, ByVal SADBELDB As ADODB.Connection, _
                                Optional ByVal blnShowMessage As Boolean, Optional ByVal strMessage As String, _
                                Optional ByVal strTitle As String) As Long
    Dim strSQL As String
    Dim rst As ADODB.Recordset

    ADORecordsetOpen "Select * from Products where Prod_Num = '" & strProdNum & "'", SADBELDB, rst, adOpenKeyset, adLockOptimistic
    'rst.Open "Select * from Products where Prod_Num = '" & strProdNum & "'", SADBELDB, adOpenForwardOnly, adLockReadOnly

    If rst.BOF And rst.EOF Then
        If blnShowMessage Then
            MsgBox strMessage, vbInformation, strTitle
        End If
    Else
        rst.MoveFirst
        GetProd_ID = rst!Prod_ID

    End If

    ADORecordsetClose rst

End Function

Public Function GetCountryDesc(ByVal strCtryCode As String, _
                               ByVal SADBELDB As ADODB.Connection, _
                               ByVal strLanguage As String) As String
    Dim strSQL As String
    Dim rst As ADODB.Recordset

        strSQL = "SELECT Code AS [Key Code], [Description " & IIf(UCase(strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Key Description] " & _
                 "FROM [PICKLIST MAINTENANCE " & IIf(UCase(strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] INNER JOIN [PICKLIST DEFINITION] " & _
                 "ON [PICKLIST MAINTENANCE " & IIf(UCase(strLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(strLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "].[INTERNAL CODE] = [PICKLIST DEFINITION].[INTERNAL CODE] " & _
                 "WHERE Document = 'Import' and [BOX CODE] = 'C2' and Code = '" & strCtryCode & "'"
         
    ADORecordsetOpen strSQL, SADBELDB, rst, adOpenKeyset, adLockOptimistic
    'rst.Open strSQL, SADBELDB, adOpenKeyset, adLockReadOnly
    
    If Not (rst.BOF And rst.EOF) Then
        rst.MoveFirst
        
        GetCountryDesc = rst.Fields("Key Description").Value
    Else
        GetCountryDesc = "ALL YOUR BASE ARE BELONG TO US"
    End If
    
    ADORecordsetClose rst

End Function


Public Function Translate(ByVal StringToTranslate As Variant) As String
    Dim cTranslated As String * 520
    
' ********** Commented July 2, 2003 **********
' ********** IsNumeric() test fails in regional settings where the currency symbol is a letter
' ********** (e.g. in French where F is used for francs), causing box captions F1, F2, F3 to go blank.
'    If IsNumeric(StringToTranslate) Then
'        LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
'        cTranslated = StripNullTerminator(cTranslated)
'        Translate = RTrim(cTranslated)
'    Else
'        Translate = StringToTranslate
'    End If
' ********** End Comment *********************
    
    If StringToTranslate Like "*[A-Z]*" Or StringToTranslate Like "*[a-z]*" Then
        Translate = StringToTranslate
    Else
        'If (Trim(StringToTranslate) <> "") Then
        If (IsNumeric(StringToTranslate) = True) Then
            LoadString ResourceHandler, CLng(StringToTranslate), cTranslated, 520
        End If
        cTranslated = StripNullTerminator(cTranslated)
        Translate = RTrim(cTranslated)
    End If
End Function

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

'For database creation 1/5/05 rac
' TO UNCOMMENT
Public Sub CreateHistoryMdb(ByRef DBSADBEL As ADODB.Connection, _
                            ByRef DBDATA As ADODB.Connection, _
                            ByVal HistoryYear As String) ' 2chars.year
    
    ' TO DO FOR CP.NET UNCOMMENT AND DO
'    Dim wksHistory As DAO.Workspace
'    Dim dbsHistory As DAO.Database
'    Dim tdfTableToTransfer As DAO.TableDef
'
'    Dim strHistoryDBName As String
'    Dim strTableNames() As String
'    Dim lngCtr As Long
'    Dim gstrMDBPath As String
'
'    Set wksHistory = DBEngine.CreateWorkspace("History", "Admin", "", dbUseJet)
'    gstrMDBPath = GetSetting("ClearingPoint", "Settings", "MdbPath", "C:\Program Files\Cubepoint\ClearingPoint")
'    strHistoryDBName = gstrMDBPath & "\mdb_history" & cYear
'
'    If Len(Trim(Dir(strHistoryDBName & ".mdb"))) = 0 Then
'        Set dbsHistory = wksHistory.CreateDatabase(strHistoryDBName, dbLangGeneral)
'
'        ReDim strTableNames(33)
'
'        strTableNames(0) = "Import"
'        strTableNames(1) = "Export"
'        strTableNames(2) = "Transit"
'
'        strTableNames(3) = "Import Header"
'        strTableNames(4) = "Export Header"
'        strTableNames(5) = "Transit Header"
'
'        strTableNames(6) = "Import Detail"
'        strTableNames(7) = "Export Detail"
'        strTableNames(8) = "Transit Detail"
'
'        strTableNames(9) = "NCTS"
'        strTableNames(10) = "NCTS Header"
'        strTableNames(11) = "NCTS Header Zekerheid"
'        strTableNames(12) = "NCTS Detail"
'        strTableNames(13) = "NCTS Detail Colli"
'        strTableNames(14) = "NCTS Detail Container"
'        strTableNames(15) = "NCTS Detail Documenten"
'        strTableNames(16) = "NCTS Detail Bijzondere"
'
'        strTableNames(17) = "Combined NCTS"
'        strTableNames(18) = "Combined NCTS Header"
'        strTableNames(19) = "Combined NCTS Header Zekerheid"
'        strTableNames(20) = "Combined NCTS Detail"
'        strTableNames(21) = "Combined NCTS Detail Bijzondere"
'        strTableNames(22) = "Combined NCTS Detail Colli"
'        strTableNames(23) = "Combined NCTS Detail Container"
'        strTableNames(24) = "Combined NCTS Detail Documenten"
'        strTableNames(25) = "Combined NCTS Detail Gevoelige"
'        strTableNames(26) = "Combined NCTS Detail Goederen"
'
'        strTableNames(27) = "Inbounds"
'        strTableNames(28) = "InboundDocs"
'        strTableNames(29) = "Outbounds"
'        strTableNames(30) = "OutboundDocs"
'
'        strTableNames(31) = "Remarks"
'        strTableNames(32) = "Master"
'        strTableNames(33) = "MasterNCTS"
'
'        For lngCtr = LBound(strTableNames()) To UBound(strTableNames())
'            If lngCtr <= 26 Then
'                Set tdfTableToTransfer = datSADBEL.TableDefs(strTableNames(lngCtr))
'                CreateFields dbsHistory, tdfTableToTransfer
'            ElseIf lngCtr >= 27 And lngCtr <= 30 Then
'                Set tdfTableToTransfer = datSADBEL.TableDefs(strTableNames(lngCtr))
'                CreateFields dbsHistory, tdfTableToTransfer, True
'            Else
'                Set tdfTableToTransfer = datData.TableDefs(strTableNames(lngCtr))
'                CreateFields dbsHistory, tdfTableToTransfer
'            End If
'        Next
'    End If
'
'    Erase strTableNames()
'
'    dbsHistory.Close
'    wksHistory.Close
'
'    Set tdfTableToTransfer = Nothing
'    Set dbsHistory = Nothing
'    Set wksHistory = Nothing
End Sub

' TO UNCOMMENT
'Private Sub CreateFields(ByRef HistoryDB As DAO.Database, ByVal SADBELTable As DAO.TableDef, Optional ByVal blnEntrepotHis As Boolean)
'    Dim tdfHistory As DAO.TableDef
'    Dim idxHistory As DAO.Index
'    Dim idxSADBEL As DAO.Index
'    Dim fldSADBEL As DAO.Field
'
'    Set tdfHistory = HistoryDB.CreateTableDef(SADBELTable.Name)
'
'    With tdfHistory
'        For Each fldSADBEL In SADBELTable.Fields
'            .Fields.Append .CreateField(fldSADBEL.Name, fldSADBEL.Type, fldSADBEL.Size)
'            .Fields(fldSADBEL.Name).AllowZeroLength = True
'
'            'Sets Entrepot History ID fields type as autonumber Primary Key
'            If blnEntrepotHis = True Then
'                Select Case UCase(SADBELTable.Name)
'                    Case "INBOUNDS"
'                        If fldSADBEL.Name = "In_ID" Then
'                            .Fields("In_ID").Attributes = .Fields("In_ID").Attributes Or dbAutoIncrField
'                        End If
'                    Case "INBOUNDDOCS"
'                        If fldSADBEL.Name = "InDoc_ID" Then
'                            .Fields("InDoc_ID").Attributes = .Fields("InDoc_ID").Attributes Or dbAutoIncrField
'                        End If
'                    Case "OUTBOUNDS"
'                        If fldSADBEL.Name = "Out_ID" Then
'                            .Fields("Out_ID").Attributes = .Fields("Out_ID").Attributes Or dbAutoIncrField
'                        End If
'                    Case "OUTBOUNDDOCS"
'                        If fldSADBEL.Name = "OutDoc_ID" Then
'                            .Fields("OutDoc_ID").Attributes = .Fields("OutDoc_ID").Attributes Or dbAutoIncrField
'                        End If
'                End Select
'            End If
'
'        Next
'
'        For Each idxSADBEL In SADBELTable.Indexes
'            Set idxHistory = .CreateIndex(idxSADBEL.Name)
'
'            For Each fldSADBEL In idxSADBEL.Fields
'                idxHistory.Fields.Append .CreateField(fldSADBEL.Name)
'            Next
'
'            'Makes primary key Entrepot History table's ID field
'            If blnEntrepotHis = True Then
'                If UCase(idxSADBEL.Name) = "PRIMARYKEY" Then idxHistory.Primary = True
'            End If
'
'            .Indexes.Append idxHistory
'        Next
'    End With
'
'    HistoryDB.TableDefs.Append tdfHistory
'
'    'Sets Entrepot History ID increment type to random
'    If blnEntrepotHis = True Then
'        Select Case UCase(SADBELTable.Name)
'            Case "INBOUNDS"
'                HistoryDB.TableDefs("Inbounds").Fields("In_ID").DefaultValue = "GenUniqueID()"
'            Case "INBOUNDDOCS"
'                HistoryDB.TableDefs("InBoundDocs").Fields("InDoc_ID").DefaultValue = "GenUniqueID()"
'            Case "OUTBOUNDS"
'                HistoryDB.TableDefs("Outbounds").Fields("Out_ID").DefaultValue = "GenUniqueID()"
'            Case "OUTBOUNDDOCS"
'                HistoryDB.TableDefs("OutboundDocs").Fields("OutDoc_ID").DefaultValue = "GenUniqueID()"
'        End Select
'    End If
'
'    Set fldSADBEL = Nothing
'    Set idxSADBEL = Nothing
'    Set idxHistory = Nothing
'    Set tdfHistory = Nothing
'End Sub


Public Function IsValidEntrepot(ByVal strEntrepotNum As String, ByVal connSadbel As ADODB.Connection) As Boolean
    Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    IsValidEntrepot = cpiEntrepotFunc.IsValidEntrepot(strEntrepotNum, connSadbel)
    Set cpiEntrepotFunc = Nothing

End Function

Public Function GetEntrepotInBoxY4(ByVal strBoxY4 As String) As String
    Dim cpiEntrepotFunc As cEntrepotFunc

    Set cpiEntrepotFunc = New cEntrepotFunc
    
    GetEntrepotInBoxY4 = cpiEntrepotFunc.GetEntrepotInBoxY4(strBoxY4)
    Set cpiEntrepotFunc = Nothing

End Function


Public Function PadL(ByVal strStringToPad As Variant, ByVal intDesiredLength As Integer, ByVal strCharToPad As String, Optional ByVal blnIncludeLeadingTrailingSpaces As Boolean = False) As String
    Dim intActualLength As Integer
    
    strStringToPad = IIf(IsNull(strStringToPad), "", strStringToPad)
    
    intActualLength = Len(strStringToPad)
    
    If intActualLength < intDesiredLength Then
        If blnIncludeLeadingTrailingSpaces Then
            PadL = ReplicateString(strCharToPad, intDesiredLength - intActualLength) & strStringToPad
        Else
            PadL = ReplicateString(strCharToPad, intDesiredLength - intActualLength) & Trim(strStringToPad)
        End If
    Else
        If blnIncludeLeadingTrailingSpaces Then
            PadL = strStringToPad
        Else
            PadL = Trim(strStringToPad)
        End If
    End If
End Function

Public Function ShowPrinters(ByRef PrinterName As String) As Boolean

    Dim strTempPrinter As String
    Dim strDriverName As String
    Dim strPort As String
    Dim lngCtr As Long
    
    For lngCtr = 0 To Printers.Count - 1
        If UCase(Printers(lngCtr).DeviceName) = UCase(PrinterName) Then
            On Error Resume Next
            strTempPrinter = PrinterName
            strDriverName = Printers(lngCtr).DriverName
            strPort = Printer(lngCtr).Port
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Exit For
        End If
    Next
                
    ShowPrinters = ShowPrinter(strTempPrinter, strDriverName, strPort, GetActiveWindow)
    PrinterName = strTempPrinter
    
End Function

Public Function ShowPrinter(ByRef NewPrinterName As String, ByVal strDriverName As String, ByVal strPort As String, ByVal lngHandle As Long) As Boolean
    
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long
    Dim lpDevName As Long
    Dim bReturn As Integer
    Dim astrExtra() As String
    
    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = lngHandle

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(strDriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(strPort)
        .wDefault = 0
    End With

    DevName.extra = strDriverName & Chr(0) & NewPrinterName & Chr(0) & strPort & Chr(0)

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Printer Dialog
    If PrintDialog(PrintDlg) <> 0 Then
        'Get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)

        astrExtra() = Split(DevName.extra, Chr$(0))
        NewPrinterName = astrExtra(1)
        ShowPrinter = True
    Else
        ShowPrinter = False
    End If
    
    'Free allocated memory
    GlobalFree PrintDlg.hDevMode
    GlobalFree PrintDlg.hDevNames
    
End Function

