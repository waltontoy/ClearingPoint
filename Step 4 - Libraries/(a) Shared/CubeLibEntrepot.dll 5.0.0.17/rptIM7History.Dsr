VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptIM7History 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cpentrepotdll - rptIM7History (ActiveReport)"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "rptIM7History.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   _ExtentX        =   26935
   _ExtentY        =   13679
   SectionData     =   "rptIM7History.dsx":000C
End
Attribute VB_Name = "rptIM7History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngUserID As Long

Public mstrIndoc_Num                As String
Public mconSADBEL                   As ADODB.Connection
Public mstrEntrepot_Type_Num        As String
Public mdatDateGiven                As Date
Public mstrLicCompanyName           As String
Public mblnLicIsDemo                As Boolean

Private mrstIM7History              As ADODB.Recordset
Private mcolPackages                As Collection
Private mcolValues                  As Collection
Private mstrPrevious_Inbounddoc     As String
Private mstrPrevious_Inbound        As String
Private mstrPrevious_Outbounddoc    As String
Private mstrPrevious_Stockard       As String
Private mlngRunning_Total           As Double
Private mlngGrand_Total             As Double
Private mlngRow_Ctr                 As Long
Private mlngYear_Ctr                As Long
Private blnPrinted_Subtotal         As Boolean
Private mintBlankRow                As Boolean
Private mblnGrandtotal              As Boolean
Private mblnLastSubtotal            As Boolean

Private mblnPrinted                 As Boolean

'for the loopings
Private mlngPackagesCtr             As Long
Private mblnLoop_Subtotal1          As Boolean
Private mblnLoop_Subtotal2          As Boolean


Private Sub ActiveReport_PageEnd()
    If Me.pageNumber = 1 And mrstIM7History.EOF = False Then
        'mlngRow_Ctr = mlngRow_Ctr - 1
    ElseIf mrstIM7History.EOF And Me.pageNumber > 1 Then
        mlngRow_Ctr = mlngRow_Ctr + 1
    End If
    If mlngRow_Ctr > 66 Then
        mlngRow_Ctr = 66
    End If
    
    'Canvas.DrawLine Me.PageLeftMargin, (mlngRow_Ctr * Detail.Height) + Me.PageHeader.Height + Me.PageTopMargin, Me.PageLeftMargin + 10440, (mlngRow_Ctr * Detail.Height) + Me.PageHeader.Height + Me.PageTopMargin
    Canvas.DrawLine Me.PageSettings.LeftMargin, (mlngRow_Ctr * Detail.Height) + Me.PageHeader.Height + Me.PageSettings.TopMargin, Me.PageSettings.LeftMargin + 10440, (mlngRow_Ctr * Detail.Height) + Me.PageHeader.Height + Me.PageSettings.TopMargin
    mlngRow_Ctr = 0
End Sub

Private Sub ActiveReport_ReportEnd()
    
    On Error Resume Next
        
    ExecuteNonQuery mconSADBEL, "DROP TABLE IM7History_Rep" & "_" & Format(m_lngUserID, "00")
    'mconSADBEL.Execute "DROP TABLE IM7History_Rep" & "_" & Format(m_lngUserID, "00")
    
    On Error GoTo 0
    
    ADORecordsetClose mrstIM7History
End Sub

Private Sub ActiveReport_ReportStart()
    Dim strSQL                  As String
    Dim strDBPath               As String
    Dim lngYear_Ctr             As Long
    Dim rstIM7History           As ADODB.Recordset
    Dim rstValues_Buffer        As ADODB.Recordset
    Dim strHistoryPath          As String
    Dim blnTable_1_Created      As Boolean
    Dim blnTable_2_Created      As Boolean
    Dim rstBuffer               As ADODB.Recordset
    Dim astrHistory()           As String
    Dim strDBTemp               As String
    
    'paul for checking
    'Dim lngSeqNumber_Check      As Long
    Dim strSeqNumber            As String 'para po sa mga multiple seqnum for the year
    Dim colStock_ID             As Collection
    'Dim blnGlobal_Check         As Boolean
    'Dim strOffice               As String
    Dim colOffice               As Collection
    Dim colCertificate_Type     As Collection
    Dim colCertificate_Number   As Collection
    
    Dim strDocType              As String
    Dim blnIsHistoryDBExisting  As Boolean
    
    Dim lngTimer As Long
    
    Dim strHistoryDBYear As String
    
    
    ReDim astrHistory(0)
    
    blnIsHistoryDBExisting = False
    
    strDBPath = GetSetting("ClearingPoint", "Settings", "MDBPATH", "")
    strDBTemp = Dir(strDBPath & "\mdb_history??.mdb")
    Do While strDBTemp <> ""
        ReDim Preserve astrHistory(UBound(astrHistory) + 1)
        astrHistory(UBound(astrHistory)) = strDBTemp
        strDBTemp = Dir()
    Loop
    
    On Error Resume Next
    ExecuteNonQuery mconSADBEL, "DROP TABLE IM7History" & "_" & Format(m_lngUserID, "00") & " "
    ExecuteNonQuery mconSADBEL, "DROP TABLE IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " "
    'mconSADBEL.Execute "DROP TABLE IM7History" & "_" & Format(m_lngUserID, "00") & " "
    'mconSADBEL.Execute "DROP TABLE IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " "
    On Error GoTo 0
    
    For lngYear_Ctr = 1 To UBound(astrHistory)
        blnIsHistoryDBExisting = True
        strHistoryPath = strDBPath & "\" & astrHistory(lngYear_Ctr)
        
        If Len(Dir(strHistoryPath)) Then
            
            strHistoryDBYear = Replace(astrHistory(lngYear_Ctr), "mdb_history", vbNullString)
            strHistoryDBYear = Replace(strHistoryDBYear, ".mdb", vbNullString)
            
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_ID, "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_Type, "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_Num, "
                strSQL = strSQL & "DateValue(Format(HistoryInboundDocs.InDoc_Date, 'mm/dd/yyyy')) AS InDoc_Date, "
                strSQL = strSQL & "CStr(HistoryInbounds.In_ID) AS In_ID, "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_SeqNum, "
                strSQL = strSQL & "HistoryInboundDocs.Indoc_Cert_Num, "
                strSQL = strSQL & "HistoryInboundDocs.Indoc_Cert_Type, "
                strSQL = strSQL & "HistoryInbounds.In_Batch_Num, "
                strSQL = strSQL & "HistoryInbounds.In_Job_Num, "
                strSQL = strSQL & "HistoryInboundDocs.Indoc_Global, "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_Office, "
                strSQL = strSQL & "HistoryInboundDocs.InDoc_Date AS InDoc_DateMain, "
                strSQL = strSQL & "HistoryInbounds.In_Avl_Qty_Wgt, "
                strSQL = strSQL & "HistoryInbounds.Stock_ID, "
                strSQL = strSQL & "HistoryInbounds.In_Orig_Packages_Type, "
                strSQL = strSQL & "HistoryInbounds.In_Source_In_ID, "
                strSQL = strSQL & "HistoryInbounds.In_Orig_Packages_Qty, "
                strSQL = strSQL & "HistoryInbounds.In_Orig_Gross_Weight, "
                strSQL = strSQL & "HistoryInbounds.In_Orig_Net_Weight "
                strSQL = strSQL & "INTO "
                'strSQL = strSQL & "IM7History" & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & " "
                strSQL = strSQL & "FROM Entrepots INNER JOIN ("
                strSQL = strSQL & "Products INNER JOIN ("
                strSQL = strSQL & "StockCards INNER JOIN ("
                strSQL = strSQL & "HistoryInbounds INNER JOIN HistoryInboundDocs "
                strSQL = strSQL & "ON HistoryInbounds.InDoc_ID = HistoryInboundDocs.InDoc_ID) "
                strSQL = strSQL & "ON HistoryInbounds.Stock_ID = StockCards.Stock_ID) "
                strSQL = strSQL & "ON StockCards.Prod_ID = Products.Prod_ID) "
                strSQL = strSQL & "ON Products.Entrepot_ID = Entrepots.Entrepot_ID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "HistoryInboundDocs.Indoc_Num = '" & ProcessQuotes(mstrIndoc_Num) & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "IIf(IsNull(HistoryInbounds.IN_CODE), '', Right(HistoryInbounds.IN_CODE,8)) <> '<<TEST>>' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "IIf(IsNull(HistoryInbounds.IN_CODE), '', Right(HistoryInbounds.IN_CODE,11)) <> '<<CLOSURE>>' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & ProcessQuotes(mstrEntrepot_Type_Num) & "' "
                strSQL = strSQL & "ORDER BY "
                strSQL = strSQL & "HistoryInbounds.Stock_ID, "
                strSQL = strSQL & "HistoryInbounds.In_ID, "
                strSQL = strSQL & "HistoryInbounds.In_Orig_Packages_Type "
            
            '============ for inbounds and outbounds ===============
            
                strSQL = Replace(strSQL, "HistoryInbounds", "HistoryInboundsForIM7Report" & strHistoryDBYear)
                strSQL = Replace(strSQL, "HistoryInboundDocs", "HistoryInboundDocsForIM7Report" & strHistoryDBYear)
                
            ExecuteNonQuery mconSADBEL, strSQL
            'mconSADBEL.Execute strSQL
            
            
            If Not blnTable_1_Created Then
                
                strSQL = " SELECT IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Net_Weight, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Gross_Weight, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Packages_Qty, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_SeqNum, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_DateMain, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Source_In_ID, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Packages_Type, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Stock_ID, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Avl_Qty_Wgt, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Office, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Cert_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Cert_Type, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Job_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Batch_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Date, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Global, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Type, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_ID, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_ID, Outbounds.Out_Batch_Num, Outbounds.Out_Code, OutboundDocs.Outdoc_MRN, " & _
                         " Outbounds.Out_Job_Num, Outbounds.Out_Packages_Qty_Wgt, OutboundDocs.OutDoc_Date, OutboundDocs.OutDoc_Num, OutboundDocs.outdoc_Global, " & _
                         " OutboundDocs.OutDoc_Type, iif(isnull(Outbounds.Out_ID),NULL,cstr(Outbounds.Out_ID)) as Out_ID " & _
                         " INTO IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " " & _
                         " FROM IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & " LEFT JOIN ( Outbounds LEFT JOIN OutboundDocs ON Outbounds.OutDoc_ID = OutboundDocs.OutDoc_ID) " & _
                         " ON VAL(IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_ID) = Outbounds.In_ID " & _
                         " WHERE iif(isnull(Outbounds.Out_CODE), '',RIGHT(Outbounds.Out_CODE,8)) <> '<<TEST>>' AND iif(isnull(Outbounds.Out_CODE), '',RIGHT(Outbounds.Out_CODE,11)) <> '<<CLOSURE>>'"
                
            Else
                
                strSQL = " INSERT INTO IM7History_Rep" & "_" & Format(m_lngUserID, "00") & _
                         " SELECT IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Net_Weight, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Gross_Weight, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Packages_Qty, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_SeqNum, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Cert_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Cert_Type, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Source_In_ID, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Orig_Packages_Type, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Stock_ID, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Avl_Qty_Wgt,  IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Office, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_DateMain, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Job_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_Batch_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Date, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Num, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".Indoc_Global, " & _
                         " IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_Type, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".InDoc_ID, IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_ID, Outbounds.Out_Batch_Num, Outbounds.Out_Code, OutboundDocs.Outdoc_MRN, " & _
                         " Outbounds.Out_Job_Num, Outbounds.Out_Packages_Qty_Wgt, OutboundDocs.OutDoc_Date, OutboundDocs.OutDoc_Num, OutboundDocs.outdoc_Global, " & _
                         " OutboundDocs.OutDoc_Type, iif(isnull(Outbounds.Out_ID),NULL,cstr(Outbounds.Out_ID)) as OUt_ID" & _
                         " FROM IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & " LEFT JOIN ( Outbounds LEFT JOIN OutboundDocs ON Outbounds.OutDoc_ID = OutboundDocs.OutDoc_ID) " & _
                         " ON VAL(IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ".In_ID) = Outbounds.In_ID " & _
                         " WHERE iif(isnull(Outbounds.Out_CODE), '',RIGHT(Outbounds.Out_CODE,8)) <> '<<TEST>>' AND iif(isnull(Outbounds.Out_CODE), '',RIGHT(Outbounds.Out_CODE,11)) <> '<<CLOSURE>>'"
                
            End If
            
            ExecuteNonQuery mconSADBEL, strSQL
            'mconSADBEL.Execute strSQL
            
            
            blnTable_1_Created = True
            
            On Error Resume Next
            ExecuteNonQuery mconSADBEL, "DROP TABLE IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ""
            'mconSADBEL.Execute "DROP TABLE IM7History" & strHistoryDBYear & "_" & Format(m_lngUserID, "00") & ""
            On Error GoTo 0
            
        End If
    Next lngYear_Ctr
    
    
    If (blnIsHistoryDBExisting = True) Then
        mconSADBEL.Close
        mconSADBEL.Open

        'paul records cleaning, due to this program the processing time will be around O(n exp 2)
        
        Set colCertificate_Number = New Collection
        Set colCertificate_Type = New Collection
        Set colOffice = New Collection
        
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "IM7History_Rep_" & Format(m_lngUserID, "00") & " "
            strSQL = strSQL & "ORDER BY "
            strSQL = strSQL & "Stock_ID ASC, "
            strSQL = strSQL & "Indoc_Date ASC, "
            strSQL = strSQL & "In_Source_In_ID "
            strSQL = strSQL & "DESC "
        ADORecordsetOpen strSQL, mconSADBEL, rstBuffer, adOpenKeyset, adLockOptimistic
        'rstBuffer.Open strSQL, mconSADBEL, adOpenKeyset, adLockBatchOptimistic
            With rstBuffer
                If Not (.EOF And .BOF) Then
                    'check if meron record with the given Date
200005                  .Filter = 0
                        .Filter = "Indoc_Date = #" & Format(mdatDateGiven, "mm/dd/yyyy") & "#"
                        If rstBuffer.RecordCount <= 0 Then 'exit na pag wala
                            GoTo ExitNa
                        End If
                        
                        'for multiple indoc_seqnum
                        Set colStock_ID = New Collection
                        .MoveFirst
                        strSeqNumber = ""
                        
                        Do While Not .EOF
                            If InStr(strSeqNumber, "*~*" & .Fields("Indoc_SeqNum") & "*~*") = 0 Then
                                strSeqNumber = strSeqNumber & "*~*" & .Fields("Indoc_SeqNum") & "*~*"
                                If Not IsNull(.Fields("Indoc_Cert_Num")) Then
                                    colCertificate_Number.Add IIf(IsNull(.Fields("Indoc_Cert_Num")), "", .Fields("Indoc_Cert_Num")), CStr(.Fields("Indoc_SeqNum"))
                                End If
                                If Not IsNull(.Fields("Indoc_Cert_Type")) Then
                                    colCertificate_Type.Add IIf(IsNull(.Fields("Indoc_Cert_Type")), "", .Fields("Indoc_Cert_Type")), CStr(.Fields("Indoc_SeqNum"))
                                End If
                                If Not IsNull(.Fields("Indoc_Office")) Then
                                    colOffice.Add .Fields("Indoc_Office"), CStr(.Fields("Indoc_SeqNum"))
                                End If
                            End If
                            .MoveNext
                        Loop
                        
                        
                        'check if merong mga entries na yung dates di magkatugma
200020                  .Filter = 0
                        .Filter = "Indoc_Date <> #" & Format(mdatDateGiven, "mm/dd/yyyy") & "#"
                        If .RecordCount > 0 Then
                            .MoveFirst
                            Do While Not .EOF
                                'check values; available lang is Indoc_seqnum, indoc_global
                                'If lngSeqNumber_Check <> .Fields("Indoc_SeqNum") Or blnGlobal_Check <> .Fields("Indoc_Global") Then
                                If InStr(strSeqNumber, "*~*" & .Fields("Indoc_SeqNum") & "*~*") = 0 Then
                                    
                                    If IsNull(.Fields("In_Source_In_ID")) Then
                                        .Delete adAffectCurrent
                                        .UpdateBatch
                                    Else
                                        If Not Check_Insource_InID(.Fields("In_Source_In_ID")) Then
                                            .Delete adAffectCurrent
                                            .UpdateBatch
                                        End If
                                    End If
                                Else
                                    If IIf(IsNull(.Fields("Indoc_Office")), "", .Fields("Indoc_Office")) <> colOffice(CStr(.Fields("Indoc_SeqNum"))) Then
                                             .Delete adAffectCurrent
                                            .UpdateBatch
                                    ElseIf Not IsNull(.Fields("Indoc_Cert_Num")) Then
                                        If IIf(IsNull(.Fields("Indoc_Cert_Num")), "", .Fields("Indoc_Cert_num")) <> colCertificate_Number(CStr(.Fields("Indoc_SeqNum"))) Then
                                             .Delete adAffectCurrent
                                             .UpdateBatch
                                        End If
                                    ElseIf Not IsNull(.Fields("Indoc_Cert_Type")) Then
                                        If IIf(IsNull(.Fields("Indoc_Cert_Type")), "", .Fields("Indoc_Cert_Type")) <> colCertificate_Type(CStr(.Fields("Indoc_SeqNum"))) Then
                                             .Delete adAffectCurrent
                                             .UpdateBatch
                                        End If
                                    End If
                                End If
                                
                                .MoveNext
                            Loop
                        
                        End If
                End If
                .Filter = 0
                .Requery
            End With
        
        ADORecordsetClose rstBuffer
        
        Set colCertificate_Number = Nothing
        Set colCertificate_Type = Nothing
        Set colOffice = Nothing
        '=========
        
        ADORecordsetOpen "SELECT In_ID FROM IM7History_Rep" & "_" & Format(m_lngUserID, "00"), mconSADBEL, rstBuffer, adOpenKeyset, adLockOptimistic
        'rstBuffer.Open "SELECT In_ID FROM IM7History_Rep" & "_" & Format(m_lngUserID, "00"), mconSADBEL, adOpenKeyset, adLockBatchOptimistic
            
        If rstBuffer.EOF And rstBuffer.BOF Then
ExitNa:
            MsgBox "There are no Stocks with this Document Number.", vbOKOnly + vbInformation, "Stock Not Found"
            
            ADORecordsetClose rstBuffer
            
            Me.Cancel
            
        End If
        
        ADORecordsetClose rstBuffer
        
            strSQL = " SELECT DISTINCT * FROM IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " ORDER BY Stock_ID ASC, Indoc_DateMain ASC, In_ID ASC, Outdoc_Date ASC, OutDoc_Num ASC"
        ADORecordsetOpen strSQL, mconSADBEL, mrstIM7History, adOpenKeyset, adLockOptimistic
        'mrstIM7History.Open strSQL, mconSADBEL, adOpenKeyset, adLockOptimistic

        ' =======================
       
        mstrPrevious_Inbound = ""
        mstrPrevious_Stockard = ""
        mlngRunning_Total = 0
        mlngGrand_Total = 0
        mblnGrandtotal = False
        mblnLastSubtotal = False
        Set mcolPackages = New Collection
        Set mcolValues = New Collection
        
        ' Translate labels - joy 05/21/2006
        lblIM7.Caption = "IM7 " & Translate(2361)
        lblStock_Number.Caption = UCase$(Translate(2359))
        lblIN_BN_JN.Caption = Translate(2360)
        lblOUT_BN_JN.Caption = Translate(2360)
        lblbRunningTotal.Caption = UCase$(Translate(529))
        
        If mblnLicIsDemo Then
            lblVersionNum.Caption = "ClearingPoint v" & GetVersion() & " Demo version"
        Else
            lblVersionNum.Caption = "ClearingPoint v" & GetVersion() & IIf(Len(mstrLicCompanyName) > 0, " Licensed to: " & mstrLicCompanyName, "")
        End If
        
        lblPrintDate.Caption = Translate(2328) & " " & Now()
        
    Else
        
        MsgBox "No IM7 History Report to show.", vbInformation + vbOKOnly
        Me.Cancel
        
    End If
    
End Sub

Private Sub Detail_Format()
        
        Dim lngYear_Ctr         As Long
        Dim lngInbound_Value    As Double
        Dim lngOutbound_value   As Double
        Dim lngbuffer           As Double
        Dim blnNew_Inbound      As Boolean
        Dim blnNew_Stock        As Boolean
        Dim blnRepackaged       As Boolean

        Dim intSubtotal_Ctr     As Integer
        Dim strKey              As String
        Dim strDocument         As String

    If Not mblnGrandtotal Then
        If mrstIM7History.EOF = False Then
            
            'for total/ subtotal per stockcard
            
            fldRunningTotal.Font.Bold = False
            fldOut_batch_Job_Num.Font.Bold = False
            fldStock_Num.Font.Bold = True
            fldIn_Batch_Job_Num.Text = ""
            fldIn_DocType.Text = ""
            fldInbound.Text = ""
            fldOut_batch_Job_Num.Text = ""
            fldOutbounds.Text = ""
            fldStock_Num.Text = ""
            fldRunningTotal.Text = ""
           
           'paul if there is a looping for subtotal1 the go there
           If mblnLoop_Subtotal1 Then GoTo Subtotal_Loop1
            
            mblnPrinted = False
            
            If mstrPrevious_Stockard <> mrstIM7History![Stock_ID] And mstrPrevious_Stockard <> "" Then
            
Subtotal_Loop1:
                    If mblnLoop_Subtotal1 = False Then
                        mblnLoop_Subtotal1 = True
                        mlngPackagesCtr = 1
                    End If
                    
                     'Do While mintPackageCounter <> mcolPackages.Count
                     'the ifs will be the do while
                 
                    If mlngPackagesCtr <= mcolPackages.Count Then
                        fldOut_batch_Job_Num.Font.Bold = True
                        fldOutbounds.Font.Bold = True
    
                        If mblnPrinted = False Then
                            fldOutbounds.Text = "SUBTOTAL"
                            mblnPrinted = True
                        End If
    
                        If mstrPrevious_Stockard <> "0" Then
                            strKey = CStr(mcolPackages.Item(mlngPackagesCtr))
                           ' Debug.Assert CStr(mcolPackages.Item(mlngPackagesCtr)) = "CT"
                            fldOut_batch_Job_Num.Text = Replace(CStr(mcolValues.Item(strKey)), ",", ".") & " " & strKey
'                        Else
'                            mlngRow_Ctr = mlngRow_Ctr + 1
                        End If
                        'mlngRow_Ctr = mlngRow_Ctr + 1
                        mstrPrevious_Stockard = mcolPackages.Item(mlngPackagesCtr)
                        Me.LayoutAction = ddLAPrintSection + ddLAMoveLayout
                        Detail.PrintSection
                        mlngPackagesCtr = mlngPackagesCtr + 1
                        If mlngPackagesCtr > mcolPackages.Count Then
                            mblnLoop_Subtotal1 = False
                            
                        Else
                            mlngRow_Ctr = mlngRow_Ctr + 1
                            Exit Sub
                        End If
                    
                        
                    End If
                    
                    
                'end if the loop
                    
                    
                    
                    
                    Set mcolPackages = Nothing
                    Set mcolValues = Nothing
                    
                    Set mcolPackages = New Collection
                    Set mcolValues = New Collection
                    
                    lineRowBottom.Visible = True
                    blnPrinted_Subtotal = True
                    mstrPrevious_Stockard = mrstIM7History![Stock_ID]
                    '==DIFFERENCE
                    mlngGrand_Total = mlngGrand_Total + mlngRunning_Total
                    mlngRunning_Total = 0
                    '===
                    mlngRow_Ctr = mlngRow_Ctr + 1
                    Detail.PrintSection
                    Exit Sub
                
            End If
            
            fldOutbounds.Font.Bold = False
            fldOut_batch_Job_Num.Font.Bold = False
            
            fldIn_Batch_Job_Num.Text = ""
            fldIn_DocType.Text = ""
            fldInbound.Text = ""
            fldOut_batch_Job_Num.Text = ""
            fldOutbounds.Text = ""
            fldStock_Num.Text = ""
            fldRunningTotal.Text = ""
                    
            blnNew_Stock = False
            blnNew_Inbound = False
            
            LineRowTop.Visible = False
            lineRowBottom.Visible = False
            
            If mstrPrevious_Inbound <> CStr(mrstIM7History![In_ID] & mrstIM7History![InDoc_ID]) And mstrPrevious_Stockard = mrstIM7History![Stock_ID] Then
                mstrPrevious_Inbound = CStr(mrstIM7History![In_ID] & mrstIM7History![InDoc_ID])
                blnNew_Inbound = True
                'mlngRunning_Total =  mrstIM7History![]
            End If
            
            If mstrPrevious_Stockard <> mrstIM7History![Stock_ID] Or blnPrinted_Subtotal Then
           ' If blnPrinted_Subtotal Then
                mstrPrevious_Stockard = mrstIM7History![Stock_ID]
                blnNew_Stock = True
                
                fldStock_Num.Text = Get_Stock_Number(CStr(mrstIM7History![Stock_ID]))
                'mlngRow_Ctr = mlngRow_Ctr + 1
                Detail.PrintSection
                blnPrinted_Subtotal = False
            End If
           
            If blnNew_Inbound Then
                lngInbound_Value = mrstIM7History.Fields(GetHandling(mrstIM7History![Stock_ID]))
                
                'Packages ==
                strKey = mrstIM7History![In_Orig_Packages_Type]
                If In_Col(strKey, mcolPackages) Then
                    lngbuffer = CLng(mcolValues.Item(strKey)) + lngInbound_Value
                    mcolValues.Remove (strKey)
                    mcolValues.Add lngbuffer, strKey
                Else
                    mcolPackages.Add strKey, strKey
                    mcolValues.Add lngInbound_Value, strKey
                End If
                '===
                
                If lngInbound_Value < 0 Then
                    fldInbound.Text = Replace(CStr(lngInbound_Value), ",", ".") & "R " & mrstIM7History![In_Orig_Packages_Type]
                Else
                    fldInbound.Text = fldInbound.Text & Replace(CStr(lngInbound_Value), ",", ".") & IIf(IsNull(mrstIM7History![In_Source_In_Id]), " ", "R ") & mrstIM7History![In_Orig_Packages_Type]
                End If
                
                fldIn_Batch_Job_Num.Text = IIf(IsNull(mrstIM7History![In_Batch_Num]), "", mrstIM7History![In_Batch_Num]) & " " & IIf(IsNull(mrstIM7History![In_Job_Num]), "", mrstIM7History![In_Job_Num])
                mlngRunning_Total = mlngRunning_Total + lngInbound_Value
                fldRunningTotal.Text = Replace(CStr(mlngRunning_Total), ",", ".")
                
                If IIf(IsNull(mrstIM7History![Out_ID]), True, False) Then
                    Me.LayoutAction = ddLAPrintSection + ddLAMoveLayout + ddLANextRecord
                    mrstIM7History.MoveNext
                    If mrstIM7History.EOF Then
                        mblnGrandtotal = True
                    End If
                End If
                 mlngRow_Ctr = mlngRow_Ctr + 1
                 Detail.PrintSection
            End If
            
                
            If blnNew_Inbound = False And blnNew_Stock = False Then
                
                fldOut_batch_Job_Num.Text = IIf(IsNull(mrstIM7History![Out_Batch_Num]), "", mrstIM7History![Out_Batch_Num]) & " " & IIf(IsNull(mrstIM7History![Out_Job_Num]), "", mrstIM7History![Out_Job_Num])
                lngOutbound_value = IIf(IsNull(mrstIM7History![Out_Packages_Qty_Wgt]), 0, mrstIM7History![Out_Packages_Qty_Wgt])
                
                If IIf(IsNull(mrstIM7History![OutDoc_Num]), "", mrstIM7History![OutDoc_Num]) = "" Then
                    strDocument = IIf(IsNull(mrstIM7History![OutDoc_MRN]), "", mrstIM7History![OutDoc_MRN])
                Else
                    strDocument = mrstIM7History![OutDoc_Type] & " " & mrstIM7History![OutDoc_Num]
                End If
                
11121           If InStr(IIf(IsNull(mrstIM7History![Out_Code]), "", mrstIM7History![Out_Code]), "ICorrection") > 0 Then
                        'if icorrections then add to packages
                        lngbuffer = CLng(mcolValues.Item(mrstIM7History![In_Orig_Packages_Type])) + lngOutbound_value
                        mcolValues.Remove (mrstIM7History![In_Orig_Packages_Type])
                        mcolValues.Add lngbuffer, mrstIM7History![In_Orig_Packages_Type]
                        
                        fldInbound.Text = Replace(CStr(lngOutbound_value), ",", ".") & " " & mrstIM7History![In_Orig_Packages_Type]
11122                   mlngRunning_Total = mlngRunning_Total + lngOutbound_value
                        
                        If mstrPrevious_Outbounddoc <> strDocument Then
                            fldIn_DocType.Text = strDocument & " C" & Format(mrstIM7History![OutDoc_Date], "Short date")
                            mstrPrevious_Outbounddoc = strDocument
                        End If

3333            ElseIf InStr(IIf(IsNull(mrstIM7History![Out_Code]), "", mrstIM7History![Out_Code]), "OCorrection") > 0 Then
                        'if ocorrections then subtract from packages
                        lngbuffer = CLng(mcolValues.Item(mrstIM7History![In_Orig_Packages_Type])) - lngOutbound_value
                        mcolValues.Remove (mrstIM7History![In_Orig_Packages_Type])
                        mcolValues.Add lngbuffer, mrstIM7History![In_Orig_Packages_Type]
                        
                        fldOutbounds.Text = Replace(CStr(lngOutbound_value), ",", ".") & " " & mrstIM7History![In_Orig_Packages_Type]
4444                    mlngRunning_Total = mlngRunning_Total - lngOutbound_value

                        If mstrPrevious_Outbounddoc <> strDocument Then
                            fldIn_DocType.Text = strDocument & " C" & Format(mrstIM7History![OutDoc_Date], "Short date")
                            mstrPrevious_Outbounddoc = strDocument
                        End If
                    
                    
                Else
                    'Packages minus =====
                    lngbuffer = CLng(mcolValues.Item(mrstIM7History![In_Orig_Packages_Type])) - lngOutbound_value
                    mcolValues.Remove (mrstIM7History![In_Orig_Packages_Type])
                    mcolValues.Add lngbuffer, mrstIM7History![In_Orig_Packages_Type]
                    ' =====
                    fldOutbounds.Text = Replace(CStr(lngOutbound_value), ",", ".") & " " & mrstIM7History![In_Orig_Packages_Type]
                    
                   
                    If mstrPrevious_Outbounddoc <> strDocument Then
                        fldIn_DocType.Text = strDocument & IIf(IsNull(mrstIM7History![Out_Code]), " M", " ") & Format(mrstIM7History![OutDoc_Date], "Short date")
                        mstrPrevious_Outbounddoc = strDocument
                    End If
                    
                    mlngRunning_Total = mlngRunning_Total - lngOutbound_value
                End If
                
                
                fldRunningTotal.Text = Replace(CStr(mlngRunning_Total), ",", ".")
                mlngRow_Ctr = mlngRow_Ctr + 1
                Detail.PrintSection
                mrstIM7History.MoveNext
                If mrstIM7History.EOF Then
                   mblnGrandtotal = True
                End If
            End If
                 
         End If
    Else
            'for Grandtotal
        
        fldIn_Batch_Job_Num.Text = ""
        fldIn_DocType.Text = ""
        fldInbound.Text = ""
        fldOut_batch_Job_Num.Text = ""
        fldOutbounds.Text = ""
        fldStock_Num.Text = ""
        fldRunningTotal.Text = ""
        
        If mblnLoop_Subtotal2 Then GoTo Subtotal_Loop2
        
        mblnPrinted = False
        

        If Not mblnLastSubtotal Then

Subtotal_Loop2:
                    If mblnLoop_Subtotal2 = False Then
                        mblnLoop_Subtotal2 = True
                        mlngPackagesCtr = 1
                    End If
                     'Do While mintPackageCounter <> mcolPackages.Count
                     'the ifs will be the do while
                 
                    If mlngPackagesCtr <= mcolPackages.Count Then
                        fldOut_batch_Job_Num.Font.Bold = True
                        fldOutbounds.Font.Bold = True
    
                        If mblnPrinted = False Then
                            fldOutbounds.Text = "SUBTOTAL"
                            mblnPrinted = True
                        End If
    
                        If mstrPrevious_Stockard <> "0" Then
                            strKey = CStr(mcolPackages.Item(mlngPackagesCtr))
                            fldOut_batch_Job_Num.Text = Replace(CStr(mcolValues.Item(strKey)), ",", ".") & " " & strKey
                            mlngRow_Ctr = mlngRow_Ctr + 1
                        End If
                        
                        mstrPrevious_Stockard = mcolPackages.Item(mlngPackagesCtr)
                        Me.LayoutAction = ddLAPrintSection + ddLAMoveLayout
                        Detail.PrintSection
                        mlngPackagesCtr = mlngPackagesCtr + 1
                        If mlngPackagesCtr > mcolPackages.Count Then
                            mblnLoop_Subtotal2 = False
                        Else
                            Exit Sub
                        End If
                    
                        
                    End If


            'DIFFERENCE ===
            mlngGrand_Total = mlngGrand_Total + mlngRunning_Total
            mlngRunning_Total = 0
            '====
            mblnLastSubtotal = True
            mlngRow_Ctr = mlngRow_Ctr + 1
            Exit Sub
        'paul comments that we will continue after printing the subtotal
        Else
            LineRowTop.Visible = True
            
            fldOut_batch_Job_Num.Font.Bold = True
            fldOutbounds.Text = "GRAND TOTAL"
            fldRunningTotal.Font.Bold = True
            'DIFFERENCE===
            'fldRunningTotal.Text = CStr(mlngRunning_Total)
            fldRunningTotal.Text = Replace(CStr(mlngGrand_Total), ",", ".")
            '====
            lineRowBottom.Visible = True
            Me.LayoutAction = 7
             
            Detail.PrintSection
            mblnGrandtotal = False
        End If
    End If
        
       
        
End Sub



Private Sub PageHeader_Format()
        
        If Not mrstIM7History.EOF Then
            fldDoc_Num_Date.Text = "IM7 " & mstrIndoc_Num & " " & IIf(mrstIM7History![Indoc_Global] = -1, " M", "") & Format(mrstIM7History![Indoc_Date], "Short Date")
            PageHeader.PrintSection
        End If
        
End Sub

Private Function Get_Stock_Number(ByVal Stock_ID As String) As String
    
    Dim rstBuffer As ADODB.Recordset
    
    ADORecordsetOpen "SELECT Stock_Card_Num FROM Stockcards WHERE Stock_Id = " & Stock_ID, mconSADBEL, rstBuffer, adOpenKeyset, adLockOptimistic
    'rstBuffer.Open "SELECT Stock_Card_Num FROM Stockcards WHERE Stock_Id = " & Stock_ID, mconSADBEL, adOpenKeyset, adLockOptimistic
    
    Get_Stock_Number = "0"
    
    If Not (rstBuffer.EOF And rstBuffer.BOF) Then
        rstBuffer.MoveFirst
        
        Get_Stock_Number = CStr(rstBuffer![Stock_Card_Num])
    End If
    
End Function

Private Function GetHandling(Stock_ID As Long) As String

    Dim rstProd As ADODB.Recordset
    Dim strSQL As String
    
        strSQL = "SELECT Prod_Handling " & _
                "FROM Products INNER JOIN Stockcards ON " & _
                "Stockcards.Prod_ID = Products.Prod_ID " & _
                "WHERE StockCards!Stock_ID = " & Stock_ID
    ADORecordsetOpen strSQL, mconSADBEL, rstProd, adOpenKeyset, adLockOptimistic
    'rstProd.Open strSQL, mconSADBEL, adOpenKeyset, adLockOptimistic
    
    If Not (rstProd.EOF And rstProd.BOF) Then
        rstProd.MoveFirst
        
        Select Case rstProd!Prod_Handling
        
            Case 0
                GetHandling = "In_Orig_Packages_Qty"
            Case 1
                GetHandling = "In_Orig_Gross_Weight"
            Case 2
                GetHandling = "In_Orig_Net_Weight"
                
        End Select
    Else
        GetHandling = "In_Orig_Packages_Qty"
    End If
    
    ADORecordsetClose rstProd

End Function

Private Function In_Col(ByVal Key As String, ByRef COLL As Collection) As Boolean
    
    
    On Error GoTo Err_handler
        In_Col = Not IsNull(COLL.Item(Key))
    Exit Function

Err_handler:
    In_Col = False
    
End Function
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not mrstIM7History.EOF Or mblnGrandtotal Then
        EOF = False
    Else
        EOF = True
    End If
End Sub

Private Function GetVersion() As String

    Dim rstProd As ADODB.Recordset
    Dim strSQL As String
    Dim conbuffer As ADODB.Connection
    
        
        
    '<<< dandan 112306
    '<<< Uodate database with password
    ADOConnectDB conbuffer, g_objDataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
    'OpenADODatabase conbuffer, GetSetting("ClearingPoint", "Settings", "MDBPATH", ""), "TemplateCP.mdb"
    
        strSQL = "SELECT dbprops_Version FROM DBProperties "    'Table updated by BCo 2006-05-08
    ADORecordsetOpen strSQL, conbuffer, rstProd, adOpenKeyset, adLockOptimistic
    'rstProd.Open strSQL, conbuffer, adOpenKeyset, adLockOptimistic
    
    
    If Not (rstProd.EOF And rstProd.BOF) Then
        rstProd.MoveFirst
        
        GetVersion = rstProd![dbprops_Version]
    End If
    
    ADORecordsetClose rstProd
    ADODisconnectDB conbuffer
    
End Function

Private Function Check_Insource_InID(ByVal strInsource_InID) As Boolean
    
    Dim rstBuffer   As ADODB.Recordset
    
    Check_Insource_InID = False
    
    ADORecordsetOpen "SELECT In_ID FROM IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " WHERE In_ID = '" & strInsource_InID & "'", mconSADBEL, rstBuffer, adOpenKeyset, adLockOptimistic
    'rstBuffer.Open "SELECT In_ID FROM IM7History_Rep" & "_" & Format(m_lngUserID, "00") & " WHERE In_ID = '" & strInsource_InID & "'", mconSADBEL, adOpenKeyset, adLockBatchOptimistic
        
    If Not (rstBuffer.EOF And rstBuffer.BOF) Then
        Check_Insource_InID = True
    End If
    
    ADORecordsetClose rstBuffer
End Function

Public Property Get UserID() As Long
     UserID = m_lngUserID
End Property

Public Property Let UserID(ByVal Value As Long)
    m_lngUserID = Value
End Property
