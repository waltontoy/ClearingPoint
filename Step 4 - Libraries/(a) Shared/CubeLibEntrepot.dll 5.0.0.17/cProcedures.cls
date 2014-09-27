VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Sub ReturnReserveToAvailable(ByVal strCode As String, _
                                    ByVal bytDocType As Byte, _
                                    ByRef SADBELDB As ADODB.Connection)
    
    Dim strSQL As String
    Dim strBaseTable As String
    Dim rst As ADODB.Recordset
    Dim rstInbound As ADODB.Recordset
    
    Dim lngProd_Handling As Long
    Dim dblQty_Wgt As Double
    Dim blnIsOutbound As Boolean
    
    'Glenn 4/5/2006
    Dim blnPrevDetT7IsF As Boolean
    Dim lngHeadWithT7F As Long

    Select Case bytDocType
        Case 1
            strBaseTable = "IMPORT"
        Case 2
            strBaseTable = "EXPORT"
        Case 3
            strBaseTable = "TRANSIT"
        Case 7
            strBaseTable = "NCTS"
        Case 9
            strBaseTable = "COMBINED NCTS"
        
        'Glenn - added codes for PLDA
        Case 14
            strBaseTable = "PLDA IMPORT"
        Case 18
            strBaseTable = "PLDA COMBINED"
        
        Case Else
            'do not process duplicatas!
            Exit Sub
    End Select
    
    
    'Glenn 4/5/2006 - ADDED: added HEADER and DETAIL for sorting
    If bytDocType = 1 Or _
        bytDocType = 2 Or _
        bytDocType = 3 Or _
        bytDocType = 14 Or _
        bytDocType = 18 Then
        
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* , "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].HEADER AS HEADER, "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].DETAIL AS DETAIL "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "[" & strBaseTable & "] "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "[" & strBaseTable & " HEADER] "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "[" & strBaseTable & " DETAIL] "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "([" & strBaseTable & " HEADER].HEADER = [" & strBaseTable & " DETAIL].HEADER) "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "([" & strBaseTable & " HEADER].CODE = [" & strBaseTable & " DETAIL].CODE)"
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "[" & strBaseTable & "].CODE = [" & strBaseTable & " HEADER].CODE "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[" & strBaseTable & "].CODE = '" & strCode & "' "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].HEADER ASC, "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].DETAIL ASC "
    ElseIf bytDocType = 7 Then
        
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "*, "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].HEADER AS HEADER, "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].DETAIL AS DETAIL "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
            strSQL = strSQL & "( "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "[" & strBaseTable & "] "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "[" & strBaseTable & " HEADER] "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "[" & strBaseTable & "].CODE = [" & strBaseTable & " HEADER].CODE "
                strSQL = strSQL & ") "
                strSQL = strSQL & "INNER JOIN '"
                strSQL = strSQL & "[" & strBaseTable & " DETAIL] "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "[" & strBaseTable & " HEADER].CODE = [" & strBaseTable & " DETAIL].CODE "
            strSQL = strSQL & ") "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "[" & strBaseTable & " DETAIL DOCUMENTEN] "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "([" & strBaseTable & " DETAIL].CODE = [" & strBaseTable & " DETAIL DOCUMENTEN].CODE) "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "([" & strBaseTable & " DETAIL].HEADER = [" & strBaseTable & " DETAIL DOCUMENTEN].HEADER) "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "([" & strBaseTable & " DETAIL].DETAIL = [" & strBaseTable & " DETAIL DOCUMENTEN].DETAIL) "
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL COLLI] "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "([" & strBaseTable & " DETAIL].CODE = [" & strBaseTable & " DETAIL COLLI].CODE) "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "([" & strBaseTable & " DETAIL].HEADER = [" & strBaseTable & " DETAIL COLLI].HEADER) "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "([" & strBaseTable & " DETAIL].DETAIL = [" & strBaseTable & " DETAIL COLLI].DETAIL) "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL COLLI].ORDINAL = 1 "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL DOCUMENTEN].ORDINAL = 1 "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "[" & strBaseTable & "].CODE = '" & strCode & "' "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].HEADER ASC, "
        strSQL = strSQL & "[" & strBaseTable & " DETAIL].DETAIL ASC "

    ElseIf bytDocType = 9 Then
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "*, "
        strSQL = strSQL & "[COMBINED NCTS DETAIL].HEADER AS HEADER, "
        strSQL = strSQL & "[COMBINED NCTS DETAIL].DETAIL AS DETAIL "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
            strSQL = strSQL & "( "
                strSQL = strSQL & "[COMBINED NCTS] "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "[COMBINED NCTS DETAIL] "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "[COMBINED NCTS].CODE = [COMBINED NCTS DETAIL].CODE "
            strSQL = strSQL & ") "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "[COMBINED NCTS DETAIL COLLI] "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "([COMBINED NCTS DETAIL].CODE = [COMBINED NCTS DETAIL COLLI].CODE) "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "([COMBINED NCTS DETAIL].HEADER = [COMBINED NCTS DETAIL COLLI].HEADER) "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "([COMBINED NCTS DETAIL].DETAIL = [COMBINED NCTS DETAIL COLLI].DETAIL) "
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "[COMBINED NCTS DETAIL GOEDEREN] "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "([COMBINED NCTS DETAIL COLLI].CODE = [COMBINED NCTS DETAIL GOEDEREN].CODE) "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "([COMBINED NCTS DETAIL COLLI].HEADER = [COMBINED NCTS DETAIL GOEDEREN].HEADER) "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "([COMBINED NCTS DETAIL COLLI].DETAIL = [COMBINED NCTS DETAIL GOEDEREN].DETAIL) "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "[COMBINED NCTS DETAIL COLLI].ORDINAL = 1 "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "[COMBINED NCTS DETAIL GOEDEREN].ORDINAL = 1 "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "[COMBINED NCTS].CODE = '" & strCode & "' "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & "[COMBINED NCTS DETAIL].HEADER ASC, "
        strSQL = strSQL & "[COMBINED NCTS DETAIL].DETAIL ASC "
    
    End If
    
    ADORecordsetOpen strSQL, SADBELDB, rst, adOpenKeyset, adLockOptimistic
    'rst.Open strSQL, SADBELDB, adOpenForwardOnly, adLockReadOnly
    
    ADORecordsetOpen "Select * from Inbounds", SADBELDB, rstInbound, adOpenKeyset, adLockOptimistic
    'rstInbound.Open "Select * from Inbounds", SADBELDB, adOpenKeyset, adLockPessimistic
    
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        
        If UCase(rst!Type) <> "T" Then
            Do While Not rst.EOF
                'Glenn 4/5/2006
                '---------------------------------------------------------
                Debug.Print "H:" & rst!Header; " D:" & rst!Detail
                If lngHeadWithT7F <> rst!Header Then
                    'New Header
                    blnPrevDetT7IsF = False
                End If
                
                If blnPrevDetT7IsF = True And rst!Header = lngHeadWithT7F Then
                    'Disregard detail since previous detail already end of declaration.
                    GoTo NextDetail
                End If
                    
                If rst!T7 = "F" Then
                    blnPrevDetT7IsF = True
                    lngHeadWithT7F = rst!Header
                End If
                '---------------------------------------------------------
                
                If rst!In_ID <> 0 Then
                
                    If Not (rstInbound.BOF And rstInbound.EOF) Then
                        rstInbound.MoveFirst
                        rstInbound.Find "In_ID = " & rst!In_ID
                       
                        If Not rstInbound.EOF Then
                            lngProd_Handling = GetProd_Handling(rstInbound!In_ID, SADBELDB)
                                                                       
                            If lngProd_Handling <> 9 Then
                                'Glenn - modifications for addition of PLDA
                                Select Case lngProd_Handling
                                    Case 0
                                        'If bytDocType = 1 Or bytDocType = 2 Or bytDocType = 3 Then
                                        If bytDocType = 1 Then
                                            dblQty_Wgt = CDbl(Val(rst!T6))
                                        ElseIf bytDocType = 14 Or bytDocType = 18 Then
                                            dblQty_Wgt = CDbl(Val(rst!S2))
                                        Else
                                            dblQty_Wgt = CDbl(Val(rst!S3))
                                        End If
                                    Case 1
                                        If bytDocType = 14 Or bytDocType = 18 Then
                                            dblQty_Wgt = CDbl(Val(rst!L9))
                                        Else
                                            dblQty_Wgt = CDbl(Val(rst!M1))
                                        End If
                                    Case 2
                                        If bytDocType = 14 Or bytDocType = 18 Then
                                            dblQty_Wgt = CDbl(Val(rst!LA))
                                        Else
                                            dblQty_Wgt = CDbl(Val(rst!M2))
                                        End If
                                    Case 9  'if not found
                                                
                                End Select
                                
                                Select Case bytDocType
                                    Case 1
                                        If IIf(IsNull(rst!T1), "", rst!T1) = "ENTNR" Then
                                            If IsValidEntrepot(IIf(IsNull(rst!BB), "", rst!BB), SADBELDB) Then
                                                blnIsOutbound = True
                                            End If
                                        End If
                                        
                                    Case 2, 3, 9
                                        If IIf(IsNull(rst!T1), "", rst!T1) = "ENTNR" Then
                                            If IsValidEntrepot(IIf(IsNull(rst!L6), "", rst!L6), SADBELDB) Then
                                                blnIsOutbound = True
                                            End If
                                        End If
                                    
                                    Case 7
                                        If IIf(IsNull(rst!Y2), "", rst!Y2) = "ENTNR" Then
                                            If IsValidEntrepot(GetEntrepotInBoxY4(IIf(IsNull(rst!Y4), "", rst!Y4)), SADBELDB) Then
                                                blnIsOutbound = True
                                            End If
                                        End If
                                    
                                    Case 14, 18
                                        'GLENN: needs checking if outbound.
                                        If IsValidEntrepot(IIf(IsNull(rst!M3), "", rst!M3) & "-" & IIf(IsNull(rst!M4), "", rst!M4), SADBELDB) Then
                                            blnIsOutbound = True
                                        End If
                                        
                                End Select
                                
                                If blnIsOutbound = True Then
                                    rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - dblQty_Wgt
                                    Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                                    
                                    rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + dblQty_Wgt
                                    
                                    rstInbound.Update
                                    
                                    UpdateRecordset SADBELDB, rstInbound, "Inbounds"
                                End If
                            End If
                        End If
                        
                    End If
                    
                    '====================== for mdb_history ===================================
                    If blnIsOutbound = True Then
                        UpdateReserveInHistory rst!In_ID, dblQty_Wgt
                    End If
                   '=============================================================================
                    
                End If
NextDetail:
                rst.MoveNext
            Loop
        End If
    End If

    ADORecordsetClose rst
    ADORecordsetClose rstInbound
End Sub


Private Sub UpdateReserveInHistory(ByVal In_ID As Long, ByVal dblQty_Wgt As Double)
    Dim conHistory As ADODB.Connection
    Dim rstInboundHistory As ADODB.Recordset
    Dim intYear As Integer
    Dim blnFoundInHistory As Boolean
    
    intYear = Year(Now)
    Do Until blnFoundInHistory
        
        ' TO DO CPNET UPDATE FOR SQLSERVER
        If Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(CStr(intYear), 2) & ".mdb") = "" Then
            If Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(CStr(intYear - 1), 2) & ".mdb") = "" Then
                Exit Do
            End If
        Else

            ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(CStr(intYear), 2)
                        
            ADORecordsetOpen "Select * from Inbounds where In_ID = " & In_ID, conHistory, rstInboundHistory, adOpenKeyset, adLockOptimistic
            'rstInboundHistory.Open "Select * from Inbounds where In_ID = " & In_ID, conHistory, adOpenKeyset, adLockPessimistic
        
            If Not (rstInboundHistory.BOF And rstInboundHistory.EOF) Then
                rstInboundHistory.MoveFirst
                
                blnFoundInHistory = True
                rstInboundHistory!In_Reserved_Qty_Wgt = rstInboundHistory!In_Reserved_Qty_Wgt - dblQty_Wgt
                Debug.Assert rstInboundHistory!In_Reserved_Qty_Wgt >= 0
    
                rstInboundHistory!In_Avl_Qty_Wgt = rstInboundHistory!In_Avl_Qty_Wgt + dblQty_Wgt
                rstInboundHistory.Update
                
                UpdateRecordset conHistory, rstInboundHistory, "Inbounds"
            End If
            
            ADORecordsetClose rstInboundHistory
            ADODisconnectDB conHistory

        End If
        
        intYear = intYear - 1
    Loop
    
    ADORecordsetClose rstInboundHistory
    ADODisconnectDB conHistory
End Sub
