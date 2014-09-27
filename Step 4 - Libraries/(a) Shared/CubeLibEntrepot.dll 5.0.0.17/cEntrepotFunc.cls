VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntrepotFunc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eCodiType
    eCodi_Import = 0
    eCodi_Export = 1
    eCodi_OTS = 2
    eCodi_TransitNCTS = 3
    eCodi_CombinedNCTS = 4
    eCodi_EDIDeparture = 5
    eCodi_PLDAImport = 6
    eCodi_PLDAExport = 7
    eCodi_PLDACombined = 8
End Enum

Public Function IsValidEntrepot(ByVal strEntrepotNum As String, _
                                ByRef SADBELDB As ADODB.Connection) As Boolean
    Dim intDashPos As Integer
    Dim strType As String
    Dim strNumber As String
    Dim rstEntrepot As ADODB.Recordset
    Dim strSQL As String 'allanSQL
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Entrepots "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "(Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num) = " & Chr(39) & ProcessQuotes(strEntrepotNum) & Chr(39) & " "
    ADORecordsetOpen strSQL, SADBELDB, rstEntrepot, adOpenKeyset, adLockOptimistic
    'rstEntrepot.Open strSQL, SADBELDB, adOpenForwardOnly, adLockReadOnly
    If rstEntrepot.BOF And rstEntrepot.EOF Then
        IsValidEntrepot = False
    Else
        IsValidEntrepot = True
    End If
    
    ADORecordsetClose rstEntrepot
End Function


Public Function IsInbound_CP(ByVal strBoxA7 As String) As Boolean
    If Mid(strBoxA7, 2, 1) = "7" Then
        IsInbound_CP = True
    End If
End Function

Public Function IsOutbound_CP(ByVal strBoxEntnr As String) As Boolean
    If UCase(strBoxEntnr) = "ENTNR" Then
        IsOutbound_CP = True
    End If
End Function


Public Function IsEntrepotActivated(ByRef SADBELDB As ADODB.Connection) As Boolean
    Dim rstEntrepot As ADODB.Recordset

    ADORecordsetOpen "Select * from Entrepots", SADBELDB, rstEntrepot, adOpenKeyset, adLockOptimistic
    'rstEntrepot.Open "Select * from Entrepots", SADBELDB, adOpenForwardOnly, adLockReadOnly
    
    If rstEntrepot.BOF And rstEntrepot.EOF Then
        IsEntrepotActivated = False
    Else
        IsEntrepotActivated = True
    End If
    
    ADORecordsetClose rstEntrepot
End Function

Public Function GetEntrepotType(ByVal strEntrepot As String) As String
    Dim intDashPos As Integer
    
    intDashPos = InStr(1, strEntrepot, "-")
    
    If intDashPos > 0 Then
        GetEntrepotType = Trim(Left(strEntrepot, intDashPos - 1))
    End If
End Function

Public Function GetEntrepotNum(ByVal strEntrepot As String) As String
    Dim intDashPos As Integer
    
    intDashPos = InStr(1, strEntrepot, "-")
    
    If intDashPos > 0 Then
        GetEntrepotNum = Trim(Mid(strEntrepot, intDashPos + 1))
    End If
End Function


Public Function WillShowEntrepotForOutbound(ByVal strBoxValueTrigger As String) As Boolean

    If UCase(strBoxValueTrigger) = "ENTNR" Then
        WillShowEntrepotForOutbound = True
    End If

End Function

Public Function GetProd_HandlingUsingStockID(ByVal StockID As Long, _
                                             ByRef SADBELDB As ADODB.Connection) As Long
    Dim rstProducts As ADODB.Recordset
    Dim strSQL As String 'allanSQL
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Prod_Handling "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Products "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "StockCards "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Products.Prod_ID = StockCards.Prod_ID "
        strSQL = strSQL & "WHERE Stock_ID = " & StockID & " "
    ADORecordsetOpen "Select Prod_Handling from Products INNER JOIN StockCards ON Products.Prod_ID = StockCards.Prod_ID where Stock_ID = " & StockID, SADBELDB, rstProducts, adOpenKeyset, adLockOptimistic
    'rstProducts.Open "Select Prod_Handling from Products INNER JOIN StockCards ON Products.Prod_ID = StockCards.Prod_ID where Stock_ID = " & StockID, SADBELDB, adOpenForwardOnly, adLockReadOnly
    
    If Not (rstProducts.BOF And rstProducts.EOF) Then
        GetProd_HandlingUsingStockID = rstProducts!Prod_Handling
    Else
        GetProd_HandlingUsingStockID = 9    'means not found
    End If
    
    ADORecordsetClose rstProducts
End Function

Public Sub checkAndChangeStockArchive(ByVal StockID As Long, _
                                      ByRef SADBELDB As ADODB.Connection)
    Dim rstEntrepotProductStockCard As ADODB.Recordset
    Dim strSQL As String 'allanSQL
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
        strSQL = strSQL & "Entrepots "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "Products "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID"
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "StockCards "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Products.Prod_ID = StockCards.Prod_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "StockCards.Stock_ID = " & StockID & " "
    ADORecordsetOpen strSQL, SADBELDB, rstEntrepotProductStockCard, adOpenKeyset, adLockOptimistic
    'rstEntrepotProductStockCard.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
    If Not (rstEntrepotProductStockCard.EOF And rstEntrepotProductStockCard.BOF) Then
        rstEntrepotProductStockCard.MoveFirst
        
        If rstEntrepotProductStockCard("Entrepot_StockCard_Archiving") = 0 Then
            rstEntrepotProductStockCard("Stock_Archive") = True
        End If
        
        rstEntrepotProductStockCard.Update
        
        ' TO DO CPNET UNCOMMENT
        'UpdateRecordset SADBELDB, rstEntrepotProductStockCard, ????
    End If
    
    ADORecordsetClose rstEntrepotProductStockCard
End Sub

Public Function NeedInboundUpdateIETEDI(ByVal blnIsEDIDeparture As Boolean, ByVal strOldTree As String, _
                                    ByVal strNewTree As String, ByVal blnNewCS_OR_blnCopy As Boolean) As eUpdateInbound

    If Not blnIsEDIDeparture Then
        If (strNewTree = "WL2" And strOldTree <> "WL2") Or (blnNewCS_OR_blnCopy And strNewTree = "WL2") Then
            NeedInboundUpdateIETEDI = eUpdate_FromOtherFolderToOutbox
        ElseIf strNewTree = "WL2" And strOldTree = "WL2" Then
            NeedInboundUpdateIETEDI = eUpdate_FromOutboxToOutbox
        ElseIf strNewTree = "DD" And strOldTree = "WL2" Then
            NeedInboundUpdateIETEDI = eUpdate_ToDeleted
        ElseIf strNewTree = "DE" And strOldTree = "WL2" Then
            NeedInboundUpdateIETEDI = eUpdate_ToRejected
        Else
            NeedInboundUpdateIETEDI = eUpdate_None
        End If
    Else

        If (strNewTree = "31ED" And strOldTree <> "31ED") Or (blnNewCS_OR_blnCopy And strNewTree = "31ED") Then
            NeedInboundUpdateIETEDI = eUpdate_FromOtherFolderToOutbox
        ElseIf strNewTree = "31ED" And strOldTree = "31ED" Then
            NeedInboundUpdateIETEDI = eUpdate_FromOutboxToOutbox
        ElseIf strNewTree = "-1ED" And (strOldTree = "31ED" Or strOldTree = "32ED") Then
            NeedInboundUpdateIETEDI = eUpdate_ToDeleted
        ElseIf strNewTree = "33ED" And (strOldTree = "31ED" Or strOldTree = "32ED") Then
            NeedInboundUpdateIETEDI = eUpdate_ToRejected
        Else
            NeedInboundUpdateIETEDI = eUpdate_None
        End If

    End If

End Function

Public Function UpdateInbound_CP(ByRef SADBELDB As ADODB.Connection, _
                                 ByVal lngIn_ID As Long, _
                                 ByVal UpdateType As eUpdateInbound, _
                                 ByVal strNewValue As String, _
                                 ByVal strOrigValUE As String, _
                                 ByVal strFieldToUse As String, _
                                 ByVal strMDBpath As String) As Byte

'if UpdateInbound_CP = 0
'if UpdateInbound_CP = 1--> mali updating ng eUpdate_Outbox
'if UpdateInbound_CP = 2--> tama updating ng eUpdate_Outbox

    Dim rstInbound As ADODB.Recordset
    Dim strSQL As String
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Inbounds.* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
            strSQL = strSQL & "Inbounds "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "InboundDocs "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "StockCards "
            strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                strSQL = strSQL & "Products "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "Entrepots "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Products.Entrepot_ID = Entrepots.Entrepot_ID "
                strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "StockCards.Prod_ID = Products.Prod_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Inbounds.Stock_ID = StockCards.Stock_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Inbounds.In_ID = " & CStr(lngIn_ID) & " "
    ADORecordsetOpen strSQL, SADBELDB, rstInbound, adOpenKeyset, adLockOptimistic
    'rstInbound.Open strSQL, SADBELDB, adOpenKeyset, adLockPessimistic

    If Not (rstInbound.BOF And rstInbound.EOF) Then
        
        Select Case UpdateType
            Case eUpdate_FromOtherFolderToOutbox
                rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt + Val(strNewValue)
                
                Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                
                rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt - Val(strNewValue)
            Case eUpdate_FromOutboxToOutbox
            
                '===== IF statement removed because there will be inconsistency on the quantity/weight defined in the codisheet! ======
                'If MsgBox("Would you like to add the " & CStr(dblOrigValUE) & " items you previously reserved?", vbQuestion + vbYesNo) = vbYes Then
                '    rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt + val(strNewValue)
                '    rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt - val(strNewValue)
                'Else
                    rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strOrigValUE) + Val(strNewValue)
                    Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                    
                    rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strOrigValUE) - Val(strNewValue)
                'End If
            
            Case eUpdate_ToDeleted
                rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strNewValue)
                Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                
                rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strNewValue)

            Case eUpdate_ToRejected
                rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strNewValue)
                
                Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                
                rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strNewValue)
                
            Case Else
                Debug.Assert False

        End Select

        rstInbound.Update
        
        ' TO DO CPNET UNCOMMENT
        'UpdateRecordset SADBELDB, rstInbound, ????
    End If
    
    ADORecordsetClose rstInbound
        
    '================== update correct history =====================================
    Dim intYear As Integer
    Dim conHistory As ADODB.Connection
    Dim blnFoundInHistory As Boolean
    
    intYear = Year(Now)
    
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Inbounds.* "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Inbounds "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Inbounds.In_ID = " & CStr(lngIn_ID) & " "
    Do Until blnFoundInHistory
        
        ' TO DO CPNET - EDIT FOR SQLSERVER
        If Dir(strMDBpath & "\mdb_history" & Right(CStr(intYear), 2) & ".mdb") = "" Then
            If Dir(strMDBpath & "\mdb_history" & Right(CStr(intYear - 1), 2) & ".mdb") = "" Then
                Exit Do
            End If
        Else
            '<<< dandan 112306
            '<<< Update with database password
            ADOConnectDB conHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(CStr(intYear), 2)
            
            ADORecordsetOpen strSQL, conHistory, rstInbound, adOpenKeyset, adLockOptimistic
            'rstInbound.Open strSQL, conHistory, adOpenKeyset, adLockOptimistic
    
            If Not (rstInbound.BOF And rstInbound.EOF) Then
            
                blnFoundInHistory = True
                rstInbound.MoveFirst
                
                Select Case UpdateType
                    Case eUpdate_FromOtherFolderToOutbox
                        rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt + Val(strNewValue)
                        'Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                        
                        rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt - Val(strNewValue)
                    Case eUpdate_FromOutboxToOutbox
                    
                        rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strOrigValUE) + Val(strNewValue)
                        'Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                        
                        rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strOrigValUE) - Val(strNewValue)
                    
                    Case eUpdate_ToDeleted
                        rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strNewValue)
                        
                        'Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                        
                        rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strNewValue)
        
                    Case eUpdate_ToRejected
                        rstInbound!In_Reserved_Qty_Wgt = rstInbound!In_Reserved_Qty_Wgt - Val(strNewValue)
                        
                        'Debug.Assert rstInbound!In_Reserved_Qty_Wgt >= 0
                        
                        rstInbound!In_Avl_Qty_Wgt = rstInbound!In_Avl_Qty_Wgt + Val(strNewValue)
                        
                    Case Else
                        Debug.Assert False
        
                End Select
        
                rstInbound.Update
                
                UpdateRecordset conHistory, rstInbound, "Inbounds"
            End If
            
            ADORecordsetClose rstInbound
            ADODisconnectDB conHistory

        End If
        
        intYear = intYear - 1
    Loop
    
    ADORecordsetClose rstInbound
    ADODisconnectDB conHistory
End Function

Public Function GetProd_Handling(ByVal lngIn_ID As Long, _
                                 ByVal SADBELDB As ADODB.Connection, _
                        Optional ByVal blnIsDIA As Boolean = False, _
                        Optional ByVal strDocDate As String) As Long

    Dim rstTmp As ADODB.Recordset
    Dim strSQL As String

        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Products.Prod_ID, "
        strSQL = strSQL & "Products.Prod_Handling "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "( "
            strSQL = strSQL & "Inbounds "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "InboundDocs "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "StockCards "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "Products "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "Entrepots "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Products.Entrepot_ID = Entrepots.Entrepot_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "StockCards.Prod_ID = Products.Prod_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Inbounds.Stock_ID = StockCards.Stock_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Inbounds.In_ID = " & CStr(lngIn_ID) & " "
    ADORecordsetOpen strSQL, SADBELDB, rstTmp, adOpenKeyset, adLockOptimistic
    'rstTmp.Open strSQL, SADBELDB

    If Not (rstTmp.BOF And rstTmp.EOF) Then
        rstTmp.MoveFirst
        
        GetProd_Handling = rstTmp!Prod_Handling
    Else
        If blnIsDIA Then    'check in history
                        
                        
            If Dir(NoBackSlash(g_objDataSourceProperties.InitialCatalogPath) & "\Mdb_History" & Right(strDocDate, 2) & ".mdb") <> "" Then
                Dim connHistory As ADODB.Connection
                Dim rstHistoryInbounds As ADODB.Recordset

                ADOConnectDB connHistory, g_objDataSourceProperties, DBInstanceType_DATABASE_HISTORY, Right(strDocDate, 2)
                'OpenADODatabase connHistory, strMDBpath, "Mdb_History" & Right(strDocDate, 2) & ".mdb"
                
                    strSQL = "Select Stock_ID from Inbounds where In_ID = " & CStr(lngIn_ID)
                ADORecordsetOpen strSQL, connHistory, rstHistoryInbounds, adOpenKeyset, adLockOptimistic
                'rstHistoryInbounds.Open strSQL, connHistory, adOpenForwardOnly, adLockReadOnly
                
                If Not (rstHistoryInbounds.BOF And rstHistoryInbounds.EOF) Then
                    rstHistoryInbounds.MoveFirst
                    
                    GetProd_Handling = GetProd_HandlingUsingStockID(rstHistoryInbounds!Stock_ID, SADBELDB)
                    
                Else
                    GetProd_Handling = 9
                End If
                
                
                ADORecordsetClose rstHistoryInbounds
                ADODisconnectDB connHistory

            End If
        Else
            GetProd_Handling = 9    'means not found
        End If
    End If

    ADORecordsetClose rstTmp
End Function

Public Function GetAndUpdateLastSeqNum(ByVal Stock_ID As Long, _
                                       ByVal SADBELDB As ADODB.Connection) As Long
    Dim strSQL As String
    Dim rstEntrepots As ADODB.Recordset
    
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Entrepot_LastSeqNum "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Entrepots "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "( "
            strSQL = strSQL & "Products "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "StockCards "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Products.Prod_ID = StockCards.Prod_ID "
        strSQL = strSQL & ") "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "StockCards.Stock_ID = " & Stock_ID & " "
    ADORecordsetOpen strSQL, SADBELDB, rstEntrepots, adOpenKeyset, adLockOptimistic
    'rstEntrepots.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
        
    With rstEntrepots
        If Not (.BOF And .EOF) Then
            .MoveFirst
            
            If IsNull(!Entrepot_LastSeqNum) Then
                !Entrepot_LastSeqNum = 1
                
                GetAndUpdateLastSeqNum = 1
            Else
                !Entrepot_LastSeqNum = !Entrepot_LastSeqNum + 1
                
                GetAndUpdateLastSeqNum = !Entrepot_LastSeqNum
            End If
            
            .Update
            
            ' TO DO CPNET UNCOMMENT
            'UpdateRecordset SADBELDB, rstEntrepots, ????
        Else
            GetAndUpdateLastSeqNum = 1
        End If
    End With
    
    ADORecordsetClose rstEntrepots
End Function

Public Function ConvertDDMMYY(ByVal DDMMYY As String) As String
    'return value will be in mm-dd-yyyy format
    
    ConvertDDMMYY = Format(Mid(DDMMYY, 3, 2) & "/" & Left(DDMMYY, 2) & "/" & Right(DDMMYY, 2), "mm/dd/yyyy")

End Function

'Public Function GetOutboundDocType(ByVal strDeclarationBox As String, ByVal blnIsImport As Boolean) As String
Public Function GetOutboundDocType(ByVal strDeclarationBox As String, ByVal bytDocType As Byte) As String
    'If blnIsImport Then
    Select Case bytDocType
        Case 1  'Import
            Select Case UCase(Left(strDeclarationBox, 2))
                Case "I0"
                    GetOutboundDocType = "IM0"
                Case "U0"
                    GetOutboundDocType = "EU0"
                Case "C4"
                    GetOutboundDocType = "CM4"
                Case "I4"
                    GetOutboundDocType = "IM4"
                Case "U4"
                    GetOutboundDocType = "EU4"
                Case "C5"
                    GetOutboundDocType = "CM5"
                Case "I5"
                    GetOutboundDocType = "IM5"
                Case "U5"
                    GetOutboundDocType = "EU5"
                Case "C6"
                    GetOutboundDocType = "CM6"
                Case "I6"
                    GetOutboundDocType = "IM6"
                Case "U6"
                    GetOutboundDocType = "EU6"
                Case "C7"
                    GetOutboundDocType = "CM7"
                Case "I7"
                    GetOutboundDocType = "IM7"
                Case "U7"
                    GetOutboundDocType = "EU7"
                Case "C9"
                    GetOutboundDocType = "CM9"
                Case "I9"
                    GetOutboundDocType = "IM9"
                Case "U9"
                    GetOutboundDocType = "EU9"
                Case "A4"
                    GetOutboundDocType = "AC4"
                Case Else
                    GetOutboundDocType = strDeclarationBox
                
            End Select
        Case 2, 3
    
            GetOutboundDocType = strDeclarationBox
        Case 9
            GetOutboundDocType = Left(Trim(strDeclarationBox), 2) & "0"
    End Select
    
End Function

Public Function GetEntrepotInBoxY4(ByVal strBoxY4 As String) As String
    'Rachelle Oct 20, 2005 : Use 'Trim' instead of 'Mid' due to cases where there can be two dashes used in the Entrepot type -
    'Entrepot number combination and where there is a possibility that they contain more than 6 characters.
'    Dim intDashPos As Integer
'    Dim strEntrepotType As String
'    Dim strEntrepotNum As String
    
'    intDashPos = InStr(1, strBoxY4, "-")
    
'    If intDashPos > 1 Then
'        GetEntrepotInBoxY4 = Mid(strBoxY4, intDashPos - 1, 6)
    GetEntrepotInBoxY4 = Trim(strBoxY4)
'    Else
'        GetEntrepotInBoxY4 = strBoxY4
'    End If

End Function
