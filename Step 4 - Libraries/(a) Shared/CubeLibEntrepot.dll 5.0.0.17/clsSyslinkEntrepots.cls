VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyslinkEntrepots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colSyslinkEntrepot As Collection

Public Property Get Count() As Long
    Count = colSyslinkEntrepot.Count
End Property

Public Property Get Item(vKey As Long) As clsSyslinkEntrepot
    Set Item = colSyslinkEntrepot.Item(vKey)
End Property

Public Sub Add(vHeaderDetail As String, Optional vProdNum As String, Optional vStockNum As String, Optional vJobNum As String, Optional vBatchNum As String, Optional vEntrepotNum As String, Optional vEntrepot As Boolean, Optional vInbound As Boolean, Optional vQuantity As String, Optional vNetWeight As String, Optional vGrossWeight As String, Optional vIsENTNR As Boolean)
    
    Dim lngCtr As Long
    Dim clsTemp As clsSyslinkEntrepot
        
    For lngCtr = 1 To colSyslinkEntrepot.Count
        If UCase(colSyslinkEntrepot.Item(lngCtr).Key) = UCase(vHeaderDetail) Then
            Exit For
        End If
    Next
    
    If lngCtr > colSyslinkEntrepot.Count Then
        
        Set clsTemp = New clsSyslinkEntrepot
        
        clsTemp.Key = vHeaderDetail
        clsTemp.EntrepotNum = vEntrepotNum
        clsTemp.ProdNum = vProdNum
        clsTemp.StockNum = vStockNum
        clsTemp.JobNum = vJobNum
        clsTemp.BatchNum = vBatchNum
        clsTemp.ValidEntrepot = vEntrepot
        clsTemp.ValidInbound = vInbound
        clsTemp.Quantity = Val(clsTemp.Quantity) + Val(vQuantity)
        clsTemp.NetWeight = vNetWeight
        clsTemp.GrossWeight = vGrossWeight
        clsTemp.IsENTNR = vIsENTNR
        
        colSyslinkEntrepot.Add clsTemp, vHeaderDetail
            
        Set clsTemp = Nothing
        
    Else
        
        If Trim(vEntrepotNum) <> "" Then
            colSyslinkEntrepot(vHeaderDetail).EntrepotNum = vEntrepotNum
        End If
        If Trim(vProdNum) <> "" Then
            colSyslinkEntrepot(vHeaderDetail).ProdNum = vProdNum
        End If
        If Trim(vStockNum) <> "" Then
            colSyslinkEntrepot(vHeaderDetail).StockNum = vStockNum
        End If
        If Trim(vJobNum) <> "" Then
            colSyslinkEntrepot(vHeaderDetail).JobNum = vJobNum
        End If
        If Trim(vBatchNum) <> "" Then
            colSyslinkEntrepot(vHeaderDetail).BatchNum = vBatchNum
        End If
        If vEntrepot Then
            colSyslinkEntrepot(vHeaderDetail).ValidEntrepot = vEntrepot
        End If
        If vInbound Then
            colSyslinkEntrepot(vHeaderDetail).ValidInbound = vInbound
        End If
        If vIsENTNR Then
            colSyslinkEntrepot(vHeaderDetail).IsENTNR = vIsENTNR
        End If
        If vQuantity <> "" Then
            colSyslinkEntrepot(vHeaderDetail).Quantity = vQuantity
        End If
        If vNetWeight <> "" Then
            colSyslinkEntrepot(vHeaderDetail).NetWeight = vNetWeight
        End If
        If vGrossWeight <> "" Then
            colSyslinkEntrepot(vHeaderDetail).GrossWeight = vGrossWeight
        End If
        
    End If
    
End Sub

Public Sub GenerateValidEntrepots(strHead As String, strDet As String, Optional ByVal ForPLDA As Boolean = False)
    
    Dim lngCtr As Long
    Dim lngCtr2 As Long
    Dim clsDet As clsSyslinkEntrepot
    Dim clsHed As clsSyslinkEntrepot
    
    For lngCtr = 1 To colSyslinkEntrepot.Count
        Set clsDet = colSyslinkEntrepot.Item(lngCtr)
        If InStr(clsDet.Key, strDet) Then
            For lngCtr2 = 1 To colSyslinkEntrepot.Count
                If colSyslinkEntrepot.Item(lngCtr2).Key = UCase(strHead & Mid(clsDet.Key, 2, 1)) Then
                    Set clsHed = colSyslinkEntrepot.Item(lngCtr2)
                    Exit For
                End If
            Next
            
            If lngCtr2 <= colSyslinkEntrepot.Count And Not ForPLDA Then
                If clsHed.EntrepotNum <> "" Then
                    clsDet.EntrepotNum = clsHed.EntrepotNum
                End If
                If clsDet.ValidEntrepot = False Then
                    clsDet.ValidEntrepot = clsHed.ValidEntrepot
                End If
                clsDet.ValidInbound = clsHed.ValidInbound
                
            ElseIf Not ForPLDA Then
                clsDet.EntrepotNum = ""
                clsDet.ValidEntrepot = False
                clsDet.ValidInbound = False
                
            ElseIf ForPLDA Then
                If Not clsDet.ValidInbound And clsDet.ValidEntrepot Then
                    clsDet.IsENTNR = True
                End If
            End If
        End If
    Next
    
    Set clsHed = Nothing
    Set clsDet = Nothing
    
    lngCtr = 1
    
    'Glenn - added checker to avoid error if colSyslinkEntrepot.Count = 0
    If colSyslinkEntrepot.Count >= lngCtr Then
        Do While True
                If InStr(colSyslinkEntrepot.Item(lngCtr).Key, strHead) Then
                    colSyslinkEntrepot.Remove lngCtr
                ElseIf InStr(colSyslinkEntrepot.Item(lngCtr).Key, strDet) Then
                    If colSyslinkEntrepot.Item(lngCtr).ValidEntrepot = False Or (colSyslinkEntrepot.Item(lngCtr).ValidInbound = False And colSyslinkEntrepot.Item(lngCtr).IsENTNR = False) Then
                        colSyslinkEntrepot.Remove lngCtr
                    Else
                        lngCtr = lngCtr + 1
                    End If
                Else
                    lngCtr = lngCtr + 1
                End If
                
                If lngCtr > colSyslinkEntrepot.Count Then
                    Exit Do
                End If
        Loop
    End If
    
End Sub

Public Sub GenerateValidEntrepotDetails(ByRef SADBELDB As ADODB.Connection)
    
    Dim rstEntrepotDetail As ADODB.Recordset
    Dim clsTemp As clsSyslinkEntrepot
    Dim lngCtr As Long
    Dim strSQL As String
    Dim strFilter As String
    Dim blnRemove As Boolean
    
    If colSyslinkEntrepot.Count = 0 Then
        Exit Sub
    End If
    
    lngCtr = 1
    
    Do While True
                
        Set clsTemp = colSyslinkEntrepot.Item(lngCtr)
        
        'Get SQL for outbound
        If GetSQLForOutbound(clsTemp, strSQL, strFilter) Then
            blnRemove = False
        Else
            blnRemove = True
        End If
        
        If blnRemove = False And clsTemp.ValidInbound = False Then
                                                    
            ADORecordsetOpen strSQL, SADBELDB, rstEntrepotDetail, adOpenKeyset, adLockOptimistic
            'rstEntrepotDetail.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
            
            rstEntrepotDetail.Filter = strFilter
            If rstEntrepotDetail.RecordCount = 0 Then
                rstEntrepotDetail.Filter = 0
            End If
            
            If rstEntrepotDetail.RecordCount = 0 And _
                rstEntrepotDetail.EOF Then
                
                colSyslinkEntrepot.Remove lngCtr
            Else
                
                rstEntrepotDetail.MoveFirst
                    
                clsTemp.In_ID = rstEntrepotDetail!In_ID
                clsTemp.Stock_ID = rstEntrepotDetail!Stock_ID
                clsTemp.TaricDesc = rstEntrepotDetail!Prod_Desc
                clsTemp.TaricCode = rstEntrepotDetail!Taric_Code
                clsTemp.BatchNum = rstEntrepotDetail!Batch_Num
                
                lngCtr = lngCtr + 1
                
            End If
                
            ADORecordsetClose rstEntrepotDetail
        
        ElseIf clsTemp.ValidInbound And Trim(clsTemp.ProdNum) <> "" Then
        
            If UCase(Trim(clsTemp.StockNum)) = "NEW" Then
                Call CreateNewStockCard(clsTemp, SADBELDB)
            Else
                Call CheckStockcardNumber(clsTemp, SADBELDB)
            End If
            
            If clsTemp.StockNum = "" Then
                colSyslinkEntrepot.Remove lngCtr
            Else
                lngCtr = lngCtr + 1
            End If
            
        Else
            colSyslinkEntrepot.Remove lngCtr
        End If
        
        If lngCtr > colSyslinkEntrepot.Count Then
            Exit Do
        End If
        
    Loop
        
    Set clsTemp = Nothing
    
    ADORecordsetClose rstEntrepotDetail
End Sub

Private Function GetSQLForOutbound(ByVal clsSyslink As clsSyslinkEntrepot, _
                                   ByRef strSQL As String, _
                                   ByRef strFilter As String) As Boolean
        
    If Trim(clsSyslink.BatchNum) <> "" Then
    
        If Trim(clsSyslink.JobNum) = "" Or _
            Trim(clsSyslink.StockNum) = "" Or _
            Trim(clsSyslink.ProdNum) = "" Then
            
            GetSQLForOutbound = False
                            
        Else
            GetSQLForOutbound = True
            If clsSyslink.ValidInbound = False Then
                
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "Stockcards.Stock_ID AS Stock_ID, "
                strSQL = strSQL & "Inbounds.In_ID AS In_ID, "
                strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
                strSQL = strSQL & "Products.Taric_Code AS Taric_Code, "
                strSQL = strSQL & "Inbounds.In_Batch_Num AS Batch_Num "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "Entrepots "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "Products "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Stockcards "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "Inbounds "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "InboundDocs "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Stockcards.Stock_ID = Inbounds.Stock_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Stockcards.Stock_Card_Num = '" & clsSyslink.StockNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "VAL(Inbounds.In_Avl_Qty_Wgt) > 0 "
                strSQL = strSQL & "ORDER BY "
                strSQL = strSQL & "InboundDocs.InDoc_Date ASC, "
                strSQL = strSQL & "Inbounds.In_ID ASC "
                
                strFilter = "Batch_Num = '" & clsSyslink.BatchNum & "'"
                
            End If
        End If
    ElseIf Trim(clsSyslink.JobNum) <> "" Then
        If Trim(clsSyslink.StockNum) = "" Or _
            Trim(clsSyslink.ProdNum) = "" Then
            
            GetSQLForOutbound = False
            
        Else
            GetSQLForOutbound = True
            If clsSyslink.ValidInbound = False Then
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "Stockcards.Stock_ID AS Stock_ID, "
                strSQL = strSQL & "Inbounds.In_ID AS In_ID, "
                strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
                strSQL = strSQL & "Products.Taric_Code AS Taric_Code, "
                strSQL = strSQL & "Stockcards.Stock_Card_Num AS Stock_Num, "
                strSQL = strSQL & "Inbounds.In_Batch_Num AS Batch_Num "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "Entrepots "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "Products "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Stockcards "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "Inbounds "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "InboundDocs "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Stockcards.Stock_ID = Inbounds.Stock_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
                strSQL = strSQL & ")  "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
                strSQL = strSQL & "WHERE Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "VAL(Inbounds.In_Avl_Qty_Wgt) > 0 "
                strSQL = strSQL & "ORDER BY "
                strSQL = strSQL & "Stockcards.Stock_Card_Num ASC, "
                strSQL = strSQL & "InboundDocs.InDoc_Date ASC, "
                strSQL = strSQL & "Inbounds.In_ID ASC "
            
                strFilter = "Stock_Num = '" & clsSyslink.StockNum & "'"
            End If
        End If
    ElseIf Trim(clsSyslink.StockNum) <> "" Then
        If Trim(clsSyslink.ProdNum) = "" Then
        
            GetSQLForOutbound = False
            
        Else
            GetSQLForOutbound = True
            If clsSyslink.ValidInbound = False Then
                
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "Stockcards.Stock_ID AS Stock_ID, "
                strSQL = strSQL & "Inbounds.In_ID AS In_ID, "
                strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
                strSQL = strSQL & "Products.Taric_Code AS Taric_Code, "
                strSQL = strSQL & "Stockcards.Stock_Card_Num AS Stock_Num, "
                strSQL = strSQL & "Inbounds.In_Batch_Num AS Batch_Num "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "Entrepots "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "Products "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Stockcards "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "( "
                            strSQL = strSQL & "Inbounds "
                            strSQL = strSQL & "INNER JOIN "
                            strSQL = strSQL & "InboundDocs "
                            strSQL = strSQL & "ON "
                            strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
                        strSQL = strSQL & ") "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Stockcards.Stock_ID = Inbounds.Stock_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
                strSQL = strSQL & "AND VAL(Inbounds.In_Avl_Qty_Wgt) > 0 "
                strSQL = strSQL & "ORDER BY "
                strSQL = strSQL & "Stockcards.Stock_Card_Num ASC, "
                strSQL = strSQL & "InboundDocs.InDoc_Date ASC, "
                strSQL = strSQL & "Inbounds.In_ID ASC "
                
                strFilter = "Stock_Num = '" & clsSyslink.StockNum & "'"
            End If
        End If
    ElseIf Trim(clsSyslink.ProdNum) <> "" Then
        GetSQLForOutbound = True
        If clsSyslink.ValidInbound = False Then
        
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "Stockcards.Stock_ID AS Stock_ID, "
            strSQL = strSQL & "Inbounds.In_ID AS In_ID, "
            strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
            strSQL = strSQL & "Products.Taric_Code AS Taric_Code, "
            strSQL = strSQL & "Stockcards.Stock_Card_Num AS Stock_Num, "
            strSQL = strSQL & "Inbounds.In_Batch_Num AS Batch_Num "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "Entrepots "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "Products "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "( "
                    strSQL = strSQL & "Stockcards "
                    strSQL = strSQL & "INNER JOIN "
                    strSQL = strSQL & "( "
                        strSQL = strSQL & "Inbounds "
                        strSQL = strSQL & "INNER JOIN "
                        strSQL = strSQL & "InboundDocs "
                        strSQL = strSQL & "ON "
                        strSQL = strSQL & "Inbounds.InDoc_ID = InboundDocs.InDoc_ID "
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "ON "
                    strSQL = strSQL & "Stockcards.Stock_ID = Inbounds.Stock_ID "
                strSQL = strSQL & ") "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
            strSQL = strSQL & "AND VAL(Inbounds.In_Avl_Qty_Wgt) > 0 "
            strSQL = strSQL & "ORDER BY "
            strSQL = strSQL & "Stockcards.Stock_Card_Num ASC, "
            strSQL = strSQL & "InboundDocs.InDoc_Date ASC, "
            strSQL = strSQL & "Inbounds.In_ID ASC "
            
            strFilter = ""
        End If
    ElseIf Trim(clsSyslink.ProdNum) = "" Then
        GetSQLForOutbound = False
    End If

End Function

Private Sub CheckStockcardNumber(ByRef clsSyslink As clsSyslinkEntrepot, _
                                 ByRef SADBELDB As ADODB.Connection)
    
    Dim rstStockcards As ADODB.Recordset
    Dim rstTemp As ADODB.Recordset
    Dim strSQL As String
    Dim lngProdID As Long
        
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Products.Prod_ID AS Prod_ID, "
        strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
        strSQL = strSQL & "Products.Taric_Code AS Taric_Code "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Entrepots "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "Products "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
    ADORecordsetOpen strSQL, SADBELDB, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
    
    'Check if Product Number exists
    If rstTemp.EOF And rstTemp.BOF Then
    
        clsSyslink.StockNum = ""
        
        ADORecordsetClose rstTemp
        
    Else
        rstTemp.MoveFirst
        
        'Save Product Id, description, and taric code
        lngProdID = rstTemp!Prod_ID
        clsSyslink.TaricCode = rstTemp!Taric_Code
        clsSyslink.TaricDesc = rstTemp!Prod_Desc
        
        ADORecordsetClose rstTemp
        
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "Products.Prod_ID AS Prod_ID, "
            strSQL = strSQL & "Products.Prod_Num AS Prod_Num, "
            strSQL = strSQL & "Stockcards.Stock_ID AS Stock_ID, "
            strSQL = strSQL & "Stockcards.Stock_Card_Num AS Stock_Card_Num "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "Entrepots "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "( "
                strSQL = strSQL & "Products "
                strSQL = strSQL & "INNER JOIN "
                strSQL = strSQL & "Stockcards "
                strSQL = strSQL & "ON "
                strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
            strSQL = strSQL & ") "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
            strSQL = strSQL & "ORDER BY "
            strSQL = strSQL & "Products.Prod_ID ASC, "
            strSQL = strSQL & "Stockcards.Stock_Card_Num ASC "
        ADORecordsetOpen strSQL, SADBELDB, rstTemp, adOpenKeyset, adLockOptimistic
        'rstTemp.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
        
        'Check if Entrepot contains Stockcards
        If rstTemp.EOF And rstTemp.BOF Then
                
                'allanSQL
                strSQL = vbNullString
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "* "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "Stockcards "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "Stockcards.Prod_ID = " & lngProdID & " "
            ADORecordsetOpen strSQL, SADBELDB, rstStockcards, adOpenKeyset, adLockOptimistic
            'rstStockcards.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
            
            rstStockcards.AddNew
            
            rstStockcards!Stock_Card_Num = clsSyslink.StockNum
            rstStockcards!Prod_ID = lngProdID
            rstStockcards!Stock_Archive = False
            
            'clsSyslink.Stock_ID = rstTemp!Stock_ID
            
            rstStockcards.Update
            
            clsSyslink.Stock_ID = InsertRecordset(SADBELDB, rstStockcards, "Stockcards")
            
            ADORecordsetClose rstStockcards
            
        Else
            
            rstTemp.Filter = "Prod_ID = " & lngProdID
            
            'Check if there are existing Stock Cards in the Product, if there is, then find the provided stockcard, else, create one
            If rstTemp.EOF And rstTemp.BOF Then
                
                'If stock card number was not provided then generate a new one
                If Trim(clsSyslink.StockNum) = "" Then
                    Call CreateNewStockCard(clsSyslink, SADBELDB)
                Else
                
                    rstTemp.Filter = 0
                    rstTemp.Filter = "Stock_Num = '" & clsSyslink.StockNum & "'"
                    
                    'Check if stock card number already exist in the entrepot, if not then create it
                    If rstTemp.EOF And rstTemp.BOF Then
                            'allanSQL
                            strSQL = vbNullString
                            strSQL = strSQL & "SELECT "
                            strSQL = strSQL & "* "
                            strSQL = strSQL & "FROM "
                            strSQL = strSQL & "Stockcards "
                            strSQL = strSQL & "WHERE "
                            strSQL = strSQL & "Stockcards.Prod_ID = " & lngProdID & " "
                        ADORecordsetOpen strSQL, SADBELDB, rstStockcards, adOpenKeyset, adLockOptimistic
                        'rstStockcards.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
                        
                        rstStockcards.AddNew
                        
                        rstStockcards!Stock_Card_Num = clsSyslink.StockNum
                        rstStockcards!Prod_ID = lngProdID
                        rstStockcards!Stock_Archive = False
                        
                        'clsSyslink.Stock_ID = rstTemp!Stock_ID
                        
                        rstStockcards.Update
                        
                        clsSyslink.Stock_ID = InsertRecordset(SADBELDB, rstStockcards, "Stockcards")
                        
                        ADORecordsetClose rstStockcards
                        
                    Else    'If the stock card number already exist in other product then generate a new one
                        
                        Call CreateNewStockCard(clsSyslink, SADBELDB)
                        
                    End If
                    
                End If
                
            Else
                
                rstTemp.MoveFirst
                rstTemp.Find "Stock_Card_Num = '" & clsSyslink.StockNum & "'"
                
                'If the stock card number was found then use it, else get the least stock card number
                If rstTemp.EOF Then
                    rstTemp.MoveFirst
                    
                    clsSyslink.StockNum = rstTemp!Stock_Card_Num
                    clsSyslink.Stock_ID = rstTemp!Stock_ID
                Else
                    clsSyslink.Stock_ID = rstTemp!Stock_ID
                End If
                
                ADORecordsetClose rstTemp
                
            End If
            
        End If
        
    End If
    
    ADORecordsetClose rstTemp
    ADORecordsetClose rstStockcards
End Sub

Private Sub Class_Initialize()
    Set colSyslinkEntrepot = New Collection
End Sub

Private Sub CreateNewStockCard(ByRef clsSyslink As clsSyslinkEntrepot, _
                               ByRef SADBELDB As ADODB.Connection)
    
    Dim strSQL As String
    Dim strNewStockNum As String
    Dim strStartingNum As String
    Dim rstTemp As ADODB.Recordset
    Dim rstClone As ADODB.Recordset
    Dim lngEntrepotID As Long
    Dim lngProdID As Long
    Dim blnAutoNumber As Boolean

        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT '"
        strSQL = strSQL & "Entrepots.Entrepot_ID AS Entrepot_ID, "
        strSQL = strSQL & "Entrepots.Entrepot_StockCard_Numbering AS Stock_Numbering, "
        strSQL = strSQL & "Entrepots.Entrepot_Starting_Num AS Starting_Num, "
        strSQL = strSQL & "Products.Prod_ID AS Prod_ID, "
        strSQL = strSQL & "Products.Prod_Desc AS Prod_Desc, "
        strSQL = strSQL & "Products.Taric_Code AS Taric_Code "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Entrepots "
        strSQL = strSQL & "INNER JOIN "
        strSQL = strSQL & "Products "
        strSQL = strSQL & "ON "
        strSQL = strSQL & "Entrepots.Entrepot_ID = Products.Entrepot_ID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Products.Prod_Num = '" & clsSyslink.ProdNum & "' "
        strSQL = strSQL & "AND "
        strSQL = strSQL & "Entrepots.Entrepot_Type & '-' & Entrepots.Entrepot_Num = '" & clsSyslink.EntrepotNum & "' "
    ADORecordsetOpen strSQL, SADBELDB, rstTemp, adOpenKeyset, adLockOptimistic
    'rstTemp.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
        
    'Check if Product Number does exist, if not, then invalid inbound
    If rstTemp.EOF And rstTemp.BOF Then
    
        clsSyslink.StockNum = ""
        
        ADORecordsetClose rstTemp
        
    Else    'Save Taric Code, Product Description, and search or create a Stock Card Number
        
        rstTemp.MoveFirst
        
        clsSyslink.TaricDesc = rstTemp!Prod_Desc
        clsSyslink.TaricCode = rstTemp!Taric_Code
        lngProdID = rstTemp!Prod_ID
        lngEntrepotID = rstTemp!Entrepot_ID
        blnAutoNumber = (rstTemp!Stock_Numbering = 0)
        strStartingNum = rstTemp!Starting_Num
        
        ADORecordsetClose rstTemp

            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "Stockcards.Stock_Card_Num AS Stock_Num "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "Products "
            strSQL = strSQL & "INNER JOIN "
            strSQL = strSQL & "Stockcards "
            strSQL = strSQL & "ON "
            strSQL = strSQL & "Products.Prod_ID = Stockcards.Prod_ID "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Products.Entrepot_ID = " & lngEntrepotID & " "
            strSQL = strSQL & "ORDER BY "
            strSQL = strSQL & "Val(Stockcards.Stock_Card_Num) "
        ADORecordsetOpen strSQL, SADBELDB, rstTemp, adOpenKeyset, adLockOptimistic
        'rstTemp.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
        Set rstClone = rstTemp.Clone
        
        If rstTemp.EOF And rstTemp.BOF Then
            
            'If no Stock Card in Entrepot then generate one
            If blnAutoNumber Then
                strNewStockNum = strStartingNum
            Else
                strNewStockNum = 1
            End If
        Else
            rstTemp.MoveFirst
            
            'Create least Stock Card Number
            Do While Not rstTemp.EOF
            
                strNewStockNum = Val(rstTemp!Stock_Num) + 1
                
                If Len(strNewStockNum) < Len(strStartingNum) And blnAutoNumber Then
                    strNewStockNum = String$(Len(strStartingNum) - Len(strNewStockNum), "0") & strNewStockNum
                End If
                
                rstClone.Filter = "Stock_Num = '" & strNewStockNum & "'"
                If rstClone.RecordCount > 0 Then
                    rstTemp.MoveNext
                    rstClone.Filter = 0
                Else
                    Exit Do
                End If
                
            Loop
                            
        End If
        
        ADORecordsetClose rstClone
        ADORecordsetClose rstTemp

        'Save to database and to class
            'allanSQL
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "* "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "Stockcards "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Stockcards.Prod_ID = " & lngProdID & " "
        ADORecordsetOpen strSQL, SADBELDB, rstTemp, adOpenKeyset, adLockOptimistic
        'rstTemp.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
        
        rstTemp.AddNew
        
        rstTemp!Stock_Card_Num = strNewStockNum
        rstTemp!Prod_ID = lngProdID
        rstTemp!Stock_Archive = False
        
        'clsSyslink.StockNum = strNewStockNum
        'clsSyslink.Stock_ID = rstTemp!Stock_ID
        
        rstTemp.Update
        
        clsSyslink.StockNum = strNewStockNum
        clsSyslink.Stock_ID = InsertRecordset(SADBELDB, rstTemp, "Stockcards")

    End If
    
    ADORecordsetClose rstTemp
    
End Sub

Private Sub Class_Terminate()
    Set colSyslinkEntrepot = Nothing
End Sub
