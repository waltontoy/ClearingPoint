VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStockCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCancelled As Boolean
Private strStockCardSQL As String

Private rstFindSeveral As ADODB.Recordset
Public m_rstPass2GridOff As ADODB.Recordset

Public mvarConn_Sadbel As ADODB.Connection
Public mvarConn_Taric As ADODB.Connection
Public mvarLanguage As String
Public mvarTaricProp As Integer

Public Sub ShowStockCard(ByRef CallingForm As Object, _
                         ByRef SADBELDB As ADODB.Connection, _
                         ByRef TARICDB As ADODB.Connection, _
                         ByVal AutoUnload As Boolean, _
                         ByVal Language As String, _
                         ByVal TaricProperties As Integer, _
                         Optional Pass2CtrlName As String, _
                         Optional Search As String)

    Dim strFirst(0 To 6) As String
    Dim strSecond(0 To 4) As String
    Dim blnSeveral(2 To 4) As Boolean               '0 not included since the "Stock Card No" field can't be <Several>.
    Dim lngCtr As Long
    Dim strIM7 As String
    Dim strSCNumX As String
    Dim strSQL As String 'allanSQL
    
    ' rstFindSeveral    :: Primary recordset: contains all possible records.
    ' m_rstPass2GridOff      :: Secondary recordset: contains the word <Several>.
    
    'For use when loading Products picklist.
    Set mvarConn_Sadbel = SADBELDB
    Set mvarConn_Taric = TARICDB
    mvarTaricProp = TaricProperties
    mvarLanguage = Language
    
    'allanSQL
    strStockCardSQL = vbNullString
    strStockCardSQL = strStockCardSQL & "SELECT "
    strStockCardSQL = strStockCardSQL & "SC.Stock_ID AS [ID], "
    strStockCardSQL = strStockCardSQL & "SC.Stock_Card_Num AS [Stock Card No], "
    strStockCardSQL = strStockCardSQL & "I.In_Job_Num AS [Job No], "
    strStockCardSQL = strStockCardSQL & "I.In_Batch_Num AS [Batch No], "
    strStockCardSQL = strStockCardSQL & "ID.InDoc_Type AS [Doc Type], "
    strStockCardSQL = strStockCardSQL & "ID.InDoc_Num AS [Doc Num], "
    strStockCardSQL = strStockCardSQL & "P.Prod_ID AS [Product ID], "
    strStockCardSQL = strStockCardSQL & "P.Prod_Num AS [Prod Num], "
    strStockCardSQL = strStockCardSQL & "SC.Prod_ID AS [Prod ID], "
    strStockCardSQL = strStockCardSQL & "E.Entrepot_ID AS [Entrepot ID], "
    strStockCardSQL = strStockCardSQL & "E.Entrepot_Type AS [Entrepot Type], "
    strStockCardSQL = strStockCardSQL & "E.Entrepot_Num AS [Entrepot Num], "
    strStockCardSQL = strStockCardSQL & "E.Entrepot_StockCard_Numbering AS [Numbering], "
    strStockCardSQL = strStockCardSQL & "E.Entrepot_Starting_Num AS [Starting Num] "
    strStockCardSQL = strStockCardSQL & "FROM "
    strStockCardSQL = strStockCardSQL & "( "
        strStockCardSQL = strStockCardSQL & "Entrepots AS [E] "
        strStockCardSQL = strStockCardSQL & "INNER JOIN "
        strStockCardSQL = strStockCardSQL & "( "
            strStockCardSQL = strStockCardSQL & "StockCards AS [SC] "
            strStockCardSQL = strStockCardSQL & "INNER JOIN "
            strStockCardSQL = strStockCardSQL & "Products AS [P] "
            strStockCardSQL = strStockCardSQL & "ON "
            strStockCardSQL = strStockCardSQL & "SC.Prod_ID = P.Prod_ID "
        strStockCardSQL = strStockCardSQL & ") "
        strStockCardSQL = strStockCardSQL & "ON "
        strStockCardSQL = strStockCardSQL & "E.Entrepot_ID = P.Entrepot_ID "
    strStockCardSQL = strStockCardSQL & ") "
    strStockCardSQL = strStockCardSQL & "LEFT JOIN "
    strStockCardSQL = strStockCardSQL & "( "
        strStockCardSQL = strStockCardSQL & "Inbounds AS [I] "
        strStockCardSQL = strStockCardSQL & "LEFT JOIN "
        strStockCardSQL = strStockCardSQL & "InboundDocs AS [ID] "
        strStockCardSQL = strStockCardSQL & "ON "
        strStockCardSQL = strStockCardSQL & "I.Indoc_ID = ID.Indoc_ID "
    strStockCardSQL = strStockCardSQL & ") "
    strStockCardSQL = strStockCardSQL & "ON "
    strStockCardSQL = strStockCardSQL & "SC.Stock_ID = I.Stock_ID "
    strStockCardSQL = strStockCardSQL & "ORDER BY "
    strStockCardSQL = strStockCardSQL & "P.Prod_ID ASC, "
    strStockCardSQL = strStockCardSQL & "SC.Stock_ID ASC, "
    strStockCardSQL = strStockCardSQL & "I.In_Job_Num ASC, "
    strStockCardSQL = strStockCardSQL & "SC.Stock_Card_Num ASC "
    
    'Query for rstFindSeveral.
    ADORecordsetOpen strStockCardSQL, SADBELDB, rstFindSeveral, adOpenKeyset, adLockOptimistic
    'rstFindSeveral.Open strStockCardSQL, SADBELDB, adOpenKeyset, adLockOptimistic
    
    'Prepare second recordset for manual population.
    With m_rstPass2GridOff
        .CursorLocation = adUseClient
        .Fields.Append "Stock ID", adVarNumeric, 10
        .Fields.Append "Stock Card No", adVarChar, 10
        .Fields.Append "Job No", adVarChar, 50
        .Fields.Append "Batch No", adVarChar, 50
        .Fields.Append "IM7", adVarChar, 100
        .Fields.Append "Entrepot ID", adVarChar, 10
        .Fields.Append "Product ID", adVarChar, 10
        .Fields.Append "New", adBoolean
        .Fields.Append "Length", adVarChar, 10
        .Open
    End With
    '---------------- Locating those severals ----------------
    With rstFindSeveral
        If Not (.EOF And .BOF) Then .MoveFirst
        
        'Process primary recordset until last record or Stock Card No.
        Do Until .EOF = True
            '==========================================================================
            '===== m_rstPass2GridOff Fields    =====       ===== rstFindSeveral Fields =====
            '===== (strFirst & strSecond) =====       =====                       =====
            '===== 0 = Stock ID           =====       =====  0 = ID               =====
            '===== 1 = Stock Card No      =====       =====  1 = Stock Card No    =====
            '===== 2 = Job No             =====       =====  2 = Job No           =====
            '===== 3 = Batch No           =====       =====  3 = Batch No         =====
            '===== 4 = IM7                =====       =====  4 = Doc Type         =====
            '===== 5 = Entrepot ID        =====       =====  5 = Doc Num          =====
            '===== 6 = Product ID         =====       =====  6 = Product ID       =====
            '===== 7 = New                =====       =====  7 = Prod Num         =====
            '===== 8 = Length             =====       =====  8 = Entrepot ID      =====
            '=====                        =====       =====  9 = Entrepot Type    =====
            '=====                        =====       ===== 10 = Entrepot Num     =====
            '=====                        =====       ===== 11 = Numbering        =====
            '=====                        =====       ===== 12 = Starting Num     =====
            '==========================================================================
            
            'Stores first unique instance of:
            'Stock ID | Stock Card No | Job No | Batch No
            For lngCtr = 0 To 6
                Select Case lngCtr
                    Case Is < 4
                        If Not IsNull(.Fields(lngCtr).Value) Then
                            strFirst(lngCtr) = .Fields(lngCtr).Value
                        Else
                            'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                            strFirst(lngCtr) = Empty
                        End If
                    Case 4
                        If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                            'Sets format of IM7 to "Type"-"Num".
                            strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                        Else
                            strIM7 = Empty
                        End If
                        strFirst(lngCtr) = strIM7
                    Case 5
                        strFirst(lngCtr) = .Fields("Entrepot ID").Value
                    Case 6
                        strFirst(lngCtr) = .Fields("Product ID").Value
                End Select
            Next lngCtr
            .MoveNext
            
            'If there's more than 1 record in the rst.
            If .RecordCount > 1 And Not .EOF Then
                'Stores second instance (possibly unique).
                'Stock ID | Stock Card No | Job No | Batch No
                For lngCtr = 0 To 4
                    If Not IsNull(.Fields(lngCtr).Value) Then
                        If Not lngCtr = 4 Then
                            strSecond(lngCtr) = .Fields(lngCtr).Value
                        Else
                            If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                                'Sets format of IM7 to "Type"-"Num".
                                strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                            End If
                            strSecond(lngCtr) = strIM7
                        End If
                    Else
                        'MsgBox "Null value encountered in " & lngCtr & "."      'CUSTOM
                        strSecond(lngCtr) = Empty
                    End If
                Next lngCtr
            'In case there's only 1 record in the rst or an EOF.
            Else
                'Stores dummy second instance (definitely unique).
                For lngCtr = 0 To 4
                    strSecond(lngCtr) = "dummy"
                Next lngCtr
            End If
            
            'Check if first and second are alike.
            'Perform record transfer when distinct.
            If strFirst(1) <> strSecond(1) Then
                'Copy first instance to a new recordset.
                'This is the same action perform after processing multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 6
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    
                    strSCNumX = Empty
                End With
            'When First = Second.
            ElseIf strFirst(1) = strSecond(1) Then
            
                'Initialize/reset "Several" flag.
                For lngCtr = 2 To 4
                    blnSeveral(lngCtr) = False
                Next lngCtr
                    
                'Proceed with locating severals for every alike Stock Card Num.
                Do While strFirst(1) = strSecond(1)
                    'Cycles through other fields to identify which records will be labelled "Several".
                    For lngCtr = 2 To 4
                        'Used boolean flags to avoid too much string comparisons.
                        If blnSeveral(lngCtr) = False Then
                            'Copies next record to strSecond for comparison with strFirst.
                            If Not IsNull(.Fields(lngCtr).Value) Then           'CUSTOM
                                If Not lngCtr = 4 Then
                                    strSecond(lngCtr) = .Fields(lngCtr).Value
                                Else
                                    If Not (IsNull(.Fields(4).Value) And IsNull(.Fields(5).Value)) Then
                                        'Sets format of IM7 to "Type"-"Num".
                                        strIM7 = .Fields(4).Value & "-" & .Fields(5).Value
                                    End If
                                    strSecond(lngCtr) = strIM7
                                End If
                            Else
                                'MsgBox "Null value encountered in " & lngCtr & "."
                                strSecond(lngCtr) = Empty
                            End If
                                                                
                            'When strFirst != strSecond, a flag is raised to avoid processing this field.
                            If strFirst(lngCtr) <> strSecond(lngCtr) And IsEmpty(strSecond(lngCtr)) = False Then
                                blnSeveral(lngCtr) = True
                                'Stores the value "Several" for recording in first recordset.
                                strFirst(lngCtr) = "<Several>"
                                strSecond(lngCtr) = Empty
                            End If
                        End If
                    Next lngCtr

                    'Goes to next record or exits if last record of last distinct set.
                    .MoveNext
                    
                    If .EOF = True Then Exit Do
                    
                    strSecond(1) = .Fields(1).Value
                    
                    'Just adds the Entrepot ID and Product ID to the recordset.
                    strFirst(5) = .Fields("Entrepot ID").Value
                    strFirst(6) = .Fields("Product ID").Value
                Loop
                
                'Copy first instance to a new recordset. After multiple identical Stock Card Nos.
                With m_rstPass2GridOff
                    .AddNew
                    For lngCtr = 0 To 6
                        .Fields(lngCtr).Value = strFirst(lngCtr)
                    Next lngCtr
                    'Added for use with Stock Card numbering since they can contain leading zeroes.
                    strSCNumX = Replace(.Fields("Stock Card No").Value, "9", "")
                    strSCNumX = Trim(strSCNumX)
                    If Len(strSCNumX) > 0 Or Len(strSCNumX) = Len(.Fields("Stock Card No").Value) Then
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value)
                    Else
                    'Appends a "9" if incrementing the Stock Card No will increase its length.
                    'This is for cases when the ceiling is reached. E.g. 9, 99, 999, etc.
                        .Fields("Length").Value = Len(.Fields("Stock Card No").Value) & "9"
                    End If
                    .Update
                    
                    strSCNumX = Empty
                End With
            End If
        Loop
    End With
    
    ADORecordsetClose rstFindSeveral
    
'    m_rstPass2GridOff.Sort = "[Stock ID], [Stock Card No]"
    m_rstPass2GridOff.Sort = "[Length], [Stock Card No]"
    '------------ Finished locating those severals ------------
    
    'Display Stock Card picklist using modified recordset (containing the identified "<Several>").
    frmStockcardPicklist.Pre_Load Me, mvarCancelled
    
    'Open a new rstFindSeveral for commit purposes.
        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Stock_Card_Num AS [Stock Card No], "
        strSQL = strSQL & "Prod_ID AS [Prod ID] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "StockCards "
    ADORecordsetOpen strSQL, SADBELDB, rstFindSeveral, adOpenKeyset, adLockOptimistic
    'rstFindSeveral.Open strSQL, SADBELDB, adOpenKeyset, adLockOptimistic
    
    'Write to database routine, when user clicks on OK.
    If mvarCancelled = False Then
        m_rstPass2GridOff.Filter = "New = -1"
        If Not (m_rstPass2GridOff.BOF Or m_rstPass2GridOff.EOF) Then
            Do Until m_rstPass2GridOff.EOF
            
                rstFindSeveral.AddNew
                rstFindSeveral.Fields("Stock Card No").Value = m_rstPass2GridOff.Fields("Stock Card No").Value
                rstFindSeveral.Fields("Prod ID").Value = m_rstPass2GridOff.Fields("Product ID").Value
                rstFindSeveral.Update
                
                InsertRecordset SADBELDB, rstFindSeveral, "StockCards"
                
                m_rstPass2GridOff.MoveNext
            Loop
            
            
        End If
    End If
    
    'Clean up.
    ADORecordsetClose rstFindSeveral
    ADORecordsetClose m_rstPass2GridOff
    
    Set mvarConn_Sadbel = Nothing
    Set mvarConn_Taric = Nothing
End Sub
