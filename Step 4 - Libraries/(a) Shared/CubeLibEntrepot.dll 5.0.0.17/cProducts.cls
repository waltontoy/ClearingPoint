VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents pckProducts As PCubeLibPick.CPicklist
Attribute pckProducts.VB_VarHelpID = -1
Private mTaricProperties As Long
Private mLanguage As String
Private mConnection As ADODB.Connection
Private mConnectionTaric As ADODB.Connection
Private mvarEntrepot_Num As String
Private mvarCancelled As Boolean
Private mvarTaric_Code As String
Private mvarCtry_Origin As String
Private mvarCtry_Export As String
Private mvarProduct_ID As Long
Private mvarProduct_Num As String
                    
Private mvarOrigin_Desc As String
Private mvarExport_Desc As String
Private mvarProd_Desc As String
Private mvarEntrepot_ID As String
Private mvarEntrepot_Type As String
Private mvarEntrepot_Num_Only As String
Private mvarNumbering As Byte
Private mvarProd_Handling As Byte
Private mvarStartingNum As String

'local variable(s) to hold property value(s)
Private mvarResourceHandler As Long 'local copy
Private m_blnCheckSelectedProduct As Boolean

Public Property Let Prod_Handling(ByVal vData As Byte)
    mvarProd_Handling = vData
End Property

Public Property Get Prod_Handling() As Byte
    Prod_Handling = mvarProd_Handling
End Property

Public Property Let StartingNum(ByVal vData As String)
    mvarStartingNum = vData
End Property

Public Property Get StartingNum() As String
    StartingNum = mvarStartingNum
End Property


Public Property Let Numbering(ByVal vData As Byte)
    mvarNumbering = vData
End Property

Public Property Get Numbering() As Byte
    Numbering = mvarNumbering
End Property

Public Property Let Entrepot_Num_Only(ByVal vData As String)
    mvarEntrepot_Num_Only = vData
End Property

Public Property Get Entrepot_Num_Only() As String
    Entrepot_Num_Only = mvarEntrepot_Num_Only
End Property

Public Property Let Entrepot_Type(ByVal vData As String)
    mvarEntrepot_Type = vData
End Property

Public Property Get Entrepot_Type() As String
    Entrepot_Type = mvarEntrepot_Type
End Property

Public Property Let Entrepot_ID(ByVal vData As Long)
    mvarEntrepot_ID = vData
End Property

Public Property Get Entrepot_ID() As Long
    Entrepot_ID = mvarEntrepot_ID
End Property

Public Property Let Prod_Desc(ByVal vData As String)
    mvarProd_Desc = vData
End Property

Public Property Get Prod_Desc() As String
    Prod_Desc = mvarProd_Desc
End Property

Public Property Let Export_Desc(ByVal vData As String)
    mvarExport_Desc = vData
End Property

Public Property Get Export_Desc() As String
    Export_Desc = mvarExport_Desc
End Property

Public Property Let Origin_Desc(ByVal vData As String)
    mvarOrigin_Desc = vData
End Property

Public Property Get Origin_Desc() As String
    Origin_Desc = mvarOrigin_Desc
End Property

Public Property Let ResourceHandler(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResourceHandler = 5
    mvarResourceHandler = vData
End Property


Public Property Get ResourceHandler() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResourceHandler
    ResourceHandler = mvarResourceHandler
End Property




Public Property Let Product_Num(ByVal vData As String)
    mvarProduct_Num = vData
End Property

Public Property Get Product_Num() As String
    Product_Num = mvarProduct_Num
End Property

Public Property Let Product_ID(ByVal vData As Long)
    mvarProduct_ID = vData
End Property

Public Property Get Product_ID() As Long
    Product_ID = mvarProduct_ID
End Property

Public Property Let Ctry_Export(ByVal vData As String)
    mvarCtry_Export = vData
End Property

Public Property Get Ctry_Export() As String
    Ctry_Export = mvarCtry_Export
End Property

Public Property Let Ctry_Origin(ByVal vData As String)
    mvarCtry_Origin = vData
End Property

Public Property Get Ctry_Origin() As String
    Ctry_Origin = mvarCtry_Origin
End Property

Public Property Let Taric_Code(ByVal vData As String)
    mvarTaric_Code = vData
End Property

Public Property Get Taric_Code() As String
    Taric_Code = mvarTaric_Code
End Property

Public Property Let Entrepot_Num(ByVal vData As String)
    mvarEntrepot_Num = vData
End Property

Public Property Get Entrepot_Num() As String
    Entrepot_Num = mvarEntrepot_Num
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = mvarCancelled
End Property

Public Sub ShowProducts(Index As Long, CallingForm As Object, Connection As ADODB.Connection, ConnectionTaric As ADODB.Connection, Language As String, ByVal TaricProperties As Long, _
                        ByVal MyResourceHandler As Long, Optional ByVal strTxtBox As String, Optional UseSimple As Boolean, Optional ByVal blnForSummaryReport As Boolean = False)

    
    Dim gsdProducts As PCubeLibPick.CGridSeed
    Dim strProductsSQL As String
    Dim strEntrepotFields As String
    Dim strEntrepotWhere As String
    'For picklist's filter functionality.
    Dim clsPicklistFilter As CPicklistFilter
    Dim strSQL As String 'allanSQL
    
    ResourceHandler = MyResourceHandler
    
    Set clsPicklistFilter = New CPicklistFilter
    Set pckProducts = New CPicklist
    Set gsdProducts = New CGridSeed
    'Set gsdProducts = pckProducts.SeedGrid("Product Num", 1300, "Left", "Description", 2000, "Left", "Taric Code", 1100, "Left", "Entrepot Type", 1100, "Left", "Entrepot Num", 1050, "Left")
    Set gsdProducts = pckProducts.SeedGrid("Product Num", 1300, "Left", "Description", 2000, "Left", "Taric Code", 1100, "Left", "Entrepot Type", 1100, "Left", "Entrepot Num", 1050, "Left", "Origin Code", 1000, "Left", _
    "Origin Description", 1500, "Left", "Export Code", 1000, "Left", "Export Description", 1500, "Left")
    Set mConnection = Connection
    Set mConnectionTaric = ConnectionTaric
    
    mTaricProperties = TaricProperties
    mLanguage = Language
    
    strEntrepotFields = "Prod_ID,ID,Entrepot Type,Entrepot Num,Entrepot_ID,Auth_ID,Entrepot_StartDate,Entrepot_EndDate," & _
    "Entrepot_StockCard_Numbering,Entrepot_Starting_Num,Entrepot_StockCard_Format,Entrepot_StockCard_Archiving,Entrepot_LastSeqNum,Origin Description,Export Description,M1.Internal Code,M1.Code,M2.Internal Code,M2.Code"
    
    'For use with codisheets when accessing StockProd and AvailStocks
    'to filter picklist according to Entrepot Type-Num, Taric_Code,
    'Country of Origin and Country of Export.
    strEntrepotWhere = "WHERE "
    If Len(Entrepot_Num) <> 0 Then
        strEntrepotWhere = strEntrepotWhere & _
                           "E.Entrepot_Type = '" & Mid$(mvarEntrepot_Num, 1, 1) & "' AND " & _
                           "E.Entrepot_Num = '" & Mid$(mvarEntrepot_Num, 3) & "'"
    ElseIf blnForSummaryReport Then
        strEntrepotWhere = strEntrepotWhere & _
                           "E.Entrepot_Type = '" & Mid$(mvarEntrepot_Num, 1, 1) & "' AND " & _
                           "E.Entrepot_Num = '" & Mid$(mvarEntrepot_Num, 3) & "'"
    End If

    If Len(strEntrepotWhere) > 10 Then strEntrepotWhere = strEntrepotWhere & " AND "
    strEntrepotWhere = strEntrepotWhere & "M1.[Internal Code] = '8.29801619052887E+19' AND M2.[Internal Code] = '8.29801619052887E+19' "
    
    Select Case Index
        Case 0
            ' The primary key is mentioned twice to conform to the design of the picklist class.
            ' Products, Description, Taric Code, Entrepot No.
            
            'allanSQL
            strProductsSQL = vbNullString
            strProductsSQL = strProductsSQL & "SELECT "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [ID], "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [Prod_ID], "
            strProductsSQL = strProductsSQL & "P.Prod_Handling as [Prod_Handling], "
            strProductsSQL = strProductsSQL & "P.Prod_Num AS [Product Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Desc AS [Description], "
            strProductsSQL = strProductsSQL & "P.Taric_Code AS [Taric Code], "
            strProductsSQL = strProductsSQL & "P.Entrepot_ID AS [Entrepot ID], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Type AS [Entrepot Type], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Num AS [Entrepot Num], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Starting_Num AS [Starting Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin AS [Origin Code], "
            strProductsSQL = strProductsSQL & "M1.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Origin Description], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export AS [Export Code], "
            strProductsSQL = strProductsSQL & "M2.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Export Description] "
            strProductsSQL = strProductsSQL & "FROM "
            strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "( "
                    strProductsSQL = strProductsSQL & "Products AS [P] "
                    strProductsSQL = strProductsSQL & "INNER JOIN "
                    strProductsSQL = strProductsSQL & "Entrepots AS [E] "
                    strProductsSQL = strProductsSQL & "ON "
                    strProductsSQL = strProductsSQL & "P.Entrepot_ID = E.Entrepot_ID "
                strProductsSQL = strProductsSQL & ") "
                strProductsSQL = strProductsSQL & "INNER JOIN "
                strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 "
                strProductsSQL = strProductsSQL & "ON "
                strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin = M1.Code "
            strProductsSQL = strProductsSQL & ") "
            strProductsSQL = strProductsSQL & "INNER JOIN "
            strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 "
            strProductsSQL = strProductsSQL & "ON "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export = M2.Code "
            strProductsSQL = strProductsSQL & "WHERE "
            strProductsSQL = strProductsSQL & "M1.[Internal Code] = '8.29801619052887E+19' "
            strProductsSQL = strProductsSQL & "AND "
            strProductsSQL = strProductsSQL & "M2.[Internal Code] = '8.29801619052887E+19' "
            strProductsSQL = strProductsSQL & "ORDER BY P.Prod_Num"

            With pckProducts
                'This will be used to create a non inner joined recordset for commit purposes.
                'allanSQL check
                .DeleteSQL4InnerJoinCases = "SELECT Prod_ID, " & _
                                            "Prod_Num AS [Product Num], " & _
                                            "Prod_Desc AS Description, " & _
                                            "Taric_Code AS [Taric Code], " & _
                                            "Prod_Ctry_Origin AS [Origin Code], " & _
                                            "Prod_Archive, " & _
                                            "Entrepot_ID AS [Entrepot ID], " & _
                                            "Prod_Handling, " & _
                                            "Prod_Ctry_Export AS [Export Code], " & _
                                            "Prod_Packages_Type FROM Products"    'IAN 01-18-04
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick CallingForm, cpiFilterCatalog, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, clsPicklistFilter, , False, cpiKeyF2, strEntrepotFields  ' , True, "Product Num"
                
                If .CancelTrans = False Then
                    
                End If
            End With
            
        Case 1
            ' The primary key is mentioned twice to conform to the design of the picklist class.
            ' Products, Description, Taric Code, Entrepot No.
            'allanSQL
            strProductsSQL = vbNullString
            strProductsSQL = strProductsSQL & "SELECT "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [ID], "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [Prod_ID], "
            strProductsSQL = strProductsSQL & "P.Prod_Handling as [Prod_Handling], "
            strProductsSQL = strProductsSQL & "P.Prod_Num AS [Product Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Desc AS [Description], "
            strProductsSQL = strProductsSQL & "P.Taric_Code AS [Taric Code], "
            strProductsSQL = strProductsSQL & "P.Entrepot_ID AS [Entrepot ID], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Type AS [Entrepot Type], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Num AS [Entrepot Num], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Starting_Num AS [Starting Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin AS [Origin Code], "
            strProductsSQL = strProductsSQL & "M1.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Origin Description], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export AS [Export Code], "
            strProductsSQL = strProductsSQL & "M2.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Export Description] "
            strProductsSQL = strProductsSQL & "FROM "
            strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "( "
                    strProductsSQL = strProductsSQL & "Products AS [P] "
                    strProductsSQL = strProductsSQL & "INNER JOIN "
                    strProductsSQL = strProductsSQL & "Entrepots AS [E] "
                    strProductsSQL = strProductsSQL & "ON "
                    strProductsSQL = strProductsSQL & "P.Entrepot_ID = E.Entrepot_ID "
                strProductsSQL = strProductsSQL & ") "
                strProductsSQL = strProductsSQL & "INNER JOIN "
                strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 "
                strProductsSQL = strProductsSQL & "ON "
                strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin = M1.Code "
            strProductsSQL = strProductsSQL & ") "
            strProductsSQL = strProductsSQL & "INNER JOIN "
            strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 "
            strProductsSQL = strProductsSQL & "ON "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export = M2.Code "
            strProductsSQL = strProductsSQL & strEntrepotWhere & " "
            
            'This case was designed specifically for PCubeLibEntrepot.frmSummaryReports
            With pckProducts
                If UseSimple = True Then
                    .Search True, "Product Num", strTxtBox
                    .Pick CallingForm, cpiSimplePicklist, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, , , True, cpiKeyF2, strEntrepotFields ' ,  True, "Product Num"
                    If .CancelTrans = False And Not (.SelectedRecord Is Nothing) Then
                        mvarProduct_ID = pckProducts.SelectedRecord.RecordSource.Fields("Prod_ID").Value
                        mvarProduct_Num = pckProducts.SelectedRecord.RecordSource.Fields("Product Num").Value
                        mvarEntrepot_ID = .SelectedRecord.RecordSource.Fields("Entrepot ID").Value
                        mvarEntrepot_Type = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value
                        mvarEntrepot_Num_Only = .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                        mvarEntrepot_Num = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value & "-" & .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                    End If
                    mvarCancelled = .CancelTrans
                Else
                    If UCase(CallingForm.Name) = "FRMINITIALSTOCK" Then
                        m_blnCheckSelectedProduct = True
                    Else
                        m_blnCheckSelectedProduct = False
                    End If
                    
                    If Trim(strTxtBox) <> "" Then
                        .Search True, "Product Num", Trim(strTxtBox)
                    End If
                    clsPicklistFilter.PicklistFilters.Add "Key", "Products", "Prod_ID <> 0", True
                    
                    'allanSQL check
                    .DeleteSQL4InnerJoinCases = "SELECT Prod_ID, " & _
                                                "Prod_Num AS [Product Num], " & _
                                                "Prod_Desc AS Description, " & _
                                                "Taric_Code AS [Taric Code], " & _
                                                "Prod_Ctry_Origin AS [Origin Code], " & _
                                                "Prod_Archive, " & _
                                                "Entrepot_ID AS [Entrepot ID], " & _
                                                "Prod_Handling, " & _
                                                "Prod_Ctry_Export AS [Export Code], " & _
                                                "Prod_Packages_Type FROM Products"    'IAN 01-18-04
                    
                    ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                    .Pick CallingForm, cpiFilterCatalog, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, clsPicklistFilter, , True, cpiKeyF2, strEntrepotFields ', True, "Product Num"
                    
                    If Not .SelectedRecord Is Nothing Then
                        mvarProduct_ID = pckProducts.SelectedRecord.RecordSource.Fields("Prod_ID").Value
                        mvarProduct_Num = pckProducts.SelectedRecord.RecordSource.Fields("Product Num").Value
                        mvarProd_Desc = .SelectedRecord.RecordSource.Fields("Description").Value
                        mvarProd_Handling = .SelectedRecord.RecordSource.Fields("Prod_Handling").Value
                        mvarCtry_Origin = .SelectedRecord.RecordSource.Fields("Origin Code").Value
                        Me.Ctry_Export = .SelectedRecord.RecordSource.Fields("Export Code").Value
                        mvarOrigin_Desc = .SelectedRecord.RecordSource.Fields("Origin Description").Value
                        mvarExport_Desc = .SelectedRecord.RecordSource.Fields("Export Description").Value
                        mvarTaric_Code = .SelectedRecord.RecordSource.Fields("Taric Code").Value
                        mvarEntrepot_ID = .SelectedRecord.RecordSource.Fields("Entrepot ID").Value
                        mvarEntrepot_Type = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value
                        mvarEntrepot_Num_Only = .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                        mvarEntrepot_Num = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value & "-" & .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                        
                        mvarCancelled = pckProducts.CancelTrans
                    Else
                        mvarCancelled = True
                    End If

                End If
            End With
            
        Case 2
            ' The primary key is mentioned twice to conform to the design of the picklist class.
            ' Products, Description, Taric Code, Entrepot No.
            'allanSQL
            strProductsSQL = vbNullString
            strProductsSQL = strProductsSQL & "SELECT "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [ID], "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [Prod_ID], "
            strProductsSQL = strProductsSQL & "P.Prod_Handling AS [Prod_Handling], "
            strProductsSQL = strProductsSQL & "P.Prod_Num AS [Product Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Desc AS [Description], "
            strProductsSQL = strProductsSQL & "P.Taric_Code AS [Taric Code], "
            strProductsSQL = strProductsSQL & "P.Entrepot_ID AS [Entrepot ID], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Type AS [Entrepot Type], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Num AS [Entrepot Num], "
            strProductsSQL = strProductsSQL & "E.Entrepot_StockCard_Numbering AS [Numbering], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Starting_Num AS [Starting Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin AS [Origin Code], "
            strProductsSQL = strProductsSQL & "M1.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Origin Description], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export AS [Export Code], "
            strProductsSQL = strProductsSQL & "M2.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Export Description] "
            strProductsSQL = strProductsSQL & "FROM "
            strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "Products AS [P] "
                strProductsSQL = strProductsSQL & "INNER JOIN "
                strProductsSQL = strProductsSQL & "Entrepots AS [E] "
                strProductsSQL = strProductsSQL & "ON "
                strProductsSQL = strProductsSQL & "P.Entrepot_ID = E.Entrepot_ID "
            strProductsSQL = strProductsSQL & ") "
            strProductsSQL = strProductsSQL & "INNER JOIN "
            strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 "
            strProductsSQL = strProductsSQL & "ON "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin = M1.Code "
        strProductsSQL = strProductsSQL & ") "
        strProductsSQL = strProductsSQL & "INNER JOIN "
        strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 "
        strProductsSQL = strProductsSQL & "ON "
        strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export = M2.Code "
        strProductsSQL = strProductsSQL & strEntrepotWhere & " "

            'This case was designed specifically for PCubeLibEntrepot.frmStockcard
            With pckProducts
                .Search True, "Product Num", Trim(CallingForm.txtProductNo.Text)
                
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick CallingForm, cpiSimplePicklist, mConnection, strProductsSQL, "ID", "Product Num", vbModal, gsdProducts, , , True, cpiKeyF2, strEntrepotFields
                
                If Not .SelectedRecord Is Nothing Then
                    CallingForm.txtProductNo.Text = .SelectedRecord.RecordSource.Fields("Product Num").Value
                    CallingForm.txtProductNo.Tag = .SelectedRecord.RecordSource.Fields("Prod_ID").Value
                    CallingForm.lblEntrepotNumX.Tag = .SelectedRecord.RecordSource.Fields("Entrepot ID").Value
                    CallingForm.strEntrepotType = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value
                    CallingForm.strEntrepotNum = .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                    CallingForm.bytSequential = .SelectedRecord.RecordSource.Fields("Numbering").Value
                    CallingForm.strSeqStart = .SelectedRecord.RecordSource.Fields("Starting Num").Value
                    
''                   --------- Will obtain required fields below within Stock Card Maintainance --------
''                    CallingForm.lngStockNum = .SelectedRecord.RecordSource.Fields("Stock Card Number").Value
''                   -----------------------------------------------------------------------------------
                    mvarCancelled = pckProducts.CancelTrans
                Else
                    mvarCancelled = True
                
                End If
            End With
            strProductsSQL = Empty
                    
        Case 3
            ' The primary key is mentioned twice to conform to the design of the picklist class.
            ' Products, Description, Taric Code, Entrepot No.
            'allanSQL
            strProductsSQL = vbNullString
            strProductsSQL = strProductsSQL & "SELECT "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [ID], "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [Prod_ID], "
            strProductsSQL = strProductsSQL & "P.Prod_Handling as [Prod_Handling], "
            strProductsSQL = strProductsSQL & "P.Prod_Num AS [Product Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Desc AS [Description], "
            strProductsSQL = strProductsSQL & "P.Taric_Code AS [Taric Code], "
            strProductsSQL = strProductsSQL & "P.Entrepot_ID AS [Entrepot ID], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Type AS [Entrepot Type], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Num AS [Entrepot Num], "
            strProductsSQL = strProductsSQL & "E.Entrepot_StockCard_Numbering AS [Numbering], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Starting_Num AS [Starting Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin AS [Origin Code], "
            strProductsSQL = strProductsSQL & "M1.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Origin Description], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export AS [Export Code], "
            strProductsSQL = strProductsSQL & "M2.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Export Description] "
            strProductsSQL = strProductsSQL & "FROM "
            strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "( "
                    strProductsSQL = strProductsSQL & "Products AS [P] "
                    strProductsSQL = strProductsSQL & "INNER JOIN "
                    strProductsSQL = strProductsSQL & "Entrepots AS [E] "
                    strProductsSQL = strProductsSQL & "ON "
                    strProductsSQL = strProductsSQL & "P.Entrepot_ID = E.Entrepot_ID "
                strProductsSQL = strProductsSQL & ") "
                strProductsSQL = strProductsSQL & "INNER JOIN "
                strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 "
                strProductsSQL = strProductsSQL & "ON "
                strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin = M1.Code "
            strProductsSQL = strProductsSQL & ") "
            strProductsSQL = strProductsSQL & "INNER JOIN "
            strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 "
            strProductsSQL = strProductsSQL & "ON "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export = M2.Code "
            strProductsSQL = strProductsSQL & strEntrepotWhere & " "
            
            'This case was designed specifically for PCubeLibEntrepot.frmStockProdPicklist
            With pckProducts
                .Search True, "Product Num", Trim(CallingForm.txtProductNo.Text)
                
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                If CallingForm.cmdNew.Visible = True Then
                    
                    clsPicklistFilter.PicklistFilters.Add "Key", "Products", "Prod_ID <> 0", True
                    
                    .DeleteSQL4InnerJoinCases = "SELECT Prod_ID, " & _
                                                "Prod_Num AS [Product Num], " & _
                                                "Prod_Desc AS Description, " & _
                                                "Taric_Code AS [Taric Code], " & _
                                                "Prod_Ctry_Origin AS [Origin Code], " & _
                                                "Prod_Archive, " & _
                                                "Entrepot_ID AS [Entrepot ID], " & _
                                                "Prod_Handling, " & _
                                                "Prod_Ctry_Export AS [Export Code], " & _
                                                "Prod_Packages_Type FROM Products"    'IAN 01-18-04
                                                
                    'This is called if StockProd was called from codisheet.
                    .Pick CallingForm, cpiFilterCatalog, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, clsPicklistFilter, , True, cpiKeyF2, strEntrepotFields
                ElseIf CallingForm.cmdNew.Visible = False Then
                    'This is called if StockProd was called from Summary Report.
                    .Pick CallingForm, cpiSimplePicklist, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, , , True, cpiKeyF2, strEntrepotFields ', True, "Product Num"
                End If
                
                If Not .SelectedRecord Is Nothing Then
                                            
                    Me.Ctry_Origin = .SelectedRecord.RecordSource.Fields("Origin Code").Value
                    Me.Ctry_Export = .SelectedRecord.RecordSource.Fields("Export Code").Value
                    Me.Origin_Desc = .SelectedRecord.RecordSource.Fields("Origin Description").Value
                    Me.Export_Desc = .SelectedRecord.RecordSource.Fields("Export Description").Value
                    Me.Product_Num = .SelectedRecord.RecordSource.Fields("Product Num").Value
                    Me.Product_ID = .SelectedRecord.RecordSource.Fields("Prod_ID").Value
                    Me.Prod_Desc = .SelectedRecord.RecordSource.Fields("Description").Value
                    Me.Taric_Code = .SelectedRecord.RecordSource.Fields("Taric Code").Value
                    Me.Entrepot_ID = .SelectedRecord.RecordSource.Fields("Entrepot ID").Value
                    Me.Entrepot_Type = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value
                    Me.Entrepot_Num_Only = .SelectedRecord.RecordSource.Fields("Entrepot Num").Value
                    Me.Numbering = .SelectedRecord.RecordSource.Fields("Numbering").Value
                    Me.StartingNum = .SelectedRecord.RecordSource.Fields("Starting Num").Value

                    'dto ko na lng nga ilagay prod_handling...sana cls na lng nakapass instead na form... -alg
                    Me.Prod_Handling = .SelectedRecord.RecordSource.Fields("Prod_Handling").Value
                    
                    mvarCancelled = pckProducts.CancelTrans
                Else
                    mvarCancelled = True
                End If
            End With
            strProductsSQL = Empty
            
            
            
        Case 4
            Set gsdProducts = pckProducts.SeedGrid("Product Num", 1300, "Left", "Description", 2000, "Left", "Taric Code", 1100, "Left", "Entrepot Num", 1050, "Left", "Origin Code", 1000, "Left", _
            "Origin Description", 1500, "Left", "Export Code", 1000, "Left", "Export Description", 1500, "Left")
            
            strEntrepotFields = "Prod_ID,ID,Entrepot Num,Entrepot_ID,Auth_ID,Entrepot_StartDate,Entrepot_EndDate," & _
            "Entrepot_StockCard_Numbering,Entrepot_Starting_Num,Entrepot_StockCard_Format,Entrepot_StockCard_Archiving,Entrepot_LastSeqNum,Origin Description,Export Description,M1.Internal Code,M1.Code,M2.Internal Code,M2.Code"
                        
            ' The primary key is mentioned twice to conform to the design of the picklist class.
            ' Products, Description, Taric Code, Entrepot No.
            'allanSQL
            strProductsSQL = vbNullString
            strProductsSQL = strProductsSQL & "SELECT "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [ID], "
            strProductsSQL = strProductsSQL & "P.Prod_ID AS [Prod_ID], "
            strProductsSQL = strProductsSQL & "P.Prod_Handling as [Prod_Handling], "
            strProductsSQL = strProductsSQL & "P.Prod_Num AS [Product Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Desc AS [Description], "
            strProductsSQL = strProductsSQL & "P.Taric_Code AS [Taric Code], "
            strProductsSQL = strProductsSQL & "P.Entrepot_ID AS [Entrepot ID], "
            strProductsSQL = strProductsSQL & "E.Entrepot_Type  & '-' & E.Entrepot_Num AS [Entrepot Num], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin AS [Origin Code], "
            strProductsSQL = strProductsSQL & "M1.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Origin Description], "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export AS [Export Code], "
            strProductsSQL = strProductsSQL & "M2.[Description " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS [Export Description] "
            strProductsSQL = strProductsSQL & "FROM "
            strProductsSQL = strProductsSQL & "( "
                strProductsSQL = strProductsSQL & "( "
                    strProductsSQL = strProductsSQL & "Products AS [P] "
                    strProductsSQL = strProductsSQL & "INNER JOIN "
                    strProductsSQL = strProductsSQL & "Entrepots AS [E]  "
                    strProductsSQL = strProductsSQL & "ON "
                    strProductsSQL = strProductsSQL & "P.Entrepot_ID = E.Entrepot_ID "
                strProductsSQL = strProductsSQL & ") "
                strProductsSQL = strProductsSQL & "INNER JOIN "
                strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M1 "
                strProductsSQL = strProductsSQL & "ON "
                strProductsSQL = strProductsSQL & "P.Prod_Ctry_Origin = M1.Code "
            strProductsSQL = strProductsSQL & ") "
            strProductsSQL = strProductsSQL & "INNER JOIN "
            strProductsSQL = strProductsSQL & "[Picklist Maintenance " & IIf(UCase(mLanguage) = "ENGLISH", "ENGLISH", IIf(UCase(mLanguage) = "FRENCH", "FRENCH", "DUTCH")) & "] AS M2 "
            strProductsSQL = strProductsSQL & "ON "
            strProductsSQL = strProductsSQL & "P.Prod_Ctry_Export = M2.Code "
            strProductsSQL = strProductsSQL & "WHERE "
            strProductsSQL = strProductsSQL & "M1.[Internal Code] = '8.29801619052887E+19' "
            strProductsSQL = strProductsSQL & "AND "
            strProductsSQL = strProductsSQL & "M2.[Internal Code] = '8.29801619052887E+19' "
            strProductsSQL = strProductsSQL & "ORDER BY "
            strProductsSQL = strProductsSQL & "P.Prod_Num ASC "
            'This case was designed specifically when accessed via F4 under Box L1 of Import codisheet.
            'Does not have button due to Entrepot Type & "-" & Entrepot Num
            With pckProducts
                'This will be used to create a non inner joined recordset for commit purposes.
                .DeleteSQL4InnerJoinCases = "SELECT Prod_ID, " & _
                                            "Prod_Num AS [Product Num], " & _
                                            "Prod_Desc AS Description, " & _
                                            "Taric_Code AS [Taric Code], " & _
                                            "Prod_Ctry_Origin AS [Origin Code], " & _
                                            "Prod_Archive, " & _
                                            "Entrepot_ID AS [Entrepot ID], " & _
                                            "Prod_Handling, " & _
                                            "Prod_Ctry_Export AS [Export Code], " & _
                                            "Prod_Packages_Type FROM Products"    'IAN 01-18-04
                
                .AddButtonVisible = False
                .ModifyButtonVisible = False
                .CopyButtonVisible = False
                .DeleteButtonVisible = False
                
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick CallingForm, cpiFilterCatalog, mConnection, strProductsSQL, "ID", "Products", vbModal, gsdProducts, clsPicklistFilter, , False, cpiKeyF2, strEntrepotFields

                If .CancelTrans = False Then
                
                End If
            End With
    End Select
    
    Set gsdProducts = Nothing
    Set pckProducts = Nothing
    Set clsPicklistFilter = Nothing
End Sub

Private Sub pckProducts_BeforeOK(ByVal SelectedRecord As PCubeLibPick.CRecord, Cancel As Boolean)
    If m_blnCheckSelectedProduct = True Then
        If UCase(SelectedRecord.RecordSource.Fields("Entrepot Type").Value & "-" & SelectedRecord.RecordSource.Fields("Entrepot Num")) <> UCase(mvarEntrepot_Num) Then
            Cancel = True
            MsgBox "The product you have selected is not under the currently defined entrepot name. Please select a different product number.", vbInformation, Translate(2216)
        End If
    End If
End Sub




Private Sub pckProducts_BeforeDelete(ByVal BaseName As String, ByVal ID As Variant, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Dim rstChkProduct As ADODB.Recordset
    Dim strSQL As String 'allanSQL

        'allanSQL
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Prod_ID AS [Prod ID], "
        strSQL = strSQL & "Stock_Card_Num AS [Stock Card Num] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "StockCards "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Prod_ID = " & ID & " "
    ADORecordsetOpen strSQL, mConnection, rstChkProduct, adOpenKeyset, adLockOptimistic
    'rstChkProduct.Open strSQL, mConnection, adOpenKeyset, adLockOptimistic
    If Not (rstChkProduct.BOF And rstChkProduct.EOF) Then
        MsgBox Translate(2290) & vbCrLf & _
               rstChkProduct.RecordCount & Space(1) & Translate(2293), _
               vbOKOnly + vbInformation, "Stock Cards"
        Cancel = True
    Else
        Cancel = False
    End If
    
    ADORecordsetClose rstChkProduct
End Sub

Private Sub pckProducts_BtnClick(Record As PCubeLibPick.CRecord, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Select Case Button
        Case cpiDelete
        Case Else
            frmProducts.MyLoad Record.RecordSource, Button, Cancel, mLanguage, mConnection, mConnectionTaric, mTaricProperties, ResourceHandler
    End Select
End Sub
