VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStockProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarJobNo As String 'local copy
Private mvarBatchNo As String 'local copy
Private mvarProduct_ID As Long 'local copy
Private mvarProductNo As String 'local copy
Private mvarStockCardNo As String 'local copy
Private mvarTaricCode As String 'local copy
Private mvarStock_ID As Long 'local copy
Private mvarCancel As Boolean 'local copy
Private mvarEntrepot_Num As String
Private mvarCtryOrigin As String 'local copy
Private mvarCtryExport As String 'local copy
Private mvarMemo As String 'local copy
Private mvarProductDesc As String 'local copy
Private mvarProductHandling As Byte 'local copy

'used for passing required values to products picklist form.
'Private mvarCallingForm As Object
Public m_conSADBEL As ADODB.Connection
Public m_conTaric As ADODB.Connection
Public strLanguage As String
Public intTaricProperties As Integer

Public Property Let ProductHandling(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductHandling = 5
    mvarProductHandling = vData
End Property

Public Property Get ProductHandling() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductHandling
    ProductHandling = mvarProductHandling
End Property

Public Property Let ProductDesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductDesc = 5
    mvarProductDesc = vData
End Property

Public Property Get ProductDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductDesc
    ProductDesc = mvarProductDesc
End Property

Public Property Let Product_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Product_ID = 5
    mvarProduct_ID = vData
End Property

Public Property Get Product_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Product_ID
    Product_ID = mvarProduct_ID
End Property

Public Property Let CtryExport(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CtryExport = 5
    mvarCtryExport = vData
End Property

Public Property Get CtryExport() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CtryExport
    CtryExport = mvarCtryExport
End Property

Public Property Let CtryOrigin(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CtryOrigin = 5
    mvarCtryOrigin = vData
End Property

Public Property Get CtryOrigin() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CtryOrigin
    CtryOrigin = mvarCtryOrigin
End Property

Public Property Let Memo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Memo = 5
    mvarMemo = vData
End Property

Public Property Get Memo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Memo
    Memo = mvarMemo
End Property

Public Property Let Entrepot_Num(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Entrepot_Num = 5
    mvarEntrepot_Num = vData
End Property

Public Property Get Entrepot_Num() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Entrepot_Num
    Entrepot_Num = mvarEntrepot_Num
End Property

Public Sub ShowPicklist(CallingForm As Object, _
                        ByVal conn_Sadbel As ADODB.Connection, _
                        ByVal conn_Taric As ADODB.Connection, _
                        Language As String, _
                        TaricProperties As Integer, _
                        ByVal ResourceHandler As Long, _
                        Optional ByVal blnDontShowBatchJob As Boolean, _
                        Optional ByVal blnInitialStock As Boolean, _
                        Optional ByVal blnWithEntrepot As Boolean, _
                        Optional ByVal strEntrepotNum As String, _
                        Optional blnFromSummaryReports As Boolean = False)

                        
'    Set mvarCallingForm = CallingForm
    Set m_conSADBEL = conn_Sadbel
    Set m_conTaric = conn_Taric
    strLanguage = Language
    intTaricProperties = TaricProperties
                        
    frmStockProdPicklist.Pre_Load Me, mvarCancel, ResourceHandler, blnDontShowBatchJob, blnInitialStock, blnWithEntrepot, strEntrepotNum, blnFromSummaryReports
    
    Set frmStockProdPicklist = Nothing
End Sub

Public Property Let StockCardNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StockCardNo = 5
    mvarStockCardNo = vData
End Property

Public Property Get StockCardNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StockCardNo
    StockCardNo = mvarStockCardNo
End Property

Public Property Let ProductNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductNo = 5
    mvarProductNo = vData
End Property

Public Property Get ProductNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductNo
    ProductNo = mvarProductNo
End Property

Public Property Let BatchNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BatchNo = 5
    mvarBatchNo = vData
End Property

Public Property Get BatchNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BatchNo
    BatchNo = mvarBatchNo
End Property

Public Property Let JobNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JobNo = 5
    mvarJobNo = vData
End Property

Public Property Get JobNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JobNo
    JobNo = mvarJobNo
End Property

Public Property Let TaricCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TaricCode = 5
    mvarTaricCode = vData
End Property

Public Property Get TaricCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaricCode
    TaricCode = mvarTaricCode
End Property

Public Property Let Stock_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stock_ID = 5
    mvarStock_ID = vData
End Property

Public Property Get Stock_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stock_ID
    Stock_ID = mvarStock_ID
End Property

Public Property Let Cancel(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cancel = 5
    mvarCancel = vData
End Property

Public Property Get Cancel() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cancel
    Cancel = mvarCancel
End Property

Private Sub Class_Terminate()
    Set m_conSADBEL = Nothing
    Set m_conTaric = Nothing
'    strLanguage = Empty
'    intTaricProperties = Empty
End Sub
