VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAvailableStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarStock_Num As String 'local copy
'Private mvarRequest As Double 'local copy
Private mvarStock_ID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarEntrepot_Num As String 'local copy
'local variable(s) to hold property value(s)
'Private mvarQtyToReserve As Double 'local copy
Private mvarQtyToReserve As String 'local copy
'local variable(s) to hold property value(s)
Private mvarTaricCode As String 'local copy
Private mvarProductNum As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSelectedRecord As Boolean 'local copy

Private mvarTaricQueryName As String
Public Enum eTaricDisplay
    eTaric_None = 0
    eTaric_Simplified
    eTaric_Full
    
End Enum

''Public Enum eCodiType
''    eCodi_Import = 0
''    eCodi_Export = 1
''    eCodi_OTS = 2
''    eCodi_TransitNCTS = 3
''    eCodi_CombinedNCTS = 4
''    eCodi_EDIDeparture = 5
''End Enum
    
Public Enum eProdHandling
    eProd_Quantity = 0
    eProd_GrossWeight = 1
    eProd_NetWeight = 2
End Enum

'local variable(s) to hold property value(s)
Private mvarTaricDisplay As eTaricDisplay 'local copy
'local variable(s) to hold property value(s)
Private mvarCodiType As eCodiType 'local copy
Private mvarCommon As cCommon 'local copy
'local variable(s) to hold property value(s)
Private mvarBatchNumber As String 'local copy
Private mvarJobNumber As String 'local copy
Dim WithEvents mvarAvailableStocks As frmAvailableStocks
Attribute mvarAvailableStocks.VB_VarHelpID = -1
Event cmdTransactClick(ByVal Index As Integer)
'local variable(s) to hold property value(s)
'Private mvarGrossWeight As Double 'local copy
Private mvarGrossWeight As String    'local copy
'Private mvarNetWeight As Double 'local copy
Private mvarNetWeight As String 'local copy
'local variable(s) to hold property value(s)
Private mvarIn_ID As Long 'local copy
'local variable(s) to hold property value(s)
'Private mvarPackageQuantity As Double 'local copy
Private mvarPackageQuantity As String    'local copy
'local variable(s) to hold property value(s)
Private mvarProd_Handling As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarCtryOfOrigin As String 'local copy
Private mvarCtryOfExport As String 'local copy
Public Property Let CtryOfExport(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CtryOfExport = 5
    mvarCtryOfExport = vData
End Property


Public Property Get CtryOfExport() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CtryOfExport
    CtryOfExport = mvarCtryOfExport
End Property



Public Property Let CtryOfOrigin(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CtryOfOrigin = 5
    mvarCtryOfOrigin = vData
End Property


Public Property Get CtryOfOrigin() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CtryOfOrigin
    CtryOfOrigin = mvarCtryOfOrigin
End Property



Public Property Let Prod_Handling(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prod_Handling = 5
    mvarProd_Handling = vData
End Property


Public Property Get Prod_Handling() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prod_Handling
    Prod_Handling = mvarProd_Handling
End Property



Public Property Let PackageQuantity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PackageQuantity = 5
    mvarPackageQuantity = vData
End Property


Public Property Get PackageQuantity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PackageQuantity
    PackageQuantity = mvarPackageQuantity
End Property



Public Property Let In_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IN_ID = 5
    mvarIn_ID = vData
End Property


Public Property Get In_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IN_ID
    In_ID = mvarIn_ID
End Property



Public Property Let NetWeight(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NetWeight = 5
    mvarNetWeight = vData
End Property


Public Property Get NetWeight() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NetWeight
    NetWeight = mvarNetWeight
End Property



Public Property Let GrossWeight(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GrossWeight = 5
    mvarGrossWeight = vData
End Property


Public Property Get GrossWeight() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GrossWeight
    GrossWeight = mvarGrossWeight
End Property




Public Property Let JobNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JobNumber = 5
    mvarJobNumber = vData
End Property


Public Property Get JobNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JobNumber
    JobNumber = mvarJobNumber
End Property



Public Property Let BatchNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BatchNumber = 5
    mvarBatchNumber = vData
End Property


Public Property Get BatchNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BatchNumber
    BatchNumber = mvarBatchNumber
End Property



Public Property Set Common(ByVal vData As cCommon)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Common = Form1
    Set mvarCommon = vData
End Property


Public Property Get Common() As cCommon
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Common
    Set Common = mvarCommon
End Property



Public Property Let CodiType(ByVal vData As eCodiType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodiType = 5
    mvarCodiType = vData
End Property


Public Property Get CodiType() As eCodiType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodiType
    CodiType = mvarCodiType
End Property



Public Property Let TaricDisplay(ByVal vData As eTaricDisplay)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TaricDisplay = 5
    mvarTaricDisplay = vData
End Property


Public Property Get TaricDisplay() As eTaricDisplay
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaricDisplay
    TaricDisplay = mvarTaricDisplay
End Property



Public Property Let SelectedRecord(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OKClicked = 5
    mvarSelectedRecord = vData
End Property


Public Property Get SelectedRecord() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OKClicked
    SelectedRecord = mvarSelectedRecord
End Property



Public Property Let ProductNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductNum = 5
    mvarProductNum = vData
End Property


Public Property Get ProductNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductNum
    ProductNum = mvarProductNum
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



Public Property Let QtyToReserve(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QtyToReserve = 5
    mvarQtyToReserve = vData
End Property


Public Property Get QtyToReserve() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QtyToReserve
    QtyToReserve = mvarQtyToReserve
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



''Public Property Let Request(ByVal vData As Double)
'''used when assigning a value to the property, on the left side of an assignment.
'''Syntax: X.Request = 5
''    mvarRequest = vData
''End Property
''
''
''Public Property Get Request() As Double
'''used when retrieving value of a property, on the right side of an assignment.
'''Syntax: Debug.Print X.Request
''    Request = mvarRequest
''End Property



Public Property Let Stock_Num(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stock_Num = 5
    mvarStock_Num = vData
End Property


Public Property Get Stock_Num() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stock_Num
    Stock_Num = mvarStock_Num
End Property

Public Sub Show(ByRef connSadbel As ADODB.Connection, _
                ByRef connTaric As ADODB.Connection, _
                ByVal strLanguage As String, ByVal cpiEDetails As cEntrepotDetails, _
                ByVal strTabCaption As String, ByVal ResourceHandler As Long, ByVal lngIn_ID As Long, _
                ByVal blnIsDIA As Boolean, Optional ByVal strMDBpath As String, _
                Optional ByVal strDocType As String, Optional ByVal strDocNumber As String, Optional ByVal strDocDate As String, _
                Optional ByVal strTaric As String, Optional ByVal strEntrepotNum As String)
                
    'mclsCommon.ActiveLanguage = Language
    'Set frmAvailableStocks.AvailableStocks = Me
    'frmAvailableStocks.MyLoad connSadbel, connTaric, strLanguage, Me
    Set mvarAvailableStocks = frmAvailableStocks
    Me.In_ID = lngIn_ID
    Me.TaricCode = strTaric
    
    mvarAvailableStocks.MousePointer = vbHourglass
    mvarAvailableStocks.MyLoad connSadbel, connTaric, strLanguage, Me, cpiEDetails, strTabCaption, ResourceHandler, _
                                blnIsDIA, strMDBpath, strDocType, strDocNumber, strDocDate

    Set mvarAvailableStocks = Nothing
    
End Sub

 
Public Property Let TaricQueryName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductNum = 5
    mvarTaricQueryName = vData
End Property


Public Property Get TaricQueryName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductNum
    TaricQueryName = mvarTaricQueryName
End Property

Private Sub Class_Initialize()
    Set mvarCommon = New cCommon
End Sub

Private Sub Class_Terminate()
    Set mvarCommon = Nothing
End Sub
