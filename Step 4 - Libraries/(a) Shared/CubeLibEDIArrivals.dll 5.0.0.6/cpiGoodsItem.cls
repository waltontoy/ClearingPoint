VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiGoodsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_CODE As String ' 1
Private mvarFIELD_COMMODITY_CODE As String ' 2
Private mvarFIELD_GOODS_DESCRIPTION As String ' 3
Private mvarFIELD_GOODS_DESCRIPTION_LNG As String ' 4
Private mvarFIELD_ITEM_NUMBER As String ' 5
Private mvarFIELD_GROSS_MASS As String ' 6
Private mvarFIELD_NET_MASS As String ' 7
Private mvarFIELD_ORDINAL_PARENT As Long ' 8
Private mvarFIELD_ORDINAL As Long ' 9

Private mvarFIELD_T7 As String

' as   cpiDocCertificates
Private mvarContainers As cpiContainers
Private mvarPackages As cpiPackages
Private mvarSGICodes As cpiSGICodes
Private mvarDocumentCertificates As cpiDocCertificates
Private mvarResultOfControls As cpiResultOfControls  ' 9d

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_COMMODITY_CODE(ByVal vData As String)
   mvarFIELD_COMMODITY_CODE = vData
End Property

Public Property Get FIELD_COMMODITY_CODE() As String
   FIELD_COMMODITY_CODE = mvarFIELD_COMMODITY_CODE
End Property

'  3
Public Property Let FIELD_GOODS_DESCRIPTION(ByVal vData As String)
   mvarFIELD_GOODS_DESCRIPTION = vData
End Property

Public Property Get FIELD_GOODS_DESCRIPTION() As String
   FIELD_GOODS_DESCRIPTION = mvarFIELD_GOODS_DESCRIPTION
End Property

'  4
Public Property Let FIELD_GOODS_DESCRIPTION_LNG(ByVal vData As String)
   mvarFIELD_GOODS_DESCRIPTION_LNG = vData
End Property

Public Property Get FIELD_GOODS_DESCRIPTION_LNG() As String
   FIELD_GOODS_DESCRIPTION_LNG = mvarFIELD_GOODS_DESCRIPTION_LNG
End Property

'  5
Public Property Let FIELD_ITEM_NUMBER(ByVal vData As String)
   mvarFIELD_ITEM_NUMBER = vData
End Property

Public Property Get FIELD_ITEM_NUMBER() As String
   FIELD_ITEM_NUMBER = mvarFIELD_ITEM_NUMBER
End Property

'  6
Public Property Let FIELD_GROSS_MASS(ByVal vData As String)
   mvarFIELD_GROSS_MASS = vData
End Property

Public Property Get FIELD_GROSS_MASS() As String
   FIELD_GROSS_MASS = mvarFIELD_GROSS_MASS
End Property

'  7
Public Property Let FIELD_NET_MASS(ByVal vData As String)
   mvarFIELD_NET_MASS = vData
End Property

Public Property Get FIELD_NET_MASS() As String
   FIELD_NET_MASS = mvarFIELD_NET_MASS
End Property

'  8
Public Property Let FIELD_ORDINAL_PARENT(ByVal vData As Long)
   mvarFIELD_ORDINAL_PARENT = vData
End Property

Public Property Get FIELD_ORDINAL_PARENT() As Long
   FIELD_ORDINAL_PARENT = mvarFIELD_ORDINAL_PARENT
End Property

'  9
Public Property Let FIELD_ORDINAL(ByVal vData As Long)
   mvarFIELD_ORDINAL = vData
End Property

Public Property Get FIELD_ORDINAL() As Long
   FIELD_ORDINAL = mvarFIELD_ORDINAL
End Property

' 10
Public Property Set Containers(ByVal vData As cpiContainers)
    Set mvarContainers = vData
End Property

Public Property Get Containers() As cpiContainers
    Set Containers = mvarContainers
End Property

' 11
Public Property Set Packages(ByVal vData As cpiPackages)
    Set mvarPackages = vData
End Property

Public Property Get Packages() As cpiPackages
    Set Packages = mvarPackages
End Property

' 12
Public Property Set SGICodes(ByVal vData As cpiSGICodes)
    Set mvarSGICodes = vData
End Property

Public Property Get SGICodes() As cpiSGICodes
    Set SGICodes = mvarSGICodes
End Property

' 13
Public Property Set DocumentCertificates(ByVal vData As cpiDocCertificates)
    Set mvarDocumentCertificates = vData
End Property

Public Property Get DocumentCertificates() As cpiDocCertificates
    Set DocumentCertificates = mvarDocumentCertificates
End Property

' 14
Public Property Set ResultOfControls(ByVal vData As cpiResultOfControls)
    Set mvarResultOfControls = vData
End Property

Public Property Get ResultOfControls() As cpiResultOfControls
    Set ResultOfControls = mvarResultOfControls
End Property

Public Property Let FIELD_T7(ByVal vData As String)
    mvarFIELD_T7 = vData
End Property
Public Property Get FIELD_T7() As String
    FIELD_T7 = mvarFIELD_T7
End Property

