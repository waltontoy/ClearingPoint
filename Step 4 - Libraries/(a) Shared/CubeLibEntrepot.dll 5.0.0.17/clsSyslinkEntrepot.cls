VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyslinkEntrepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_blnEntrepot As Boolean
Private m_blnInbound As Boolean
Private m_strEntrepotNum As String
Private m_strProdNum As String
Private m_strStockNum As String
Private m_strJobNum As String
Private m_strBatchNum As String
Private m_strKey As String
Private m_lngStock_ID As Long
Private m_lngIn_ID As Long
Private m_strQuantity As String
Private m_strNetWeight As String
Private m_strGrossWeight As String
Private m_strTaricDesc As String
Private m_strTaricCode As String
Private m_blnIsENTNR As String

Public Property Let Key(vKey As String)
    m_strKey = vKey
End Property

Public Property Get Key() As String
    Key = m_strKey
End Property

Public Property Let ValidEntrepot(vEntrepot As Boolean)
    m_blnEntrepot = vEntrepot
End Property

Public Property Get ValidEntrepot() As Boolean
    ValidEntrepot = m_blnEntrepot
End Property

Public Property Let ValidInbound(vInbound As Boolean)
    m_blnInbound = vInbound
End Property

Public Property Get ValidInbound() As Boolean
    ValidInbound = m_blnInbound
End Property

Public Property Let EntrepotNum(vEntrepotNum As String)
    m_strEntrepotNum = vEntrepotNum
End Property

Public Property Get EntrepotNum() As String
    EntrepotNum = m_strEntrepotNum
End Property

Public Property Let ProdNum(vProdNum As String)
    m_strProdNum = vProdNum
End Property

Public Property Get ProdNum() As String
    ProdNum = m_strProdNum
End Property

Public Property Let StockNum(vStockNum As String)
    m_strStockNum = vStockNum
End Property

Public Property Get StockNum() As String
    StockNum = m_strStockNum
End Property

Public Property Let JobNum(vJobNum As String)
    m_strJobNum = vJobNum
End Property

Public Property Get JobNum() As String
    JobNum = m_strJobNum
End Property

Public Property Let BatchNum(vBatchNum As String)
    m_strBatchNum = vBatchNum
End Property

Public Property Get BatchNum() As String
    BatchNum = m_strBatchNum
End Property

Public Property Get In_ID() As Long
    In_ID = m_lngIn_ID
End Property

Public Property Let In_ID(vIn_ID As Long)
    m_lngIn_ID = vIn_ID
End Property

Public Property Get Stock_ID() As Long
    Stock_ID = m_lngStock_ID
End Property

Public Property Let Stock_ID(vStock_ID As Long)
    m_lngStock_ID = vStock_ID
End Property

Public Property Let Quantity(vQuantity As String)
    m_strQuantity = vQuantity
End Property

Public Property Get Quantity() As String
    Quantity = m_strQuantity
End Property

Public Property Let NetWeight(vNetWeight As String)
    m_strNetWeight = vNetWeight
End Property

Public Property Get NetWeight() As String
    NetWeight = m_strNetWeight
End Property

Public Property Let GrossWeight(vGrossWeight As String)
    m_strGrossWeight = vGrossWeight
End Property

Public Property Get GrossWeight() As String
    GrossWeight = m_strGrossWeight
End Property

Public Property Let TaricDesc(vTaricDesc As String)
    m_strTaricDesc = vTaricDesc
End Property

Public Property Get TaricDesc() As String
    TaricDesc = m_strTaricDesc
End Property

Public Property Let TaricCode(vTaricCode As String)
    m_strTaricCode = vTaricCode
End Property

Public Property Get TaricCode() As String
    TaricCode = m_strTaricCode
End Property

Public Property Let IsENTNR(vIsENTNR As Boolean)
    m_blnIsENTNR = vIsENTNR
End Property

Public Property Get IsENTNR() As Boolean
    IsENTNR = m_blnIsENTNR
End Property
