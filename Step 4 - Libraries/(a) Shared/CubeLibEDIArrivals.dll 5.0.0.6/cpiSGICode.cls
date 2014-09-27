VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiSGICode"
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
Private mvarFIELD_SENSITIVE_GOODS_CODE As String ' 2
Private mvarFIELD_SENSITIVE_QUANTITY As String ' 3
Private mvarFIELD_ORDINAL_PARENT As Long ' 4
Private mvarFIELD_ORDINAL As Long ' 5

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_SENSITIVE_GOODS_CODE(ByVal vData As String)
   mvarFIELD_SENSITIVE_GOODS_CODE = vData
End Property

Public Property Get FIELD_SENSITIVE_GOODS_CODE() As String
   FIELD_SENSITIVE_GOODS_CODE = mvarFIELD_SENSITIVE_GOODS_CODE
End Property

'  3
Public Property Let FIELD_SENSITIVE_QUANTITY(ByVal vData As String)
   mvarFIELD_SENSITIVE_QUANTITY = vData
End Property

Public Property Get FIELD_SENSITIVE_QUANTITY() As String
   FIELD_SENSITIVE_QUANTITY = mvarFIELD_SENSITIVE_QUANTITY
End Property

'  4
Public Property Let FIELD_ORDINAL_PARENT(ByVal vData As Long)
   mvarFIELD_ORDINAL_PARENT = vData
End Property

Public Property Get FIELD_ORDINAL_PARENT() As Long
   FIELD_ORDINAL_PARENT = mvarFIELD_ORDINAL_PARENT
End Property

'  5
Public Property Let FIELD_ORDINAL(ByVal vData As Long)
   mvarFIELD_ORDINAL = vData
End Property

Public Property Get FIELD_ORDINAL() As Long
   FIELD_ORDINAL = mvarFIELD_ORDINAL
End Property


