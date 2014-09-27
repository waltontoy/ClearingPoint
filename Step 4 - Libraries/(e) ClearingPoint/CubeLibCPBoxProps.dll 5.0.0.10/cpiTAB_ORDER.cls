VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiTAB_ORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_BOX_CODE As String ' 1
Private mvarFIELD_VALUE As String ' 2
Private mvarFIELD_EMPTY As Boolean ' 3
Private mvarFIELD_REFERENCE As String ' 4
Private mvarFIELD_USER_NO As Long ' 5
Private mvarFIELD_TYPE As String ' 6

'  1
Public Property Let FIELD_BOX_CODE(ByVal vData As String)
   mvarFIELD_BOX_CODE = vData
End Property

Public Property Get FIELD_BOX_CODE() As String
   FIELD_BOX_CODE = mvarFIELD_BOX_CODE
End Property

'  2
Public Property Let FIELD_VALUE(ByVal vData As String)
   mvarFIELD_VALUE = vData
End Property

Public Property Get FIELD_VALUE() As String
   FIELD_VALUE = mvarFIELD_VALUE
End Property

'  3
Public Property Let FIELD_EMPTY(ByVal vData As Boolean)
   mvarFIELD_EMPTY = vData
End Property

Public Property Get FIELD_EMPTY() As Boolean
   FIELD_EMPTY = mvarFIELD_EMPTY
End Property

'  4
Public Property Let FIELD_REFERENCE(ByVal vData As String)
   mvarFIELD_REFERENCE = vData
End Property

Public Property Get FIELD_REFERENCE() As String
   FIELD_REFERENCE = mvarFIELD_REFERENCE
End Property

'  5
Public Property Let FIELD_USER_NO(ByVal vData As Long)
   mvarFIELD_USER_NO = vData
End Property

Public Property Get FIELD_USER_NO() As Long
   FIELD_USER_NO = mvarFIELD_USER_NO
End Property

'  6
Public Property Let FIELD_TYPE(ByVal vData As String)
   mvarFIELD_TYPE = vData
End Property

Public Property Get FIELD_TYPE() As String
   FIELD_TYPE = mvarFIELD_TYPE
End Property


