VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBOX_DEF_VAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_LOGID_DESCRIPTION As String ' 1
Private mvarFIELD_BOX_CODE As String ' 2
Private mvarFIELD_DEFAULT_VALUE As String ' 3
Private mvarFIELD_user_no As Long ' 4
Private mvarFIELD_TYPE As String ' 5

'  1
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property

Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

'  2
Public Property Let FIELD_BOX_CODE(ByVal vData As String)
   mvarFIELD_BOX_CODE = vData
End Property

Public Property Get FIELD_BOX_CODE() As String
   FIELD_BOX_CODE = mvarFIELD_BOX_CODE
End Property

'  3
Public Property Let FIELD_DEFAULT_VALUE(ByVal vData As String)
   mvarFIELD_DEFAULT_VALUE = vData
End Property

Public Property Get FIELD_DEFAULT_VALUE() As String
   FIELD_DEFAULT_VALUE = mvarFIELD_DEFAULT_VALUE
End Property

'  4
Public Property Let FIELD_user_no(ByVal vData As Long)
   mvarFIELD_user_no = vData
End Property

Public Property Get FIELD_user_no() As Long
   FIELD_user_no = mvarFIELD_user_no
End Property

'  5
Public Property Let FIELD_TYPE(ByVal vData As String)
   mvarFIELD_TYPE = vData
End Property

Public Property Get FIELD_TYPE() As String
   FIELD_TYPE = mvarFIELD_TYPE
End Property


