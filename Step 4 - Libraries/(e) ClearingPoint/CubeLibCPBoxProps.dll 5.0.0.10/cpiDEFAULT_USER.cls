VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiDEFAULT_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_USER_NO As Long ' 1
Private mvarFIELD_BOX_CODE As String ' 2
Private mvarFIELD_DEFAULT_VALUE As String ' 3
Private mvarFIELD_SEQUENTIAL_TABBING_DEFAULT As Boolean ' 4
Private mvarFIELD_SEQUENTIAL_TABBING_ACTIVE As Boolean ' 5
Private mvarFIELD_LOGID_DESCRIPTION As String ' 6
Private mvarFIELD_TYPE As String ' 7

'  1
Public Property Let FIELD_USER_NO(ByVal vData As Long)
   mvarFIELD_USER_NO = vData
End Property

Public Property Get FIELD_USER_NO() As Long
   FIELD_USER_NO = mvarFIELD_USER_NO
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
Public Property Let FIELD_SEQUENTIAL_TABBING_DEFAULT(ByVal vData As Boolean)
   mvarFIELD_SEQUENTIAL_TABBING_DEFAULT = vData
End Property

Public Property Get FIELD_SEQUENTIAL_TABBING_DEFAULT() As Boolean
   FIELD_SEQUENTIAL_TABBING_DEFAULT = mvarFIELD_SEQUENTIAL_TABBING_DEFAULT
End Property

'  5
Public Property Let FIELD_SEQUENTIAL_TABBING_ACTIVE(ByVal vData As Boolean)
   mvarFIELD_SEQUENTIAL_TABBING_ACTIVE = vData
End Property

Public Property Get FIELD_SEQUENTIAL_TABBING_ACTIVE() As Boolean
   FIELD_SEQUENTIAL_TABBING_ACTIVE = mvarFIELD_SEQUENTIAL_TABBING_ACTIVE
End Property

'  6
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property

Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

'  7
Public Property Let FIELD_TYPE(ByVal vData As String)
   mvarFIELD_TYPE = vData
End Property

Public Property Get FIELD_TYPE() As String
   FIELD_TYPE = mvarFIELD_TYPE
End Property


