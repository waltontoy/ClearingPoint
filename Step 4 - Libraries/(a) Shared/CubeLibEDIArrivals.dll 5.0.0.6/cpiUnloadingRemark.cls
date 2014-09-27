VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiUnloadingRemark"
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
Private mvarFIELD_UNLOADING_DATE As String ' 2
Private mvarFIELD_UNLOADING_COMPLETION As String ' 3
Private mvarFIELD_STATE_OF_SEALS_OK As String ' 4
Private mvarFIELD_CONFORM As String ' 5
Private mvarFIELD_UNLOADING_REMARK As String ' 6
Private mvarFIELD_UNLOADING_REMARK_LNG As String ' 7
Private mvarFIELD_ORDINAL_PARENT As Long ' 8
Private mvarFIELD_ORDINAL As Long ' 9

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_UNLOADING_DATE(ByVal vData As String)
   mvarFIELD_UNLOADING_DATE = vData
End Property

Public Property Get FIELD_UNLOADING_DATE() As String
   FIELD_UNLOADING_DATE = mvarFIELD_UNLOADING_DATE
End Property

'  3
Public Property Let FIELD_UNLOADING_COMPLETION(ByVal vData As String)
   mvarFIELD_UNLOADING_COMPLETION = vData
End Property

Public Property Get FIELD_UNLOADING_COMPLETION() As String
   FIELD_UNLOADING_COMPLETION = mvarFIELD_UNLOADING_COMPLETION
End Property

'  4
Public Property Let FIELD_STATE_OF_SEALS_OK(ByVal vData As String)
   mvarFIELD_STATE_OF_SEALS_OK = vData
End Property

Public Property Get FIELD_STATE_OF_SEALS_OK() As String
   FIELD_STATE_OF_SEALS_OK = mvarFIELD_STATE_OF_SEALS_OK
End Property

'  5
Public Property Let FIELD_CONFORM(ByVal vData As String)
   mvarFIELD_CONFORM = vData
End Property

Public Property Get FIELD_CONFORM() As String
   FIELD_CONFORM = mvarFIELD_CONFORM
End Property

'  6
Public Property Let FIELD_UNLOADING_REMARK(ByVal vData As String)
   mvarFIELD_UNLOADING_REMARK = vData
End Property

Public Property Get FIELD_UNLOADING_REMARK() As String
   FIELD_UNLOADING_REMARK = mvarFIELD_UNLOADING_REMARK
End Property

'  7
Public Property Let FIELD_UNLOADING_REMARK_LNG(ByVal vData As String)
   mvarFIELD_UNLOADING_REMARK_LNG = vData
End Property

Public Property Get FIELD_UNLOADING_REMARK_LNG() As String
   FIELD_UNLOADING_REMARK_LNG = mvarFIELD_UNLOADING_REMARK_LNG
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


