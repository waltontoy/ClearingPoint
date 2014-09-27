VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiOldSeals_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarSeals  As cpiSealItems

Private mvarFIELD_CODE As String ' 1
Private mvarFIELD_SEALS_NUMBER As String ' 2
Private mvarFIELD_SEAL_START As String ' 3
Private mvarFIELD_SEAL_END As String ' 4
Private mvarFIELD_ORDINAL_PARENT As Long ' 5
Private mvarFIELD_ORDINAL As Long ' 6

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_SEALS_NUMBER(ByVal vData As String)
   mvarFIELD_SEALS_NUMBER = vData
End Property

Public Property Get FIELD_SEALS_NUMBER() As String
   FIELD_SEALS_NUMBER = mvarFIELD_SEALS_NUMBER
End Property

'  3
Public Property Let FIELD_SEAL_START(ByVal vData As String)
   mvarFIELD_SEAL_START = vData
End Property

Public Property Get FIELD_SEAL_START() As String
   FIELD_SEAL_START = mvarFIELD_SEAL_START
End Property

'  4
Public Property Let FIELD_SEAL_END(ByVal vData As String)
   mvarFIELD_SEAL_END = vData
End Property

Public Property Get FIELD_SEAL_END() As String
   FIELD_SEAL_END = mvarFIELD_SEAL_END
End Property

'  5
Public Property Let FIELD_ORDINAL_PARENT(ByVal vData As Long)
   mvarFIELD_ORDINAL_PARENT = vData
End Property

Public Property Get FIELD_ORDINAL_PARENT() As Long
   FIELD_ORDINAL_PARENT = mvarFIELD_ORDINAL_PARENT
End Property

'  6
Public Property Let FIELD_ORDINAL(ByVal vData As Long)
   mvarFIELD_ORDINAL = vData
End Property

Public Property Get FIELD_ORDINAL() As Long
   FIELD_ORDINAL = mvarFIELD_ORDINAL
End Property

' 7
Public Property Set Seals(ByVal vData As cpiSealItems)
    Set mvarSeals = vData
End Property

Public Property Get Seals() As cpiSealItems
    Set Seals = mvarSeals
End Property

