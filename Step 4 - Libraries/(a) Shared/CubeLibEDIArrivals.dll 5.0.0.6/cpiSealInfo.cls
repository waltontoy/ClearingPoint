VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSealInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

Private mvarSeals As cpiSealItems '

Private mvarCODE_FIELD As String ' 1
Private mvarNEW_SEALS_NO As String ' 2  AK
Private mvarORDINAL_PARENT As Long ' 3
Private mvarORDINAL As Long ' 4
Private mvarSEAL_START As String ' 5  AL
Private mvarSEAL_END As String ' 6  AM

' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

Public Property Set Seals(ByVal vData As cpiSealItems)
    Set mvarSeals = vData
End Property

Public Property Get Seals() As cpiSealItems
    Set Seals = mvarSeals
End Property

' 2
Public Property Let NEW_SEALS_NO(ByVal vData As String)
    mvarNEW_SEALS_NO = vData
End Property

Public Property Get NEW_SEALS_NO() As String
    NEW_SEALS_NO = mvarNEW_SEALS_NO
End Property

' 3
Public Property Let ORDINAL_PARENT(ByVal vData As Long)
    mvarORDINAL_PARENT = vData
End Property

Public Property Get ORDINAL_PARENT() As Long
    ORDINAL_PARENT = mvarORDINAL_PARENT
End Property

' 4
Public Property Let ORDINAL(ByVal vData As Long)
    mvarORDINAL = vData
End Property

Public Property Get ORDINAL() As Long
    ORDINAL = mvarORDINAL
End Property

' 5
Public Property Let SEAL_START(ByVal vData As String)
    mvarSEAL_START = vData
End Property

Public Property Get SEAL_START() As String
    SEAL_START = mvarSEAL_START
End Property

' 6
Public Property Let SEAL_END(ByVal vData As String)
    mvarSEAL_END = vData
End Property

Public Property Get SEAL_END() As String
    SEAL_END = mvarSEAL_END
End Property

