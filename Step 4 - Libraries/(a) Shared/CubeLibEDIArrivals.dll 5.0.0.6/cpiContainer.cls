VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

Private mvarCODE_FIELD As String ' 1
Private mvarNEW_CONTAINER_NUMBER As String ' SC 2
Private mvarORDINAL_PARENT As Long ' 3
Private mvarORDINAL As Long ' 4
Private mvarFIELD_SB As String

' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 2
Public Property Let NEW_CONTAINER_NUMBER(ByVal vData As String)
    mvarNEW_CONTAINER_NUMBER = vData
End Property

Public Property Get NEW_CONTAINER_NUMBER() As String
    NEW_CONTAINER_NUMBER = mvarNEW_CONTAINER_NUMBER
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


Public Property Let FIELD_SB(ByVal vData As String)
    mvarFIELD_SB = vData
End Property

Public Property Get FIELD_SB() As String
    FIELD_SB = mvarFIELD_SB
End Property

