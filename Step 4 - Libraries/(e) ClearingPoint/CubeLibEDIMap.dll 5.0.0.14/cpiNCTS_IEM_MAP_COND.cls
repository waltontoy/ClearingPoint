VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiNCTS_IEM_MAP_COND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_IEM_MCN_ID As Long ' 1
Private mvarFIELD_NCTS_IEM_MCN_Source As String ' 2
Private mvarFIELD_NCTS_IEM_MCN_Condition As String ' 3

'  1
Public Property Let FIELD_NCTS_IEM_MCN_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_MCN_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_MCN_ID() As Long
   FIELD_NCTS_IEM_MCN_ID = mvarFIELD_NCTS_IEM_MCN_ID
End Property

'  2
Public Property Let FIELD_NCTS_IEM_MCN_Source(ByVal vData As String)
   mvarFIELD_NCTS_IEM_MCN_Source = vData
End Property

Public Property Get FIELD_NCTS_IEM_MCN_Source() As String
   FIELD_NCTS_IEM_MCN_Source = mvarFIELD_NCTS_IEM_MCN_Source
End Property

'  3
Public Property Let FIELD_NCTS_IEM_MCN_Condition(ByVal vData As String)
   mvarFIELD_NCTS_IEM_MCN_Condition = vData
End Property

Public Property Get FIELD_NCTS_IEM_MCN_Condition() As String
   FIELD_NCTS_IEM_MCN_Condition = mvarFIELD_NCTS_IEM_MCN_Condition
End Property


