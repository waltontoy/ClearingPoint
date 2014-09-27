VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiPICK_MAINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_INTERNAL_CODE As String ' 1
Private mvarFIELD_CODE As String ' 2
Private mvarFIELD_DESCRIPTION As String ' 3

'  1
Public Property Let FIELD_INTERNAL_CODE(ByVal vData As String)
   mvarFIELD_INTERNAL_CODE = vData
End Property

Public Property Get FIELD_INTERNAL_CODE() As String
   FIELD_INTERNAL_CODE = mvarFIELD_INTERNAL_CODE
End Property

'  2
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  3
Public Property Let FIELD_DESCRIPTION(ByVal vData As String)
   mvarFIELD_DESCRIPTION = vData
End Property

Public Property Get FIELD_DESCRIPTION() As String
   FIELD_DESCRIPTION = mvarFIELD_DESCRIPTION
End Property


