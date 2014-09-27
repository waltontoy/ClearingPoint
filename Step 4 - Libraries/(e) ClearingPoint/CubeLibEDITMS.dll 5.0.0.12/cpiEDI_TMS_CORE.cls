VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiEDI_TMS_CORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_EDI_TMS_COR_ID As Long ' 1
Private mvarFIELD_EDI_TMS_COR_Name As String ' 2

'  1
Public Property Let FIELD_EDI_TMS_COR_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_COR_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_COR_ID() As Long
   FIELD_EDI_TMS_COR_ID = mvarFIELD_EDI_TMS_COR_ID
End Property

'  2
Public Property Let FIELD_EDI_TMS_COR_Name(ByVal vData As String)
   mvarFIELD_EDI_TMS_COR_Name = vData
End Property

Public Property Get FIELD_EDI_TMS_COR_Name() As String
   FIELD_EDI_TMS_COR_Name = mvarFIELD_EDI_TMS_COR_Name
End Property


