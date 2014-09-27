VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiNCTS_IEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_NCTS_IEM_ID As Long ' 1
Private mvarFIELD_Done As Boolean ' 2
Private mvarFIELD_EDI_TMS_COR_ID As Long ' 3
Private mvarFIELD_NCTS_IEM_Name As String ' 4
Private mvarFIELD_NCTS_IEM_Code As String ' 5
Private mvarFIELD_NCTS_IEM_Description_E As String ' 6
Private mvarFIELD_NCTS_IEM_Description_D As String ' 7
Private mvarFIELD_NCTS_IEM_Description_F As String ' 8
Private mvarFIELD_NCTS_IEM_Reference As String ' 9

'  1
Public Property Let FIELD_NCTS_IEM_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_ID() As Long
   FIELD_NCTS_IEM_ID = mvarFIELD_NCTS_IEM_ID
End Property

'  2
Public Property Let FIELD_Done(ByVal vData As Boolean)
   mvarFIELD_Done = vData
End Property

Public Property Get FIELD_Done() As Boolean
   FIELD_Done = mvarFIELD_Done
End Property

'  3
Public Property Let FIELD_EDI_TMS_COR_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_COR_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_COR_ID() As Long
   FIELD_EDI_TMS_COR_ID = mvarFIELD_EDI_TMS_COR_ID
End Property

'  4
Public Property Let FIELD_NCTS_IEM_Name(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Name = vData
End Property

Public Property Get FIELD_NCTS_IEM_Name() As String
   FIELD_NCTS_IEM_Name = mvarFIELD_NCTS_IEM_Name
End Property

'  5
Public Property Let FIELD_NCTS_IEM_Code(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Code = vData
End Property

Public Property Get FIELD_NCTS_IEM_Code() As String
   FIELD_NCTS_IEM_Code = mvarFIELD_NCTS_IEM_Code
End Property

'  6
Public Property Let FIELD_NCTS_IEM_Description_E(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Description_E = vData
End Property

Public Property Get FIELD_NCTS_IEM_Description_E() As String
   FIELD_NCTS_IEM_Description_E = mvarFIELD_NCTS_IEM_Description_E
End Property

'  7
Public Property Let FIELD_NCTS_IEM_Description_D(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Description_D = vData
End Property

Public Property Get FIELD_NCTS_IEM_Description_D() As String
   FIELD_NCTS_IEM_Description_D = mvarFIELD_NCTS_IEM_Description_D
End Property

'  8
Public Property Let FIELD_NCTS_IEM_Description_F(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Description_F = vData
End Property

Public Property Get FIELD_NCTS_IEM_Description_F() As String
   FIELD_NCTS_IEM_Description_F = mvarFIELD_NCTS_IEM_Description_F
End Property

'  9
Public Property Let FIELD_NCTS_IEM_Reference(ByVal vData As String)
   mvarFIELD_NCTS_IEM_Reference = vData
End Property

Public Property Get FIELD_NCTS_IEM_Reference() As String
   FIELD_NCTS_IEM_Reference = mvarFIELD_NCTS_IEM_Reference
End Property


