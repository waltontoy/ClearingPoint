VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEDI_TMS_SEG_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_EDI_TMS_SEG_ID As Long ' 1
Private mvarFIELD_EDI_TMS_SEG_Tag As String ' 2
Private mvarFIELD_EDI_TMS_SEG_Description As String ' 3

'  1
Public Property Let FIELD_EDI_TMS_SEG_ID(ByVal vData As Long)
   mvarFIELD_EDI_TMS_SEG_ID = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_ID() As Long
   FIELD_EDI_TMS_SEG_ID = mvarFIELD_EDI_TMS_SEG_ID
End Property

'  2
Public Property Let FIELD_EDI_TMS_SEG_Tag(ByVal vData As String)
   mvarFIELD_EDI_TMS_SEG_Tag = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_Tag() As String
   FIELD_EDI_TMS_SEG_Tag = mvarFIELD_EDI_TMS_SEG_Tag
End Property

'  3
Public Property Let FIELD_EDI_TMS_SEG_Description(ByVal vData As String)
   mvarFIELD_EDI_TMS_SEG_Description = vData
End Property

Public Property Get FIELD_EDI_TMS_SEG_Description() As String
   FIELD_EDI_TMS_SEG_Description = mvarFIELD_EDI_TMS_SEG_Description
End Property


