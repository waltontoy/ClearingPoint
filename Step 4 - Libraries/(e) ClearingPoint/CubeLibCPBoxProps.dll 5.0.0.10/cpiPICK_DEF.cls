VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiPICK_DEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_PICKLIST_DESCRIPTION_ENGLISH As String ' 1
Private mvarFIELD_PICKLIST_DESCRIPTION_DUTCH As String ' 2
Private mvarFIELD_PICKLIST_DESCRIPTION_FRENCH As String ' 3
Private mvarFIELD_DOCUMENT As String ' 4
Private mvarFIELD_BOX_CODE As String ' 5
Private mvarFIELD_FROM As Long ' 6
Private mvarFIELD_TO As Long ' 7
Private mvarFIELD_VALIDS As String ' 8
Private mvarFIELD_WIDTH As Long ' 9
Private mvarFIELD_AUTO_ADD As Boolean ' 10
Private mvarFIELD_DATA_TYPE As String ' 11
Private mvarFIELD_INTERNAL_CODE As String ' 12

'  1
Public Property Let FIELD_PICKLIST_DESCRIPTION_ENGLISH(ByVal vData As String)
   mvarFIELD_PICKLIST_DESCRIPTION_ENGLISH = vData
End Property

Public Property Get FIELD_PICKLIST_DESCRIPTION_ENGLISH() As String
   FIELD_PICKLIST_DESCRIPTION_ENGLISH = mvarFIELD_PICKLIST_DESCRIPTION_ENGLISH
End Property

'  2
Public Property Let FIELD_PICKLIST_DESCRIPTION_DUTCH(ByVal vData As String)
   mvarFIELD_PICKLIST_DESCRIPTION_DUTCH = vData
End Property

Public Property Get FIELD_PICKLIST_DESCRIPTION_DUTCH() As String
   FIELD_PICKLIST_DESCRIPTION_DUTCH = mvarFIELD_PICKLIST_DESCRIPTION_DUTCH
End Property

'  3
Public Property Let FIELD_PICKLIST_DESCRIPTION_FRENCH(ByVal vData As String)
   mvarFIELD_PICKLIST_DESCRIPTION_FRENCH = vData
End Property

Public Property Get FIELD_PICKLIST_DESCRIPTION_FRENCH() As String
   FIELD_PICKLIST_DESCRIPTION_FRENCH = mvarFIELD_PICKLIST_DESCRIPTION_FRENCH
End Property

'  4
Public Property Let FIELD_DOCUMENT(ByVal vData As String)
   mvarFIELD_DOCUMENT = vData
End Property

Public Property Get FIELD_DOCUMENT() As String
   FIELD_DOCUMENT = mvarFIELD_DOCUMENT
End Property

'  5
Public Property Let FIELD_BOX_CODE(ByVal vData As String)
   mvarFIELD_BOX_CODE = vData
End Property

Public Property Get FIELD_BOX_CODE() As String
   FIELD_BOX_CODE = mvarFIELD_BOX_CODE
End Property

'  6
Public Property Let FIELD_FROM(ByVal vData As Long)
   mvarFIELD_FROM = vData
End Property

Public Property Get FIELD_FROM() As Long
   FIELD_FROM = mvarFIELD_FROM
End Property

'  7
Public Property Let FIELD_TO(ByVal vData As Long)
   mvarFIELD_TO = vData
End Property

Public Property Get FIELD_TO() As Long
   FIELD_TO = mvarFIELD_TO
End Property

'  8
Public Property Let FIELD_VALIDS(ByVal vData As String)
   mvarFIELD_VALIDS = vData
End Property

Public Property Get FIELD_VALIDS() As String
   FIELD_VALIDS = mvarFIELD_VALIDS
End Property

'  9
Public Property Let FIELD_WIDTH(ByVal vData As Long)
   mvarFIELD_WIDTH = vData
End Property

Public Property Get FIELD_WIDTH() As Long
   FIELD_WIDTH = mvarFIELD_WIDTH
End Property

'  10
Public Property Let FIELD_AUTO_ADD(ByVal vData As Boolean)
   mvarFIELD_AUTO_ADD = vData
End Property

Public Property Get FIELD_AUTO_ADD() As Boolean
   FIELD_AUTO_ADD = mvarFIELD_AUTO_ADD
End Property

'  11
Public Property Let FIELD_DATA_TYPE(ByVal vData As String)
   mvarFIELD_DATA_TYPE = vData
End Property

Public Property Get FIELD_DATA_TYPE() As String
   FIELD_DATA_TYPE = mvarFIELD_DATA_TYPE
End Property

'  12
Public Property Let FIELD_INTERNAL_CODE(ByVal vData As String)
   mvarFIELD_INTERNAL_CODE = vData
End Property

Public Property Get FIELD_INTERNAL_CODE() As String
   FIELD_INTERNAL_CODE = mvarFIELD_INTERNAL_CODE
End Property


