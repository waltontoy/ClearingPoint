VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiDocCertificate"
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
Private mvarFIELD_DOCUMENT_TYPE As String ' 2
Private mvarFIELD_DOCUMENT_REFERENCE As String ' 3
Private mvarFIELD_DOCUMENT_REFERENCE_LNG As String ' 4
Private mvarFIELD_COMPLEMENT_INFORMATION As String ' 5
Private mvarFIELD_COMPLEMENT_INFORMATION_LNG As String ' 6
Private mvarFIELD_ORDINAL_PARENT As Long ' 7
Private mvarFIELD_ORDINAL As Long ' 8

Private mvarFIELD_Y1 As String
Private mvarFIELD_Y5 As String

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_DOCUMENT_TYPE(ByVal vData As String)
   mvarFIELD_DOCUMENT_TYPE = vData
End Property

Public Property Get FIELD_DOCUMENT_TYPE() As String
   FIELD_DOCUMENT_TYPE = mvarFIELD_DOCUMENT_TYPE
End Property

'  3
Public Property Let FIELD_DOCUMENT_REFERENCE(ByVal vData As String)
   mvarFIELD_DOCUMENT_REFERENCE = vData
End Property

Public Property Get FIELD_DOCUMENT_REFERENCE() As String
   FIELD_DOCUMENT_REFERENCE = mvarFIELD_DOCUMENT_REFERENCE
End Property

'  4
Public Property Let FIELD_DOCUMENT_REFERENCE_LNG(ByVal vData As String)
   mvarFIELD_DOCUMENT_REFERENCE_LNG = vData
End Property

Public Property Get FIELD_DOCUMENT_REFERENCE_LNG() As String
   FIELD_DOCUMENT_REFERENCE_LNG = mvarFIELD_DOCUMENT_REFERENCE_LNG
End Property

'  5
Public Property Let FIELD_COMPLEMENT_INFORMATION(ByVal vData As String)
   mvarFIELD_COMPLEMENT_INFORMATION = vData
End Property

Public Property Get FIELD_COMPLEMENT_INFORMATION() As String
   FIELD_COMPLEMENT_INFORMATION = mvarFIELD_COMPLEMENT_INFORMATION
End Property

'  6
Public Property Let FIELD_COMPLEMENT_INFORMATION_LNG(ByVal vData As String)
   mvarFIELD_COMPLEMENT_INFORMATION_LNG = vData
End Property

Public Property Get FIELD_COMPLEMENT_INFORMATION_LNG() As String
   FIELD_COMPLEMENT_INFORMATION_LNG = mvarFIELD_COMPLEMENT_INFORMATION_LNG
End Property

'  7
Public Property Let FIELD_ORDINAL_PARENT(ByVal vData As Long)
   mvarFIELD_ORDINAL_PARENT = vData
End Property

Public Property Get FIELD_ORDINAL_PARENT() As Long
   FIELD_ORDINAL_PARENT = mvarFIELD_ORDINAL_PARENT
End Property

'  8
Public Property Let FIELD_ORDINAL(ByVal vData As Long)
   mvarFIELD_ORDINAL = vData
End Property

Public Property Get FIELD_ORDINAL() As Long
   FIELD_ORDINAL = mvarFIELD_ORDINAL
End Property


Public Property Let FIELD_Y1(ByVal vData As String)
    mvarFIELD_Y1 = vData
End Property
Public Property Get FIELD_Y1() As String
    FIELD_Y1 = mvarFIELD_Y1
End Property
Public Property Let FIELD_Y5(ByVal vData As String)
    mvarFIELD_Y5 = vData
End Property
Public Property Get FIELD_Y5() As String
    FIELD_Y5 = mvarFIELD_Y5
End Property

Private Sub Class_Initialize()
    mvarFIELD_Y5 = ""
    mvarFIELD_Y1 = ""
End Sub
