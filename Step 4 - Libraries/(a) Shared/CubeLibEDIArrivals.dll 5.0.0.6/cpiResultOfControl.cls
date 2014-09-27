VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiResultOfControl"
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
Private mvarFIELD_CONTROL_INDICATOR As String ' 2
Private mvarFIELD_DESCRIPTION As String ' 3
Private mvarFIELD_DESCRIPTION_LNG As String ' 4
Private mvarFIELD_POINTER_TO_THE_ATTRIBUTE As String ' 5
Private mvarFIELD_CORRECTED_VALUE As String ' 6
Private mvarFIELD_ORDINAL_PARENT As Long ' 7
Private mvarFIELD_ORDINAL As Long ' 8

Private mvarFIELD_CL As String
Private mvarFIELD_CM As String

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_CONTROL_INDICATOR(ByVal vData As String)
   mvarFIELD_CONTROL_INDICATOR = vData
End Property

Public Property Get FIELD_CONTROL_INDICATOR() As String
   FIELD_CONTROL_INDICATOR = mvarFIELD_CONTROL_INDICATOR
End Property

'  3
Public Property Let FIELD_DESCRIPTION(ByVal vData As String)
   mvarFIELD_DESCRIPTION = vData
End Property

Public Property Get FIELD_DESCRIPTION() As String
   FIELD_DESCRIPTION = mvarFIELD_DESCRIPTION
End Property

'  4
Public Property Let FIELD_DESCRIPTION_LNG(ByVal vData As String)
   mvarFIELD_DESCRIPTION_LNG = vData
End Property

Public Property Get FIELD_DESCRIPTION_LNG() As String
   FIELD_DESCRIPTION_LNG = mvarFIELD_DESCRIPTION_LNG
End Property

'  5
Public Property Let FIELD_POINTER_TO_THE_ATTRIBUTE(ByVal vData As String)
   mvarFIELD_POINTER_TO_THE_ATTRIBUTE = vData
End Property

Public Property Get FIELD_POINTER_TO_THE_ATTRIBUTE() As String
   FIELD_POINTER_TO_THE_ATTRIBUTE = mvarFIELD_POINTER_TO_THE_ATTRIBUTE
End Property

'  6
Public Property Let FIELD_CORRECTED_VALUE(ByVal vData As String)
   mvarFIELD_CORRECTED_VALUE = vData
End Property

Public Property Get FIELD_CORRECTED_VALUE() As String
   FIELD_CORRECTED_VALUE = mvarFIELD_CORRECTED_VALUE
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


Public Property Let FIELD_CL(ByVal vData As String)
    mvarFIELD_CL = vData
End Property
Public Property Get FIELD_CL() As String
    FIELD_CL = mvarFIELD_CL
End Property

Public Property Let FIELD_CM(ByVal vData As String)
    mvarFIELD_CM = vData
End Property
Public Property Get FIELD_CM() As String
    FIELD_CM = mvarFIELD_CM
End Property

