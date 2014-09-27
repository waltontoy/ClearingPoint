VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiPackage"
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
Private mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES As String ' 2
Private mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG As String ' 3
Private mvarFIELD_KIND_OF_PACKAGES As String ' 4
Private mvarFIELD_NUMBER_OF_PACKAGES As String ' 5
Private mvarFIELD_NUMBER_OF_PIECES As String ' 6
Private mvarFIELD_ORDINAL_PARENT As Long ' 7
Private mvarFIELD_ORDINAL As Long  ' 8
Private mvarFIELD_S5 As String

'  1
Public Property Let FIELD_CODE(ByVal vData As String)
   mvarFIELD_CODE = vData
End Property

Public Property Get FIELD_CODE() As String
   FIELD_CODE = mvarFIELD_CODE
End Property

'  2
Public Property Let FIELD_MARKS_AND_NUMBERS_OF_PACKAGES(ByVal vData As String)
   mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES = vData
End Property

Public Property Get FIELD_MARKS_AND_NUMBERS_OF_PACKAGES() As String
   FIELD_MARKS_AND_NUMBERS_OF_PACKAGES = mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES
End Property

'  3
Public Property Let FIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG(ByVal vData As String)
   mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG = vData
End Property

Public Property Get FIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG() As String
   FIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG = mvarFIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG
End Property

'  4
Public Property Let FIELD_KIND_OF_PACKAGES(ByVal vData As String)
   mvarFIELD_KIND_OF_PACKAGES = vData
End Property

Public Property Get FIELD_KIND_OF_PACKAGES() As String
   FIELD_KIND_OF_PACKAGES = mvarFIELD_KIND_OF_PACKAGES
End Property

'  5
Public Property Let FIELD_NUMBER_OF_PACKAGES(ByVal vData As String)
   mvarFIELD_NUMBER_OF_PACKAGES = vData
End Property

Public Property Get FIELD_NUMBER_OF_PACKAGES() As String
   FIELD_NUMBER_OF_PACKAGES = mvarFIELD_NUMBER_OF_PACKAGES
End Property

'  6
Public Property Let FIELD_NUMBER_OF_PIECES(ByVal vData As String)
   mvarFIELD_NUMBER_OF_PIECES = vData
End Property

Public Property Get FIELD_NUMBER_OF_PIECES() As String
   FIELD_NUMBER_OF_PIECES = mvarFIELD_NUMBER_OF_PIECES
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
Public Property Let FIELD_S5(ByVal vData As String)
   mvarFIELD_S5 = vData
End Property

Public Property Get FIELD_S5() As String
   FIELD_S5 = mvarFIELD_S5
End Property


