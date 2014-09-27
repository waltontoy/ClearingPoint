VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBoxDefEdIIE44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_BOX_CODE As String ' 1
Private mvarFIELD_ENGLISH_DESCRIPTION As String ' 2
Private mvarFIELD_DUTCH_DESCRIPTION As String ' 3
Private mvarFIELD_FRENCH_DESCRIPTION As String ' 4
Private mvarFIELD_EMPTY_FIELD_VALUE As String ' 5
Private mvarFIELD_INSERT As Integer ' 6
Private mvarFIELD_JUSTIFY As String ' 7
Private mvarFIELD_SEQUENTIAL_TABBING_DEFAULT As Boolean ' 8
Private mvarFIELD_SEQUENTIAL_TABBING_ACTIVE As Boolean ' 9
Private mvarFIELD_CHECK_VAT As Boolean ' 10
Private mvarFIELD_CALCULATE_NET_WEIGHT As Boolean ' 11
Private mvarFIELD_CALCULATE_NO_OF_ITEMS As Boolean ' 12
Private mvarFIELD_COPY_TO_NEXT As Boolean ' 13
Private mvarFIELD_CHANGE_ONLY_IN_H1 As Boolean ' 14
Private mvarFIELD_SEND_ONLY_IN_H1 As Boolean ' 15
Private mvarFIELD_WIDTH As Integer ' 16
Private mvarFIELD_DECIMAL As Integer ' 17
Private mvarFIELD_DATA_TYPE As String ' 18
Private mvarFIELD_AUTO_ADD As Boolean ' 19
Private mvarFIELD_DEFAULT_VALUE As String ' 20
Private mvarFIELD_CALCULATE_CUSTOMS_VALUE As Boolean ' 21
Private mvarFIELD_RELATE_L1_TO_S1 As Boolean ' 22
Private mvarFIELD_BOX_TYPE As String ' 23
Private mvarFIELD_BOX_INDEX As Integer ' 24
Private mvarFIELD_INTERNAL_CODE As String ' 25

'  1
Public Property Let FIELD_BOX_CODE(ByVal vData As String)
   mvarFIELD_BOX_CODE = vData
End Property

Public Property Get FIELD_BOX_CODE() As String
   FIELD_BOX_CODE = mvarFIELD_BOX_CODE
End Property

'  2
Public Property Let FIELD_ENGLISH_DESCRIPTION(ByVal vData As String)
   mvarFIELD_ENGLISH_DESCRIPTION = vData
End Property

Public Property Get FIELD_ENGLISH_DESCRIPTION() As String
   FIELD_ENGLISH_DESCRIPTION = mvarFIELD_ENGLISH_DESCRIPTION
End Property

'  3
Public Property Let FIELD_DUTCH_DESCRIPTION(ByVal vData As String)
   mvarFIELD_DUTCH_DESCRIPTION = vData
End Property

Public Property Get FIELD_DUTCH_DESCRIPTION() As String
   FIELD_DUTCH_DESCRIPTION = mvarFIELD_DUTCH_DESCRIPTION
End Property

'  4
Public Property Let FIELD_FRENCH_DESCRIPTION(ByVal vData As String)
   mvarFIELD_FRENCH_DESCRIPTION = vData
End Property

Public Property Get FIELD_FRENCH_DESCRIPTION() As String
   FIELD_FRENCH_DESCRIPTION = mvarFIELD_FRENCH_DESCRIPTION
End Property

'  5
Public Property Let FIELD_EMPTY_FIELD_VALUE(ByVal vData As String)
   mvarFIELD_EMPTY_FIELD_VALUE = vData
End Property

Public Property Get FIELD_EMPTY_FIELD_VALUE() As String
   FIELD_EMPTY_FIELD_VALUE = mvarFIELD_EMPTY_FIELD_VALUE
End Property

'  6
Public Property Let FIELD_INSERT(ByVal vData As Integer)
   mvarFIELD_INSERT = vData
End Property

Public Property Get FIELD_INSERT() As Integer
   FIELD_INSERT = mvarFIELD_INSERT
End Property

'  7
Public Property Let FIELD_JUSTIFY(ByVal vData As String)
   mvarFIELD_JUSTIFY = vData
End Property

Public Property Get FIELD_JUSTIFY() As String
   FIELD_JUSTIFY = mvarFIELD_JUSTIFY
End Property

'  8
Public Property Let FIELD_SEQUENTIAL_TABBING_DEFAULT(ByVal vData As Boolean)
   mvarFIELD_SEQUENTIAL_TABBING_DEFAULT = vData
End Property

Public Property Get FIELD_SEQUENTIAL_TABBING_DEFAULT() As Boolean
   FIELD_SEQUENTIAL_TABBING_DEFAULT = mvarFIELD_SEQUENTIAL_TABBING_DEFAULT
End Property

'  9
Public Property Let FIELD_SEQUENTIAL_TABBING_ACTIVE(ByVal vData As Boolean)
   mvarFIELD_SEQUENTIAL_TABBING_ACTIVE = vData
End Property

Public Property Get FIELD_SEQUENTIAL_TABBING_ACTIVE() As Boolean
   FIELD_SEQUENTIAL_TABBING_ACTIVE = mvarFIELD_SEQUENTIAL_TABBING_ACTIVE
End Property

'  10
Public Property Let FIELD_CHECK_VAT(ByVal vData As Boolean)
   mvarFIELD_CHECK_VAT = vData
End Property

Public Property Get FIELD_CHECK_VAT() As Boolean
   FIELD_CHECK_VAT = mvarFIELD_CHECK_VAT
End Property

'  11
Public Property Let FIELD_CALCULATE_NET_WEIGHT(ByVal vData As Boolean)
   mvarFIELD_CALCULATE_NET_WEIGHT = vData
End Property

Public Property Get FIELD_CALCULATE_NET_WEIGHT() As Boolean
   FIELD_CALCULATE_NET_WEIGHT = mvarFIELD_CALCULATE_NET_WEIGHT
End Property

'  12
Public Property Let FIELD_CALCULATE_NO_OF_ITEMS(ByVal vData As Boolean)
   mvarFIELD_CALCULATE_NO_OF_ITEMS = vData
End Property

Public Property Get FIELD_CALCULATE_NO_OF_ITEMS() As Boolean
   FIELD_CALCULATE_NO_OF_ITEMS = mvarFIELD_CALCULATE_NO_OF_ITEMS
End Property

'  13
Public Property Let FIELD_COPY_TO_NEXT(ByVal vData As Boolean)
   mvarFIELD_COPY_TO_NEXT = vData
End Property

Public Property Get FIELD_COPY_TO_NEXT() As Boolean
   FIELD_COPY_TO_NEXT = mvarFIELD_COPY_TO_NEXT
End Property

'  14
Public Property Let FIELD_CHANGE_ONLY_IN_H1(ByVal vData As Boolean)
   mvarFIELD_CHANGE_ONLY_IN_H1 = vData
End Property

Public Property Get FIELD_CHANGE_ONLY_IN_H1() As Boolean
   FIELD_CHANGE_ONLY_IN_H1 = mvarFIELD_CHANGE_ONLY_IN_H1
End Property

'  15
Public Property Let FIELD_SEND_ONLY_IN_H1(ByVal vData As Boolean)
   mvarFIELD_SEND_ONLY_IN_H1 = vData
End Property

Public Property Get FIELD_SEND_ONLY_IN_H1() As Boolean
   FIELD_SEND_ONLY_IN_H1 = mvarFIELD_SEND_ONLY_IN_H1
End Property

'  16
Public Property Let FIELD_WIDTH(ByVal vData As Integer)
   mvarFIELD_WIDTH = vData
End Property

Public Property Get FIELD_WIDTH() As Integer
   FIELD_WIDTH = mvarFIELD_WIDTH
End Property

'  17
Public Property Let FIELD_DECIMAL(ByVal vData As Integer)
   mvarFIELD_DECIMAL = vData
End Property

Public Property Get FIELD_DECIMAL() As Integer
   FIELD_DECIMAL = mvarFIELD_DECIMAL
End Property

'  18
Public Property Let FIELD_DATA_TYPE(ByVal vData As String)
   mvarFIELD_DATA_TYPE = vData
End Property

Public Property Get FIELD_DATA_TYPE() As String
   FIELD_DATA_TYPE = mvarFIELD_DATA_TYPE
End Property

'  19
Public Property Let FIELD_AUTO_ADD(ByVal vData As Boolean)
   mvarFIELD_AUTO_ADD = vData
End Property

Public Property Get FIELD_AUTO_ADD() As Boolean
   FIELD_AUTO_ADD = mvarFIELD_AUTO_ADD
End Property

'  20
Public Property Let FIELD_DEFAULT_VALUE(ByVal vData As String)
   mvarFIELD_DEFAULT_VALUE = vData
End Property

Public Property Get FIELD_DEFAULT_VALUE() As String
   FIELD_DEFAULT_VALUE = mvarFIELD_DEFAULT_VALUE
End Property

'  21
Public Property Let FIELD_CALCULATE_CUSTOMS_VALUE(ByVal vData As Boolean)
   mvarFIELD_CALCULATE_CUSTOMS_VALUE = vData
End Property

Public Property Get FIELD_CALCULATE_CUSTOMS_VALUE() As Boolean
   FIELD_CALCULATE_CUSTOMS_VALUE = mvarFIELD_CALCULATE_CUSTOMS_VALUE
End Property

'  22
Public Property Let FIELD_RELATE_L1_TO_S1(ByVal vData As Boolean)
   mvarFIELD_RELATE_L1_TO_S1 = vData
End Property

Public Property Get FIELD_RELATE_L1_TO_S1() As Boolean
   FIELD_RELATE_L1_TO_S1 = mvarFIELD_RELATE_L1_TO_S1
End Property

'  23
Public Property Let FIELD_BOX_TYPE(ByVal vData As String)
   mvarFIELD_BOX_TYPE = vData
End Property

Public Property Get FIELD_BOX_TYPE() As String
   FIELD_BOX_TYPE = mvarFIELD_BOX_TYPE
End Property

'  24
Public Property Let FIELD_BOX_INDEX(ByVal vData As Integer)
   mvarFIELD_BOX_INDEX = vData
End Property

Public Property Get FIELD_BOX_INDEX() As Integer
   FIELD_BOX_INDEX = mvarFIELD_BOX_INDEX
End Property

'  25
Public Property Let FIELD_INTERNAL_CODE(ByVal vData As String)
   mvarFIELD_INTERNAL_CODE = vData
End Property

Public Property Get FIELD_INTERNAL_CODE() As String
   FIELD_INTERNAL_CODE = mvarFIELD_INTERNAL_CODE
End Property


