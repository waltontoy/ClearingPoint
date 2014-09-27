VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiTrader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

' DESTINATION_TIN ' W8
' DESTINATION_NAME ' W9
' DESTINATION_STREET_AND_NUMBER ' WA
' DESTINATION_COUNTRY_CODE ' WE
' DESTINATION_POSTAL_CODE ' WD
' DESTINATION_CITY ' WB
' DESTINATION_NAD_LNG
Private mvarDESTINATION_NAD_LNG As String '

Private mvarCODE_FIELD As String ' 1
Private mvarDESTINATION_TIN As String ' 2
Private mvarDESTINATION_NAME As String ' 3
Private mvarDESTINATION_STREET_AND_NUMBER As String ' 4
Private mvarDESTINATION_COUNTRY_CODE As String ' 5
Private mvarDESTINATION_POSTAL_CODE As String ' 6
Private mvarDESTINATION_CITY As String ' 7
Private mvarORDINAL_PARENT As Long ' 8
Private mvarORDINAL As Long ' 9

' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 2
Public Property Let DESTINATION_TIN(ByVal vData As String)
    mvarDESTINATION_TIN = vData
End Property

Public Property Get DESTINATION_TIN() As String
    DESTINATION_TIN = mvarDESTINATION_TIN
End Property

' 3
Public Property Let DESTINATION_NAME(ByVal vData As String)
    mvarDESTINATION_NAME = vData
End Property

Public Property Get DESTINATION_NAME() As String
    DESTINATION_NAME = mvarDESTINATION_NAME
End Property

' 4
Public Property Let DESTINATION_STREET_AND_NUMBER(ByVal vData As String)
    mvarDESTINATION_STREET_AND_NUMBER = vData
End Property

Public Property Get DESTINATION_STREET_AND_NUMBER() As String
    DESTINATION_STREET_AND_NUMBER = mvarDESTINATION_STREET_AND_NUMBER
End Property

' 5
Public Property Let DESTINATION_COUNTRY_CODE(ByVal vData As String)
    mvarDESTINATION_COUNTRY_CODE = vData
End Property

Public Property Get DESTINATION_COUNTRY_CODE() As String
    DESTINATION_COUNTRY_CODE = mvarDESTINATION_COUNTRY_CODE
End Property

' 6
Public Property Let DESTINATION_POSTAL_CODE(ByVal vData As String)
    mvarDESTINATION_POSTAL_CODE = vData
End Property

Public Property Get DESTINATION_POSTAL_CODE() As String
    DESTINATION_POSTAL_CODE = mvarDESTINATION_POSTAL_CODE
End Property

' 7
Public Property Let DESTINATION_CITY(ByVal vData As String)
    mvarDESTINATION_CITY = vData
End Property

Public Property Get DESTINATION_CITY() As String
    DESTINATION_CITY = mvarDESTINATION_CITY
End Property


Public Property Let DESTINATION_NAD_LNG(ByVal vData As String)
    mvarDESTINATION_NAD_LNG = vData
End Property

Public Property Get DESTINATION_NAD_LNG() As String
    DESTINATION_NAD_LNG = mvarDESTINATION_NAD_LNG
End Property

' 8
Public Property Let ORDINAL_PARENT(ByVal vData As Long)
    mvarORDINAL_PARENT = vData
End Property

Public Property Get ORDINAL_PARENT() As Long
    ORDINAL_PARENT = mvarORDINAL_PARENT
End Property

' 9
Public Property Let ORDINAL(ByVal vData As Long)
    mvarORDINAL = vData
End Property

Public Property Get ORDINAL() As Long
    ORDINAL = mvarORDINAL
End Property

