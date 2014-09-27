VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiTranshipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

Private mvarNEW_TRANSPORT_MEANS_IDENTITY_LNG As String ' ?
Private mvarENDORSEMENT_AUTHORITY_LNG As String ' ?
Private mvarENDORSEMENT_PLACE_LNG As String ' ?

Private mvarCODE_FIELD As String ' 1
Private mvarNEW_TRANSPORT_MEANS_IDENTITY As String ' 2 BL
Private mvarNEW_TRANSPORT_MEANS_NATIONALITY As String ' 3 BG
Private mvarENDORSEMENT_DATE As String ' 4 BH
Private mvarENDORSEMENT_AUTHORITY As String ' 5 BI
Private mvarENDORSEMENT_PLACE As String ' 6 BJ
Private mvarENDORSEMENT_COUNTRY As String ' 7  BK
Private mvarORDINAL_PARENT As Long ' 8
Private mvarORDINAL As Long ' 9

Private mvarContainers As cpiContainers ' 10


' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 1
Public Property Set Containers(ByVal vData As cpiContainers)
    Set mvarContainers = vData
End Property

Public Property Get Containers() As cpiContainers
    Set Containers = mvarContainers
End Property

' 2
Public Property Let NEW_TRANSPORT_MEANS_IDENTITY(ByVal vData As String)
    mvarNEW_TRANSPORT_MEANS_IDENTITY = vData
End Property

Public Property Get NEW_TRANSPORT_MEANS_IDENTITY() As String
    NEW_TRANSPORT_MEANS_IDENTITY = mvarNEW_TRANSPORT_MEANS_IDENTITY
End Property

' 3
Public Property Let NEW_TRANSPORT_MEANS_IDENTITY_LNG(ByVal vData As String)
    mvarNEW_TRANSPORT_MEANS_IDENTITY_LNG = vData
End Property

Public Property Get NEW_TRANSPORT_MEANS_IDENTITY_LNG() As String
    NEW_TRANSPORT_MEANS_IDENTITY_LNG = mvarNEW_TRANSPORT_MEANS_IDENTITY_LNG
End Property

' 4
Public Property Let NEW_TRANSPORT_MEANS_NATIONALITY(ByVal vData As String)
    mvarNEW_TRANSPORT_MEANS_NATIONALITY = vData
End Property

Public Property Get NEW_TRANSPORT_MEANS_NATIONALITY() As String
    NEW_TRANSPORT_MEANS_NATIONALITY = mvarNEW_TRANSPORT_MEANS_NATIONALITY
End Property

' 5
Public Property Let ENDORSEMENT_DATE(ByVal vData As String)
    mvarENDORSEMENT_DATE = vData
End Property

Public Property Get ENDORSEMENT_DATE() As String
    ENDORSEMENT_DATE = mvarENDORSEMENT_DATE
End Property

' 6
Public Property Let ENDORSEMENT_AUTHORITY(ByVal vData As String)
    mvarENDORSEMENT_AUTHORITY = vData
End Property

Public Property Get ENDORSEMENT_AUTHORITY() As String
    ENDORSEMENT_AUTHORITY = mvarENDORSEMENT_AUTHORITY
End Property

' 7
Public Property Let ENDORSEMENT_AUTHORITY_LNG(ByVal vData As String)
    mvarENDORSEMENT_AUTHORITY_LNG = vData
End Property

Public Property Get ENDORSEMENT_AUTHORITY_LNG() As String
    ENDORSEMENT_AUTHORITY_LNG = mvarENDORSEMENT_AUTHORITY_LNG
End Property

' 8
Public Property Let ENDORSEMENT_PLACE(ByVal vData As String)
    mvarENDORSEMENT_PLACE = vData
End Property

Public Property Get ENDORSEMENT_PLACE() As String
    ENDORSEMENT_PLACE = mvarENDORSEMENT_PLACE
End Property

' 9
Public Property Let ENDORSEMENT_PLACE_LNG(ByVal vData As String)
    mvarENDORSEMENT_PLACE_LNG = vData
End Property

Public Property Get ENDORSEMENT_PLACE_LNG() As String
    ENDORSEMENT_PLACE_LNG = mvarENDORSEMENT_PLACE_LNG
End Property

' 10
Public Property Let ENDORSEMENT_COUNTRY(ByVal vData As String)
    mvarENDORSEMENT_COUNTRY = vData
End Property

Public Property Get ENDORSEMENT_COUNTRY() As String
    ENDORSEMENT_COUNTRY = mvarENDORSEMENT_COUNTRY
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

