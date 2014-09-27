VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiIncident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

' INCIDENT_FLAG ' ER
' INCIDENT_INFORMATION ' ES
' INCIDENT_INFORMATION_LNG '
' ENDORSEMENT DATE ' C8
' ENDORSEMENT AUTHORITY ' C9
' ENDORSEMENT AUTHORITY LNG
' ENDORSEMENT PLACE ' CA
' ENDORSEMENT PLACE LNG
' ENDORSEMENT COUNTRY ' CB

Private mvarINCIDENT_INFORMATION_LNG As String ' ?
Private mvarENDORSEMENT_AUTHORITY_LNG As String ' ?
Private mvarENDORSEMENT_PLACE_LNG As String ' ?

Private mvarCODE_FIELD As String ' 1
Private mvarINCIDENT_FLAG As String ' 2
Private mvarINCIDENT_INFORMATION As String ' 3
Private mvarENDORSEMENT_DATE As String ' 4
Private mvarENDORSEMENT_AUTHORITY As String ' 5
Private mvarENDORSEMENT_PLACE As String ' 6
Private mvarENDORSEMENT_COUNTRY As String ' 7
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
Public Property Let INCIDENT_FLAG(ByVal vData As String)
    mvarINCIDENT_FLAG = vData
End Property

Public Property Get INCIDENT_FLAG() As String
    INCIDENT_FLAG = mvarINCIDENT_FLAG
End Property

' 3
Public Property Let INCIDENT_INFORMATION(ByVal vData As String)
    mvarINCIDENT_INFORMATION = vData
End Property

Public Property Get INCIDENT_INFORMATION() As String
    INCIDENT_INFORMATION = mvarINCIDENT_INFORMATION
End Property

' 4
Public Property Let INCIDENT_INFORMATION_LNG(ByVal vData As String)
    mvarINCIDENT_INFORMATION_LNG = vData
End Property

Public Property Get INCIDENT_INFORMATION_LNG() As String
    INCIDENT_INFORMATION_LNG = mvarINCIDENT_INFORMATION_LNG
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

