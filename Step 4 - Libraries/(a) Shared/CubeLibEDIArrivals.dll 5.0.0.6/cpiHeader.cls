VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

' MOVEMENT_REFERENCE_NUMBER ' MR
' ARRIVAL_NOTIFICATION_DATE ' MC
' ARRIVAL_NOTIFICATION_PLACE ' BC
' ARRIVAL_NOTIFICATION_PLACE_LNG '
' ARRIVAL_AGREED_LOCATION_CODE ' AH
' ARRIVAL_AGREED_LOCATION_OF_GOODS ' AG
' ARRIVAL AGREED LOCATION OF GOODS LNG
' ARRIVAL_AUTHORISED_LOCATION_OF_GOODS ' AI
' DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION ' AJ
' SIMPLIFIED_PROCEDURE_FLAG ' EP

Private mvarARRIVAL_NOTIFICATION_PLACE_LNG As String ' 4
Private mvarARRIVAL_AGREED_LOCATION_OF_GOODS_LNG As String ' 7

Private mvarCODE_FIELD As String ' 1
Private mvarMOVEMENT_REFERENCE_NUMBER As String ' 2
Private mvarARRIVAL_NOTIFICATION_DATE As String ' 3
Private mvarARRIVAL_NOTIFICATION_PLACE As String ' 4
Private mvarARRIVAL_AGREED_LOCATION_CODE As String ' 5
Private mvarARRIVAL_AGREED_LOCATION_OF_GOODS As String ' 6
Private mvarARRIVAL_AUTHORISED_LOCATION_OF_GOODS As String ' 7
Private mvarDIALOG_LANGUAGE_INDICATOR_AT_DESTINATION As String ' 8
Private mvarSIMPLIFIED_PROCEDURE_FLAG As String ' 9
Private mvarORDINAL_PARENT As Long ' 10
Private mvarORDINAL As Long ' 11
Private mvarTIR_DISCHARGE As String ' 12
Private mvarTIR_VOLET_PAGE_NUMBER As String ' 13
' additional
'TOTAL NUMBER OF ITEMS
'TOTAL NUMBER OF PACKAGES
'IDENTITY MEANS OF TRANSPORT AT DEPARTURE
'IDENTITY MEANS OF TRANSPORT AT DEPARTURE LNG
'NATIONALITY OF MEANS OF TRANSPORT AT DEPARTURE
'TOTAL GROSS MASS

Private mvarTOTAL_NUMBER_OF_ITEMS As String ' 12
Private mvarTOTAL_NUMBER_OF_PACKAGES As String ' 13
Private mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE As String ' 14
Private mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG As String ' 15
Private mvarNATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE As String ' 16
Private mvarTOTAL_GROSS_MASS As String ' 17



' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 2
Public Property Let MOVEMENT_REFERENCE_NUMBER(ByVal vData As String)
    mvarMOVEMENT_REFERENCE_NUMBER = vData
End Property

Public Property Get MOVEMENT_REFERENCE_NUMBER() As String
    MOVEMENT_REFERENCE_NUMBER = mvarMOVEMENT_REFERENCE_NUMBER
End Property

' 12
Public Property Let TIR_DISCHARGE(ByVal vData As String)
    mvarTIR_DISCHARGE = vData
End Property

Public Property Get TIR_DISCHARGE() As String
    TIR_DISCHARGE = mvarTIR_DISCHARGE
End Property


' 13
Public Property Let TIR_VOLET_PAGE_NUMBER(ByVal vData As String)
    mvarTIR_VOLET_PAGE_NUMBER = vData
End Property

Public Property Get TIR_VOLET_PAGE_NUMBER() As String
    TIR_VOLET_PAGE_NUMBER = mvarTIR_VOLET_PAGE_NUMBER
End Property

' 3
Public Property Let ARRIVAL_NOTIFICATION_DATE(ByVal vData As String)
    mvarARRIVAL_NOTIFICATION_DATE = vData
End Property

Public Property Get ARRIVAL_NOTIFICATION_DATE() As String
    ARRIVAL_NOTIFICATION_DATE = mvarARRIVAL_NOTIFICATION_DATE
End Property

' 4
Public Property Let ARRIVAL_NOTIFICATION_PLACE(ByVal vData As String)
    mvarARRIVAL_NOTIFICATION_PLACE = vData
End Property

Public Property Get ARRIVAL_NOTIFICATION_PLACE() As String
    ARRIVAL_NOTIFICATION_PLACE = mvarARRIVAL_NOTIFICATION_PLACE
End Property

'
Public Property Let ARRIVAL_NOTIFICATION_PLACE_LNG(ByVal vData As String)
    mvarARRIVAL_NOTIFICATION_PLACE_LNG = vData
End Property

Public Property Get ARRIVAL_NOTIFICATION_PLACE_LNG() As String
    ARRIVAL_NOTIFICATION_PLACE_LNG = mvarARRIVAL_NOTIFICATION_PLACE_LNG
End Property

' 5
Public Property Let ARRIVAL_AGREED_LOCATION_CODE(ByVal vData As String)
    mvarARRIVAL_AGREED_LOCATION_CODE = vData
End Property

Public Property Get ARRIVAL_AGREED_LOCATION_CODE() As String
    ARRIVAL_AGREED_LOCATION_CODE = mvarARRIVAL_AGREED_LOCATION_CODE
End Property

' 6
Public Property Let ARRIVAL_AGREED_LOCATION_OF_GOODS(ByVal vData As String)
    mvarARRIVAL_AGREED_LOCATION_OF_GOODS = vData
End Property

Public Property Get ARRIVAL_AGREED_LOCATION_OF_GOODS() As String
    ARRIVAL_AGREED_LOCATION_OF_GOODS = mvarARRIVAL_AGREED_LOCATION_OF_GOODS
End Property

'
Public Property Let ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG(ByVal vData As String)
    mvarARRIVAL_AGREED_LOCATION_OF_GOODS_LNG = vData
End Property

Public Property Get ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG() As String
    ARRIVAL_AGREED_LOCATION_OF_GOODS_LNG = mvarARRIVAL_AGREED_LOCATION_OF_GOODS_LNG
End Property

' 7
Public Property Let ARRIVAL_AUTHORISED_LOCATION_OF_GOODS(ByVal vData As String)
    mvarARRIVAL_AUTHORISED_LOCATION_OF_GOODS = vData
End Property

Public Property Get ARRIVAL_AUTHORISED_LOCATION_OF_GOODS() As String
    ARRIVAL_AUTHORISED_LOCATION_OF_GOODS = mvarARRIVAL_AUTHORISED_LOCATION_OF_GOODS
End Property

' 8
Public Property Let DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION(ByVal vData As String)
    mvarDIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = vData
End Property

Public Property Get DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION() As String
    DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = mvarDIALOG_LANGUAGE_INDICATOR_AT_DESTINATION
End Property

' 9
Public Property Let SIMPLIFIED_PROCEDURE_FLAG(ByVal vData As String)
    mvarSIMPLIFIED_PROCEDURE_FLAG = vData
End Property

Public Property Get SIMPLIFIED_PROCEDURE_FLAG() As String
    SIMPLIFIED_PROCEDURE_FLAG = mvarSIMPLIFIED_PROCEDURE_FLAG
End Property

' 10
Public Property Let ORDINAL_PARENT(ByVal vData As Long)
    mvarORDINAL_PARENT = vData
End Property

Public Property Get ORDINAL_PARENT() As Long
    ORDINAL_PARENT = mvarORDINAL_PARENT
End Property

' 11
Public Property Let ORDINAL(ByVal vData As Long)
    mvarORDINAL = vData
End Property

Public Property Get ORDINAL() As Long
    ORDINAL = mvarORDINAL
End Property

' 12
Public Property Let TOTAL_NUMBER_OF_ITEMS(ByVal vData As String)
    mvarTOTAL_NUMBER_OF_ITEMS = vData
End Property

Public Property Get TOTAL_NUMBER_OF_ITEMS() As String
    TOTAL_NUMBER_OF_ITEMS = mvarTOTAL_NUMBER_OF_ITEMS
End Property

' 13
Public Property Let TOTAL_NUMBER_OF_PACKAGES(ByVal vData As String)
    mvarTOTAL_NUMBER_OF_PACKAGES = vData
End Property

Public Property Get TOTAL_NUMBER_OF_PACKAGES() As String
    TOTAL_NUMBER_OF_PACKAGES = mvarTOTAL_NUMBER_OF_PACKAGES
End Property

' 14
Public Property Let IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE(ByVal vData As String)
    mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE = vData
End Property

Public Property Get IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE() As String
    IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE = mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE
End Property

' 15
Public Property Let IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG(ByVal vData As String)
    mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG = vData
End Property

Public Property Get IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG() As String
    IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG = mvarIDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG
End Property

' 16
Public Property Let NATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE(ByVal vData As String)
    mvarNATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE = vData
End Property

Public Property Get NATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE() As String
    NATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE = mvarNATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE
End Property

' 17
Public Property Let TOTAL_GROSS_MASS(ByVal vData As String)
    mvarTOTAL_GROSS_MASS = vData
End Property

Public Property Get TOTAL_GROSS_MASS() As String
    TOTAL_GROSS_MASS = mvarTOTAL_GROSS_MASS
End Property


