VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiEnRouteEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

' Incidents
' Transhipments
' NewSeals
' Ctl_ Controls
' PLACE
' PLACE_LNG
' COUNTRY_CODE
Private mvarPLACE_LNG As String '

Private mvarCODE_FIELD As String ' 1
Private mvarPLACE As String ' 2 ' BF
Private mvarCOUNTRY_CODE As String ' 3 ' C7
Private mvarORDINAL_PARENT As Long ' 4
Private mvarORDINAL As Long ' 5

Private mvarIncidents As cpiIncidents ' 6
Private mvarTranshipments As cpiTranshipments  ' 7
Private mvarNewSeals As cpiSealInfos  ' 8
Private mvarCtl_Controls As cpiControls  ' 9

' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 2
Public Property Let PLACE(ByVal vData As String)
    mvarPLACE = vData
End Property

Public Property Get PLACE() As String
    PLACE = mvarPLACE
End Property

'
Public Property Let PLACE_LNG(ByVal vData As String)
    mvarPLACE_LNG = vData
End Property

Public Property Get PLACE_LNG() As String
    PLACE_LNG = mvarPLACE_LNG
End Property

' 3
Public Property Let COUNTRY_CODE(ByVal vData As String)
    mvarCOUNTRY_CODE = vData
End Property

Public Property Get COUNTRY_CODE() As String
    COUNTRY_CODE = mvarCOUNTRY_CODE
End Property

' 4
Public Property Let ORDINAL_PARENT(ByVal vData As Long)
    mvarORDINAL_PARENT = vData
End Property

Public Property Get ORDINAL_PARENT() As Long
    ORDINAL_PARENT = mvarORDINAL_PARENT
End Property

' 5
Public Property Let ORDINAL(ByVal vData As Long)
    mvarORDINAL = vData
End Property

Public Property Get ORDINAL() As Long
    ORDINAL = mvarORDINAL
End Property

' 6
Public Property Set Incidents(ByVal vData As cpiIncidents)
    Set mvarIncidents = vData
End Property

Public Property Get Incidents() As cpiIncidents
    Set Incidents = mvarIncidents
End Property

' 7
Public Property Set Transhipments(ByVal vData As cpiTranshipments)
    Set mvarTranshipments = vData
End Property

Public Property Get Transhipments() As cpiTranshipments
    Set Transhipments = mvarTranshipments
End Property

' 8
Public Property Set NewSeals(ByVal vData As cpiSealInfos)
    Set mvarNewSeals = vData
End Property

Public Property Get NewSeals() As cpiSealInfos
    Set NewSeals = mvarNewSeals
End Property

' 9
Public Property Set Ctl_Controls(ByVal vData As cpiControls)
    Set mvarCtl_Controls = vData
End Property

Public Property Get Ctl_Controls() As cpiControls
    Set Ctl_Controls = mvarCtl_Controls
End Property


