VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiIE44Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Index As Long

' Headers
' GoodsItems
' CustomOffices
' Traders
' DATE_OF_PREPARATION
' TIME_OF_PREPARATION

Private mvarCODE_FIELD As String ' 1
Private mvarNCTS_IEM_ID As Long   ' 2
Private mvarDATE_OF_PREPARATION As String ' 3
Private mvarTIME_OF_PREPARATION As String ' 4
Private mvarORDINAL As Long  ' 5

Private mvarHeaders As cpiHeaders ' 6
Private mvarGoodsItems As cpiGoodsItems   ' 7
Private mvarCustomOffices As cpiCustomOffices   ' 8
Private mvarTraders As cpiTraders  ' 9

Private mvarUnloadingRemarks As cpiUnloadingRemarks  ' 9b
Private mvarSealInfos As cpiOldSeals_Tables ' 9c
Private mvarResultOfControls As cpiResultOfControls  ' 9d


Private mvarMESSAGE_IDENTIFICATION As String '10
Private mvarINTERCHAGE_CONTROL_COUNT As Integer ' 11
Private mvarINTERCHAGE_CONTROL_REFERENCE As String ' 12
Private mvarTYPE_OF_DECLARATION As String ' 13
Private mvarNUMBER_OF_SEGMENTS As Integer

' 1
Public Property Let CODE_FIELD(ByVal vData As String)
    mvarCODE_FIELD = vData
End Property

Public Property Get CODE_FIELD() As String
    CODE_FIELD = mvarCODE_FIELD
End Property

' 2
Public Property Let NCTS_IEM_ID(ByVal vData As Long)
    mvarNCTS_IEM_ID = vData
End Property

Public Property Get NCTS_IEM_ID() As Long
    NCTS_IEM_ID = mvarNCTS_IEM_ID
End Property

' 3
Public Property Let DATE_OF_PREPARATION(ByVal vData As String)
    mvarDATE_OF_PREPARATION = vData
End Property

Public Property Get DATE_OF_PREPARATION() As String
    DATE_OF_PREPARATION = mvarDATE_OF_PREPARATION
End Property

' 4
Public Property Let TIME_OF_PREPARATION(ByVal vData As String)
    mvarTIME_OF_PREPARATION = vData
End Property

Public Property Get TIME_OF_PREPARATION() As String
    TIME_OF_PREPARATION = mvarTIME_OF_PREPARATION
End Property

' 5
Public Property Let ORDINAL(ByVal vData As Long)
    mvarORDINAL = vData
End Property

Public Property Get ORDINAL() As Long
    ORDINAL = mvarORDINAL
End Property

' 6
Public Property Set Headers(ByVal vData As cpiHeaders)
    Set mvarHeaders = vData
End Property

Public Property Get Headers() As cpiHeaders
    Set Headers = mvarHeaders
End Property

' 7
Public Property Set GoodsItems(ByVal vData As cpiGoodsItems)
    Set mvarGoodsItems = vData
End Property

Public Property Get GoodsItems() As cpiGoodsItems
    Set GoodsItems = mvarGoodsItems
End Property

' 8
Public Property Set CustomOffices(ByVal vData As cpiCustomOffices)
    Set mvarCustomOffices = vData
End Property

Public Property Get CustomOffices() As cpiCustomOffices
    Set CustomOffices = mvarCustomOffices
End Property

' 9
Public Property Set Traders(ByVal vData As cpiTraders)
    Set mvarTraders = vData
End Property

Public Property Get Traders() As cpiTraders
    Set Traders = mvarTraders
End Property

' 10
Public Property Let MESSAGE_IDENTIFICATION(ByVal vData As String)
    mvarMESSAGE_IDENTIFICATION = vData
End Property

Public Property Get MESSAGE_IDENTIFICATION() As String
    MESSAGE_IDENTIFICATION = mvarMESSAGE_IDENTIFICATION
End Property

' 11
Public Property Let INTERCHAGE_CONTROL_COUNT(ByVal vData As Integer)
    mvarINTERCHAGE_CONTROL_COUNT = vData
End Property

Public Property Get INTERCHAGE_CONTROL_COUNT() As Integer
    INTERCHAGE_CONTROL_COUNT = mvarINTERCHAGE_CONTROL_COUNT
End Property

' 12
Public Property Let INTERCHAGE_CONTROL_REFERENCE(ByVal vData As String)
    mvarINTERCHAGE_CONTROL_REFERENCE = vData
End Property

Public Property Get INTERCHAGE_CONTROL_REFERENCE() As String
    INTERCHAGE_CONTROL_REFERENCE = mvarINTERCHAGE_CONTROL_REFERENCE
End Property

' 13
Public Property Let TYPE_OF_DECLARATION(ByVal vData As String)
    mvarTYPE_OF_DECLARATION = vData
End Property

Public Property Get TYPE_OF_DECLARATION() As String
    TYPE_OF_DECLARATION = mvarTYPE_OF_DECLARATION
End Property

' 14
Public Property Let NUMBER_OF_SEGMENTS(ByVal vData As Integer)
    mvarNUMBER_OF_SEGMENTS = vData
End Property

Public Property Get NUMBER_OF_SEGMENTS() As Integer
    NUMBER_OF_SEGMENTS = mvarNUMBER_OF_SEGMENTS
End Property

Private Sub Class_Terminate()

    Set mvarHeaders = Nothing
    Set mvarGoodsItems = Nothing
    Set mvarCustomOffices = Nothing
    Set mvarTraders = Nothing
    Set mvarResultOfControls = Nothing
    

End Sub

' 9b
Public Property Set UnloadingRemarks(ByVal vData As cpiUnloadingRemarks)
    Set mvarUnloadingRemarks = vData
End Property

Public Property Get UnloadingRemarks() As cpiUnloadingRemarks
    Set UnloadingRemarks = mvarUnloadingRemarks
End Property

' 9c
Public Property Set SealInfos(ByVal vData As cpiOldSeals_Tables)
    Set mvarSealInfos = vData
End Property

Public Property Get SealInfos() As cpiOldSeals_Tables
    Set SealInfos = mvarSealInfos
End Property

' 9d
Public Property Set ResultOfControls(ByVal vData As cpiResultOfControls)
    Set mvarResultOfControls = vData
End Property

Public Property Get ResultOfControls() As cpiResultOfControls
    Set ResultOfControls = mvarResultOfControls
End Property


