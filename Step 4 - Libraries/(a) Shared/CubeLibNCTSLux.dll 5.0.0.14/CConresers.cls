VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConresers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarConResCodERS16 As String
Private mvarDatLimERS69 As String

Public Property Let ConResCodERS16(ByVal vData As String)
    mvarConResCodERS16 = vData
End Property

Public Property Get ConResCodERS16() As String
    ConResCodERS16 = mvarConResCodERS16
End Property

Public Property Let DatLimERS69(ByVal vData As String)
    mvarDatLimERS69 = vData
End Property

Public Property Get DatLimERS69() As String
    DatLimERS69 = mvarDatLimERS69
End Property
