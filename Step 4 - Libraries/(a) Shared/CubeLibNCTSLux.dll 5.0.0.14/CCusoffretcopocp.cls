VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCusoffretcopocp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRefNumOCP1 As String
Private mCusOffNamOCP2 As String
Private mStrAndNumOCP3 As String
Private mCouOCP4 As String
Private mPosCodOCP6 As String
Private mCitOCP7 As String

'//////////////////////////////////////////////////////////////////////////////
Public Property Let RefNumOCP1(ByVal vData As String)
    mRefNumOCP1 = vData
End Property

Public Property Get RefNumOCP1() As String
        RefNumOCP1 = mRefNumOCP1
End Property

Public Property Let CusOffNamOCP2(ByVal vData As String)
    mCusOffNamOCP2 = vData
End Property

Public Property Get CusOffNamOCP2() As String
        CusOffNamOCP2 = mCusOffNamOCP2
End Property

Public Property Let StrAndNumOCP3(ByVal vData As String)
    mStrAndNumOCP3 = vData
End Property

Public Property Get StrAndNumOCP3() As String
        StrAndNumOCP3 = mStrAndNumOCP3
End Property

Public Property Let CouOCP4(ByVal vData As String)
    mCouOCP4 = vData
End Property

Public Property Get CouOCP4() As String
        CouOCP4 = mCouOCP4
End Property

Public Property Let PosCodOCP6(ByVal vData As String)
    mPosCodOCP6 = vData
End Property

Public Property Get PosCodOCP6() As String
        PosCodOCP6 = mPosCodOCP6
End Property

'/
Public Property Let CitOCP7(ByVal vData As String)
    mCitOCP7 = vData
End Property

Public Property Get CitOCP7() As String
        CitOCP7 = mCitOCP7
End Property



