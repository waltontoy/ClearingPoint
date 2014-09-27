VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lngFeatureID As Long
Private m_lngFeatureCode As Long
Private m_strFeatureName As String


'Feature ID
Public Property Let FeatureID(ByVal vData As Long)
    m_lngFeatureID = vData
End Property
Public Property Get FeatureID() As Long
    FeatureID = m_lngFeatureID
End Property


'Feature Code
Public Property Let FeatureCode(ByVal vData As Long)
    m_lngFeatureCode = vData
End Property
Public Property Get FeatureCode() As Long
    FeatureCode = m_lngFeatureCode
End Property


'Feature Name
Public Property Let FeatureName(ByVal vData As String)
    m_strFeatureName = vData
End Property
Public Property Get FeatureName() As String
    FeatureName = m_strFeatureName
End Property
