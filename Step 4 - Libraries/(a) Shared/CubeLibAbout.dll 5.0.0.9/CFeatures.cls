VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CFeature"
Attribute VB_Ext_KEY = "Member0" ,"CFeature"
Option Explicit

Private m_colActiveFeatures As Collection

Public Function Add(ByVal FeatureID As Long, _
                    ByVal FeatureCode As Long, _
                    ByVal FeatureName As String, _
           Optional ByVal Key As String) As CFeature
    
    Dim objNewMember As CFeature
    
    
    ' Create a new object
    Set objNewMember = New CFeature
    
    
    ' Set the properties passed into the method
    objNewMember.FeatureID = FeatureID
    objNewMember.FeatureCode = FeatureCode
    objNewMember.FeatureName = FeatureName
    
    
    ' Add new member into the collection
    If (Len(Key) = 0) Then
        m_colActiveFeatures.Add objNewMember
    Else
        m_colActiveFeatures.Add objNewMember, Key
    End If
    
    
    ' Return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(ByVal IndexKey As Variant) As CFeature
Attribute Item.VB_UserMemId = 0
    Set Item = m_colActiveFeatures(IndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colActiveFeatures.Count
End Property

Public Sub Remove(IndexKey As Variant)
    m_colActiveFeatures.Remove IndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colActiveFeatures.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colActiveFeatures = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colActiveFeatures = Nothing
End Sub
