VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSHEET_PROPS_Frms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CLINK"
Attribute VB_Ext_KEY = "Member0" ,"CLINK"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' local variable to hold collection
Private mCol As Collection

Public Function Add(ByRef sKey As String) As cpiSHEET_PROPS_Frm
    ' create a new object
    Dim objNewMember As cpiSHEET_PROPS_Frm
    Set objNewMember = New cpiSHEET_PROPS_Frm
    
    ' set the properties passed into the method
    
    objNewMember.Index = mCol.Count + 1
    objNewMember.Key = sKey
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    
    ' return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
        
End Function

Public Property Get Item(vntIndexKey As Variant) As cpiSHEET_PROPS_Frm
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    If (mCol.Count <> 0) Then
        mCol.Remove vntIndexKey
        'Unload mCol(vntIndexKey)
    End If
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()

    Dim frm As frmBoxProperties

    'destroys collection when this class is terminated
    
'    Debug.Print mCol.Count
    'For Each frm In mCol
    '    Unload frm
    'Next
    
    Set mCol = Nothing
End Sub

