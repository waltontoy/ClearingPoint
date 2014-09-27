VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDIProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CEDIProperty"
Attribute VB_Ext_KEY = "Member0" ,"CEDIProperty"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, QueueName As String, UserName As String, Password As String, RemoteHost As String, RremotePort As String, Optional sKey As String) As CEDIProperty
    'create a new object
    Dim objNewMember As CEDIProperty
    Set objNewMember = New CEDIProperty


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.QueueName = QueueName
    objNewMember.UserName = UserName
    objNewMember.Password = Password
    objNewMember.RemoteHost = RemoteHost
    objNewMember.RemotePort = RremotePort
    objNewMember.Disabled = False
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CEDIProperty
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


    mCol.Remove vntIndexKey
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
    Dim lngItemCtr As Long
    Dim clsProperty As CEDIProperty
    
    For Each clsProperty In Me
        Set clsProperty = Nothing
    Next
    
    'destroys collection when this class is terminated
    'For lngItemCtr = 1 To mCol.Count
    '    Set mCol.Item(lngItemCtr) = Nothing
    'Next
    
    Set mCol = Nothing
End Sub

