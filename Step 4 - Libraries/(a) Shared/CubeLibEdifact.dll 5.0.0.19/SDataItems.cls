VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDataItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SDataItem"
Attribute VB_Ext_KEY = "Member0" ,"SDataItem"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5CC0366"
Option Explicit

'local variable to hold collection
Private mCol As Collection

'Public Function Add(Tag As String, Name As String, Usage As String, DataFormat As String, CodeListID As Long, Value As String, MapFieldName As String, MapTableName As String, Key As String) As SDataItem
Public Function Add(Key As String, _
                    NCTS_ITM_ID As Long, _
                    EDIName As String, _
                    EDIDataFormat As String, _
                    EDIUsage As String, _
                    EDIOrdinal As Long, _
                    NCTSName As String, _
                    NCTSDataFormat As String, _
                    NCTSUsage As String, _
                    CodeListID As Long, _
                    EDIGroupID As Long, _
                    Value As String) _
                    As SDataItem

        
    Dim objNewMember As SDataItem
    Set objNewMember = New SDataItem
        
    objNewMember.NCTS_ITM_ID = NCTS_ITM_ID
    objNewMember.EDIName = EDIName
    objNewMember.EDIDataFormat = EDIDataFormat
    objNewMember.EDIUsage = EDIUsage
    objNewMember.EDIOrdinal = EDIOrdinal
    objNewMember.NCTSName = NCTSName
    objNewMember.NCTSDataFormat = NCTSDataFormat
    objNewMember.NCTSUsage = NCTSUsage
    
    objNewMember.CodeListID = CodeListID
    objNewMember.EDIGroupID = EDIGroupID
    objNewMember.Value = Value
    objNewMember.Key = Key
            
    mCol.Add objNewMember, Key

    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As SDataItem
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
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
