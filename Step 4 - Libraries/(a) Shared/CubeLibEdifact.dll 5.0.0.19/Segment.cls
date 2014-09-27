VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Segment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SDataItems"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5CC021C"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Attribute VB_Ext_KEY = "Member1" ,"Segments"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNCTS_IEM_TMS_ID   As Long
Private mvarSegmentTag        As String
Private mvarDescription       As String
Private mvarEDIUsage          As String
Private mvarEDIOccurrence     As Long
Private mvarNCTSUsage         As String
Private mvarNCTSOccurrence    As Long
Private mvarKey               As String
Private mvarDataNCTSSegmentID As Long
Private mblnIsSaved           As Boolean
Private mvarKeysChildren      As String
Private mvarKeyParent         As String
Private mvarSDataItems        As SDataItems

Public Property Let KeyParent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeyParent = 5
    mvarKeyParent = vData
End Property

Public Property Get KeyParent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeyParent
    KeyParent = mvarKeyParent
End Property

Public Property Let KeysChildren(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeysChildren = 5
    mvarKeysChildren = vData
End Property

Public Property Get KeysChildren() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeysChildren
    KeysChildren = mvarKeysChildren
End Property

Friend Property Let IsSaved(ByVal blnData As Boolean)
    mblnIsSaved = blnData
End Property

Friend Property Get IsSaved() As Boolean
    IsSaved = mblnIsSaved
End Property

Public Property Let DataNCTSSegmentID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataNCTSSegmentID = 5
    mvarDataNCTSSegmentID = vData
End Property


Public Property Get DataNCTSSegmentID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataNCTSSegmentID
    DataNCTSSegmentID = mvarDataNCTSSegmentID
End Property




Public Property Let NCTS_IEM_TMS_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsInUse = 5
    mvarNCTS_IEM_TMS_ID = vData
End Property


Public Property Get NCTS_IEM_TMS_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsInUse
    NCTS_IEM_TMS_ID = mvarNCTS_IEM_TMS_ID
End Property


Friend Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property


Public Property Let EDIUsage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EDIUsage = 5
    mvarEDIUsage = vData
End Property


Public Property Get EDIUsage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EDIUsage
    EDIUsage = mvarEDIUsage
End Property


Public Property Let EDIOccurrence(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EDIOccurrence = 5
    mvarEDIOccurrence = vData
End Property


Public Property Get EDIOccurrence() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EDIOccurrence
    EDIOccurrence = mvarEDIOccurrence
End Property

Public Property Let NCTSUsage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Usage = 5
    mvarNCTSUsage = vData
End Property

Public Property Get NCTSUsage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Usage
    NCTSUsage = mvarNCTSUsage
End Property

Public Property Let NCTSOccurrence(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Occurrence = 5
    mvarNCTSOccurrence = vData
End Property


Public Property Get NCTSOccurrence() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Occurrence
    NCTSOccurrence = mvarNCTSOccurrence
End Property




Public Property Get SDataItems() As SDataItems
    If mvarSDataItems Is Nothing Then
        Set mvarSDataItems = New SDataItems
    End If
    Set SDataItems = mvarSDataItems
End Property

Public Property Set SDataItems(vData As SDataItems)
    Set mvarSDataItems = vData
End Property


Private Sub Class_Terminate()
    Set mvarSDataItems = Nothing
End Sub

Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarDescription = vData
End Property

Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Description = mvarDescription
End Property

Public Property Let SegmentTag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarSegmentTag = vData
End Property

Public Property Get SegmentTag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    SegmentTag = mvarSegmentTag
End Property
