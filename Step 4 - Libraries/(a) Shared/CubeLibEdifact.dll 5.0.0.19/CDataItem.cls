VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5C80280"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTag As String 'local copy
Private mvarName As String 'local copy
Private mvarUsage As String 'local copy

Public Property Let Usage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Usage = 5
    mvarUsage = vData
End Property


Public Property Get Usage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Usage
    Usage = mvarUsage
End Property



Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property



Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property
