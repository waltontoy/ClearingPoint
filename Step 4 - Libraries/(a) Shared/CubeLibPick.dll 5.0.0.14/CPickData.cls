VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPickData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarFieldName As String 'local copy
Private mvarValue As Variant 'local copy
Public Property Let Value(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Set Value(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Value = Form1
    Set mvarValue = vData
End Property


Public Property Get Value() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else
        Value = mvarValue
    End If
End Property



Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



