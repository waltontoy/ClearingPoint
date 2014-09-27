VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SegmentCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'local variable(s) to hold property value(s)
Private mvarCount As Long 'local copy
Public Property Let Count(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Count = 5
    mvarCount = vData
End Property


Public Property Get Count() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    Count = mvarCount
End Property



