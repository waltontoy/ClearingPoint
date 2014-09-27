VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSQLPickRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cpiPickDataFields"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCancelled As Boolean 'local copy
Private mvarcpiPickDataFields As CPickDataFields


Public Property Get cpiPickDataFields() As CPickDataFields
    If mvarcpiPickDataFields Is Nothing Then
        Set mvarcpiPickDataFields = New CPickDataFields
    End If


    Set cpiPickDataFields = mvarcpiPickDataFields
End Property


Public Property Set cpiPickDataFields(vData As CPickDataFields)
    Set mvarcpiPickDataFields = vData
End Property
Private Sub Class_Terminate()
    Set mvarcpiPickDataFields = Nothing
End Sub



Public Property Let Cancelled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cancelled = 5
    mvarCancelled = vData
End Property


Public Property Get Cancelled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cancelled
    Cancelled = mvarCancelled
End Property



