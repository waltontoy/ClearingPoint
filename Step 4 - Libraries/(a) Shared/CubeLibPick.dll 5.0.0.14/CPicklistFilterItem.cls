VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPicklistFilterItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'Public Enum cpiCheckValueConstants
'    cpiUnchecked
'    cpiChecked
'    cpiGrayed
'End Enum

Public Key As String

Private mvarFilter As String
Private mvarFilterCaption As String
'local variable(s) to hold property value(s)
Private mvarState As Boolean 'local copy
Private mvarEnabled As Boolean
'Private mvarValue As cpiCheckValueConstants
Private mvarValue As Boolean
Private mvarIndex As Long
Private mvarVisible As Boolean

Event OptionButtonClick(Index As Long)

Public Property Let Value(ByVal vData As Boolean)
    mvarValue = vData
    ' trigger the  event
    'RaiseEvent OptionButtonClick(mvarIndex)
End Property

Public Property Get Value() As Boolean
    Value = mvarValue
End Property

Public Property Let Index(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarIndex = vData
End Property

Public Property Get Index() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    Index = mvarIndex
End Property


Public Property Let Enabled(ByVal vData)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarEnabled = vData
End Property

Public Property Get Enabled()
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    Enabled = mvarEnabled
End Property

Public Property Let State(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property

Public Property Get State() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    State = mvarState
End Property

Public Property Let Filter(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filter = 5
    mvarFilter = vData
End Property

Public Property Get Filter() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filter
    Filter = mvarFilter
End Property

Public Property Let FilterCaption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filter = 5
    mvarFilterCaption = vData
End Property


Public Property Get FilterCaption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filter
    FilterCaption = mvarFilterCaption
End Property



Private Sub Class_Initialize()
   mvarEnabled = True
   mvarValue = False
End Sub

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property


