VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPicklistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PicklistFilter"
Attribute VB_Ext_KEY = "Member1" ,"PicklistFilters"
Option Explicit
    
    Private mvarFilterCount As Long
    
    Public Enum enuPicklistFilter
        cpiRadioOptions = 0
        cpiCheckOptions = 1
        cpiComboRecords = 2
    End Enum

Private mvarFilterType As enuPicklistFilter
'Private mvarFilterConnection As Variant
Private mvarPicklistFilters As CPicklistFilters


Public Property Get PicklistFilters() As CPicklistFilters
    If mvarPicklistFilters Is Nothing Then
        Set mvarPicklistFilters = New CPicklistFilters
    End If

    Set PicklistFilters = mvarPicklistFilters
End Property

Public Property Set PicklistFilters(vData As CPicklistFilters)
    Set mvarPicklistFilters = vData
End Property

Private Sub Class_Terminate()
    Set mvarPicklistFilters = Nothing
End Sub

Public Property Let FilterType(ByVal vData As enuPicklistFilter)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FilterType = 5
    mvarFilterType = vData
End Property


Public Property Get FilterType() As enuPicklistFilter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FilterType
    FilterType = mvarFilterType
End Property
    
Friend Property Let FilterCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FilterCount = 5
    mvarFilterCount = vData
End Property

Friend Property Get FilterCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FilterCount
    FilterCount = mvarFilterCount
End Property


