VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CDataItem"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5C70258"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

    Private mvarNCTS_ITM_ID As Long
    Private mvarEDIName As String 'local copy
    Private mvarEDIDataFormat As String 'local copy
    Private mvarEDIUsage As String 'local copy
    Private mvarEDIOrdinal As Long 'local copy
    
    Private mvarNCTSName As String 'local copy
    Private mvarNCTSDataFormat As String 'local copy
    Private mvarNCTSUsage As String 'local copy
    
    Private mvarCodeListID As Long 'local copy
    Private mvarValue As String 'local copy
    
    Private mvarKey As String
    
    Public Enum EConditionTabType
        ECondition_Header = 1
        ECondition_Detail = 2
    End Enum
    
'local variable(s) to hold property value(s)
Private mvarEDIGroupID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarNCTS_IEM_MCN_ID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarConditionTabType As EConditionTabType 'local copy
Public Property Let ConditionTabType(ByVal vData As EConditionTabType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConditionTabType = 5
    mvarConditionTabType = vData
End Property


Public Property Get ConditionTabType() As EConditionTabType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConditionTabType
    ConditionTabType = mvarConditionTabType
End Property



Public Property Let NCTS_IEM_MCN_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NCTS_IEM_MCN_ID = 5
    mvarNCTS_IEM_MCN_ID = vData
End Property


Public Property Get NCTS_IEM_MCN_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCTS_IEM_MCN_ID
    NCTS_IEM_MCN_ID = mvarNCTS_IEM_MCN_ID
End Property



Public Property Let EDIGroupID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarEDIGroupID = vData
End Property


Public Property Get EDIGroupID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    EDIGroupID = mvarEDIGroupID
End Property





Public Property Let NCTS_ITM_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsInUse = 5
    mvarNCTS_ITM_ID = vData
End Property


Public Property Get NCTS_ITM_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsInUse
    NCTS_ITM_ID = mvarNCTS_ITM_ID
End Property



Public Property Let EDIOrdinal(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsInUse = 5
    mvarEDIOrdinal = vData
End Property

Public Property Get EDIOrdinal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsInUse
    EDIOrdinal = mvarEDIOrdinal
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



Public Property Let EDIDataFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EDIDataFormat = 5
    mvarEDIDataFormat = vData
End Property


Public Property Get EDIDataFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EDIDataFormat
    EDIDataFormat = mvarEDIDataFormat
End Property



Public Property Let EDIName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EDIName = 5
    mvarEDIName = vData
End Property


Public Property Get EDIName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EDIName
    EDIName = mvarEDIName
End Property




Public Property Let CodeListID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeListID = 5
    mvarCodeListID = vData
End Property


Public Property Get CodeListID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeListID
    CodeListID = mvarCodeListID
End Property

Public Property Let Value(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Get Value() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property


Public Property Let NCTSDataFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataFormat = 5
    mvarNCTSDataFormat = vData
End Property


Public Property Get NCTSDataFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataFormat
    NCTSDataFormat = mvarNCTSDataFormat
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

Public Property Let NCTSName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarNCTSName = vData
End Property


Public Property Get NCTSName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    NCTSName = mvarNCTSName
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
