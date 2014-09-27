VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGridColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

    Public Key As String
    
    Private mvarDataType As Variant
    Private mvarColumnWidth As Double
    Private mvarColumnAlignment As Variant
    Private mvarColumnFieldAias As String
    Private mvarColumnBaseFieldName As String
    Private mvarFormat As String

Public Property Let Format(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnFieldAias = 5
    mvarFormat = vData
End Property


Public Property Get Format() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnFieldAias
    Format = mvarFormat
End Property


Public Property Let ColumnBaseFieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnFieldAias = 5
    mvarColumnBaseFieldName = vData
End Property


Public Property Get ColumnBaseFieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnFieldAias
    ColumnBaseFieldName = mvarColumnBaseFieldName
End Property

Public Property Let ColumnFieldAias(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnFieldAias = 5
    mvarColumnFieldAias = vData
End Property


Public Property Get ColumnFieldAias() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnFieldAias
    ColumnFieldAias = mvarColumnFieldAias
End Property



Public Property Let ColumnAlignment(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnAlignment = 5
    mvarColumnAlignment = vData
End Property


Public Property Set ColumnAlignment(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ColumnAlignment = Form1
    Set mvarColumnAlignment = vData
End Property


Public Property Get ColumnAlignment() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnAlignment
    If IsObject(mvarColumnAlignment) Then
        Set ColumnAlignment = mvarColumnAlignment
    Else
        ColumnAlignment = mvarColumnAlignment
    End If
End Property



Public Property Let ColumnWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnWidth = 5
    mvarColumnWidth = vData
End Property


Public Property Get ColumnWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnWidth
    ColumnWidth = mvarColumnWidth
End Property



Public Property Let DataType(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataType = 5
    mvarDataType = vData
End Property


Public Property Set DataType(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DataType = Form1
    Set mvarDataType = vData
End Property


Public Property Get DataType() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataType
    If IsObject(mvarDataType) Then
        Set DataType = mvarDataType
    Else
        DataType = mvarDataType
    End If
End Property



