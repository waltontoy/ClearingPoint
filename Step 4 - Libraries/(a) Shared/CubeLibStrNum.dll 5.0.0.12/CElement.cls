VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CElement"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' ************************************************************************************** '
' Class Filename        : clsElement.cls
' Classname             : clsElement
' No of used classes    : 0
' No of Properties      : 3
'   Element             : String (Get/Let)
'   ElementType         : String (Get/Let)
'   ElementValue        : Double (Get/Let)
' No of Methods         : 0
' No of Events          : 0
' Author                : Jason B. Sia
' Company               : Cubepoint, Inc.
' Revisions             : 1
' Revised Date          : May 29, 2002
' Description           : Use by other classes in creating list of arrays
' ************************************************************************************** '

Option Explicit

Private mvarElement As String
Private mvarElementType As String
Private mvarElementValue As Double

Public Property Get Element() As String
    Element = mvarElement
End Property

Public Property Let Element(ByVal vData As String)
    mvarElement = vData
End Property

Public Property Get ElementType() As String
    ElementType = mvarElementType
End Property

Public Property Let ElementType(ByVal vData As String)
    mvarElementType = vData
End Property

Public Property Get ElementValue() As Double
    ElementValue = mvarElementValue
End Property

Public Property Let ElementValue(ByVal vData As Double)
    mvarElementValue = vData
End Property






