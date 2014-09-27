VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSendValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarHeaderDetailNum As String 'local copy
Private mvarBoxCode As String 'local copy
Private mvarDescription As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSolution As String 'local copy
Public Property Let Solution(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Solution = 5
    mvarSolution = vData
End Property


Public Property Get Solution() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Solution
    Solution = mvarSolution
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let BoxCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoxCode = 5
    mvarBoxCode = vData
End Property


Public Property Get BoxCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoxCode
    BoxCode = mvarBoxCode
End Property



Public Property Let HeaderDetailNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeaderDetailNum = 5
    mvarHeaderDetailNum = vData
End Property


Public Property Get HeaderDetailNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeaderDetailNum
    HeaderDetailNum = mvarHeaderDetailNum
End Property



