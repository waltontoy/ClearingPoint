VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCusoffdepept"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy

Private mvarRefNumEPT1 As String

Public Property Let RefNumEPT1(ByVal vData As String)
    mvarRefNumEPT1 = vData
End Property

Public Property Get RefNumEPT1() As String
    RefNumEPT1 = mvarRefNumEPT1
End Property


'Public Property Set DateTimeOfPreparation(ByVal vdata As CDateTimeElement)
''used when assigning an Object to the property, on the left side of a Set statement.
''Syntax: Set x.DateTimeOfPreparation = Form1
'    Set mvarDateTimeOfPreparation = vdata
'End Property
'
'
'Public Property Get DateTimeOfPreparation() As CDateTimeElement
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.DateTimeOfPreparation
'    Set DateTimeOfPreparation = mvarDateTimeOfPreparation
'End Property
'
'
'
'Public Property Let MessageSender(ByVal vdata As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.MessageSender = 5
'    mvarMessageSender = vdata
'End Property
'
'
'Public Property Get MessageSender() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.MessageSender
'    MessageSender = mvarMessageSender
'End Property
'
'
'
'
