VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrapripc1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy

Private mvarNamPC17 As String
Private mvarStrAndNumPC122 As String
Private mvarPosCodPC123 As String
Private mvarCitPC124 As String
Private mvarCouPC125 As String
Private mvarNADLNGPC As String
Private mvarTINPC159 As String

'LET
Public Property Let NamPC17(ByVal vData As String)
    mvarNamPC17 = vData
End Property
Public Property Let StrAndNumPC122(ByVal vData As String)
    mvarStrAndNumPC122 = vData
End Property
Public Property Let PosCodPC123(ByVal vData As String)
    mvarPosCodPC123 = vData
End Property
Public Property Let CitPC124(ByVal vData As String)
    mvarCitPC124 = vData
End Property
Public Property Let CouPC125(ByVal vData As String)
    mvarCouPC125 = vData
End Property
Public Property Let NADLNGPC(ByVal vData As String)
    mvarNADLNGPC = vData
End Property
Public Property Let TINPC159(ByVal vData As String)
    mvarTINPC159 = vData
End Property

'GET
Public Property Get NamPC17() As String
    NamPC17 = mvarNamPC17
End Property
Public Property Get StrAndNumPC122() As String
    StrAndNumPC122 = mvarStrAndNumPC122
End Property
Public Property Get PosCodPC123() As String
    PosCodPC123 = mvarPosCodPC123
End Property
Public Property Get CitPC124() As String
    CitPC124 = mvarCitPC124
End Property
Public Property Get CouPC125() As String
    CouPC125 = mvarCouPC125
End Property
Public Property Get NADLNGPC() As String
    NADLNGPC = mvarNADLNGPC
End Property
Public Property Get TINPC159() As String
    TINPC159 = mvarTINPC159
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

