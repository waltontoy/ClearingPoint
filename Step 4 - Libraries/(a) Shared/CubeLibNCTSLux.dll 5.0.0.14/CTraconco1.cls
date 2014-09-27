VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTraconco1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy

Private mvarNamCO17 As String
Private mvarStrAndNumCO122 As String
Private mvarPosCodCO123 As String
Private mvarCitCO124 As String
Private mvarCouCO125 As String
Private mvarNADLNGCO As String
Private mvarTINCO159 As String

'LET
Public Property Let NamCO17(ByVal vData As String)
    mvarNamCO17 = vData
End Property
Public Property Let StrAndNumCO122(ByVal vData As String)
    mvarStrAndNumCO122 = vData
End Property
Public Property Let PosCodCO123(ByVal vData As String)
    mvarPosCodCO123 = vData
End Property
Public Property Let CitCO124(ByVal vData As String)
    mvarCitCO124 = vData
End Property
Public Property Let CouCO125(ByVal vData As String)
    mvarCouCO125 = vData
End Property
Public Property Let NADLNGCO(ByVal vData As String)
    mvarNADLNGCO = vData
End Property
Public Property Let TINCO159(ByVal vData As String)
    mvarTINCO159 = vData
End Property

'GET
Public Property Get NamCO17() As String
    NamCO17 = mvarNamCO17
End Property
Public Property Get StrAndNumCO122() As String
    StrAndNumCO122 = mvarStrAndNumCO122
End Property
Public Property Get PosCodCO123() As String
    PosCodCO123 = mvarPosCodCO123
End Property
Public Property Get CitCO124() As String
    CitCO124 = mvarCitCO124
End Property
Public Property Get CouCO125() As String
    CouCO125 = mvarCouCO125
End Property
Public Property Get NADLNGCO() As String
    NADLNGCO = mvarNADLNGCO
End Property
Public Property Get TINCO159() As String
    TINCO159 = mvarTINCO159
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

