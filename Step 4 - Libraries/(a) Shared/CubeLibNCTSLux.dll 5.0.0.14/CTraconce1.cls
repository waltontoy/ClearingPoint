VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTraconce1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy

Private mvarNamCE17 As String
Private mvarStrAndNumCE122 As String
Private mvarPosCodCE123 As String
Private mvarCitCE124 As String
Private mvarCouCE125 As String
Private mvarNADLNGCE As String
Private mvarTINCE159 As String

'LET
Public Property Let NamCE17(ByVal vData As String)
    mvarNamCE17 = vData
End Property
Public Property Let StrAndNumCE122(ByVal vData As String)
    mvarStrAndNumCE122 = vData
End Property
Public Property Let PosCodCE123(ByVal vData As String)
    mvarPosCodCE123 = vData
End Property
Public Property Let CitCE124(ByVal vData As String)
    mvarCitCE124 = vData
End Property
Public Property Let CouCE125(ByVal vData As String)
    mvarCouCE125 = vData
End Property
Public Property Let NADLNGCE(ByVal vData As String)
    mvarNADLNGCE = vData
End Property
Public Property Let TINCE159(ByVal vData As String)
    mvarTINCE159 = vData
End Property

'GET
Public Property Get NamCE17() As String
    NamCE17 = mvarNamCE17
End Property
Public Property Get StrAndNumCE122() As String
    StrAndNumCE122 = mvarStrAndNumCE122
End Property
Public Property Get PosCodCE123() As String
    PosCodCE123 = mvarPosCodCE123
End Property
Public Property Get CitCE124() As String
    CitCE124 = mvarCitCE124
End Property
Public Property Get CouCE125() As String
    CouCE125 = mvarCouCE125
End Property
Public Property Get NADLNGCE() As String
    NADLNGCE = mvarNADLNGCE
End Property
Public Property Get TINCE159() As String
    TINCE159 = mvarTINCE159
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

