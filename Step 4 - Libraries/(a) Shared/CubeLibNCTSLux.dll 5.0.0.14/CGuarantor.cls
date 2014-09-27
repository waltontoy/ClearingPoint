VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuarantor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarNamGU20 As String
Private mvarStrAndNumGU621 As String
Private mvarPosCodGU622 As String
Private mvarCitGU623 As String
Private mvarCouGU624 As String
Private mvarNADLNGGU As String

'LET
Public Property Let NamGU20(ByVal vData As String)
    mvarNamGU20 = vData
End Property
Public Property Let StrAndNumGU621(ByVal vData As String)
    mvarStrAndNumGU621 = vData
End Property
Public Property Let PosCodGU622(ByVal vData As String)
    mvarPosCodGU622 = vData
End Property
Public Property Let CitGU623(ByVal vData As String)
    mvarCitGU623 = vData
End Property
Public Property Let CouGU624(ByVal vData As String)
    mvarCouGU624 = vData
End Property
Public Property Let NADLNGGU(ByVal vData As String)
    mvarNADLNGGU = vData
End Property

'GET
Public Property Get NamGU20() As String
    NamGU20 = mvarNamGU20
End Property
Public Property Get StrAndNumGU621() As String
    StrAndNumGU621 = mvarStrAndNumGU621
End Property
Public Property Get PosCodGU622() As String
    PosCodGU622 = mvarPosCodGU622
End Property
Public Property Get CitGU623() As String
    CitGU623 = mvarCitGU623
End Property
Public Property Get CouGU624() As String
    CouGU624 = mvarCouGU624
End Property
Public Property Get NADLNGGU() As String
    NADLNGGU = mvarNADLNGGU
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

