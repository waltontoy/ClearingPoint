VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunerrer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy


Private mvarErrTypER11 As String
Private mvarErrPoiER12 As String
Private mvarErrReaER13 As String
Private mvarOriAttValER14 As String

'LET
Public Property Let ErrTypER11(ByVal vData As String)
    mvarErrTypER11 = vData
End Property
Public Property Let ErrPoiER12(ByVal vData As String)
    mvarErrPoiER12 = vData
End Property
Public Property Let ErrReaER13(ByVal vData As String)
    mvarErrReaER13 = vData
End Property
Public Property Let OriAttValER14(ByVal vData As String)
    mvarOriAttValER14 = vData
End Property
  
'GET
Public Property Get ErrTypER11() As String
    ErrTypER11 = mvarErrTypER11
End Property
Public Property Get ErrPoiER12() As String
    ErrPoiER12 = mvarErrPoiER12
End Property
Public Property Get ErrReaER13() As String
    ErrReaER13 = mvarErrReaER13
End Property
Public Property Get OriAttValER14() As String
    OriAttValER14 = mvarOriAttValER14
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
