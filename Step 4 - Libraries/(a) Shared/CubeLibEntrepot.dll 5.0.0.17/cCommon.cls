VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eUpdateInbound
    eUpdate_None = 0
    eUpdate_FromOtherFolderToOutbox = 1
    eUpdate_FromOutboxToOutbox = 2
    eUpdate_ToDeleted = 3
    eUpdate_ToRejected = 4
    
End Enum

'local variable(s) to hold property value(s)
Private mvarActiveLanguage As String 'local copy
Public Property Let ActiveLanguage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ActiveLanguage = 5
    mvarActiveLanguage = vData
End Property


Public Property Get ActiveLanguage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActiveLanguage
    ActiveLanguage = mvarActiveLanguage
End Property

