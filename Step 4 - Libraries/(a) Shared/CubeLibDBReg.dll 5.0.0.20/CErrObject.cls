VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErrObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNumber As Long 'local copy
Private mvarSource As String 'local copy
Private mvarDescription As String 'local copy
Private mvarHelpContext As Long 'local copy
Private mvarHelpFile As String 'local copy
Private mvarLastDllError As Long 'local copy
Private mvarErrorLine As Long 'local copy

Public Property Get ErrorLine() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorLine
    ErrorLine = mvarErrorLine
End Property




Public Property Get LastDllError() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastDllError
    LastDllError = mvarLastDllError
End Property

Public Sub CloneErrObject(ByRef e As ErrObject)
    mvarNumber = e.Number
    mvarLastDllError = e.LastDllError
    mvarSource = e.Source
    mvarDescription = e.Description
    mvarHelpContext = e.HelpContext
    mvarHelpFile = e.HelpFile
    mvarErrorLine = Erl
End Sub

Public Property Get HelpFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HelpFile
    HelpFile = mvarHelpFile
End Property

Public Property Get HelpContext() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HelpContext
    HelpContext = mvarHelpContext
End Property

Public Property Let Description(ByVal NewValue As String)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    mvarDescription = NewValue
End Property

Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property

Public Property Get Source() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    Source = mvarSource
End Property

Public Property Get Number() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number
    Number = mvarNumber
End Property


Friend Property Let Number(ByVal NewValue As Long)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number
    mvarNumber = NewValue
End Property

