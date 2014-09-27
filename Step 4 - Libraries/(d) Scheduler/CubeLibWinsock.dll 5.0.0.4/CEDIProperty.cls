VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEDIProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarQueueName As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPassword As String 'local copy
Private mvarRemoteHost As String 'local copy
Private mvarRemotePort As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDisabled As Boolean 'local copy
Public Property Let Disabled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Disabled = 5
    mvarDisabled = vData
End Property


Public Property Get Disabled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Disabled
    Disabled = mvarDisabled
End Property




Public Property Let RemotePort(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RremotePort = 5
    mvarRemotePort = vData
End Property


Public Property Get RemotePort() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RremotePort
    RemotePort = mvarRemotePort
End Property



Public Property Let RemoteHost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteHost = 5
    mvarRemoteHost = vData
End Property


Public Property Get RemoteHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteHost
    RemoteHost = mvarRemoteHost
End Property



Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property



Public Property Let QueueName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QueueName = 5
    mvarQueueName = vData
End Property


Public Property Get QueueName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QueueName
    QueueName = mvarQueueName
End Property



