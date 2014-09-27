VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Connections"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarConnections As clsConnections 'local copy
Public Event Response(ResponseType As enuResponse, MessageInfo As String)
Private Sub mvarConnections_Response(ResponseType As enuResponse, MessageInfo As String)

    RaiseEvent Response(ResponseType, MessageInfo)
End Sub

Public Property Set Connections(ByVal vData As clsConnections)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Connections = Form1
    Set mvarConnections = vData
End Property


Public Property Get Connections() As clsConnections
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Connections
    Set Connections = mvarConnections
End Property

Private Sub Class_Initialize()
    Set mvarConnections = New clsConnections
End Sub
