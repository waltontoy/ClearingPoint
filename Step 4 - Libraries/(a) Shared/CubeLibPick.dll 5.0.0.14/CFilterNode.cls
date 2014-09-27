VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mvarIndex  As Integer
Dim mvarFilter As String
Public Key As String

Public Property Get Filter() As String
    Filter = mvarFilter
End Property

Public Property Get Index() As Integer
    Index = mvarIndex
End Property

Friend Property Let Index(ByVal vData As Integer)
    mvarIndex = vData
End Property

Public Property Let Filter(ByVal vData As String)
    mvarFilter = vData
End Property




