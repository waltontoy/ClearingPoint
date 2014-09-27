VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' cResults
' 2003/05/19 Copyright © 2003, Larry Rebich, using the DELL7500
' 2003/05/19 larry@larryrebich.com, www.larryrebich.com, 760-771-4730

Option Explicit
DefLng A-Z
'

Private miAdded         As Integer
Private miReplaced      As Integer
Private miSkipped       As Integer

Public Property Let iAdded(iAdded As Integer)
   miAdded = iAdded
End Property

Public Property Get iAdded() As Integer
   iAdded = miAdded
End Property

Public Property Let iReplaced(iReplaced As Integer)
   miReplaced = iReplaced
End Property

Public Property Get iReplaced() As Integer
   iReplaced = miReplaced
End Property

Public Property Let iSkipped(iSkipped As Integer)
   miSkipped = iSkipped
End Property

Public Property Get iSkipped() As Integer
   iSkipped = miSkipped
End Property

