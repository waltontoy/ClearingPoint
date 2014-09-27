VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntrepotProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private blnCancelled As Boolean

Public Sub ShowProp(ByVal conSource, ByVal ResourceHandler As Long)
    
    blnCancelled = frmEntrepotProperties.LoadProp(conSource, ResourceHandler)
    
End Sub

Public Property Get PropCancelled() As Boolean

    PropCancelled = blnCancelled
    
End Property
