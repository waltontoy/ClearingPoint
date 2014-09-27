VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGridSeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GridColumns"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
Private mvarGridColumns As CGridColumns

Public Property Get GridColumns() As CGridColumns
    If mvarGridColumns Is Nothing Then
        Set mvarGridColumns = New CGridColumns
    End If


    Set GridColumns = mvarGridColumns
End Property


Public Property Set GridColumns(vData As CGridColumns)
    Set mvarGridColumns = vData
End Property

Private Sub Class_Terminate()
    Set mvarGridColumns = Nothing
End Sub



