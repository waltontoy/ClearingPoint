VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiIE07Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarSetResourceHandler As Long

Public Function InitializeMenus(ByRef tbrIE07 As Object, ByVal AppName As String, ByVal AppPath As String)
    G_InitializeMenus tbrIE07, AppName, AppPath
End Function



Public Property Get SetResourceHandler() As Long
    SetResourceHandler = mvarSetResourceHandler
End Property

Public Property Let SetResourceHandler(ByVal vNewValue As Long)
    mvarSetResourceHandler = vNewValue
    ResourceHandler = vNewValue
End Property

Public Function SearchRecordInDB(ByRef ActiveConnection As ADODB.Connection, _
                                                    ByVal TableName, _
                                                    ByVal SearchField, _
                                                    ByVal SearchValue) As Boolean
'
    SearchRecordInDB = SearchRecord(ActiveConnection, TableName, SearchField, SearchValue)

End Function

'''''Public Function FNullField(ByRef ActiveField As ADODB.Field) As Variant
'''''    CheckNull = G_FNullField(ActiveField)
'''''End Function

