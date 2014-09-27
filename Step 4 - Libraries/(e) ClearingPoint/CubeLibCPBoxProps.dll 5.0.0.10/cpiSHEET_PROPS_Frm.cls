VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSHEET_PROPS_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CLINK"
Attribute VB_Ext_KEY = "Member0" ,"CLINK"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Key As String
Public Index As Long

Dim mvarBoxDefaultAdminTable As String ' 1
Dim mvarBoxDefaultValueTable As String ' 2
Dim mvarDefaultUserTable As String ' 3
Dim mvarCodisheetType As cpiCodiSheetTypeEnums ' 4
Dim mvarActiveBoxCode As String ' 5
Dim mvarUserNo As Long ' 6

' 1
Public Property Get BoxDefaultAdminTable() As String
    BoxDefaultAdminTable = mvarBoxDefaultAdminTable
End Property

Public Property Let BoxDefaultAdminTable(ByVal vNewValue As String)
    mvarBoxDefaultAdminTable = vNewValue
End Property

' 2
Public Property Get BoxDefaultValueTable() As String
    BoxDefaultValueTable = mvarBoxDefaultValueTable
End Property

Public Property Let BoxDefaultValueTable(ByVal vNewValue As String)
    mvarBoxDefaultValueTable = vNewValue
End Property

' 3
Public Property Get DefaultUserTable() As String
    DefaultUserTable = mvarDefaultUserTable
End Property

Public Property Let DefaultUserTable(ByVal vNewValue As String)
    mvarDefaultUserTable = vNewValue
End Property

' 4
Public Property Get CodisheetType() As cpiCodiSheetTypeEnums
    CodisheetType = mvarCodisheetType
End Property

Public Property Let CodisheetType(ByVal vNewValue As cpiCodiSheetTypeEnums)
    mvarCodisheetType = vNewValue
End Property

' 5
Public Property Get ActiveBoxCode() As String
    ActiveBoxCode = mvarActiveBoxCode
End Property

Public Property Let ActiveBoxCode(ByVal vNewValue As String)
    mvarActiveBoxCode = vNewValue
End Property

' 6
Public Property Get UserNo() As Long
    UserNo = mvarUserNo
End Property

Public Property Let UserNo(ByVal vNewValue As Long)
    mvarUserNo = vNewValue
End Property

' functions here
Public Function ShowForm(ByRef OwnerForm As Object, _
                                            ByRef ActiveConnection As ADODB.Connection, _
                                            ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                            ByRef UserNo As Long, _
                                            ByRef ActiveLanguage As String, _
                                            ByRef ResourceHandle As Long) As Boolean
                                            

    Dim clsForm As frmSheetProperties
    
    Set clsForm = New frmSheetProperties

    ' load frmBoxProperties here
    clsForm.ShowForm OwnerForm, ActiveConnection, ActiveCodisheet, _
                            UserNo, ActiveLanguage, ResourceHandle
    
    Unload clsForm
    Set clsForm = Nothing

End Function



