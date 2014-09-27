VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMachinePaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cpiResourcePath"
Attribute VB_Ext_KEY = "Member0" ,"cpiResourcePath"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Private mCol As Collection

    Public Event BeforeRestart(ByVal NewDBLocation As String, ByRef Cancel As Boolean)
    Public Event RegistryChanged(ByRef NewPath As String)
    
Public Function LoadPaths(ParamArray MachinePaths())
    Dim objNewMember As CMachinePath
    
    Dim lngMachPathCtr As Long
    Dim arrMachinePathDetails
    
1000    For lngMachPathCtr = 0 To UBound(MachinePaths)
    
1010        Set objNewMember = New CMachinePath
        
1020        arrMachinePathDetails = Split(CStr(MachinePaths(lngMachPathCtr)), "|")
        
        ' Set the Properties Passed Into the Method
1030        objNewMember.Key = CStr(MachinePaths(lngMachPathCtr))
1040        Select Case CStr(arrMachinePathDetails(0))
            Case "F"
1060                objNewMember.PathType = "File"
            Case "P"
1080                objNewMember.PathType = "Folder"
            Case "D"
1100                objNewMember.PathType = "Database"
1110        End Select
1120        objNewMember.RegistryKey = CStr(arrMachinePathDetails(1))
1130        objNewMember.RegistrySetting = CStr(arrMachinePathDetails(2))
1140        objNewMember.DisplayName = CStr(arrMachinePathDetails(3))
1150        If CStr(arrMachinePathDetails(0)) = "D" Then
1160            objNewMember.DBType = CStr(arrMachinePathDetails(4))
1170        Else
1180            objNewMember.DBType = ""
1190        End If
1200        If CStr(arrMachinePathDetails(0)) = "D" Then
1210            objNewMember.DBName = CStr(arrMachinePathDetails(5))
1220        Else
1230            objNewMember.DBName = ""
1240        End If
        
1250        mCol.Add objNewMember, CStr(MachinePaths(lngMachPathCtr))
                
1260        Set objNewMember = Nothing
1270    Next lngMachPathCtr
End Function

Public Property Get Item(vntIndexKey As Variant) As CMachinePath
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Friend Sub RestartEvents(ByVal DBLocation As String, ByRef Cancel As Boolean)

    RaiseEvent BeforeRestart(DBLocation, Cancel)
    
End Sub

Friend Sub TriggerRegistryChanged(NewDBPath As String)

    RaiseEvent RegistryChanged(NewDBPath)
    
End Sub
