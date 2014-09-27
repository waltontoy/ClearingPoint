VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMachinePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

    Public Key As String
    
    Private mvarRegistrySource As String
    Private mvarPathType As String
    Private mvarRegistryKey As String
    Private mvarRegistrySetting As String
    Private mvarDisplayName As String
    Private mvarDBType As String
    Private mvarDBName As String

Friend Property Let DBName(vData As String)
    mvarDBName = vData
End Property

Public Property Get DBName() As String
    DBName = mvarDBName
End Property


Friend Property Let DBType(vData As String)
    mvarDBType = vData
End Property

Public Property Get DBType() As String
    DBType = mvarDBType
End Property


Friend Property Let DisplayName(vData As String)
    mvarDisplayName = vData
End Property

Public Property Get DisplayName() As String
    DisplayName = mvarDisplayName
End Property


Friend Property Let RegistrySetting(vData As String)
    mvarRegistrySetting = vData
End Property

Public Property Get RegistrySetting() As String
    RegistrySetting = mvarRegistrySetting
End Property


Friend Property Let RegistryKey(vData As String)
    mvarRegistryKey = vData
End Property

Public Property Get RegistryKey() As String
    RegistryKey = mvarRegistryKey
End Property


Friend Property Let PathType(vData As String)
    mvarPathType = vData
End Property

Public Property Get PathType() As String
    PathType = mvarPathType
End Property
