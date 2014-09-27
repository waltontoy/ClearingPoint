VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistryGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Function HasRegistryAdminRights(ByVal RegistryGroup As TargetRegistry, _
                                          ByVal RegistryPath As String, _
                                          ByVal RegistryKey As String, _
                                          ByVal RegistrySetting As String, _
                                          HasAdminRights As Boolean) As String
    
    HasRegistryAdminRights = HasRegistryAdminRights_F(RegistryGroup, _
                                                          RegistryPath, _
                                                          RegistryKey, _
                                                          RegistrySetting, _
                                                          HasAdminRights)
End Function


