VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCPLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum FeaturesConstant
    [PLDA Import] = 2 ^ 0
    [PLDA Export] = 2 ^ 1
    [PLDA Combined] = 2 ^ 2
    [EDIFACT NCTS] = 2 ^ 3
    [SADBEL Import] = 2 ^ 4
    [SADBEL Export/OTS] = 2 ^ 5
    [SADBEL NCTS/Combined] = 2 ^ 6
    [SysLink Automatic Input] = 2 ^ 7
    [SysLink Semi-Automatic Input] = 2 ^ 8
    [SysLink Automatic Output] = 2 ^ 9
    [Email Automatic] = 2 ^ 10
    [Email Semi-Automatic] = 2 ^ 11
    [Entrepot] = 2 ^ 12
    [Repertory] = 2 ^ 13
    [Remote Print] = 2 ^ 14
    [Fixed Type License] = 2 ^ 15    'Server licensing type (checked = Fixed, unchecked = Floating)
    [EDIFACT NCTS FollowUpRequest] = 2 ^ 16
    [PDF Ouptput] = 2 ^ 17
    [Backup DB] = 2 ^ 18
End Enum

Private m_enuFeatures As FeaturesConstant
Private mvarLicenseType As String

Public Function IsFeatureLicensed(ByVal Feature As FeaturesConstant)
    IsFeatureLicensed = (m_enuFeatures And Feature)
End Function

Public Property Let LicenseType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LicenseType = 5
    mvarLicenseType = vData
End Property


Public Property Get LicenseType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LicenseType
    LicenseType = mvarLicenseType
End Property



