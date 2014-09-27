VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPickGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' CONSTANTS
'Const CPI_TRUE = -1
'Const CPI_FALSE = 0
'Const CPI_AUTOCANCEL = 1
'
' ENUMERATIONS
Public Enum ButtonType
    cpiAdd = 0
    cpiModify = 1
    cpiCopy = 2
    cpiDelete = 3
End Enum

Public Enum cpiGridSettingConstants
   cpiSaveGridSetting
   cpiRestoreGridSetting
End Enum

Public Enum cpiActiveStatusConstants
  cpiManyRecord
  cpiOneRecord
  cpiNotFound
  cpiOneRecordExact

End Enum

Public Enum PicklistStyle
    cpiSimplePicklist = 0
    cpiCatalog = 1
    cpiFilterCatalog = 2
End Enum
    
Public Enum cpiKeyPickConstants
   cpiKeyEnter = 0
   cpiKeyF2 = 1
   cpiKeyTabNoAction = 2
   cpiKeyTabEnter = 3
End Enum

Public Enum cpiListTypeConstants
   cpiUserList = 0
   cpiSimpleList = 1
End Enum


