VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNetworkInfo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum NetworkStatusConstant
    [No Network]
    [AOL Network]
    [LAN Network]
    [WAN Network]
    [Undetermined]
End Enum

Public Function GetNetworkStatus() As NetworkStatusConstant
    
    Dim enuNetworkStatus As NetworkStatusConstant
    Dim lngReturnNetworkConnection As Long
    
    On Error GoTo ERR_HANDLER
    
    IsNetworkAlive lngReturnNetworkConnection
    
    Select Case lngReturnNetworkConnection
        Case 0
            enuNetworkStatus = NetworkStatusConstant.[No Network]
        Case NETWORK_ALIVE_AOL
            enuNetworkStatus = NetworkStatusConstant.[AOL Network]
        Case NETWORK_ALIVE_LAN
            enuNetworkStatus = NetworkStatusConstant.[LAN Network]
        Case Else
            enuNetworkStatus = NetworkStatusConstant.[WAN Network]
    End Select

ERR_HANDLER:

    Select Case Err.Number
        Case 0
        Case Else
            enuNetworkStatus = NetworkStatusConstant.Undetermined
    End Select
    
    GetNetworkStatus = enuNetworkStatus
End Function
