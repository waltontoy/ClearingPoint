Attribute VB_Name = "MNetowrkInfo"
Option Explicit

Public Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long
    
Public Const NETWORK_ALIVE_AOL = &H4
Public Const NETWORK_ALIVE_LAN = &H1
Public Const NETWORK_ALIVE_WAN = &H2
