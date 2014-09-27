Attribute VB_Name = "MDriveInfo"
Option Explicit

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_TYPE_UNDTERMINED = 0
Public Const DRIVE_ROOT_NOT_EXIST = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
