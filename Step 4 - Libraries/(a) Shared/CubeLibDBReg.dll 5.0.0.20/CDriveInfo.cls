VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDriveInfo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetSpaceSerializedLogicalDrives() As String
    Dim LDs As Long
    Dim Cnt As Long
    Dim sDrives As String
    
    'get the available drives
    LDs = GetLogicalDrives
    sDrives = ""
    
    For Cnt = 0 To 25
        If (LDs And 2 ^ Cnt) <> 0 Then
            sDrives = sDrives + " " + Chr$(65 + Cnt)
        End If
    Next Cnt
    
    GetSpaceSerializedLogicalDrives = Trim(sDrives)
End Function

Public Function GetDriveTypeDescription(ByVal DriveToCheck As String) As String
    Select Case GetDriveType(DriveToCheck)
        Case DRIVE_TYPE_UNDTERMINED
            GetDriveTypeDescription = "Undetermined"
        Case DRIVE_ROOT_NOT_EXIST
            GetDriveTypeDescription = "Does Not Exist"
        Case DRIVE_REMOVABLE
            GetDriveTypeDescription = "Removable"
        Case DRIVE_FIXED
            GetDriveTypeDescription = "Drive Fixed"
        Case DRIVE_REMOTE
            GetDriveTypeDescription = "Remote/Network/Mapped"
        Case DRIVE_CDROM
            GetDriveTypeDescription = "Cd-Rom"
        Case DRIVE_RAMDISK
            GetDriveTypeDescription = "Ram disk"
        Case Else
            GetDriveTypeDescription = "Undetermined"
    End Select
End Function

