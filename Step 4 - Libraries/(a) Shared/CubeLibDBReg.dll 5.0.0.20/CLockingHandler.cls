VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLockingHandler"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum enumLockingProcedure
    enumLockAdd = 0
    enumLockModify = 1
    enumLockDelete = 2
End Enum

'lngUpdateType => 0 - Add ; 1 - Modify ; 2 - Delete
Public Sub RecordLockingWorkAround(ByRef rstToBeUpdated As ADODB.Recordset, ByVal lngUpdateType As enumLockingProcedure)

    Dim blnLocked As Boolean
    Dim aFieldValues() As Variant
    Dim lngCtr As Long
    Dim lngTimeStart As Long

    On Error GoTo LockingError
    
    lngTimeStart = Timer()
    
    Select Case lngUpdateType
        
        Case enumLockingProcedure.enumLockAdd
            blnLocked = False
            ReDim aFieldValues(0)
            For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
                ReDim Preserve aFieldValues(lngCtr)
                aFieldValues(lngCtr) = rstToBeUpdated.Fields(lngCtr).Value
            Next
            rstToBeUpdated.Update
            Do While (blnLocked And Abs(CLng(Timer) - lngTimeStart) <= 10)
                rstToBeUpdated.AddNew
                blnLocked = False
                For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
                    If rstToBeUpdated.Fields(lngCtr).Properties("IsAutoIncrement").Value = False Then
                        rstToBeUpdated.Fields(lngCtr).Value = aFieldValues(lngCtr)
                    End If
                Next
                rstToBeUpdated.Update
            Loop
        Case enumLockingProcedure.enumLockModify
            blnLocked = False
            ReDim aFieldValues(0)
            For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
                ReDim Preserve aFieldValues(lngCtr)
                aFieldValues(lngCtr) = rstToBeUpdated.Fields(lngCtr).Value
            Next
            rstToBeUpdated.Update
            Do While (blnLocked And Abs(CLng(Timer) - lngTimeStart) <= 10)
                blnLocked = False
                For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
                    If rstToBeUpdated.Fields(lngCtr).Properties("IsAutoIncrement").Value = False Then
                        rstToBeUpdated.Fields(lngCtr).Value = aFieldValues(lngCtr)
                    End If
                Next
                rstToBeUpdated.Update
            Loop
        Case enumLockingProcedure.enumLockDelete
            blnLocked = True
            Do While blnLocked And (Abs(CLng(Timer) - lngTimeStart) <= 10)
                blnLocked = False
                rstToBeUpdated.Delete
            Loop
    End Select
    
    On Error GoTo 0
    Exit Sub
    
LockingError:
    If Err.Number = -2147217887 Then
        rstToBeUpdated.CancelUpdate
        blnLocked = True
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

''''''lngUpdateType => 0 - Add ; 1 - Modify ; 2 - Delete
'''''Public Sub RecordLockingWorkAroundForDAO(ByRef rstToBeUpdated As DAO.Recordset, ByVal lngUpdateType As enumLockingProcedure)
'''''
'''''    Dim blnLocked As Boolean
'''''    Dim aFieldValues() As Variant
'''''    Dim lngCtr As Long
'''''    Dim lngTimeStart As Long
'''''
'''''    On Error GoTo LockingError
'''''
'''''    lngTimeStart = Timer()
'''''
'''''    Select Case lngUpdateType
'''''
'''''        Case enumLockingProcedure.enumLockAdd
'''''            blnLocked = False
'''''            ReDim aFieldValues(0)
'''''            For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
'''''                ReDim Preserve aFieldValues(lngCtr)
'''''                aFieldValues(lngCtr) = rstToBeUpdated.Fields(lngCtr).Value
'''''            Next
'''''            rstToBeUpdated.Update
'''''            Do While (blnLocked And Abs(CLng(Timer) - lngTimeStart) <= 10)
'''''                rstToBeUpdated.AddNew
'''''                blnLocked = False
'''''                For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
'''''                    If Not (rstToBeUpdated.Fields(lngCtr).Attributes = 49 And rstToBeUpdated.Fields(lngCtr).Type = 4) Then
'''''                        rstToBeUpdated.Fields(lngCtr).Value = aFieldValues(lngCtr)
'''''                    End If
'''''                Next
'''''                rstToBeUpdated.Update
'''''            Loop
'''''        Case enumLockingProcedure.enumLockModify
'''''            blnLocked = False
'''''            ReDim aFieldValues(0)
'''''            For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
'''''                ReDim Preserve aFieldValues(lngCtr)
'''''                aFieldValues(lngCtr) = rstToBeUpdated.Fields(lngCtr).Value
'''''            Next
'''''            rstToBeUpdated.Update
'''''            Do While (blnLocked And Abs(CLng(Timer) - lngTimeStart) <= 10)
'''''                blnLocked = False
'''''                For lngCtr = 0 To rstToBeUpdated.Fields.Count - 1
'''''                    If Not (rstToBeUpdated.Fields(lngCtr).Attributes = 49 And rstToBeUpdated.Fields(lngCtr).Type = 4) Then
'''''                        rstToBeUpdated.Fields(lngCtr).Value = aFieldValues(lngCtr)
'''''                    End If
'''''                Next
'''''                rstToBeUpdated.Update
'''''            Loop
'''''        Case enumLockingProcedure.enumLockDelete
'''''            blnLocked = True
'''''            Do While blnLocked And (Abs(CLng(Timer) - lngTimeStart) <= 10)
'''''                blnLocked = False
'''''                rstToBeUpdated.Delete
'''''            Loop
'''''    End Select
'''''
'''''    On Error GoTo 0
'''''    Exit Sub
'''''
'''''LockingError:
'''''    If Err.Number = -2147217887 Then
'''''        rstToBeUpdated.CancelUpdate
'''''        blnLocked = True
'''''        Resume Next
'''''    Else
'''''        Err.Raise Err.Number, Err.Source, Err.Description
'''''    End If
'''''End Sub


