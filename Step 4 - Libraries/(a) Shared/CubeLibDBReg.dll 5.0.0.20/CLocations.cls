VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent AdvanceClick[(arg1, arg2, ... , argn)]
Public Event AdvanceClick()


Public Function ViewLocations(ByRef OwnerForm As Object, ByVal Application As Object, ByVal UseAdvanced As Boolean, ParamArray PathArgs()) As Boolean
    Dim clsRegistry As CRegistry
    Dim lngArgCtr  As Long
    Dim arrSubArguments
    Dim strPaths As String
    Dim strRegistryValue As String

    Set clsRegistry = New CRegistry
    
    strPaths = ""
    
    For lngArgCtr = 0 To UBound(PathArgs)
        arrSubArguments = Split(CStr(PathArgs(lngArgCtr)), "|")
        
        clsRegistry.GetRegistry cpiLocalMachine, Application.ProductName, CStr(arrSubArguments(1)), CStr(arrSubArguments(2))
        strRegistryValue = Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT)
        
        Select Case UCase(CStr(arrSubArguments(0)))
            Case "P", "D"
                strPaths = strPaths & "|||||" & CStr(PathArgs(lngArgCtr))
            Case "F"
                If Trim(strRegistryValue) <> "" And Len(strRegistryValue) > 3 Then
                    If Mid(strRegistryValue, Len(strRegistryValue) - 3, 1) = "." Then
                        If Not PathFileExists(strRegistryValue) Then
                            strPaths = strPaths & "|||||" & CStr(PathArgs(lngArgCtr))
                        End If
                    End If
                Else
                    strPaths = strPaths & "|||||" & CStr(PathArgs(lngArgCtr))
                End If
        End Select
    Next lngArgCtr

    If Left(strPaths, 5) = "|||||" Then
        strPaths = Mid(strPaths, 6)
    End If
    
    ViewLocations = frmMissingPathsList.ShowMissingPaths(OwnerForm, Application, strPaths, Me, UseAdvanced)

    Set clsRegistry = Nothing

End Function

Friend Sub TriggerAdvanced()
    RaiseEvent AdvanceClick
End Sub

