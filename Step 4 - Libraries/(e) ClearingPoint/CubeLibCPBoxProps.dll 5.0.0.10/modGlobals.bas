Attribute VB_Name = "modGlobals"
Option Explicit

Public Const PCK_TIME = "8.98888888209135E+19"
Public Const PCK_DATE = "8.98888888841205E+19"



Public Function GetInternalCode(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveDocument As String, _
                                ByRef ActiveBoxCode As String) As String
    '
    Dim strSql As String
    Dim clsPICKLIST_DEFINITION As cpiPICK_DEF
    Dim clsPICKLIST_DEFINITIONs As cpiPICK_DEFs
    
    Set clsPICKLIST_DEFINITION = New cpiPICK_DEF
    Set clsPICKLIST_DEFINITIONs = New cpiPICK_DEFs
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [PICKLIST DEFINITION] "
    strSql = strSql & " WHERE [DOCUMENT]='" & ActiveDocument & "'"
    strSql = strSql & " AND [BOX CODE]='" & ActiveBoxCode & "'"
    
    Dim rstPicklistDefinitions As ADODB.Recordset
    ADORecordsetOpen strSql, ActiveConnection, rstPicklistDefinitions, adOpenKeyset, adLockOptimistic
    
    Set clsPICKLIST_DEFINITIONs.Recordset = rstPicklistDefinitions
    
    If (clsPICKLIST_DEFINITIONs.Recordset.EOF = False) Then
        Set clsPICKLIST_DEFINITION = clsPICKLIST_DEFINITIONs.GetClassRecord(clsPICKLIST_DEFINITIONs.Recordset)
        GetInternalCode = clsPICKLIST_DEFINITION.FIELD_INTERNAL_CODE
    End If
    
    Set clsPICKLIST_DEFINITION = Nothing
    Set clsPICKLIST_DEFINITIONs = Nothing
'
End Function



Public Sub LoadResStrings(ByRef frmFormToLoad As Form, ByVal lngHandler As Long)
    Dim ctlControlToLoad As Control
    Dim strTypeName As String
    Dim intCtrlCount As Integer
    Dim i As Integer
        
    On Error Resume Next
    
    
    For Each ctlControlToLoad In frmFormToLoad.Controls
        strTypeName = LCase(TypeName(ctlControlToLoad))
        
        Select Case strTypeName
            Case "sstab"
                intCtrlCount = ctlControlToLoad.Tabs
                                
                For i = 0 To intCtrlCount - 1
                    ctlControlToLoad.TabCaption(i) = Translate_B(ctlControlToLoad.TabCaption(i), lngHandler)
                Next
            Case "tabstrip"
                intCtrlCount = ctlControlToLoad.Tabs.Count
                For i = 1 To intCtrlCount
                    If (ctlControlToLoad.Tabs(i).Caption <> "") Then
                        ctlControlToLoad.Tabs(i).Caption = Translate_B(ctlControlToLoad.Tabs(i).Caption, lngHandler)
                    End If
                Next
            Case "label", "optionbutton", "frame", "commandbutton", "sscommand", "sspanel", "checkbox"
                If (Trim$(ctlControlToLoad.Tag) <> "") Then
                    ctlControlToLoad.Caption = Translate_B(ctlControlToLoad.Tag, lngHandler)
                End If

        End Select
    Next
End Sub

