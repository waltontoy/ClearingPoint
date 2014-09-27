VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBOX_PROPS_Fxs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarActiveType  As String
Private mvarActiveDocument   As String
Private mvarBoxDefaultAdminTable  As String
Private mvarBoxDefaultValueTable  As String
Private mvarDefaultUserTable  As String

Public Function GetEmptyBoxValue(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String) As String
    '
    InitCodisheetType ActiveCodisheet
    '
    Dim clsBoxDefaultAdmins As cpiBOX_DEF_ADMINs
    Dim clsBoxDefaultAdmin As cpiBOX_DEF_ADMIN
    ' Dim strInternalCode As String
    '
    Set clsBoxDefaultAdmins = New cpiBOX_DEF_ADMINs
    Set clsBoxDefaultAdmin = New cpiBOX_DEF_ADMIN
    '
    clsBoxDefaultAdmins.SetSqlParameters mvarBoxDefaultAdminTable
    '
    ' open general admin box here
    clsBoxDefaultAdmin.FIELD_BOX_CODE = ActiveBoxCode
    '
    clsBoxDefaultAdmins.GetRecord ActiveConnection, clsBoxDefaultAdmin
    '
    ' check empty box value here
    GetEmptyBoxValue = clsBoxDefaultAdmin.FIELD_EMPTY_FIELD_VALUE
    '
    Set clsBoxDefaultAdmin = Nothing
    Set clsBoxDefaultAdmins = Nothing
    '
End Function

Public Function GetGeneralDefaultValue(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef ActiveLogicalID As String, _
                                ByRef ActiveUserNo As Long) As String
    '
    Dim strSql As String
    Dim clsBOX_DEFAULT_VALUE As cpiBOX_DEF_VAL
    Dim clsBOX_DEFAULT_VALUEs As cpiBOX_DEF_VALs
    '
    InitCodisheetType ActiveCodisheet
    '
    strSql = "SELECT * "
    strSql = strSql & "FROM [" & mvarBoxDefaultValueTable & "]"
    strSql = strSql & "WHERE [BOX CODE]='" & ActiveBoxCode & "'"
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(ActiveUserNo) & ""
    strSql = strSql & " AND [LOGID DESCRIPTION]='" & CStr(ActiveLogicalID) & "'"
    '
    Set clsBOX_DEFAULT_VALUE = New cpiBOX_DEF_VAL
    Set clsBOX_DEFAULT_VALUEs = New cpiBOX_DEF_VALs
    '
    Set clsBOX_DEFAULT_VALUEs.Recordset = ActiveConnection.Execute(strSql)
    '
    If (clsBOX_DEFAULT_VALUEs.Recordset.EOF = False) Then
        Set clsBOX_DEFAULT_VALUE = clsBOX_DEFAULT_VALUEs.GetClassRecord(clsBOX_DEFAULT_VALUEs.Recordset)
        GetGeneralDefaultValue = clsBOX_DEFAULT_VALUE.FIELD_DEFAULT_VALUE
    ElseIf (clsBOX_DEFAULT_VALUEs.Recordset.EOF = True) Then
        GetGeneralDefaultValue = ""
    End If
    '
    Set clsBOX_DEFAULT_VALUE = Nothing
    Set clsBOX_DEFAULT_VALUEs = Nothing
    '
End Function

'
Public Function GetUserDefaultValue(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef ActiveLogicalID As String, _
                                ByRef ActiveUserNo As Long) As String
    '
    Dim strSql As String
    Dim clsDEFAULT_USER As cpiDEFAULT_USER
    Dim clsDEFAULT_USERs As cpiDEFAULT_USERs
    '
    InitCodisheetType ActiveCodisheet
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [" & mvarDefaultUserTable & "]"
    strSql = strSql & "WHERE [BOX CODE]='" & ActiveBoxCode & "'"
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(ActiveUserNo) & ""
    strSql = strSql & " AND [LOGID DESCRIPTION]='" & CStr(ActiveLogicalID) & "'"
    '
    Set clsDEFAULT_USER = New cpiDEFAULT_USER
    Set clsDEFAULT_USERs = New cpiDEFAULT_USERs
    '
    Set clsDEFAULT_USERs.Recordset = ActiveConnection.Execute(strSql)
    '
    If (clsDEFAULT_USERs.Recordset.EOF = False) Then
        Set clsDEFAULT_USER = clsDEFAULT_USERs.GetClassRecord(clsDEFAULT_USERs.Recordset)
        GetUserDefaultValue = clsDEFAULT_USER.FIELD_DEFAULT_VALUE
    ElseIf (clsDEFAULT_USERs.Recordset.EOF = True) Then
        GetUserDefaultValue = ""
    End If
    
    Set clsDEFAULT_USER = Nothing
    Set clsDEFAULT_USERs = Nothing
'
End Function

Public Function GetDefaultValues(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef ActiveLogicalID As String, _
                                ByRef ActiveUserNo As Long) As String
'
    Dim strDefaultValue As String
    
    ' check user default value
    strDefaultValue = GetUserDefaultValue(ActiveConnection, ActiveCodisheet, ActiveBoxCode, _
                                ActiveLanguage, ActiveLogicalID, ActiveUserNo)
                                
    ' check general default value
    If (strDefaultValue = "") Then
        
        strDefaultValue = GetGeneralDefaultValue(ActiveConnection, ActiveCodisheet, ActiveBoxCode, _
                                ActiveLanguage, ActiveLogicalID, ActiveUserNo)
        
        If (strDefaultValue = "") Then
            '
            Dim blnBoxIsDate As Boolean
            
            blnBoxIsDate = BoxTypeIsDate(ActiveConnection, ActiveCodisheet, ActiveBoxCode, ActiveLanguage)
            If (blnBoxIsDate = False) Then
                strDefaultValue = GetEmptyBoxValue(ActiveConnection, ActiveCodisheet, ActiveBoxCode, ActiveLanguage)
            ElseIf (blnBoxIsDate = True) Then
                strDefaultValue = GetEmptyBoxValue(ActiveConnection, ActiveCodisheet, ActiveBoxCode, ActiveLanguage)
                strDefaultValue = GetDateValue(strDefaultValue, ActiveCodisheet)
            End If
        End If
        
    End If
                                
    GetDefaultValues = strDefaultValue
                                    
End Function

Public Function BoxTypeIsDate(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String) As Boolean

    InitCodisheetType ActiveCodisheet
    
    Dim strInternalCode As String
    
    strInternalCode = GetInternalCode(ActiveConnection, mvarActiveDocument, ActiveBoxCode)
    
    BoxTypeIsDate = False
    If (strInternalCode = PCK_DATE) Then
        BoxTypeIsDate = True
    End If
    
End Function

' GetDateValue
Public Function GetDateValue(ByRef ActiveDateCode As String, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums) As String
'
    Select Case ActiveCodisheet
    
        Case cpiImportCodisheet
        
        Case cpiExportCodisheet
        
        Case cpiTransitCodisheet

        Case cpiSadbelNCTSCodisheet
        
        Case cpiCombinedNCTSCodisheet
        
        Case cpiDepartureIE15Codisheet
        
        Case cpiArrivalIE07Codisheet
        
            Select Case UCase$(ActiveDateCode)
                Case "A" ' yesterday
                    GetDateValue = Format$(Now - 1, "yyyyMMdd")
                Case "B" ' today
                    GetDateValue = Format$(Now, "yyyyMMdd")
                Case "C" ' tomorrow
                    GetDateValue = Format$(Now + 1, "yyyyMMdd")
            End Select
            
        Case cpiArrivalIE44Codisheet
        
            Select Case UCase$(ActiveDateCode)
                Case "A" ' yesterday
                    GetDateValue = Format$(Now - 1, "yyyyMMdd")
                Case "B" ' today
                    GetDateValue = Format$(Now, "yyyyMMdd")
                Case "C" ' tomorrow
                    GetDateValue = Format$(Now + 1, "yyyyMMdd")
            End Select
            
    End Select
    

'
End Function

Private Function InitCodisheetType(ByVal ActiveCodisheetType As cpiCodiSheetTypeEnums)
    
'    ActiveCodisheetType = vNewValue
    
    Select Case ActiveCodisheetType
    
        Case cpiImportCodisheet
        
            mvarActiveType = "I"
            mvarActiveDocument = "Import"
            mvarBoxDefaultAdminTable = "BOX DEFAULT IMPORT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE IMPORT"
            mvarDefaultUserTable = "DEFAULT USER IMPORT"
        
        Case cpiExportCodisheet
        
            mvarActiveType = "E"
            mvarActiveDocument = "Export/Transit"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EXPORT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EXPORT"
            mvarDefaultUserTable = "DEFAULT USER EXPORT"
        
        Case cpiTransitCodisheet
        
            mvarActiveType = "T"
            mvarActiveDocument = "Export/Transit"
            mvarBoxDefaultAdminTable = "BOX DEFAULT TRANSIT ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE TRANSIT"
            mvarDefaultUserTable = "DEFAULT USER TRANSIT"

        Case cpiSadbelNCTSCodisheet
        
            mvarActiveType = "N"
            mvarActiveDocument = "Transit NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT TRANSIT NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE TRANSIT NCTS"
            mvarDefaultUserTable = "DEFAULT USER TRANSIT NCTS"
        
        Case cpiCombinedNCTSCodisheet
        
            mvarActiveType = "C"
            mvarActiveDocument = "Combined NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT COMBINED NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE COMBINED NCTS"
            mvarDefaultUserTable = "DEFAULT USER COMBINED NCTS"
        
        Case cpiDepartureIE15Codisheet
        
            mvarActiveType = "D"
            mvarActiveDocument = "EDI NCTS"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS"
        
        Case cpiArrivalIE07Codisheet
            
            mvarActiveType = "A"
            mvarActiveDocument = "EDI NCTS2"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS2 ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS2"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS2"
        
        Case cpiArrivalIE44Codisheet
        
            mvarActiveType = "U"
            mvarActiveDocument = "EDI NCTS IE44"
            mvarBoxDefaultAdminTable = "BOX DEFAULT EDI NCTS IE44 ADMIN"
            mvarBoxDefaultValueTable = "BOX DEFAULT VALUE EDI NCTS IE44"
            mvarDefaultUserTable = "DEFAULT USER EDI NCTS IE44"
        
    End Select
    
End Function

Public Function GetSkipTabRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef UserNo As Long) As cpiSKIPs
    
    InitCodisheetType ActiveCodisheet
    '
    Dim strSql As String
    'Dim clsSKIP As cpiSKIP
    Dim clsSKIPs  As cpiSKIPs
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [SKIP] "
    strSql = strSql & "WHERE [REFERENCE]='" & ActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(UserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    '
    'Set clsSKIP = New cpiSKIP
    Set clsSKIPs = New cpiSKIPs
    '
    Set clsSKIPs.Recordset = ActiveConnection.Execute(strSql)
    
    Set GetSkipTabRecord = clsSKIPs
    
'    Do While (clsSKIPs.Recordset.EOF = False)
'
'        Set clsSKIP = clsSKIPs.GetClassRecord(clsSKIPs.Recordset)
'
'        '
'
'        clsSKIPs.Recordset.MoveNext
'
'    Loop
    
    'Set clsSKIP = Nothing
    Set clsSKIPs = Nothing
    '
End Function

Public Function GetTabOrderRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef UserNo As Long) As cpiTAB_ORDERs
    
    InitCodisheetType ActiveCodisheet
    
    Dim strSql As String
    ' Dim clsTAB_ORDER As cpiTAB_ORDER
    Dim clsTAB_ORDERs As cpiTAB_ORDERs
    
    strSql = "SELECT * "
    strSql = strSql & "FROM [TAB ORDER] "
    strSql = strSql & "WHERE [REFERENCE]='" & ActiveBoxCode & "'"
    strSql = strSql & " AND [USER NO]=" & CStr(UserNo) & ""
    strSql = strSql & " AND [TYPE]='" & mvarActiveType & "'"
    
    'Set clsTAB_ORDER = New cpiTAB_ORDER
    Set clsTAB_ORDERs = New cpiTAB_ORDERs
    '
    Set clsTAB_ORDERs.Recordset = ActiveConnection.Execute(strSql)
    
    Set GetTabOrderRecord = clsTAB_ORDERs
    
'    Do While (clsTAB_ORDERs.Recordset.EOF = False)
'
'        Set clsTAB_ORDER = clsTAB_ORDERs.GetClassRecord(clsTAB_ORDERs.Recordset)
'
'        clsTAB_ORDERs.Recordset.MoveNext
'
'    Loop
    
    'Set clsTAB_ORDER = Nothing
    Set clsTAB_ORDERs = Nothing
    '
    '
End Function

Public Function GetInternalCode_B(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef UserNo As Long) As String
    
    InitCodisheetType ActiveCodisheet
    
    GetInternalCode_B = GetInternalCode(ActiveConnection, mvarActiveDocument, ActiveBoxCode)
    
    
    
End Function

Public Function GetPicklistRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef UserNo As Long) As cpiPICK_DEFs

    InitCodisheetType ActiveCodisheet
    
    Dim strSql As String 'PICKLIST_DEFINITION
    Dim clsPICKLIST_DEFINITION As cpiPICK_DEF 'cpiPICK_DEF
    Dim clsPICKLIST_DEFINITIONs  As cpiPICK_DEFs
    '
    strSql = "SELECT * "
    strSql = strSql & " FROM [PICKLIST DEFINITION] "
    strSql = strSql & " WHERE [DOCUMENT]='" & mvarActiveDocument & "'"
    strSql = strSql & " AND [BOX CODE]='" & ActiveBoxCode & "'"
    
    Set clsPICKLIST_DEFINITION = New cpiPICK_DEF
    Set clsPICKLIST_DEFINITIONs = New cpiPICK_DEFs
'
    Set clsPICKLIST_DEFINITIONs.Recordset = ActiveConnection.Execute(strSql)
    
    Set GetPicklistRecord = clsPICKLIST_DEFINITIONs
    
'    Do While (clsPICKLIST_DEFINITIONs.Recordset.EOF = False)
'
'        Set clsPICKLIST_DEFINITION = clsPICKLIST_DEFINITIONs.GetClassRecord(clsPICKLIST_DEFINITIONs.Recordset)
'
'        '
'        clsPICKLIST_DEFINITIONs.Recordset.MoveNext
'
'    Loop
    
    Set clsPICKLIST_DEFINITION = Nothing
    Set clsPICKLIST_DEFINITIONs = Nothing

End Function

Public Function GetPicklistItemRecord(ByRef ActiveConnection As ADODB.Connection, _
                                ByRef ActiveCodisheet As cpiCodiSheetTypeEnums, _
                                ByRef ActiveBoxCode As String, _
                                ByRef ActiveLanguage As String, _
                                ByRef UserNo As Long) As cpiPICK_MAINTs
    '
    InitCodisheetType ActiveCodisheet
    Dim strInternalCode As String

    ' load box values here if any
    Dim clsPICKLIST_MAINTENANCE As cpiPICK_MAINT
    Dim clsPICKLIST_MAINTENANCEs As cpiPICK_MAINTs
        
    Set clsPICKLIST_MAINTENANCE = New cpiPICK_MAINT
    Set clsPICKLIST_MAINTENANCEs = New cpiPICK_MAINTs
    
    Dim strSql As String
    
    strInternalCode = GetInternalCode(ActiveConnection, mvarActiveDocument, ActiveBoxCode)
    
    If (strInternalCode <> "") Then
    
        clsPICKLIST_MAINTENANCEs.SetSqlParameters "[PICKLIST MAINTENANCE " & ActiveLanguage & "]", ActiveLanguage
        
        strSql = "SELECT * "
        strSql = strSql & " FROM [PICKLIST MAINTENANCE " & ActiveLanguage & "] "
        strSql = strSql & " WHERE [INTERNAL CODE]='" & strInternalCode & "'"
        
        Set clsPICKLIST_MAINTENANCEs.Recordset = ActiveConnection.Execute(strSql)
        
        Set GetPicklistItemRecord = clsPICKLIST_MAINTENANCEs
        
'        Do While (clsPICKLIST_MAINTENANCEs.Recordset.EOF = False)
'
'            ' CBO_BOXVALUE_TABORDER
'            Set clsPICKLIST_MAINTENANCE = clsPICKLIST_MAINTENANCEs.GetClassRecord(clsPICKLIST_MAINTENANCEs.Recordset)
'
'            clsPICKLIST_MAINTENANCEs.Recordset.MoveNext
'        Loop
        
    End If

    Set clsPICKLIST_MAINTENANCE = Nothing
    Set clsPICKLIST_MAINTENANCEs = Nothing

'
End Function
