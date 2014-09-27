VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCusres"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCusoffretcopocp As New CCusoffretcopocp
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Private mvarCDxxxA As New CCDxxxA
Private mvarHeaHea As New CHea
Private mvarTrapripc1 As New CTrapripc1
Private mvarTraconco1 As New CTraconco1
Private mvarTraconce1 As New CTraconce1
Private mvarFunerrer As New CFunerrer
Private mvarGuarantor As New CGuarantor
Private mvarCusoffDepEpt As New CCusoffdepept
Private mvarConresers As New CConresers
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private mvarLanguageCode As String 'local copy

Private mvarMessageResponse As String ' local copy

Private mvarPreviousNode As String
Private mvarNodesFound As Long
Private mvarNodeValues() As String

Private Enum DeclarantTypeConstant
    [Consignor]
    [Consignee]
    [Representative]
    [Principal]
    [Invoice Issuer]
    [Beneficiary]
    [Declarant]
    [Intracom]
    [Warehouse Depositor]
End Enum

'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb 022008
Private mvarSynIdeMES1 As String
Private mvarSynVerNumMES2 As String
Private mvarMesSenMES3 As String
Private mvarSenIdeCodQuaMES4 As String
Private mvarMesRecMES6 As String
Private mvarRecIdeCodQuaMES7 As String
Private mvarDatOfPreMES9 As String
Private mvarTimOfPreMES10 As String
Private mvarIntConRefMES11 As String
Private mvarRecRefMES12 As String
Private mvarRecRefQuaMES13 As String
Private mvarAppRefMES14 As String
Private mvarPriMES15 As String
Private mvarAckReqMES16 As String
Private mvarComAgrIdMES17 As String
Private mvarTesIndMES18 As String
Private mvarMesIdeMES19 As String
Private mvarMesTypMES20 As String
Private mvarComAccRefMES21 As String
Private mvarMesSeqNumMES22 As String
Private mvarFirAndLasTraMES23 As String

'LET

Public Property Let SynIdeMES1(ByVal vData As String)
    mvarSynIdeMES1 = vData
End Property
Public Property Let SynVerNumMES2(ByVal vData As String)
    mvarSynVerNumMES2 = vData
End Property
Public Property Let MesSenMES3(ByVal vData As String)
    mvarMesSenMES3 = vData
End Property
Public Property Let SenIdeCodQuaMES4(ByVal vData As String)
    mvarSenIdeCodQuaMES4 = vData
End Property
Public Property Let MesRecMES6(ByVal vData As String)
    mvarMesRecMES6 = vData
End Property
Public Property Let RecIdeCodQuaMES7(ByVal vData As String)
    mvarRecIdeCodQuaMES7 = vData
End Property
Public Property Let DatOfPreMES9(ByVal vData As String)
    mvarDatOfPreMES9 = vData
End Property
Public Property Let TimOfPreMES10(ByVal vData As String)
    mvarTimOfPreMES10 = vData
End Property
Public Property Let IntConRefMES11(ByVal vData As String)
    mvarIntConRefMES11 = vData
End Property
Public Property Let RecRefMES12(ByVal vData As String)
    mvarRecRefMES12 = vData
End Property
Public Property Let RecRefQuaMES13(ByVal vData As String)
    mvarRecRefQuaMES13 = vData
End Property
Public Property Let AppRefMES14(ByVal vData As String)
    mvarAppRefMES14 = vData
End Property
Public Property Let PriMES15(ByVal vData As String)
    mvarPriMES15 = vData
End Property
Public Property Let AckReqMES16(ByVal vData As String)
    mvarAckReqMES16 = vData
End Property
Public Property Let ComAgrIdMES17(ByVal vData As String)
    mvarComAgrIdMES17 = vData
End Property
Public Property Let TesIndMES18(ByVal vData As String)
    mvarTesIndMES18 = vData
End Property
Public Property Let MesIdeMES19(ByVal vData As String)
    mvarMesIdeMES19 = vData
End Property
Public Property Let MesTypMES20(ByVal vData As String)
    mvarMesTypMES20 = vData
End Property
Public Property Let ComAccRefMES21(ByVal vData As String)
    mvarComAccRefMES21 = vData
End Property
Public Property Let MesSeqNumMES22(ByVal vData As String)
    mvarMesSeqNumMES22 = vData
End Property
Public Property Let FirAndLasTraMES23(ByVal vData As String)
    mvarFirAndLasTraMES23 = vData
End Property

'GET

Public Property Get SynIdeMES1() As String
    SynIdeMES1 = mvarSynIdeMES1
End Property
Public Property Get SynVerNumMES2() As String
    SynVerNumMES2 = mvarSynVerNumMES2
End Property
Public Property Get MesSenMES3() As String
    MesSenMES3 = mvarMesSenMES3
End Property
Public Property Get SenIdeCodQuaMES4() As String
    SenIdeCodQuaMES4 = mvarSenIdeCodQuaMES4
End Property
Public Property Get MesRecMES6() As String
    MesRecMES6 = mvarMesRecMES6
End Property
Public Property Get RecIdeCodQuaMES7() As String
    RecIdeCodQuaMES7 = mvarRecIdeCodQuaMES7
End Property
Public Property Get DatOfPreMES9() As String
    DatOfPreMES9 = mvarDatOfPreMES9
End Property
Public Property Get TimOfPreMES10() As String
    TimOfPreMES10 = mvarTimOfPreMES10
End Property
Public Property Get IntConRefMES11() As String
    IntConRefMES11 = mvarIntConRefMES11
End Property
Public Property Get RecRefMES12() As String
    RecRefMES12 = mvarRecRefMES12
End Property
Public Property Get RecRefQuaMES13() As String
    RecRefQuaMES13 = mvarRecRefQuaMES13
End Property
Public Property Get AppRefMES14() As String
    AppRefMES14 = mvarAppRefMES14
End Property
Public Property Get PriMES15() As String
    PriMES15 = mvarPriMES15
End Property
Public Property Get AckReqMES16() As String
    AckReqMES16 = mvarAckReqMES16
End Property
Public Property Get ComAgrIdMES17() As String
    ComAgrIdMES17 = mvarComAgrIdMES17
End Property
Public Property Get TesIndMES18() As String
    TesIndMES18 = mvarTesIndMES18
End Property
Public Property Get MesIdeMES19() As String
    MesIdeMES19 = mvarMesIdeMES19
End Property
Public Property Get MesTypMES20() As String
    MesTypMES20 = mvarMesTypMES20
End Property
Public Property Get ComAccRefMES21() As String
    ComAccRefMES21 = mvarComAccRefMES21
End Property
Public Property Get MesSeqNumMES22() As String
    MesSeqNumMES22 = mvarMesSeqNumMES22
End Property
Public Property Get FirAndLasTraMES23() As String
    FirAndLasTraMES23 = mvarFirAndLasTraMES23
End Property
'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb

Public Property Get Guarantor() As CGuarantor
    Set Guarantor = mvarGuarantor
End Property

'Public Property Get CDxxxA() As CCDxxxA
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.ResponseInterchangeHeader
'    Set CDxxxA = mvarCDxxxA
'End Property

Public Property Get HeaHea() As CHea
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResponseInterchangeHeader
    Set HeaHea = mvarHeaHea
End Property

Public Property Get Trapripc1() As CTrapripc1
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResponseInterchangeHeader
    Set Trapripc1 = mvarTrapripc1
End Property

Public Property Get Funerrer() As CFunerrer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResponseInterchangeHeader
    Set Funerrer = mvarFunerrer
End Property

Public Property Get CusoffDepEpt() As CCusoffdepept
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResponseInterchangeHeader
    Set CusoffDepEpt = mvarCusoffDepEpt
End Property

Public Property Get Conresers() As CConresers
    Set Conresers = mvarConresers
End Property

'CSCLP-402
Public Property Get Cusoffretcopocp() As CCusoffretcopocp
    Set Cusoffretcopocp = mvarCusoffretcopocp
End Property

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Function SetProp(ByVal CustomsSystem As Long, _
                              ByVal MessageResponse As String) As Boolean
    
    Select Case CustomsSystem
        'Case CPCustomsSystemConstant.[PLDA - Belgium]
            'GetPropertiesFromPLDAEDI MessageResponseFilePath
        
        'Case CPCustomsSystemConstant.[PLDA - Luxembourg]
            'If GetPropertiesFromPLDAXML(MessageResponse) = False Then
            '    SetProperties = False
            '    Exit Function
            'End If
        
        Case 1
            If GetFromNCTSXML(MessageResponse) = False Then
                SetProp = False
                Exit Function
            End If
            'GetPropertiesFromNCTSEDI MessageResponseFilePath
                
        'Case CPCustomsSystemConstant.[SADBEL]
            'GetPropertiesFromSadbelEDI MessageResponseFilePath
            
        Case Else
            Debug.Assert False
            
    End Select
    
    SetProp = True
    
End Function

Public Property Let LanguageCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LanguageCode = 5
    mvarLanguageCode = vData
End Property

Public Property Get LanguageCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LanguageCode
    LanguageCode = mvarLanguageCode
End Property

Public Property Get MessageResponse() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MessageResponse
    MessageResponse = mvarMessageResponse
End Property

'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
Public Function GetNCTSTemporaryFilename() As String
    Dim strTemporaryFileName As String
    Dim strWindowsTemporaryPath As String
    
    strWindowsTemporaryPath = WindowsTempPath
    
    ' Create a buffer
    strTemporaryFileName = String(260, 0)
    ' Get a temporary filename
    GetTempFileName AddBackSlashOnPath(strWindowsTemporaryPath), "FFF", 0, strTemporaryFileName
        
    ' Remove all the unnecessary chr$(0)'s
    strTemporaryFileName = Left$(strTemporaryFileName, InStr(1, strTemporaryFileName, Chr$(0)) - 1)
    
    ' Set the file attributes
    SetFileAttributes strTemporaryFileName, FILE_ATTRIBUTE_TEMPORARY
        
    GetNCTSTemporaryFilename = strTemporaryFileName
End Function
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
'Private Function SetCDxxxA(ByVal ParentNode As IXMLDOMNode) As Boolean
'
'    Dim strTemp As String
'    'Dim objDateTime As New CDateTimeElement
'
'    On Error GoTo ErrHandler
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "SynIdeMES1", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.SynIdeMES1 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "SynVerNumMES2", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.SynVerNumMES2 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "MesSenMES3", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.MesSenMES3 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "SenIdeCodQuaMES4", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.SenIdeCodQuaMES4 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "MesRecMES6", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.MesRecMES6 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "RecIdeCodQuaMES7", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.RecIdeCodQuaMES7 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "DatOfPreMES9", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.DatOfPreMES9 = strTemp
'
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "TimOfPreMES10", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.TimOfPreMES10 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "IntConRefMES11", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.IntConRefMES11 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "RecRefMES12", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.RecRefMES12 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "RecRefQuaMES13", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.RecRefQuaMES13 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "AppRefMES14", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.AppRefMES14 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "PriMES15", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.PriMES15 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "AckReqMES16", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.AckReqMES16 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "ComAgrIdMES17", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.ComAgrIdMES17 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "TesIndMES18", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.TesIndMES18 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "MesIdeMES19", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.MesIdeMES19 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "MesTypMES20", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.MesTypMES20 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "ComAccRefMES21", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.ComAccRefMES21 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "MesSeqNumMES22", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.MesSeqNumMES22 = strTemp
'
'    strTemp = vbNullString
'    If FindNode(ParentNode, "FirAndLasTraMES23", 0, strTemp) = False Then GoTo ErrHandler
'    mvarCDxxxA.FirAndLasTraMES23 = strTemp
'
'    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
'
'    SetCDxxxA = True
'
'    'Set objDateTime = Nothing
'
'    Exit Function
'
'ErrHandler:
'
'    SetCDxxxA = False
'
'    'Set objDateTime = Nothing
'
'End Function
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA

'CSCLP-402
Private Function SetCusoffretcopocp(ByVal ParentNode As IXMLDOMNode) As Boolean
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
    strTemp = vbNullString
    
    If FindNode(ParentNode, "RefNumOCP1", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.RefNumOCP1 = strTemp
    
    If FindNode(ParentNode, "CusOffNamOCP2", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.CusOffNamOCP2 = strTemp
    
    If FindNode(ParentNode, "StrAndNumOCP3", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.StrAndNumOCP3 = strTemp
    
    If FindNode(ParentNode, "CouOCP4", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.CouOCP4 = strTemp
    
    If FindNode(ParentNode, "PosCodOCP6", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.PosCodOCP6 = strTemp
    
    If FindNode(ParentNode, "CitOCP7", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffretcopocp.CitOCP7 = strTemp
        
    SetCusoffretcopocp = True
        
    Exit Function
    
ErrHandler:

    SetCusoffretcopocp = False
    
End Function

Private Function SetCusoffdepept(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "RefNumEPT1", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffDepEpt.RefNumEPT1 = strTemp
    
    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
    
    SetCusoffdepept = True
    
    'Set objDateTime = Nothing
    
    Exit Function
    
ErrHandler:

    SetCusoffdepept = False
    
    'Set objDateTime = Nothing
    
End Function
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
Private Function SetGuarantor(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "RefNumEPT1", 0, strTemp) = False Then GoTo ErrHandler
    mvarCusoffDepEpt.RefNumEPT1 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NamGU20", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.NamGU20 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "StrAndNumGU621", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.StrAndNumGU621 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "PosCodGU622", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.PosCodGU622 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CitGU623", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.CitGU623 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouGU624", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.CouGU624 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NADLNGGU", 0, strTemp) = False Then GoTo ErrHandler
    mvarGuarantor.NADLNGGU = strTemp
    
    SetGuarantor = True
    Exit Function
    
ErrHandler:

    SetGuarantor = False
    
End Function

'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ

Private Function SetConresers(ByVal ParentNode As IXMLDOMNode) As Boolean

    Dim strTemp As String
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "ConResCodERS16", 0, strTemp) = False Then GoTo ErrHandler
    mvarConresers.ConResCodERS16 = strTemp
        
    strTemp = vbNullString
    If FindNode(ParentNode, "DatLimERS69", 0, strTemp) = False Then GoTo ErrHandler
    mvarConresers.DatLimERS69 = strTemp
        
    SetConresers = True
    
    Exit Function
    
ErrHandler:

    SetConresers = False
    

End Function

'YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
Private Function SetTrapripc1(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "NamPC17", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.NamPC17 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "StrAndNumPC122", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.StrAndNumPC122 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "PosCodPC123", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.PosCodPC123 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CitPC124", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.CitPC124 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouPC125", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.CouPC125 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NADLNGPC", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.NADLNGPC = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TINPC159", 0, strTemp) = False Then GoTo ErrHandler
    mvarTrapripc1.TINPC159 = strTemp
        
    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
    
    SetTrapripc1 = True
    
    'Set objDateTime = Nothing
    
    Exit Function
    
ErrHandler:

    SetTrapripc1 = False
    
    'Set objDateTime = Nothing
    
End Function

'BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB
Private Function SetTraconce1(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "NamCE17", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.NamCE17 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "StrAndNumCE122", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.StrAndNumCE122 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "PosCodCE123", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.PosCodCE123 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CitCE124", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.CitCE124 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouCE125", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.CouCE125 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NADLNGCE", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.NADLNGCE = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TINCE159", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconce1.TINCE159 = strTemp
        
    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
    
    SetTraconce1 = True
    
    'Set objDateTime = Nothing
    
    Exit Function
    
ErrHandler:

    SetTraconce1 = False
    
    'Set objDateTime = Nothing
    
End Function

'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
Private Function SetTraconco1(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
      
    strTemp = vbNullString
    If FindNode(ParentNode, "NamCO17", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.NamCO17 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "StrAndNumCO122", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.StrAndNumCO122 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "PosCodCO123", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.PosCodCO123 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CitCO124", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.CitCO124 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouCO125", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.CouCO125 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NADLNGCO", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.NADLNGCO = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TINCO159", 0, strTemp) = False Then GoTo ErrHandler
    mvarTraconco1.TINCO159 = strTemp
        
    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
    
    SetTraconco1 = True
    
    'Set objDateTime = Nothing
    
    Exit Function
    
ErrHandler:

    SetTraconco1 = False
    
    'Set objDateTime = Nothing
    
End Function

'BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB

Private Function SetHeahea(ByVal ParentNode As IXMLDOMNode) As Boolean
    
    Dim strTemp As String
    'Dim objDateTime As New CDateTimeElement
    
    On Error GoTo ErrHandler
        
    strTemp = vbNullString
    If FindNode(ParentNode, "DocNumHEA5 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DocNumHEA5 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanDecHEA93 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanDecHEA93 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DatOfCanReqHEA147 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DatOfCanReqHEA147 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanIniByCusHEA94 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanIniByCusHEA94 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DatOfCanDecHEA146 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DatOfCanDecHEA146 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanJusHEA248 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanJusHEA248 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanJusHEA248LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanJusHEA248LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanReaHEA250 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanReaHEA250 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CanReaHEA250LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CanReaHEA250LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "RefNumHEA4 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.RefNumHEA4 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TypOfDecHEA24 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TypOfDecHEA24 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouOfDesCodHEA30 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CouOfDesCodHEA30 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AgrLocOfGooCodHEA38 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AgrLocOfGooCodHEA38 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AgrLocOfGooHEA39 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AgrLocOfGooHEA39 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AgrLocOfGooHEA39LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AgrLocOfGooHEA39LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AutLocOfGooCodHEA41 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AutLocOfGooCodHEA41 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "PlaOfLoaCodHEA46 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.PlaOfLoaCodHEA46 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CouOfDisCodHEA55 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CouOfDisCodHEA55 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "CusSubPlaHEA66 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.CusSubPlaHEA66 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "InlTraModHEA75 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.InlTraModHEA75 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TraModAtBorHEA76 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TraModAtBorHEA76 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "IdeOfMeaOfTraAtDHEA78 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.IdeOfMeaOfTraAtDHEA78 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "IdeOfMeaOfTraAtDHEA78LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.IdeOfMeaOfTraAtDHEA78LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NatOfMeaOfTraAtDHEA80 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.NatOfMeaOfTraAtDHEA80 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "IdeOfMeaOfTraCroHEA85 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.IdeOfMeaOfTraCroHEA85 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "IdeOfMeaOfTraCroHEA85LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.IdeOfMeaOfTraCroHEA85LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NatOfMeaOfTraCroHEA87 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.NatOfMeaOfTraCroHEA87 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TypOfMeaOfTraCroHEA88 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TypOfMeaOfTraCroHEA88 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "ConInHEA96 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.ConInHEA96 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DiaLanIndAtDepHEA254 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DiaLanIndAtDepHEA254 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NCTSAccDocHEA601LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.NCTSAccDocHEA601LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NumOfLoaLisHEA304 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.NumOfLoaLisHEA304 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TotNumOfIteHEA305 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TotNumOfIteHEA305 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TotNumOfPacHEA306 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TotNumOfPacHEA306 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "TotGroMasHEA307 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.TotGroMasHEA307 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecDatHEA383 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecDatHEA383 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecPlaHEA394 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecPlaHEA394 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecPlaHEA394LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecPlaHEA394LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecRejDatHEA159 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecRejDatHEA159 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecRejReaHEA252 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecRejReaHEA252 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DecRejReaHEA252LNG ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DecRejReaHEA252LNG = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AccDotHEA158 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AccDotHEA158 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "NCTRetCopHEA104 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.NCTRetCopHEA104 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "AccDatHEA158 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.AccDatHEA158 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "IssDatHEA186 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.IssDatHEA186 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "WriOffDat ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.WriOffDat = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "DatOfConNotHEA148 ", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.DatOfConNotHEA148 = strTemp
    
    strTemp = vbNullString
    If FindNode(ParentNode, "OriMesIdeMES22", 0, strTemp) = False Then GoTo ErrHandler
    mvarHeaHea.OriMesIdeMES22 = strTemp

    'Set mvarResponseInterchangeHeader.DateTimeOfPreparation = objDateTime
    
    SetHeahea = True
    'Set objDateTime = Nothing
    Exit Function
    
ErrHandler:

    SetHeahea = False
    'Set objDateTime = Nothing
    
End Function


Private Function GetFromNCTSXML(ByVal MessageResponse As String) As Boolean

    Dim objDOM As DOMDocument
    Dim objChildNode As IXMLDOMNode
    Dim strFunctionCode As String
    
    Dim strTemporaryFileName As String
    Dim lngFreeFile As Long
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    
    strTemporaryFileName = GetTemporaryFilename
    
    lngFreeFile = FreeFile()
    Open strTemporaryFileName For Output As #lngFreeFile
    Print #lngFreeFile, MessageResponse
    Close #lngFreeFile

    ' Rename temporary file
    Name strTemporaryFileName As Replace(UCase$(strTemporaryFileName), ".TMP", ".xml")
    
    strTemporaryFileName = Replace(UCase$(strTemporaryFileName), ".TMP", ".xml")
        
    Set objDOM = New MSXML2.DOMDocument
    
    ' Load XML Document to XML Object
    objDOM.Load strTemporaryFileName

    ' Set MessageResponse Property
    mvarMessageResponse = objDOM.xml
    
    For Each objChildNode In objDOM.documentElement.childNodes
        
        Select Case UCase(objChildNode.nodeName)
            'Case "CDXXXA"
            '    If SetCDxxxA(objChildNode) = False Then GoTo ErrHandler
            
            '>>>>>>>>>
            Case UCase("SynIdeMES1")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "SynIdeMES1", 0, strTemp) = False Then GoTo ErrHandler
                mvarSynIdeMES1 = strTemp
            
            Case UCase("SynVerNumMES2")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "SynVerNumMES2", 0, strTemp) = False Then GoTo ErrHandler
                mvarSynVerNumMES2 = strTemp

            Case UCase("MesSenMES3")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "MesSenMES3", 0, strTemp) = False Then GoTo ErrHandler
                mvarMesSenMES3 = strTemp

            Case UCase("SenIdeCodQuaMES4")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "SenIdeCodQuaMES4", 0, strTemp) = False Then GoTo ErrHandler
                mvarSenIdeCodQuaMES4 = strTemp

            Case UCase("MesRecMES6")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "MesRecMES6", 0, strTemp) = False Then GoTo ErrHandler
                mvarMesRecMES6 = strTemp

            Case UCase("RecIdeCodQuaMES7")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "RecIdeCodQuaMES7", 0, strTemp) = False Then GoTo ErrHandler
                mvarRecIdeCodQuaMES7 = strTemp

            Case UCase("DatOfPreMES9")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "DatOfPreMES9", 0, strTemp) = False Then GoTo ErrHandler
                mvarDatOfPreMES9 = strTemp

            Case UCase("TimOfPreMES10")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "TimOfPreMES10", 0, strTemp) = False Then GoTo ErrHandler
                mvarTimOfPreMES10 = strTemp

            Case UCase("IntConRefMES11")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "IntConRefMES11", 0, strTemp) = False Then GoTo ErrHandler
                mvarIntConRefMES11 = strTemp

            Case UCase("RecRefMES12")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "RecRefMES12", 0, strTemp) = False Then GoTo ErrHandler
                mvarRecRefMES12 = strTemp

            Case UCase("RecRefQuaMES13")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "RecRefQuaMES13", 0, strTemp) = False Then GoTo ErrHandler
                mvarRecRefQuaMES13 = strTemp

            Case UCase("AppRefMES14")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "AppRefMES14", 0, strTemp) = False Then GoTo ErrHandler
                mvarAppRefMES14 = strTemp

            Case UCase("PriMES15")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "PriMES15", 0, strTemp) = False Then GoTo ErrHandler
                mvarPriMES15 = strTemp
            
            Case UCase("AckReqMES16")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "AckReqMES16", 0, strTemp) = False Then GoTo ErrHandler
                mvarAckReqMES16 = strTemp

            Case UCase("ComAgrIdMES17")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "ComAgrIdMES17", 0, strTemp) = False Then GoTo ErrHandler
                mvarComAgrIdMES17 = strTemp

            Case UCase("TesIndMES18")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "TesIndMES18", 0, strTemp) = False Then GoTo ErrHandler
                mvarTesIndMES18 = strTemp

            Case UCase("MesIdeMES19")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "MesIdeMES19", 0, strTemp) = False Then GoTo ErrHandler
                mvarMesIdeMES19 = strTemp

            Case UCase("MesTypMES20")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "MesTypMES20", 0, strTemp) = False Then GoTo ErrHandler
                mvarMesTypMES20 = strTemp

            Case UCase("ComAccRefMES21")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "ComAccRefMES21", 0, strTemp) = False Then GoTo ErrHandler
                mvarComAccRefMES21 = strTemp

            Case UCase("MesSeqNumMES22")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "MesSeqNumMES22", 0, strTemp) = False Then GoTo ErrHandler
                mvarMesSeqNumMES22 = strTemp

            Case UCase("FirAndLasTraMES23")
                strTemp = vbNullString
                If FindNode(objChildNode.ParentNode, "FirAndLasTraMES23", 0, strTemp) = False Then GoTo ErrHandler
                mvarFirAndLasTraMES23 = strTemp
            '>>>>>>>>>
            
            Case "HEAHEA"
                If SetHeahea(objChildNode) = False Then GoTo ErrHandler

            Case "TRAPRIPC1"
                If SetTrapripc1(objChildNode) = False Then GoTo ErrHandler

            Case "TRACONCO1"
                If SetTraconco1(objChildNode) = False Then GoTo ErrHandler

            Case "TRACONCE1"
                If SetTraconce1(objChildNode) = False Then GoTo ErrHandler

            Case "CUSOFFDEPEPT"
                If SetCusoffdepept(objChildNode) = False Then GoTo ErrHandler
            
            Case "GUARANTOR"
                If SetGuarantor(objChildNode) = False Then GoTo ErrHandler
            
            Case "CONRESERS"
                If SetConresers(objChildNode) = False Then GoTo ErrHandler
            
            'CSCLP-402
            Case "CUSOFFRETCOPOCP"
                If SetCusoffretcopocp(objChildNode) = False Then GoTo ErrHandler
                
            Case Else
            '    Debug.Assert False
                
        End Select
        
    Next
        
    GetFromNCTSXML = True
    Set objChildNode = Nothing
    Set objDOM = Nothing
        
    On Error Resume Next
    Kill strTemporaryFileName
    On Error GoTo 0
        
    Exit Function
    
ErrHandler:
    
    GetFromNCTSXML = False
    
    Set objChildNode = Nothing
    Set objDOM = Nothing
    
    On Error Resume Next
    Kill strTemporaryFileName
    On Error GoTo 0
    
End Function
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA

Private Function FindNode(ByVal ParentNode As IXMLDOMNode, _
                          ByVal NodeNameToFind As String, _
                          ByVal NodeLocation As Long, _
                          ByRef NodeValue As String, _
                          Optional ByVal Required As Boolean) As Boolean
    
    ' ParentNode        - XML's parent node
    ' NodeNameToFind    - Name of the node to search from ParentNode
    ' NodeLocation      - Nth instance of target NodeName in case of having nodes with same name
    ' NodeValue         - Value of the node being search
    
    InitiliatizeFindNodeFromXML
    If FindNodeFromXML(ParentNode, NodeNameToFind) = True Then
        If mvarNodesFound > 0 Then
            If UBound(mvarNodeValues) >= NodeLocation Then
                NodeValue = CStr(mvarNodeValues(NodeLocation))
            Else
                Debug.Print "Node member " & NodeLocation & " is not found for node name: " & NodeNameToFind
                If Required = True Then
                    Debug.Assert False
                    FindNode = False
                    Exit Function
                End If
            End If
        Else
            Debug.Print "Node " & NodeNameToFind & " is not found in the XML Message!!!"
            If Required = True Then
                Debug.Assert False
                FindNode = False
                Exit Function
            End If
        End If
    Else
        Debug.Assert False
        FindNode = False
        Exit Function
    End If
    
    FindNode = True
    
End Function

Private Function FindNodes(ByVal ParentNode As IXMLDOMNode, _
                           ByVal NodeNameToFind As String, _
                           ByRef NodeCount As Long, _
                           ByRef NodeValues As Variant, _
                           Optional Required As Boolean) As Boolean
    
    ' ParentNode        - XML's parent node
    ' NodeNameToFind    - Name of the node to search from ParentNode
    ' NodeCount         - Count of target NodeName occurence
    ' NodeValues        - Array containing the node values
    
    InitiliatizeFindNodeFromXML
    If FindNodeFromXML(ParentNode, NodeNameToFind) = True Then
        If mvarNodesFound > 0 Then
            NodeValues = mvarNodeValues
            NodeCount = mvarNodesFound
        Else
            Debug.Print "Node " & NodeNameToFind & " is not found in the XML Message!!!"
            If Required = True Then
                Debug.Assert False
                FindNodes = False
                Exit Function
            End If
            NodeCount = mvarNodesFound
        End If
    Else
        Debug.Print "Node " & NodeNameToFind & " is not found in the XML Message!!!"
        If Required = True Then
            Debug.Assert False
            FindNodes = False
            Exit Function
        End If
    End If
    
    FindNodes = True
    
End Function

Private Function FindNodeFromXML(ByVal Node As IXMLDOMNode, _
                                ByVal NodeNameToFind As String) As Boolean

    Dim objChildNode As IXMLDOMNode
    Dim blnNodeFound As Boolean
    
    On Error GoTo ErrHandler

    For Each objChildNode In Node.childNodes
        
        If LCase(mvarPreviousNode) = Trim(LCase(NodeNameToFind)) Then
            ReDim Preserve mvarNodeValues(mvarNodesFound)
            
            If objChildNode.nodeType = NODE_TEXT Then
                If IsNull(objChildNode.NodeValue) = False Then
                    mvarNodeValues(mvarNodesFound) = objChildNode.NodeValue
                Else
                    mvarNodeValues(mvarNodesFound) = vbNullString
                End If
            End If
            
            Debug.Print mvarPreviousNode & " - " & mvarNodeValues(mvarNodesFound)
            mvarNodesFound = mvarNodesFound + 1
            
        ElseIf LCase(objChildNode.nodeName) = Trim(LCase(NodeNameToFind)) Then
            blnNodeFound = True
        End If
        
        mvarPreviousNode = objChildNode.nodeName
        
        If objChildNode.hasChildNodes = True Then
            If FindNodeFromXML(objChildNode, NodeNameToFind) = False Then
                FindNodeFromXML = False
                Exit Function
            End If
            blnNodeFound = False
        Else
            If blnNodeFound = True Then
                ReDim Preserve mvarNodeValues(mvarNodesFound)
                mvarNodeValues(mvarNodesFound) = vbNullString
                Debug.Print mvarPreviousNode & " - " & mvarNodeValues(mvarNodesFound)
                mvarNodesFound = mvarNodesFound + 1
                mvarPreviousNode = vbNullString
                blnNodeFound = False
            End If
        End If
        
    Next
    
    FindNodeFromXML = True

    Set objChildNode = Nothing
    
    Exit Function

ErrHandler:

    FindNodeFromXML = False
    
    Set objChildNode = Nothing
    
End Function

Private Sub InitiliatizeFindNodeFromXML()

    ' Call this procedure before using Function FindNodeFromXML
    mvarPreviousNode = vbNullString
    mvarNodesFound = 0
    Erase mvarNodeValues
    
End Sub

