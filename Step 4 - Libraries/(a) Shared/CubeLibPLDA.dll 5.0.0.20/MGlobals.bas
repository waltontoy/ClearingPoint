Attribute VB_Name = "MGlobals"
Option Explicit

Public Const EDI_SEP_SEGMENT                As String = "'"
Public Const EDI_SEP_COMPOSITE_DATA_ELEMENT As String = ":"
Public Const EDI_SEP_DATA_ELEMENT           As String = "+"
Public Const EDI_SEP_RELEASE_CHARACTER      As String = "?"

Public g_strDigiSignFieldVerifier As String             'Field Verifier - Edwin Oct 24
Public g_rstDigiSignData As ADODB.Recordset             'Recordset for the data to digisign
Public g_rstDigiSign As ADODB.Recordset                 'Recordset for the actual data to pack
Public g_blnDigiSignActivated As Boolean                'BooleanIndicator that Digital Signature is Activated
Public g_lngDigiSignCounter As Long
    
'**********
Public g_rstHeaderSeals As ADODB.Recordset
Public g_rstHeaderTransitOffices As ADODB.Recordset
Public g_rstHeaderGuarantee As ADODB.Recordset

Public g_rstDetailBijzondere As ADODB.Recordset
Public g_rstDetailBerekenings As ADODB.Recordset
Public g_rstDetailDocumenten As ADODB.Recordset
Public g_rstDetailZelf As ADODB.Recordset
Public g_rstDetailContainer As ADODB.Recordset
Public g_rstDetailSensitiveGoods As ADODB.Recordset
'**********
    
Public g_lngDType As Long                               'DType

Public g_clsSignData As CDataToSign                     'Added by Philip on 02-22-2007. Digital signature string
                                                
Public g_intDigitalSignatureType As Integer              'Added by Philip 02-26-2007. Selected option for the digital signature
Public g_strCertificateToUse As String                  'Added by Philip 02-26-2007. Certificate to use for the digital signature

Public Enum DigitalSignatureType                        'Added by Philip 02-26-2007. Option for digital signature
    [None] = 0
    [Fixed] = 1
    [User Defined] = 2
End Enum

Public g_rstDetails As ADODB.Recordset         'Added by Migs on 04-04-2006. This is needed for the procs in MProcedures.
Public g_rstDetailsHandelaars As ADODB.Recordset    'Added by Migs on 08-09-2006. This is needed for the procs in MProcedures.

Public Enum DECLARATION_MODE
    enuAddition = 2
    enuCancellation = 3
    enuAmendment = 4
    enuOriginal = 9
End Enum

Public Enum DECLARATION_TYPE
    enuImport = 14
    enuExport = 18
End Enum

Public Enum HEADER_OPERATOR_TYPE
    enuHeaderConsignee = 1
    enuHeaderDeclarant = 2
    enuHeaderIntracommunautaireVerwerving = 3
    enuHeaderResponsibleRepresentative = 4
End Enum

Public Enum EXPORT_HEADER_OPERATOR_TYPE
    enuExportHeaderDeclarant = 1
    enuExportHeaderResponsibleRepresentative = 2
    enuExportHeaderBenificiary = 3
    enuExportHeaderConsignee = 4
    enuexportHeaderExporter = 5
End Enum

Public Enum DETAIL_OPERATOR_TYPE
    enuDetailConsignee = 1
    enuDetailIntracommunautaireVerwerving = 2
    enuDetailResponsibleRepresentative = 3
    enuDetailWarehouseDepositor = 4
End Enum

Public Enum EXPORT_DETAIL_OPERATOR_TYPE
    enuExportDetailExporter = 1
    enuExportDetailConsignee = 2
    enuExportDetailWarehouseDepositor = 3
End Enum

Public Const G_MAIN_PASSWORD = "wack2"

Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public m_strDeclarantVentureNumber As String

Public g_conSADBEL As ADODB.Connection

Public Function ReplaceSpecialCharacters(ByVal SourceString As String) As String
    Dim strReturnValue As String
    strReturnValue = Trim(SourceString)
    If strReturnValue <> vbNullString Then
        '----->  RELEASE CHARACTER
        strReturnValue = Replace(strReturnValue, EDI_SEP_RELEASE_CHARACTER, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_RELEASE_CHARACTER)
        '----->  SEGMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_SEGMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_SEGMENT)
        '----->  COMPOSITE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_COMPOSITE_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_COMPOSITE_DATA_ELEMENT)
        '----->  SIMPLE DATA ELEMENT SEPARATOR
        strReturnValue = Replace(strReturnValue, EDI_SEP_DATA_ELEMENT, EDI_SEP_RELEASE_CHARACTER & EDI_SEP_DATA_ELEMENT)
    End If
    ReplaceSpecialCharacters = strReturnValue
End Function
