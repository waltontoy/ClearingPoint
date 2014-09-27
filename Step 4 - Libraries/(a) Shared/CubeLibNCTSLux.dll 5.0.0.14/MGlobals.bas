Attribute VB_Name = "MGlobals"
Option Explicit
'***************************************For CSCLP-439

'Public g_conEdifact As ADODB.Connection
Public g_conEdifact As ADODB.Connection     'Public EdifactDAO As DAO.Database
Public g_conScheduler As ADODB.Connection   'Public G_datScheduler As DAO.Database
Public g_conSADBEL As ADODB.Connection      'Public G_datSADBEL As DAO.Database

Public g_rstSegment As ADODB.Recordset      'Public G_rstSegment As DAO.Recordset
Public g_rstEdifact As ADODB.Recordset      'Public rstEdifact As DAO.Recordset

Public lngctr As Long
Public strTableName As String
Public strDATA_NCTS_ID As String
Public strDATA_NCTS_MSG_ID As String
Public strUniqueCode As String
Public rstUNB As New ADODB.Recordset
Public rstUNH As New ADODB.Recordset
Public rstRFF As New ADODB.Recordset
Public rstBGM As New ADODB.Recordset
Public rstLOC As New ADODB.Recordset
Public rstTDT As New ADODB.Recordset
Public rstGIS As New ADODB.Recordset
Public rstFTX As New ADODB.Recordset
Public rstCNT As New ADODB.Recordset
Public rstMEA As New ADODB.Recordset
Public rstNAD As New ADODB.Recordset
Public rstDTM As New ADODB.Recordset
Public rstPAC As New ADODB.Recordset
Public rstPCI As New ADODB.Recordset
Public rstDetailCount As ADODB.Recordset
Public rstDetailCST As New ADODB.Recordset
Public rstDetailFTX As New ADODB.Recordset
Public rstDetailTOD As New ADODB.Recordset
Public rstDetailMEA As New ADODB.Recordset
Public rstDetailRFF As New ADODB.Recordset
Public rstDetailDOC As New ADODB.Recordset
Public rstDetailPAC As New ADODB.Recordset
Public rstDetailPCI As New ADODB.Recordset
Public rstDetailLOC As New ADODB.Recordset
Public rstDetailGIR As New ADODB.Recordset
Public conEdifactADO As ADODB.Connection
'****************************************************

Public G_strUniqueCode As String

'Public G_strMdbPath As String

'************************************************************************************
'Departure / Cancellation
'************************************************************************************
Public G_clsEDIDeparture As EdifactMessage
Public G_rstDepartureMap As ADODB.Recordset
Public G_strEDIMessage As String
Public G_strEDICancellation As String

'************************************************************************************

'************************************************************************************
'Arrival
'************************************************************************************
Public G_clsEDIArrival As PCubeLibEDIArrivals.cpiMessage
Public G_clsIE44Arrival As PCubeLibEDIArrivals.cpiIE44Message
Public G_strHeaderKey As String
Public G_strCustomOfcKey As String
Public G_strTraderKey As String
'************************************************************************************

Public G_strSendMode As String
Public G_strLogicalId As String

Public Const G_CONST_EDINCTS1_TYPE = "EDI NCTS"
Public Const G_CONST_NCTS1_TYPE = "Transit NCTS"
Public Const G_CONST_NCTS1_SHEET = "NCTS1Sheet"
Public Const G_CONST_NCTS2_TYPE = "Combined NCTS"

Public Enum eTabType
    eTab_Header = 1
    eTab_Detail = 2
End Enum

Public Enum SealsCreationModes
    SealsCreationMode_InvalidInput = 0
    SealsCreationMode_NoSeal
    SealsCreationMode_AEValueOnly
    SealsCreationMode_Repetition
    SealsCreationMode_Increment
End Enum

Public Enum IE29Values
    enuIE29Val_NotFromIE29 = 1
    
    enuIEVal_IE43_Marks_And_Numbers
    enuIEVal_IE43_Number_of_Packages
    enuIEVal_IE43_Kind_of_Packages
    enuIEVal_IE43_Container_Numbers
    enuIEVal_IE43_Description_of_Goods
    enuIEVal_IE43_Sensitivity_Code
    enuIEVal_IE43_Sensitive_Quantity
    enuIEVal_IE43_Country_of_Dispatch_Export
    enuIEVal_IE43_Country_of_Destination
    enuIEVal_IE43_CO_Departure                      'LOC+118(2)
    enuIEVal_IE43_Gross_Mass
    enuIEVal_IE43_Net_Mass
    enuIEVal_IE43_Additional_Information
    enuIEVal_IE43_Consignor_TIN
    enuIEVal_IE43_Consignor_Name
    enuIEVal_IE43_Consignor_Street_And_Number
    enuIEVal_IE43_Consignor_Postal_Code
    enuIEVal_IE43_Consignor_City
    enuIEVal_IE43_Consignor_Country
    enuIEVal_IE43_Consignee_TIN
    enuIEVal_IE43_Consignee_Name
    enuIEVal_IE43_Consignee_Street_And_Number
    enuIEVal_IE43_Consignee_Postal_Code
    enuIEVal_IE43_Consignee_City
    enuIEVal_IE43_Consignee_Country
    enuIEVal_IE43_Document_Type
    enuIEVal_IE43_Document_Reference
    enuIEVal_IE43_Document_Complement_Information
    enuIEVal_IE43_Detail_Number
    enuIEVal_IE43_Commodity_Code
    
    enuIE29Val_MessageIdentification                'UNH(1)
    enuIE29Val_ReferenceNumber                      'BGM(5)
    enuIE29Val_AuthorizedLocationOfGoods            'LOC+14(6)
    enuIE29Val_DeclarationPlace                     'LOC+91(5)
    enuIE29Val_COReferencNumber                     'LOC+168(2) - CO = Customs Office
    enuIE29Val_COName                               'LOC+168(5) - CO = Customs Office
    enuIE29Val_COCountry                            'LOC+168(6) - CO = Customs Office
    enuIE29Val_COStreetAndNumber                    'LOC+168(9) - CO = Customs Office
    enuIE29Val_COPostalCode                         'LOC+168(10) - CO = Customs Office
    enuIE29Val_COCity                               'LOC+168(13) - CO = Customs Office
    enuIE29Val_COLanguage                           'LOC+168(14) - CO = Customs Office
    enuIE29Val_DateApproval                         'DTM+148(2)
    enuIE29Val_DateIssuance                         'DTM+182(2)
    enuIE29Val_DateControl                          'DTM+9(2)
    enuIEVal_IE29_DateLimitTransit                  'DTM+268(2)
    enuIE29Val_ReturnCopy                           'GIS 62(2)
    enuIE29Val_BindingItinerary                     'FTX+ABL(6)
    enuIE29Val_NotValidForEC                        'PCI+19(2)
    enuIE29Val_TPName                               'NAD+AF(10) - TP = Transit Principal
    enuIE29Val_TPStreetAndNumber                    'NAD+AF(16) - TP = Transit Principal
    enuIE29Val_TPCity                               'NAD+AF(20) - TP = Transit Principal
    enuIE29Val_TPPostalCode                         'NAD+AF(22) - TP = Transit Principal
    enuIE29Val_TPCountry                            'NAD+AF(23) - TP = Transit Principal
    enuIE29Val_ControlledBy                         'NAD+EI(2)
    
    enuIEVal_IE28_TPTIN                             'NAD+AF(2)  - TP = Transit Principal
    enuIEVal_IE28_TPName                            'NAD+AF(10) - TP = Transit Principal
    enuIEVal_IE28_TPStreetAndNumber                 'NAD+AF(16) - TP = Transit Principal
    enuIEVal_IE28_TPCity                            'NAD+AF(20) - TP = Transit Principal
    enuIEVal_IE28_TPPostalCode                      'NAD+AF(22) - TP = Transit Principal
    enuIEVal_IE28_TPCountry                         'NAD+AF(23) - TP = Transit Principal
End Enum

