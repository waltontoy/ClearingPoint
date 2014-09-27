VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfigWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarCompanyName As String 'local copy
Private mvarLogicalID As String 'local copy
Private mvarCustomsRegNo As String 'local copy
Private mvarAccount49 As String 'local copy
Private mvarCustomsOffice As String 'local copy
Private mvarLanguageofDeclaration As String 'local copy
Private mvarPlaceofLoading As String 'local copy
Private mvarSendingPassword As String 'local copy
Private mvarPrintingPassword As String 'local copy
Private mvarAdministratorName As String 'local copy
Private mvarPassword As String 'local copy
Private mvarIPAddress As String 'local copy
Private mvarGatewayName As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPrinterName As String 'local copy
Private mvarPrintingProgram As String 'local copy

Public Property Let PrintingProgram(ByVal vData As String)
    mvarPrintingProgram = vData
End Property

Public Property Get PrintingProgram() As String
    PrintingProgram = mvarPrintingProgram
End Property

Public Property Let PrinterName(ByVal vData As String)
    mvarPrinterName = vData
End Property

Public Property Get PrinterName() As String
    PrinterName = mvarPrinterName
End Property

Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property

Public Property Get UserName() As String
    UserName = mvarUserName
End Property

Public Property Let GatewayName(ByVal vData As String)
    mvarGatewayName = vData
End Property

Public Property Get GatewayName() As String
    GatewayName = mvarGatewayName
End Property

Public Property Let IPAddress(ByVal vData As String)
    mvarIPAddress = vData
End Property

Public Property Get IPAddress() As String
    IPAddress = mvarIPAddress
End Property

Public Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property

Public Property Get Password() As String
    Password = mvarPassword
End Property

Public Property Let AdministratorName(ByVal vData As String)
    mvarAdministratorName = vData
End Property

Public Property Get AdministratorName() As String
    AdministratorName = mvarAdministratorName
End Property

Public Property Let PrintingPassword(ByVal vData As String)
    mvarPrintingPassword = vData
End Property

Public Property Get PrintingPassword() As String
    PrintingPassword = mvarPrintingPassword
End Property

Public Property Let SendingPassword(ByVal vData As String)
    mvarSendingPassword = vData
End Property

Public Property Get SendingPassword() As String
    SendingPassword = mvarSendingPassword
End Property

Public Property Let PlaceofLoading(ByVal vData As String)
    mvarPlaceofLoading = vData
End Property

Public Property Get PlaceofLoading() As String
    PlaceofLoading = mvarPlaceofLoading
End Property

Public Property Let LanguageofDeclaration(ByVal vData As String)
    mvarLanguageofDeclaration = vData
End Property

Public Property Get LanguageofDeclaration() As String
    LanguageofDeclaration = mvarLanguageofDeclaration
End Property

Public Property Let CustomsOffice(ByVal vData As String)
    mvarCustomsOffice = vData
End Property

Public Property Get CustomsOffice() As String
    CustomsOffice = mvarCustomsOffice
End Property

Public Property Let Account49(ByVal vData As String)
    mvarAccount49 = vData
End Property

Public Property Get Account49() As String
    Account49 = mvarAccount49
End Property

Public Property Let CustomsRegNo(ByVal vData As String)
    mvarCustomsRegNo = vData
End Property

Public Property Get CustomsRegNo() As String
    CustomsRegNo = mvarCustomsRegNo
End Property

Public Property Let LogicalID(ByVal vData As String)
    mvarLogicalID = vData
End Property

Public Property Get LogicalID() As String
    LogicalID = mvarLogicalID
End Property

Public Property Let CompanyName(ByVal vData As String)
    mvarCompanyName = vData
End Property

Public Property Get CompanyName() As String
    CompanyName = mvarCompanyName
End Property

Public Function ShowConfigurationWizard(ByRef DataSourceProperties As CDataSourceProperties, _
                                        ByRef CallingForm As Object) As Boolean
    
'Public Function ShowConfigurationWizard(ByVal MdbPath As String, _
                                        ByRef CallingForm As Object) As Boolean
                                        
    ' Open connections
    '<<< dandan 112306
    '<<< Update with database
    'Set g_conSadbel = OpenDatabase(MdbPath & "\mdb_sadbel.mdb")
    ' MUCP-159 - Start
    'OpenDAODatabase g_conSadbel, MdbPath, "mdb_sadbel.mdb"
    ADOConnectDB g_conSadbel, DataSourceProperties, DBInstanceType_DATABASE_SADBEL

    'Set g_conTemplate = OpenDatabase(MdbPath & "\TemplateCP.mdb")
    'OpenDAODatabase g_conTemplate, MdbPath, "TemplateCP.mdb"
    ADOConnectDB g_conTemplate, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
    
    'Set g_conScheduler = OpenDatabase(MdbPath & "\mdb_scheduler.mdb")
    'OpenDAODatabase g_conScheduler, MdbPath, "mdb_scheduler.mdb"
    ADOConnectDB g_conScheduler, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
    ' MUCP-159 - End
    
    ' Start configuration wizard
    ShowConfigurationWizard = frmConfigWizardPage1.ShowConfigUtility(CallingForm)
    
    
    ' Close connection objects and destroy
    ADODisconnectDB g_conSadbel
    'g_conSadbel.Close
    'Set g_conSadbel = Nothing
    
    ADODisconnectDB g_conTemplate
    'g_conTemplate.Close
    'Set g_conTemplate = Nothing
    
    ADODisconnectDB g_conScheduler
    'g_conScheduler.Close
    'Set g_conScheduler = Nothing
    
End Function
