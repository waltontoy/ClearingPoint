Attribute VB_Name = "MConfigWizard"
Option Explicit
    Public g_blnConfigurationFinished As Boolean
    
    Public g_strAdminUserLabel As String
    
    Public g_rstUser As ADODB.Recordset
    Public g_rstLogID As ADODB.Recordset
    Public g_rstUserLogID As ADODB.Recordset
    Public g_rstSetupSched As ADODB.Recordset
    Public g_rstPrinterDef As ADODB.Recordset
    Public g_rstLogIDSched As ADODB.Recordset
    Public g_rstSetupSadbel As ADODB.Recordset
    
    Public g_rstTemplate As ADODB.Recordset
    
    Public g_conSadbel As ADODB.Connection
    Public g_conTemplate As ADODB.Connection
    Public g_conScheduler As ADODB.Connection
    
'    Public g_rstUser As DAO.Recordset
'    Public g_rstLogID As DAO.Recordset
'    Public g_rstUserLogID As DAO.Recordset
'    Public g_rstSetupSched As DAO.Recordset
'    Public g_rstPrinterDef As DAO.Recordset
'    Public g_rstLogIDSched As DAO.Recordset
'    Public g_rstSetupSadbel As DAO.Recordset
'
'    Public g_rstTemplate As DAO.Recordset
'
'    Public g_conSadbel As DAO.Database
'    Public g_conTemplate As DAO.Database
'    Public g_conScheduler As DAO.Database

    Public clsConfigWizard As CConfigWizard
    Public Const G_Main_Password = "wack2"


