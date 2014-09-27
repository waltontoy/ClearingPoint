VERSION 5.00
Begin VB.Form FCubeLibDBRegUnit 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   5955
   ClientTop       =   5265
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6585
End
Attribute VB_Name = "FCubeLibDBRegUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim conSQL As ADODB.Connection
    Dim conModel As ADODB.Connection
    Dim conTarget As ADODB.Connection
    
    Dim objConProps As CConnectionProperties
    
    Dim strNewDB As String
    Dim blnSuccess As Boolean
    
    strNewDB = "TemplateCP_AA.mdb"
    
    Set objConProps = GetDataSourceProperties(App.Path)
    
    blnSuccess = ADOXCreateDatabase(objConProps, strNewDB)
    
    blnSuccess = ADOConnectDB(conTarget, objConProps, DBInstanceType_DATABASE_OTHER, , strNewDB)
    blnSuccess = ADOConnectDB(conModel, objConProps, DBInstanceType_DATABASE_TEMPLATE)
    
    blnSuccess = ADOXSyncDatabase(conTarget, conModel)
    
    'CreateLinkedTable objConProps, DBInstanceType_DATABASE_TEMPLATE, "tblPLDAImport", DBInstanceType_DATABASE_SADBEL, "PLDA IMPORT"
    
    'ADOConnectDB conSQL, objConProps, DBInstanceType_DATABASE_TEMPLATE
    
    'ADOConnectDB conSQL, objConProps, DBInstanceType_DATABASE_TEMPLATE, , True
    
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", "wack2", "TemplateCP", "sa", True
'''''
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", , "TemplateCP"
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", , "TemplateCP", , True
'''''
'''''
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", "wack2", "TemplateCP", "sa"
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", "wack2", "TemplateCP", "sa", True
'''''
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", , "TemplateCP"
'''''
'''''    ADOConnectDB conSQL, [SQL Server Provider], "WINXPSP3\SQLEXPRESS", , "TemplateCP", , True
End Sub
