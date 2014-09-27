VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub ShowReports(ByRef SummaryReport As DDActiveReports2.ActiveReport, ByVal Connections As PCubeLibDBReg.CConnections, ByVal Language As String, ByVal TaricUse As Integer, ByVal AppVersion As String, ByVal LicenseeName As String, ByVal ResourceHandler As Long, ByVal strMDBpath, ByVal strPrinterName As String, ByVal LicCompanyName As String, ByVal LicIsDemo As Boolean, ByVal UserID As String)
    frmSummaryReports.My_Load SummaryReport, Connections, Language, TaricUse, AppVersion, LicenseeName, ResourceHandler, strMDBpath, strPrinterName, LicCompanyName, LicIsDemo, UserID
End Sub
