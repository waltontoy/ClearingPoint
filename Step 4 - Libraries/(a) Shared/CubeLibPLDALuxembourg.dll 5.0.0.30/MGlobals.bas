Attribute VB_Name = "MGlobals"
Option Explicit
    
    Public G_rstMain As ADODB.Recordset
    
    Public G_rstHeader As ADODB.Recordset
    Public G_rstHeaderHandelaars As ADODB.Recordset
    Public G_rstHeaderZegels As ADODB.Recordset
    
    Public G_rstDetails As ADODB.Recordset
    Public G_rstDetailsBijzondere As ADODB.Recordset
    Public G_rstDetailsBerekeningsEenheden As ADODB.Recordset
    Public G_rstDetailsContainer As ADODB.Recordset
    Public G_rstDetailsHandelaars As ADODB.Recordset
    Public G_rstDetailsDocumenten As ADODB.Recordset
    Public G_rstDetailsZelf As ADODB.Recordset
        
    Public G_strXMLUserName As String
    Public G_strMessageSender As String
    Public G_strMessageRecipient As String
    Public G_strCancelReason As String
    Public G_strIEFunctionCode As String
    Public G_lngTestOnly As Long
    
    Public G_rstLogIDFields As ADODB.Recordset '03242008
    
    Public Const G_BlnDoNotIncludeIfEmpty = True
