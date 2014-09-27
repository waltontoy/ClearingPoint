Attribute VB_Name = "MGlobals"
Option Explicit

Public G_MyStore As Store
Public G_MyCertificate As Certificate
Public G_MyCertificates As Certificates
Public G_EncodingType As Utilities

Public G_strMdbPath As String
Public G_conDigiSign As New ADODB.Connection
Public G_rstDigiSign As New ADODB.Recordset
Public G_rstSadbel As New ADODB.Recordset

Public Const G_MAIN_PASSWORD = "wack2"
'Public Const G_DIGISIGN_DATABASE_NAME = "mdb_digisign.mdb"
Public Const G_DIGISIGN_DATABASE_NAME = "mdb_sadbel.mdb"

Public Enum SEARCH_TYPE
    enuXMLDesc = 1
    enuFieldTable = 2
    enuFieldName = 3
    enuFieldType = 4
End Enum

