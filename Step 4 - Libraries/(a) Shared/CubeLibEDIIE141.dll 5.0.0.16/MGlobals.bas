Attribute VB_Name = "MGlobals"
Option Explicit

Public G_CallingForm As Object

Public G_conEdifact As ADODB.Connection
Public G_conData As ADODB.Connection
Public G_conSadbel As ADODB.Connection

Public G_rstFollowUpRequest As ADODB.Recordset
Public G_rstMasterEDI As ADODB.Recordset
Public G_rstNCTSData As ADODB.Recordset

Public Const EDI_SEP_SEGMENT                As String = "'"
Public Const EDI_SEP_COMPOSITE_DATA_ELEMENT As String = ":"
Public Const EDI_SEP_DATA_ELEMENT           As String = "+"
Public Const EDI_SEP_RELEASE_CHARACTER      As String = "?"

Public Const G_MAIN_PASSWORD = "wack2"
