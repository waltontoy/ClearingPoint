VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiLOGICAL_ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_LOGID_DESCRIPTION As String ' 1
Private mvarFIELD_SEND_OPERATIONAL_CORR As String ' 2
Private mvarFIELD_SEND_OPERATIONAL_LOGID As String ' 3
Private mvarFIELD_SEND_OPERATIONAL_PASS As String ' 4
Private mvarFIELD_SEND_TEST_CORR As String ' 5
Private mvarFIELD_SEND_TEST_LOGID As String ' 6
Private mvarFIELD_SEND_TEST_PASS As String ' 7
Private mvarFIELD_PRINT_OPERATIONAL_CORR As String ' 8
Private mvarFIELD_PRINT_OPERATIONAL_LOGID As String ' 9
Private mvarFIELD_PRINT_OPERATIONAL_PASS As String ' 10
Private mvarFIELD_PRINT_TEST_CORR As String ' 11
Private mvarFIELD_PRINT_TEST_LOGID As String ' 12
Private mvarFIELD_PRINT_TEST_PASS As String ' 13
Private mvarFIELD_A1 As String ' 14
Private mvarFIELD_A2 As String ' 15
Private mvarFIELD_HISTORY As String ' 16
Private mvarFIELD_VAT As String ' 17
Private mvarFIELD_TIN As String ' 18
Private mvarFIELD_USAGE As Byte ' 19
Private mvarFIELD_LRN_USAGE As Byte ' 20
Private mvarFIELD_BRANCH As String ' 21
Private mvarFIELD_PROCEDURE As Byte ' 22
Private mvarFIELD_LOCATION As String ' 23
Private mvarFIELD_VOLGBRIEFJE_PRINTER As String ' 24
Private mvarFIELD_NCTS_DOC_PRINTER As String ' 25
Private mvarFIELD_PRINT_MODE As Byte ' 26
Private mvarFIELD_SFTP_ADDRESS As String ' 27
Private mvarFIELD_EMAIL_ADDRESS As String ' 28
Private mvarFIELD_SEND_EDI_RECIPIENT_OPERATIONAL As String ' 29
Private mvarFIELD_SEND_EDI_RECIPIENT_TEST As String ' 30
Private mvarFIELD_LAST_EDI_REFERENCE As String ' 31

'  1
Public Property Let FIELD_LOGID_DESCRIPTION(ByVal vData As String)
   mvarFIELD_LOGID_DESCRIPTION = vData
End Property

Public Property Get FIELD_LOGID_DESCRIPTION() As String
   FIELD_LOGID_DESCRIPTION = mvarFIELD_LOGID_DESCRIPTION
End Property

'  2
Public Property Let FIELD_SEND_OPERATIONAL_CORR(ByVal vData As String)
   mvarFIELD_SEND_OPERATIONAL_CORR = vData
End Property

Public Property Get FIELD_SEND_OPERATIONAL_CORR() As String
   FIELD_SEND_OPERATIONAL_CORR = mvarFIELD_SEND_OPERATIONAL_CORR
End Property

'  3
Public Property Let FIELD_SEND_OPERATIONAL_LOGID(ByVal vData As String)
   mvarFIELD_SEND_OPERATIONAL_LOGID = vData
End Property

Public Property Get FIELD_SEND_OPERATIONAL_LOGID() As String
   FIELD_SEND_OPERATIONAL_LOGID = mvarFIELD_SEND_OPERATIONAL_LOGID
End Property

'  4
Public Property Let FIELD_SEND_OPERATIONAL_PASS(ByVal vData As String)
   mvarFIELD_SEND_OPERATIONAL_PASS = vData
End Property

Public Property Get FIELD_SEND_OPERATIONAL_PASS() As String
   FIELD_SEND_OPERATIONAL_PASS = mvarFIELD_SEND_OPERATIONAL_PASS
End Property

'  5
Public Property Let FIELD_SEND_TEST_CORR(ByVal vData As String)
   mvarFIELD_SEND_TEST_CORR = vData
End Property

Public Property Get FIELD_SEND_TEST_CORR() As String
   FIELD_SEND_TEST_CORR = mvarFIELD_SEND_TEST_CORR
End Property

'  6
Public Property Let FIELD_SEND_TEST_LOGID(ByVal vData As String)
   mvarFIELD_SEND_TEST_LOGID = vData
End Property

Public Property Get FIELD_SEND_TEST_LOGID() As String
   FIELD_SEND_TEST_LOGID = mvarFIELD_SEND_TEST_LOGID
End Property

'  7
Public Property Let FIELD_SEND_TEST_PASS(ByVal vData As String)
   mvarFIELD_SEND_TEST_PASS = vData
End Property

Public Property Get FIELD_SEND_TEST_PASS() As String
   FIELD_SEND_TEST_PASS = mvarFIELD_SEND_TEST_PASS
End Property

'  8
Public Property Let FIELD_PRINT_OPERATIONAL_CORR(ByVal vData As String)
   mvarFIELD_PRINT_OPERATIONAL_CORR = vData
End Property

Public Property Get FIELD_PRINT_OPERATIONAL_CORR() As String
   FIELD_PRINT_OPERATIONAL_CORR = mvarFIELD_PRINT_OPERATIONAL_CORR
End Property

'  9
Public Property Let FIELD_PRINT_OPERATIONAL_LOGID(ByVal vData As String)
   mvarFIELD_PRINT_OPERATIONAL_LOGID = vData
End Property

Public Property Get FIELD_PRINT_OPERATIONAL_LOGID() As String
   FIELD_PRINT_OPERATIONAL_LOGID = mvarFIELD_PRINT_OPERATIONAL_LOGID
End Property

'  10
Public Property Let FIELD_PRINT_OPERATIONAL_PASS(ByVal vData As String)
   mvarFIELD_PRINT_OPERATIONAL_PASS = vData
End Property

Public Property Get FIELD_PRINT_OPERATIONAL_PASS() As String
   FIELD_PRINT_OPERATIONAL_PASS = mvarFIELD_PRINT_OPERATIONAL_PASS
End Property

'  11
Public Property Let FIELD_PRINT_TEST_CORR(ByVal vData As String)
   mvarFIELD_PRINT_TEST_CORR = vData
End Property

Public Property Get FIELD_PRINT_TEST_CORR() As String
   FIELD_PRINT_TEST_CORR = mvarFIELD_PRINT_TEST_CORR
End Property

'  12
Public Property Let FIELD_PRINT_TEST_LOGID(ByVal vData As String)
   mvarFIELD_PRINT_TEST_LOGID = vData
End Property

Public Property Get FIELD_PRINT_TEST_LOGID() As String
   FIELD_PRINT_TEST_LOGID = mvarFIELD_PRINT_TEST_LOGID
End Property

'  13
Public Property Let FIELD_PRINT_TEST_PASS(ByVal vData As String)
   mvarFIELD_PRINT_TEST_PASS = vData
End Property

Public Property Get FIELD_PRINT_TEST_PASS() As String
   FIELD_PRINT_TEST_PASS = mvarFIELD_PRINT_TEST_PASS
End Property

'  14
Public Property Let FIELD_A1(ByVal vData As String)
   mvarFIELD_A1 = vData
End Property

Public Property Get FIELD_A1() As String
   FIELD_A1 = mvarFIELD_A1
End Property

'  15
Public Property Let FIELD_A2(ByVal vData As String)
   mvarFIELD_A2 = vData
End Property

Public Property Get FIELD_A2() As String
   FIELD_A2 = mvarFIELD_A2
End Property

'  16
Public Property Let FIELD_HISTORY(ByVal vData As String)
   mvarFIELD_HISTORY = vData
End Property

Public Property Get FIELD_HISTORY() As String
   FIELD_HISTORY = mvarFIELD_HISTORY
End Property

'  17
Public Property Let FIELD_VAT(ByVal vData As String)
   mvarFIELD_VAT = vData
End Property

Public Property Get FIELD_VAT() As String
   FIELD_VAT = mvarFIELD_VAT
End Property

'  18
Public Property Let FIELD_TIN(ByVal vData As String)
   mvarFIELD_TIN = vData
End Property

Public Property Get FIELD_TIN() As String
   FIELD_TIN = mvarFIELD_TIN
End Property

'  19
Public Property Let FIELD_USAGE(ByVal vData As Byte)
   mvarFIELD_USAGE = vData
End Property

Public Property Get FIELD_USAGE() As Byte
   FIELD_USAGE = mvarFIELD_USAGE
End Property

'  20
Public Property Let FIELD_LRN_USAGE(ByVal vData As Byte)
   mvarFIELD_LRN_USAGE = vData
End Property

Public Property Get FIELD_LRN_USAGE() As Byte
   FIELD_LRN_USAGE = mvarFIELD_LRN_USAGE
End Property

'  21
Public Property Let FIELD_BRANCH(ByVal vData As String)
   mvarFIELD_BRANCH = vData
End Property

Public Property Get FIELD_BRANCH() As String
   FIELD_BRANCH = mvarFIELD_BRANCH
End Property

'  22
Public Property Let FIELD_PROCEDURE(ByVal vData As Byte)
   mvarFIELD_PROCEDURE = vData
End Property

Public Property Get FIELD_PROCEDURE() As Byte
   FIELD_PROCEDURE = mvarFIELD_PROCEDURE
End Property

'  23
Public Property Let FIELD_LOCATION(ByVal vData As String)
   mvarFIELD_LOCATION = vData
End Property

Public Property Get FIELD_LOCATION() As String
   FIELD_LOCATION = mvarFIELD_LOCATION
End Property

'  24
Public Property Let FIELD_VOLGBRIEFJE_PRINTER(ByVal vData As String)
   mvarFIELD_VOLGBRIEFJE_PRINTER = vData
End Property

Public Property Get FIELD_VOLGBRIEFJE_PRINTER() As String
   FIELD_VOLGBRIEFJE_PRINTER = mvarFIELD_VOLGBRIEFJE_PRINTER
End Property

'  25
Public Property Let FIELD_NCTS_DOC_PRINTER(ByVal vData As String)
   mvarFIELD_NCTS_DOC_PRINTER = vData
End Property

Public Property Get FIELD_NCTS_DOC_PRINTER() As String
   FIELD_NCTS_DOC_PRINTER = mvarFIELD_NCTS_DOC_PRINTER
End Property

'  26
Public Property Let FIELD_PRINT_MODE(ByVal vData As Byte)
   mvarFIELD_PRINT_MODE = vData
End Property

Public Property Get FIELD_PRINT_MODE() As Byte
   FIELD_PRINT_MODE = mvarFIELD_PRINT_MODE
End Property

'  27
Public Property Let FIELD_SFTP_ADDRESS(ByVal vData As String)
   mvarFIELD_SFTP_ADDRESS = vData
End Property

Public Property Get FIELD_SFTP_ADDRESS() As String
   FIELD_SFTP_ADDRESS = mvarFIELD_SFTP_ADDRESS
End Property

'  28
Public Property Let FIELD_EMAIL_ADDRESS(ByVal vData As String)
   mvarFIELD_EMAIL_ADDRESS = vData
End Property

Public Property Get FIELD_EMAIL_ADDRESS() As String
   FIELD_EMAIL_ADDRESS = mvarFIELD_EMAIL_ADDRESS
End Property

'  29
Public Property Let FIELD_SEND_EDI_RECIPIENT_OPERATIONAL(ByVal vData As String)
   mvarFIELD_SEND_EDI_RECIPIENT_OPERATIONAL = vData
End Property

Public Property Get FIELD_SEND_EDI_RECIPIENT_OPERATIONAL() As String
   FIELD_SEND_EDI_RECIPIENT_OPERATIONAL = mvarFIELD_SEND_EDI_RECIPIENT_OPERATIONAL
End Property

'  30
Public Property Let FIELD_SEND_EDI_RECIPIENT_TEST(ByVal vData As String)
   mvarFIELD_SEND_EDI_RECIPIENT_TEST = vData
End Property

Public Property Get FIELD_SEND_EDI_RECIPIENT_TEST() As String
   FIELD_SEND_EDI_RECIPIENT_TEST = mvarFIELD_SEND_EDI_RECIPIENT_TEST
End Property

'  31
Public Property Let FIELD_LAST_EDI_REFERENCE(ByVal vData As String)
   mvarFIELD_LAST_EDI_REFERENCE = vData
End Property

Public Property Get FIELD_LAST_EDI_REFERENCE() As String
   FIELD_LAST_EDI_REFERENCE = mvarFIELD_LAST_EDI_REFERENCE
End Property


