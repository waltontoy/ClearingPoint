VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiSETUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_ENGLISH As Boolean ' 1
Private mvarFIELD_FRENCH As Boolean ' 2
Private mvarFIELD_DUTCH As Boolean ' 3
Private mvarFIELD_CURRENT As Integer ' 4
Private mvarFIELD_LAST_USER As String ' 5
Private mvarFIELD_USER_NO_SERIES As Long ' 6
Private mvarFIELD_WITH_SECURITY As Boolean ' 7
Private mvarFIELD_EDIT_TIME As Date ' 8
Private mvarFIELD_REPERTORIUM_FOR_EXPORT As Boolean ' 9
Private mvarFIELD_SENT_TIME As Date ' 10
Private mvarFIELD_TREE_TIME As Date ' 11
Private mvarFIELD_LAST_COMPACT_DATE As Date ' 12
Private mvarFIELD_NUMBER_OF_USERS As Integer ' 13
Private mvarFIELD_FIRSTRUN As Boolean ' 14
Private mvarFIELD_Memo_Field As Boolean ' 15
Private mvarFIELD_DataTransferredToNCTS As Boolean ' 16

'  1
Public Property Let FIELD_ENGLISH(ByVal vData As Boolean)
   mvarFIELD_ENGLISH = vData
End Property

Public Property Get FIELD_ENGLISH() As Boolean
   FIELD_ENGLISH = mvarFIELD_ENGLISH
End Property

'  2
Public Property Let FIELD_FRENCH(ByVal vData As Boolean)
   mvarFIELD_FRENCH = vData
End Property

Public Property Get FIELD_FRENCH() As Boolean
   FIELD_FRENCH = mvarFIELD_FRENCH
End Property

'  3
Public Property Let FIELD_DUTCH(ByVal vData As Boolean)
   mvarFIELD_DUTCH = vData
End Property

Public Property Get FIELD_DUTCH() As Boolean
   FIELD_DUTCH = mvarFIELD_DUTCH
End Property

'  4
Public Property Let FIELD_CURRENT(ByVal vData As Integer)
   mvarFIELD_CURRENT = vData
End Property

Public Property Get FIELD_CURRENT() As Integer
   FIELD_CURRENT = mvarFIELD_CURRENT
End Property

'  5
Public Property Let FIELD_LAST_USER(ByVal vData As String)
   mvarFIELD_LAST_USER = vData
End Property

Public Property Get FIELD_LAST_USER() As String
   FIELD_LAST_USER = mvarFIELD_LAST_USER
End Property

'  6
Public Property Let FIELD_USER_NO_SERIES(ByVal vData As Long)
   mvarFIELD_USER_NO_SERIES = vData
End Property

Public Property Get FIELD_USER_NO_SERIES() As Long
   FIELD_USER_NO_SERIES = mvarFIELD_USER_NO_SERIES
End Property

'  7
Public Property Let FIELD_WITH_SECURITY(ByVal vData As Boolean)
   mvarFIELD_WITH_SECURITY = vData
End Property

Public Property Get FIELD_WITH_SECURITY() As Boolean
   FIELD_WITH_SECURITY = mvarFIELD_WITH_SECURITY
End Property

'  8
Public Property Let FIELD_EDIT_TIME(ByVal vData As Date)
   mvarFIELD_EDIT_TIME = vData
End Property

Public Property Get FIELD_EDIT_TIME() As Date
   FIELD_EDIT_TIME = mvarFIELD_EDIT_TIME
End Property

'  9
Public Property Let FIELD_REPERTORIUM_FOR_EXPORT(ByVal vData As Boolean)
   mvarFIELD_REPERTORIUM_FOR_EXPORT = vData
End Property

Public Property Get FIELD_REPERTORIUM_FOR_EXPORT() As Boolean
   FIELD_REPERTORIUM_FOR_EXPORT = mvarFIELD_REPERTORIUM_FOR_EXPORT
End Property

'  10
Public Property Let FIELD_SENT_TIME(ByVal vData As Date)
   mvarFIELD_SENT_TIME = vData
End Property

Public Property Get FIELD_SENT_TIME() As Date
   FIELD_SENT_TIME = mvarFIELD_SENT_TIME
End Property

'  11
Public Property Let FIELD_TREE_TIME(ByVal vData As Date)
   mvarFIELD_TREE_TIME = vData
End Property

Public Property Get FIELD_TREE_TIME() As Date
   FIELD_TREE_TIME = mvarFIELD_TREE_TIME
End Property

'  12
Public Property Let FIELD_LAST_COMPACT_DATE(ByVal vData As Date)
   mvarFIELD_LAST_COMPACT_DATE = vData
End Property

Public Property Get FIELD_LAST_COMPACT_DATE() As Date
   FIELD_LAST_COMPACT_DATE = mvarFIELD_LAST_COMPACT_DATE
End Property

'  13
Public Property Let FIELD_NUMBER_OF_USERS(ByVal vData As Integer)
   mvarFIELD_NUMBER_OF_USERS = vData
End Property

Public Property Get FIELD_NUMBER_OF_USERS() As Integer
   FIELD_NUMBER_OF_USERS = mvarFIELD_NUMBER_OF_USERS
End Property

'  14
Public Property Let FIELD_FIRSTRUN(ByVal vData As Boolean)
   mvarFIELD_FIRSTRUN = vData
End Property

Public Property Get FIELD_FIRSTRUN() As Boolean
   FIELD_FIRSTRUN = mvarFIELD_FIRSTRUN
End Property

'  15
Public Property Let FIELD_Memo_Field(ByVal vData As Boolean)
   mvarFIELD_Memo_Field = vData
End Property

Public Property Get FIELD_Memo_Field() As Boolean
   FIELD_Memo_Field = mvarFIELD_Memo_Field
End Property

'  16
Public Property Let FIELD_DataTransferredToNCTS(ByVal vData As Boolean)
   mvarFIELD_DataTransferredToNCTS = vData
End Property

Public Property Get FIELD_DataTransferredToNCTS() As Boolean
   FIELD_DataTransferredToNCTS = mvarFIELD_DataTransferredToNCTS
End Property


