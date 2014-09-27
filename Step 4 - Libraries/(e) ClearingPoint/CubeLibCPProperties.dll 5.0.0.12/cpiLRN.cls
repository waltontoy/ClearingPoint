VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiLRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_Branch_Code As String ' 1
Private mvarFIELD_Year As String ' 2
Private mvarFIELD_Last_Used_Counter As Single ' 3
Private mvarFIELD_Restart_Next_Year As Boolean ' 4
Private mvarFIELD_Starting_Counter As Single ' 5

'  1
Public Property Let FIELD_Branch_Code(ByVal vData As String)
   mvarFIELD_Branch_Code = vData
End Property

Public Property Get FIELD_Branch_Code() As String
   FIELD_Branch_Code = mvarFIELD_Branch_Code
End Property

'  2
Public Property Let FIELD_Year(ByVal vData As String)
   mvarFIELD_Year = vData
End Property

Public Property Get FIELD_Year() As String
   FIELD_Year = mvarFIELD_Year
End Property

'  3
Public Property Let FIELD_Last_Used_Counter(ByVal vData As Single)
   mvarFIELD_Last_Used_Counter = vData
End Property

Public Property Get FIELD_Last_Used_Counter() As Single
   FIELD_Last_Used_Counter = mvarFIELD_Last_Used_Counter
End Property

'  4
Public Property Let FIELD_Restart_Next_Year(ByVal vData As Boolean)
   mvarFIELD_Restart_Next_Year = vData
End Property

Public Property Get FIELD_Restart_Next_Year() As Boolean
   FIELD_Restart_Next_Year = mvarFIELD_Restart_Next_Year
End Property

'  5
Public Property Let FIELD_Starting_Counter(ByVal vData As Single)
   mvarFIELD_Starting_Counter = vData
End Property

Public Property Get FIELD_Starting_Counter() As Single
   FIELD_Starting_Counter = mvarFIELD_Starting_Counter
End Property


