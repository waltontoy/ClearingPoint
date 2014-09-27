VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "cpiDATA_NCTS_PCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_DATA_NCTS_PCI_ID As Long ' 1
Private mvarFIELD_DATA_NCTS_MSG_ID As Long ' 2
Private mvarFIELD_NCTS_IEM_TMS_ID As Long ' 3
Private mvarFIELD_DATA_NCTS_PCI_ParentID As Long ' 4
Private mvarFIELD_DATA_NCTS_PCI_Seq1 As String ' 5
Private mvarFIELD_DATA_NCTS_PCI_Seq2 As String ' 6
Private mvarFIELD_DATA_NCTS_PCI_Seq3 As String ' 7
Private mvarFIELD_DATA_NCTS_PCI_Seq4 As String ' 8
Private mvarFIELD_DATA_NCTS_PCI_Seq5 As String ' 9
Private mvarFIELD_DATA_NCTS_PCI_Seq6 As String ' 10
Private mvarFIELD_DATA_NCTS_PCI_Seq7 As String ' 11
Private mvarFIELD_DATA_NCTS_PCI_Seq8 As String ' 12
Private mvarFIELD_DATA_NCTS_PCI_Seq9 As String ' 13
Private mvarFIELD_DATA_NCTS_PCI_Seq10 As String ' 14
Private mvarFIELD_DATA_NCTS_PCI_Seq11 As String ' 15
Private mvarFIELD_DATA_NCTS_PCI_Seq12 As String ' 16
Private mvarFIELD_DATA_NCTS_PCI_Seq13 As String ' 17
Private mvarFIELD_DATA_NCTS_PCI_Seq14 As String ' 18
Private mvarFIELD_DATA_NCTS_PCI_Seq15 As String ' 19
Private mvarFIELD_DATA_NCTS_PCI_Instance As Long ' 20

'  1
Public Property Let FIELD_DATA_NCTS_PCI_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_PCI_ID = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_ID() As Long
   FIELD_DATA_NCTS_PCI_ID = mvarFIELD_DATA_NCTS_PCI_ID
End Property

'  2
Public Property Let FIELD_DATA_NCTS_MSG_ID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_MSG_ID = vData
End Property

Public Property Get FIELD_DATA_NCTS_MSG_ID() As Long
   FIELD_DATA_NCTS_MSG_ID = mvarFIELD_DATA_NCTS_MSG_ID
End Property

'  3
Public Property Let FIELD_NCTS_IEM_TMS_ID(ByVal vData As Long)
   mvarFIELD_NCTS_IEM_TMS_ID = vData
End Property

Public Property Get FIELD_NCTS_IEM_TMS_ID() As Long
   FIELD_NCTS_IEM_TMS_ID = mvarFIELD_NCTS_IEM_TMS_ID
End Property

'  4
Public Property Let FIELD_DATA_NCTS_PCI_ParentID(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_PCI_ParentID = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_ParentID() As Long
   FIELD_DATA_NCTS_PCI_ParentID = mvarFIELD_DATA_NCTS_PCI_ParentID
End Property

'  5
Public Property Let FIELD_DATA_NCTS_PCI_Seq1(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq1 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq1() As String
   FIELD_DATA_NCTS_PCI_Seq1 = mvarFIELD_DATA_NCTS_PCI_Seq1
End Property

'  6
Public Property Let FIELD_DATA_NCTS_PCI_Seq2(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq2 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq2() As String
   FIELD_DATA_NCTS_PCI_Seq2 = mvarFIELD_DATA_NCTS_PCI_Seq2
End Property

'  7
Public Property Let FIELD_DATA_NCTS_PCI_Seq3(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq3 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq3() As String
   FIELD_DATA_NCTS_PCI_Seq3 = mvarFIELD_DATA_NCTS_PCI_Seq3
End Property

'  8
Public Property Let FIELD_DATA_NCTS_PCI_Seq4(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq4 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq4() As String
   FIELD_DATA_NCTS_PCI_Seq4 = mvarFIELD_DATA_NCTS_PCI_Seq4
End Property

'  9
Public Property Let FIELD_DATA_NCTS_PCI_Seq5(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq5 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq5() As String
   FIELD_DATA_NCTS_PCI_Seq5 = mvarFIELD_DATA_NCTS_PCI_Seq5
End Property

'  10
Public Property Let FIELD_DATA_NCTS_PCI_Seq6(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq6 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq6() As String
   FIELD_DATA_NCTS_PCI_Seq6 = mvarFIELD_DATA_NCTS_PCI_Seq6
End Property

'  11
Public Property Let FIELD_DATA_NCTS_PCI_Seq7(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq7 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq7() As String
   FIELD_DATA_NCTS_PCI_Seq7 = mvarFIELD_DATA_NCTS_PCI_Seq7
End Property

'  12
Public Property Let FIELD_DATA_NCTS_PCI_Seq8(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq8 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq8() As String
   FIELD_DATA_NCTS_PCI_Seq8 = mvarFIELD_DATA_NCTS_PCI_Seq8
End Property

'  13
Public Property Let FIELD_DATA_NCTS_PCI_Seq9(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq9 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq9() As String
   FIELD_DATA_NCTS_PCI_Seq9 = mvarFIELD_DATA_NCTS_PCI_Seq9
End Property

'  14
Public Property Let FIELD_DATA_NCTS_PCI_Seq10(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq10 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq10() As String
   FIELD_DATA_NCTS_PCI_Seq10 = mvarFIELD_DATA_NCTS_PCI_Seq10
End Property

'  15
Public Property Let FIELD_DATA_NCTS_PCI_Seq11(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq11 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq11() As String
   FIELD_DATA_NCTS_PCI_Seq11 = mvarFIELD_DATA_NCTS_PCI_Seq11
End Property

'  16
Public Property Let FIELD_DATA_NCTS_PCI_Seq12(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq12 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq12() As String
   FIELD_DATA_NCTS_PCI_Seq12 = mvarFIELD_DATA_NCTS_PCI_Seq12
End Property

'  17
Public Property Let FIELD_DATA_NCTS_PCI_Seq13(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq13 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq13() As String
   FIELD_DATA_NCTS_PCI_Seq13 = mvarFIELD_DATA_NCTS_PCI_Seq13
End Property

'  18
Public Property Let FIELD_DATA_NCTS_PCI_Seq14(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq14 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq14() As String
   FIELD_DATA_NCTS_PCI_Seq14 = mvarFIELD_DATA_NCTS_PCI_Seq14
End Property

'  19
Public Property Let FIELD_DATA_NCTS_PCI_Seq15(ByVal vData As String)
   mvarFIELD_DATA_NCTS_PCI_Seq15 = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Seq15() As String
   FIELD_DATA_NCTS_PCI_Seq15 = mvarFIELD_DATA_NCTS_PCI_Seq15
End Property

'  20
Public Property Let FIELD_DATA_NCTS_PCI_Instance(ByVal vData As Long)
   mvarFIELD_DATA_NCTS_PCI_Instance = vData
End Property

Public Property Get FIELD_DATA_NCTS_PCI_Instance() As Long
   FIELD_DATA_NCTS_PCI_Instance = mvarFIELD_DATA_NCTS_PCI_Instance
End Property


