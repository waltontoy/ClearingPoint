VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiSHEET_PROPS_Tbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

Public Key As String
Public Index As Long

Private mvarFIELD_user_no As Long ' 1
Private mvarFIELD_font_name As String ' 2
Private mvarFIELD_bold As Boolean ' 3
Private mvarFIELD_italic As Boolean ' 4
Private mvarFIELD_BC_active_box As Long ' 5
Private mvarFIELD_BC_inactive_box As Long ' 6
Private mvarFIELD_BC_disabled_box As Long ' 7
Private mvarFIELD_BC_werror_active As Long ' 8
Private mvarFIELD_BC_werror_inactive As Long ' 9
Private mvarFIELD_FC_active_box As Long ' 10
Private mvarFIELD_FC_inactive_box As Long ' 11
Private mvarFIELD_FC_disabled_box As Long ' 12
Private mvarFIELD_FC_werror_active As Long ' 13
Private mvarFIELD_FC_werror_inactive As Long ' 14
Private mvarFIELD_size As Integer ' 15

'  1
Public Property Let FIELD_user_no(ByVal vData As Long)
   mvarFIELD_user_no = vData
End Property

Public Property Get FIELD_user_no() As Long
   FIELD_user_no = mvarFIELD_user_no
End Property

'  2
Public Property Let FIELD_font_name(ByVal vData As String)
   mvarFIELD_font_name = vData
End Property

Public Property Get FIELD_font_name() As String
   FIELD_font_name = mvarFIELD_font_name
End Property

'  3
Public Property Let FIELD_bold(ByVal vData As Boolean)
   mvarFIELD_bold = vData
End Property

Public Property Get FIELD_bold() As Boolean
   FIELD_bold = mvarFIELD_bold
End Property

'  4
Public Property Let FIELD_italic(ByVal vData As Boolean)
   mvarFIELD_italic = vData
End Property

Public Property Get FIELD_italic() As Boolean
   FIELD_italic = mvarFIELD_italic
End Property

'  5
Public Property Let FIELD_BC_active_box(ByVal vData As Long)
   mvarFIELD_BC_active_box = vData
End Property

Public Property Get FIELD_BC_active_box() As Long
   FIELD_BC_active_box = mvarFIELD_BC_active_box
End Property

'  6
Public Property Let FIELD_BC_inactive_box(ByVal vData As Long)
   mvarFIELD_BC_inactive_box = vData
End Property

Public Property Get FIELD_BC_inactive_box() As Long
   FIELD_BC_inactive_box = mvarFIELD_BC_inactive_box
End Property

'  7
Public Property Let FIELD_BC_disabled_box(ByVal vData As Long)
   mvarFIELD_BC_disabled_box = vData
End Property

Public Property Get FIELD_BC_disabled_box() As Long
   FIELD_BC_disabled_box = mvarFIELD_BC_disabled_box
End Property

'  8
Public Property Let FIELD_BC_werror_active(ByVal vData As Long)
   mvarFIELD_BC_werror_active = vData
End Property

Public Property Get FIELD_BC_werror_active() As Long
   FIELD_BC_werror_active = mvarFIELD_BC_werror_active
End Property

'  9
Public Property Let FIELD_BC_werror_inactive(ByVal vData As Long)
   mvarFIELD_BC_werror_inactive = vData
End Property

Public Property Get FIELD_BC_werror_inactive() As Long
   FIELD_BC_werror_inactive = mvarFIELD_BC_werror_inactive
End Property

'  10
Public Property Let FIELD_FC_active_box(ByVal vData As Long)
   mvarFIELD_FC_active_box = vData
End Property

Public Property Get FIELD_FC_active_box() As Long
   FIELD_FC_active_box = mvarFIELD_FC_active_box
End Property

'  11
Public Property Let FIELD_FC_inactive_box(ByVal vData As Long)
   mvarFIELD_FC_inactive_box = vData
End Property

Public Property Get FIELD_FC_inactive_box() As Long
   FIELD_FC_inactive_box = mvarFIELD_FC_inactive_box
End Property

'  12
Public Property Let FIELD_FC_disabled_box(ByVal vData As Long)
   mvarFIELD_FC_disabled_box = vData
End Property

Public Property Get FIELD_FC_disabled_box() As Long
   FIELD_FC_disabled_box = mvarFIELD_FC_disabled_box
End Property

'  13
Public Property Let FIELD_FC_werror_active(ByVal vData As Long)
   mvarFIELD_FC_werror_active = vData
End Property

Public Property Get FIELD_FC_werror_active() As Long
   FIELD_FC_werror_active = mvarFIELD_FC_werror_active
End Property

'  14
Public Property Let FIELD_FC_werror_inactive(ByVal vData As Long)
   mvarFIELD_FC_werror_inactive = vData
End Property

Public Property Get FIELD_FC_werror_inactive() As Long
   FIELD_FC_werror_inactive = mvarFIELD_FC_werror_inactive
End Property

'  15
Public Property Let FIELD_size(ByVal vData As Integer)
   mvarFIELD_size = vData
End Property

Public Property Get FIELD_size() As Integer
   FIELD_size = mvarFIELD_size
End Property


