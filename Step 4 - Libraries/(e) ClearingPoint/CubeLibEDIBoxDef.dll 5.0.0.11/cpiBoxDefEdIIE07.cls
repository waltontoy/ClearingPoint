VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiBoxDefEdIIE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarBOX_CODE As String  ' 1
Private mvarENGLISH_DESCRIPTION As String   ' 2
Private mvarDUTCH_DESCRIPTION As String ' 3
Private mvarFRENCH_DESCRIPTION  As String   ' 4
Private mvarEMPTY_FIELD_VALUE  As String   ' 5
Private mvarINSERT_FIELD  As Integer   ' 6
Private mvarJUSTIFY_FIELD As String   ' 7
Private mvarSEQUENTIAL_TABBING_DEFAULT  As Boolean   ' 8
Private mvarSEQUENTIAL_TABBING_ACTIVE  As Boolean   ' 9
Private mvarCHECK_VAT  As Boolean   ' 10
Private mvarCALCULATE_NET_WEIGHT  As Boolean   ' 11
Private mvarCALCULATE_NO_OF_ITEMS  As Boolean   ' 12
Private mvarCOPY_TO_NEXT  As Boolean   ' 13
Private mvarCHANGE_ONLY_IN_H1  As Boolean   ' 14
Private mvarSEND_ONLY_IN_H1  As Boolean   ' 15
Private mvarWIDTH_FIELD  As Integer   ' 16
Private mvarDECIMAL_FIELD  As Integer   ' 17
Private mvarDATA_TYPE As String   ' 18
Private mvarAUTO_ADD  As Boolean   ' 19
Private mvarDEFAULT_VALUE  As String   ' 20
Private mvarCALCULATE_CUSTOMS_VALUE  As Boolean   ' 21
Private mvarRELATE_L1_TO_S1  As Boolean   ' 22
Private mvarBOX_TYPE As String    ' 23
Private mvarBOX_INDEX As Integer     ' 24
Private mvarINTERNAL_CODE As String ' 25
'INTERNAL_CODE

' 1
Public Property Let BOX_CODE(ByVal vData As String)
    
    mvarBOX_CODE = vData

End Property

' 1
Public Property Get BOX_CODE() As String
    
    BOX_CODE = mvarBOX_CODE

End Property

' 2
Public Property Let ENGLISH_DESCRIPTION(ByVal vData As String)
    
    mvarENGLISH_DESCRIPTION = vData

End Property

' 2
Public Property Get ENGLISH_DESCRIPTION() As String
    
    ENGLISH_DESCRIPTION = mvarENGLISH_DESCRIPTION

End Property

' 3
Public Property Let DUTCH_DESCRIPTION(ByVal vData As String)
    
    mvarDUTCH_DESCRIPTION = vData

End Property

' 3
Public Property Get DUTCH_DESCRIPTION() As String
    
    DUTCH_DESCRIPTION = mvarDUTCH_DESCRIPTION

End Property

' 4
Public Property Let FRENCH_DESCRIPTION(ByVal vData As String)
    
    mvarFRENCH_DESCRIPTION = vData

End Property

' 4
Public Property Get FRENCH_DESCRIPTION() As String
    
    FRENCH_DESCRIPTION = mvarFRENCH_DESCRIPTION

End Property

' 5
Public Property Let EMPTY_FIELD_VALUE(ByVal vData As String)
    
    mvarEMPTY_FIELD_VALUE = vData

End Property

' 5
Public Property Get EMPTY_FIELD_VALUE() As String
    
    EMPTY_FIELD_VALUE = mvarEMPTY_FIELD_VALUE

End Property

' 6
Public Property Let INSERT_FIELD(ByVal vData As Integer)
    
    mvarINSERT_FIELD = vData

End Property

' 6
Public Property Get INSERT_FIELD() As Integer
    
    INSERT_FIELD = mvarINSERT_FIELD

End Property

' 7
Public Property Let JUSTIFY_FIELD(ByVal vData As String)
    
    mvarJUSTIFY_FIELD = vData

End Property

' 7
Public Property Get JUSTIFY_FIELD() As String
    
    JUSTIFY_FIELD = mvarJUSTIFY_FIELD

End Property

' 8
Public Property Let SEQUENTIAL_TABBING_DEFAULT(ByVal vData As Boolean)
    
    mvarSEQUENTIAL_TABBING_DEFAULT = vData

End Property

' 8
Public Property Get SEQUENTIAL_TABBING_DEFAULT() As Boolean
    
    SEQUENTIAL_TABBING_DEFAULT = mvarSEQUENTIAL_TABBING_DEFAULT

End Property

' 9
Public Property Let SEQUENTIAL_TABBING_ACTIVE(ByVal vData As Boolean)
    
    mvarSEQUENTIAL_TABBING_ACTIVE = vData

End Property

' 9
Public Property Get SEQUENTIAL_TABBING_ACTIVE() As Boolean
    
    SEQUENTIAL_TABBING_ACTIVE = mvarSEQUENTIAL_TABBING_ACTIVE

End Property

' 10
Public Property Let CHECK_VAT(ByVal vData As Boolean)
    
    mvarCHECK_VAT = vData

End Property

' 10
Public Property Get CHECK_VAT() As Boolean
    
    CHECK_VAT = mvarCHECK_VAT

End Property

' 11
Public Property Let CALCULATE_NET_WEIGHT(ByVal vData As Boolean)
    
    mvarCALCULATE_NET_WEIGHT = vData

End Property

' 11
Public Property Get CALCULATE_NET_WEIGHT() As Boolean
    
    CALCULATE_NET_WEIGHT = mvarCALCULATE_NET_WEIGHT

End Property

' 12
Public Property Let CALCULATE_NO_OF_ITEMS(ByVal vData As Boolean)
    
    mvarCALCULATE_NO_OF_ITEMS = vData

End Property

' 12
Public Property Get CALCULATE_NO_OF_ITEMS() As Boolean
    
    CALCULATE_NO_OF_ITEMS = mvarCALCULATE_NO_OF_ITEMS

End Property

' 13
Public Property Let COPY_TO_NEXT(ByVal vData As Boolean)
    
    mvarCOPY_TO_NEXT = vData

End Property

' 13
Public Property Get COPY_TO_NEXT() As Boolean
    
    COPY_TO_NEXT = mvarCOPY_TO_NEXT

End Property

' 14
Public Property Let CHANGE_ONLY_IN_H1(ByVal vData As Boolean)
    
    mvarCHANGE_ONLY_IN_H1 = vData

End Property

' 14
Public Property Get CHANGE_ONLY_IN_H1() As Boolean
    
    CHANGE_ONLY_IN_H1 = mvarCHANGE_ONLY_IN_H1

End Property

' 15
Public Property Let SEND_ONLY_IN_H1(ByVal vData As Boolean)
    
    mvarSEND_ONLY_IN_H1 = vData

End Property

' 15
Public Property Get SEND_ONLY_IN_H1() As Boolean
    
    SEND_ONLY_IN_H1 = mvarSEND_ONLY_IN_H1

End Property

' 16
Public Property Let WIDTH_FIELD(ByVal vData As Integer)
    
    mvarWIDTH_FIELD = vData

End Property

' 16
Public Property Get WIDTH_FIELD() As Integer
    
    WIDTH_FIELD = mvarWIDTH_FIELD

End Property

' 17
Public Property Let DECIMAL_FIELD(ByVal vData As Integer)
    
    mvarDECIMAL_FIELD = vData

End Property

' 17
Public Property Get DECIMAL_FIELD() As Integer
    
    DECIMAL_FIELD = mvarDECIMAL_FIELD

End Property

' 18
Public Property Let DATA_TYPE(ByVal vData As String)
    
    mvarDATA_TYPE = vData

End Property

' 18
Public Property Get DATA_TYPE() As String
    
    DATA_TYPE = mvarDATA_TYPE

End Property

' 19
Public Property Let AUTO_ADD(ByVal vData As Boolean)
    
    mvarAUTO_ADD = vData

End Property

' 19
Public Property Get AUTO_ADD() As Boolean
    
    AUTO_ADD = mvarAUTO_ADD

End Property

' 20
Public Property Let DEFAULT_VALUE(ByVal vData As String)
    
    mvarDEFAULT_VALUE = vData

End Property

' 20
Public Property Get DEFAULT_VALUE() As String
    
    DEFAULT_VALUE = mvarDEFAULT_VALUE

End Property

' 21
Public Property Let CALCULATE_CUSTOMS_VALUE(ByVal vData As Boolean)
    
    mvarCALCULATE_CUSTOMS_VALUE = vData

End Property

' 21
Public Property Get CALCULATE_CUSTOMS_VALUE() As Boolean
    
    CALCULATE_CUSTOMS_VALUE = mvarCALCULATE_CUSTOMS_VALUE

End Property

' 22
Public Property Let RELATE_L1_TO_S1(ByVal vData As Boolean)
    
    mvarRELATE_L1_TO_S1 = vData

End Property

' 22
Public Property Get RELATE_L1_TO_S1() As Boolean
    
    RELATE_L1_TO_S1 = mvarRELATE_L1_TO_S1

End Property

' 23
Public Property Let BOX_TYPE(ByVal vData As String)
    
    mvarBOX_TYPE = vData

End Property

' 23
Public Property Get BOX_TYPE() As String
    
    BOX_TYPE = mvarBOX_TYPE

End Property

' 24
Public Property Let BOX_INDEX(ByVal vData As Integer)
    
    mvarBOX_INDEX = vData

End Property

' 24
Public Property Get BOX_INDEX() As Integer
    
    BOX_INDEX = mvarBOX_INDEX

End Property

' 25
Public Property Let INTERNAL_CODE(ByVal vData As String)
    
    mvarINTERNAL_CODE = vData

End Property

' 25
Public Property Get INTERNAL_CODE() As String
    
    INTERNAL_CODE = mvarINTERNAL_CODE

End Property


'INTERNAL_CODE
