Attribute VB_Name = "modDeclarations"
Option Explicit

Public Const PREFIX_DATA_TABLE           As String = "Data"
Public Const PREFIX_EDI_TABLE            As String = "EDI"
Public Const PREFIX_NCTS_ITEMS_TABLE     As String = "NCTS_ITM"
Public Const PREFIX_SEGMENT_KEY_TMS_ID   As String = "TMSID"
Public Const PREFIX_SEGMENT_KEY_INSTANCE As String = "INSTANCE"

Public Const EDI_SEP_SEGMENT                As String = "'"
Public Const EDI_SEP_COMPOSITE_DATA_ELEMENT As String = ":"
Public Const EDI_SEP_DATA_ELEMENT           As String = "+"
Public Const EDI_SEP_RELEASE_CHARACTER      As String = "?"

Public Const EDI_USAGE_REQUIRED  As String = "R"
Public Const EDI_USAGE_OPTIONAL  As String = "O"
Public Const EDI_USAGE_DEPENDENT As String = "D"

Public Const EDI_DATATYPE_QUALIFIER As String = "Q"

Public Const DELIMETER_QUALIFIER_STRING As String = ","

Public Const MESSAGE_STATUS_DOCUMENT As String = "Document"
Public Const MESSAGE_STATUS_QUEUED   As String = "Queued"
Public Const MESSAGE_STATUS_RECEIVED As String = "Received"
Public Const MESSAGE_STATUS_SENT     As String = "Sent"

Public Const OPERATOR_EQUAL         As String = "="
Public Const OPERATOR_GREATER_THAN  As String = ">"
Public Const OPERATOR_LESS_THAN     As String = "<"
Public Const OPERATOR_NOT_EQUAL     As String = "<>"
' MUCP-65 - Start
Public Const OPERATOR_LEFT          As String = "%"
' MUCP-65 - End

Public Const OPERATOR_AND           As String = "&"
Public Const OPERATOR_OR            As String = "|"

Public Const BOX_GROUP_DETAIL            As String = "DETAIL"
Public Const BOX_GROUP_DETAIL_BIJZONDERE As String = "DETAIL_BIJZONDERE"
Public Const BOX_GROUP_DETAIL_COLLI      As String = "DETAIL_COLLI"
Public Const BOX_GROUP_DETAIL_CONTAINER  As String = "DETAIL_CONTAINER"
Public Const BOX_GROUP_DETAIL_DOCUMENTEN As String = "DETAIL_DOCUMENTEN"
Public Const BOX_GROUP_HEADER            As String = "HEADER"
Public Const BOX_GROUP_HEADER_ZEKERHEID  As String = "HEADER_ZEKERHEID"

Global G_strQuery As String

Global Const G_IE15_NCTS_IEM_ID = 5
Global Const G_IE_MESSAGE_TYPE_ARRAY_INDEX = 5

Public Enum ConditionTokenIndexes
    ConditionTokenIndex_LeftOperand = 0
    ConditionTokenIndex_Operator
    ConditionTokenIndex_RightOperand
    ConditionTokenIndex_LogicalOperator
End Enum

Public Const G_Main_Password = "wack2"
