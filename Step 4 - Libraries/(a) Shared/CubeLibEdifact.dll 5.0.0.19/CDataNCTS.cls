VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataNCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
'local variable(s) to hold property value(s)
Private mvarData_NCTS_ID As Long 'local copy
Private mvarDATA_NCTS_MSG_ID As Long 'local copy
Private mvarCode As String 'local copy
Private mvarLogicalIDDescription As String 'local copy
Private mvarSendMode As String 'local copy
Private mvarCommunicationMode As String 'local copy
Private mvarUserID As Long 'local copy
Private mvarTreeID As String 'local copy
Private mvarSupplierName As String 'local copy
Private mvarDocumentType As Long 'local copy
Private mvarLogicalID As String 'local copy
Private mvarMRN As String 'local copy
Public Property Let MRN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MRN = 5
    mvarMRN = vData
End Property


Public Property Get MRN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MRN
    MRN = mvarMRN
End Property



Public Property Let LogicalID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LogicalID = 5
    mvarLogicalID = vData
End Property


Public Property Get LogicalID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogicalID
    LogicalID = mvarLogicalID
End Property



Public Property Let DocumentType(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DocumentType = 5
    mvarDocumentType = vData
End Property


Public Property Get DocumentType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DocumentType
    DocumentType = mvarDocumentType
End Property



Public Property Let SupplierName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplierName = 5
    mvarSupplierName = vData
End Property


Public Property Get SupplierName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplierName
    SupplierName = mvarSupplierName
End Property



Public Property Let TreeID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreeID = 5
    mvarTreeID = vData
End Property


Public Property Get TreeID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeID
    TreeID = mvarTreeID
End Property



Public Property Let UserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserID = 5
    mvarUserID = vData
End Property


Public Property Get UserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserID
    UserID = mvarUserID
End Property



Public Property Let CommunicationMode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommunicationMode = 5
    mvarCommunicationMode = vData
End Property


Public Property Get CommunicationMode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommunicationMode
    CommunicationMode = mvarCommunicationMode
End Property



Public Property Let SendMode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SendMode = 5
    mvarSendMode = vData
End Property


Public Property Get SendMode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SendMode
    SendMode = mvarSendMode
End Property



Public Property Let LogicalIDDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LogicalIDDescription = 5
    mvarLogicalIDDescription = vData
End Property


Public Property Get LogicalIDDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogicalIDDescription
    LogicalIDDescription = mvarLogicalIDDescription
End Property



Public Property Let Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
    mvarCode = vData
End Property


Public Property Get Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property



Public Property Let DATA_NCTS_MSG_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DATA_NCTS_MSG_ID = 5
    mvarDATA_NCTS_MSG_ID = vData
End Property


Public Property Get DATA_NCTS_MSG_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DATA_NCTS_MSG_ID
    DATA_NCTS_MSG_ID = mvarDATA_NCTS_MSG_ID
End Property



Public Property Let Data_NCTS_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Data_NCTS_ID = 5
    mvarData_NCTS_ID = vData
End Property


Public Property Get Data_NCTS_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Data_NCTS_ID
    Data_NCTS_ID = mvarData_NCTS_ID
End Property
