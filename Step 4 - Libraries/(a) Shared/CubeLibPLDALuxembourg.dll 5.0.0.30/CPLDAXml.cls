VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPLDAXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' 05/25/2007 By Philip
'
' Note: This library is used to read PLDA Response XML file.
'       Please see CP Task Scheduler for the implementation.
'

Option Explicit

'local variable(s) to hold property value(s)
Private mvarSender As String 'local copy
Private mvarRecipient As String 'local copy
Private mvarVersion As String 'local copy
Private mvarTestIndicator As String 'local copy
Private mvarFunctionCode As String 'local copy
Private mvarDateofPreparation As String 'local copy
Private mvarTimeOfPreparation As String 'local copy
Private mvarLRN As String 'local copy
Private mvarMRN As String 'local copy
Private mvarCancellationDate As Variant 'local copy
Private mvarCancellationTime As String 'local copy
Private mvarCancellationReason As Variant 'local copy
Private mvarExportOffice As String 'local copy
Private mvarDestinationOffice As String 'local copy
Private mvarDeclarantOperatorIdentity As String 'local copy
Private mvarDeclarantOperatorName As String 'local copy
Private mvarDeclarantOperatorPostal As String 'local copy
Private mvarDeclarantOperatorStreetNo1 As String 'local copy
Private mvarDeclarantOperatorStreetNo2 As String 'local copy
Private mvarDeclarantOperatorCity As String 'local copy
Private mvarDeclarantOperatorCountry As String 'local copy
Private mvarElementPath As String 'local copy
Private mvarElementSequence As String 'local copy
Private mvarErrorCode As String 'local copy
Private mvarErrorContent As String 'local copy
Private mvarErrorDescription As String 'local copy
Private mvarSequence As String 'local copy
Private mvarTaxType As String 'local copy
Private mvarTaxBase As String 'local copy
Private mvarTaxAmount As String 'local copy
Private mvarTaxPaymentMethod As String 'local copy
Private mvarXMLMessage As String

'Edwin - July 4, 2007 - Added for finding node values
Private mvarPreviousNode As String 'local copy
Private mvarParentNode As IXMLDOMNode 'local copy
Private mvarNodesFound As Long 'local copy
Private mvarNodeValues() As String 'local copy

Public Property Get TaxPaymentMethod() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaxPaymentMethod
    TaxPaymentMethod = mvarTaxPaymentMethod
End Property

Public Property Get TaxAmount() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaxAmount
    TaxAmount = mvarTaxAmount
End Property

Public Property Get TaxBase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaxBase
    TaxBase = mvarTaxBase
End Property

Public Property Get TaxType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaxType
    TaxType = mvarTaxType
End Property

Public Property Get Sequence() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sequence
    Sequence = mvarSequence
End Property

Public Property Get ErrorDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorDescription
    ErrorDescription = mvarErrorDescription
End Property

Public Property Get ErrorContent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorContent
    ErrorContent = mvarErrorContent
End Property

Public Property Get ErrorCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorCode
    ErrorCode = mvarErrorCode
End Property

Public Property Get ElementSequence() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ElementSequence
    ElementSequence = mvarElementSequence
End Property

Public Property Get ElementPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ElementPath
    ElementPath = mvarElementPath
End Property

Public Property Get DeclarantOperatorCountry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorCountry
    DeclarantOperatorCountry = mvarDeclarantOperatorCountry
End Property

Public Property Get DeclarantOperatorCity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorCity
    DeclarantOperatorCity = mvarDeclarantOperatorCity
End Property

Public Property Get DeclarantOperatorStreetNo2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorStreetNo2
    DeclarantOperatorStreetNo2 = mvarDeclarantOperatorStreetNo2
End Property

Public Property Get DeclarantOperatorStreetNo1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorStreetNo1
    DeclarantOperatorStreetNo1 = mvarDeclarantOperatorStreetNo1
End Property

Public Property Get DeclarantOperatorPostal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorPostal
    DeclarantOperatorPostal = mvarDeclarantOperatorPostal
End Property

Public Property Get DeclarantOperatorName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorName
    DeclarantOperatorName = mvarDeclarantOperatorName
End Property

Public Property Get DeclarantOperatorIdentity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeclarantOperatorIdentity
    DeclarantOperatorIdentity = mvarDeclarantOperatorIdentity
End Property

Public Property Get DestinationOffice() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationOffice
    DestinationOffice = mvarDestinationOffice
End Property

Public Property Get ExportOffice() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExportOffice
    ExportOffice = mvarExportOffice
End Property

Public Property Get CancellationReason() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CancellationReason
    If IsObject(mvarCancellationReason) Then
        Set CancellationReason = mvarCancellationReason
    Else
        CancellationReason = mvarCancellationReason
    End If
End Property

Public Property Get CancellationTime() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CancellationTime
    CancellationTime = mvarCancellationTime
End Property

Public Property Get CancellationDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CancellationDate
    If IsObject(mvarCancellationDate) Then
        Set CancellationDate = mvarCancellationDate
    Else
        CancellationDate = mvarCancellationDate
    End If
End Property

Public Property Get MRN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MRN
    MRN = mvarMRN
End Property

Public Property Get LRN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LRN
    LRN = mvarLRN
End Property

Public Property Get TimeOfPreparation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeOfPreparation
    TimeOfPreparation = mvarTimeOfPreparation
End Property

Public Property Get DateOfPreparation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateofPreparation
    DateOfPreparation = mvarDateofPreparation
End Property

Public Property Get FunctionCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FunctionCode
    FunctionCode = mvarFunctionCode
End Property

Public Property Get TestIndicator() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TestIndicator
    TestIndicator = mvarTestIndicator
End Property

Public Property Get Version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property

Public Property Get Recipient() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Recipient
    Recipient = mvarRecipient
End Property

Public Property Get Sender() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sender
    Sender = mvarSender
End Property

Public Function ReadPLDAResponseXML(ByVal objParentNode As IXMLDOMNode)
    
    If LCase(objParentNode.nodeName) = "responsepldaheader" Then
        GetResponsePLDAHeader objParentNode
    End If

    If LCase(objParentNode.nodeName) = "responsepldaitem" Then
        GetResponsePLDAItem objParentNode
    End If
    
    If LCase(objParentNode.nodeName) = "calculationresults" Then
        GetCalculationResults objParentNode
    End If

End Function

Public Function ReadPLDAXML(ByVal strXML As String) As Boolean

    On Error GoTo ErrHandler
    
    Dim objDOM As DOMDocument
    Dim objParentNode As IXMLDOMNode
        
    Set objDOM = New MSXML2.DOMDocument
    objDOM.Load strXML
        
    mvarXMLMessage = objDOM.xml
    mvarParentNode = objDOM.documentElement.ParentNode
    
    For Each objParentNode In objDOM.documentElement.childNodes
        
        If LCase(objParentNode.nodeName) = "interchangeheader" Then
            GetInterchangeHeader objParentNode
        End If
        
        Select Case UCase(mvarFunctionCode)
            Case "IE629"    'PLDA Import - Release
                If LCase(objParentNode.nodeName) = "pldadv1" Then
                    GetResponsePLDAImport objParentNode
                End If
                
            Case "IE529"    'PLDA Export - Release
                If LCase(objParentNode.nodeName) = "goodsdeclaration" Then
                    GetResponsePLDAExport objParentNode
                End If
                
            Case "IE674"    'PLDA Import - Invoice
                'Wala pa dito. Hindi pa tapo sni reims
                                
            Case Else       'PLDA Response
                ReadPLDAResponseXML objParentNode
        End Select
        Set mvarParentNode = objParentNode.ParentNode
    Next
        
    ReadPLDAXML = True
    
    Set objParentNode = Nothing
    Set objDOM = Nothing
    
    Exit Function
    
ErrHandler:
    
    ReadPLDAXML = False
    
End Function
    
Private Function GetInterchangeHeader(ByVal objParentNode As IXMLDOMNode)

    Dim objNode As IXMLDOMNode
    Dim objChildNode As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "messagesender" Then
            mvarSender = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "messagerecipient" Then
            mvarRecipient = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "messageversion" Then
            mvarVersion = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "testindicator" Then
            mvarTestIndicator = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "functioncode" Then
            mvarFunctionCode = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "datetimeofpreparation" Then
            For Each objNode In objChildNode.childNodes
                If LCase(objNode.nodeName) = "dateofpreparation" Then
                    mvarDateofPreparation = objNode.Text
                End If
                
                If LCase(objNode.nodeName) = "timeofpreparation" Then
                    mvarTimeOfPreparation = objNode.Text
                End If
            Next
            
        End If
                            
    Next
    
    If Not objNode Is Nothing Then Set objNode = Nothing
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    
End Function

Private Function GetResponsePLDAHeader(ByVal objParentNode As IXMLDOMNode)

    Dim objChildNode As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    Dim objNode1 As IXMLDOMNode
    Dim objNode2 As IXMLDOMNode
    Dim objNode3 As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "localreferencenumber" Then
            mvarLRN = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "customsmainreferencenumber" Then
            mvarMRN = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "cancellationrequestdate" Then
            mvarCancellationDate = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "cancellationrequesttime" Then
            mvarCancellationTime = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "cancellationreason" Then
            mvarCancellationReason = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "exportoffice" Then
            mvarExportOffice = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "destinationoffice" Then
            mvarDestinationOffice = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "declarant" Then
            For Each objNode In objChildNode.childNodes
                If LCase(objNode.nodeName) = "declarantoperator" Then
                    For Each objNode1 In objNode.childNodes
                        If LCase(objNode1.nodeName) = "operatoridentity" Then
                            For Each objNode2 In objNode1.childNodes
                                If LCase(objNode2.nodeName) = "operatoridentity" Then
                                    mvarDeclarantOperatorIdentity = objNode2.Text
                                End If
                            Next
                        End If
                        
                        If LCase(objNode1.nodeName) = "operator" Then
                            For Each objNode2 In objNode1.childNodes
                                If LCase(objNode2.nodeName) = "operatorname" Then
                                    mvarDeclarantOperatorName = objNode2.Text
                                End If
                                If LCase(objNode2.nodeName) = "operatoraddress" Then
                                    For Each objNode3 In objNode2.childNodes
                                        If LCase(objNode3.nodeName) = "postalcode" Then
                                            mvarDeclarantOperatorPostal = objNode3.Text
                                        End If
                                        If LCase(objNode3.nodeName) = "streetandnumber1" Then
                                            mvarDeclarantOperatorStreetNo1 = objNode3.Text
                                        End If
                                        If LCase(objNode3.nodeName) = "streetandnumber2" Then
                                            mvarDeclarantOperatorStreetNo2 = objNode3.Text
                                        End If
                                        If LCase(objNode3.nodeName) = "city" Then
                                            mvarDeclarantOperatorCity = objNode3.Text
                                        End If
                                        If LCase(objNode3.nodeName) = "country" Then
                                            mvarDeclarantOperatorCountry = objNode3.Text
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
                                    
    Next
    
    If Not objNode3 Is Nothing Then Set objNode3 = Nothing
    If Not objNode2 Is Nothing Then Set objNode2 = Nothing
    If Not objNode1 Is Nothing Then Set objNode1 = Nothing
    If Not objNode Is Nothing Then Set objNode = Nothing
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    
End Function

Private Function GetResponsePLDAExport(ByVal objParentNode As IXMLDOMNode)

    Dim objChildNode As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    Dim objNode1 As IXMLDOMNode
    Dim objNode2 As IXMLDOMNode
    Dim objNode3 As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "localreferencenumber" Then
            mvarLRN = objChildNode.Text
        End If
    
        If LCase(objChildNode.nodeName) = "customsmainreferencenumber" Then
            mvarMRN = objChildNode.Text
        End If
    
'        If LCase(objChildNode.nodeName) = "dv1header" Then '"declarant" Then
'            For Each objNode In objChildNode.childNodes
'                If LCase(objNode.nodeName) = "customsreference" Then '"declarantoperator" Then
'                    For Each objNode1 In objNode.childNodes
'                        If LCase(objNode1.nodeName) = "operatoridentity" Then
'                            For Each objNode2 In objNode1.childNodes
'                                If LCase(objNode2.nodeName) = "operatoridentity" Then
'                                    mvarDeclarantOperatorIdentity = objNode2.Text
'                                End If
'                            Next
'                        End If
'
'                        '<<< dandan 062707
'                        'retrieving lrn and mrn
'                        If LCase(objNode1.nodeName) = "localreferencenumber" Then
'                            mvarLRN = objNode1.Text
'                        End If
'
'                        If LCase(objNode1.nodeName) = "customsmainreferencenumber" Then
'                            mvarMRN = objNode1.Text
'                        End If
'
'
'                        If LCase(objNode1.nodeName) = "operator" Then
'                            For Each objNode2 In objNode1.childNodes
'                                If LCase(objNode2.nodeName) = "operatorname" Then
'                                    mvarDeclarantOperatorName = objNode2.Text
'                                End If
'                                If LCase(objNode2.nodeName) = "operatoraddress" Then
'                                    For Each objNode3 In objNode2.childNodes
'                                        If LCase(objNode3.nodeName) = "postalcode" Then
'                                            mvarDeclarantOperatorPostal = objNode3.Text
'                                        End If
'                                        If LCase(objNode3.nodeName) = "streetandnumber1" Then
'                                            mvarDeclarantOperatorStreetNo1 = objNode3.Text
'                                        End If
'                                        If LCase(objNode3.nodeName) = "streetandnumber2" Then
'                                            mvarDeclarantOperatorStreetNo2 = objNode3.Text
'                                        End If
'                                        If LCase(objNode3.nodeName) = "city" Then
'                                            mvarDeclarantOperatorCity = objNode3.Text
'                                        End If
'                                        If LCase(objNode3.nodeName) = "country" Then
'                                            mvarDeclarantOperatorCountry = objNode3.Text
'                                        End If
'                                    Next
'                                End If
'                            Next
'                        End If
'                    Next
'                End If
'            Next
'        End If
                                    
    Next
    
    If Not objNode3 Is Nothing Then Set objNode3 = Nothing
    If Not objNode2 Is Nothing Then Set objNode2 = Nothing
    If Not objNode1 Is Nothing Then Set objNode1 = Nothing
    If Not objNode Is Nothing Then Set objNode = Nothing
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    
End Function

Private Function GetResponsePLDAImport(ByVal objParentNode As IXMLDOMNode)

    Dim objChildNode As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    Dim objNode1 As IXMLDOMNode
    'Dim objNode2 As IXMLDOMNode
    'Dim objNode3 As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "dv1header" Then
            For Each objNode In objChildNode.childNodes
                If LCase(objNode.nodeName) = "customsreference" Then
                    For Each objNode1 In objNode.childNodes
                        
                        '<<< dandan 062707
                        'retrieving lrn and mrn
                        If LCase(objNode1.nodeName) = "localreferencenumber" Then
                            mvarLRN = objNode1.Text
                        End If
                    
                        If LCase(objNode1.nodeName) = "customsmainreferencenumber" Then
                            mvarMRN = objNode1.Text
                        End If
                        
                    Next
                End If
            Next
        End If
                                    
    Next
    
    'If Not objNode3 Is Nothing Then Set objNode3 = Nothing
    'If Not objNode2 Is Nothing Then Set objNode2 = Nothing
    If Not objNode1 Is Nothing Then Set objNode1 = Nothing
    If Not objNode Is Nothing Then Set objNode = Nothing
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    
End Function

Private Function GetResponsePLDAItem(ByVal objParentNode As IXMLDOMNode)
    
    Dim objChildNode As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "dataelementpath" Then
            mvarElementPath = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "dataelementsequence" Then
            mvarElementSequence = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "errorcode" Then
            mvarErrorCode = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "errorcontent" Then
            mvarErrorContent = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "errordescription" Then
            mvarErrorDescription = objChildNode.Text
        End If
        
    Next
        
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    
End Function

Private Function GetCalculationResults(ByVal objParentNode As IXMLDOMNode)
    
    Dim objChildNode As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    
    For Each objChildNode In objParentNode.childNodes
        
        If LCase(objChildNode.nodeName) = "sequence" Then
            mvarSequence = objChildNode.Text
        End If
        
        If LCase(objChildNode.nodeName) = "duties" Then
                        
            For Each objNode In objChildNode.childNodes
                
                If LCase(objNode.nodeName) = "taxtype" Then
                    mvarTaxType = objNode.Text
                End If
                
                If LCase(objNode.nodeName) = "taxbase" Then
                    mvarTaxBase = objNode.Text
                End If
                
                If LCase(objNode.nodeName) = "taxamount" Then
                    mvarTaxAmount = objNode.Text
                End If

                If LCase(objNode.nodeName) = "paymentmethodtaxes" Then
                    mvarTaxPaymentMethod = objNode.Text
                End If

            Next
            
        End If
                
    Next
        
    If Not objChildNode Is Nothing Then Set objChildNode = Nothing
    If Not objNode Is Nothing Then Set objNode = Nothing
    
End Function

Public Property Get XMLMessage() As String

    XMLMessage = mvarXMLMessage

End Property


Public Function FindNodeFromXML(ByVal Node As IXMLDOMNode, _
                                ByVal NodeNameToFind As String) As Boolean

    Dim objChildNode As IXMLDOMNode
    
    On Error GoTo ErrHandler

    For Each objChildNode In Node.childNodes
        
        If LCase(mvarPreviousNode) = Trim(LCase(NodeNameToFind)) Then
            ReDim Preserve mvarNodeValues(mvarNodesFound)
            
            If IsNull(objChildNode.nodeValue) = False And objChildNode.nodeType = NODE_TEXT Then
                mvarNodeValues(mvarNodesFound) = objChildNode.nodeValue
            ElseIf IsNull(objChildNode.nodeValue) = True And objChildNode.nodeType = NODE_TEXT Then
                mvarNodeValues(mvarNodesFound) = ""
            Else
                mvarNodeValues(mvarNodesFound) = ""
                Debug.Assert False
            End If
            Debug.Print mvarPreviousNode & " - " & mvarNodeValues(mvarNodesFound)
            mvarNodesFound = mvarNodesFound + 1
        End If
                
        mvarPreviousNode = objChildNode.nodeName
        
        If objChildNode.hasChildNodes = True Then
            If FindNodeFromXML(objChildNode, NodeNameToFind) = False Then
                FindNodeFromXML = False
                Exit Function
            End If
        End If
        
    Next

    FindNodeFromXML = True

    Set objChildNode = Nothing
    
    Exit Function

ErrHandler:

    FindNodeFromXML = False
    
End Function

Public Property Get NodeValues() As Variant

    NodeValues = mvarNodeValues

End Property


Public Property Get ParentNode() As IXMLDOMNode

    Set ParentNode = mvarParentNode

End Property


Public Function InitiliatizeFindNodeFromXML()

    mvarPreviousNode = ""
    mvarNodesFound = 0
    ReDim mvarNodeValues(0)
    
End Function

