VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataToSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private m_varSignerName As String
Private m_varStringToSign As String
Private m_varErrorMessage As String
Private m_varDigitallySignedString As String

Public Property Get SignedString() As String
    SignedString = m_varDigitallySignedString
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_varErrorMessage
End Property

Public Property Let StringToSign(ByVal vData As String)
    m_varStringToSign = vData
End Property

Public Property Get StringToSign() As String
    StringToSign = m_varStringToSign
End Property

'Public Property Let SignerName(ByVal vData As String)
'    m_varSignerName = vData
'End Property

'Public Property Get SignerName() As String
'    SignerName = m_varSignerName
'End Property

Public Function ValidSigner(ByVal SignerName As String) As Boolean
    
    If Trim$(SignerName) = "" Then
        m_varErrorMessage = "Signer Name is not Specified!"
        GoTo ErrHandler
    End If
    
    Set G_MyStore = New Store
    G_MyStore.Open CAPICOM_CURRENT_USER_STORE, CAPICOM_MY_STORE, CAPICOM_STORE_OPEN_READ_ONLY
    
    If G_MyStore.Certificates.Count > 0 Then
            
        Set G_MyCertificates = G_MyStore.Certificates
        
        'Modified by Philip on 02-26-2007.
        'Allow certificates that are not stored on Trusted Root Store since we are only creating digital signature
        'We only need certificates on Trusted Root Store if we are verifying digital signature
        'Set G_MyCertificates = G_MyCertificates.Find(CAPICOM_CERTIFICATE_FIND_ISSUER_NAME, SignerName, True)
        Set G_MyCertificates = G_MyCertificates.Find(CAPICOM_CERTIFICATE_FIND_ISSUER_NAME, SignerName)
    
        Select Case G_MyCertificates.Count
        Case 1
                 
            Set G_MyCertificate = G_MyCertificates.Item(1)
            
            If Not G_MyCertificate.HasPrivateKey Then
                m_varErrorMessage = "The Signer Certificate has no associated private key to use!"
                GoTo ErrHandler
            End If
                   
            If Trim(m_varStringToSign) = "" Then
                m_varErrorMessage = "No data to sign!"
                GoTo ErrHandler
            End If
            
            m_varDigitallySignedString = SignMessage(G_MyCertificate, m_varStringToSign, m_varErrorMessage)
            
            If Trim(m_varErrorMessage) <> "" Then
                GoTo ErrHandler
            End If
            
            m_varErrorMessage = ""
            ValidSigner = True
                                                        
        Case Else
            m_varErrorMessage = "Could not find the certificate to use!"
            GoTo ErrHandler
        End Select
        
    Else
        m_varErrorMessage = "The selected certificate is not valid!"
        GoTo ErrHandler
    End If
    
    G_MyStore.Close
    Set G_MyStore = Nothing
    
    Exit Function
    
ErrHandler:
    ValidSigner = False
    Exit Function
    
End Function

