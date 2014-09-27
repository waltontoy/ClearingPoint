VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSignedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_varContentString As String
Private m_varSignedString As String
Private m_varErrorMessage As String

Public Function IsValid() As Boolean
Attribute IsValid.VB_Description = "Verifies the SignedString using the installed certificate(s) found on Trusted Root Certification Authorities Store"

    Dim MySignedData As SignedData
                
    If Trim(ContentString) = "" Then
        m_varErrorMessage = "No signed data to verify!"
        GoTo ErrHandler
    End If
    
    Set MySignedData = New SignedData
    
    ' Convert From Hex To Binary
    Set G_EncodingType = New Utilities
    SignedString = G_EncodingType.HexToBinary(SignedString)

    ' Set the detached content upon which the signature is based.
    MySignedData.Content = ContentString
    
    ' Verify the detached signature.
    ' Signer Certificate s/b in the Trusted Root Certification Authorities Store
    ' The 'True' parameter indicates that the content signed is not included in the signature string.
    On Error Resume Next
        MySignedData.Verify SignedString, True
    
    If Err.Number <> 0 Then
        m_varErrorMessage = Err.Description
        GoTo ErrHandler
    Else
        m_varErrorMessage = ""
        IsValid = True
    End If
    
    ' Release the SignedData object.
    Set MySignedData = Nothing
    
    Exit Function

ErrHandler:
    IsValid = False
    Exit Function
    
End Function

Public Property Get ErrorMessage() As String
    ErrorMessage = m_varErrorMessage
End Property

Public Property Let SignedString(ByVal vData As String)
Attribute SignedString.VB_Description = "Holds the digitally signed data."
    m_varSignedString = vData
End Property

Public Property Get SignedString() As String
    SignedString = m_varSignedString
End Property

Public Property Let ContentString(ByVal vData As String)
Attribute ContentString.VB_Description = "Holds the original data that was digitally signed."
    m_varContentString = vData
End Property

Public Property Get ContentString() As String
    ContentString = m_varContentString
End Property



