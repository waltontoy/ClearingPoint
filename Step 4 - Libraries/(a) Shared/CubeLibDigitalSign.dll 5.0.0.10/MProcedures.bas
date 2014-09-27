Attribute VB_Name = "MProcedures"
Option Explicit

Public Sub Main()
    
    'Initialize database path
    G_strMdbPath = NoBackSlash(GetSetting("ClearingPoint", "Settings", "MdbPath", "C:\Program Files\Cubepoint\ClearingPoint"))
    
End Sub

Public Function SignMessage(ByVal Cert As Certificate, ByVal StringToSign As String, ByRef ErrorMessage) As String
    
    Dim Signobj As New SignedData
    Dim Signer As New Signer
    
    ' Use HEX encoding as required by PLDA
    ' No available HEX encoding for signing and needs to be converted using CAPICOM.Utilities
    
    Signer.Certificate = Cert

    On Error Resume Next
        Signobj.Content = StringToSign

    ' Sign the content using the signer's private key.
    ' The 'True' parameter indicates that the content signed is not included in the signature string
    SignMessage = Signobj.Sign(Signer, True, CAPICOM_ENCODE_BINARY)
    'SignMessage = Signobj.Sign(Signer, True, CAPICOM_ENCODE_BASE64)
                
    ' Get the HEX equivalent of the signature
    Set G_EncodingType = New Utilities
    SignMessage = G_EncodingType.BinaryToHex(SignMessage)
    Set G_EncodingType = Nothing
    
    Set Signobj = Nothing
    Set Signer = Nothing
    
    
    If Err.Number <> 0 Then
        ErrorMessage = Err.Description
    Else
        ErrorMessage = ""
    End If
    
    Exit Function
    
End Function
