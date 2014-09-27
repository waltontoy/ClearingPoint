VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CConnection"
Attribute VB_Ext_KEY = "Member0" ,"CConnection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

'''''Public Enum ProviderType
'''''
'''''    cpiAccess = 1
'''''    cpiSQLServer = 2
'''''    cpiDBCFS = 3
'''''    cpiInformix = 4
'''''
'''''End Enum

Public Enum ProviderType
    [Access Provider] = 1
    [SQL Server Provider] = 2
    [DBCFS Provider] = 3
    [Informix Provider] = 4
End Enum

Public Function Add(Key As String, ByRef Application As Object, ByVal ConnectionStream As String, ByVal DBPassword As String) As CConnection
    
    ' create a new object
    Dim objNewMember As CConnection
    Dim objConnection As ADODB.Connection
    Dim clsRegistry As CRegistry
    
    Dim enuProvider As ProviderType
    Dim strDBPath As String
    Dim strRegKey() As String
    
    Dim arrParseConnection As Variant
    Dim blnFlag As Boolean
    Dim intHkey As Integer
        
        
    Set objNewMember = New CConnection
    Set clsRegistry = New CRegistry

    If (ConnectionStream <> "") Then
    
        arrParseConnection = Split(ConnectionStream, "|")
        
        blnFlag = (InStr(1, arrParseConnection(1), "~", vbTextCompare) = 0)
        
        ' check if standard
        If (blnFlag = True) Then
        
            ' get database path
            clsRegistry.GetRegistry cpiLocalMachine, Application.ProductName, _
                      CStr(arrParseConnection(1)), CStr(arrParseConnection(2)), cpiStandard
        
        ElseIf (blnFlag = False) Then
        
            ' not standard
            strRegKey = Split(arrParseConnection(1), "~", , vbTextCompare)
            
            Select Case UCase$(strRegKey(0))
                
                Case "HKEY_CURRENT_USER"
                    
                    intHkey = 1
                
                Case "HKEY_LOCAL_MACHINE"
                    
                    intHkey = 2
                
                Case "HKEY_CURRENT_CONFIG"
                    
                    intHkey = 3
                
                Case "HKEY_CLASSES_ROOT"
                    
                    intHkey = 4
                            
            End Select
            
            ' check if registry key exist else create one
            ' get database path
            clsRegistry.GetRegistry intHkey, Application.ProductName, _
                      CStr(strRegKey(1)), CStr(arrParseConnection(2)), cpiUserDefined

        End If
        
        Select Case UCase$(CStr(arrParseConnection(4)))
            
            Case "ACCESS"
                
                enuProvider = cpiAccess
            
            Case "SQL"
                
                enuProvider = cpiSQLServer
        
        End Select
        
        
        blnFlag = (UBound(arrParseConnection) = 6)
        ' set the properties passed into the method
        
        If (blnFlag = True) Then
        
            objNewMember.Connect Key, objConnection, enuProvider, Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) & "\" & _
                                    CStr(arrParseConnection(5)), Decrypt(DBPassword, KEY_ENCRYPT), CStr(arrParseConnection(6))
                                    
        ElseIf (blnFlag = False) Then
        
            objNewMember.Connect Key, objConnection, enuProvider, Decrypt(clsRegistry.RegistryValue, KEY_ENCRYPT) _
                                        & "\" & CStr(arrParseConnection(5)), Decrypt(DBPassword, KEY_ENCRYPT)
            
        End If
    
    ElseIf (ConnectionStream = "") Then
    
        Err.Raise 1, , "Invalid Connection"
    
    End If
    
    mCol.Add objNewMember, Key

    ' return the object created
    Set Add = objNewMember
    
    ' destroy objects
    Set objNewMember = Nothing
    Set clsRegistry = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As CConnection
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

