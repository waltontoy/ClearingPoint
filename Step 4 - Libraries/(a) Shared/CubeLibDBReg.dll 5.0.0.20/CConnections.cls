VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarConnection As Variant 'local copy
Private mvarConnectionType As String 'local copy
Private mvarConnectionName As String 'local copy
Private mvarDatabasePath As String 'local copy
Public Property Let DatabasePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatabasePath = 5
    mvarDatabasePath = vData
End Property


Public Property Get DatabasePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabasePath
    DatabasePath = mvarDatabasePath
End Property



Public Property Let ConnectionName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectionName = 5
    mvarConnectionName = vData
End Property


Public Property Get ConnectionName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectionName
    ConnectionName = mvarConnectionName
End Property



Public Property Let ConnectionType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectionType = 5
    mvarConnectionType = vData
End Property


Public Property Get ConnectionType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectionType
    ConnectionType = mvarConnectionType
End Property



Public Property Let Connection(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Connection = 5
    mvarConnection = vData
End Property


Public Property Set Connection(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Connection = Form1
    Set mvarConnection = vData
End Property


Public Property Get Connection() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Connection
    If IsObject(mvarConnection) Then
        Set Connection = mvarConnection
    Else
        Connection = mvarConnection
    End If
End Property


Public Function Connect(ByVal ConnectionName As String, _
                        ByRef ADOConnection As ADODB.Connection, _
                        ByVal Provider As ProviderType, _
                        ByVal DataSource As String, _
                        ByVal Password As String, _
                        Optional ByVal InitialCatalog As String) _
                        As Boolean
                                        
    Connect = False
    
    Key = ConnectionName
    mvarConnectionName = ConnectionName
    
    Select Case Provider
        
        Case 1 'ACCESS
            
            mvarConnectionType = "MS Access"
        
        Case 2 'SQLServer
            
            mvarConnectionType = "MS SQLServer"
        
        Case 3 'DBCFS
            
            mvarConnectionType = "DB/C FS"
        
        Case 4 'Informix
            
            mvarConnectionType = "Informix"
    
    End Select
    
    If ((ADOConnection Is Nothing) = False) Then
    
        If (ADOConnection.State = adStateOpen) Then
            
            ADOConnection.Close
        
        End If
        
        Set ADOConnection = Nothing
    
    End If
    
    Set ADOConnection = New ADODB.Connection
            
    On Error GoTo Error_Handler
    
    ' update database location
    mvarDatabasePath = DataSource
    
    Select Case Provider
    
        Case cpiAccess
        
            ADOConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataSource & ";Persist Security Info=False;Jet OLEDB:Database Password=" & Password
            
        Case cpiSQLServer
        
            ADOConnection.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & _
                                            InitialCatalog & ";Data Source=" & DataSource
            
        Case cpiDBCFS
            
            denDB.conDB.ConnectionString = "Extended Properties=" & Chr(34) & "FileDSN=" & DataSource & ";" & Chr(34)
            denDB.conDB.Open
            
            Set ADOConnection = denDB.conDB
            
        Case cpiInformix
        
            ADOConnection.Open Provider & "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & _
                                            InitialCatalog & ";Data Source=" & DataSource
            
    End Select
    
    On Error GoTo 0
    
    Set mvarConnection = ADOConnection
    Set ADOConnection = Nothing
        
    Connect = True
    
    Exit Function

Error_Handler:
    
    Err.Raise 1, , "Database not found."

End Function

