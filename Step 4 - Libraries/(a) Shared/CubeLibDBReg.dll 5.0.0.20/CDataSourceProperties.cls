VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataSourceProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarUsername As String
Private mvarPassword As String
Private mvarDatabaseType As CubeLibDataSource.DatabaseType
Private mvarDataSource As String
Private mvarInitialCatalog As String
Private mvarOutputFilePath As String
Private mvarInitialCatalogPath As String 'local copy
Private mvarTracefileType As TracefileTypeConsant

Friend Property Let TracefileType(ByVal vData As TracefileTypeConsant)
    mvarTracefileType = vData
End Property

Public Property Get TracefileType() As TracefileTypeConsant
    TracefileType = mvarTracefileType
End Property

Public Property Let InitialCatalogPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InitialCatalogPath = 5
    mvarInitialCatalogPath = vData
End Property


Public Property Get InitialCatalogPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InitialCatalogPath
    InitialCatalogPath = mvarInitialCatalogPath
End Property

Public Property Let OutputFilePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OutputPath = 5
    mvarOutputFilePath = vData
End Property

' Tracefiles Output Folder (MDBPATH)
' ProtectionPLUS License FIle (CPLic.lf)
Public Property Get OutputFilePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OutputPath
    OutputFilePath = mvarOutputFilePath
End Property

Public Property Let InitialCatalog(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InitialCatalog = 5
    mvarInitialCatalog = vData
End Property


Public Property Get InitialCatalog() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InitialCatalog
    InitialCatalog = mvarInitialCatalog
End Property



Public Property Let DataSource(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataSource = 5
    mvarDataSource = vData
End Property


Public Property Get DataSource() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataSource
    DataSource = mvarDataSource
End Property



Public Property Let DatabaseType(ByVal vData As CubeLibDataSource.DatabaseType)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DatabaseType = Form1
    mvarDatabaseType = vData
End Property


Public Property Get DatabaseType() As CubeLibDataSource.DatabaseType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabaseType
    DatabaseType = mvarDatabaseType
End Property



Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let Username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    mvarUsername = vData
End Property


Public Property Get Username() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    Username = mvarUsername
End Property





