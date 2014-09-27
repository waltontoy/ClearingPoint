VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClearingPoint"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function ADOConnectDBRepertory(ByRef ADOConnection As ADODB.Connection, _
                                      ByRef DataSourceProperties As CDataSourceProperties, _
                                      ByVal RepertoryDBName As String) As ErrObject
    Dim strUsedDBYear As String
    
    strUsedDBYear = GetRepertoryDBYear(RepertoryDBName)
    If LenB(Trim$(strUsedDBYear)) > 0 Then
        Set ADOConnectDBRepertory = ADOConnectDB_F(ADOConnection, DataSourceProperties, DBInstanceType_DATABASE_REPERTORY, strUsedDBYear)
    Else
        Set ADOConnectDBRepertory = ADOConnectDB_F(ADOConnection, DataSourceProperties, DBInstanceType_DATABASE_REPERTORY)
    End If
 
End Function

Public Function GetRepertoryDBYear(ByVal RepertoryDBName As String) As String
    
    GetRepertoryDBYear = GetRepertoryDBYear_F(RepertoryDBName)
    
End Function

Public Function GetEDIHistoryDBYear(ByVal EDIHistoryDBName As String) As String

    GetEDIHistoryDBYear = GetEDIHistoryDBYear_F(EDIHistoryDBName)
    
End Function

Public Function GetHistoryDBYear(ByVal HistoryDBName As String) As String
    
    GetHistoryDBYear = GetHistoryDBYear_F(HistoryDBName)
    
End Function

Public Function IsCPDatabase(ByVal DBName As String) As Boolean
    
    IsCPDatabase = IsCPDatabase_F(DBName)
    
End Function

Public Function IsEDIHistoryDB(ByVal DBName As String) As Boolean
    
    IsEDIHistoryDB = IsEDIHistoryDB_F(DBName)
    
End Function

Public Function IsHistoryDB(ByVal DBName As String) As Boolean
    
    IsHistoryDB = IsHistoryDB_F(DBName)
    
End Function

Public Function IsRepertoryDB(ByVal DBName As String) As Boolean
    
    IsRepertoryDB = IsRepertoryDB_F(DBName)
    
End Function
