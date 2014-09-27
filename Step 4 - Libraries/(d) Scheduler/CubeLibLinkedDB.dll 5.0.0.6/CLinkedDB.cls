VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLinkedDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Friend Function AddLinkedTable(ByRef DataSourceProperties As CDataSourceProperties, _
                               ByVal NewTableName As String, _
                               ByVal DestinationDB As DBInstanceType, _
                               ByVal SourceTable As String, _
                               ByVal SourceDB As DBInstanceType, _
                Optional ByVal DestinationDBYearSuffix As String = vbNullString, _
                Optional ByVal SourceDBYearSuffix As String = vbNullString) As PCubeLibLinkedDB.QueryResultConstants
                
    On Error GoTo Error_Handler

    CreateLinkedTable DataSourceProperties, DestinationDB, NewTableName, SourceDB, SourceTable, DestinationDBYearSuffix, SourceDBYearSuffix
    
    AddLinkedTable = QueryResultSuccessful
    
    Exit Function
    
Error_Handler:
    
    AddLinkedTable = QueryResultError

End Function

Friend Function GetLinkedTables(ByRef ADOConnection As ADODB.Connection, ByRef LinkedTables As ADODB.Recordset) As PCubeLibLinkedDB.QueryResultConstants
    
    Dim rstLinkedTables As ADODB.Recordset
    Dim strCommandText As String
    
        strCommandText = vbNullString
        strCommandText = strCommandText & "SELECT "
        strCommandText = strCommandText & "* "
        strCommandText = strCommandText & "FROM "
        strCommandText = strCommandText & "LinkedTables "
        strCommandText = strCommandText & "INNER JOIN "
        strCommandText = strCommandText & "Databases "
        strCommandText = strCommandText & "ON "
        strCommandText = strCommandText & "LinkedTables.DB_ID = Databases.DB_ID "
    ADORecordsetOpen strCommandText, ADOConnection, rstLinkedTables, adOpenKeyset, adLockOptimistic
    
    If (rstLinkedTables.EOF And rstLinkedTables.BOF) Then
        GetLinkedTables = QueryResultNoRecord
    Else
        Set LinkedTables = rstLinkedTables
        GetLinkedTables = QueryResultSuccessful
    End If
    
    Set rstLinkedTables = Nothing
    
End Function
