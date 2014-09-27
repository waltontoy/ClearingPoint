VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILinkedDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CreateLinkedTables(ByRef DataSourceProperties As CDataSourceProperties, _
                              ByRef UserID As Long, _
                    Optional ByVal ContinueMissingDB As Boolean = False)

'Public Sub CreateLinkedTables(ByVal DBPath As String, _
                              ByVal TemplateFileName As String, _
                              ByVal DBPassword As String, _
                              ByRef ADOConnection As ADODB.Connection, _
                              ByRef UserID As Long, _
                    Optional ByVal ContinueMissingDB As Boolean = False)
                    
    Dim clsDB As CLinkedDB
    Dim rstLinkedTables As ADODB.Recordset
    Dim enuResult As PCubeLibLinkedDB.QueryResultConstants
    Dim objADOConnection As ADODB.Connection
    Dim enuDBInstanceType As DBInstanceType
    Dim strSourceDBName As String
    Dim strYearSuffix As String
    
    Set clsDB = New CLinkedDB
    
    ADOConnectDB objADOConnection, DataSourceProperties, DBInstanceType_DATABASE_TEMPLATE
    
    enuResult = clsDB.GetLinkedTables(objADOConnection, rstLinkedTables)
    
    Select Case enuResult
        Case QueryResultSuccessful
            Do While Not rstLinkedTables.EOF
            
                strSourceDBName = UCase$(Trim$(FNullField(rstLinkedTables![DB_Name])))
                strSourceDBName = Replace(strSourceDBName, ".MDB", "")
                
                Select Case strSourceDBName
                    Case "MDB_DATA"
                        enuDBInstanceType = DBInstanceType_DATABASE_DATA
                    Case "MDB_SADBEL"
                        enuDBInstanceType = DBInstanceType_DATABASE_SADBEL
                    Case "MDB_SCHEDULER"
                        enuDBInstanceType = DBInstanceType_DATABASE_SCHEDULER
                    Case "MDB_TARIC"
                        enuDBInstanceType = DBInstanceType_DATABASE_TARIC
                    Case "MDB_REPERTORY"
                        enuDBInstanceType = DBInstanceType_DATABASE_REPERTORY
                        strYearSuffix = vbNullString
                    Case "EDIFACT"
                        enuDBInstanceType = DBInstanceType_DATABASE_EDIFACT
                    Case "TEMPLATECP"
                        enuDBInstanceType = DBInstanceType_DATABASE_TEMPLATE
                    Case Else
                        
                        If Left$(strSourceDBName, 11) = UCase$("MDB_HISTORY") Then
                            enuDBInstanceType = DBInstanceType_DATABASE_HISTORY
                            strYearSuffix = Right$(strSourceDBName, 2)
                        ElseIf Left$(strSourceDBName, 14) = UCase$("MDB_EDIHISTORY") Then
                            enuDBInstanceType = DBInstanceType_DATABASE_EDI_HISTORY
                            strYearSuffix = Right$(strSourceDBName, 2)
                        ElseIf Left$(strSourceDBName, 13) = UCase$("MDB_REPERTORY") Then
                            enuDBInstanceType = DBInstanceType_DATABASE_REPERTORY
                            strYearSuffix = Right$(strSourceDBName, 4)
                        Else
                            Debug.Assert False
                        End If
                        
                End Select
                
                enuResult = clsDB.AddLinkedTable(DataSourceProperties, _
                                                 rstLinkedTables![Table_NewName] & Format(UserID, "00"), _
                                                 DBInstanceType_DATABASE_TEMPLATE, _
                                                 rstLinkedTables![Table_SourceName], _
                                                 enuDBInstanceType, _
                                                 "", _
                                                 strYearSuffix)


                If enuResult = QueryResultError Then
                    If ContinueMissingDB = False Then
                        Exit Do
                    End If
                End If
                rstLinkedTables.MoveNext
            Loop
            
        Case QueryResultNoRecord
        
        Case QueryResultError
    
    End Select
    
    Set clsDB = Nothing
    
    ' hobbes 10/18/2005
    ADORecordsetClose rstLinkedTables
    
    ADODisconnectDB objADOConnection
    
    If enuResult = QueryResultError Then
    
    End If
    
End Sub
