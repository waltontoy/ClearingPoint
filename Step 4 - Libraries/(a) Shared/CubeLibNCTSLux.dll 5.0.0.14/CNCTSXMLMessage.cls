VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNCTSXMLMessage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Function CreateXMLMessage(ByRef DataSourceProperties As CDataSourceProperties, _
                                 ByVal UniqueCode As String, _
                                 ByVal DATA_NCTS_ID As Long, _
                                 ByRef DBEdifact As ADODB.Connection, _
                                 ByVal MDBPath As String, _
                                 ByVal XMLMessageType As NCTS_IEM_IDs, _
                                 ByVal SendMode As String, _
                                 ByVal LogicalID As String, _
                                 ByVal MessageID As Long, _
                        Optional ByRef CallingForm As Object) As String
    
        On Error GoTo ErrorHandler
        
        'Set Global Edifact ADO Connection
1000    Set g_conEdifact = DBEdifact
    
        'Set Global variable for Logical ID
1005    G_strLogicalId = LogicalID
    
        'Set Global variable for Send mode
1010    G_strSendMode = SendMode
    
        'Open SADBEL Database
1015     ADOConnectDB g_conSADBEL, DataSourceProperties, DBInstanceType_DATABASE_SADBEL
'1015    OpenDAODatabase G_datSADBEL, MDBPath, "mdb_sadbel.mdb"
    
        'Open Scheduler Database
1020    ADOConnectDB g_conScheduler, DataSourceProperties, DBInstanceType_DATABASE_SCHEDULER
'1020    OpenDAODatabase G_datScheduler, MDBPath, "mdb_scheduler.mdb"
    
        'Set Database Path
'1025    G_strMdbPath = MDBPath
         
        'Set global variable for Unique Code
1030    G_strUniqueCode = Trim(UniqueCode)

        'Set global variable for NOT SENDING language field when segment members are empty
1035    ADORecordsetOpen "SELECT * FROM SEGMENT", g_conScheduler, g_rstSegment, adOpenKeyset, adLockOptimistic
'1035    Set g_rstSegment = G_datScheduler.OpenRecordset("SEGMENT")
        
        'Prepare Class Data
        Select Case XMLMessageType
            Case NCTS_IEM_IDs.NCTS_IEM_ID_IE15
1040            PrepareEDIDepartureClass DataSourceProperties, UniqueCode, MessageID
                
            Case NCTS_IEM_IDs.NCTS_IEM_ID_IE14
1045            PrepareEDICancellationClass DataSourceProperties, UniqueCode, DATA_NCTS_ID, MessageID
                
            Case NCTS_IEM_IDs.NCTS_IEM_ID_IE07
1050            PrepareEDIArrivalClass UniqueCode
            
            Case NCTS_IEM_IDs.NCTS_IEM_ID_IE44
1055            PrepareIE44ArrivalClass DataSourceProperties, UniqueCode
        
            Case Else
                Debug.Assert False
                
        End Select
        
        'Generate XML from Data
1060    CreateXMLMessage = CreateXML(DataSourceProperties, XMLMessageType)
        
        '*********************************************************************************
        'FOR TESTING ONLY
        '*********************************************************************************
        'Dim strTemp As String
        'strTemp = CreateXMLMessage
        
        'Open App.Path & "/test_" & XMLMessageType & ".xml" For Output As #1
        'Print #1, strTemp
        'Close #1
        '*********************************************************************************
        
        '*********************************************************************************
        'Free Memory
        '*********************************************************************************
        
1065    ADORecordsetClose g_rstSegment
'1065    DAORstClose g_rstSegment
        
1070    Set G_clsEDIDeparture = Nothing
1075    Set G_clsEDIArrival = Nothing
        '*********************************************************************************
    
ErrorHandler:
    
Select Case Err.Number
    Case 0      'No Error
        
    Case Else   'Generic error handling
        CallingForm.NCTSLuxAddToTrace "Error in SendDocument, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description, True
        
End Select
    
End Function



