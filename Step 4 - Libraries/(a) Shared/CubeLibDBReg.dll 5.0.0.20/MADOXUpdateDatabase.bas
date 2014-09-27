Attribute VB_Name = "MADOXUpdateDatabase"
' #VBIDEUtils#************************************************************
' * Author           : Larry Rebich
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           : larry@buygold.net
' * Date             : 12/10/2003
' * Purpose          :
' * Project Name     : DBUpdateADO
' * Module Name      : modUpdateDatabaseWithModelUsingADOX
' **********************************************************************
' * Comments         :
' *
' *
' * Example          :
' *
' * History          : Updated by Waty Thierry
' * 2003/03/12 Copyright © 2003, Larry Rebich, using the DELL7500
' * 2003/03/12 larry@larryrebich.com, www.larryrebich.com, 760-771-4730
' * 2003/05/13 Made a tip-of-the-month
' * 2003/12/11 Updated by Waty Thierry to add fields in right order
' *            and some other modifications, optimization...
' *            Better Error Handling
' *
' * See Also         :
' *
' *
' **********************************************************************

Option Explicit
DefLng A-Z

Public gsAppPath        As String              'User can set this
Public gsFileNameLog    As String
Public gsFileNameErrorLog As String

Const mcsLog = "UpdateDatabaseWithModelUsingADOX.log"               '2003/03/22 Default log file name
Const mcsErrorLog = "UpdateDatabaseWithModelUsingADOXErrorLog.log"  'default

Private msAppPath       As String
Private mbLogFilesClearedAndOpened As Boolean   '2003/05/30
Private mbSkipWriteToLog As Boolean      '2003/03/22 Module Level switch
Private mbSkipWriteToErrorLog As Boolean    '2003/05/19
Private msFileNameLog   As String 'fully qualified file name
Private msFileNameErrorLog As String

Private Type udtAddReplaceSkip      'can be added, replaced, skipped
   iAdded               As Integer
   iReplaced            As Integer
   iSkipped             As Integer
End Type

Public Type udtUpdateDatabaseWithModelUsingADOX     'count changes
   bSkipWriteToLog      As Boolean      'if true don't write changes to a log file
   bSkipUpdateRelationships As Boolean  '2003/05/30 Don't do update of relationships, do after copying data
   bAnyErrors           As Boolean      '2003/05/30 Set True if any error
   sFileNameLog         As String       'fully qualified names
   sFileNameErrorLog    As String
   sAppPath             As String
   iAddedTables         As Integer
   iAddedColumns        As Integer
   tIndexes             As udtAddReplaceSkip
   tRelationships       As udtAddReplaceSkip
   tStoredProcedures    As udtAddReplaceSkip
   tViews               As udtAddReplaceSkip
End Type

Private Type udtErrors          'store errors in an array
   lNumber              As Long
   sDescription         As String
End Type
Dim maryErrors()        As udtErrors   'store errors here
Dim miErrorsCount       As Integer

Public Function ADOXIsLinkedTableExisting_F(ByRef ADODatabase As ADODB.Connection, _
                                            ByVal LinkedTableName As String) As Boolean

    Dim objCat As ADOX.Catalog
    Dim objTable As ADOX.Table
        
    ADOXIsLinkedTableExisting_F = False
    
    Set objCat = New ADOX.Catalog
    Set objCat.ActiveConnection = ADODatabase

    For Each objTable In objCat.Tables
        
        If UCase$(Trim$(objTable.Type)) = UCase$(Trim$("LINK")) And _
           UCase$(Trim$(objTable.Name)) = UCase$(Trim$(LinkedTableName)) Then
        
            ADOXIsLinkedTableExisting_F = True
            Exit For
        End If
    Next
    
    objCat.Tables.Refresh
    
    Set objCat = Nothing
End Function

Public Function ADOXIsTableExisting_F(ByRef ADODatabase As ADODB.Connection, _
                                    ByVal TableName As String) As Boolean

    Dim objCat As ADOX.Catalog
    Dim objTable As ADOX.Table
        
    ADOXIsTableExisting_F = False
    
    Set objCat = New ADOX.Catalog
    Set objCat.ActiveConnection = ADODatabase

    For Each objTable In objCat.Tables
        
        If UCase$(Trim$(objTable.Name)) = UCase$(Trim$(TableName)) Then
        
            ADOXIsTableExisting_F = True
            Exit For
        End If
    Next
    
    objCat.Tables.Refresh
    
    Set objCat = Nothing
End Function


Private Sub LogFilesClearedAndOpened(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogFilesClearedAndOpened
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Sub created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   If Not mbLogFilesClearedAndOpened Then
      ClearCounters tUpdateDatabaseWithModelUsingADOX                         'clear the counters by resetting the UDT
      LogOpenAndClear connDB, connModel, tUpdateDatabaseWithModelUsingADOX    '2003/03/22 Open the log file
      LogErrorOpenAndClear tUpdateDatabaseWithModelUsingADOX                  '2003/03/21
      mbLogFilesClearedAndOpened = True
   End If
End Sub

Public Function UpdateDatabaseWithModelUsingADOX(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateDatabaseWithModelUsingADOX
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   LogFilesClearedAndOpened connDB, connModel, tUpdateDatabaseWithModelUsingADOX      '2003/05/30

   On Error GoTo UpdateDatabaseWithModelUsingADOXEH

   LogWriteHeading "Tables"
   Call UpdateTables(connDB, connModel, tUpdateDatabaseWithModelUsingADOX)
   
   LogWriteHeading "Indexes"
   Call UpdateIndexes(connDB, connModel, tUpdateDatabaseWithModelUsingADOX)
   
   If Not tUpdateDatabaseWithModelUsingADOX.bSkipUpdateRelationships Then  '2003/05/30 Added
      LogWriteHeading "Relationships"
      Call UpdateRelationships(connDB, connModel, tUpdateDatabaseWithModelUsingADOX)
   End If
   
   LogWriteHeading "Stored Procedures and Views"
   Call UpdateStoredProceduresAndViews(connDB, connModel, tUpdateDatabaseWithModelUsingADOX)

   If miErrorsCount > 0 Then
      ReportErrors tUpdateDatabaseWithModelUsingADOX, miErrorsCount, maryErrors()
   End If

   Erase maryErrors()
   miErrorsCount = 0

   UpdateDatabaseWithModelUsingADOX = True     'made it this far

   Exit Function

UpdateDatabaseWithModelUsingADOXEH:
   Err.Raise Err.Number, "modUpdateDatabaseWithModelUsingADOX:UpdateDatabaseWithModelUsingADOX", Err.Description

End Function

Public Sub ReportErrorsADOX(tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ReportErrorsADOX
   ' * Parameters       :
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   ReportErrors tUpdateDatabaseWithModelUsingADOX, miErrorsCount, maryErrors

End Sub

Private Function UpdateStoredProceduresAndViews(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateStoredProceduresAndViews
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/06/11 Add error checking, SQL which passes in DAO may fail using ADOX. _
   ' *  A table named 'Names' which is OK in DAO fails when processed by ADOX. _
   ' *  Names' is a reserved SQL word.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   Dim catDB            As New ADOX.Catalog
   Dim catMD            As New ADOX.Catalog
   Dim viewMD           As ADOX.View
   Dim proMD            As ADOX.Procedure
   Dim cmdDB            As New ADODB.Command
   Dim cmdMD            As ADODB.Command
   Dim bReplace         As Boolean
   Dim sSQL             As String

   catDB.ActiveConnection = connDB.ConnectionString
   catMD.ActiveConnection = connModel.ConnectionString

   ' Procedures
   For Each proMD In catMD.Procedures
      On Error GoTo UpdateStoredProceduresAndViewsProcedureEH
      bReplace = False
      With proMD
         If NewerModelProcedureOrDoesNotExist(.Name, catMD.Procedures(.Name).DateModified, catDB) Then
            If ExistsProcedure(.Name, catDB) Then   'delete it if it exists
               catDB.Procedures.Delete .Name
               bReplace = True
            End If
            Set cmdMD = .Command        'get SQL
            sSQL = cmdMD.CommandText    'SQL for new Command
            cmdDB.CommandText = sSQL    'set SQL in new Command
            catDB.Procedures.Append .Name, cmdDB    'store in the application db
            Set cmdDB = Nothing         'no longer needed
            With tUpdateDatabaseWithModelUsingADOX.tStoredProcedures    'update counters
               If bReplace Then
                  .iReplaced = .iReplaced + 1
                  LogWrite "Stored procedure '" & proMD.Name & "' replaced."
               Else
                  .iAdded = .iAdded + 1
                  LogWrite "Stored procedure '" & proMD.Name & "' added."
               End If
            End With
         End If
      End With
      
      DoEvents
UpdateStoredProceduresAndViewsProcedureContinue:
   Next

   ' Views
   For Each viewMD In catMD.Views
      On Error GoTo UpdateStoredProceduresAndViewsViewEH
      bReplace = False
      With viewMD
         If NewerModelViewOrDoesNotExist(.Name, catMD.Views(.Name).DateModified, catDB) Then
            If ExistsView(.Name, catDB) Then    'delete it if it exists
               catDB.Views.Delete .Name
               bReplace = True
            End If
            Set cmdMD = .Command        'get SQL
            sSQL = cmdMD.CommandText    'SQL for new Command
            cmdDB.CommandText = sSQL    'set SQL in new Command
            catDB.Views.Append .Name, cmdDB     'store in the application db
            Set cmdDB = Nothing                 'no longer needed
            With tUpdateDatabaseWithModelUsingADOX.tViews       'update counters
               If bReplace Then
                  .iReplaced = .iReplaced + 1
                  LogWrite "View '" & viewMD.Name & "' replaced."
               Else
                  .iAdded = .iAdded + 1
                  LogWrite "View '" & viewMD.Name & "' added."
               End If
            End With
         End If
      End With
      
      DoEvents
UpdateStoredProceduresAndViewsViewContinue:
   Next

   UpdateStoredProceduresAndViews = True
   Exit Function

UpdateStoredProceduresAndViewsProcedureEH:
   LogWrite "Error, procedure '" & proMD.Name & "', Error: &H" & Hex(Err.Number) & ", " & Err.Description
   miErrorsCount = miErrorsCount + 1
   ReDim Preserve maryErrors(1 To miErrorsCount) As udtErrors
   With maryErrors(miErrorsCount)
      .lNumber = Err.Number
      .sDescription = Err.Description
   End With
   Resume UpdateStoredProceduresAndViewsProcedureContinue

UpdateStoredProceduresAndViewsViewEH:
   LogWrite "Error, view '" & viewMD.Name & "', Error: &H" & Hex(Err.Number) & ", " & Err.Description
   miErrorsCount = miErrorsCount + 1
   ReDim Preserve maryErrors(1 To miErrorsCount) As udtErrors
   With maryErrors(miErrorsCount)
      .lNumber = Err.Number
      .sDescription = Err.Description
   End With
   Resume UpdateStoredProceduresAndViewsViewContinue

End Function

Private Function NewerModelViewOrDoesNotExist(sName As String, dDateModified As Date, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : NewerModelViewOrDoesNotExist
   ' * Parameters       :
   ' *                    sName As String
   ' *                    dDateModified As Date
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   Dim viewDB           As ADOX.View
   
   If ExistsView(sName, catDB) Then
      For Each viewDB In catDB.Views
         With viewDB
            If LCase$(.Name) = LCase$(sName) Then
               If .DateModified < dDateModified Then
                  NewerModelViewOrDoesNotExist = True
               End If
               Exit Function
            End If
         End With
         
         DoEvents
      Next
   Else
      NewerModelViewOrDoesNotExist = True     'does not exist so can is 'newer'
   End If
   
End Function

Private Function NewerModelProcedureOrDoesNotExist(sName As String, dDateModified As Date, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : NewerModelProcedureOrDoesNotExist
   ' * Parameters       :
   ' *                    sName As String
   ' *                    dDateModified As Date
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim proDB            As ADOX.Procedure

   If ExistsProcedure(sName, catDB) Then
      For Each proDB In catDB.Procedures
         With proDB
            If LCase$(.Name) = LCase$(sName) Then
               If .DateModified < dDateModified Then
                  NewerModelProcedureOrDoesNotExist = True
               End If
               Exit Function
            End If
         End With
         
         DoEvents
      Next
   Else
      NewerModelProcedureOrDoesNotExist = True     'does not exist so can is 'newer'
   End If

End Function

Private Function DeleteAllRelationships(connDB As ADODB.Connection) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : DeleteAllRelationships
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   Dim catDB            As New ADOX.Catalog
   Dim tblDB            As ADOX.Table
   Dim keyDB            As ADOX.Key
   Dim aryNames()       As String
   Dim iNamesCount      As Integer
   Dim i                As Integer

   catDB.ActiveConnection = connDB.ConnectionString

   For Each tblDB In catDB.Tables
      With tblDB
         If .Type = "TABLE" Then
            Erase aryNames()
            iNamesCount = 0
            For Each keyDB In .Keys
               With keyDB
                  If .Type = adKeyForeign Then
                     iNamesCount = iNamesCount + 1
                     ReDim Preserve aryNames(1 To iNamesCount) As String
                     aryNames(iNamesCount) = .Name
                  End If
               End With
            Next
            For i = 1 To iNamesCount    'now delete them
               .Keys.Delete aryNames(i)
            Next
            If iNamesCount = 1 Then     'write to log
               LogWrite "One relationship [Foreign Key] removed from table '" & tblDB.Name & "'"
            ElseIf iNamesCount > 1 Then
               LogWrite i & " relationships [Foreign Keys] removed from table '" & tblDB.Name & "'"
            End If
         End If
      End With
      
      DoEvents
   Next

   Set catDB = Nothing
   Set tblDB = Nothing
   Set keyDB = Nothing
   Erase aryNames()
   iNamesCount = 0

   DeleteAllRelationships = True
End Function

Private Function UpdateRelationships(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateRelationships
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   Dim catDB            As New ADOX.Catalog
   Dim catMD            As New ADOX.Catalog
   Dim tblDB            As ADOX.Table
   Dim tblMD            As ADOX.Table
   Dim keyDB            As ADOX.Key
   Dim keyMD            As ADOX.Key
   Dim colDB            As ADOX.Column
   Dim colMD            As ADOX.Column
   Dim bAdd             As Boolean
   Dim bReplace         As Boolean
   Dim i                As Integer

   catDB.ActiveConnection = connDB.ConnectionString
   catMD.ActiveConnection = connModel.ConnectionString

   For Each tblMD In catMD.Tables
      With tblMD
         If .Type = "TABLE" Then
            Set tblDB = catDB.Tables(.Name)
            For Each keyMD In tblMD.Keys
               bAdd = False
               bReplace = False
               With keyMD
                  If .Type = adKeyForeign Then
                     If ExistsRelationship(.Name, tblDB) Then
                        bReplace = True
                        tblDB.Keys.Delete .Name
                     Else
                        bAdd = True
                     End If
                     Set keyDB = New ADOX.Key
                     With keyDB
                        '2003/03/20 Can't use old name, ADOX thinks it is still there even though it has been deleted.
                        .Name = CreateGUIDWithPrefix("rel")
                        .RelatedTable = keyMD.RelatedTable
                        .Type = keyMD.Type
                        .DeleteRule = keyMD.DeleteRule
                        .UpdateRule = keyMD.UpdateRule
                        .Type = keyMD.Type
                        For Each colMD In keyMD.Columns
                           Set colDB = New ADOX.Column
                           With colDB
                              .Name = colMD.Name
                              .RelatedColumn = colMD.RelatedColumn
                           End With
                        Next
                        keyDB.Columns.Append colDB
                     End With
                     On Local Error GoTo UpdateRelationshipsEH        '-2147467259 You cannot add or change a record because a related record is required in table 'tblOrganization'.
                     tblDB.Keys.Append keyDB
                     With tUpdateDatabaseWithModelUsingADOX.tRelationships
                        If bReplace Then
                           .iReplaced = .iReplaced + 1
                           LogWrite "Relationship [Foreign Key] replaced in table '" & tblDB.Name & "' related to table '" & keyDB.RelatedTable & "'"
                        Else
                           .iAdded = .iAdded + 1
                           LogWrite "Relationship [Foreign Key] added to table '" & tblDB.Name & "' related to table '" & keyDB.RelatedTable & "'"
                        End If
                     End With
UpdateRelationshipsContinue:
                  End If
               End With
               
               DoEvents
            Next
         End If
      End With
   Next
   UpdateRelationships = True
   Exit Function

UpdateRelationshipsEH:
   LogWrite "Error, table '" & tblDB.Name & "' Key '" & keyDB.Name & "' Error: &H" & Hex(Err.Number) & ", " & Err.Description
   miErrorsCount = miErrorsCount + 1
   ReDim Preserve maryErrors(1 To miErrorsCount) As udtErrors
   With maryErrors(miErrorsCount)
      .lNumber = Err.Number
      .sDescription = Err.Description
   End With
   Resume UpdateRelationshipsContinue
End Function

Private Function UpdateIndexes(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateIndexes
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   Dim catDB            As New ADOX.Catalog
   Dim catMD            As New ADOX.Catalog
   Dim tblDB            As ADOX.Table
   Dim tblMD            As ADOX.Table
   Dim idxDB            As ADOX.Index
   Dim idxMD            As ADOX.Index
   Dim colMD            As ADOX.Column
   Dim colDB            As ADOX.Column
   Dim bReplace         As Boolean

   DeleteAllRelationships connDB   '2003/03/20 Needed to prevent "Cannot delete this index or table.  It is either the current index or is used in a relationship."

   catDB.ActiveConnection = connDB.ConnectionString
   catMD.ActiveConnection = connModel.ConnectionString

   For Each tblMD In catMD.Tables
      If tblMD.Type = "TABLE" Then
         Set tblDB = catDB.Tables(tblMD.Name)
         With tblMD
            For Each idxMD In .Indexes
               If Not IsKeyForeign(idxMD.Name, tblMD) Then   'all Keys are reported in the index collection so make sure it is not a 'relationship'
                  bReplace = False
                  On Local Error GoTo UpdateIndexesEH
                  If ExistsIndex(idxMD.Name, tblDB) Then
                     tblDB.Indexes.Delete idxMD.Name
                     bReplace = True
                  End If
                  Set idxDB = New ADOX.Index
                  With idxDB
                     .Clustered = idxMD.Clustered
                     .IndexNulls = idxMD.IndexNulls
                     .Name = idxMD.Name
                     .PrimaryKey = idxMD.PrimaryKey
                     .Unique = idxMD.Unique
                     For Each colMD In idxMD.Columns
                        Set colDB = New ADOX.Column
                        With colDB
                           .Name = colMD.Name
                        End With
                        idxDB.Columns.Append colDB
                     Next
                  End With
                  tblDB.Indexes.Append idxDB
                  With tUpdateDatabaseWithModelUsingADOX.tIndexes
                     If bReplace Then
                        .iReplaced = .iReplaced + 1
                        LogWrite "Index '" & idxDB.Name & "' replaced in table '" & tblDB.Name & "'"
                     Else
                        .iAdded = .iAdded + 1
                        LogWrite "Index '" & idxDB.Name & "' added to table '" & tblDB.Name & "'"
                     End If
                  End With
               End If
               
               DoEvents
            Next
         End With
      End If
UpdateIndexesNext:
   Next

   UpdateIndexes = True
   Exit Function

UpdateIndexesEH:
   LogWrite "Error, table '" & tblMD.Name & "' Index '" & idxMD.Name & "' Error: &H" & Hex(Err.Number) & ", " & Err.Description
   Select Case Err.Number
      Case -2147217866    'Cannot delete this index or table.  It is either the current index or is used in a relationship.
         With tUpdateDatabaseWithModelUsingADOX.tIndexes
            .iSkipped = .iSkipped + 1
         End With
         Resume UpdateIndexesNext
      Case Else
         Err.Raise Err.Number, , Err.Description
         Resume UpdateIndexesNext
   End Select
End Function

Private Function ExistsIndex(sName As String, tblDB As ADOX.Table) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsIndex
   ' * Parameters       :
   ' *                    sName As String
   ' *                    tblDB As ADOX.Table
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   Dim idxTB            As ADOX.Index
   
   For Each idxTB In tblDB.Indexes
      With idxTB
         If LCase$(.Name) = LCase$(sName) Then
            ExistsIndex = True
            Exit Function
         End If
      End With
   Next
   
End Function

Private Function UpdateTables(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateTables
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/12/12 : Taking care of setting the properties of the table
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim catDB            As New ADOX.Catalog
   Dim oCatalogSource   As New ADOX.Catalog
   Dim oTableDest       As ADOX.Table
   Dim oTableSource     As ADOX.Table
   Dim oColDest         As ADOX.Column
   Dim colModel        As ADOX.Column
   Dim oRecord          As ADODB.Recordset
   Dim oPropertySrc     As ADODB.Property
   Dim oProperty        As ADODB.Property
   Dim i                As Long
   Dim lngPropertiesCtr As Long

   catDB.ActiveConnection = connDB.ConnectionString
   oCatalogSource.ActiveConnection = connModel.ConnectionString

   For Each oTableSource In oCatalogSource.Tables
      With oTableSource
         If .Type = "TABLE" Then
            If Not ExistsTable(.Name, catDB) Then
            
                If UCase$(oTableSource.Name) = UCase$("Users") Then
                
                       Debug.Assert UCase$(oTableSource.Name) <> UCase$("Users")
                        
                       Set oTableDest = New ADOX.Table
                       With oTableDest
                          .Name = oTableSource.Name
                          
                          ' *** Get all fields in the right order to create them
                          Set oRecord = New ADODB.Recordset
                          oRecord.Open "Select * From [" & .Name & "]", connModel, adOpenForwardOnly
        
                          Set oColDest = New ADOX.Column
                          
                          
                          If oRecord.Fields(oRecord.Fields(0).Name).Properties("ISAUTOINCREMENT").Value = True Then
                            oColDest.Name = oTableSource.Columns(oRecord.Fields(0).Name).Name
                            oColDest.Type = oTableSource.Columns(oRecord.Fields(0).Name).Type
                            
                            ' Must set before setting properties
                            Set oColDest.ParentCatalog = catDB
                            
                            'colTarget.DefinedSize = oTableSource.Columns(oRecord.Fields(0).Name).DefinedSize
                            'colTarget.Attributes = oTableSource.Columns(oRecord.Fields(0).Name).Attributes
                            oColDest.Properties("AutoIncrement").Value = True
                            '
                            
                            Set colModel = oTableSource.Columns(oRecord.Fields(0).Name)
                        
                            For lngPropertiesCtr = 0 To oTableSource.Columns(oRecord.Fields(0).Name).Properties.Count - 1
                                Set oPropertySrc = oTableSource.Columns(oRecord.Fields(0).Name).Properties(lngPropertiesCtr)
                                oColDest.Properties(lngPropertiesCtr).Value = oPropertySrc.Value
                            Next
                           Else
                                With oColDest
                                   .Type = oTableSource.Columns(oRecord.Fields(0).Name).Type
                                   .Name = oTableSource.Columns(oRecord.Fields(0).Name).Name
                                   .DefinedSize = oTableSource.Columns(oRecord.Fields(0).Name).DefinedSize
                                   '.Attributes = oTableSource.Columns(oRecord.Fields(0).Name).Attributes
                                End With
                                
                                Set oColDest.ParentCatalog = catDB
                                
                                Set colModel = oTableSource.Columns(oRecord.Fields(0).Name)
                        
                                For lngPropertiesCtr = 0 To oTableSource.Columns(oRecord.Fields(0).Name).Properties.Count - 1
                                    Set oPropertySrc = oTableSource.Columns(oRecord.Fields(0).Name).Properties(lngPropertiesCtr)
                                    oColDest.Properties(lngPropertiesCtr).Value = oPropertySrc.Value
                                Next
                           End If
                            
                          .Columns.Append oColDest
                          
                           Set oTableDest.Columns(oRecord.Fields(0).Name).ParentCatalog = catDB
                          
                           '*** Now as the column is added, set all his properties
                            On Error Resume Next
                            ' *** Ignore all errors
                            For i = 0 To oTableSource.Columns(oRecord.Fields(0).Name).Properties.Count - 1
        
                               Set oPropertySrc = oTableSource.Columns(oRecord.Fields(0).Name).Properties(i)
                               .Columns(oRecord.Fields(0).Name).Properties(i).Value = oPropertySrc.Value
                            Next
                            On Error GoTo 0
                            
                       End With
                       
                       
                       'Call AddFields(oTableSource, oTableDest, catDB, tUpdateDatabaseWithModelUsingADOX) 'add fields if necessary
                       
                       'catDB.Tables.Refresh
                       
                       Set oTableDest.ParentCatalog = catDB
                       
                       catDB.Tables.Append oTableDest    'append it
                       
                       ' *** Now as the table is added, set all his properties
                       
                       
                       
                       On Error Resume Next
                       ' *** Ignore all errors
                       For i = 0 To oTableSource.Properties.Count - 1
                          Set oPropertySrc = oTableSource.Properties(i)
                          With oTableDest.Properties(i)
                             .Value = oPropertySrc.Value
                       
                          End With
                       Next
                       On Error GoTo 0
                       
                       LogWrite "Table '" & oTableSource.Name & "' added"
                       LogWriteColumn oColDest.Name, oTableDest.Name
                       tUpdateDatabaseWithModelUsingADOX.iAddedTables = tUpdateDatabaseWithModelUsingADOX.iAddedTables + 1
                       UpdateTables = True
                       
                       Set oTableDest = catDB.Tables(.Name)
            
                       Call AddFields(oTableSource, oTableDest, catDB, tUpdateDatabaseWithModelUsingADOX) 'add fields if necessary
                End If
            End If
            
            
            DoEvents
         End If
      End With
   Next

   Set catDB = Nothing
   Set oCatalogSource = Nothing

   UpdateTables = True
   
End Function

Private Function AddFields(oTableSource As ADOX.Table, oTableDest As ADOX.Table, oCatDest As ADOX.Catalog, tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : AddFields
   ' * Parameters       :
   ' *                    oTableSource As ADOX.Table
   ' *                    oTableDest As ADOX.Table
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/12/12 : Taking care of same column order in destination DB
   ' * 2003/12/12 : Taking care of setting the properties of the column
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_AddFields

   Dim colMD            As ADOX.Column
   Dim colDB            As New ADOX.Column
   Dim i                As Integer                '2003/05/30
   Dim oRecord          As ADODB.Recordset
   Dim oField           As ADODB.Field
   Dim oPropertySrc     As ADODB.Property
   Dim oProperty        As ADODB.Property
   
   ' *** Get all fields in the right order to create them
   Set oRecord = New ADODB.Recordset
   oRecord.Open "Select * From [" & oTableSource.Name & "]", oTableSource.ParentCatalog.ActiveConnection, adOpenForwardOnly
   
   For Each oField In oRecord.Fields
      
      'Debug.Assert UCase$(oField.Name) <> UCase$("User_ID")
      
      Set colMD = oTableSource.Columns(oField.Name)
      With colMD
         If Not ExistsColumn(.Name, oTableDest) Then
            Set colDB = New ADOX.Column
            With colDB
               .Name = colMD.Name
               .Type = colMD.Type
               .DefinedSize = colMD.DefinedSize
               .Attributes = colMD.Attributes
            
                'Set .ParentCatalog = oCatDest
            End With
            
            oTableDest.Columns.Append colDB
            
            'oTableDest.Columns.Append colMD.Name, colMD.Type, colMD.DefinedSize
            
            'oTableDest.Columns(colMD.Name).Attributes = colMD.Attributes
            Set oTableDest.Columns(colMD.Name).ParentCatalog = oCatDest
            
            ' *** Now as the column is added, set all his properties
            
            ' *** Ignore all errors
            For i = 0 To colMD.Properties.Count - 1
               
               Debug.Assert Not (UCase$(colMD.Properties(i).Name) = UCase$("Autoincrement") And _
                                  UCase$(oTableSource.Name) = UCase$("Users") And _
                            UCase$(oField.Name) = UCase$("User_ID"))
                  
               
               Set oPropertySrc = colMD.Properties(i)
               With oTableDest.Columns(colMD.Name).Properties(i)
                
                  Debug.Assert UCase$(.Name) <> "Autoincrement"
                  
                  On Error Resume Next
                  .Value = oPropertySrc.Value
                  On Error GoTo 0
                  
               End With
            Next
            
            'oTableDest.Columns.Append colDB
            
            LogWriteColumn colMD.Name, oTableDest.Name
            tUpdateDatabaseWithModelUsingADOX.iAddedColumns = tUpdateDatabaseWithModelUsingADOX.iAddedColumns + 1
            AddFields = True
            Set colDB = Nothing
         End If
      End With
      Set colMD = Nothing
      
      DoEvents
   Next
   
   oRecord.Close
   Set oRecord = Nothing

EXIT_AddFields:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_AddFields:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in AddFields", vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_AddFields
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   
   Resume EXIT_AddFields

End Function

Private Function ExistsTable(sName As String, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsTable
   ' * Parameters       :
   ' *                    sName As String
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim tblDB            As ADOX.Table

   For Each tblDB In catDB.Tables
      With tblDB
         If .Type = "TABLE" Then
            If LCase$(.Name) = LCase$(sName) Then
               ExistsTable = True  'exists, exit
               Exit Function
            End If
         End If
      End With
   Next

End Function

Private Function ExistsColumn(sName As String, tblDB As ADOX.Table) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsColumn
   ' * Parameters       :
   ' *                    sName As String
   ' *                    tblDB As ADOX.Table
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim colDB            As ADOX.Column
   For Each colDB In tblDB.Columns
      With colDB
         If LCase$(.Name) = LCase$(sName) Then
            ExistsColumn = True
            Exit Function
         End If
      End With
   Next
End Function

Private Function ExistsRelationship(sName As String, tblDB As ADOX.Table) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsRelationship
   ' * Parameters       :
   ' *                    sName As String
   ' *                    tblDB As ADOX.Table
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim keyDB            As ADOX.Key
   For Each keyDB In tblDB.Keys
      With keyDB
         If LCase$(.Name) = LCase$(sName) Then
            ExistsRelationship = True
            Exit Function
         End If
      End With
   Next
End Function

Private Function ExistsView(sName As String, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsView
   ' * Parameters       :
   ' *                    sName As String
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/20 Function created by Larry Rebich while in La Quinta, CA.
   Dim viewDB           As ADOX.View
   For Each viewDB In catDB.Views
      With viewDB
         If LCase$(.Name) = LCase$(sName) Then
            ExistsView = True
            Exit Function
         End If
      End With
   Next
End Function

Private Function ExistsProcedure(sName As String, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ExistsProcedure
   ' * Parameters       :
   ' *                    sName As String
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/19 Function created by Larry Rebich while in La Quinta, CA.
   Dim proDB            As ADOX.Procedure
   For Each proDB In catDB.Procedures
      With proDB
         If LCase$(.Name) = LCase$(sName) Then
            ExistsProcedure = True
            Exit Function
         End If
      End With
   Next
End Function

Private Function NewerModelTable(sName As String, dDateModified As Date, catDB As ADOX.Catalog) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : NewerModelTable
   ' * Parameters       :
   ' *                    sName As String
   ' *                    dDateModified As Date
   ' *                    catDB As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/18 Function created by Larry Rebich while in La Quinta, CA.
   Dim tblDB            As ADOX.Table
   Set tblDB = catDB.Tables(sName)
   With tblDB
      If .DateModified < dDateModified Then
         NewerModelTable = True
      End If
   End With
End Function

Private Sub ClearCounters(tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ClearCounters
   ' * Parameters       :
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim bTemp            As Boolean
   Dim bTemp2           As Boolean
   Dim tTemp            As udtUpdateDatabaseWithModelUsingADOX

   bTemp = tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog       'save this
   bTemp2 = tUpdateDatabaseWithModelUsingADOX.bSkipUpdateRelationships 'and this

   tUpdateDatabaseWithModelUsingADOX = tTemp                       'copy storage causes clear of counters

   tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog = bTemp           'restore this
   tUpdateDatabaseWithModelUsingADOX.bSkipUpdateRelationships = bTemp2 'and this
End Sub

Private Sub LogOpenAndClear(connDB As ADODB.Connection, connModel As ADODB.Connection, tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogOpenAndClear
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim iFN              As Integer

   mbSkipWriteToLog = tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog    'local variable

   On Error GoTo LogOpenAndClearEH     '2003/05/19

   If Not mbSkipWriteToLog Then
      If gsAppPath = "" Then
         msAppPath = AddBackSlashOnPath(App.Path)
      Else
         msAppPath = gsAppPath
      End If
      If gsFileNameLog = "" Then
         msFileNameLog = AddBackSlashOnPath(msAppPath) & mcsLog
      Else
         msFileNameLog = gsFileNameLog
      End If
      tUpdateDatabaseWithModelUsingADOX.sFileNameLog = msFileNameLog      '2003/05/22
      tUpdateDatabaseWithModelUsingADOX.sFileNameErrorLog = msFileNameErrorLog    '2003/05/22
      tUpdateDatabaseWithModelUsingADOX.sAppPath = msAppPath              '2003/05/22
      iFN = FreeFile
      Open msFileNameLog For Output As #iFN
      Print #iFN, msFileNameLog, Now
      Print #iFN, "Connection; Application=" & connDB
      Print #iFN, "Connection; ModelDB    =" & connModel
      Close #iFN
   End If
   Exit Sub
LogOpenAndClearEH:
   tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog = True
   mbSkipWriteToLog = tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog    'local variable
   Err.Raise Err.Number, , Err.Description
End Sub

Private Sub LogWrite(sMessage As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogWrite
   ' * Parameters       :
   ' *                    sMessage As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/22 Sub created by Larry Rebich while in La Quinta, CA.
   Dim iFN              As Integer

   If Not mbSkipWriteToLog Then
      iFN = FreeFile
      Open msFileNameLog For Append As #iFN
      Print #iFN, sMessage
      Close #iFN
   End If
End Sub

Private Sub LogWriteHeading(sHeading As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogWriteHeading
   ' * Parameters       :
   ' *                    sHeading As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim sMessage         As String
   sMessage = "Processing: " & sHeading & " " & String$(120, "-")
   sMessage = Left$(sMessage, 120)
   LogWrite sMessage
End Sub

Private Sub LogWriteColumn(sNameColumn As String, sNameTable As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogWriteColumn
   ' * Parameters       :
   ' *                    sNameColumn As String
   ' *                    sNameTable As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/24 Sub created by Larry Rebich while in La Quinta, CA.
   ' 2003/03/24 Standardize in one location
   LogWrite "Column '" & sNameColumn & "' added to table '" & sNameTable & "'"
End Sub

Private Sub LogErrorKill()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogErrorKill
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/24 Sub created by Larry Rebich while in La Quinta, CA.
   On Error Resume Next        'no big deal if this fails.
   Kill msFileNameErrorLog
End Sub

Private Sub LogErrorOpenAndClear(t As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : LogErrorOpenAndClear
   ' * Parameters       :
   ' *                    t As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim iFN              As Integer

   On Local Error GoTo LogErrorOpenAndClearEH      '2003/05/19

   Erase maryErrors()
   miErrorsCount = 0

   If gsFileNameErrorLog = "" Then
      msFileNameErrorLog = AddBackSlashOnPath(msAppPath) & mcsErrorLog
   Else
      msFileNameErrorLog = gsFileNameErrorLog
   End If
   t.sFileNameErrorLog = msFileNameErrorLog
   t.bAnyErrors = False

   iFN = FreeFile
   Open msFileNameErrorLog For Output As #iFN
   Print #iFN, msFileNameErrorLog, Now
   Print #iFN, "Err.Number ", "Err.Description"
   Close #iFN
   Exit Sub
LogErrorOpenAndClearEH:
   mbSkipWriteToErrorLog = True
   Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ReportErrors(t As udtUpdateDatabaseWithModelUsingADOX, iErrorsCount As Integer, aryErrors() As udtErrors)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : ReportErrors
   ' * Parameters       :
   ' *                    t As udtUpdateDatabaseWithModelUsingADOX
   ' *                    iErrorsCount As Integer
   ' *                    aryErrors() As udtErrors
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim sMsg             As String
   Dim sMsg2            As String

   If mbSkipWriteToErrorLog Then Exit Sub  '2003/05/19

   If iErrorsCount > 0 Then
      t.bAnyErrors = True
   End If

   WriteErrorLog iErrorsCount, aryErrors()

   sMsg2 = vbCrLf & "See log file '" & msFileNameErrorLog & "' for more information."

   If iErrorsCount > 0 Then
      sMsg = iErrorsCount & " errors where encountered while updating the database." & sMsg2
      MsgBox sMsg, vbExclamation, iErrorsCount & " Errors"
   End If

   Erase maryErrors()
   iErrorsCount = 0
End Sub

Private Sub WriteErrorLog(iErrorsCount As Integer, aryErrors() As udtErrors)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : WriteErrorLog
   ' * Parameters       :
   ' *                    iErrorsCount As Integer
   ' *                    aryErrors() As udtErrors
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/21 Sub created by Larry Rebich while in La Quinta, CA.
   Dim iFN              As Integer
   Dim i                As Integer

   If iErrorsCount > 0 Then
      iFN = FreeFile
      Open msFileNameErrorLog For Append As #iFN
      For i = 1 To iErrorsCount
         With aryErrors(i)
            Print #iFN, "&H" & Hex(.lNumber), .sDescription
         End With
      Next
      Close #iFN
   End If
End Sub

Private Function IsKeyForeign(sName As String, tbl As ADOX.Table) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : IsKeyForeign
   ' * Parameters       :
   ' *                    sName As String
   ' *                    tbl As ADOX.Table
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/03/24 Function created by Larry Rebich while in La Quinta, CA.
   Dim keyFN            As ADOX.Key

   On Local Error GoTo IsKeyForeignEH      'in unlikely case not there!
   Set keyFN = tbl.Keys(sName)
   If keyFN.Type = adKeyForeign Then
      IsKeyForeign = True
   End If
IsKeyForeignEH:
End Function

Public Function UpdateTablesWithModelUsingADOX(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateTablesWithModelUsingADOX
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.

   LogFilesClearedAndOpened connDB, connModel, tUpdateDatabaseWithModelUsingADOX   '2003/05/30

   LogWriteHeading "Tables"
   If UpdateTables(connDB, connModel, tUpdateDatabaseWithModelUsingADOX) Then
      UpdateTablesWithModelUsingADOX = True
   End If

End Function

Public Function UpdateIndexesWithModelUsingADOX(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateIndexesWithModelUsingADOX
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.

   LogFilesClearedAndOpened connDB, connModel, tUpdateDatabaseWithModelUsingADOX   '2003/05/30

   LogWriteHeading "Indexes"
   If UpdateIndexes(connDB, connModel, tUpdateDatabaseWithModelUsingADOX) Then
      UpdateIndexesWithModelUsingADOX = True
   End If

End Function

Public Function UpdateStoredProceduresAndViewsWithModelUsingADOX(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateStoredProceduresAndViewsWithModelUsingADOX
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.

   LogFilesClearedAndOpened connDB, connModel, tUpdateDatabaseWithModelUsingADOX   '2003/05/30

   LogWriteHeading "Stored Procedures and Views"
   If UpdateStoredProceduresAndViews(connDB, connModel, tUpdateDatabaseWithModelUsingADOX) Then
      UpdateStoredProceduresAndViewsWithModelUsingADOX = True
   End If
End Function

Public Function UpdateRelationshipsWithModelUsingADOX(connDB As ADODB.Connection, connModel As ADODB.Connection, _
   tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modUpdateDatabaseWithModelUsingADOX
   ' * Procedure Name   : UpdateRelationshipsWithModelUsingADOX
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' *                    connModel As ADODB.Connection
   ' *                    tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.

   LogFilesClearedAndOpened connDB, connModel, tUpdateDatabaseWithModelUsingADOX   '2003/05/30

   LogWriteHeading "Relationships"
   If UpdateRelationships(connDB, connModel, tUpdateDatabaseWithModelUsingADOX) Then
      UpdateRelationshipsWithModelUsingADOX = True
   End If
End Function


