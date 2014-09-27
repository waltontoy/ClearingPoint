VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
Public Enum RecordOperation
    cpiRecordAdd = 0
    cpiRecordEdit = 1
    cpiRecordCopy = 2
    cpiRecordDelete = 3
    cpiRecordReturn = 4
End Enum

Public Enum DataStatus
    cpiStateNew = 0
    cpiStateOld = 1
    cpiStateModified = 2
    cpiStateDeleted = 3
End Enum

' native
Private mvarStatus As Variant
Private mvarKey As String
Private mvarRecordSQL As String
Private mvarTempDBID As Variant
Private mvarDBID As Variant

' ADO objects
Private mvarADOFields As CADOFields
Private mvarRecordSource As ADODB.Recordset
Private mvarOldRecordSource As ADODB.Recordset

'cubepoint classes
'transactions classes
Private mvarChildTransactions As CTransactions
Private mvarMotherTransactions As CTransactions
Private mvarActiveFilters As CPicklistFilters

Public Property Set ActiveFilters(ByVal vData As CPicklistFilters)
    Set mvarActiveFilters = vData
End Property

Public Property Get ActiveFilters() As CPicklistFilters
    Set ActiveFilters = mvarActiveFilters
End Property

Public Property Let DBID(ByVal vData)
    mvarDBID = vData
End Property

Public Property Get DBID()
    DBID = mvarDBID
End Property

Public Property Let TempDBID(ByVal vData)
    mvarTempDBID = vData
End Property

Public Property Get TempDBID()
    TempDBID = mvarTempDBID
End Property

Public Property Set MotherTransactions(ByVal vData As CTransactions)
    Set mvarMotherTransactions = vData
End Property

Public Property Get MotherTransactions() As CTransactions
    Set MotherTransactions = mvarMotherTransactions
End Property

Public Property Set ChildTransactions(ByVal vData As CTransactions)
    Set mvarChildTransactions = vData
End Property

Public Property Get ChildTransactions() As CTransactions
    If mvarChildTransactions Is Nothing Then
        Set mvarChildTransactions = New CTransactions
    End If
    
    Set ChildTransactions = mvarChildTransactions
End Property

'mvarChildTransactions.Add
Friend Property Let RecordSQL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecordSQL = 5
    mvarRecordSQL = vData
End Property

Friend Property Get RecordSQL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecordSQL
    RecordSQL = mvarRecordSQL
End Property

Public Property Let Status(ByVal vData As DataStatus)
    mvarStatus = vData
End Property

Public Property Get Status() As DataStatus
    Status = mvarStatus
End Property

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Set RecordSource(ByVal vData As ADODB.Recordset)
'    If vData.ActiveConnection Is Nothing Then
        Set mvarRecordSource = vData
'    Else
'        Err.Raise "The recordsource should be disconnected from the database"
'        Err.Clear
'    End If
End Property

Public Property Get RecordSource() As ADODB.Recordset
    Set RecordSource = mvarRecordSource
End Property

Public Property Set OldRecordSource(ByVal vData As ADODB.Recordset)
    Set mvarOldRecordSource = vData
End Property

Public Property Get OldRecordSource() As ADODB.Recordset
    Set OldRecordSource = mvarOldRecordSource
End Property

Friend Property Get ADOFields() As CADOFields
    If mvarADOFields Is Nothing Then
        Set mvarADOFields = New CADOFields
    End If

    Set ADOFields = mvarADOFields
End Property


Friend Property Set ADOFields(vData As CADOFields)
    Set mvarADOFields = vData
End Property

Private Sub Class_Initialize()
   'ado classes
    Set mvarADOFields = New CADOFields
    
    ' cubepoint classes
   'Set mvarActiveFilters = New CPicklistFilters
   
'    Set mvarRecordSource = New CRecord
'    Set mvarOldRecordSource = New CRecord
'    Set mvarChildTransactions = New CTransactions
'    Set mvarMotherTransactions = New CTransactions

End Sub

Private Sub Class_Terminate()
   'ado classes
    Set mvarADOFields = Nothing
    
     ' cubepoint classes
    Set mvarRecordSource = Nothing
    Set mvarOldRecordSource = Nothing
    Set mvarChildTransactions = Nothing
    Set mvarMotherTransactions = Nothing
    Set mvarActiveFilters = Nothing
    
   
End Sub

Public Function LoadData(ByRef DBConnection As ADODB.Connection _
                                                , ByVal SQL As String _
                                                , ByVal Operation As RecordOperation _
                                                , ByVal CurrentPosition As Long _
                                                , Optional ByVal PKBaseFieldName As String _
                                                , Optional ByVal PKFieldValue As String _
                                                , Optional ByVal PKFieldDataType As String _
                                                , Optional ByVal SelRecordTag) As Boolean
                                                
    Dim clsRecordset As CRecordset
    Dim rstParsed As ADODB.Recordset
    Dim strDummySQL As String
    
    LoadData = True
    mvarRecordSQL = SQL
    
    Set clsRecordset = New CRecordset
    Set rstParsed = New ADODB.Recordset
    
    If (IsMissing(PKBaseFieldName) Or IsMissing(PKFieldValue)) _
         And (Operation = cpiRecordCopy Or Operation = cpiRecordEdit) Then
         
        LoadData = False
        Err.Raise "1567", , "Missing parameter."
        Err.Clear
        Exit Function
    End If
        
    strDummySQL = SQL
'    strOldSQL = SQL
    
    Select Case Operation
      Case cpiRecordAdd ', cpiRecordCopy
        strDummySQL = RegenerateSQL(strDummySQL)
      
      Case cpiRecordDelete, cpiRecordEdit, cpiRecordCopy
         If SelRecordTag = "A" Then
             strDummySQL = RegenerateSQL(strDummySQL, True)
'            Call RstCopy(mvarOldRecordSource, rstParsed, True, mvarOldRecordSource.AbsolutePosition - 1 _
                                    , mvarOldRecordSource.AbsolutePosition - 1, False)
'             Set mvarRecordSource = rstParsed
'            'mvarRecordSource.MoveFirst
'            'Debug.Print "BREAKPOINT"
'            Exit Function
         Else
            strDummySQL = RegenerateRecordForPKSQL(strDummySQL, PKBaseFieldName, _
                  PKFieldValue, PKFieldDataType)
            'strDummySQL = RegenerateSQL(strDummySQL, True)
         End If
    End Select
    
'    If (Operation <> cpiRecordAdd) Then
'        ' Modify SQL to return specific data
'        strDummySQL = RegenerateRecordForPKSQL(strDummySQL, PKBaseFieldName, _
'               PKFieldValue, PKFieldDataType)
'    Else
'        ' Recreate SQL to return at most only one line record
'        ' use by mvarRecordSource
'        strDummySQL = RegenerateSQL(strDummySQL)
'        ' get all the fields from the database
''        strOldSQL = RegenerateSQL(strOldSQL, False)
'   End If
        
    ' Open Recordset for mvarRecordSource  ' bug HERE
    ADORecordsetOpen strDummySQL, DBConnection, rstParsed, adOpenKeyset, adLockOptimistic
    'Call clsRecordset.cpiOpen(strDummySQL, DBConnection, rstParsed, _
                                 adOpenKeyset, adLockOptimistic, , True)
'    Call UpdateOldRecordSource(CatalogForm.jgxPicklist.ADORecordset)
    
    'Call clsRecordset.cpiOpen(strDummySQL, DBConnection, mvarOldRecordSource, _
                                 adOpenKeyset, adLockOptimistic, , True)
    Set mvarRecordSource = rstParsed
    
    Select Case Operation
        Case cpiRecordDelete, cpiRecordEdit, cpiRecordCopy
         mvarRecordSource.AbsolutePosition = CurrentPosition
      End Select
   Call CheckIfAdd(Operation)
        
   Set clsRecordset = Nothing
        
End Function

Public Function RegenerateRecordForPKSQL(ByVal SQL As String, ByVal PKBaseFieldName As String, ByVal PKFieldValue As String, ByVal PKBaseFieldDataType As Long, Optional ByVal BaseTable As String = "") As String
'''''    Dim clsSQLQuotes As cpiRecordset
    
    Dim strDummySQL As String
    Dim strWhereClauseCheck As String
    Dim strDummySQLLeft As String
    Dim strDummySQLRight As String
    Dim strBaseTable As String
    
    Dim strPKCondition As String
    
    Dim intCharcterCtr As Long
    
    Dim lngWherePosition As Long
    Dim lngGroupByPosition As Long
    Dim lngOrderByPosition As Long
    Dim lngHavingPosition As Long
    
    Dim lngInsertPosition As Long
    Dim bytClauseInsertType As Byte
    
    
    
    strDummySQLRight = SQL
    
    strPKCondition = GetRecordCriteria(PKBaseFieldName, PKFieldValue, PKBaseFieldDataType, BaseTable)
    
    lngWherePosition = InStr(1, UCase(strDummySQLRight), " WHERE ")
    
    If lngWherePosition > 0 Then
        strDummySQLLeft = Left(strDummySQLRight, InStr(1, UCase(strDummySQLRight), " WHERE ") + 6)
        strDummySQLRight = Mid(strDummySQLRight, InStr(1, UCase(strDummySQLRight), " WHERE ") + 7)
    End If
        
    lngGroupByPosition = InStr(1, UCase(strDummySQLRight), " GROUP BY ")
    lngHavingPosition = InStr(1, UCase(strDummySQLRight), " HAVING ")
    lngOrderByPosition = InStr(1, UCase(strDummySQLRight), " ORDER BY ")
    
    If Not (lngGroupByPosition = 0 And lngOrderByPosition = 0 And lngHavingPosition = 0) Then
        If lngGroupByPosition > 0 Then
            lngInsertPosition = lngGroupByPosition
            bytClauseInsertType = 1
        End If
        If lngHavingPosition > 0 And (lngHavingPosition < lngInsertPosition Or lngInsertPosition = 0) Then
            lngInsertPosition = lngHavingPosition
            bytClauseInsertType = 2
        End If
        If lngOrderByPosition > 0 And (lngOrderByPosition < lngInsertPosition Or lngInsertPosition = 0) Then
            lngInsertPosition = lngOrderByPosition
            bytClauseInsertType = 3
        End If
    End If
    
    If lngInsertPosition > 0 Then
        If lngWherePosition > 0 Then
            If Trim(strPKCondition) <> "" Then
                strDummySQLLeft = strDummySQLLeft & " WHERE " & strPKCondition & " AND ("
                strDummySQLLeft = strDummySQLLeft & Left(strDummySQLRight, lngInsertPosition - 1) & ") "
                
                strDummySQLRight = Mid(strDummySQLRight, lngInsertPosition)
            End If
            
            RegenerateRecordForPKSQL = strDummySQLLeft & strDummySQLRight
        Else
            strDummySQLLeft = Left(strDummySQLRight, lngInsertPosition - 1)
            strDummySQLRight = Mid(strDummySQLRight, lngInsertPosition)
            
            If Trim(strPKCondition) <> "" Then
                strDummySQLLeft = strDummySQLLeft & " WHERE " & strPKCondition
            End If
            
            RegenerateRecordForPKSQL = strDummySQLLeft & strDummySQLRight
        End If
    Else
        If lngWherePosition > 0 Then
            If Trim(strPKCondition) <> "" Then
                strDummySQLLeft = strDummySQLLeft & strPKCondition & " AND ("
                strDummySQLRight = strDummySQLRight & ") "
            End If
            
            RegenerateRecordForPKSQL = strDummySQLLeft & strDummySQLRight
        Else
            If Trim(strPKCondition) <> "" Then
                RegenerateRecordForPKSQL = strDummySQLRight & " WHERE " & strPKCondition
            Else
                RegenerateRecordForPKSQL = strDummySQLRight
            End If
        End If
    End If
    
    strDummySQLRight = RegenerateRecordForPKSQL
    strDummySQLLeft = Left(strDummySQLRight, InStr(1, UCase(strDummySQLRight), " FROM ") - 1)
    strDummySQLRight = Mid(strDummySQLRight, InStr(1, UCase(strDummySQLRight), " FROM "))
    
    If InStr(1, strDummySQLLeft, "*") <= 0 Then
        RegenerateRecordForPKSQL = strDummySQLLeft & ", *" & strDummySQLRight
    End If
End Function

'Public Sub UpdateOldRecordSource(ByVal rstOldRecordSource As ADODB.Recordset _
'                                                                     , Optional ByVal Start As Long _
'                                                                     , Optional ByVal Count As Long)
'
'   Dim intIndex As Integer
'   Dim lngPosition As Long
'   Dim rstTemp As ADODB.Recordset
'   Dim lngRecordCount As Long
'
'    ' Open Recordset for mvarOldRecordSource
'    lngPosition = rstOldRecordSource.AbsolutePosition
'
'    'Set mvarOldRecordSource = rstOldRecordSource
'
'    Call RstCopy(rstOldRecordSource, mvarOldRecordSource, True, Start, Count)
'
'    mvarOldRecordSource.MoveFirst
'    For intIndex = 0 To mvarOldRecordSource.RecordCount - 1
'         If mvarOldRecordSource!Tag.Value = "D" Then
'            mvarOldRecordSource.Delete
'         End If
'         mvarOldRecordSource.MoveNext
'    Next intIndex
'    If mvarOldRecordSource.EOF Then
'      mvarOldRecordSource.MoveLast
'   Else
'      mvarOldRecordSource.AbsolutePosition = lngPosition
'   End If
'
'End Sub


Private Sub CheckIfAdd(ByVal Operation As RecordOperation)
    
    Dim fldDummy As ADODB.Field
    If mvarRecordSource.RecordCount > 0 And Operation = cpiRecordAdd Then
        mvarRecordSource.MoveFirst
                
        For Each fldDummy In mvarRecordSource.Fields
            'mvarADOFields.Add fldDummy
                
            If Operation = cpiRecordAdd Then
                Select Case fldDummy.Type
                    ' adSmallInt = 2
                    ' adInteger = 3
                    ' adSingle = 4
                    ' adDouble = 5
                    ' adCurrency = 6
                    ' adIDispatch = 9
                    ' adDecimal = 14
                    ' adTinyInt= 16
                    ' adUnsignedTinyInt = 17
                    ' adUnsignedSmallInt = 18
                    ' adUnsignedInt = 19
                    ' adBigInt = 20
                    ' adUnsignedBigInt = 21
                    ' adBinary = 128
                    ' adNumeric = 131
                    ' adChapter = 136
                    ' adPropVariant= 138
                    ' adLongVarBinary = 205
                    Case adUnsignedInt, adUnsignedSmallInt, adInteger, adUnsignedTinyInt, _
                            adSingle, adDouble, adCurrency, adNumeric, adDecimal, _
                            adLongVarBinary, adIDispatch, adPropVariant, adChapter, _
                            adBinary, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt
                            
                        mvarRecordSource.Fields(fldDummy.Name) = 0
    
                    Case adDate, adDBDate      ' Date
                        ' Do nothing
                    
                    ' adDBFileTime = 137
                    Case adDBTime, adFileTime, 137, adDBTimeStamp  ' Time
                        ' Do nothing
                        
                    ' dbBoolean = 1
                    ' adBoolean = 11
                    Case adBoolean, 1   ' Boolean

                        ' Do nothing
                    
                    ' adBSTR = 8                                            ' String
                    ' dbText = 10
                    ' dbMemo = 12
                    ' adChar = 129
                    ' adWChar = 130
                    ' adVarChar = 200
                    ' adLongVarChar = 201
                    ' adVarWChar = 202
                    ' adLongVarWChar= 203
                    ' adVarBinary = 204
                    Case adChar, adLongVarChar, adVarBinary, adVarChar, _
                            adVarWChar, adWChar, 10, 12, adBSTR
                        
                        mvarRecordSource.Fields(fldDummy.Name) = ""
                        
                    Case 205            ' OLE Object
                        ' Do nothing
                    Case Else
                        ' Do nothing
                End Select
            End If
        Next
    End If

End Sub
