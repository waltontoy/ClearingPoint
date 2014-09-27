VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTemplateGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'''''' Function #1
'''''' Function Name: GetRecordset
'''''' Description: Returns created recordset based on query and connection passed.
'''''' Syntax: Set [rstObject] = GetRecordset([conObject],[strSQL])
'''''' Scope: Public
'''''' Fan-In: <none>
'''''Public Function GetRecordset(ByRef ActiveConnection As ADODB.Connection, _
'''''                             ByRef CommandText As String) As ADODB.Recordset
'''''
'''''    Dim rstRecordset As ADODB.Recordset
'''''
'''''    Set rstRecordset = New ADODB.Recordset
'''''
'''''    On Error GoTo ERROR_RECORDSET
'''''    rstRecordset.Open CommandText, ActiveConnection, adOpenKeyset, adLockOptimistic
'''''
'''''    Set GetRecordset = rstRecordset
'''''    Set rstRecordset = Nothing
'''''    Exit Function
'''''
'''''ERROR_RECORDSET:
'''''    Set rstRecordset = Nothing
'''''End Function

' Function #4
' Function Name: GetChunkData
' Description: get next 100 bytes of current  binary values from database field
' Syntax: [blnFlag] = GetChunkData([fldObject],[varData])
' Scope: Public
' Fan-In: G_GetChunkData
Public Function GetChunkData(ByRef ActiveField As ADODB.Field, _
                             ByRef varData As Variant) As Boolean
    GetChunkData = G_GetChunkData(ActiveField, varData)
End Function

' Function #5
' Function Name: GetOLEFromDB
' Description: get recordset with binary data (not yet implemented)
' Syntax: [blnFlag] = GetOLEFromDB([conObject],[strSQL],[varData])
' Scope: Public
' Fan-In: G_GetOLEFromDB
Public Function GetOLEFromDB(ByRef ActiveConnection As ADODB.Connection, _
                             ByRef CommandText As String, _
                             ByRef varData As Variant) As Boolean
    GetOLEFromDB = G_GetOLEFromDB(ActiveConnection, CommandText, varData)
End Function

' Function #6
' Function Name: SetEnclosedChr
' Description: set proper enclosed character for each data type
' Syntax: [varData] = SetEnclosedChr([varDataType])
' Scope: Public
' Fan-In: G_SetEnclosedChr
Public Function SetEnclosedChr(ByRef varValue As Variant) As Variant
    SetEnclosedChr = G_SetEnclosedChr(varValue)
End Function

' Function #7
' Function Name: Translate_B
' Description: get corresponding language value based on ID passed
' Syntax: [varValue] = Translate_B([varData],[lngResourceHandler])
' Scope: Public
' Fan-In: Translate
Public Function Translate_B(ByVal StringToTranslate As Variant, _
                            ByVal ResourceHandler As Long) As String
    Translate_B = Translate(StringToTranslate, ResourceHandler)
End Function

'''''' Function #8
'''''' Function Name: SetRecordset
'''''' Description: load static recordset from target database
'''''' Syntax: [blnData] = SetRecordset([conObject],[strSQL],[rstObject])
'''''' Scope: Public
'''''' Fan-In: <none>
'''''Public Function SetRecordset(ByRef ActiveConnection As ADODB.Connection, _
'''''                             ByRef CommandText As String, _
'''''                             ByRef ActiveRecord As ADODB.Recordset) As Boolean
'''''
'''''    Dim lngRecordsAffected As Long
'''''
'''''    On Error GoTo ERROR_SEARCH
'''''    Set ActiveRecord = ActiveConnection.Execute(CommandText, lngRecordsAffected)
'''''
'''''    SetRecordset = True
'''''    Exit Function
'''''
'''''ERROR_SEARCH:
'''''    SetRecordset = False
'''''End Function

Public Function ADORecordsetCopy(ByRef Source As ADODB.Recordset) As ADODB.Recordset
    Dim rstRecordsetCopy As ADODB.Recordset
    Dim pb As PropertyBag
    
    Dim strFilter As String
    Dim strSort As String
    Dim lngAbsolutePosition As Long
    Dim strDateNow As String
    
    strDateNow = Format(Now, "_" & "mmddyyy" & "_" & "hhmmss")
    
    lngAbsolutePosition = Source.AbsolutePosition
    
    strFilter = Source.Filter
    strSort = Source.Sort
    
    ' Create a copy of the filtered recordset
    Set pb = New PropertyBag
    pb.WriteProperty "rstCopy" & strDateNow, Source
    
    Set rstRecordsetCopy = New ADODB.Recordset
    Set rstRecordsetCopy = pb.ReadProperty("rstCopy" & strDateNow)
    
    Set pb = Nothing
    
    rstRecordsetCopy.Filter = strFilter
    rstRecordsetCopy.Sort = strSort
    rstRecordsetCopy.AbsolutePosition = lngAbsolutePosition
    
    Set ADORecordsetCopy = rstRecordsetCopy
    
    Set rstRecordsetCopy = Nothing

End Function

' Function #9
' Function Name: RstCopy
' Description: returns a copy of recordset properties and records
' Syntax: Set [varRecordset] = RstCopy([rstSource],[blnDisconnected],
'              [lngRecordStart],[lngRecordEnd],[lngAbsolutePosition],[blnFieldOnly])
' Scope: Public
' Fan-In: <none>
Public Function RstCopy(ByRef Source As ADODB.Recordset, _
                        ByRef Disconnected As Boolean, _
                        ByRef RecordStart As Long, _
                        ByRef RecordEnd As Long, _
                        ByRef AbsolutePosition As Long, _
                        Optional ByVal FieldOnly As Boolean) As Variant

    Dim rstTemp As ADODB.Recordset
    Dim intIndex As Long
    Dim intPropertyIdx As Long
    Dim fldTemp As ADODB.Field
    Dim lngSourcePos As Long
   
    If (Disconnected = False) Then
        Set RstCopy = Source
        Exit Function
    End If
   
    Set rstTemp = New ADODB.Recordset
    
    ' STEP -1 copy all the fields properties
    For intIndex = 0 To Source.Fields.Count - 1
        rstTemp.Fields.Append Source.Fields(intIndex).Name, _
                              Source.Fields(intIndex).Type, _
                              Source.Fields(intIndex).DefinedSize, _
                              Source.Fields(intIndex).Attributes
    Next intIndex
      
    rstTemp.Open
   
    If (FieldOnly = True) Then
        rstTemp.AddNew
        rstTemp.Update
        GoTo rstEnd
    End If
   
    If ((Source.RecordCount = 0) Or (Source.RecordCount = -1)) Then
        GoTo rstEnd
    End If
        
     lngSourcePos = Source.AbsolutePosition
     Source.MoveFirst
     Source.Move RecordStart '+ 1

     ' copy all the record values of the source
     For intIndex = RecordStart To RecordEnd
         rstTemp.AddNew
         
         For Each fldTemp In Source.Fields
             rstTemp.Fields(fldTemp.Name) = fldTemp.Value
         Next ' fldTemp
         Source.MoveNext
     Next intIndex

     rstTemp.Update

     rstTemp.AbsolutePosition = AbsolutePosition
     
     If Source.EOF Then
         Source.MoveLast
     End If

     Source.AbsolutePosition = lngSourcePos

rstEnd:
   Set RstCopy = rstTemp
   Set rstTemp = Nothing
End Function


