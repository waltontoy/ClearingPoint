VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Segments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Segment"
Attribute VB_Ext_KEY = "Member0" ,"Segment"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E87F5CE0316"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mrstTechnicalMessageStructure As ADODB.Recordset

Friend Property Set TechnicalMessageStructure(ByRef TMSRecordset As ADODB.Recordset)
    Set mrstTechnicalMessageStructure = TMSRecordset
End Property

'Public Function Add(Key As String, NCTS_IEM_TMS_ID As Long, SegmentTag As String, Description As String, EDIUsage As String, EDIOccurrence As Long, NCTSUsage As String, NCTSOccurrence As Long) As Segment
Public Function Add(NCTS_IEM_TMS_ID As Long, SegmentInstance As Long, ParentInstance As Long) As Segment
    Dim objNewMember As Segment
    
    Dim lngSegmentEDIOccurrence     As Long
    Dim lngSegmentNCTSOccurrence    As Long
    Dim lngSegmentParentID          As Long
    Dim lngEDIOrdinal               As Long
    Dim lngNCTSCodelistID           As Long
    Dim lngEDI_TMS_ITM_ID           As Long
    Dim lngNCTS_ITM_ID              As Long
    Dim lngEDIGroupID               As Long
    
    Dim strSegmentTag               As String
    Dim strSegmentDescription       As String
    Dim strSegmentEDIUsage          As String
    Dim strSegmentNCTSUsage         As String
    Dim strSegmentKey               As String
    Dim strSegmentKeyParent         As String
    Dim strEDIName                  As String
    Dim strEDIDataFormat            As String
    Dim strEDIUsage                 As String
    Dim strNCTSName                 As String
    Dim strNCTSDataFormat           As String
    Dim strNCTSUsage                As String
    Dim strNCTSValue                As String
    
    Dim colSDataItems               As SDataItems
    Dim rstNCTSItems                As ADODB.Recordset
    Dim rstEDIItems                 As ADODB.Recordset
    
    Set objNewMember = New Segment
    
    If IsInRecordset(mrstTechnicalMessageStructure, "NCTS_IEM_TMS_ID = " & CStr(NCTS_IEM_TMS_ID)) Then
        strSegmentTag = mrstTechnicalMessageStructure.Fields("EDI_TMS_SEG_Tag").Value
        strSegmentDescription = mrstTechnicalMessageStructure.Fields("EDI_TMS_SEG_Description").Value
        strSegmentEDIUsage = mrstTechnicalMessageStructure.Fields("EDI_TMS_Usage").Value
        lngSegmentEDIOccurrence = mrstTechnicalMessageStructure.Fields("EDI_TMS_Occurrence").Value
        strSegmentNCTSUsage = mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_Usage").Value
        lngSegmentNCTSOccurrence = mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_Occurrence").Value
        lngSegmentParentID = mrstTechnicalMessageStructure.Fields("NCTS_IEM_TMS_ParentID").Value
    Else
        Debug.Assert Not (mrstTechnicalMessageStructure.BOF Or mrstTechnicalMessageStructure.EOF)
    End If
    
    'set the properties passed into the method
    objNewMember.SegmentTag = strSegmentTag
    objNewMember.Description = strSegmentDescription
    objNewMember.EDIUsage = strSegmentEDIUsage
    objNewMember.EDIOccurrence = lngSegmentEDIOccurrence
    objNewMember.NCTSUsage = strSegmentNCTSUsage
    objNewMember.NCTSOccurrence = lngSegmentNCTSOccurrence
    objNewMember.DataNCTSSegmentID = 0
    objNewMember.IsSaved = False
    
    Set colSDataItems = objNewMember.SDataItems
    Set rstNCTSItems = mrstTechnicalMessageStructure.Fields(strSegmentTag).Value
    Set rstEDIItems = mrstTechnicalMessageStructure.Fields("EDI_TMS_ITEMS").Value
    Do While Not rstEDIItems.EOF
        If IsNull(rstEDIItems.Fields("EDI_TMS_GRP_ID").Value) Then
            lngEDIGroupID = 0
        Else
            lngEDIGroupID = rstEDIItems.Fields("EDI_TMS_GRP_ID").Value
        End If
        If IsNull(rstEDIItems.Fields("EDI_TMS_ITM_Ordinal").Value) Then
            lngEDIOrdinal = 0
        Else
            lngEDIOrdinal = rstEDIItems.Fields("EDI_TMS_ITM_Ordinal").Value
        End If
        If IsNull(rstEDIItems.Fields("EDI_TMS_ITM_ID").Value) Then
            lngEDI_TMS_ITM_ID = 0
        Else
            lngEDI_TMS_ITM_ID = rstEDIItems.Fields("EDI_TMS_ITM_ID").Value
        End If
        If IsNull(rstEDIItems.Fields("EDI_TMS_ITM_Description").Value) Then
            strEDIName = vbNullString
        Else
            strEDIName = rstEDIItems.Fields("EDI_TMS_ITM_Description").Value
        End If
        If IsNull(rstEDIItems.Fields("EDI_TMS_ITM_DataType").Value) Then
            strEDIDataFormat = vbNullString
        Else
            strEDIDataFormat = rstEDIItems.Fields("EDI_TMS_ITM_DataType").Value
        End If
        If IsNull(rstEDIItems.Fields("EDI_TMS_ITM_Usage").Value) Then
            strEDIUsage = vbNullString
        Else
            strEDIUsage = rstEDIItems.Fields("EDI_TMS_ITM_Usage").Value
        End If
        If IsInRecordset(rstNCTSItems, "EDI_TMS_ITM_ID = " & CStr(lngEDI_TMS_ITM_ID)) Then
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Codelist").Value) Then
                lngNCTSCodelistID = 0
            Else
                ' MUCP-51 - Start
                Dim strTempCodelist As String
                strTempCodelist = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Codelist").Value
                If IsNumeric(strTempCodelist) Then
                    lngNCTSCodelistID = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Codelist").Value
                Else
                    lngNCTSCodelistID = 0
                End If
                'lngNCTSCodelistID = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Codelist").Value
                ' MUCP-51 - End
            End If
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_ID").Value) Then
                lngNCTS_ITM_ID = 0
            Else
                lngNCTS_ITM_ID = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_ID").Value
            End If
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Description").Value) Then
                strNCTSName = vbNullString
            Else
                strNCTSName = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Description").Value
            End If
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_DataType").Value) Then
                strNCTSDataFormat = vbNullString
            Else
                strNCTSDataFormat = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_DataType").Value
            End If
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Usage").Value) Then
                strNCTSUsage = vbNullString
            Else
                strNCTSUsage = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Usage").Value
            End If
            If IsNull(rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Value").Value) Then
                strNCTSValue = vbNullString
            Else
                strNCTSValue = rstNCTSItems.Fields("NCTS_ITM_" & strSegmentTag & "_Value").Value
            End If
        Else
            lngNCTSCodelistID = 0
            lngNCTS_ITM_ID = 0
            strNCTSName = vbNullString
            strNCTSDataFormat = vbNullString
            strNCTSUsage = vbNullString
            strNCTSValue = vbNullString
        End If
        colSDataItems.Add "S" & CStr(lngEDIOrdinal), lngNCTS_ITM_ID, strEDIName, strEDIDataFormat, strEDIUsage, lngEDIOrdinal, strNCTSName, strNCTSDataFormat, strNCTSUsage, lngNCTSCodelistID, lngEDIGroupID, strNCTSValue
        rstEDIItems.MoveNext
    Loop
    
    strSegmentKey = "S_" & CStr(NCTS_IEM_TMS_ID) & "_" & CStr(SegmentInstance)
    'strSegmentKeyParent = vbNullString
    If lngSegmentParentID > 0 Then
        strSegmentKeyParent = "S_" & CStr(lngSegmentParentID) & "_" & CStr(ParentInstance)
'''''        If Me(strSegmentKeyParent).KeysChildren = vbNullString Then
'''''            Me(strSegmentKeyParent).KeysChildren = strSegmentKey
'''''        Else
'''''            Me(strSegmentKeyParent).KeysChildren = Me(strSegmentKeyParent).KeysChildren & "|" & strSegmentKey
'''''        End If
    Else
        strSegmentKeyParent = vbNullString
    End If
    
    objNewMember.Key = strSegmentKey
    objNewMember.KeyParent = strSegmentKeyParent
    objNewMember.KeysChildren = vbNullString
    
    mCol.Add objNewMember, strSegmentKey
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Segment
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
