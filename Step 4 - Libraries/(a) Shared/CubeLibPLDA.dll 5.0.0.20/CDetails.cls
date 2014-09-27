VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function CreateDetails(ByRef lngSegmentCount As Long, _
                              ByRef RecordsHeader As ADODB.Recordset, _
                              ByRef RecordsDetails As ADODB.Recordset, _
                              ByRef RecordsDetailsBijzondere As ADODB.Recordset, _
                              ByRef RecordsDetailsContainer As ADODB.Recordset, _
                              ByRef RecordsDetailsDocumenten As ADODB.Recordset, _
                              ByRef RecordsDetailsZelf As ADODB.Recordset, _
                              ByRef RecordsDetailsBerekeningsEenheden As ADODB.Recordset, _
                              ByRef RecordsDetailsSensitiveGoods As ADODB.Recordset, _
                              ByVal lngDType As Long, _
                              ByRef ReturnTotalNumberCST As Long, _
                              Optional ByVal Code As String) As String
    
    Dim rstDetailsHandelaars As ADODB.Recordset
    
    Dim strMessage As String
    Dim CImpDetail As CIDetail
    Dim CExpDetail As CEDetail
    
    Dim lngDetailCnt As Long
    
    If RecordsDetails.RecordCount > 0 Then
        RecordsDetails.MoveFirst
    End If
    
    lngDetailCnt = 0
    strMessage = ""
    
    Do While RecordsDetails.EOF = False
        
        Set rstDetailsHandelaars = g_rstDetailsHandelaars.Clone
        rstDetailsHandelaars.Filter = adFilterNone
        rstDetailsHandelaars.Filter = "Detail = " & FNullField(RecordsDetails.Fields("Detail").Value)
        
        'Import Declarations
        If lngDType = 14 Then
            Set CImpDetail = New CIDetail
            
            strMessage = strMessage & CImpDetail.CreateDetail(RecordsHeader, _
                                                              RecordsDetails, _
                                                              RecordsDetailsBijzondere, _
                                                              RecordsDetailsContainer, _
                                                              RecordsDetailsDocumenten, _
                                                              RecordsDetailsZelf, _
                                                              rstDetailsHandelaars, _
                                                              RecordsDetailsBerekeningsEenheden, _
                                                              lngSegmentCount, _
                                                              FNullField(RecordsDetails.Fields("Detail").Value), _
                                                              enuOriginal)

            lngDetailCnt = lngDetailCnt + 1
            ReturnTotalNumberCST = lngDetailCnt
        
        'Export Declarations
        ElseIf lngDType = 18 Then
            Set CExpDetail = New CEDetail
            
            strMessage = strMessage & CExpDetail.CreateDetail(RecordsHeader, _
                                                              RecordsDetails, _
                                                              RecordsDetailsBijzondere, _
                                                              RecordsDetailsContainer, _
                                                              RecordsDetailsDocumenten, _
                                                              rstDetailsHandelaars, _
                                                              RecordsDetailsSensitiveGoods, _
                                                              lngSegmentCount, _
                                                              FNullField(RecordsDetails.Fields("Detail").Value), _
                                                              FNullField(RecordsHeader.Fields("A9").Value), _
                                                              enuOriginal, Code)

            lngDetailCnt = lngDetailCnt + 1
            ReturnTotalNumberCST = lngDetailCnt
        End If
        
        rstDetailsHandelaars.Filter = adFilterNone
        
        If UCase(FNullField(RecordsDetails.Fields("T7").Value)) = "F" Then
            Exit Do
        End If
        
        RecordsDetails.MoveNext
        
        Set CImpDetail = Nothing
        Set CExpDetail = Nothing
            
    Loop
    
    'Additional Digital Signature for Total number of Details
    'Based on total number of details code in CISummary
    'Number of Details
    AddFieldToDigiSign "Total Number of Details", RecordsDetails, False
    
    CreateDetails = strMessage
    
End Function
