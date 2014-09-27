Attribute VB_Name = "MUnloadingRemarks"
Option Explicit

Private strTemp As String

'p4tric 021908
Private Function GetTOD(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)

    Dim clsTODs As PCubeLibEDIDataTag.cpiDATA_NCTS_TODs
    Dim clsTOD As PCubeLibEDIDataTag.cpiDATA_NCTS_TOD

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsTODs = New cpiDATA_NCTS_TODs
    Set clsTOD = New cpiDATA_NCTS_TOD
    
    
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsTOD.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsTOD.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsTODs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        0, 1)

    'loop here
    Do While (clsTODs.Recordset.EOF = False)

        Set clsTOD = clsTODs.GetClassRecord(clsTODs.Recordset)
        clsMessage.UnloadingRemarks(1).FIELD_STATE_OF_SEALS_OK = clsTOD.FIELD_DATA_NCTS_TOD_Seq2
        clsMessage.UnloadingRemarks(1).FIELD_CONFORM = clsTOD.FIELD_DATA_NCTS_TOD_Seq3
        clsMessage.UnloadingRemarks(1).FIELD_UNLOADING_COMPLETION = clsTOD.FIELD_DATA_NCTS_TOD_Seq3
        clsMessage.UnloadingRemarks(1).FIELD_UNLOADING_DATE = clsTOD.FIELD_DATA_NCTS_TOD_Seq6
        clsTODs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsTODs = Nothing
    Set clsTOD = Nothing

End Function



Private Function GetBGM(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)
'
    Dim clsBGMs As PCubeLibEDIDataTag.cpiDATA_NCTS_BGMs
    Dim clsBGM As PCubeLibEDIDataTag.cpiDATA_NCTS_BGM

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsBGMs = New cpiDATA_NCTS_BGMs
    Set clsBGM = New cpiDATA_NCTS_BGM

'    ' delete all existing records with same message id
'    strSql = "DELETE * FROM [DATA_NCTS_" & strTagName & "] WHERE [DATA_NCTS_MSG_ID]=" & CStr(lngDATA_NCTS_MSG_ID)
'    clsBGMs.GetRecordset EdifactDB, strSql

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsBGM.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsBGM.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsBGMs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        0, 1)

    'loop here
    Do While (clsBGMs.Recordset.EOF = False)

        Set clsBGM = clsBGMs.GetClassRecord(clsBGMs.Recordset)
        clsMessage.Headers(1).MOVEMENT_REFERENCE_NUMBER = clsBGM.FIELD_DATA_NCTS_BGM_Seq5
        clsBGMs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsBGMs = Nothing
    Set clsBGM = Nothing
'
End Function

Private Function GetLOC_22(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)
'
    Dim clsLOCs As cpiDATA_NCTS_LOCs
    Dim clsLOC As cpiDATA_NCTS_LOC

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsLOCs = New cpiDATA_NCTS_LOCs
    Set clsLOC = New cpiDATA_NCTS_LOC

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsLOC.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsLOC.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsLOCs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)

    'loop here
    Do While (clsLOCs.Recordset.EOF = False)

        Set clsLOC = clsLOCs.GetClassRecord(clsLOCs.Recordset)

        ' Reference number (R, an8)
        clsMessage.CustomOffices(1).REFERENCE_NUMBER = clsLOC.FIELD_DATA_NCTS_LOC_Seq2

        clsLOCs.Recordset.MoveNext

    Loop


    ' map values here

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsLOCs = Nothing
    Set clsLOC = Nothing

'
End Function

Private Function GetMEA_AAD(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)
'
    Dim clsMEAs As cpiDATA_NCTS_MEAs
    Dim clsMEA As cpiDATA_NCTS_MEA

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsMEAs = New cpiDATA_NCTS_MEAs
    Set clsMEA = New cpiDATA_NCTS_MEA

    ' ---
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsMEA.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsMEA.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' convert class to recordset
    'Set clsMEAs.Recordset = clsMEAs.GetTableRecord(EdifactDB, clsMEA)

    ' MAP - start
    ' map message and clsMEA here
    ' get tag recordset
    Set clsMEAs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)

    'loop here
    Do While (clsMEAs.Recordset.EOF = False)

        Set clsMEA = clsMEAs.GetClassRecord(clsMEAs.Recordset)

        ' Total gross mass (O, n..11,3)
        clsMessage.Headers(1).TOTAL_GROSS_MASS = clsMEA.FIELD_DATA_NCTS_MEA_Seq7

        clsMEAs.Recordset.MoveNext

    Loop


    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsMEAs = Nothing
    Set clsMEA = Nothing

'
End Function

Private Function GetSEL(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44) As Boolean
'
    Dim clsSELs As cpiDATA_NCTS_SELs
    Dim clsSEL As cpiDATA_NCTS_SEL

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim intSealCtr As Integer
    Dim blnLoadTrueSeals As Boolean
    Set clsSELs = New cpiDATA_NCTS_SELs
    Set clsSEL = New cpiDATA_NCTS_SEL

    ' seq_1
    ' seq_5 Seals identity (R, an..20)
    ' ---
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsSEL.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsSEL.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID
    
    ' convert class to recordset
    'Set clsSELs.Recordset = clsSELs.GetTableRecord(EdifactDB, clsSEL)

    ' MAP - start
    ' map message and clsSEL_here

    ' get tag recordset
    Set clsSELs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0)

    ' loop here
    
    Do While (clsSELs.Recordset.EOF = False)
        blnLoadTrueSeals = True
        Set clsSEL = clsSELs.GetClassRecord(clsSELs.Recordset)

        intSealCtr = intSealCtr + 1
        clsMessage.SealInfos(1).Seals.Add CStr(clsSEL.FIELD_DATA_NCTS_SEL_Instance), clsMasterRecord.FIELD_CODE, 0

        ' seq 1 - Seals identity LNG (R, a2)
        ' clsSEL.FIELD_DATA_NCTS_SEL_Seq1 = clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION

        ' seq 5 - Seals identity (R, an..20)
        clsMessage.SealInfos(1).Seals(intSealCtr).SEALS_IDENTITY = clsSEL.FIELD_DATA_NCTS_SEL_Seq5

        clsSELs.Recordset.MoveNext

    Loop  ' @@@ break the BREAK ! ! !  :-)
    
    ' update seal's format here
    If blnLoadTrueSeals Then
        LoadTrueSeals clsMessage.SealInfos, clsMasterRecord
    End If
    
    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsSELs = Nothing
    Set clsSEL = Nothing

End Function

Private Function IE44_GetFTX_ABV(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44)

    Dim clsFTXs As cpiDATA_NCTS_FTXs
    Dim clsFTX As cpiDATA_NCTS_FTX

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strDescription As String

    Dim intResultOfControlCtr As Integer

    ' strDescription
    Set clsFTXs = New cpiDATA_NCTS_FTXs
    Set clsFTX = New cpiDATA_NCTS_FTX

    ' seq 3,6,7,8,9,11
    ' --- >>>
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsFTX.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsFTX.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsFTXs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0)

    intResultOfControlCtr = 0

    Set clsMessage.ResultOfControls = Nothing
    Set clsMessage.ResultOfControls = New cpiResultOfControls

    ' loop here
    Do While (clsFTXs.Recordset.EOF = False)

        Set clsFTX = clsFTXs.GetClassRecord(clsFTXs.Recordset)

        intResultOfControlCtr = intResultOfControlCtr + 1 '= clsFTX.FIELD_DATA_NCTS_FTX_Instance
        clsMessage.ResultOfControls.Add CStr(intResultOfControlCtr), clsMasterRecord.FIELD_CODE   ',  0

        ' seq 3 - Control indicator (R, an2)
        clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_CONTROL_INDICATOR = clsFTX.FIELD_DATA_NCTS_FTX_Seq3

        ' seq 6 - Pointer to the attribute (D, an..35)
        clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_POINTER_TO_THE_ATTRIBUTE = clsFTX.FIELD_DATA_NCTS_FTX_Seq6

        ' strDescription = clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_DESCRIPTION
        ' seq 7,8 - Description (O, an..140) 1..70
        strDescription = clsFTX.FIELD_DATA_NCTS_FTX_Seq7
        strDescription = strDescription & clsFTX.FIELD_DATA_NCTS_FTX_Seq8
        clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_DESCRIPTION = strDescription

        ' seq 9 - Corrected value (D, an..15)
        clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_CORRECTED_VALUE = clsFTX.FIELD_DATA_NCTS_FTX_Seq9

        
        clsFTXs.Recordset.MoveNext

    Loop  ' @@@ break the BREAK ! ! !  :-)

'    ' MAP - start
'    ' map message and clsFTX here
'
'    ' seq 3 - Control indicator (R, an2)
'    clsFTX.FIELD_DATA_NCTS_FTX_Seq3 = clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_CONTROL_INDICATOR
'
'    ' seq 6 - Pointer to the attribute (D, an..35)
'    clsFTX.FIELD_DATA_NCTS_FTX_Seq6 = clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_POINTER_TO_THE_ATTRIBUTE
'
'    strDescription = clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_DESCRIPTION
'
'    If (Len(strDescription) <= 70) Then
'        ' seq 7 - Description (O, an..140) 1..70
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq7 = strDescription
'    ElseIf (Len(strDescription) <= 140) Then
'        ' seq 7 - Description (O, an..140) 1..70
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq7 = Left$(strDescription, 70)
'        ' seq 8 - Description (O, an..140) 71..140
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq8 = Mid$(strDescription, 71)
'    End If
'
'    ' seq 9 - Corrected value (D, an..15)
'    clsFTX.FIELD_DATA_NCTS_FTX_Seq9 = clsMessage.ResultOfControls(intResultOfControlCtr).FIELD_CORRECTED_VALUE
'
'    ' seq 11 - Description LNG (D, a2)
'    clsFTX.FIELD_DATA_NCTS_FTX_Seq11 = clsMessage.Headers(intResultOfControlCtr).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION
'
'    ' set instance here
'    clsFTX.FIELD_DATA_NCTS_FTX_Instance = 1
'
'    ' map values here
'    ' add pk +1
'    clsFTXs.GetMaxID EdifactDB, clsFTX
'    clsFTX.FIELD_DATA_NCTS_FTX_ID = clsFTX.FIELD_DATA_NCTS_FTX_ID + 1
'
'    ' add record to db
'    clsFTXs.AddRecord EdifactDB, clsFTX
'
'    'clsmessage.Headers(intResultOfControlCtr).
'    ' set instance count

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsFTXs = Nothing
    Set clsFTX = Nothing

'
End Function


Private Function GetTDT_12(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)

    Dim clsTDTs As cpiDATA_NCTS_TDTs
    Dim clsTDT As cpiDATA_NCTS_TDT

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsTDTs = New cpiDATA_NCTS_TDTs
    Set clsTDT = New cpiDATA_NCTS_TDT

    ' seq 18,19
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsTDT.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsTDT.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start
    ' map message and clsTDT here
    ' get tag recordset
    Set clsTDTs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)



    ' loop here
    Do While (clsTDTs.Recordset.EOF = False)

        Set clsTDT = clsTDTs.GetClassRecord(clsTDTs.Recordset)

        ' Total gross mass (O, n..11,3)
        ' seq 18 - Identity of MEANS of transport at departure (exp/trans) (O, an..27)
        clsMessage.Headers(1).IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE = clsTDT.FIELD_DATA_NCTS_TDT_Seq18

        ' seq 19 - Nationality of MEANS of transport at departure (O, a2) [9]
        clsMessage.Headers(1).NATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE = clsTDT.FIELD_DATA_NCTS_TDT_Seq19

        ' add children here
        'IE44_GetTPL udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "TPL", 19, MSG_IE44, "", _
                    clsMasterRecord, clsTDT.FIELD_DATA_NCTS_TDT_ID


        clsTDTs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsTDTs = Nothing
    Set clsTDT = Nothing

'
End Function


Private Function GetNAD_CPD(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)
'
    Dim clsNADs As cpiDATA_NCTS_NADs
    Dim clsNAD As cpiDATA_NCTS_NAD

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Set clsNADs = New cpiDATA_NCTS_NADs
    Set clsNAD = New cpiDATA_NCTS_NAD

    ' seq 2, 10, 16, 20-23
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsNAD.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsNAD.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start
    ' map message and clsNAD here
    ' get tag recordset
    Set clsNADs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)

    ' loop here
    Do While (clsNADs.Recordset.EOF = False)

        Set clsNAD = clsNADs.GetClassRecord(clsNADs.Recordset)

        ' TIN
        clsMessage.Traders(1).DESTINATION_TIN = clsNAD.FIELD_DATA_NCTS_NAD_Seq2

        ' Name
        clsMessage.Traders(1).DESTINATION_NAME = clsNAD.FIELD_DATA_NCTS_NAD_Seq10

        ' Street and number
        clsMessage.Traders(1).DESTINATION_STREET_AND_NUMBER = clsNAD.FIELD_DATA_NCTS_NAD_Seq16

        ' City
        clsMessage.Traders(1).DESTINATION_CITY = clsNAD.FIELD_DATA_NCTS_NAD_Seq20

        ' NAD_LNG
        clsMessage.Traders(1).DESTINATION_NAD_LNG = clsNAD.FIELD_DATA_NCTS_NAD_Seq21

        ' Postal code
        clsMessage.Traders(1).DESTINATION_POSTAL_CODE = clsNAD.FIELD_DATA_NCTS_NAD_Seq22

        ' Country
        clsMessage.Traders(1).DESTINATION_COUNTRY_CODE = clsNAD.FIELD_DATA_NCTS_NAD_Seq23
'
'        Exit Do
        clsNADs.Recordset.MoveNext

    Loop


    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsNADs = Nothing
    Set clsNAD = Nothing

'
End Function

Private Function GetCST(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef NCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByVal NewCode As Boolean)

    Dim clsCSTs As cpiDATA_NCTS_CSTs
    Dim clsCST As cpiDATA_NCTS_CST

    Dim intSealsTotal As Integer
    Dim intContainerTotal As Integer

    Dim intResultOfControlTotal As Integer
    Dim intPackagesTotal As Integer
    Dim intDocCertTotal As Integer
    Dim intSGICodeTotal As Integer

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Dim intGoodsItemCtr As Integer

    Set clsCSTs = New cpiDATA_NCTS_CSTs
    Set clsCST = New cpiDATA_NCTS_CST

    ' get total result of control's count = intResultOfControlTotal
'    intResultOfControlTotal = 0
'    For intGoodsItemCtr = 1 To clsMessage.GoodsItems.Count
'        intResultOfControlTotal = intResultOfControlTotal + clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls.Count
'    Next intGoodsItemCtr

    ' get total package count = intPackagesTotal
    intPackagesTotal = 0
'    For intGoodsItemCtr = 1 To clsMessage.GoodsItems.Count
'        intPackagesTotal = intPackagesTotal + clsMessage.GoodsItems(intGoodsItemCtr).Packages.Count
'    Next intGoodsItemCtr

    ' get total container count
    intContainerTotal = 0
    'For intGoodsItemCtr = 1 To clsMessage.GoodsItems.Count
    '    intContainerTotal = intContainerTotal + clsMessage.GoodsItems(intGoodsItemCtr).Containers.Count
    'Next intGoodsItemCtr

    ' get total document/certificates count
    intDocCertTotal = 0
    'For intGoodsItemCtr = 1 To clsMessage.GoodsItems.Count
    '    intDocCertTotal = intDocCertTotal + clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates.Count
    'Next intGoodsItemCtr

    'intSgiCodeTotal
    ' get SGI codes total
    intSGICodeTotal = 0
    'For intGoodsItemCtr = 1 To clsMessage.GoodsItems.Count
    '    intSgiCodeTotal = intSgiCodeTotal + clsMessage.GoodsItems(intGoodsItemCtr).SGICodes.Count
    'Next intGoodsItemCtr


    ' seq 1, 2
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, NCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsCST.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsCST.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsCSTs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0)

    If Not NewCode Then
        If (clsCSTs.Recordset.EOF = True) Then
            ' get IE43 CST here
            Dim lngDATA_NCTS_MSG_ID_IE43 As Long
            
            'p4tric commento
            lngDATA_NCTS_MSG_ID_IE43 = GetNCTS_MSG_ID_IE43(EdifactDB, clsMasterRecord, clsMessage)
            
            'GetCST EdifactDB, clsMasterRecord, clsMessage, lngDATA_NCTS_MSG_ID_IE43, "CST", 93, lngDATA_NCTS_MSG_ID_IE43, "", NewCode    ', ActiveBar
            
            Exit Function
        End If
    End If

    ' loop here
    Do While (clsCSTs.Recordset.EOF = False)

        Set clsCST = clsCSTs.GetClassRecord(clsCSTs.Recordset)

        intGoodsItemCtr = intGoodsItemCtr + 1
        clsMessage.GoodsItems.Add CStr(clsCST.FIELD_DATA_NCTS_CST_Instance), clsMasterRecord.FIELD_CODE

        ' Item number (R, n..5)
        ' clsMessage.GoodsItems(intGoodsItemCtr).FIELD_ITEM_NUMBER = intGoodsItemCtr
        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_ITEM_NUMBER = clsCST.FIELD_DATA_NCTS_CST_Seq1
        ' Commodity code (taric code) (O, n..10)
        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_COMMODITY_CODE = clsCST.FIELD_DATA_NCTS_CST_Seq2

        ' Total Gross Mass
        GetFTX_AAA EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "FTX", 94, NCTS_IEM_ID, "AAA", _
                clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr

        'If Not NewCode Then
            ' Result of  Control
            IE44_GetFTX_ABV_Detail EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "FTX", 94, NCTS_IEM_ID, "ABV", _
                    clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intResultOfControlTotal
        'End If

        If Not NewCode Then
            ' Gross Mass (O, n..11,3)
            GetMEA_AAB EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 97, NCTS_IEM_ID, "WT:AAB:KGM", _
                    clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr
    
            
            ' Net mass
            GetMEA_AAA EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 97, NCTS_IEM_ID, "WT:AAA:KGM", _
                    clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr
        Else
            ' Gross Mass (O, n..11,3)
            GetMEA_AAB EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 97, NCTS_IEM_ID, "WT:AAB:KGM", _
                    clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr
    
            
            ' Net mass
            GetMEA_AAA EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 97, NCTS_IEM_ID, "WT:AAA:KGM", _
                    clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr
        End If
        
        ' Packages
        GetPAC_6 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "PAC", 100, NCTS_IEM_ID, "6", _
                clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intPackagesTotal

        If Not NewCode Then
            ' Containers
            IE44_GetRFF_AAQ EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "RFF", 106, NCTS_IEM_ID, "AAQ", _
                    clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intContainerTotal
        Else
            IE43_GetRFF_AAQ EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "RFF", 106, NCTS_IEM_ID, "AAQ", _
                    clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intContainerTotal
        End If

        ' Produce Documents / Certificates   ' intDocCertTotal
        GetDOC_916 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "DOC", 112, NCTS_IEM_ID, "916", _
                clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intDocCertTotal, NewCode

        ' SGI Codes
        If Not NewCode Then
            GetGIR_3 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "GIR", 132, NCTS_IEM_ID, "3", _
                    clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intSGICodeTotal
        Else
            GetGIR_3 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "GIR", 132, NCTS_IEM_ID, "3", _
                    clsMasterRecord, clsCST.FIELD_DATA_NCTS_CST_ID, intGoodsItemCtr, intSGICodeTotal
        End If

        clsCSTs.Recordset.MoveNext

    Loop

    If Not NewCode Then
        GetOtherNonSegmentBox_UnloadingRemarks EdifactDB, clsMessage
    End If



    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsCSTs = Nothing
    Set clsCST = Nothing

End Function

Private Function GetCNT_5(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)
    '
    Dim clsCNTs As cpiDATA_NCTS_CNTs
    Dim clsCNT As cpiDATA_NCTS_CNT

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table
    Dim intCNTCertCtr As Integer

    Dim strSQL As String

    Set clsCNTs = New cpiDATA_NCTS_CNTs
    Set clsCNT = New cpiDATA_NCTS_CNT
    '
    ' seq 4, 5, 6, 7, 8
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)
    '
    clsCNT.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsCNT.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID
    '
    ' MAP-start    ' get tag recordset
    Set clsCNTs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, 0, 1)

'    ' map message and clsCNT here
     ' loop here i love u jesus
    Do While (clsCNTs.Recordset.EOF = False)

        ' i'm dying
        Set clsCNT = clsCNTs.GetClassRecord(clsCNTs.Recordset)

        clsMessage.Headers(1).TOTAL_NUMBER_OF_ITEMS = clsCNT.FIELD_DATA_NCTS_CNT_Seq2

        clsCNTs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsCNTs = Nothing
    Set clsCNT = Nothing

End Function


Private Function GetCNT_11(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)

    Dim clsCNTs As cpiDATA_NCTS_CNTs
    Dim clsCNT As cpiDATA_NCTS_CNT

    Dim intSealsTotal As Integer
    Dim intContainerTotal As Integer

    Dim intResultOfControlTotal As Integer
    Dim intPackagesTotal As Integer
    Dim intDocCertTotal As Integer
    Dim intSGICodeTotal As Integer

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Dim intGoodsItemCtr As Integer

    Set clsCNTs = New cpiDATA_NCTS_CNTs
    Set clsCNT = New cpiDATA_NCTS_CNT

    ' seq 1, 2
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsCNT.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsCNT.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsCNTs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)

    ' loop here
    Do While (clsCNTs.Recordset.EOF = False)

        Set clsCNT = clsCNTs.GetClassRecord(clsCNTs.Recordset)

        clsMessage.Headers(1).TOTAL_NUMBER_OF_PACKAGES = clsCNT.FIELD_DATA_NCTS_CNT_Seq2

        clsCNTs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsCNTs = Nothing
    Set clsCNT = Nothing

End Function


Private Function GetCNT_16(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String)

    Dim clsCNTs As cpiDATA_NCTS_CNTs
    Dim clsCNT As cpiDATA_NCTS_CNT

    Dim intSealsTotal As Integer
    Dim intContainerTotal As Integer

    Dim intResultOfControlTotal As Integer
    Dim intPackagesTotal As Integer
    Dim intDocCertTotal As Integer
    Dim intSGICodeTotal As Integer

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Dim intGoodsItemCtr As Integer

    Set clsCNTs = New cpiDATA_NCTS_CNTs
    Set clsCNT = New cpiDATA_NCTS_CNT

    ' seq 1, 2
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsCNT.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsCNT.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsCNTs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, 0, 1)

    ' loop here
    Do While (clsCNTs.Recordset.EOF = False)

        Set clsCNT = clsCNTs.GetClassRecord(clsCNTs.Recordset)

        'clsMessage.GoodsItems.Add CStr(clsCNT.FIELD_DATA_NCTS_CNT_Instance), clsMasterRecord.FIELD_CODE

        'clsMessage.Headers(1).sel = clsCNT.FIELD_DATA_NCTS_CNT_Seq2

         ' ????
        ' Item number (R, n..5)
        'clsMessage.GoodsItems(intGoodsItemCtr).FIELD_ITEM_NUMBER = intGoodsItemCtr
        'clsMessage.GoodsItems(intGoodsItemCtr).FIELD_ITEM_NUMBER = clsCNT.FIELD_DATA_NCTS_CNT_Seq1
        ' Commodity code (taric code) (O, n..10)
        'clsMessage.GoodsItems(intGoodsItemCtr).FIELD_COMMODITY_CODE = clsCNT.FIELD_DATA_NCTS_CNT_Seq2
        clsMessage.SealInfos(1).FIELD_SEALS_NUMBER = clsCNT.FIELD_DATA_NCTS_CNT_Seq2

        clsCNTs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsCNTs = Nothing
    Set clsCNT = Nothing

End Function

Private Function GetNCTS_IEM_TMS_Table(ByRef ActiveConnection As ADODB.Connection, _
                            ByRef lngNCTS_IEM_ID As Long, _
                            ByRef lngEDI_TMS_ID As Long, _
                            ByRef strNCTS_IEM_TMS_RemarksQualifier As String) As PCubeLibEDIMap.cpiNCTS_IEM_TMS_Table
'
    Dim strSQL As String
    Dim clsNCTS_IEM_TMS_Tables As cpiNCTS_IEM_TMS_Tables
    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Set clsNCTS_IEM_TMS_Tables = New cpiNCTS_IEM_TMS_Tables

    If (strNCTS_IEM_TMS_RemarksQualifier = "") Then
        strSQL = "SELECT * FROM [NCTS_IEM_TMS] "
        strSQL = strSQL & " WHERE (([NCTS_IEM_ID]=" & CStr(lngNCTS_IEM_ID) & ")"
        strSQL = strSQL & " AND ([EDI_TMS_ID]=" & CStr(lngEDI_TMS_ID) & ")"
        strSQL = strSQL & " AND ([NCTS_IEM_TMS_RemarksQualifier]='" & strNCTS_IEM_TMS_RemarksQualifier & "'))"
        strSQL = strSQL & " OR (([NCTS_IEM_ID]=" & CStr(lngNCTS_IEM_ID) & ")"
        strSQL = strSQL & " AND ([EDI_TMS_ID]=" & CStr(lngEDI_TMS_ID) & ")"
        strSQL = strSQL & " AND ([NCTS_IEM_TMS_RemarksQualifier] IS NULL))"
    
    ElseIf (strNCTS_IEM_TMS_RemarksQualifier <> "") Then
        strSQL = "SELECT * FROM [NCTS_IEM_TMS] "
        strSQL = strSQL & " WHERE (([NCTS_IEM_ID]=" & CStr(lngNCTS_IEM_ID) & ")"
        strSQL = strSQL & " AND ([EDI_TMS_ID]=" & CStr(lngEDI_TMS_ID) & ")"
        strSQL = strSQL & " AND ([NCTS_IEM_TMS_RemarksQualifier]='" & strNCTS_IEM_TMS_RemarksQualifier & "'))"
    End If

    clsNCTS_IEM_TMS_Tables.GetRecordset ActiveConnection, strSQL

    If (clsNCTS_IEM_TMS_Tables.Recordset.EOF = False) Then
        Set clsNCTS_IEM_TMS_Table = clsNCTS_IEM_TMS_Tables.GetClassRecord(clsNCTS_IEM_TMS_Tables.Recordset)
    End If

    Set GetNCTS_IEM_TMS_Table = clsNCTS_IEM_TMS_Table

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsNCTS_IEM_TMS_Tables = Nothing
'
End Function


Private Function GetDATA_NCTS_TAG_Record(ByRef ActiveConnection As ADODB.Connection, _
                                            ByRef lngDATA_NCTS_MSG_ID As Long, _
                                            ByRef lngNCTS_IEM_TMS_ID As Long, _
                                            ByRef strTagName As String, _
                                            ByRef lngDATA_NCTS_TAG_ParentID As Long, _
                                            Optional ByVal lngDATA_NCTS_TAG_Instance As Long = 0) As ADODB.Recordset
    ' instance included
    Dim strSQL As String
    Dim rstDB As ADODB.Recordset

    If (lngDATA_NCTS_TAG_Instance > 0) Then

        strSQL = "SELECT * FROM [DATA_NCTS_" & strTagName & "] "
        strSQL = strSQL & " WHERE [DATA_NCTS_MSG_ID]=" & CStr(lngDATA_NCTS_MSG_ID)
        strSQL = strSQL & " AND [NCTS_IEM_TMS_ID]=" & CStr(lngNCTS_IEM_TMS_ID)
        strSQL = strSQL & " AND [DATA_NCTS_" & strTagName & "_ParentID]=" & CStr(lngDATA_NCTS_TAG_ParentID)
        strSQL = strSQL & " AND [DATA_NCTS_" & strTagName & "_Instance]=" & CStr(lngDATA_NCTS_TAG_Instance)

    ' allow multiple instance
    ElseIf ((lngDATA_NCTS_TAG_Instance) = 0) Then

        strSQL = "SELECT * FROM [DATA_NCTS_" & strTagName & "] "
        strSQL = strSQL & " WHERE [DATA_NCTS_MSG_ID]=" & CStr(lngDATA_NCTS_MSG_ID)
        strSQL = strSQL & " AND [NCTS_IEM_TMS_ID]=" & CStr(lngNCTS_IEM_TMS_ID)
        strSQL = strSQL & " AND [DATA_NCTS_" & strTagName & "_ParentID]=" & CStr(lngDATA_NCTS_TAG_ParentID)

    End If

    ' huhuhu :(
    ' FIELD_DATA_NCTS_BGM_Instance
    ' lngDATA_NCTS_TAG_ParentID
    
    ADORecordsetOpen strSQL, ActiveConnection, rstDB, adOpenKeyset, adLockOptimistic
    'Set rstDB = ActiveConnection.Execute(strSQL)

    If ((rstDB Is Nothing) = True) Then
        Set rstDB = New ADODB.Recordset
    End If

    Set GetDATA_NCTS_TAG_Record = rstDB

    ADORecordsetClose rstDB
    'Set rstDB = Nothing

'
End Function


Public Function LoadTrueSeals(ByRef ActiveSeals As cpiOldSeals_Tables, _
                                                    ByRef clsMasterRecord As cpiMASTEREDINCTSIE44) As Boolean
'
    'select case
'
    ' ActiveSeals(1).se
    Dim strFirstSeal As String
    Dim strLastSeal As String

    If (ActiveSeals(1).Seals.Count = 1) Then
        strFirstSeal = ActiveSeals(1).Seals(1).SEALS_IDENTITY
        strLastSeal = ""
    ElseIf (ActiveSeals(1).Seals.Count = 2) Then
        strFirstSeal = ActiveSeals(1).Seals(1).SEALS_IDENTITY
        strLastSeal = ActiveSeals(1).Seals(2).SEALS_IDENTITY
    ElseIf (ActiveSeals(1).Seals.Count > 2) Then

        ' check if there is repetiton
        Dim blnRepetition As Boolean
        Dim intSealCtr As Integer

        blnRepetition = True
        For intSealCtr = 2 To ActiveSeals(1).Seals.Count
            If (ActiveSeals(1).Seals(intSealCtr).SEALS_IDENTITY <> ActiveSeals(1).Seals(1).SEALS_IDENTITY) Then
                blnRepetition = False
                Exit For
            End If
        Next intSealCtr

        If (blnRepetition = True) Then
            strFirstSeal = ActiveSeals(1).Seals(1).SEALS_IDENTITY
            strLastSeal = "*" & CStr(ActiveSeals(1).Seals.Count)

        ElseIf (blnRepetition = False) Then

            ' check if there if seqence default check last two digit
            Dim blnInSequence As Boolean
            Dim intSealNo As Integer

            blnInSequence = True
            For intSealCtr = 1 To ActiveSeals(1).Seals.Count
                'If (IsNumeric(Right$(ActiveSeals(1).Seals(intSealCtr).SEALS_IDENTITY, 2)) = True) And _
                    (IsNumeric(Right$(ActiveSeals(1).Seals(intSealCtr - 1).SEALS_IDENTITY, 2)) = True) Then
                    'If ((Right$(ActiveSeals(1).Seals(intSealCtr).SEALS_IDENTITY, 2) - _
                        Right$(ActiveSeals(1).Seals(intSealCtr - 1).SEALS_IDENTITY, 2)) <> 1) Then

                If (IsNumeric(Right$(ActiveSeals(1).Seals(intSealCtr).SEALS_IDENTITY, 2)) = True) Then
                    intSealNo = Val(Right$(ActiveSeals(1).Seals(intSealCtr).SEALS_IDENTITY, 2))
                    If (intSealNo <> intSealCtr) Then
                        blnInSequence = False
                        Exit For
                    End If
                Else
                    blnInSequence = False
                    Exit For
                End If
            Next intSealCtr

            If (blnInSequence = True) Then
                strFirstSeal = ActiveSeals(1).Seals(1).SEALS_IDENTITY
                strLastSeal = ActiveSeals(1).Seals(ActiveSeals(1).Seals.Count).SEALS_IDENTITY
            ElseIf (blnInSequence = False) Then
                strFirstSeal = "*"
                strLastSeal = "*"
            End If

        End If

    End If

    'ActiveSeals.SEAL_START = strFirstSeal
    '
    'ActiveSeals(1).SEAL_END = strLastSeal

    ActiveSeals(1).FIELD_SEAL_START = strFirstSeal
    ActiveSeals(1).FIELD_SEAL_END = strLastSeal
    'clsMasterRecord.FIELD_AE = strFirstSeal
    'clsMasterRecord.FIELD_AF = strLastSeal

End Function


Private Function GetNCTS_MSG_ID_IE43(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef clsMessage As cpiIE44Message) As Long
'
    Dim strSQL As String

    Dim clsDataNctsTable As cpiDataNctsTable
    Dim clsDataNctsTables As cpiDataNctsTables
    Dim clsDataNctsMessage As cpiDataNctsMessage
    Dim clsDataNctsMessages As cpiDataNctsMessages

    Dim lngDATA_NCTS_MSG_ID As Long

    Set clsDataNctsTables = New cpiDataNctsTables
    Set clsDataNctsMessages = New cpiDataNctsMessages
    'Set clsDataNctsMessage = New cpiDataNctsMessage

        strSQL = "SELECT * FROM [DATA_NCTS] "
        strSQL = strSQL & " WHERE [CODE]=" & Chr(39) & ProcessQuotes(clsMasterRecord.FIELD_CODE) & Chr(39)
        strSQL = strSQL & " AND [MRN]=" & Chr(39) & ProcessQuotes(clsMasterRecord.FIELD_MR) & Chr(39)
    ADORecordsetOpen strSQL, EdifactDB, clsDataNctsTables.Recordset, adOpenKeyset, adLockOptimistic
    'Set clsDataNctsTables.Recordset = EdifactDB.Execute(strSQL)

    If (clsDataNctsTables.Recordset.EOF = False) Then

        Set clsDataNctsTable = clsDataNctsTables.GetClassRecord(clsDataNctsTables.Recordset)

        ' get data_ncts_id
            strSQL = "SELECT * FROM [DATA_NCTS_MESSAGES] "
            strSQL = strSQL & " WHERE [NCTS_IEM_ID]=" & CStr(CONST_NCTS_IEM_ID_MSG_IE43)
            strSQL = strSQL & " AND [DATA_NCTS_ID]= " & CStr(clsDataNctsTable.FIELD_DATA_NCTS_ID)
        ADORecordsetOpen strSQL, EdifactDB, clsDataNctsMessages.Recordset, adOpenKeyset, adLockOptimistic
        'Set clsDataNctsMessages.Recordset = EdifactDB.Execute(strSQL)

        If (clsDataNctsMessages.Recordset.EOF = False) Then

            Set clsDataNctsMessage = clsDataNctsMessages.GetClassRecord(clsDataNctsMessages.Recordset)

            lngDATA_NCTS_MSG_ID = clsDataNctsMessage.FIELD_DATA_NCTS_MSG_ID
            GetNCTS_MSG_ID_IE43 = lngDATA_NCTS_MSG_ID
            ' GetIE43Data udtActiveParameters, lngDATA_NCTS_MSG_ID, clsIE44Message, clsMasterRecord
            ' GetTagsIE43 = True

        ElseIf (clsDataNctsMessages.Recordset.EOF = True) Then
            'GetTagsIE43 = False
        End If

    ElseIf (clsDataNctsTables.Recordset.EOF = True) Then
        'GetTagsIE43 = False
    End If

    Set clsDataNctsTable = Nothing
    Set clsDataNctsTables = Nothing

    Set clsDataNctsMessages = Nothing
    Set clsDataNctsMessage = Nothing

End Function

Private Function GetFTX_AAA(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_FTX_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer)

    Dim clsFTXs As cpiDATA_NCTS_FTXs
    Dim clsFTX As cpiDATA_NCTS_FTX

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strGoodsDesc As String

    Set clsFTXs = New cpiDATA_NCTS_FTXs
    Set clsFTX = New cpiDATA_NCTS_FTX

    ' seq 6,7,8,9, 11

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsFTX.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsFTX.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' convert  class to recordset
    'Set clsFTXs.Recordset = clsFTXs.GetTableRecord(EdifactDB, clsFTX)


    ' MAP - start    ' map message and clsFTX here
    ' get tag recordset
    Set clsFTXs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        lngDATA_NCTS_FTX_ParentID)

    ' loop here
    Do While (clsFTXs.Recordset.EOF = False)

        Set clsFTX = clsFTXs.GetClassRecord(clsFTXs.Recordset)

        strGoodsDesc = clsFTX.FIELD_DATA_NCTS_FTX_Seq6
        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq7
        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq8
        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq9

        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_GOODS_DESCRIPTION = strGoodsDesc

        ' LNG
        'clsFTX.FIELD_DATA_NCTS_FTX_Seq11 = clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION

        clsFTXs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsFTXs = Nothing
    Set clsFTX = Nothing
'
End Function


Private Function IE44_GetFTX_ABV_Detail(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_FTX_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intResultOfControlTotal As Integer)

    Dim clsFTXs As cpiDATA_NCTS_FTXs
    Dim clsFTX As cpiDATA_NCTS_FTX

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strResultDesc As String
    Dim intResultOfControlCtr As Integer

'    Static intResultOfControlInstance As Integer

    Set clsFTXs = New cpiDATA_NCTS_FTXs
    Set clsFTX = New cpiDATA_NCTS_FTX

    ' seq 3,6,7,8,11

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsFTX.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsFTX.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsFTXs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        lngDATA_NCTS_FTX_ParentID)

    If (clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls Is Nothing) = True Then
        Set clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls = New cpiResultOfControls
    End If


    ' loop here
    Do While (clsFTXs.Recordset.EOF = False)

        Set clsFTX = clsFTXs.GetClassRecord(clsFTXs.Recordset)

        intResultOfControlTotal = intResultOfControlTotal + 1
        intResultOfControlCtr = intResultOfControlCtr + 1 'intResultOfControlTotal

        clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls.Add CStr(intResultOfControlTotal), clsMasterRecord.FIELD_CODE

        ' seq 3 - Control indicator (R, an2)
        clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls(intResultOfControlCtr).FIELD_CONTROL_INDICATOR = clsFTX.FIELD_DATA_NCTS_FTX_Seq3

        ' seq 6 - Pointer to the attribute (O, an..35)
        clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls(intResultOfControlCtr).FIELD_POINTER_TO_THE_ATTRIBUTE = clsFTX.FIELD_DATA_NCTS_FTX_Seq6

        ' *
        strResultDesc = clsFTX.FIELD_DATA_NCTS_FTX_Seq7
        strResultDesc = strResultDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq8

        clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls(intResultOfControlCtr).FIELD_DESCRIPTION = strResultDesc
        '
'        strGoodsDesc = clsFTX.FIELD_DATA_NCTS_FTX_Seq6
'        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq7
'        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq8
'        strGoodsDesc = strGoodsDesc & clsFTX.FIELD_DATA_NCTS_FTX_Seq9
'
'        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_GOODS_DESCRIPTION = strGoodsDesc
        ' LNG
        'clsFTX.FIELD_DATA_NCTS_FTX_Seq11 = clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION

        clsFTXs.Recordset.MoveNext

    Loop

'    ' MAP - start    ' map message and clsFTX here
'    If (intResultOfControlInstance >= intResultOfControlTotal) Then
'        intResultOfControlInstance = 0
'    End If
'
'    For intResultOfControlCtr = 1 To clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls.Count
'
'        intResultOfControlInstance = intResultOfControlInstance + 1
'
'        ' seq 3 - Control indicator (R, an2)
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq3 = clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls(intResultOfControlCtr).FIELD_CONTROL_INDICATOR
'
'        ' seq 6 - Pointer to the attribute (O, an..35)
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq6 = clsMessage.GoodsItems(intGoodsItemCtr).ResultOfControls(intResultOfControlCtr).FIELD_POINTER_TO_THE_ATTRIBUTE
'
'        ' seq 7-8
'        strResultDesc = clsMessage.GoodsItems(intGoodsItemCtr).FIELD_GOODS_DESCRIPTION
'        If (Len(strResultDesc) <= 70) Then
'            ' Goods description 1..70
'            clsFTX.FIELD_DATA_NCTS_FTX_Seq7 = strResultDesc
'        ElseIf (Len(strResultDesc) <= 140) Then
'            ' Goods description 1..70
'            clsFTX.FIELD_DATA_NCTS_FTX_Seq7 = Left$(strResultDesc, 70)
'            ' Goods description 71..140
'            clsFTX.FIELD_DATA_NCTS_FTX_Seq8 = Mid$(strResultDesc, 71)
'        End If
'
'        ' LNG
'        clsFTX.FIELD_DATA_NCTS_FTX_Seq11 = clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION
'
'        'clsFTX.FIELD_DATA_NCTS_FTX_ParentID
'        ' set instance here
'        'clsFTX.FIELD_DATA_NCTS_FTX_Instance = 1
'        clsFTX.FIELD_DATA_NCTS_FTX_Instance = intResultOfControlInstance
'        clsFTX.FIELD_DATA_NCTS_FTX_ParentID = lngDATA_NCTS_FTX_ParentID
'
'        ' map values here
'        ' add pk +1
'        clsFTXs.GetMaxID EdifactDB, clsFTX
'        clsFTX.FIELD_DATA_NCTS_FTX_ID = clsFTX.FIELD_DATA_NCTS_FTX_ID + 1
'
'        ' add record to db
'        clsFTXs.AddRecord EdifactDB, clsFTX
'
'    Next intResultOfControlCtr
        ' add children here

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsFTXs = Nothing
    Set clsFTX = Nothing
'
End Function


Private Function GetMEA_AAB(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef lngDATA_NCTS_MEA_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer)
'
    Dim clsMEAs As cpiDATA_NCTS_MEAs
    Dim clsMEA As cpiDATA_NCTS_MEA

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strGoodsDesc As String

    Set clsMEAs = New cpiDATA_NCTS_MEAs
    Set clsMEA = New cpiDATA_NCTS_MEA

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsMEA.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsMEA.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsMEAs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        lngDATA_NCTS_MEA_ParentID)

    'loop here
    Do While (clsMEAs.Recordset.EOF = False)

        Set clsMEA = clsMEAs.GetClassRecord(clsMEAs.Recordset)

        ' seq 7 - Gross mass (O, n..11,3)
        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_GROSS_MASS = clsMEA.FIELD_DATA_NCTS_MEA_Seq7

        clsMEAs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsMEAs = Nothing
    Set clsMEA = Nothing

'
End Function


Private Function GetMEA_AAA(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef lngDATA_NCTS_MEA_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer)

    Dim clsMEAs As cpiDATA_NCTS_MEAs
    Dim clsMEA As cpiDATA_NCTS_MEA

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strGoodsDesc As String

    Set clsMEAs = New cpiDATA_NCTS_MEAs
    Set clsMEA = New cpiDATA_NCTS_MEA

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsMEA.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsMEA.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsMEAs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, strTagName, _
                        lngDATA_NCTS_MEA_ParentID)

    'loop here
    Do While (clsMEAs.Recordset.EOF = False)

        Set clsMEA = clsMEAs.GetClassRecord(clsMEAs.Recordset)

        ' seq 7 - Net mass (O, n..11,3)
        clsMessage.GoodsItems(intGoodsItemCtr).FIELD_NET_MASS = clsMEA.FIELD_DATA_NCTS_MEA_Seq7

     clsMEAs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsMEAs = Nothing
    Set clsMEA = Nothing

'
End Function

Private Function GetPAC_6(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_PAC_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intPackagesTotal As Integer)
'
    Dim clsPACs As cpiDATA_NCTS_PACs
    Dim clsPAC As cpiDATA_NCTS_PAC

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim intPackageCtr As Integer
    Dim intPackageInstance As Integer

    Set clsPACs = New cpiDATA_NCTS_PACs
    Set clsPAC = New cpiDATA_NCTS_PAC

    ' MAP - start
    ' map message and clsPAC here
'    If (intPackageInstance >= intPackagesTotal) Then
'        intPackageInstance = 0
'    End If

'    For intPackageCtr = 1 To clsMessage.GoodsItems(intGoodsItemCtr).Packages.Count
'
'        intPackageInstance = intPackageInstance + 1
'
'        ' seq 9 - Kind of packages (R, a2)
'        clsPAC.FIELD_DATA_NCTS_PAC_Seq9 = clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_KIND_OF_PACKAGES
'
'        ' seq10 - Number of packages (O, n..5)
'        clsPAC.FIELD_DATA_NCTS_PAC_Seq10 = clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_NUMBER_OF_PACKAGES
'
'        ' seq12 - Number of pieces (O, n..5)
'        clsPAC.FIELD_DATA_NCTS_PAC_Seq12 = clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_NUMBER_OF_PIECES
'
'        ' set instance here
'        clsPAC.FIELD_DATA_NCTS_PAC_Instance = CLng(intGoodsItemCtr)
'        clsPAC.FIELD_DATA_NCTS_PAC_ParentID = lngDATA_NCTS_PAC_ParentID
'
'        ' map values here
'        clsPACs.GetMaxID EdifactDB, clsPAC
'        clsPAC.FIELD_DATA_NCTS_PAC_ID = clsPAC.FIELD_DATA_NCTS_PAC_ID + 1
'
'        ' add record to db
'        clsPACs.AddRecord EdifactDB, clsPAC
'
'        ' add children here
'        GetPCI_28 udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "PCI", 101, MSG_IE07, "28", _
'                    clsMasterRecord, clsPAC.FIELD_DATA_NCTS_PAC_ID, intGoodsItemCtr, intPackageCtr, intPackageInstance
'
'    Next intPackageCtr

    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsPAC.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsPAC.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' get tag recordset
    Set clsPACs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, lngDATA_NCTS_PAC_ParentID)

    ' loop here i love u jesus
    Do While (clsPACs.Recordset.EOF = False)

        ' i'm dying
        Set clsPAC = clsPACs.GetClassRecord(clsPACs.Recordset)

        'intPackageCtr = intPackageCtr + 1
        intPackagesTotal = intPackagesTotal + 1
        intPackageCtr = intPackageCtr + 1

        ' <<<<<<<<<<<<<<<<<<
        clsMessage.GoodsItems(intGoodsItemCtr).Packages.Add CStr(intPackageCtr), clsMasterRecord.FIELD_CODE
        'clsMessage.SealInfos(1).Seals.Add CStr(clsPAC.FIELD_DATA_NCTS_PAC_Instance), clsMasterRecord.FIELD_CODE, 0

        ' seq 9 - Kind of packages (R, a2)
        clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_KIND_OF_PACKAGES = clsPAC.FIELD_DATA_NCTS_PAC_Seq9

        ' seq10 - Number of packages (O, n..5)
        clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_NUMBER_OF_PACKAGES = clsPAC.FIELD_DATA_NCTS_PAC_Seq10

        ' seq12 - Number of pieces (O, n..5)
        clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_NUMBER_OF_PIECES = clsPAC.FIELD_DATA_NCTS_PAC_Seq12

        ' add children here
        GetPCI_28 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "PCI", 101, lngNCTS_IEM_ID, "28", _
                    clsMasterRecord, clsPAC.FIELD_DATA_NCTS_PAC_ID, intGoodsItemCtr, intPackageCtr, intPackageInstance

        'set GroupIndex
        clsMessage.GoodsItems(intGoodsItemCtr).Packages.GroupIndex = 1

        clsPACs.Recordset.MoveNext

    Loop  ' @@@ break the BREAK ! ! !  :-)

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsPACs = Nothing
    Set clsPAC = Nothing
'
End Function


Private Function IE44_GetRFF_AAQ(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_RFF_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intContainerTotal As Integer)
'
    Dim clsRFFs As cpiDATA_NCTS_RFFs
    Dim clsRFF As cpiDATA_NCTS_RFF

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table
    Dim intContainerCtr As Integer

    Dim strSQL As String

'    Static intContainerInstance As Integer

    Set clsRFFs = New cpiDATA_NCTS_RFFs
    Set clsRFF = New cpiDATA_NCTS_RFF

    ' seq 2
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsRFF.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsRFF.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start
    ' get tag recordset
    Set clsRFFs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, lngDATA_NCTS_RFF_ParentID)

    ' loop here i love u jesus
    Do While (clsRFFs.Recordset.EOF = False)

        ' i'm dying
        Set clsRFF = clsRFFs.GetClassRecord(clsRFFs.Recordset)

        intContainerTotal = intContainerTotal + 1
        intContainerCtr = intContainerCtr + 1 'intContainerTotal

        ' add new container
        clsMessage.GoodsItems(intGoodsItemCtr).Containers.Add CStr(intContainerCtr), clsMasterRecord.FIELD_CODE
        ' Container number
        clsMessage.GoodsItems(intGoodsItemCtr).Containers(intContainerCtr).NEW_CONTAINER_NUMBER = clsRFF.FIELD_DATA_NCTS_RFF_Seq2

        clsRFFs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsRFFs = Nothing
    Set clsRFF = Nothing

End Function


Private Function IE43_GetRFF_AAQ(ByRef EdifactDB As ADODB.Connection, _
                                 ByRef lngDATA_NCTS_MSG_ID As Long, _
                                 ByRef clsMessage As cpiIE44Message, _
                                 ByRef strTagName As String, _
                                 ByRef lngEDI_TMS_ID As Long, _
                                 ByRef lngNCTS_IEM_ID As Long, _
                                 ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                 ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                 ByRef lngDATA_NCTS_RFF_ParentID As Long, _
                                 ByRef intGoodsItemCtr As Integer, _
                                 ByRef intContainerTotal As Integer)

    Dim clsRFFs As cpiDATA_NCTS_RFFs
    Dim clsRFF As cpiDATA_NCTS_RFF

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String

    Dim intContainerCtr As Integer
    Dim intExtraContainerCount As Integer
    Dim intContainerInstance As Integer

    Dim objContainer As cpiContainer

    Set clsRFFs = New cpiDATA_NCTS_RFFs
    Set clsRFF = New cpiDATA_NCTS_RFF

    ' seq 2
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                                lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsRFF.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsRFF.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start
    ' get tag recordset
    Set clsRFFs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                            lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                            strTagName, lngDATA_NCTS_RFF_ParentID)

    ' init container
    'clsMessage.GoodsItems(intGoodsItemCtr).Containers.ItemsPerGroup

    ' loop here i love u jesus
    Do Until clsRFFs.Recordset.EOF

        ' i'm dying
        Set clsRFF = clsRFFs.GetClassRecord(clsRFFs.Recordset)

        intContainerInstance = intContainerInstance + 1

        ' add new container
        'clsMessage.GoodsItems(intGoodsItemCtr).Containers.Add CStr(intContainerCtr), clsMasterRecord.FIELD_CODE
        Set objContainer = clsMessage.GoodsItems(intGoodsItemCtr).Containers.Add(CStr(intContainerInstance), clsMasterRecord.FIELD_CODE)
        objContainer.NEW_CONTAINER_NUMBER = clsRFF.FIELD_DATA_NCTS_RFF_Seq2

        'set GroupIndex
        clsMessage.GoodsItems(intGoodsItemCtr).Containers.GroupIndex = 1

        clsRFFs.Recordset.MoveNext

    Loop

    intContainerTotal = clsMessage.GoodsItems(intGoodsItemCtr).Containers.Count

    ' Containers over a multiple of 5 are treated as extra containers
    intExtraContainerCount = intContainerTotal Mod 5

    If intExtraContainerCount <> 0 Then
        ' Fill the empty container number boxes until a multiple of 5 is reached
        For intContainerCtr = 1 To 5 - intExtraContainerCount
            intContainerInstance = intContainerInstance + 1
            'clsMessage.GoodsItems(intGoodsItemCtr).Containers.Add CStr(clsMessage.GoodsItems(intGoodsItemCtr).Containers.Count + 1), clsMasterRecord.FIELD_CODE
            clsMessage.GoodsItems(intGoodsItemCtr).Containers.Add CStr(intContainerInstance), clsMasterRecord.FIELD_CODE

            'set GroupIndex
            clsMessage.GoodsItems(intGoodsItemCtr).Containers.GroupIndex = 1
        Next
    End If

    Set objContainer = Nothing
    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsRFFs = Nothing
    Set clsRFF = Nothing

End Function

Private Function GetDOC_916(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_DOC_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intDocCertTotal As Integer, _
                                        ByVal NewCode As Boolean)
'
    Dim clsDOCs As cpiDATA_NCTS_DOCs
    Dim clsDOC As cpiDATA_NCTS_DOC

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table
    Dim intDocCertCtr As Integer

    Dim strSQL As String

'    Static intDocCertInstance As Integer

    Set clsDOCs = New cpiDATA_NCTS_DOCs
    Set clsDOC = New cpiDATA_NCTS_DOC

    ' seq 4, 5, 6, 7, 8
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsDOC.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsDOC.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start    ' get tag recordset
    Set clsDOCs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, lngDATA_NCTS_DOC_ParentID)


'    ' map message and clsDOC here
'    ' seq 4, 5, 6, 7, 8
    ' loop here i love u jesus
    Do While (clsDOCs.Recordset.EOF = False)

        ' i'm dying
        Set clsDOC = clsDOCs.GetClassRecord(clsDOCs.Recordset)

        intDocCertTotal = intDocCertTotal + 1
        intDocCertCtr = intDocCertCtr + 1 'intDocCertTotal

        ' add new container
        clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates.Add CStr(intDocCertCtr), clsMasterRecord.FIELD_CODE
        ' seq 4 - Document type (O, an..3) ./.\.
        clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates(intDocCertCtr).FIELD_DOCUMENT_TYPE = clsDOC.FIELD_DATA_NCTS_DOC_Seq4

        ' seq 5 - Document reference (O, an..20)
        clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates(intDocCertCtr).FIELD_DOCUMENT_REFERENCE = clsDOC.FIELD_DATA_NCTS_DOC_Seq5

        ' seq 6 - Complement of information LNG (D, a2)
        clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = clsDOC.FIELD_DATA_NCTS_DOC_Seq6

        ' seq 7 - Complement of information (O, an..26)
        clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates(intDocCertCtr).FIELD_COMPLEMENT_INFORMATION = clsDOC.FIELD_DATA_NCTS_DOC_Seq7

        ' seq 8 - Document reference LNG (D, a2)
        'clsDOC.FIELD_DATA_NCTS_DOC_Seq8 = clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION

        If NewCode Then
            'set GroupIndex
            clsMessage.GoodsItems(intGoodsItemCtr).DocumentCertificates.GroupIndex = 1
        End If

        clsDOCs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsDOCs = Nothing
    Set clsDOC = Nothing

End Function


Private Function GetGIR_3(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_GIR_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intSGICodeTotal As Integer)
'
    Dim clsGIRs As cpiDATA_NCTS_GIRs
    Dim clsGIR As cpiDATA_NCTS_GIR

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table
    Dim intSGICodeCtr As Integer

    Dim strSQL As String

'    Static intSGICodeInstance As Integer

    Set clsGIRs = New cpiDATA_NCTS_GIRs
    Set clsGIR = New cpiDATA_NCTS_GIR

    ' seq 2, 5
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsGIR.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsGIR.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' map message and clsGIR here
    ' seq 2, 5
'    For intSgiCodeCtr = 1 To clsMessage.GoodsItems(intGoodsItemCtr).SGICodes.Count
'
'        intSgiCodeInstance = intSgiCodeInstance + 1
'        '
'        ' seq 2 - Sensitive quantity (O, n..11,3)
'        clsGIR.FIELD_DATA_NCTS_GIR_Seq2 = clsMessage.GoodsItems(intGoodsItemCtr).SGICodes(intSgiCodeCtr).FIELD_SENSITIVE_QUANTITY
'
'        ' seq 5 - Sensitive goods code (O, n..2)
'        clsGIR.FIELD_DATA_NCTS_GIR_Seq5 = clsMessage.GoodsItems(intGoodsItemCtr).SGICodes(intSgiCodeCtr).FIELD_SENSITIVE_GOODS_CODE
'
'        ' clsGIR.FIELD_DATA_NCTS_GIR_ParentID
'        ' set instance here
'        clsGIR.FIELD_DATA_NCTS_GIR_Instance = CLng(intSgiCodeInstance)   'CLng(intSgiCodeCtr) 'intSgiCodeCtr
'        clsGIR.FIELD_DATA_NCTS_GIR_ParentID = lngDATA_NCTS_GIR_ParentID
'
'        ' map values here
'        clsGIRs.GetMaxID EdifactDB, clsGIR
'        clsGIR.FIELD_DATA_NCTS_GIR_ID = clsGIR.FIELD_DATA_NCTS_GIR_ID + 1
'
'        ' add record to db
'        clsGIRs.AddRecord EdifactDB, clsGIR
'
'    Next intSgiCodeCtr

    ' MAP - start    ' get tag recordset
    Set clsGIRs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, lngDATA_NCTS_GIR_ParentID)

    ' map message and clsGIR here
    ' seq 4, 5, 6, 7, 8
    ' loop here i love u jesus
    Do While (clsGIRs.Recordset.EOF = False)

        ' i'm dying
        Set clsGIR = clsGIRs.GetClassRecord(clsGIRs.Recordset)

        intSGICodeTotal = intSGICodeTotal + 1
        intSGICodeCtr = intSGICodeCtr + 1 'intSgiCodeTotal

        ' add new container
        clsMessage.GoodsItems(intGoodsItemCtr).SGICodes.Add CStr(intSGICodeCtr), clsMasterRecord.FIELD_CODE

        ' seq 2 - Sensitive quantity (O, n..11,3)
        clsMessage.GoodsItems(intGoodsItemCtr).SGICodes(intSGICodeCtr).FIELD_SENSITIVE_QUANTITY = clsGIR.FIELD_DATA_NCTS_GIR_Seq2

        ' seq 5 - Sensitive goods code (O, n..2)
        clsMessage.GoodsItems(intGoodsItemCtr).SGICodes(intSGICodeCtr).FIELD_SENSITIVE_GOODS_CODE = clsGIR.FIELD_DATA_NCTS_GIR_Seq5

        clsGIRs.Recordset.MoveNext

    Loop

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsGIRs = Nothing
    Set clsGIR = Nothing

End Function


Public Sub GetOtherNonSegmentBox_UnloadingRemarks(ByRef EdifactDB As ADODB.Connection, _
                                                    ByRef clsMessage As cpiIE44Message)

    Dim i As Long
    Dim j As Long
    Dim rstTemp As ADODB.Recordset

    Set rstTemp = New ADODB.Recordset

    '===================== Header =====================================================
    rstTemp.Open "Select * from DATA_NCTS_HEADER_RESULTATEN where CODE = '" & clsMessage.CODE_FIELD & "' Order by ORDINAL", EdifactDB, adOpenKeyset, adLockReadOnly
    For i = 1 To clsMessage.ResultOfControls.Count
        'i-> Ordinal

        rstTemp.Find "ORDINAL = " & i
        If Not rstTemp.EOF Then
            clsMessage.ResultOfControls(i).FIELD_CM = IIf(IsNull(rstTemp!CM), "", Trim(rstTemp!CM))
        End If

    Next
    rstTemp.Close
    '==================================================================================

    '==================== Details ======================================================
    For i = 1 To clsMessage.GoodsItems.Count


        rstTemp.Open "Select * from DATA_NCTS_DETAIL_COLLI where CODE = '" & clsMessage.CODE_FIELD & "' and DETAIL = " & i & " order by ORDINAL", EdifactDB, adOpenKeyset, adLockReadOnly

        'for box S5
        For j = 1 To clsMessage.GoodsItems(i).Packages.Count
            rstTemp.Find "ORDINAL = " & j
            If Not rstTemp.EOF Then
                clsMessage.GoodsItems(i).Packages(j).FIELD_S5 = IIf(IsNull(rstTemp!S5), "", Trim(rstTemp!S5))
            End If

        Next
        rstTemp.Close

        rstTemp.Open "Select * from DATA_NCTS_DETAIL_CONTAINER where CODE = '" & clsMessage.CODE_FIELD & "' and DETAIL = " & i & " order by ORDINAL", EdifactDB, adOpenKeyset, adLockReadOnly
        'for box SB
        For j = 1 To clsMessage.GoodsItems(i).Containers.Count / 5
            rstTemp.Find "ORDINAL = " & j
            If Not rstTemp.EOF Then
                clsMessage.GoodsItems(i).Containers(((j - 1)) * 5 + 1).FIELD_SB = IIf(IsNull(rstTemp!SB), "", Trim(rstTemp!SB))
            End If

        Next
        rstTemp.Close

        rstTemp.Open "Select * from DATA_NCTS_DETAIL_DOCUMENTEN where CODE = '" & clsMessage.CODE_FIELD & "' and DETAIL = " & i & " order by ORDINAL", EdifactDB, adOpenKeyset, adLockReadOnly
        'for box Y1 and Y5
        For j = 1 To clsMessage.GoodsItems(i).DocumentCertificates.Count
            rstTemp.Find "ORDINAL = " & j
            If Not rstTemp.EOF Then
                clsMessage.GoodsItems(i).DocumentCertificates(j).FIELD_Y1 = IIf(IsNull(rstTemp!Y1), "", Trim(rstTemp!Y1))
                clsMessage.GoodsItems(i).DocumentCertificates(j).FIELD_Y5 = IIf(IsNull(rstTemp!Y5), "", Trim(rstTemp!Y5))
            End If

        Next
        rstTemp.Close

        'for box CL
        rstTemp.Open "Select * from DATA_NCTS_DETAIL_RESULTATEN where CODE = '" & clsMessage.CODE_FIELD & "' and DETAIL = " & i & " order by ORDINAL", EdifactDB, adOpenKeyset, adLockReadOnly
        For j = 1 To clsMessage.GoodsItems(i).ResultOfControls.Count
            rstTemp.Find "ORDINAL = " & j
            If Not rstTemp.EOF Then
                clsMessage.GoodsItems(i).ResultOfControls(j).FIELD_CL = IIf(IsNull(rstTemp!CL), "", Trim(rstTemp!CL))
            End If

        Next
        rstTemp.Close

        'for box T7
        rstTemp.Open "Select * from DATA_NCTS_DETAIL where CODE = '" & clsMessage.CODE_FIELD & "' and DETAIL = " & i, EdifactDB, adOpenKeyset, adLockReadOnly
        For j = 1 To clsMessage.GoodsItems.Count
            If Not rstTemp.BOF And Not rstTemp.EOF Then
                clsMessage.GoodsItems(i).FIELD_T7 = IIf(IsNull(rstTemp!T7), "", Trim(rstTemp!T7))
            End If
        Next
        rstTemp.Close

    Next
    '================= end detail ====================================================

    Set rstTemp = Nothing
End Sub

Private Function GetPCI_28(ByRef EdifactDB As ADODB.Connection, _
                                        ByRef lngDATA_NCTS_MSG_ID As Long, _
                                        ByRef clsMessage As cpiIE44Message, _
                                        ByRef strTagName As String, _
                                        ByRef lngEDI_TMS_ID As Long, _
                                        ByRef lngNCTS_IEM_ID As Long, _
                                        ByRef strNCTS_IEM_TMS_RemarksQualifier As String, _
                                        ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                                        ByRef lngDATA_NCTS_PCI_ParentID As Long, _
                                        ByRef intGoodsItemCtr As Integer, _
                                        ByRef intPackageCtr As Integer, _
                                        ByRef intPackageInstance As Integer)
'
    Dim clsPCIs As cpiDATA_NCTS_PCIs
    Dim clsPCI As cpiDATA_NCTS_PCI

    Dim clsNCTS_IEM_TMS_Table As cpiNCTS_IEM_TMS_Table

    Dim strSQL As String
    Dim strPackageDesc As String

'    Static intPackageInstance As Integer

    Set clsPCIs = New cpiDATA_NCTS_PCIs
    Set clsPCI = New cpiDATA_NCTS_PCI

    ' seq 2, 3, 12
    Set clsNCTS_IEM_TMS_Table = GetNCTS_IEM_TMS_Table(EdifactDB, lngNCTS_IEM_ID, _
                        lngEDI_TMS_ID, strNCTS_IEM_TMS_RemarksQualifier)

    clsPCI.FIELD_DATA_NCTS_MSG_ID = lngDATA_NCTS_MSG_ID
    clsPCI.FIELD_NCTS_IEM_TMS_ID = clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID

    ' MAP - start
    ' map message and clsPCI here

    ' get tag recordset
    Set clsPCIs.Recordset = GetDATA_NCTS_TAG_Record(EdifactDB, _
                        lngDATA_NCTS_MSG_ID, clsNCTS_IEM_TMS_Table.FIELD_NCTS_IEM_TMS_ID, _
                        strTagName, lngDATA_NCTS_PCI_ParentID)

    Do While clsPCIs.Recordset.EOF = False
    
        Set clsPCI = clsPCIs.GetClassRecord(clsPCIs.Recordset)
    
        strPackageDesc = clsPCI.FIELD_DATA_NCTS_PCI_Seq2
        strPackageDesc = strPackageDesc & clsPCI.FIELD_DATA_NCTS_PCI_Seq3
    
        clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_MARKS_AND_NUMBERS_OF_PACKAGES = strPackageDesc
        clsPCIs.Recordset.MoveNext
    Loop
'    If (Len(strPackageDesc) <= 35) Then
'        ' seq 2 Marks & numbers of Packages (O, an..42) 1..35
'        clsPCI.FIELD_DATA_NCTS_PCI_Seq2 = strPackageDesc
'    ElseIf (Len(strPackageDesc) <= 42) Then
'        ' seq 2 Marks & numbers of Packages (O, an..42) 1..35
'        clsPCI.FIELD_DATA_NCTS_PCI_Seq2 = Left$(strPackageDesc, 35)
'        ' seq 3     36..42
'        clsPCI.FIELD_DATA_NCTS_PCI_Seq3 = Mid$(strPackageDesc, 36)
'    End If

    ' seq 12 - Marks & numbers of Packages LNG (D, a2)
    'clsPCI.FIELD_DATA_NCTS_PCI_Seq12 = clsMessage.GoodsItems(intGoodsItemCtr).Packages(intPackageCtr).FIELD_MARKS_AND_NUMBERS_OF_PackageS_LNG

    Set clsNCTS_IEM_TMS_Table = Nothing
    Set clsPCIs = Nothing
    Set clsPCI = Nothing
'
End Function

'p4tric 021408
'Variable declarations on XML Structure
'
'   <ParentNode>
'       <ChildNode>
'           <objChildElement>
'               <objChildElement2>
'                   <objChildElement3>

' ************************************************************************************************************************************************
' IE44 get tags here weeeh
Public Function GetData(ByRef EdifactDB As ADODB.Connection, _
                            ByRef lngDATA_NCTS_MSG_ID As Long, _
                            ByRef clsMasterRecord As cpiMASTEREDINCTSIE44, _
                            ByRef clsMessage As cpiIE44Message, _
                            ByVal NCTS_IEM_ID As Long, _
                            ByVal NewCode As Boolean) As Boolean
                            
                            'ByRef ActiveBar As Shape,
                            'ByRef blnSending As Boolean) As Boolean
    
    ' set language ar dep
    clsMessage.Headers(1).DIALOG_LANGUAGE_INDICATOR_AT_DESTINATION = clsMasterRecord.FIELD_AJ
    
    GetBGM EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "BGM", 3, NCTS_IEM_ID, ""
    
    GetTOD EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "TOD", 27, NCTS_IEM_ID, "5"

    GetLOC_22 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "LOC", 5, NCTS_IEM_ID, "22"
    
    If Not NewCode Then
        GetMEA_AAD EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 9, NCTS_IEM_ID, "WT:AAD:KGM"
    Else
        GetMEA_AAD EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "MEA", 9, NCTS_IEM_ID, "WT:AAD:KGM"
    End If
    
    GetSEL EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "SEL", 11, NCTS_IEM_ID, "", clsMasterRecord
        
    'If Not NewCode Then
        IE44_GetFTX_ABV EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "FTX", 12, NCTS_IEM_ID, "ABV", clsMasterRecord
    'End If
    
    GetTDT_12 EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "TDT", 18, NCTS_IEM_ID, "12"
        ' ; TPL     child
    
    GetNAD_CPD EdifactDB, lngDATA_NCTS_MSG_ID, clsMessage, "NAD", 23, NCTS_IEM_ID, "CPD"
            
'''''    If Not NewCode Then
'''''        IE44_GetTOD_5 udtActiveParameters, lngDATA_NCTS_MSG_ID , clsMessage, "TOD",   27, NCTS_IEM_ID, "5", clsMasterRecord
'''''    End If
        
    'IE44_GetTOD_5
    'IE44_GetUNS_D udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "UNS", 33, NCTS_IEM_ID, "D", clsMasterRecord
        
    ' detail na!!!        '
    'If (clsMessage.GoodsItems.Count > 0) Then
        'AddRFF_AIV udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "RFF", 13, NCTS_IEM_ID, "AIV", clsMasterRecord ', ActiveBar
        'p4tric trail of error 021508
        GetCST EdifactDB, clsMasterRecord, clsMessage, lngDATA_NCTS_MSG_ID, "CST", 93, NCTS_IEM_ID, "", NewCode        ', ActiveBar
    '    ActiveBar.Width = CDbl(ActiveBar.Tag) * 0.75
    'End If

    'IE44_GetUNS_S udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "UNS", 145, NCTS_IEM_ID, "S", clsMasterRecord

    ' CNT here
    GetCNT_5 EdifactDB, clsMessage, lngDATA_NCTS_MSG_ID, "CNT", 146, NCTS_IEM_ID, "5"
    GetCNT_11 EdifactDB, clsMessage, lngDATA_NCTS_MSG_ID, "CNT", 146, NCTS_IEM_ID, "11"
    GetCNT_16 EdifactDB, clsMessage, lngDATA_NCTS_MSG_ID, "CNT", 146, NCTS_IEM_ID, "16"
    
    'GETSEL edifactdb,lngdata_ncts_msg_id,clsmessage,"SEL",
    'IE44_GetUNT udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "UNT", 152, NCTS_IEM_ID, "", clsMasterRecord
    'IE44_GetUNZ udtActiveParameters, lngDATA_NCTS_MSG_ID, clsMessage, "UNZ", 153, NCTS_IEM_ID, "", clsMasterRecord

End Function


Public Function MapIE07ToIE44(ByRef clsMasterEdiNcts2 As cpiMasterEdiNcts2, _
                                ByRef clsMASTEREDINCTSIE44 As cpiMASTEREDINCTSIE44)
'
    ' set default values - MAP
    clsMASTEREDINCTSIE44.FIELD_AJ = clsMasterEdiNcts2.AJ_FIELD
    clsMASTEREDINCTSIE44.FIELD_BD = clsMasterEdiNcts2.BD_FIELD
    clsMASTEREDINCTSIE44.FIELD_CODE = clsMasterEdiNcts2.CODE_FIELD
    clsMASTEREDINCTSIE44.FIELD_COMM = clsMasterEdiNcts2.COMM_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_CREATED = clsMasterEdiNcts2.DATE_CREATED_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_LAST_MODIFIED = clsMasterEdiNcts2.DATE_LAST_MODIFIED_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_LAST_RECEIVED = clsMasterEdiNcts2.DATE_LAST_RECEIVED_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_PRINTED = clsMasterEdiNcts2.DATE_PRINTED_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_REQUESTED = clsMasterEdiNcts2.DATE_REQUESTED_FIELD
    clsMASTEREDINCTSIE44.FIELD_DATE_SEND = clsMasterEdiNcts2.DATE_SEND_FIELD
    clsMASTEREDINCTSIE44.FIELD_DOC_NUMBER = clsMasterEdiNcts2.DOC_NUMBER_FIELD
    clsMASTEREDINCTSIE44.FIELD_DOC_TYPE = clsMasterEdiNcts2.DOC_TYPE_FIELD
    clsMASTEREDINCTSIE44.FIELD_DOCUMENT_NAME = clsMasterEdiNcts2.DOCUMENT_NAME_FIELD
    clsMASTEREDINCTSIE44.FIELD_DTYPE = clsMasterEdiNcts2.DTYPE_FIELD
    'clsMASTEREDINCTSIE44.FIELD_Error_HD = clsMasterEdiNcts2.Error_HD_FIELD
    clsMASTEREDINCTSIE44.FIELD_LAST_MODIFIED_BY = clsMasterEdiNcts2.LAST_MODIFIED_BY_FIELD
    clsMASTEREDINCTSIE44.FIELD_LOGID = clsMasterEdiNcts2.LOGID_FIELD
    clsMASTEREDINCTSIE44.FIELD_LOGID_DESCRIPTION = clsMasterEdiNcts2.LOGID_DESCRIPTION_FIELD
    clsMASTEREDINCTSIE44.FIELD_MR = clsMasterEdiNcts2.MR_FIELD
    clsMASTEREDINCTSIE44.FIELD_PRINT = clsMasterEdiNcts2.Print_FIELD
    clsMASTEREDINCTSIE44.FIELD_PRINTED_BY = clsMasterEdiNcts2.PRINTED_BY_FIELD
    clsMASTEREDINCTSIE44.FIELD_REMARKS = clsMasterEdiNcts2.REMARKS_FIELD
    clsMASTEREDINCTSIE44.FIELD_TREE_ID = clsMasterEdiNcts2.TREE_ID_FIELD
    clsMASTEREDINCTSIE44.FIELD_USER_NO = clsMasterEdiNcts2.USER_NO_FIELD
    clsMASTEREDINCTSIE44.FIELD_USERNAME = clsMasterEdiNcts2.USERNAME_FIELD
    clsMASTEREDINCTSIE44.FIELD_VIEWED = clsMasterEdiNcts2.VIEWED_FIELD
    clsMASTEREDINCTSIE44.FIELD_TYPE = clsMasterEdiNcts2.Type_FIELD
    clsMASTEREDINCTSIE44.FIELD_HEADER = clsMasterEdiNcts2.HEADER_FIELD
    clsMASTEREDINCTSIE44.FIELD_Memo_Field = clsMasterEdiNcts2.Memo_Field_FIELD

End Function

'p4tric 021908
Private Sub CreateMessageResultsOfControl(ByRef objDOM As DOMDocument, _
                               ByRef objParentNode As IXMLDOMNode, _
                               ByRef objChildNode As IXMLDOMNode, _
                               ByRef objChildElement As IXMLDOMNode, _
                               ByRef objChildElement2 As IXMLDOMNode)
        
    
    Dim lngctr As Long
    
        For lngctr = 1 To G_clsIE44Arrival.ResultOfControls.Count
        
            strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_CONTROL_INDICATOR
            
            If Len(Trim$(strTemp)) > 0 Then
                'Results of Control
                Set objChildNode = objDOM.createElement("RESOFCON534")
                objDOM.documentElement.appendChild objChildNode
                objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                'Description
                strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_DESCRIPTION
                If Len(Trim$(strTemp)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("DesTOC2"))
                    objChildElement.Text = strTemp
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                        
                'Description LNG
                strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_DESCRIPTION_LNG
                If Len(Trim$(strTemp)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("DesTOC2LNG"))
                    objChildElement.Text = strTemp
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                'Control Indicator
                strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_CONTROL_INDICATOR
                If Len(Trim$(strTemp)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("ConInd424"))
                    objChildElement.Text = strTemp
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                                
                'Pointer to the attribute
                strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_POINTER_TO_THE_ATTRIBUTE
                If Len(Trim$(strTemp)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("PoiToTheAttTOC5"))
                    objChildElement.Text = strTemp
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                'Corrected value
                strTemp = G_clsIE44Arrival.ResultOfControls(lngctr).FIELD_CORRECTED_VALUE
                If Len(Trim$(strTemp)) > 0 Then
                    Set objChildElement = objChildNode.appendChild(objDOM.createElement("CorValTOC4"))
                    objChildElement.Text = strTemp
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                End If
                
                objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
            End If
        
    Next lngctr
End Sub

'p4tric 021808
Private Sub CreateMessageDestinationTrader(ByRef objDOM As DOMDocument, _
                               ByRef objParentNode As IXMLDOMNode, _
                               ByRef objChildNode As IXMLDOMNode, _
                               ByRef objChildElement As IXMLDOMNode, _
                               ByRef objChildElement2 As IXMLDOMNode)

    'NamTRD7
    'StrAndNumTRD22
    'PosCodTRD23
    'CitTRD24
    'CouTRD25
    'NADLNGRD                            HINDI PA OK UNG BOX VALUE
    'TINTRD59
        
        Set objChildNode = objDOM.createElement("TRADESTRD")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                  
        'Destination Trader Name
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_NAME 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B1", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamTRD7"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
           
        'Destination Trader Street and Number
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_STREET_AND_NUMBER 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumTRD22"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Destination Trader Postal Code
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_POSTAL_CODE 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B7", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodTRD23"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Destination Trader City
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_CITY 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B1", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitTRD24"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
         
        'Destination Trader Country Code
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_COUNTRY_CODE 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouTRD25"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Destination Trader NAD LNG
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_NAD_LNG 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B7", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGRD"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
              
        'Destination Trader TIN
        strTemp = G_clsIE44Arrival.Traders(G_strTraderKey).DESTINATION_TIN 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B7", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRD59"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                      
                      
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub

'>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>

Public Sub CreateXMLMessageIE44(ByRef DataSourceProperties As CDataSourceProperties, _
                                ByRef objDOM As DOMDocument, _
                                ByRef objParentNode As IXMLDOMNode, _
                                ByRef objChildNode As IXMLDOMNode)

    Dim objChildElement As IXMLDOMNode
    Dim objChildElement2 As IXMLDOMNode

    'Interchange
    CreateMessageInterchange objDOM, objParentNode, objChildNode, objChildElement, objChildElement2

    'Header
    CreateMessageHeader objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Trader
    CreateMessageDestinationTrader objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    '(Presentation Office) Customs Office
    CreateMessagePresentationOffice objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Unloading Remarks
    CreateMessageUnloadingRemarks objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
    
    'Results of Control
    CreateMessageResultsOfControl objDOM, objParentNode, objChildNode, objChildElement, objChildElement2
  
    'Seals Info >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    CreateMessageSealsInfo objDOM, objParentNode, objChildNode, objChildElement, objChildElement2

    'Goods Item >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    CreateMessageGoodsItem objDOM, objParentNode, objChildNode, objChildElement, objChildElement2

    Set objChildElement = Nothing
    Set objChildElement2 = Nothing

End Sub

'p4tric 021808
Private Sub CreateMessageUnloadingRemarks(ByRef objDOM As DOMDocument, _
                               ByRef objParentNode As IXMLDOMNode, _
                               ByRef objChildNode As IXMLDOMNode, _
                               ByRef objChildElement As IXMLDOMNode, _
                               ByRef objChildElement2 As IXMLDOMNode)

          'Unloading Remarks
        Set objChildNode = objDOM.createElement("UNLREMREM")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
        'State of seals ok
        strTemp = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_STATE_OF_SEALS_OK
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("StaOfTheSeaOKREM19"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        'Unloading remark
        strTemp = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_UNLOADING_REMARK
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("UnlRemREM53"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Unloading remark LNG
        strTemp = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_UNLOADING_REMARK_LNG
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("UnlRemREM53LNG"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'Conform
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("ConREM65"))
        objChildElement.Text = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_CONFORM
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
        'Unloading completion
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("UnlComREM66"))
        objChildElement.Text = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_UNLOADING_COMPLETION
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Unloading date
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("UnlDatREM67"))
        objChildElement.Text = G_clsIE44Arrival.UnloadingRemarks(G_strHeaderKey).FIELD_UNLOADING_DATE
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                       
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        
End Sub

'p4tric 021808
Private Sub CreateMessagePresentationOffice(ByRef objDOM As DOMDocument, _
                                        ByRef objParentNode As IXMLDOMNode, _
                                        ByRef objChildNode As IXMLDOMNode, _
                                        ByRef objChildElement As IXMLDOMNode, _
                                        ByRef objChildElement2 As IXMLDOMNode)

    'Customs Offices of Departure
    Set objChildNode = objDOM.createElement("CUSOFFPREOFFRES")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Reference Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("RefNumRES1"))
        objChildElement.Text = G_clsIE44Arrival.CustomOffices(G_strCustomOfcKey).REFERENCE_NUMBER 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A4", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

End Sub

'p4tric 021908
Private Sub CreateMessageInterchange(ByRef objDOM As DOMDocument, _
                                    ByRef objParentNode As IXMLDOMNode, _
                                    ByRef objChildNode As IXMLDOMNode, _
                                    ByRef objChildElement As IXMLDOMNode, _
                                    ByRef objChildElement2 As IXMLDOMNode)

    'Syntax Identifier
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("SynIdeMES1"))
    objChildNode.Text = "UNOC"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Syntax Version Number
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("SynVerNumMES2"))
    objChildNode.Text = "3"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Message Sender
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesSenMES3"))
    objChildNode.Text = GetMapFunctionValue("F<RECEIVE QUEUE>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Sender Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("SenIdeCodQuaMES4"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Message Recipient
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesRecMES6"))
    objChildNode.Text = GetMapFunctionValue("F<RECIPIENT>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Recipient Identification Code Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecIdeCodQuaMES7"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Date of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("DatOfPreMES9"))
    objChildNode.Text = GetMapFunctionValue("F<DATE, YYMMDD>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Time of Preparation
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("TimOfPreMES10"))
    objChildNode.Text = GetMapFunctionValue("F<TIME, HHMM>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Interchange Control Reference
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("IntConRefMES11"))
    objChildNode.Text = GetMapFunctionValue("F<1 TIN REF>")
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Recipient's Reference/Password
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecRefMES12"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Recipient's Reference/Password Qualifier
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("RecRefQuaMES13"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Application Reference
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("AppRefMES14"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Priority
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("PriMES15"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Acknowledgement Request
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("AckReqMES16"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Communications Agreement ID
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("ComAgrIdMES17"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Test Indicator
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("TesIndMES18"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Message Identification
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesIdeMES19"))
    objChildNode.Text = "1"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Message Type
    Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesTypMES20"))
    objChildNode.Text = "CC044A"
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Common Access Reference
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("ComAccRefMES21"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'Sequence Number - NOT ON EDIFACT XML MAPPING FOR CUSDEC
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("MesSeqNumMES22"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    'First and Last Transaction - NOT ON EDIFACT XML MAPPING FOR CUSDEC
    'Set objChildNode = objParentNode.appendChild(objDOM.createElement("FirAndLasTraMES23"))
    'objChildNode.Text = vbNullString
    'objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

End Sub

Private Sub CreateMessageHeader(ByRef objDOM As DOMDocument, _
                               ByRef objParentNode As IXMLDOMNode, _
                               ByRef objChildNode As IXMLDOMNode, _
                               ByRef objChildElement As IXMLDOMNode, _
                               ByRef objChildElement2 As IXMLDOMNode)

    'Header
    Set objChildNode = objDOM.createElement("HEAHEA")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
    'DocNumHEA5
    'Document Number
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("DocNumHEA5"))
    objChildElement.Text = G_clsIE44Arrival.Headers(G_strHeaderKey).MOVEMENT_REFERENCE_NUMBER 'RetrieveRecordForXML("DocNumHEA5")
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

            'IdeOfMeaOfTraAtDHEA78
    'Identity of Means of Transport at Departure
    strTemp = G_clsIE44Arrival.Headers(G_strHeaderKey).IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B1", 1)(0)
    If Len(Trim$(strTemp)) > 0 Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraAtDHEA78"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    End If
            'IdeOfMeaOfTraAtDHEA78LNG
    'Identity of Means of Transport at Departure Language
    strTemp = G_clsIE44Arrival.Headers(G_strHeaderKey).IDENTITY_MEANS_OF_TRANSPORT_AT_DEPARTURE_LNG 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
    If Len(Trim$(strTemp)) > 0 Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("IdeOfMeaOfTraAtDHEA78LNG"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    End If
            'NatOfMeaOfTraAtDHEA80
    'Nationality of Means of Transport at Departure
    strTemp = G_clsIE44Arrival.Headers(G_strHeaderKey).NATIONALITY_OF_MEANS_OF_TRANSPORT_AT_DEPARTURE 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "B7", 1)(0)
    If Len(Trim$(strTemp)) > 0 Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NatOfMeaOfTraAtDHEA80"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    End If
            'TotNumOfIteHEA305
    'Total Number of Items
    'm_clsEDINCTSIE44Message.Headers(m_strHeaderKey).ARRIVAL_NOTIFICATION_DATE
    strTemp = G_clsIE44Arrival.Headers(G_strHeaderKey).TOTAL_NUMBER_OF_ITEMS
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotNumOfIteHEA305"))
    objChildElement.Text = IIf(Len(Trim$(strTemp)) = 0, "0", strTemp)
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

            'TotNumOfPacHEA306
    'Total Number of Packages
    strTemp = G_clsIE44Arrival.Headers(1).TOTAL_NUMBER_OF_PACKAGES 'GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "C4", 1)(0)
    If Len(Trim$(strTemp)) > 0 Then
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotNumOfPacHEA306"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    End If
            'TotGroMasHEA307
    'Total Gross Mass
    strTemp = G_clsIE44Arrival.Headers(1).TOTAL_GROSS_MASS
    Set objChildElement = objChildNode.appendChild(objDOM.createElement("TotGroMasHEA307"))
    objChildElement.Text = IIf(Len(Trim$(strTemp)) = 0, "0", strTemp)
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
           
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub

Private Sub CreateMessageConsignorHeader(ByRef objDOM As DOMDocument, _
                                        ByRef objParentNode As IXMLDOMNode, _
                                        ByRef objChildNode As IXMLDOMNode, _
                                        ByRef objChildElement As IXMLDOMNode, _
                                        ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    'Consignor
    Set objChildNode = objDOM.createElement("TRACONCO1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamCO17"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U2", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Street and Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumCO122"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U3", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Postal Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodCO123"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U4", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'City
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitCO124"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U8", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouCO125"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U7", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'NAD Language
        strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGCO"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'TIN
        strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "U6", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINCO159"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Private Sub CreateMessageConsigneeHeader(ByRef objDOM As DOMDocument, _
                                        ByRef objParentNode As IXMLDOMNode, _
                                        ByRef objChildNode As IXMLDOMNode, _
                                        ByRef objChildElement As IXMLDOMNode, _
                                        ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    
    'Consignee
    Set objChildNode = objDOM.createElement("TRACONCE1")
    objDOM.documentElement.appendChild objChildNode
    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)

        'Name
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("NamCE17"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W1", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Street and Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("StrAndNumCE122"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W2", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'Postal Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("PosCodCE123"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W4", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
    
        'City
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CitCE124"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W3", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Country Code
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("CouCE125"))
        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W5", 1)(0)
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'NAD Language
        strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "A5", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NADLNGCE"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
        'TIN
        strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W6", 1)(0)
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINCE159"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                
    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
End Sub


Private Sub CreateMessageAuthorised(ByRef objDOM As DOMDocument, _
                                   ByRef objParentNode As IXMLDOMNode, _
                                   ByRef objChildNode As IXMLDOMNode, _
                                   ByRef objChildElement As IXMLDOMNode, _
                                   ByRef objChildElement2 As IXMLDOMNode)
    
    Dim strTemp As String
    Dim lngDetailCount As Long
    Dim lngctr As Long
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    
    '************************************************************************************************************
    'Rule for Authorised Consignee
    '************************************************************************************************************
    '       If W7 = Empty String or W7 <> "Y" then do not include "Authorised Consignee" segment
    '       If W6 = Empty String then do not include "Authorised Consignee" segment
    '************************************************************************************************************
    lngDetailCount = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "F<DETAIL COUNT>", 0)(0)
    
    If VentureNumberAreTheSame("W6") = True Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "DATA_NCTS_DETAIL "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
            strSQL = strSQL & "AND "
            strSQL = strSQL & "Detail = 1"
                
        ADORecordsetOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
        'RstOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
        
        If rstTemp.RecordCount > 0 Then
            strTemp = IIf(IsNull(rstTemp.Fields("W7").Value), "", Trim$(rstTemp.Fields("W7").Value))
        Else
            strTemp = vbNullString
        End If
        
        If Not IsNull(strTemp) Then
            If UCase$(Trim$(strTemp)) = "Y" Then
        
                strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W6", 1)(0)
                
                If Len(Trim$(strTemp)) > 0 Then
                    'Authorised
                    Set objChildNode = objDOM.createElement("TRAAUTCONTRA")
                    objDOM.documentElement.appendChild objChildNode
                    objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                
                        'TIN
                        Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRA59"))
                        objChildElement.Text = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W6", 1)(0)
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                    objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
                End If
                
            End If
        End If
    
    Else
        For lngctr = 1 To lngDetailCount
            
                strSQL = vbNullString
                strSQL = strSQL & "SELECT * "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "DATA_NCTS_DETAIL "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "Code = '" & G_strUniqueCode & "' "
                strSQL = strSQL & "AND "
                strSQL = strSQL & "Detail = " & lngctr
            ADORecordsetOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic
            'RstOpen strSQL, g_conEdifact, rstTemp, adOpenKeyset, adLockOptimistic, , True
            
            If rstTemp.RecordCount > 0 Then
                strTemp = IIf(IsNull(rstTemp.Fields("W7").Value), "", Trim$(rstTemp.Fields("W7").Value))
            Else
                strTemp = vbNullString
            End If
            
            If Not IsNull(strTemp) Then
                If UCase$(Trim$(strTemp)) = "Y" Then
                    
                    strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "W6", lngctr)(0)
                    
                    If Len(Trim$(strTemp)) > 0 Then
                        'Authorised
                        Set objChildNode = objDOM.createElement("TRAAUTCONTRA")
                        objDOM.documentElement.appendChild objChildNode
                        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                    
                            'TIN
                            Set objChildElement = objChildNode.appendChild(objDOM.createElement("TINTRA59"))
                            objChildElement.Text = strTemp
                            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
                        
                        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
                        
                        Exit For
                    End If
                    
                End If
            End If
            
            strTemp = GetValueFromClass(G_clsEDIDeparture, G_rstDepartureMap, enuIE29Val_NotFromIE29, "T7", lngctr)(0)
            
            If UCase$(Trim$(strTemp)) = "F" Then Exit For
            
        Next
        
        
    End If
    
    ADORecordsetClose rstTemp
    'RstClose rstTemp
    
End Sub

'021408
Private Sub CreateMessageSealsInfo(ByRef objDOM As DOMDocument, _
                                  ByRef objParentNode As IXMLDOMNode, _
                                  ByRef objChildNode As IXMLDOMNode, _
                                  ByRef objChildElement As IXMLDOMNode, _
                                  ByRef objChildElement2 As IXMLDOMNode)
    
    Dim lngSealCount As Long
    Dim lngctr As Long
    
    strTemp = G_clsIE44Arrival.SealInfos(G_strHeaderKey).FIELD_SEALS_NUMBER
    If Len(Trim$(strTemp)) > 0 Then

        'Seals Info
        Set objChildNode = objDOM.createElement("SEAINFSLI")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Seals Number
        strTemp = G_clsIE44Arrival.SealInfos(G_strHeaderKey).FIELD_SEALS_NUMBER
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("SeaNumSLI2"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                        
            
        strTemp = G_clsIE44Arrival.SealInfos(G_strHeaderKey).Seals(G_strHeaderKey).SEALS_IDENTITY
        If Len(Trim$(strTemp)) > 0 Then
            'Seals ID
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("SEAIDSID"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
            'Seals Identity
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1"))
            objChildElement2.Text = G_clsIE44Arrival.SealInfos(G_strHeaderKey).Seals(G_strHeaderKey).SEALS_IDENTITY
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                    
            'Seals Identity Language
            strTemp = G_clsIE44Arrival.SealInfos(G_strHeaderKey).Seals(G_strHeaderKey).SEALS_IDENTITY_LNG
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SeaIdeSID1LNG"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            End If
                
            'objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
        End If
    End If
    
End Sub

'021408
Private Sub CreateMessageGoodsItem(ByRef objDOM As DOMDocument, _
                                  ByRef objParentNode As IXMLDOMNode, _
                                  ByRef objChildNode As IXMLDOMNode, _
                                  ByRef objChildElement As IXMLDOMNode, _
                                  ByRef objChildElement2 As IXMLDOMNode)
                                  
    Dim lngDetailCount As Long
    Dim lngctr As Long
    Dim lngPrevDocCtr As Long
    Dim lngContainerCtr As Long
    Dim lngPackagesCtr As Long
       
    strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_ITEM_NUMBER
    If Len(Trim$(strTemp)) > 0 Then
        
        'Goods Item
        Set objChildNode = objDOM.createElement("GOOITEGDS")
        objDOM.documentElement.appendChild objChildNode
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            
        'Item Number
        Set objChildElement = objChildNode.appendChild(objDOM.createElement("IteNumGDS7"))
        objChildElement.Text = strTemp
        objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        'Commodity Code
        strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_COMMODITY_CODE
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("ComCodTarCodGDS10"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Goods Description
        strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_GOODS_DESCRIPTION
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GooDesGDS23"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Goods Description Language
        strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_GOODS_DESCRIPTION_LNG
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GooDesGDS23LNG"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
                        
        'Gross Mass
        strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_GROSS_MASS
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("GroMasGDS46"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        'Net Mass
        strTemp = G_clsIE44Arrival.GoodsItems(1).FIELD_NET_MASS
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("NetMasGDS48"))
            objChildElement.Text = strTemp
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        End If
        
        'Produced Documents/Certificates
        If Len(Trim$(G_clsIE44Arrival.GoodsItems(1).DocumentCertificates(1).FIELD_DOCUMENT_TYPE)) > 0 Or _
           Len(Trim$(G_clsIE44Arrival.GoodsItems(1).DocumentCertificates(1).FIELD_DOCUMENT_REFERENCE)) > 0 Or _
           Len(Trim$(G_clsIE44Arrival.GoodsItems(1).DocumentCertificates(1).FIELD_DOCUMENT_REFERENCE_LNG)) > 0 Or _
           Len(Trim$(G_clsIE44Arrival.GoodsItems(1).DocumentCertificates(1).FIELD_COMPLEMENT_INFORMATION)) > 0 Or _
           Len(Trim$(G_clsIE44Arrival.GoodsItems(1).DocumentCertificates(1).FIELD_COMPLEMENT_INFORMATION_LNG)) > 0 Then
            
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PRODOCDC2"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            'Produced Document Type
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocTypDC21"))
            objChildElement2.Text = G_clsIE44Arrival.GoodsItems(G_strHeaderKey).DocumentCertificates(G_strHeaderKey).FIELD_DOCUMENT_TYPE
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            'Produced Document Reference
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocRefDC23"))
            objChildElement2.Text = G_clsIE44Arrival.GoodsItems(G_strHeaderKey).DocumentCertificates(G_strHeaderKey).FIELD_DOCUMENT_REFERENCE
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            'Produced Document Reference Language
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("DocRefDCLNG"))
            objChildElement2.Text = G_clsIE44Arrival.GoodsItems(G_strHeaderKey).DocumentCertificates(G_strHeaderKey).FIELD_DOCUMENT_REFERENCE_LNG
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            'Complement of Information
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfDC25"))
            objChildElement2.Text = G_clsIE44Arrival.GoodsItems(G_strHeaderKey).DocumentCertificates(G_strHeaderKey).FIELD_COMPLEMENT_INFORMATION
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            'Complement of Information Language
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ComOfInfDC25LNG"))
            objChildElement2.Text = G_clsIE44Arrival.GoodsItems(G_strHeaderKey).DocumentCertificates(G_strHeaderKey).FIELD_COMPLEMENT_INFORMATION_LNG
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)

            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
            objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)
    
        End If
        '*******************************************************************************************************
        'Container
        '*******************************************************************************************************
        strTemp = G_clsIE44Arrival.GoodsItems(1).Containers(1).NEW_CONTAINER_NUMBER
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("CONNR2"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
            Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("ConNumNR21"))
            objChildElement2.Text = strTemp
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
              
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        End If
        '*******************************************************************************************************
        'Packages
        '*******************************************************************************************************
        strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_KIND_OF_PACKAGES
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("PACGS2"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
                
            'Marks and Numbers
            strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_MARKS_AND_NUMBERS_OF_PACKAGES
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("MarNumOfPacGS21"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            'Marks and Numbers Language
            strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_MARKS_AND_NUMBERS_OF_PACKAGES_LNG
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("MarNumOfPacGS21LNG"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            'Package Type
            strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_KIND_OF_PACKAGES
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("KinOfPacGS23"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            'Number of Packages
            strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_NUMBER_OF_PACKAGES
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NumOfPacGS24"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            'Number of Pieces
            strTemp = G_clsIE44Arrival.GoodsItems(1).Packages(1).FIELD_NUMBER_OF_PIECES
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("NumOfPieGS25"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        End If
        '*******************************************************************************************************
        'Sensitive Goods
        '*******************************************************************************************************
        strTemp = G_clsIE44Arrival.GoodsItems(1).SGICodes(1).FIELD_SENSITIVE_QUANTITY
        If Len(Trim$(strTemp)) > 0 Then
            Set objChildElement = objChildNode.appendChild(objDOM.createElement("SGICODSD2"))
            objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            
            'Sensitive Goods Code / Sensitive Goods Quantity
            strTemp = G_clsIE44Arrival.GoodsItems(1).SGICodes(1).FIELD_SENSITIVE_GOODS_CODE
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenGooCodSD22"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            strTemp = G_clsIE44Arrival.GoodsItems(1).SGICodes(1).FIELD_SENSITIVE_QUANTITY
            If Len(Trim$(strTemp)) > 0 Then
                Set objChildElement2 = objChildElement.appendChild(objDOM.createElement("SenQuaSD23"))
                objChildElement2.Text = strTemp
                objChildElement.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab & vbTab)
            End If
            
            objChildNode.appendChild objDOM.createTextNode(vbNewLine & vbTab & vbTab)
        
        End If
    
        objParentNode.appendChild objDOM.createTextNode(vbNewLine & vbTab)

    End If
    
End Sub
    

