VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAuthorizedParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mLanguage As String
Private mConnection As ADODB.Connection
Private WithEvents pckAuthorizedParties As PCubeLibPick.CPicklist
Attribute pckAuthorizedParties.VB_VarHelpID = -1




Public Sub ShowAuthorizedParty(ByVal Index As Long, _
                               ByRef CallingForm As Object, _
                               ByRef Connection As ADODB.Connection, _
                               ByVal Language As String, _
                               ByVal MyResourceHandler As Long, _
                      Optional ByVal strTxtBox As String, _
                      Optional ByRef lngID As Long)

    Dim gsdAuthorizedParties As PCubeLibPick.CGridSeed
    Dim strAuthorizedPartiesSQL As String
    'For picklist's filter functionality.
    Dim clsPicklistFilter As CPicklistFilter
    
    ResourceHandler = MyResourceHandler
    
    Set clsPicklistFilter = New CPicklistFilter
    Set pckAuthorizedParties = New CPicklist
    Set gsdAuthorizedParties = New CGridSeed
    Set gsdAuthorizedParties = pckAuthorizedParties.SeedGrid("Name", 1300, "Left", "Address", 2970, "Left")
    ' The primary key is mentioned twice to conform to the design of the picklist class.
        'allanSQL
        strAuthorizedPartiesSQL = vbNullString
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "SELECT "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "Auth_ID AS ID, "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "Auth_ID AS Auth_ID, "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "Auth_Name AS Name, "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "Auth_Address AS Address "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "FROM "
        strAuthorizedPartiesSQL = strAuthorizedPartiesSQL & "AuthorizedParties"
    Set mConnection = Connection
    mLanguage = Language
    
    Select Case Index
        Case 0
            With pckAuthorizedParties
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick CallingForm, cpiFilterCatalog, Connection, strAuthorizedPartiesSQL, "ID", "Authorized Party", vbModal, gsdAuthorizedParties, clsPicklistFilter, , False, cpiKeyF2
                If .CancelTrans = False Then
                
                End If
            End With
        Case 1
            With pckAuthorizedParties
                .Search True, "Name", Trim(CallingForm.Controls(strTxtBox).Text)
                
                clsPicklistFilter.PicklistFilters.Add "Key", "Authorized Party", "Auth_ID <> 0", True
                
                ' Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
                .Pick CallingForm, cpiFilterCatalog, Connection, strAuthorizedPartiesSQL, "ID", "AuthorizedParty", vbModal, gsdAuthorizedParties, clsPicklistFilter, , True, cpiKeyF2
                
                If Not .SelectedRecord Is Nothing Then
                    CallingForm.Controls(strTxtBox).Text = .SelectedRecord.RecordSource.Fields("Name").Value
                    lngID = .SelectedRecord.RecordSource.Fields("ID").Value
                End If
            End With
    End Select
                
    Set gsdAuthorizedParties = Nothing
    Set pckAuthorizedParties = Nothing
    Set clsPicklistFilter = Nothing
End Sub



Private Sub pckAuthorizedParties_BeforeDelete(ByVal BaseName As String, ByVal ID As Variant, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Dim rstTmp As ADODB.Recordset
    Dim strComand As String
    
        'allanSQL
        strComand = vbNullString
        strComand = strComand & "Select "
        strComand = strComand & "Auth_ID AS [Auth ID], "
        strComand = strComand & "Entrepot_Num AS [Entrepot Num] "
        strComand = strComand & "FROM "
        strComand = strComand & "Entrepots "
        strComand = strComand & "WHERE "
        strComand = strComand & "Auth_ID = " & ID & ""
    ADORecordsetOpen strComand, mConnection, rstTmp, adOpenKeyset, adLockOptimistic
    'rstTmp.Open strComand, mConnection, adOpenKeyset, adLockReadOnly
    
    If Not (rstTmp.BOF And rstTmp.EOF) Then
        MsgBox Translate(2254) & vbCrLf & _
               rstTmp.RecordCount & Space(1) & Translate(2255), _
               vbOKOnly + vbInformation, Translate(2308)
        Cancel = True
    Else
        Cancel = False
    End If
    
    ADORecordsetClose rstTmp
End Sub

Private Sub pckAuthorizedParties_BtnClick(Record As PCubeLibPick.CRecord, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    If Button = cpiDelete Then
        'Record.RecordSource!ActionTag = "D"
        Record.RecordSource.Update
        
        UpdateRecordset mConnection, Record.RecordSource, "AuthorizedParties"
    Else
        frmAuthorizedParty.MyLoad Record.RecordSource, Button, Cancel, mLanguage, mConnection, ResourceHandler
    End If
End Sub
