VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntrepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents pckEntrepot As PCubeLibPick.CPicklist
Attribute pckEntrepot.VB_VarHelpID = -1
Private mvarLanguage As String
Private mvarConn As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarSelectedEntrepot As String 'local copy
Private mvarSelectedEntrepotCountry As String 'local copy
Private mvarCancelled As Boolean 'local copy
Private mvarEntrepot_ID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarResourceHandler As Long 'local copy
Private mvarStartingNum As String

Public Property Let SelectedEntrepotCountry(ByVal vData As String)
    mvarSelectedEntrepotCountry = vData
End Property
 
Public Property Get SelectedEntrepotCountry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedEntrepot
    SelectedEntrepotCountry = mvarSelectedEntrepotCountry
End Property

Public Property Let ResourceHandler(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResourceHandler = 5
    mvarResourceHandler = vData
End Property

Public Property Get ResourceHandler() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResourceHandler
    ResourceHandler = mvarResourceHandler
End Property

Public Property Let Entrepot_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Entrepot_ID = 5
    mvarEntrepot_ID = vData
End Property

Public Property Get Entrepot_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Entrepot_ID
    Entrepot_ID = mvarEntrepot_ID
End Property

Public Property Let Cancelled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cancelled = 5
    mvarCancelled = vData
End Property

Public Property Get Cancelled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cancelled
    Cancelled = mvarCancelled
End Property

Public Property Let SelectedEntrepot(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedEntrepot = 5
    mvarSelectedEntrepot = vData
End Property

Public Property Get SelectedEntrepot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedEntrepot
    SelectedEntrepot = mvarSelectedEntrepot
End Property

Public Property Let StartingNum(ByVal vData As String)
    mvarStartingNum = vData
End Property

Public Property Get StartingNum() As String
    StartingNum = mvarStartingNum
End Property

Public Sub ShowEntrepot(CallingForm As Object, _
                        ByVal conn As ADODB.Connection, _
                        AutoUnload As Boolean, _
                        Language As String, ByVal MyResourceHandler As Long, _
                        Optional Pass2CtrlName As String, _
                        Optional lngSearchID As Long, _
                        Optional UseSimple As Boolean)
    Dim gsdEntrepot As PCubeLibPick.CGridSeed
    Dim strEntrepotSQL As String
    
    ResourceHandler = MyResourceHandler
    
    'For picklist's filter functionality.
    Dim clsPicklistFilter As CPicklistFilter
    
    Set clsPicklistFilter = New CPicklistFilter
    Set pckEntrepot = New CPicklist
    Set gsdEntrepot = New CGridSeed
    
    Set gsdEntrepot = pckEntrepot.SeedGrid("Starting Num", 0, "Left", "Entrepot Type", 1300, "Left", "Entrepot Number", 2970, "Left")
    Set mvarConn = conn
    
    'Required by Authorized Party
    mvarLanguage = Language
    
    'The primary key is mentioned twice to conform to the design of the picklist class.
    'allanSQL
    strEntrepotSQL = vbNullString
    strEntrepotSQL = strEntrepotSQL & "SELECT "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_Starting_Num AS [Starting Num], "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_ID AS ID, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_ID AS Entrepot_ID, "
    strEntrepotSQL = strEntrepotSQL & "Auth_ID AS Auth_ID, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_Type AS [Entrepot Type], "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_Num AS [Entrepot Number], "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_Country AS [Entrepot Country], "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_StartDate, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_EndDate, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_StockCard_Numbering, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_StockCard_Format, "
    strEntrepotSQL = strEntrepotSQL & "Entrepot_StockCard_Archiving "
    strEntrepotSQL = strEntrepotSQL & "FROM "
    strEntrepotSQL = strEntrepotSQL & "Entrepots "
    
    With pckEntrepot
        .Search True, "Entrepot_ID", lngSearchID
        
        'Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        If UseSimple = True Then
            .Pick CallingForm, cpiSimplePicklist, conn, strEntrepotSQL, "ID", "Entrepot", vbModal, gsdEntrepot, , , AutoUnload, cpiKeyF2
        Else
            clsPicklistFilter.PicklistFilters.Add "Key", "Entrepot", "Entrepot_ID <> 0", True
            .Pick CallingForm, cpiFilterCatalog, conn, strEntrepotSQL, "ID", "Entrepot", vbModal, gsdEntrepot, clsPicklistFilter, , AutoUnload, cpiKeyF2
        End If
        
        If .CancelTrans = False And Not (.SelectedRecord Is Nothing) Then
            mvarStartingNum = .SelectedRecord.RecordSource.Fields("Starting Num")
            mvarEntrepot_ID = .SelectedRecord.RecordSource.Fields("Entrepot_ID")
            mvarSelectedEntrepot = .SelectedRecord.RecordSource.Fields("Entrepot Type").Value & "-" & _
                                   .SelectedRecord.RecordSource.Fields("Entrepot Number").Value
                                   
            If IsNull(.SelectedRecord.RecordSource.Fields("Entrepot Country").Value) Then
                mvarSelectedEntrepotCountry = ""
            Else
                mvarSelectedEntrepotCountry = .SelectedRecord.RecordSource.Fields("Entrepot Country").Value
            End If
        End If
        
        'Allows Entrepot to pass values back to a textbox control.
        If Len(Pass2CtrlName) <> 0 Then
            If Not .SelectedRecord Is Nothing Then
                CallingForm.Controls(Pass2CtrlName).Text = .SelectedRecord.RecordSource.Fields("Entrepot Type") & "-" & _
                                                           .SelectedRecord.RecordSource.Fields("Entrepot Number").Value
                CallingForm.Controls(Pass2CtrlName).Tag = .SelectedRecord.RecordSource.Fields("Entrepot_ID")
            End If

        End If
        mvarCancelled = .CancelTrans
    End With
    
    Set gsdEntrepot = Nothing
    Set pckEntrepot = Nothing
    Set clsPicklistFilter = Nothing
End Sub

Public Sub ShowEntrepotNumber(CallingForm As Object, _
                              ByVal conn As ADODB.Connection, _
                              AutoUnload As Boolean, _
                              Language As String, ByVal MyResourceHandler As Long, _
                              Optional EntrepotType As String = "")
    Dim gsdEntrepot As PCubeLibPick.CGridSeed
    Dim strEntrepotSQL As String
    
    ResourceHandler = MyResourceHandler
    
    'For picklist's filter functionality.
    Dim clsPicklistFilter As CPicklistFilter
    
    Set clsPicklistFilter = New CPicklistFilter
    Set pckEntrepot = New CPicklist
    Set gsdEntrepot = New CGridSeed
    
    Set gsdEntrepot = pckEntrepot.SeedGrid("Starting Num", 0, "Left", "Entrepot Type", 0, "Left", "Entrepot Number", 2970, "Left")
    Set mvarConn = conn
    
    'Required by Authorized Party
    mvarLanguage = Language
    
    'The primary key is mentioned twice to conform to the design of the picklist class.
    strEntrepotSQL = "SELECT Entrepot_Starting_Num AS [Starting Num], Entrepot_ID AS ID, Entrepot_ID AS Entrepot_ID, Auth_ID AS Auth_ID, Entrepot_Type AS [Entrepot Type], Entrepot_Num AS [Entrepot Number], Entrepot_StartDate, Entrepot_EndDate, Entrepot_StockCard_Numbering, Entrepot_StockCard_Format, Entrepot_StockCard_Archiving FROM Entrepots"
    
    If Len(EntrepotType) > 0 Then
        strEntrepotSQL = strEntrepotSQL & " WHERE Entrepot_Type = '" & EntrepotType & "'"
    End If
    
    With pckEntrepot
        '.Search True, "Entrepot_ID", lngSearchID
        
        'Setting the KeyPick argument to cpiKeyF2 positions the selected item to the branch code being searched for above.
        .Pick CallingForm, cpiSimplePicklist, conn, strEntrepotSQL, "ID", "Entrepot", vbModal, gsdEntrepot, , , AutoUnload, cpiKeyF2
        
        If .CancelTrans = False And Not (.SelectedRecord Is Nothing) Then
            mvarStartingNum = .SelectedRecord.RecordSource.Fields("Starting Num")
            mvarEntrepot_ID = .SelectedRecord.RecordSource.Fields("Entrepot_ID")
            mvarSelectedEntrepot = .SelectedRecord.RecordSource.Fields("Entrepot Type") & "-" & _
                                   .SelectedRecord.RecordSource.Fields("Entrepot Number")
        End If
        
        mvarCancelled = .CancelTrans
    End With
    
    Set gsdEntrepot = Nothing
    Set pckEntrepot = Nothing
    Set clsPicklistFilter = Nothing
End Sub

Private Sub pckEntrepot_BeforeDelete(ByVal BaseName As String, ByVal ID As Variant, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    Dim rstChkEntrepot As ADODB.Recordset
    Dim strSQL As String 'allanSQL
    
    'If There are products associated with that Entrepot ID then..
        
        strSQL = vbNullString
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "Entrepot_ID AS [Entrepot ID], "
        strSQL = strSQL & "Prod_Num AS [Product Num] "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "Products "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "Entrepot_ID = " & ID & " "
    ADORecordsetOpen strSQL, mvarConn, rstChkEntrepot, adOpenKeyset, adLockOptimistic
    'rstChkEntrepot.Open strSQL, mvarConn, adOpenKeyset, adLockOptimistic
    If Not (rstChkEntrepot.BOF And rstChkEntrepot.EOF) Then
        MsgBox Translate(2290) & vbCrLf & _
               rstChkEntrepot.RecordCount & Space(1) & Translate(2291), _
               vbOKOnly + vbInformation, "Entrepot"
        Cancel = True
    Else
        Cancel = False
    End If
    
    ADORecordsetClose rstChkEntrepot
End Sub

Private Sub pckEntrepot_BtnClick(Record As PCubeLibPick.CRecord, ByVal Button As PCubeLibPick.ButtonType, Cancel As Boolean)
    'Loads the maintainance form.
    Select Case Button
        Case cpiAdd, cpiModify, cpiCopy
            frmEntrepot.Pre_Load mvarConn, Record.RecordSource, mvarLanguage, Button, Cancel, ResourceHandler, pckEntrepot.GridRecord
'        Case cpiDelete
    End Select
End Sub

Public Sub ShowRepackaging(ByRef connSadbel As ADODB.Connection, ByRef connTaric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, ByVal UserID As Long)
                    
    frmRepackaging.MyLoad connSadbel, connTaric, TaricProperties, Language, MyResourceHandler, UserID
    Set frmRepackaging = Nothing
End Sub

Public Sub ShowInitialStock(ByRef connSadbel As ADODB.Connection, ByRef connTaric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, _
                  ByVal UserID As Long)
                    
    frmInitialStock.MyLoad connSadbel, connTaric, TaricProperties, Language, MyResourceHandler, UserID
    Set frmInitialStock = Nothing
End Sub

Public Sub ShowManualOutbound(ByRef connSadbel As ADODB.Connection, ByRef connTaric As ADODB.Connection, _
                  ByVal TaricProperties As Integer, ByVal Language As String, ByVal MyResourceHandler As Long, ByVal UserID As Long, Optional IOCorrectionMode As String)
                    
    frmManualOutbound.MyLoad connSadbel, connTaric, TaricProperties, Language, MyResourceHandler, UserID, IOCorrectionMode
    Set frmManualOutbound = Nothing
End Sub

Public Sub ShowResetStock(ByRef connSadbel As ADODB.Connection, connEdifact As ADODB.Connection, ByVal UserID As Long)
    frmResetReservedStocks.My_Load connSadbel, connEdifact, UserID
    Set frmResetReservedStocks = Nothing
End Sub

Public Sub ShowClosureForm(ByRef connSadbel As ADODB.Connection, ByVal Language As String, ByVal ResourceHandler As Long, ByVal UserID As Long)
    frmClosure.MyLoad connSadbel, Language, ResourceHandler, UserID
    Set frmClosure = Nothing
End Sub
